#!/usr/bin/env python3
"""
Loads To Archive in Ascend -> HTML action report

Delivered loads that are still unarchived in Ascend are left out of the Savings report.
This script reads an Ascend loads export (CSV), keeps only loads with Load Status = "Delivered"
(i.e. delivered but not yet archived), extracts the Load Planner name from "Users & Roles"
(removing the "Load Planner:" text), and writes a self-contained HTML report that tells each
Load Planner exactly which Load IDs they need to archive.

Usage:
    python delivered_loads_report.py ascend_loads_export.csv
    python delivered_loads_report.py ascend_loads_export.csv -o loads_to_archive.html

Requires: pandas  (pip install pandas)
"""
import argparse
import html
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

DEFAULT_MIN_DAYS = 7  # only report loads delivered at least this many days ago (e.g. run 11/24 -> up to 11/17)
BOT_PATTERN = re.compile(r"^ECBOT", re.I)  # automated planners (no person to ask)


# ---------- data prep ----------
def extract_load_planner(value):
    """'Load Planner: Tom Massam\\nLoad Builder: ...' -> 'Tom Massam' (unique names, comma-joined)."""
    if pd.isna(value):
        return "Unassigned"
    names = []
    for name in re.findall(r"Load Planner:\s*([^\n]+)", str(value)):
        name = name.strip()
        if name and name not in names:
            names.append(name)
    return ", ".join(names) if names else "Unassigned"


def to_date(series):
    return pd.to_datetime(series.astype(str).str[:10], format="%m/%d/%Y", errors="coerce")


REQUIRED_COLUMNS = ["Load ID", "Load Status", "Users & Roles", "Last Drop Date"]


def read_export(source, filename=None):
    """Read an Ascend export from a path or an uploaded file object (.csv, .xlsx or .xls)."""
    name = str(filename or getattr(source, "name", source)).lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(source)
    return pd.read_csv(source, low_memory=False)


def only_delivered(df):
    """Keep ONLY loads whose Load Status is exactly 'Delivered'. Every other status is excluded."""
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"This doesn't look like an Ascend loads export. Missing column(s): {', '.join(missing)}")
    status = df["Load Status"].astype(str).str.strip().str.casefold()
    d = df[status == "delivered"].copy()
    assert (d["Load Status"].astype(str).str.strip().str.casefold() == "delivered").all()
    return d


def cutoff_date(as_of, min_days=DEFAULT_MIN_DAYS):
    """Latest delivery date included in the report: today minus min_days (inclusive)."""
    return pd.Timestamp(as_of.date()) - pd.Timedelta(days=min_days)


def prepare(df, as_of, min_days=DEFAULT_MIN_DAYS):
    d = only_delivered(df)
    d["Load Planner"] = d["Users & Roles"].map(extract_load_planner)
    d["Delivery Date"] = to_date(d["Last Drop Date"])
    d["Days"] = (pd.Timestamp(as_of.date()) - d["Delivery Date"]).dt.days
    # Only loads delivered on or before the cutoff (past one week). Loads with no delivery date are excluded.
    d = d[d["Delivery Date"].notna() & (d["Delivery Date"] <= cutoff_date(as_of, min_days))].copy()
    d["Carrier"] = d["Carrier"].astype(str).str.replace(r"\s*-\s*\[\d+\]$", "", regex=True)
    d["Load ID"] = d["Load ID"].astype(str).str.replace(r"\.0$", "", regex=True)
    d = d.sort_values(["Load Planner", "Delivery Date"])
    return d


def summarize(d):
    g = d.groupby("Load Planner").agg(
        Loads=("Load ID", "count"),
        Oldest=("Delivery Date", "min"),
        Newest=("Delivery Date", "max"),
        MaxDays=("Days", "max"),
        Over14=("Days", lambda s: int((s > 14).sum())),
    )
    g["IsBot"] = [bool(BOT_PATTERN.match(p)) or p == "Unassigned" for p in g.index]
    # people first (most loads first), then bots / unassigned
    g = g.sort_values(["IsBot", "Loads"], ascending=[True, False])
    g["Anchor"] = [f"p{i}" for i in range(len(g))]
    return g


# ---------- formatting helpers ----------
def esc(x):
    return "" if pd.isna(x) else html.escape(str(x))


def fdate(x):
    return "" if pd.isna(x) else x.strftime("%m/%d/%Y")


def age_cls(days):
    if pd.isna(days):
        return ""
    return "a3" if days > 14 else "a2"


def ndays(x):
    return "" if pd.isna(x) else f"{int(x)}"


# ---------- HTML ----------
def build_html(d, g, source_name, as_of, min_days=DEFAULT_MIN_DAYS):
    n = len(d)
    cutoff = cutoff_date(as_of, min_days)
    first, last = d["Delivery Date"].min(), d["Delivery Date"].max()
    iso = lambda x: "" if pd.isna(x) else x.strftime("%Y-%m-%d")
    generated = as_of.strftime("%m/%d/%Y %I:%M %p")

    planners = [{"name": p, "anchor": r.Anchor, "bot": bool(r.IsBot)} for p, r in g.iterrows()]
    planners_json = json.dumps(planners).replace("</", "<\\/")
    report_key = f"archive-report-{as_of:%Y%m%d}"

    cards = ""
    for p, r in g.iterrows():
        sub = d[d["Load Planner"] == p].sort_values("Delivery Date")
        trs = "".join(
            f'<tr class="ld" data-id="{esc(x["Load ID"])}" data-p="{esc(p)}" data-date="{iso(x["Delivery Date"])}" data-days="{ndays(x["Days"])}">'
            f'<td class="cb"><input type="checkbox" class="arch" title="Tick when archived in Ascend" aria-label="Archived {esc(x["Load ID"])}"></td>'
            f'<td><b>{esc(x["Load ID"])}</b></td><td>{fdate(x["Delivery Date"])}</td>'
            f'<td class=n><span class="pill {age_cls(x["Days"])}">{ndays(x["Days"])}</span></td>'
            f'<td>{esc(x["First Pick Name"])}</td><td>{esc(x["Last Drop Name"])}'
            f'<span class=sub> · {esc(x["Last Drop City"])}, {esc(x["Last Drop State"])}</span></td>'
            f'<td>{esc(x["Carrier"])}</td></tr>'
            for _, x in sub.iterrows()
        )
        tag = '<span class="bot">Bot / unassigned</span>' if r.IsBot else ""
        cards += f"""
<details class="pcard" id="{r.Anchor}" data-p="{esc(p)}">
<summary><span class="pname">{esc(p)}</span> {tag}
<span class="pcount"><span class="c-rem">{int(r.Loads)}</span> to archive</span>
<span class="pdone"><span class="c-arch">0</span> archived</span>
<span class="pold">oldest remaining: <span class="c-old">{fdate(r.Oldest)}</span></span></summary>
<div class="ids"><span class="idl">Remaining Load IDs:</span> <code class="idlist"></code>
<button type="button" class="copy">Copy IDs</button></div>
<table><thead><tr><th class="cb"><input type="checkbox" class="arch-all" title="Mark all shown as archived" aria-label="Mark all archived"></th>
<th>Load ID</th><th>Delivered</th><th class=n>Days</th><th>Pickup</th><th>Delivery</th><th>Carrier</th></tr></thead>
<tbody>{trs}</tbody></table></details>"""

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Delivered Loads to Archive in Ascend</title>
<style>
:root{{--navy:#1F3864;--blue:#2E75B6;--red:#C00000;--amber:#B7791F;--green:#2F7D32;--line:#e3e7ee;--muted:#6b7280}}
*{{box-sizing:border-box}} body{{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f5f7fa;color:#1f2937}}
header{{background:var(--navy);color:#fff;padding:22px 32px}} header h1{{margin:0;font-size:22px}}
header p{{margin:6px 0 0;font-size:12.5px;opacity:.9}} main{{padding:24px 32px;max-width:1300px;margin:auto}}
.period{{display:flex;gap:28px;flex-wrap:wrap;margin-top:12px;font-size:12px}}
.period div span{{display:block;opacity:.7;text-transform:uppercase;font-size:10.5px;letter-spacing:.05em}} .period b{{font-size:14px}}
.ask{{background:#FFF8E1;border:1px solid #F2D27A;border-left:6px solid #E0A800;border-radius:8px;padding:16px 20px;margin-bottom:22px;font-size:14px;line-height:1.5}}
.ask h2{{margin:0 0 6px;font-size:16px;color:#7A5200}} .ask ol{{margin:6px 0 0 18px;padding:0}}
.filters{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px 18px;margin-bottom:22px;display:flex;gap:14px;align-items:end;flex-wrap:wrap}}
.filters label{{font-size:11px;font-weight:bold;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;display:block;margin-bottom:4px}}
.filters input[type=date]{{padding:7px 9px;border:1px solid #cbd2dc;border-radius:5px;font-size:13px;font-family:inherit}}
.filters button,.tools button{{padding:8px 12px;border:1px solid #cbd2dc;background:#fff;border-radius:5px;cursor:pointer;font-size:12px}}
.filters .chk{{display:flex;align-items:center;gap:6px;font-size:13px;padding-bottom:8px}}
.filters .showing{{margin-left:auto;font-size:12.5px;color:var(--muted);padding-bottom:8px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:22px}}
.kpi{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px 16px}}
.kl{{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);font-weight:bold}}
.kv{{font-size:26px;font-weight:bold;color:var(--navy);margin-top:4px}} .ks{{font-size:11px;color:var(--muted);margin-top:2px}}
.bar{{height:6px;background:#eef1f6;border-radius:3px;margin-top:8px;overflow:hidden}} .bar div{{height:100%;background:var(--green);width:0;transition:width .3s}}
.card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:18px 20px;margin-bottom:22px}}
.card h2{{margin:0 0 4px;font-size:15px;color:var(--navy)}} .card .hint{{font-size:12px;color:var(--muted);margin:0 0 12px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}}
th{{background:var(--navy);color:#fff;text-align:left;padding:8px;white-space:nowrap}}
td{{padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:top}} tbody tr:nth-child(even){{background:#fafbfd}}
.n{{text-align:right;white-space:nowrap}} .tot td{{font-weight:bold;background:#D9E1F2;border-top:2px solid var(--navy)}}
.grp td{{background:#eef1f6;font-weight:bold;color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.03em}}
a{{color:var(--blue);text-decoration:none}} a:hover{{text-decoration:underline}}
.pill{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:bold}}
.a2{{background:#FFF4D6;color:var(--amber)}} .a3{{background:#FDE7E7;color:var(--red)}} .ok{{background:#E6F4EA;color:var(--green)}}
.legend{{font-size:11.5px;color:var(--muted);margin-top:10px}} .legend .pill{{margin:0 4px 0 10px}}
td.cb,th.cb{{width:34px;text-align:center}} input.arch,input.arch-all{{width:16px;height:16px;cursor:pointer;accent-color:var(--green)}}
tr.ld.done td:not(.cb){{opacity:.45;text-decoration:line-through}} tr.ld.done{{background:#F1F8F2 !important}}
.pcard{{background:#fff;border:1px solid var(--line);border-radius:8px;margin-bottom:10px}}
.pcard.alldone{{border-color:#A5D6A7}} .pcard.alldone .pcount{{background:var(--green)}}
.pcard summary{{cursor:pointer;padding:12px 16px;display:flex;gap:14px;align-items:center;flex-wrap:wrap;list-style:none}}
.pcard summary::-webkit-details-marker{{display:none}}
.pcard summary::before{{content:"\\25B8";color:var(--muted)}} .pcard[open] summary::before{{content:"\\25BE"}}
.pname{{font-weight:bold;color:var(--navy);font-size:14px}} .pcount{{background:var(--navy);color:#fff;border-radius:12px;padding:3px 10px;font-size:12px;font-weight:bold}}
.pdone{{background:#E6F4EA;color:var(--green);border-radius:12px;padding:3px 10px;font-size:12px;font-weight:bold}}
.pold{{font-size:12px;color:var(--muted)}} .bot{{background:#eef1f6;color:var(--muted);border-radius:4px;padding:2px 6px;font-size:11px}}
.pcard .ids{{padding:0 16px 10px;font-size:12px}} .idl{{font-weight:bold;color:var(--muted)}}
.ids code{{background:#f3f5f9;padding:4px 6px;border-radius:4px;line-height:1.8;word-break:break-word}}
.copy{{margin-left:8px;padding:4px 10px;border:1px solid var(--blue);background:#fff;color:var(--blue);border-radius:4px;cursor:pointer;font-size:12px}}
.copy:hover{{background:var(--blue);color:#fff}} .pcard table{{margin:0 0 6px}} .sub{{color:var(--muted);font-size:11px}}
.tools{{display:flex;gap:10px;margin:0 0 12px}} .tools input{{flex:1;padding:8px 10px;border:1px solid #cbd2dc;border-radius:5px;font-size:13px}}
.empty{{text-align:center;color:var(--muted);padding:24px;font-size:13px;display:none}}
@media print{{.tools,.copy,.filters button,input.arch-all{{display:none}} body{{background:#fff}} .pcard{{break-inside:avoid}}}}
</style></head><body>
<header><h1>Action Needed: Archive Delivered Loads in Ascend</h1>
<p>Source: {esc(source_name)} · Load Status: Delivered only · Only loads delivered {min_days}+ days before the report date</p>
<div class="period">
<div><span>Report generated</span><b>{generated}</b></div>
<div><span>Report covers deliveries from</span><b>{fdate(first)} – {fdate(last)}</b></div>
<div><span>Delivery cutoff</span><b>on or before {fdate(cutoff)}</b></div>
<div><span>Loads in report</span><b>{n}</b></div>
</div></header>
<main>
<div class="ask"><h2>Why this matters</h2>
Loads that are <b>Delivered but not archived</b> in Ascend are <b>left out of the Savings report</b>.
Until they are archived, the savings from these loads won't be counted.
<ol><li>Find your name below and open your list (or use <b>Copy IDs</b>).</li>
<li>Archive each Load ID in Ascend, then <b>tick the box next to the Load ID</b> &mdash; the counts update as you go.</li>
<li>Start with the oldest loads (shown in red), then reply when you're done.</li></ol></div>

<section class="filters">
<div><label for="from">Delivered from</label><input type="date" id="from" value="{iso(first)}" min="{iso(first)}" max="{iso(last)}"></div>
<div><label for="to">Delivered to</label><input type="date" id="to" value="{iso(last)}" min="{iso(first)}" max="{iso(last)}"></div>
<button type="button" id="reset">Reset dates</button>
<div class="chk"><input type="checkbox" id="hide"><label for="hide" style="margin:0;text-transform:none;font-size:13px;color:inherit;font-weight:normal">Hide archived loads</label></div>
<div class="showing" id="showing"></div>
</section>

<section class="kpis">
<div class="kpi"><div class="kl">Loads Left to Archive</div><div class="kv" id="k-rem">{n}</div><div class="ks" id="k-rem-s"></div><div class="bar"><div id="k-bar"></div></div></div>
<div class="kpi"><div class="kl">Archived (ticked)</div><div class="kv" id="k-arch">0</div><div class="ks" id="k-arch-s"></div></div>
<div class="kpi"><div class="kl">Remaining &gt; 14 Days Old</div><div class="kv" id="k-old14">0</div><div class="ks">highest priority</div></div>
<div class="kpi"><div class="kl">Oldest Remaining Delivery</div><div class="kv" id="k-oldest">{fdate(first)}</div><div class="ks" id="k-oldest-s"></div></div>
</section>

<section class="card"><h2>Loads to Archive by Load Planner</h2>
<p class="hint">Counts follow the date filter and the ticked boxes. Click a name to jump to that planner's Load IDs.</p>
<table id="sum"><thead><tr><th>Load Planner</th><th class=n>Left to Archive</th><th class=n>Archived</th><th class=n>Oldest Remaining</th>
<th class=n>Oldest Age</th><th class=n>Remaining &gt; 14 Days</th><th class=n>Status</th></tr></thead>
<tbody></tbody></table>
<div class="empty" id="sum-empty">No loads in the selected date range.</div>
<p class="legend">Days since delivery:<span class="pill a2">{min_days}–14</span><span class="pill a3">15+</span></p>
</section>

<section class="card"><h2>Load IDs by Load Planner</h2>
<p class="hint">Loads are listed oldest first. Tick the box next to a Load ID once it's archived in Ascend (the header box ticks every load shown).
Ticks are saved in this browser only.</p>
<div class="tools"><input id="q" placeholder="Search a Load ID, planner, location or carrier…">
<button type="button" id="exp">Expand all</button><button type="button" id="col">Collapse all</button></div>
{cards}
<div class="empty" id="cards-empty">No loads match the current filters.</div>
</section>
</main>
<script>
const PLANNERS={planners_json};
const KEY="{report_key}";
const rows=[...document.querySelectorAll('tr.ld')];
const cards=[...document.querySelectorAll('.pcard')];
const $=id=>document.getElementById(id);
const fmt=s=>s?s.slice(5,7)+'/'+s.slice(8,10)+'/'+s.slice(0,4):'';
const esc=s=>s.replace(/[&<>"]/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}}[c]));
const ageCls=d=>d>14?'a3':'a2';

// remember ticks in this browser (if allowed)
let saved={{}}; try{{saved=JSON.parse(localStorage.getItem(KEY)||'{{}}');}}catch(e){{}}
function save(){{try{{const o={{}};rows.forEach(r=>{{if(r.querySelector('.arch').checked)o[r.dataset.id]=1;}});localStorage.setItem(KEY,JSON.stringify(o));}}catch(e){{}}}}
rows.forEach(r=>{{const c=r.querySelector('.arch');c.checked=!!saved[r.dataset.id];}});

function inRange(r){{const f=$('from').value,t=$('to').value,d=r.dataset.date;return (!f||d>=f)&&(!t||d<=t);}}

function update(){{
  const q=$('q').value.toLowerCase(), hide=$('hide').checked;
  const stats={{}}; PLANNERS.forEach(p=>stats[p.name]={{tot:0,arch:0,old14:0,oldest:'',maxDays:0,ids:[]}});
  let tot=0,arch=0,old14=0,oldest='',oldDays=0;
  rows.forEach(r=>{{
    const done=r.querySelector('.arch').checked; r.classList.toggle('done',done);
    const ir=inRange(r), s=stats[r.dataset.p], days=+r.dataset.days;
    if(ir){{
      tot++; s.tot++;
      if(done){{arch++; s.arch++;}}
      else{{s.ids.push(r.dataset.id); if(days>14){{old14++; s.old14++;}}
        if(!s.oldest||r.dataset.date<s.oldest){{s.oldest=r.dataset.date; s.maxDays=days;}}
        if(!oldest||r.dataset.date<oldest){{oldest=r.dataset.date; oldDays=days;}} }}
    }}
    const pm=r.dataset.p.toLowerCase().includes(q);
    r.style.display=(ir&&(!q||pm||r.textContent.toLowerCase().includes(q))&&!(hide&&done))?'':'none';
  }});
  const rem=tot-arch;
  $('k-rem').textContent=rem; $('k-rem-s').textContent=`of ${{tot}} loads in the selected range`;
  $('k-bar').style.width=(tot?arch/tot*100:0)+'%';
  $('k-arch').textContent=arch; $('k-arch-s').textContent=tot?`${{Math.round(arch/tot*100)}}% complete`:'';
  $('k-old14').textContent=old14;
  $('k-oldest').textContent=oldest?fmt(oldest):'—'; $('k-oldest-s').textContent=oldest?`${{oldDays}} days ago`:'all archived';
  $('showing').innerHTML=`Showing deliveries <b>${{fmt($('from').value)||'start'}}</b> – <b>${{fmt($('to').value)||'end'}}</b>`;

  // planner summary table
  let html='',groupShown=false,T={{tot:0,arch:0,old14:0}};
  PLANNERS.forEach(p=>{{const s=stats[p.name]; if(!s.tot)return;
    if(p.bot&&!groupShown){{html+='<tr class="grp"><td colspan=7>Bot-planned / unassigned — needs an owner to archive</td></tr>';groupShown=true;}}
    const left=s.tot-s.arch; T.tot+=s.tot;T.arch+=s.arch;T.old14+=s.old14;
    html+=`<tr><td><a href="#${{p.anchor}}">${{esc(p.name)}}</a></td><td class=n><b>${{left}}</b></td><td class=n>${{s.arch}}</td>`+
      `<td class=n>${{fmt(s.oldest)||'—'}}</td><td class=n>${{left?`<span class="pill ${{ageCls(s.maxDays)}}">${{s.maxDays}} days</span>`:''}}</td>`+
      `<td class=n>${{s.old14}}</td><td class=n>${{left?`${{Math.round(s.arch/s.tot*100)}}%`:'<span class="pill ok">✓ Done</span>'}}</td></tr>`;
  }});
  if(T.tot)html+=`<tr class="tot"><td>Total</td><td class=n>${{T.tot-T.arch}}</td><td class=n>${{T.arch}}</td><td class=n>${{oldest?fmt(oldest):'—'}}</td><td></td><td class=n>${{T.old14}}</td><td class=n>${{Math.round(T.arch/T.tot*100)}}%</td></tr>`;
  document.querySelector('#sum tbody').innerHTML=html; $('sum-empty').style.display=T.tot?'none':'block';
  document.querySelectorAll('#sum a').forEach(a=>a.onclick=()=>{{const c=document.querySelector(a.getAttribute('href'));if(c)c.open=true;}});

  // planner cards
  let anyCard=false;
  cards.forEach(c=>{{const s=stats[c.dataset.p], vis=[...c.querySelectorAll('tr.ld')].some(r=>r.style.display!=='none');
    c.style.display=vis?'':'none'; if(vis)anyCard=true;
    c.querySelector('.c-rem').textContent=s.tot-s.arch; c.querySelector('.c-arch').textContent=s.arch;
    c.querySelector('.c-old').textContent=fmt(s.oldest)||'—';
    c.querySelector('.idlist').textContent=s.ids.join(', ')||'All archived';
    c.classList.toggle('alldone',s.tot>0&&s.arch===s.tot);
    const shown=[...c.querySelectorAll('tr.ld')].filter(r=>r.style.display!=='none');
    const all=c.querySelector('.arch-all'); all.checked=shown.length>0&&shown.every(r=>r.querySelector('.arch').checked);
    if(q&&vis)c.open=true;
  }});
  $('cards-empty').style.display=anyCard?'none':'block';
}}

document.addEventListener('change',e=>{{
  if(e.target.classList.contains('arch')){{save();update();}}
  else if(e.target.classList.contains('arch-all')){{
    const c=e.target.closest('.pcard');
    c.querySelectorAll('tr.ld').forEach(r=>{{if(r.style.display!=='none')r.querySelector('.arch').checked=e.target.checked;}});
    save();update();}}
}});
['from','to'].forEach(id=>$(id).addEventListener('change',update));
$('hide').addEventListener('change',update); $('q').addEventListener('input',update);
$('reset').onclick=()=>{{$('from').value=$('from').min;$('to').value=$('to').max;update();}};
$('exp').onclick=()=>cards.forEach(c=>c.open=true);
$('col').onclick=()=>cards.forEach(c=>c.open=false);
function fallback(t,cb){{const a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();
 try{{document.execCommand('copy');cb();}}catch(e){{}}a.remove();}}
document.querySelectorAll('.copy').forEach(b=>b.addEventListener('click',e=>{{e.preventDefault();
 const t=b.closest('.pcard').querySelector('.idlist').textContent;
 const done=()=>{{b.textContent='Copied!';setTimeout(()=>b.textContent='Copy IDs',1500);}};
 if(navigator.clipboard)navigator.clipboard.writeText(t).then(done,()=>fallback(t,done));else fallback(t,done);}}));
update();
</script></body></html>"""


def build_report(raw_df, source_name, as_of=None, min_days=DEFAULT_MIN_DAYS):
    """Main entry point (used by the XDock dashboard): raw export -> (html, delivered_df, planner_summary)."""
    as_of = as_of or datetime.now()
    d = prepare(raw_df, as_of, min_days)
    if d.empty:
        raise ValueError(f"No Delivered loads with a delivery date on or before "
                         f"{fdate(cutoff_date(as_of, min_days))} were found in this file.")
    g = summarize(d)
    return build_html(d, g, source_name, as_of, min_days), d, g


def main():
    ap = argparse.ArgumentParser(description="Build an HTML action report of Delivered loads to archive in Ascend.")
    ap.add_argument("export", help="Path to the Ascend loads export (.csv or .xlsx)")
    ap.add_argument("--min-days", type=int, default=DEFAULT_MIN_DAYS,
                    help="Only include loads delivered at least this many days ago (default 7)")
    ap.add_argument("-o", "--output", help="Output HTML path (default: Loads_To_Archive.html next to the export)")
    args = ap.parse_args()

    src = Path(args.export)
    out = Path(args.output) if args.output else src.with_name("Loads_To_Archive.html")
    html_text, d, g = build_report(read_export(src), src.name, min_days=args.min_days)
    out.write_text(html_text, encoding="utf-8")
    print(f"Delivered loads to archive: {len(d)}  ->  {out.resolve()}")
    for p, r in g.iterrows():
        print(f"  {p:<32} {int(r.Loads):>4}")


if __name__ == "__main__":
    main()
