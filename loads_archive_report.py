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
    people = g[~g["IsBot"]]
    bots = g[g["IsBot"]]
    cutoff = cutoff_date(as_of, min_days)
    over14 = int((d["Days"] > 14).sum())
    oldest = d["Delivery Date"].min()

    kpis = [
        ("Loads to Archive", f"{n:,}", f"Delivered on or before {fdate(cutoff)}"),
        ("Load Planners", f"{len(people):,}", f"+ {len(bots)} bot / unassigned groups"),
        ("Delivered > 14 Days Ago", f"{over14:,}", "highest priority"),
        ("Oldest Delivery", fdate(oldest), f"{ndays(d['Days'].max())} days ago"),
    ]
    kpi_html = "".join(
        f'<div class="kpi"><div class="kl">{k}</div><div class="kv">{v}</div><div class="ks">{s}</div></div>'
        for k, v, s in kpis
    )

    def summary_rows(frame):
        out = ""
        for p, r in frame.iterrows():
            out += (
                f'<tr><td><a href="#{r.Anchor}">{esc(p)}</a></td><td class=n><b>{int(r.Loads)}</b></td>'
                f"<td class=n>{fdate(r.Oldest)}</td><td class=n>{fdate(r.Newest)}</td>"
                f'<td class="n"><span class="pill {age_cls(r.MaxDays)}">{ndays(r.MaxDays)} days</span></td>'
                f"<td class=n>{int(r.Over14)}</td>"
                f'<td class=n><input type="checkbox" class="done" aria-label="Archived"></td></tr>'
            )
        return out

    summary = summary_rows(people)
    if len(bots):
        summary += '<tr class="grp"><td colspan=7>Bot-planned / unassigned — needs an owner to archive</td></tr>'
        summary += summary_rows(bots)
    summary += (
        f'<tr class="tot"><td>Total</td><td class=n>{n}</td><td class=n>{fdate(oldest)}</td>'
        f'<td class=n>{fdate(d["Delivery Date"].max())}</td><td></td><td class=n>{over14}</td><td></td></tr>'
    )

    cards = ""
    for p, r in g.iterrows():
        sub = d[d["Load Planner"] == p].sort_values("Delivery Date")
        ids = ", ".join(sub["Load ID"])
        trs = "".join(
            f'<tr><td><b>{esc(x["Load ID"])}</b></td><td>{fdate(x["Delivery Date"])}</td>'
            f'<td class=n><span class="pill {age_cls(x["Days"])}">{ndays(x["Days"])}</span></td>'
            f'<td>{esc(x["First Pick Name"])}</td><td>{esc(x["Last Drop Name"])}'
            f'<span class=sub> · {esc(x["Last Drop City"])}, {esc(x["Last Drop State"])}</span></td>'
            f'<td>{esc(x["Carrier"])}</td></tr>'
            for _, x in sub.iterrows()
        )
        tag = '<span class="bot">Bot / unassigned</span>' if r.IsBot else ""
        plural = "s" if r.Loads != 1 else ""
        cards += f"""
<details class="pcard" id="{r.Anchor}" data-p="{esc(p)}">
<summary><span class="pname">{esc(p)}</span> {tag}
<span class="pcount">{int(r.Loads)} load{plural} to archive</span>
<span class="pold">oldest delivery: {fdate(r.Oldest)}</span></summary>
<div class="ids"><span class="idl">Load IDs:</span> <code>{ids}</code>
<button type="button" class="copy" data-ids="{esc(ids)}">Copy IDs</button></div>
<table><thead><tr><th>Load ID</th><th>Delivered</th><th class=n>Days</th><th>Pickup</th><th>Delivery</th><th>Carrier</th></tr></thead>
<tbody>{trs}</tbody></table></details>"""

    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Delivered Loads to Archive in Ascend</title>
<style>
:root{{--navy:#1F3864;--blue:#2E75B6;--red:#C00000;--amber:#B7791F;--green:#2F7D32;--line:#e3e7ee;--muted:#6b7280}}
*{{box-sizing:border-box}} body{{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f5f7fa;color:#1f2937}}
header{{background:var(--navy);color:#fff;padding:22px 32px}} header h1{{margin:0;font-size:22px}}
header p{{margin:6px 0 0;font-size:12px;opacity:.85}} main{{padding:24px 32px;max-width:1300px;margin:auto}}
.ask{{background:#FFF8E1;border:1px solid #F2D27A;border-left:6px solid #E0A800;border-radius:8px;padding:16px 20px;margin-bottom:22px;font-size:14px;line-height:1.5}}
.ask h2{{margin:0 0 6px;font-size:16px;color:#7A5200}} .ask ol{{margin:6px 0 0 18px;padding:0}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:22px}}
.kpi{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px 16px}}
.kl{{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);font-weight:bold}}
.kv{{font-size:26px;font-weight:bold;color:var(--navy);margin-top:4px}} .ks{{font-size:11px;color:var(--muted);margin-top:2px}}
.card{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:18px 20px;margin-bottom:22px}}
.card h2{{margin:0 0 4px;font-size:15px;color:var(--navy)}} .card .hint{{font-size:12px;color:var(--muted);margin:0 0 12px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}}
th{{background:var(--navy);color:#fff;text-align:left;padding:8px;white-space:nowrap}}
td{{padding:7px 8px;border-bottom:1px solid var(--line);vertical-align:top}} tbody tr:nth-child(even){{background:#fafbfd}}
.n{{text-align:right;white-space:nowrap}} .tot td{{font-weight:bold;background:#D9E1F2;border-top:2px solid var(--navy)}}
.grp td{{background:#eef1f6;font-weight:bold;color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.03em}}
a{{color:var(--blue);text-decoration:none}} a:hover{{text-decoration:underline}}
.pill{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:bold}}
.a1{{background:#E6F4EA;color:var(--green)}} .a2{{background:#FFF4D6;color:var(--amber)}} .a3{{background:#FDE7E7;color:var(--red)}}
.legend{{font-size:11.5px;color:var(--muted);margin-top:10px}} .legend .pill{{margin:0 4px 0 10px}}
tr.isdone td{{opacity:.45;text-decoration:line-through}}
.pcard{{background:#fff;border:1px solid var(--line);border-radius:8px;margin-bottom:10px}}
.pcard summary{{cursor:pointer;padding:12px 16px;display:flex;gap:14px;align-items:center;flex-wrap:wrap;list-style:none}}
.pcard summary::-webkit-details-marker{{display:none}}
.pcard summary::before{{content:"\\25B8";color:var(--muted)}} .pcard[open] summary::before{{content:"\\25BE"}}
.pname{{font-weight:bold;color:var(--navy);font-size:14px}} .pcount{{background:var(--navy);color:#fff;border-radius:12px;padding:3px 10px;font-size:12px;font-weight:bold}}
.pold{{font-size:12px;color:var(--muted)}} .bot{{background:#eef1f6;color:var(--muted);border-radius:4px;padding:2px 6px;font-size:11px}}
.pcard .ids{{padding:0 16px 10px;font-size:12px}} .idl{{font-weight:bold;color:var(--muted)}}
.ids code{{background:#f3f5f9;padding:4px 6px;border-radius:4px;line-height:1.8;word-break:break-word}}
.copy{{margin-left:8px;padding:4px 10px;border:1px solid var(--blue);background:#fff;color:var(--blue);border-radius:4px;cursor:pointer;font-size:12px}}
.copy:hover{{background:var(--blue);color:#fff}} .pcard table{{margin:0 0 6px}} .sub{{color:var(--muted);font-size:11px}}
.tools{{display:flex;gap:10px;margin:0 0 12px}} .tools input{{flex:1;padding:8px 10px;border:1px solid #cbd2dc;border-radius:5px;font-size:13px}}
.tools button{{padding:8px 12px;border:1px solid #cbd2dc;background:#fff;border-radius:5px;cursor:pointer;font-size:12px}}
@media print{{.tools,.copy,.done{{display:none}} body{{background:#fff}} .pcard{{break-inside:avoid}}}}
</style></head><body>
<header><h1>Action Needed: Archive Delivered Loads in Ascend</h1>
<p>As of {as_of.strftime("%m/%d/%Y")} · Source: {esc(source_name)} · Load Status: Delivered only · Delivered on or before {fdate(cutoff)} ({min_days}+ days ago) · {n} loads still unarchived</p></header>
<main>
<div class="ask"><h2>Why this matters</h2>
Loads that are <b>Delivered but not archived</b> in Ascend are <b>left out of the Savings report</b>.
Until they are archived, the savings from these {n} loads won't be counted.
<ol><li>Find your name below and open your list (or use <b>Copy IDs</b>).</li>
<li>Archive each Load ID in Ascend.</li>
<li>Start with the oldest loads (shown in red), then reply when you're done.</li></ol></div>

<section class="kpis">{kpi_html}</section>

<section class="card"><h2>Loads to Archive by Load Planner</h2>
<p class="hint">Click a name to jump to that planner's Load IDs. Tick the box as each planner confirms they're done.</p>
<table id="sum"><thead><tr><th>Load Planner</th><th class=n>Loads to Archive</th><th class=n>Oldest Delivery</th><th class=n>Newest Delivery</th>
<th class=n>Oldest Age</th><th class=n>Delivered &gt; 14 Days</th><th class=n>Done</th></tr></thead>
<tbody>{summary}</tbody></table>
<p class="legend">Days since delivery:<span class="pill a2">{min_days}–14</span><span class="pill a3">15+</span></p>
</section>

<section class="card"><h2>Load IDs by Load Planner</h2>
<p class="hint">Loads are listed oldest first.</p>
<div class="tools"><input id="q" placeholder="Search a Load ID, planner, location or carrier…">
<button type="button" id="exp">Expand all</button><button type="button" id="col">Collapse all</button></div>
{cards}
</section>
</main>
<script>
function fallback(t,cb){{const a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();
 try{{document.execCommand('copy');cb();}}catch(e){{}}a.remove();}}
document.querySelectorAll('.copy').forEach(b=>b.addEventListener('click',e=>{{e.preventDefault();
 const t=b.dataset.ids;const done=()=>{{b.textContent='Copied!';setTimeout(()=>b.textContent='Copy IDs',1500);}};
 if(navigator.clipboard)navigator.clipboard.writeText(t).then(done,()=>fallback(t,done));else fallback(t,done);}}));
document.querySelectorAll('.done').forEach(c=>c.addEventListener('change',()=>c.closest('tr').classList.toggle('isdone',c.checked)));
const cards=[...document.querySelectorAll('.pcard')];
document.getElementById('exp').onclick=()=>cards.forEach(c=>c.open=true);
document.getElementById('col').onclick=()=>cards.forEach(c=>c.open=false);
document.querySelectorAll('#sum a').forEach(a=>a.addEventListener('click',()=>{{const c=document.querySelector(a.getAttribute('href'));if(c)c.open=true;}}));
document.getElementById('q').addEventListener('input',e=>{{const t=e.target.value.toLowerCase();
 cards.forEach(c=>{{let any=false;const pm=c.dataset.p.toLowerCase().includes(t);
  c.querySelectorAll('tbody tr').forEach(r=>{{const ok=!t||pm||r.textContent.toLowerCase().includes(t);r.style.display=ok?'':'none';if(ok)any=true;}});
  c.style.display=any?'':'none';c.open=!!(t&&any);}});}});
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
