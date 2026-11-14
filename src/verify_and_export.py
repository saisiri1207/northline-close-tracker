"""Recalculate the workbook with LibreOffice headless, check formulas, and export a status page.

Writes outputs/close_status.html, outputs/completion_by_area.png and outputs/summary.csv.
"""
import shutil
import subprocess
import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"


def recalc() -> Path:
    tmp = Path(tempfile.mkdtemp())
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("LibreOffice not found; open the workbook in Excel to recalculate instead.")
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp),
                    str(ROOT / "northline_close_tracker.xlsx")], check=True, capture_output=True)
    return tmp / "northline_close_tracker.xlsx"


def main():
    wb = load_workbook(recalc(), data_only=True)
    errors = [f"{ws.title}!{c.coordinate}={c.value}" for ws in wb for row in ws.iter_rows() for c in row
              if isinstance(c.value, str) and c.value.startswith("#")]
    t, s, k = wb["Tasks"], wb["Summary"], wb["Calendar"]
    tasks = [r for r in t.iter_rows(min_row=5, values_only=True) if r[0]]
    areas = [r for r in s.iter_rows(min_row=5, max_row=13, values_only=True)]
    total = next(r for r in s.iter_rows(min_row=14, max_row=14, values_only=True))
    on_time = s["F16"].value
    cal = [r for r in k.iter_rows(min_row=4, max_row=8, values_only=True)]
    asof = wb["Settings"]["B5"].value

    summary = {"tasks": total[1], "complete": total[2], "in_progress": total[3], "not_started": total[4],
               "completion_pct": round(total[5], 4), "done_late": total[6], "open_overdue": total[7],
               "blocked": total[8], "total_days_late": total[9], "on_time_completion_pct": round(on_time, 4),
               "status_as_of_workday": asof, "formula_errors": len(errors)}
    with open(OUT / "summary.csv", "w") as f:
        f.write("metric,value\n" + "".join(f"{a},{b}\n" for a, b in summary.items()))

    fig, ax = plt.subplots(figsize=(8, 4.5))
    names = [a[0] for a in areas][::-1]; vals = [a[5] * 100 for a in areas][::-1]
    ax.barh(names, vals, color=["#2a9d8f" if v == 100 else "#e9c46a" if v >= 50 else "#e76f51" for v in vals])
    for y, v in enumerate(vals):
        ax.text(v + 1, y, f"{v:.0f}%", va="center")
    ax.set_xlim(0, 112); ax.set_xlabel("Tasks complete (%)")
    ax.set_title(f"Northline October close, completion by area as of workday {asof}")
    fig.tight_layout(); fig.savefig(OUT / "completion_by_area.png", dpi=150); plt.close(fig)

    color = {"On track": "#c6efce", "Done late": "#ffe699", "Overdue": "#f8cbad"}
    rows = "".join(
        f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>WD{r[4]}</td><td>{r[6]}</td>"
        f"<td>{r[9]}</td><td style='background:{color.get(r[11], '#f8cbad')}'>{r[11]}</td></tr>" for r in tasks)
    arows = "".join(f"<tr><td>{a[0]}</td><td>{a[1]}</td><td>{a[2]}</td><td>{a[5]*100:.1f}%</td><td>{a[7]}</td><td>{a[8]}</td></tr>" for a in areas)
    crows = "".join(f"<tr><td>WD{c[0]}</td><td>{c[1]:%a %b %d}</td><td>{c[2]}</td><td>{c[3]}</td><td>{c[4]}</td><td>{c[8]*100:.1f}%</td></tr>" for c in cal)
    html = f"""<!doctype html><html><head><meta charset="utf-8"><title>Northline close status</title>
<style>body{{font-family:Arial,sans-serif;margin:24px;color:#1f3b5c}}table{{border-collapse:collapse;margin:12px 0 24px}}
td,th{{border:1px solid #ccc;padding:4px 8px;font-size:13px}}th{{background:#1f3b5c;color:#fff}}.kpi{{display:inline-block;margin:0 24px 12px 0;font-size:15px}}</style></head><body>
<h1>Northline October 2026 close status, workday {asof}</h1><p>Fictional company, sample data. Generated from the recalculated Excel tracker.</p>
<div class="kpi"><b>{summary['completion_pct']*100:.1f}%</b> complete ({total[2]} of {total[1]})</div>
<div class="kpi"><b>{summary['blocked']}</b> blocked</div><div class="kpi"><b>{summary['open_overdue']}</b> open and overdue</div>
<div class="kpi"><b>{summary['on_time_completion_pct']*100:.1f}%</b> on time completion</div>
<h2>By area</h2><table><tr><th>Area</th><th>Tasks</th><th>Complete</th><th>Completion %</th><th>Overdue</th><th>Blocked</th></tr>{arows}</table>
<h2>Close calendar</h2><table><tr><th>Workday</th><th>Date</th><th>Due</th><th>Completed</th><th>Open</th><th>Cumulative %</th></tr>{crows}</table>
<h2>Tasks</h2><table><tr><th>ID</th><th>Area</th><th>Task</th><th>Owner</th><th>Due</th><th>Status</th><th>Days late</th><th>Flag</th></tr>{rows}</table>
</body></html>"""
    (OUT / "close_status.html").write_text(html)
    print(summary)
    if errors:
        raise SystemExit(f"formula errors: {errors}")


if __name__ == "__main__":
    main()
