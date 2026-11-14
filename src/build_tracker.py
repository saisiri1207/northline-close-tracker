"""Build the Northline month end close tracker workbook (fictional sample data).

Yellow cells are inputs, black text cells are formulas. No macros.
"""
import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="1F3B5C")
WHITE_B = Font(bold=True, color="FFFFFF")
BLACK = Font(color="000000")
BLUE = Font(color="1F4E79")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
AREAS = ["Cash", "Accounts payable", "Revenue", "Payroll", "Fixed assets",
         "Accruals and prepaids", "Intercompany", "Reconciliations", "Reporting"]


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.fill, c.font, c.border = HEAD, WHITE_B, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w


def main():
    tasks = list(csv.DictReader(open(ROOT / "data" / "close_tasks.csv")))
    n = len(tasks)
    last = 4 + n  # tasks start on row 5
    wb = Workbook()

    # ---------- Settings ----------
    s = wb.active
    s.title = "Settings"
    s["A1"] = "Northline month end close tracker (fictional company, sample data)"
    s["A1"].font = Font(bold=True, size=14)
    rows = [("Close month", "October 2026"), ("Prior month end date", "2026-10-31"),
            ("Status as of workday", 4)]
    for i, (k, v) in enumerate(rows, 3):
        s.cell(row=i, column=1, value=k)
        c = s.cell(row=i, column=2, value=v)
        c.fill, c.border = YELLOW, BOX
    from datetime import date
    s["B4"] = date(2026, 10, 31); s["B4"].number_format = "yyyy-mm-dd"
    s["A7"] = "Company holidays (excluded from workdays)"; s["A7"].font = Font(bold=True)
    for i, (d, name) in enumerate([(date(2026, 11, 11), "Veterans Day"), (date(2026, 11, 26), "Thanksgiving")], 8):
        c = s.cell(row=i, column=2, value=d); c.number_format = "yyyy-mm-dd"; c.fill, c.border = YELLOW, BOX
        s.cell(row=i, column=1, value=name)
    s["A12"] = "Legend"; s["A12"].font = Font(bold=True)
    s["A13"] = "Yellow fill"; s["B13"] = "Input you can change"; s["A13"].fill = YELLOW
    s["A14"] = "Black text, no fill"; s["B14"] = "Formula, do not overwrite"
    s.column_dimensions["A"].width = 38; s.column_dimensions["B"].width = 22

    # ---------- Tasks ----------
    t = wb.create_sheet("Tasks")
    t["A1"] = "Close task list"; t["A1"].font = Font(bold=True, size=13)
    t["A2"] = "Inputs in yellow. Days late, due date, dependency check and flag are formulas."
    cols = ["Task ID", "Area", "Task", "Owner", "Due workday", "Depends on", "Status", "Completed workday",
            "Due date", "Days late", "Dependency status", "Flag"]
    header(t, 4, cols, [9, 22, 52, 14, 10, 11, 13, 12, 12, 9, 15, 22])
    dv = DataValidation(type="list", formula1='"Not started,In progress,Complete"', allow_blank=False)
    t.add_data_validation(dv)
    for r, task in enumerate(tasks, 5):
        vals = [task["task_id"], task["area"], task["task"], task["owner"], int(task["due_wd"]),
                task["depends_on"] or None, task["status"], int(task["completed_wd"]) if task["completed_wd"] else None]
        for col, v in enumerate(vals, 1):
            c = t.cell(row=r, column=col, value=v)
            c.fill, c.border = YELLOW, BOX
        dv.add(f"G{r}")
        f = {
            "I": f"=WORKDAY(Settings!$B$4,E{r},Settings!$B$8:$B$9)",
            "J": f'=IF(G{r}="Complete",MAX(0,H{r}-E{r}),MAX(0,Settings!$B$5-E{r}))',
            "K": f'=IF(F{r}="","None",IFERROR(INDEX($G$5:$G${last},MATCH(F{r},$A$5:$A${last},0)),"Missing ID"))',
            "L": (f'=IF(AND(G{r}<>"Complete",K{r}<>"None",K{r}<>"Complete"),"Blocked by "&F{r},'
                  f'IF(AND(G{r}<>"Complete",J{r}>0),"Overdue",IF(AND(G{r}="Complete",J{r}>0),"Done late","On track")))'),
        }
        for col, formula in f.items():
            c = t[f"{col}{r}"]; c.value = formula; c.font, c.border = BLACK, BOX
        t[f"I{r}"].number_format = "ddd mmm d"
    t.freeze_panes = "A5"
    t.auto_filter.ref = f"A4:L{last}"
    red = PatternFill("solid", fgColor="F8CBAD"); green = PatternFill("solid", fgColor="C6EFCE"); amber = PatternFill("solid", fgColor="FFE699")
    t.conditional_formatting.add(f"L5:L{last}", FormulaRule(formula=[f'LEFT(L5,7)="Blocked"'], fill=red))
    t.conditional_formatting.add(f"L5:L{last}", FormulaRule(formula=[f'L5="Overdue"'], fill=red))
    t.conditional_formatting.add(f"L5:L{last}", FormulaRule(formula=[f'L5="Done late"'], fill=amber))
    t.conditional_formatting.add(f"L5:L{last}", FormulaRule(formula=[f'L5="On track"'], fill=green))

    rng = lambda c: f"Tasks!${c}$5:${c}${last}"

    # ---------- Summary by area ----------
    a = wb.create_sheet("Summary")
    a["A1"] = "Completion by area"; a["A1"].font = Font(bold=True, size=13)
    a["A2"] = '="Status as of workday "&Settings!B5&" for "&Settings!B3'
    header(a, 4, ["Area", "Tasks", "Complete", "In progress", "Not started", "Completion %",
                  "Done late", "Open and overdue", "Blocked", "Total days late"], [24, 8, 10, 11, 11, 12, 10, 15, 9, 13])
    for i, area in enumerate(AREAS, 5):
        a.cell(row=i, column=1, value=area).border = BOX
        fs = [f'=COUNTIF({rng("B")},A{i})',
              f'=COUNTIFS({rng("B")},A{i},{rng("G")},"Complete")',
              f'=COUNTIFS({rng("B")},A{i},{rng("G")},"In progress")',
              f'=COUNTIFS({rng("B")},A{i},{rng("G")},"Not started")',
              f'=IF(B{i}=0,0,C{i}/B{i})',
              f'=COUNTIFS({rng("B")},A{i},{rng("L")},"Done late")',
              f'=COUNTIFS({rng("B")},A{i},{rng("L")},"Overdue")',
              f'=COUNTIFS({rng("B")},A{i},{rng("L")},"Blocked*")',
              f'=SUMIF({rng("B")},A{i},{rng("J")})']
        for j, fx in enumerate(fs, 2):
            c = a.cell(row=i, column=j, value=fx); c.border = BOX
        a[f"F{i}"].number_format = "0.0%"
    tr = 5 + len(AREAS)
    a.cell(row=tr, column=1, value="Total").font = Font(bold=True)
    for j, col in enumerate("BCDE", 2):
        a.cell(row=tr, column=j, value=f"=SUM({col}5:{col}{tr-1})").font = Font(bold=True)
    a[f"F{tr}"] = f"=IF(B{tr}=0,0,C{tr}/B{tr})"; a[f"F{tr}"].number_format = "0.0%"; a[f"F{tr}"].font = Font(bold=True)
    for col in "GHIJ":
        a[f"{col}{tr}"] = f"=SUM({col}5:{col}{tr-1})"; a[f"{col}{tr}"].font = Font(bold=True)
    a[f"A{tr+2}"] = "On time completion rate (complete tasks finished by due workday)"
    a[f"F{tr+2}"] = f"=IF(C{tr}=0,0,(C{tr}-G{tr})/C{tr})"; a[f"F{tr+2}"].number_format = "0.0%"
    ch = BarChart(); ch.type = "bar"; ch.title = "Completion % by area"; ch.y_axis.numFmt = "0%"
    ch.add_data(Reference(a, min_col=6, min_row=4, max_row=tr - 1), titles_from_data=True)
    ch.set_categories(Reference(a, min_col=1, min_row=5, max_row=tr - 1))
    ch.height, ch.width, ch.legend = 8, 16, None
    a.add_chart(ch, "L4")

    # ---------- Calendar ----------
    k = wb.create_sheet("Calendar")
    k["A1"] = "Close calendar"; k["A1"].font = Font(bold=True, size=13)
    header(k, 3, ["Workday", "Date", "Tasks due", "Completed", "Still open", "Completed on this day",
                  "Cumulative due", "Cumulative completed", "Cumulative %"], [10, 14, 10, 11, 10, 14, 13, 15, 13])
    for i, wd in enumerate(range(1, 6), 4):
        k.cell(row=i, column=1, value=wd)
        vals = [f"=WORKDAY(Settings!$B$4,A{i},Settings!$B$8:$B$9)",
                f'=COUNTIF({rng("E")},A{i})',
                f'=COUNTIFS({rng("E")},A{i},{rng("G")},"Complete")',
                f"=C{i}-D{i}",
                f'=COUNTIFS({rng("H")},A{i},{rng("G")},"Complete")',
                f"=SUM($C$4:C{i})", f"=SUM($F$4:F{i})", f"=IF(G{i}=0,0,H{i}/G{i})"]
        for j, v in enumerate(vals, 2):
            k.cell(row=i, column=j, value=v).border = BOX
        k[f"B{i}"].number_format = "ddd mmm d"; k[f"I{i}"].number_format = "0.0%"
    k["A10"] = "Owner workload"; k["A10"].font = Font(bold=True)
    header(k, 11, ["Owner", "Tasks", "Open", "Days late"])
    for c in "ABCD":
        k[f"{c}11"].fill, k[f"{c}11"].font = HEAD, WHITE_B
    owners = sorted({x["owner"] for x in tasks})
    for i, o in enumerate(owners, 12):
        k.cell(row=i, column=1, value=o)
        k.cell(row=i, column=2, value=f'=COUNTIF({rng("D")},A{i})')
        k.cell(row=i, column=3, value=f'=COUNTIFS({rng("D")},A{i},{rng("G")},"<>Complete")')
        k.cell(row=i, column=4, value=f'=SUMIF({rng("D")},A{i},{rng("J")})')

    wb.calculation.fullCalcOnLoad = True
    out = ROOT / "northline_close_tracker.xlsx"
    wb.save(out)
    print(f"built {out} with {n} tasks")


if __name__ == "__main__":
    main()
