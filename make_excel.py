"""根据 data/ 下的 CSV 生成 Excel 到 excel/ 目录：python make_excel.py [日期|today|all]

- 日报：excel/榜单_日期.xlsx（付费榜当日最终、趋势查询、两个榜的排名变化和明细）
- 总集：excel/付费榜总集.xlsx（每天最后一版付费榜，含趋势查询），每次运行都会重建
不带参数时生成昨天（北京时间）的日报。日期格式 2026-10-05。
"""
import glob
import os
import re
import sys
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from crawl import BJ, BOARDS, ROOT, csv_path

EXCEL_DIR = os.path.join(ROOT, "excel")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FONT = Font(color="808080")


def load(board, day):
    path = csv_path(board, day)
    if not os.path.exists(path):
        return None
    df = pd.read_csv(path, dtype=str)
    df["排名"] = df["排名"].astype(int)
    df["作品ID"] = df["作品ID"].astype(int)
    return df


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")


def write_table(ws, df, widths=None):
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))
    style_header(ws, 1, len(df.columns))
    ws.freeze_panes = "A2"
    for i, w in enumerate(widths or [], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_matrix(ws, df, col):
    """每行一个作品、每列一个时间点（col 列的取值）的排名矩阵，按最后一列排名排序。
    返回 (行数, 列数)，含表头。"""
    m = df.pivot_table(index="作品名", columns=col, values="排名", aggfunc="first")
    m = m.sort_values(by=list(m.columns[::-1]))
    ws.append(["作品名"] + list(m.columns))
    for name, ranks in m.iterrows():
        ws.append([name] + [None if pd.isna(v) else int(v) for v in ranks])
    style_header(ws, 1, len(m.columns) + 1)
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 28
    return len(m) + 1, len(m.columns) + 1


def write_query(wb, ws, names, x_label, x_values, series, default_name):
    """趋势查询：在 B1 输入/选择作品名，下方表格用公式从矩阵表取出排名，并画折线图。
    series: [(系列名, 矩阵表名, 行数, 列数)]"""
    # 作品下拉列表放在隐藏的辅助表里
    lst = wb.create_sheet("作品列表")
    for n in sorted(names):
        lst.append([n])
    lst.sheet_state = "hidden"

    ws["A1"] = "作品名称："
    ws["A1"].font = Font(bold=True, size=12)
    ws["B1"] = default_name
    ws["B1"].fill = INPUT_FILL
    ws["B1"].font = Font(bold=True, size=12)
    ws["A2"] = "在黄色格子里输入作品名，或点右侧小箭头从列表里选；排名越靠上越好，空白表示当时未上榜"
    ws["A2"].font = NOTE_FONT
    dv = DataValidation(type="list", formula1=f"=作品列表!$A$1:$A${len(names)}",
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("B1")

    ws.append([])
    header_row = 4
    ws.append([x_label] + [s[0] for s in series])
    style_header(ws, header_row, len(series) + 1)
    for i, x in enumerate(x_values):
        r = header_row + 1 + i
        ws.cell(r, 1, x)
        for j, (_, sheet, nrows, ncols) in enumerate(series):
            last_col = get_column_letter(ncols)
            ref = f"'{sheet}'!"
            ws.cell(r, 2 + j, (
                f"=IFERROR(INDEX({ref}$A$1:${last_col}${nrows},"
                f"MATCH($B$1,{ref}$A$1:$A${nrows},0),"
                f"MATCH($A{r},{ref}$A$1:${last_col}$1,0)),NA())"
            ))
    last_row = header_row + len(x_values)
    ws.column_dimensions["A"].width = 20
    for j in range(len(series)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 12
    ws.freeze_panes = f"A{header_row + 1}"

    # #N/A 只是给图表留空用的，把它的字隐藏掉
    rng = f"B{header_row + 1}:{get_column_letter(1 + len(series))}{last_row}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"ISNA(B{header_row + 1})"],
                                                   font=Font(color="FFFFFF")))

    chart = LineChart()
    chart.title = "排名变化趋势"
    chart.varyColors = False  # 只有一条线时 Excel 默认会给每个点不同颜色
    chart.y_axis.scaling.orientation = "maxMin"  # 第 1 名在最上面
    chart.y_axis.scaling.min = 1
    chart.x_axis.crosses = "max"  # 纵轴倒序后横轴默认会跑到顶部，放回底部
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.legend.position = "r"
    chart.height, chart.width = 10, 24
    data = Reference(ws, min_col=2, max_col=1 + len(series), min_row=header_row, max_row=last_row)
    chart.add_data(data, titles_from_data=True)
    for s in chart.series:
        s.smooth = False
        s.marker.symbol = "circle"
        s.marker.size = 4
    chart.set_categories(Reference(ws, min_col=1, min_row=header_row + 1, max_row=last_row))
    ws.add_chart(chart, f"{get_column_letter(len(series) + 3)}4")


def build_daily(day):
    dfs = {b: load(b, day) for b in BOARDS}
    dfs = {b: df for b, df in dfs.items() if df is not None}
    if not dfs:
        print(f"{day} 没有数据，跳过")
        return

    wb = Workbook()
    final_ws = wb.active
    final_ws.title = "付费榜当日最终"
    query_ws = wb.create_sheet("趋势查询")

    pay = dfs.get("付费榜")
    if pay is not None:
        last = pay["榜单更新时间"].max()
        final = pay[pay["榜单更新时间"] == last][["排名", "作品名", "作品ID"]]
        final_ws.append([f"{day} 付费榜最终排名（当天最后一次抓取，榜单时间 {last}）"])
        final_ws["A1"].font = Font(bold=True, size=12)
        final_ws.append([])
        final_ws.append(list(final.columns))
        for row in final.itertuples(index=False):
            final_ws.append(list(row))
        style_header(final_ws, 3, 3)
        final_ws.freeze_panes = "A4"
        for col, w in zip("ABC", [8, 28, 12]):
            final_ws.column_dimensions[col].width = w

    series = []
    for board, df in dfs.items():
        nrows, ncols = write_matrix(wb.create_sheet(f"{board}排名变化"), df, "榜单更新时间")
        series.append((board, f"{board}排名变化", nrows, ncols))
    for board, df in dfs.items():
        write_table(wb.create_sheet(f"{board}明细"), df, [18, 20, 8, 28, 12])

    all_df = pd.concat(dfs.values())
    times = sorted(all_df["榜单更新时间"].unique())
    first = (pay if pay is not None else all_df).sort_values("排名").iloc[0]["作品名"]
    write_query(wb, query_ws, set(all_df["作品名"]), "榜单时间", times, series, first)

    save(wb, os.path.join(EXCEL_DIR, f"榜单_{day}.xlsx"))


def build_summary(today):
    """每天最后一版付费榜汇总（不含还没结束的今天）。"""
    days = sorted(re.search(r"(\d{4}-\d{2}-\d{2})", p).group(1)
                  for p in glob.glob(csv_path("付费榜", "*")))
    finals = []
    for day in days:
        if day >= today:
            continue
        df = load("付费榜", day)
        last = df[df["榜单更新时间"] == df["榜单更新时间"].max()].copy()
        last.insert(0, "日期", day)
        finals.append(last.drop(columns=["抓取时间"]))
    if not finals:
        print("还没有完整一天的数据，暂不生成总集")
        return
    df = pd.concat(finals, ignore_index=True)

    wb = Workbook()
    query_ws = wb.active
    query_ws.title = "趋势查询"
    nrows, ncols = write_matrix(wb.create_sheet("每日排名"), df, "日期")
    write_table(wb.create_sheet("每日最终明细"), df, [12, 18, 8, 28, 12])
    latest = df[df["日期"] == df["日期"].max()].sort_values("排名").iloc[0]["作品名"]
    write_query(wb, query_ws, set(df["作品名"]), "日期", sorted(df["日期"].unique()),
                [("付费榜", "每日排名", nrows, ncols)], latest)

    save(wb, os.path.join(EXCEL_DIR, "付费榜总集.xlsx"))


def save(wb, path):
    wb.calculation.fullCalcOnLoad = True  # 打开时重新计算公式
    os.makedirs(EXCEL_DIR, exist_ok=True)
    wb.save(path)
    print(f"已生成 {path}")


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] else ""
    today = datetime.now(BJ).strftime("%Y-%m-%d")
    if arg == "all":
        days = sorted({re.search(r"(\d{4}-\d{2}-\d{2})", p).group(1)
                       for b in BOARDS for p in glob.glob(csv_path(b, "*"))})
    elif arg == "today":
        days = [today]
    elif arg:
        days = [arg]
    else:
        days = [(datetime.now(BJ) - timedelta(days=1)).strftime("%Y-%m-%d")]
    for day in days:
        build_daily(day)
    build_summary(today)


if __name__ == "__main__":
    main()
