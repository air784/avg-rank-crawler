"""根据 data/ 下的 CSV 生成 Excel 到 excel/ 目录：python make_excel.py [日期|today|all]

- 日报：excel/榜单_日期.xlsx（付费榜当日最终、当日增长、趋势查询、各项变化和明细）
- 总集：excel/总集.xlsx（每天最后一版付费榜 + 每天人气/收藏/点赞及日增，含趋势查询），每次运行都会重建
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

from crawl import BJ, BOARDS, METRICS, ROOT, csv_path

EXCEL_DIR = os.path.join(ROOT, "excel")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FONT = Font(color="808080")
NUM_FMT = "#,##0"
METRIC_NAMES = list(METRICS)  # 人气、收藏、点赞
GROWTH_NAMES = [f"{m}日增" for m in METRIC_NAMES]


def load(board, day):
    path = csv_path(board, day)
    if not os.path.exists(path):
        return None
    df = pd.read_csv(path, dtype=str)
    df["排名"] = df["排名"].astype(int)
    df["作品ID"] = df["作品ID"].astype(int)
    for m in METRIC_NAMES:
        # 早期数据没有这几列
        df[m] = pd.to_numeric(df[m], errors="coerce") if m in df else float("nan")
    return df


def load_metrics(day):
    """两个榜合并后每个时间点每部作品一行的人气/收藏/点赞。"""
    dfs = [df for df in (load(b, day) for b in BOARDS) if df is not None]
    if not dfs:
        return None
    df = pd.concat(dfs).dropna(subset=METRIC_NAMES)
    df = df.drop_duplicates(["榜单更新时间", "作品ID"]).sort_values("榜单更新时间")
    return df[["榜单更新时间", "作品名", "作品ID"] + METRIC_NAMES]


def day_end_metrics(day):
    """每部作品当天最后一次出现时的数值，索引为作品ID。"""
    df = load_metrics(day)
    if df is None or df.empty:
        return None
    return df.groupby("作品ID").last()


def prev_day(day):
    return (datetime.strptime(day, "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m-%d")


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def number_format(ws, first_row, cols):
    for col in cols:
        for (cell,) in ws.iter_rows(min_row=first_row, min_col=col, max_col=col):
            cell.number_format = NUM_FMT


def write_table(ws, df, widths=None, title=None):
    """写一张普通表格，可选在最上面加一行说明。大数字列加千分位。"""
    if title:
        ws.append([title])
        ws["A1"].font = Font(bold=True, size=12)
        ws.append([])
    header_row = ws.max_row + 1 if title else 1
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([None if pd.isna(v) else v for v in row])
    style_header(ws, header_row, len(df.columns))
    ws.freeze_panes = f"A{header_row + 1}"
    big = [i + 1 for i, c in enumerate(df.columns) if any(c.startswith(m) for m in METRIC_NAMES)]
    number_format(ws, header_row + 1, big)
    set_widths(ws, widths or [])


def write_matrix(ws, df, col, value, ascending=True):
    """每行一个作品、每列一个时间点（col 列的取值）的矩阵，按最后一列排序。
    返回 (行数, 列数)，含表头。"""
    m = df.pivot_table(index="作品名", columns=col, values=value, aggfunc="first")
    m = m.sort_values(by=list(m.columns[::-1]), ascending=ascending)
    ws.append(["作品名"] + list(m.columns))
    for name, vals in m.iterrows():
        ws.append([name] + [None if pd.isna(v) else int(v) for v in vals])
    style_header(ws, 1, len(m.columns) + 1)
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 28
    if value != "排名":
        number_format(ws, 2, range(2, len(m.columns) + 2))
        for c in range(2, len(m.columns) + 2):
            ws.column_dimensions[get_column_letter(c)].width = 14
    return len(m) + 1, len(m.columns) + 1


def write_query(wb, ws, names, x_label, x_values, columns, charts, default_name):
    """趋势查询：在 B1 输入/选择作品名，下方表格用公式从矩阵表取数，右侧画折线图。
    columns: [(列名, 矩阵表名, 行数, 列数)]
    charts:  [(图表标题, [columns 里的序号], 是否为排名)]"""
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
    ws["A2"] = "在黄色格子里输入作品名，或点右侧小箭头从列表里选；空白表示当时没有数据（未上榜）"
    ws["A2"].font = NOTE_FONT
    dv = DataValidation(type="list", formula1=f"=作品列表!$A$1:$A${len(names)}",
                        showErrorMessage=False)
    ws.add_data_validation(dv)
    dv.add("B1")

    header_row = 4
    for j, h in enumerate([x_label] + [c[0] for c in columns], start=1):
        ws.cell(header_row, j, h)
    style_header(ws, header_row, len(columns) + 1)
    for i, x in enumerate(x_values):
        r = header_row + 1 + i
        ws.cell(r, 1, x)
        for j, (name, sheet, nrows, ncols) in enumerate(columns):
            last_col = get_column_letter(ncols)
            ref = f"'{sheet}'!"
            cell = ws.cell(r, 2 + j, (
                f"=IFERROR(INDEX({ref}$A$1:${last_col}${nrows},"
                f"MATCH($B$1,{ref}$A$1:$A${nrows},0),"
                f"MATCH($A{r},{ref}$A$1:${last_col}$1,0)),NA())"
            ))
            if name != "付费榜" and name != "活跃榜" and "排名" not in name:
                cell.number_format = NUM_FMT
    last_row = header_row + len(x_values)
    ws.column_dimensions["A"].width = 20
    for j in range(len(columns)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 14
    ws.freeze_panes = f"A{header_row + 1}"

    # #N/A 只是给图表留空用的，把它的字隐藏掉
    last_letter = get_column_letter(1 + len(columns))
    ws.conditional_formatting.add(f"B{header_row + 1}:{last_letter}{last_row}",
                                  FormulaRule(formula=[f"ISNA(B{header_row + 1})"],
                                              font=Font(color="FFFFFF")))

    anchor_col = get_column_letter(len(columns) + 3)
    for k, (title, idx, is_rank) in enumerate(charts):
        chart = LineChart()
        chart.title = title
        chart.varyColors = False  # 只有一条线时 Excel 默认会给每个点不同颜色
        if is_rank:
            chart.y_axis.scaling.orientation = "maxMin"  # 第 1 名在最上面
            chart.y_axis.scaling.min = 1
            chart.x_axis.crosses = "max"  # 纵轴倒序后横轴默认会跑到顶部，放回底部
        else:
            chart.y_axis.number_format = NUM_FMT
        chart.y_axis.delete = False
        chart.x_axis.delete = False
        chart.legend.position = "r"
        chart.height, chart.width = 7.5, 24
        for i in idx:
            col = 2 + i
            chart.add_data(Reference(ws, min_col=col, min_row=header_row, max_row=last_row),
                           titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=header_row + 1, max_row=last_row))
        for s in chart.series:
            s.smooth = False
            s.marker.symbol = "circle"
            s.marker.size = 4
        ws.add_chart(chart, f"{anchor_col}{header_row + k * 16}")


def build_daily(day):
    dfs = {b: load(b, day) for b in BOARDS}
    dfs = {b: df for b, df in dfs.items() if df is not None}
    if not dfs:
        print(f"{day} 没有数据，跳过")
        return

    wb = Workbook()
    final_ws = wb.active
    final_ws.title = "付费榜当日最终"
    growth_ws = wb.create_sheet("当日增长")
    query_ws = wb.create_sheet("趋势查询")

    pay = dfs.get("付费榜")
    if pay is not None:
        last = pay["榜单更新时间"].max()
        final = pay[pay["榜单更新时间"] == last][["排名", "作品名", "作品ID"] + METRIC_NAMES]
        write_table(final_ws, final, [8, 28, 12, 16, 14, 14],
                    f"{day} 付费榜最终排名（当天最后一次抓取，榜单时间 {last}）")

    # 当日增长：今天最后一次数值 - 昨天最后一次数值（昨天没数据就用今天第一次）
    metrics = load_metrics(day)
    has_metrics = metrics is not None and not metrics.empty
    if has_metrics:
        end = metrics.groupby("作品ID").last()
        base = day_end_metrics(prev_day(day))
        base_note = "昨天最后一次抓取"
        if base is None:
            base = metrics.groupby("作品ID").first()
            base_note = "今天第一次抓取（没有昨天的数据）"
        g = end[["作品名"]].copy()
        g.insert(1, "作品ID", end.index)
        for m in METRIC_NAMES:
            g[m] = end[m]
            g[f"{m}增长"] = end[m] - base[m].reindex(end.index)
        g = g.sort_values("人气增长", ascending=False)
        write_table(growth_ws, g, [28, 12, 16, 14, 14, 12, 14, 12],
                    f"{day} 人气/收藏/点赞增长 = 今天最后一次抓取 − {base_note}；"
                    "只统计当天上过付费榜或活跃榜的作品")
    else:
        growth_ws["A1"] = "当天的数据还没有人气/收藏/点赞"

    columns, charts = [], []
    for board, df in dfs.items():
        nrows, ncols = write_matrix(wb.create_sheet(f"{board}排名变化"), df, "榜单更新时间", "排名")
        columns.append((board, f"{board}排名变化", nrows, ncols))
    charts.append(("排名变化", list(range(len(columns))), True))
    if has_metrics:
        for m in METRIC_NAMES:
            nrows, ncols = write_matrix(wb.create_sheet(f"{m}变化"), metrics, "榜单更新时间", m,
                                        ascending=False)
            columns.append((m, f"{m}变化", nrows, ncols))
            charts.append((m, [len(columns) - 1], False))
    for board, df in dfs.items():
        write_table(wb.create_sheet(f"{board}明细"), df, [18, 20, 8, 28, 12, 16, 14, 14])

    all_df = pd.concat(dfs.values())
    times = sorted(all_df["榜单更新时间"].unique())
    first = (pay if pay is not None else all_df).sort_values("排名").iloc[0]["作品名"]
    write_query(wb, query_ws, set(all_df["作品名"]), "榜单时间", times, columns, charts, first)

    save(wb, os.path.join(EXCEL_DIR, f"榜单_{day}.xlsx"))


def build_summary(today):
    """每天最后一版付费榜 + 每天人气/收藏/点赞及日增（不含还没结束的今天）。"""
    days = sorted({re.search(r"(\d{4}-\d{2}-\d{2})", p).group(1)
                   for b in BOARDS for p in glob.glob(csv_path(b, "*"))})
    days = [d for d in days if d < today]
    pay_finals, daily = [], []
    for day in days:
        pay = load("付费榜", day)
        if pay is not None:
            last = pay[pay["榜单更新时间"] == pay["榜单更新时间"].max()].copy()
            last.insert(0, "日期", day)
            pay_finals.append(last[["日期", "榜单更新时间", "排名", "作品名", "作品ID"]])
        end = day_end_metrics(day)
        if end is None:
            continue
        prev = day_end_metrics(prev_day(day))
        d = pd.DataFrame({"日期": day, "作品名": end["作品名"], "作品ID": end.index})
        for m in METRIC_NAMES:
            d[m] = end[m]
            d[f"{m}日增"] = end[m] - prev[m].reindex(end.index) if prev is not None else float("nan")
        daily.append(d)
    if not pay_finals and not daily:
        print("还没有完整一天的数据，暂不生成总集")
        return

    wb = Workbook()
    query_ws = wb.active
    query_ws.title = "趋势查询"
    columns, charts, names, dates = [], [], set(), set()
    if pay_finals:
        pf = pd.concat(pay_finals, ignore_index=True)
        nrows, ncols = write_matrix(wb.create_sheet("付费榜每日排名"), pf, "日期", "排名")
        columns.append(("付费榜排名", "付费榜每日排名", nrows, ncols))
        charts.append(("付费榜每日最终排名", [0], True))
        names |= set(pf["作品名"])
        dates |= set(pf["日期"])
    if daily:
        dm = pd.concat(daily, ignore_index=True)
        for g in GROWTH_NAMES:
            nrows, ncols = write_matrix(wb.create_sheet(g), dm.dropna(subset=[g]), "日期", g,
                                        ascending=False)
            columns.append((g, g, nrows, ncols))
            charts.append((f"每日{g[:2]}增长", [len(columns) - 1], False))
        names |= set(dm["作品名"])
        dates |= set(dm["日期"])
        write_table(wb.create_sheet("每日数据"), dm.sort_values(["日期", "人气日增"], ascending=[True, False]),
                    [12, 28, 12, 16, 14, 14, 12, 14, 12],
                    "每天的数值取当天最后一次抓取；日增 = 当天 − 前一天（前一天没数据则为空）")
    if pay_finals:
        write_table(wb.create_sheet("付费榜每日最终"), pf, [12, 18, 8, 28, 12])

    default = (pf[pf["日期"] == pf["日期"].max()].sort_values("排名").iloc[0]["作品名"]
               if pay_finals else sorted(names)[0])
    write_query(wb, query_ws, names, "日期", sorted(dates), columns, charts, default)

    save(wb, os.path.join(EXCEL_DIR, "总集.xlsx"))


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
        days = [prev_day(today)]
    for day in days:
        build_daily(day)
    build_summary(today)


if __name__ == "__main__":
    main()
