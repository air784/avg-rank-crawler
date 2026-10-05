"""把某天的 CSV 生成 Excel：python make_excel.py [日期|today|all]

不带参数时生成昨天（北京时间）的。日期格式 2026-10-05。
"""
import glob
import os
import sys
from datetime import datetime, timedelta

import pandas as pd

from crawl import BJ, DATA_DIR, csv_path


def build(day):
    src = csv_path(day)
    if not os.path.exists(src):
        print(f"{day} 没有数据，跳过")
        return
    df = pd.read_csv(src, dtype=str)
    df["排名"] = df["排名"].astype(int)
    # 透视表：每行一个作品，每列一个榜单时间点，方便看排名变化
    pivot = df.pivot_table(index="作品名", columns="榜单更新时间", values="排名", aggfunc="first")
    pivot = pivot.sort_values(by=pivot.columns[-1])
    dst = os.path.join(DATA_DIR, f"付费榜_{day}.xlsx")
    with pd.ExcelWriter(dst) as w:
        df.to_excel(w, sheet_name="明细", index=False)
        pivot.to_excel(w, sheet_name="排名变化")
    print(f"已生成 {dst}")


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1] else ""
    now = datetime.now(BJ)
    if arg == "all":
        days = [os.path.basename(p)[4:14] for p in sorted(glob.glob(csv_path("*")))]
    elif arg == "today":
        days = [now.strftime("%Y-%m-%d")]
    elif arg:
        days = [arg]
    else:
        days = [(now - timedelta(days=1)).strftime("%Y-%m-%d")]
    for day in days:
        build(day)


if __name__ == "__main__":
    main()
