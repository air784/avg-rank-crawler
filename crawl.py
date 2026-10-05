"""易次元付费榜抓取：每次运行抓取一次完整榜单，按天追加到 data/付费榜_日期.csv。

同一个榜单版本只记录一次，所以重复运行不会产生重复数据。
Excel 由 make_excel.py 根据 CSV 生成。
"""
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests

RANK_TYPE = 1  # 1 = 付费榜
API = "https://avg.163.com/avg-portal-api/game/ranking"
PUBLISH_API = "https://avg.163.com/avg-portal-api/game/ranking/publish/time"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
                  "AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148",
    "Referer": "https://avg.163.com/m/rankList",
}
PAGE_SIZE = 20
BJ = timezone(timedelta(hours=8))  # 云端服务器是 UTC，统一用北京时间
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
COLUMNS = ["榜单更新时间", "抓取时间", "排名", "作品名", "作品ID"]


def get(url, params=None):
    for attempt in range(3):
        try:
            r = requests.get(url, params=params, headers=HEADERS, timeout=15)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            if attempt == 2:
                raise
            print(f"请求失败，重试中：{e}")
            time.sleep(3)


def fetch_ranking():
    items = []
    offset = 0
    while True:
        page = get(API, {"type": RANK_TYPE, "offset": offset, "limit": PAGE_SIZE})["data"]
        if not page:
            break
        items += page
        offset += PAGE_SIZE
        time.sleep(0.5)
    return items


def csv_path(day):
    return os.path.join(DATA_DIR, f"付费榜_{day}.csv")


def main():
    # 直接使用接口返回的榜单时间
    publish_ms = get(PUBLISH_API)["data"]["publishTime"]
    publish_time = datetime.fromtimestamp(publish_ms / 1000, BJ)
    publish_str = publish_time.strftime("%Y-%m-%d %H:%M")
    crawl_str = datetime.now(BJ).strftime("%Y-%m-%d %H:%M:%S")

    os.makedirs(DATA_DIR, exist_ok=True)
    path = csv_path(publish_time.strftime("%Y-%m-%d"))

    old = pd.read_csv(path, dtype=str) if os.path.exists(path) else pd.DataFrame(columns=COLUMNS)
    if publish_str in set(old["榜单更新时间"]):
        print(f"榜单 {publish_str} 已经抓取过，跳过")
        return

    items = fetch_ranking()
    if not items:
        sys.exit("没有抓到数据")
    new = pd.DataFrame([
        [publish_str, crawl_str, i + 1, it["gameName"], it["id"]]
        for i, it in enumerate(items)
    ], columns=COLUMNS).astype(str)

    pd.concat([old, new], ignore_index=True).to_csv(path, index=False, encoding="utf-8-sig")
    print(f"已抓取榜单 {publish_str}，共 {len(new)} 部作品 -> {path}")


if __name__ == "__main__":
    main()
