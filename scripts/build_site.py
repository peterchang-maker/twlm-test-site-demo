# -*- coding: utf-8 -*-
"""把 progress.json 灌進 site/index.html，產出可以直接發布的 public/。

刻意把資料「內嵌」進 HTML，而不是讓頁面自己 fetch：
少一次網路請求、檔案下載下來離線也打得開、而且不會有
「JSON 還沒回來所以先閃一個空白版面」的問題。
public/progress.json 照樣留一份，給之後別的東西讀。
"""
from __future__ import annotations
import argparse, json, os, shutil, sys

MARK_OPEN = '<script id="payload" type="application/json">'
MARK_CLOSE = "</script>"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="data/progress.json")
    ap.add_argument("--site", default="site")
    ap.add_argument("--out", default="public")
    ap.add_argument("--tasks", default="data/tasks.json",
                    help="fetch_issues.py 抓的任務資料；檔不在也照樣 build，頁面會說任務資料沒抓到")
    a = ap.parse_args()

    if not os.path.exists(a.data):
        sys.exit(f"找不到 {a.data}，要先跑 scan_progress.py。")
    payload = json.load(open(a.data, encoding="utf-8"))
    # 任務（GitLab Issues）是另一條資料線，抓不到不擋看板，但要讓頁面知道、顯示出來
    if os.path.exists(a.tasks):
        payload["任務資料"] = json.load(open(a.tasks, encoding="utf-8"))
    else:
        payload["任務資料"] = {"錯誤": f"沒有 {a.tasks}（這次 build 沒跑 fetch_issues.py）", "任務": []}

    tpl_path = os.path.join(a.site, "index.html")
    tpl = open(tpl_path, encoding="utf-8").read()
    i = tpl.find(MARK_OPEN)
    if i < 0:
        sys.exit(f"{tpl_path} 裡找不到資料佔位的 <script id=\"payload\">。")
    j = tpl.find(MARK_CLOSE, i)

    # </script> 不能原樣出現在 script 標籤內容裡，會提早把標籤關掉。
    blob = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    html = tpl[: i + len(MARK_OPEN)] + blob + tpl[j:]

    if os.path.isdir(a.out):
        shutil.rmtree(a.out)
    os.makedirs(a.out)
    for f in os.listdir(a.site):
        s = os.path.join(a.site, f)
        if os.path.isfile(s) and f != "index.html":
            shutil.copy2(s, os.path.join(a.out, f))
    with open(os.path.join(a.out, "index.html"), "w", encoding="utf-8") as fh:
        fh.write(html)
    shutil.copy2(a.data, os.path.join(a.out, "progress.json"))
    tk = payload["任務資料"]
    print(f"產出 {a.out}/index.html（{len(html)//1024} KB）與 progress.json"
          + (f"｜任務 {len(tk.get('任務', []))} 張" if not tk.get("錯誤") else f"｜任務資料沒抓到：{tk['錯誤']}"))


if __name__ == "__main__":
    main()
