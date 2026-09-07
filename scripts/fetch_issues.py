# -*- coding: utf-8 -*-
"""從 GitLab Issues 抓「誰負責哪份展表、第幾輪、什麼狀態」，寫成 data/tasks.json 給看板用。

階段 2 的設計（規劃書第二節）：發派任務不自己做，用 GitLab Issues——
  ・一份展表一張 Issue，標題就是展表標題（去掉「QA - 」）
  ・負責人 = Issue 的 assignee
  ・輪次／狀態／類型 = 互斥標籤 `輪次::第1輪`、`狀態::待測`、`類型::活動`
  ・版本 = milestone（`0826`）
看板拿 Issue 標題跟掃描器算出的展表標題「對起來」，對法在 task_key()：兩邊都做同樣的正規化。

認證：讀 GITLAB_READ_TOKEN（或 GITLAB_TOKEN；read_api 就夠）。CI 裡沒設就試 CI_JOB_TOKEN。
都不行時不讓 pipeline 失敗——進度數字還是對的——但會把原因寫進 tasks.json，看板上會顯示「任務資料沒抓到」。
"""
from __future__ import annotations
import argparse, datetime, json, os, re, sys, unicodedata, urllib.parse

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")

PREFIX_RE = re.compile(r"^QA\s*-\s*")
KEEP_RE = re.compile(r"[^0-9a-z一-鿿぀-ヿ가-힯]")


def task_key(title):
    """展表標題 ↔ Issue 標題 的對照鍵：去 QA- 前綴、全形轉半形、小寫、只留中英日韓字與數字。
    site/index.html 裡有一份一模一樣的 JS 版，改這裡要一起改。"""
    s = unicodedata.normalize("NFKC", str(title or ""))
    s = PREFIX_RE.sub("", s).lower()
    return KEEP_RE.sub("", s)


def guess_type(title):
    t = PREFIX_RE.sub("", str(title or "")).strip()
    body = re.sub(r"^\(.*?\)", "", t)          # 去掉「(台灣提案)」這種前綴
    if body.startswith("活動"):
        return "活動"
    if body.startswith(("BM", "IAP")):
        return "BM"
    if body.startswith("系統"):
        return "系統"
    return "基本測試"


def scoped(labels, prefix):
    for lb in labels or []:
        if lb.startswith(prefix + "::"):
            return lb.split("::", 1)[1]
    return None


def local_time(iso):
    if not iso:
        return ""
    dt = datetime.datetime.fromisoformat(iso.replace("Z", "+00:00"))
    return dt.astimezone().strftime("%Y-%m-%d %H:%M")


def parse_issues(items):
    """GitLab issues API 的 list → 看板要的欄位。純函式，selftest 有測。"""
    out = []
    for it in items:
        labels = it.get("labels") or []
        out.append({
            "iid": it.get("iid"),
            "標題": it.get("title", ""),
            "鍵": task_key(it.get("title", "")),
            "狀態": scoped(labels, "狀態") or ("完成" if it.get("state") == "closed" else "未標"),
            "輪次": scoped(labels, "輪次"),
            "類型": scoped(labels, "類型") or guess_type(it.get("title", "")),
            "負責人": [{"名": a.get("name"), "帳號": a.get("username"), "頭像": a.get("avatar_url")}
                       for a in (it.get("assignees") or [])],
            "里程碑": (it.get("milestone") or {}).get("title"),
            "網址": it.get("web_url"),
            "更新時間": local_time(it.get("updated_at")),
            "open": it.get("state") == "opened",
            "到期": it.get("due_date"),
        })
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--api", default=os.environ.get("CI_API_V4_URL", "https://gitlab.example.internal/api/v4"))
    ap.add_argument("--project", default=os.environ.get("CI_PROJECT_ID", "peterchang/twlm-test-site"))
    ap.add_argument("--out", default="data/tasks.json")
    ap.add_argument("--strict", action="store_true", help="抓不到就失敗（預設不失敗，只在 tasks.json 記原因）")
    a = ap.parse_args()

    import requests
    proj = urllib.parse.quote(str(a.project), safe="")
    web = os.environ.get("CI_PROJECT_URL") or (a.api.split("/api/")[0] + "/" + str(a.project))
    result = {
        "抓取時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "專案網址": web,
        "任務清單網址": web + "/-/issues",
        "看板網址": web + "/-/boards",
        "新任務網址": web + "/-/issues/new",
        "錯誤": None,
        "任務": [],
    }

    token = os.environ.get("GITLAB_READ_TOKEN") or os.environ.get("GITLAB_TOKEN")
    job_token = os.environ.get("CI_JOB_TOKEN")
    if token:
        headers, how = {"PRIVATE-TOKEN": token}, "GITLAB_READ_TOKEN"
    elif job_token:
        headers, how = {"JOB-TOKEN": job_token}, "CI_JOB_TOKEN"
    else:
        headers, how = None, None

    if not headers:
        result["錯誤"] = "沒有 GITLAB_READ_TOKEN（GitLab 設定 → CI/CD → 變數，Project Access Token，read_api 就夠）"
    else:
        items, page = [], 1
        try:
            while True:
                r = requests.get(f"{a.api}/projects/{proj}/issues",
                                 params={"state": "all", "per_page": 100, "page": page, "order_by": "updated_at"},
                                 headers=headers, timeout=60)
                if r.status_code != 200:
                    detail = ""
                    try:
                        detail = r.json().get("message") or r.json().get("error") or ""
                    except Exception:
                        detail = r.text[:120]
                    result["錯誤"] = f"GitLab 回 HTTP {r.status_code}（用 {how}）：{detail}"
                    if how == "CI_JOB_TOKEN":
                        result["錯誤"] += "。CI_JOB_TOKEN 不能讀 Issues 的話，請加 GITLAB_READ_TOKEN 變數（Project Access Token，read_api）"
                    items = []
                    break
                items += r.json()
                nxt = r.headers.get("X-Next-Page")
                if not nxt:
                    break
                page = int(nxt)
        except requests.RequestException as ex:
            result["錯誤"] = f"連不到 GitLab API：{ex.__class__.__name__}"
            items = []
        result["任務"] = parse_issues(items)

    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=1)

    if result["錯誤"]:
        print("任務資料沒抓到：" + result["錯誤"], file=sys.stderr)
        if a.strict:
            sys.exit(1)
    else:
        n = result["任務"]
        assigned = sum(1 for t in n if t["負責人"])
        print(f"任務 {len(n)} 張｜已指派 {assigned}｜未指派 {len(n) - assigned}"
              f"｜狀態：" + "、".join(f"{k} {v}" for k, v in sorted(
                  __import__('collections').Counter(t['狀態'] for t in n).items())))


if __name__ == "__main__":
    main()
