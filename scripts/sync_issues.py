# -*- coding: utf-8 -*-
"""把一個版本的展表開成 GitLab Issues（一份展表一張），並確保標籤與 milestone 都在。

什麼時候跑：新版本第一次掃完展表之後（或展表有新增時）。重複跑沒關係——
已經有的 Issue 不會重開、不會關、不會改人家設好的標籤與負責人；只補缺的。

    set GITLAB_TOKEN=<有 api 權限的 Project Access Token>
    python scripts\\sync_issues.py --config config.yml --from data\\progress.json
    python scripts\\sync_issues.py --config config.yml --names "活動_共同_神秘的流浪馬戲團" "BM_月底_…"

Token 只從環境變數讀，不印出來、不寫進任何檔。
"""
from __future__ import annotations
import argparse, json, os, sys, urllib.parse

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_issues import task_key, guess_type   # noqa: E402

LABELS = [
    ("狀態::待測", "#8A98A6", "還沒開始測"),
    ("狀態::測試中", "#1F5C8B", "正在測"),
    ("狀態::卡住", "#A5443C", "被擋住：等企劃回覆、等環境、等修"),
    ("狀態::完成", "#2F7A6F", "這一輪測完"),
    ("輪次::第1輪", "#6E49CB", "第一次測試"),
    ("輪次::第2輪", "#6E49CB", "第二次測試（覆測）"),
    ("輪次::第3輪", "#6E49CB", "第三次測試（覆測）"),
    ("輪次::第4輪", "#6E49CB", "第四次測試（覆測）"),
    ("類型::活動", "#E24329", "活動展表"),
    ("類型::BM", "#FC6D26", "BM／IAP 商品展表"),
    ("類型::系統", "#0E6E9A", "系統展表"),
    ("類型::基本測試", "#5C5C5C", "RC／送審／fix／持續與結束／伺服器移民 這類版本基本測試"),
]

DESCRIPTION = """展表連結：（把 Google 試算表的網址貼在這裡）

這張卡對應看板「天堂M 測試進度」上的展表「{title}」（版本 {version}）。

- **指派**：右側 Assignees 選要測的人
- **輪次／狀態**：右側 Labels 改 `輪次::第N輪`、`狀態::待測／測試中／卡住／完成`（同一組只會留一個）
- **進度數字**（母數／已測／異常）由看板從展表算出來，不在這裡填
- 有問題要問企劃：在這裡留言或另開一張 Issue 指派給企劃
"""


class GitLab:
    def __init__(self, api, project, token):
        import requests
        self.s = requests.Session()
        self.s.headers["PRIVATE-TOKEN"] = token
        self.base = f"{api}/projects/{urllib.parse.quote(str(project), safe='')}"

    def _check(self, r):
        if r.status_code >= 400:
            try:
                msg = r.json().get("message") or r.json().get("error")
            except Exception:
                msg = r.text[:200]
            sys.exit(f"GitLab 回 HTTP {r.status_code}：{msg}\n  token 有 api 權限嗎？專案路徑對嗎？")
        return r.json()

    def get_all(self, path, **params):
        out, page = [], 1
        while True:
            r = self.s.get(self.base + path, params={**params, "per_page": 100, "page": page}, timeout=60)
            out += self._check(r)
            nxt = r.headers.get("X-Next-Page")
            if not nxt:
                return out
            page = int(nxt)

    def post(self, path, **data):
        return self._check(self.s.post(self.base + path, json=data, timeout=60))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.yml")
    ap.add_argument("--api", default="https://gitlab.example.internal/api/v4")
    ap.add_argument("--project", default="peterchang/twlm-test-site")
    ap.add_argument("--from", dest="src", default=None, help="progress.json，拿裡面的展表標題")
    ap.add_argument("--names", nargs="*", default=None, help="直接給展表標題（不含 QA - ）")
    ap.add_argument("--dry-run", action="store_true", help="只印會做什麼，不真的建")
    a = ap.parse_args()

    import yaml
    cfg = yaml.safe_load(open(a.config, encoding="utf-8")) or {}
    version = str(cfg.get("版本", "")).strip()
    if not version:
        sys.exit("config.yml 沒有「版本」。")

    titles = list(a.names or [])
    if a.src:
        data = json.load(open(a.src, encoding="utf-8"))
        titles += [bk["標題"] for bk in data.get("展表", [])]
    titles = [t.strip() for t in titles if t and t.strip()]
    if not titles:
        sys.exit("沒有展表標題：給 --from data/progress.json 或 --names。")

    token = os.environ.get("GITLAB_TOKEN")
    if not token and not a.dry_run:
        sys.exit("需要環境變數 GITLAB_TOKEN（Project Access Token，api 權限）。只想看會做什麼可以加 --dry-run。")

    if a.dry_run:
        for t in titles:
            print(f"[dry-run] 會確保 Issue「{t}」存在：標籤 狀態::待測、輪次::第1輪、類型::{guess_type(t)}，milestone {version}")
        return

    gl = GitLab(a.api, a.project, token)

    # 1. 標籤
    have = {lb["name"] for lb in gl.get_all("/labels")}
    for name, color, desc in LABELS:
        if name not in have:
            gl.post("/labels", name=name, color=color, description=desc)
            print("建標籤", name)

    # 2. milestone = 版本
    ms = [m for m in gl.get_all("/milestones", state="active") if m["title"] == version]
    if not ms:
        ms_obj = gl.post("/milestones", title=version, description=f"版本 {version}：{cfg.get('站台標題', '')}".strip())
        print("建 milestone", version)
    else:
        ms_obj = ms[0]

    # 3. Issues：用正規化鍵對，避免「QA - 」前綴、全半形、斜線底線造成重複
    existing = {task_key(i["title"]): i for i in gl.get_all("/issues", state="all")}
    made, kept = 0, 0
    for t in titles:
        k = task_key(t)
        if k in existing:
            kept += 1
            continue
        gl.post("/issues", title=t, description=DESCRIPTION.format(title=t, version=version),
                labels=",".join(["狀態::待測", "輪次::第1輪", f"類型::{guess_type(t)}"]),
                milestone_id=ms_obj["id"])
        existing[k] = {"title": t}
        made += 1
        print("開卡", t)
    print(f"完成：新開 {made} 張，已存在 {kept} 張，milestone {version}")


if __name__ == "__main__":
    main()
