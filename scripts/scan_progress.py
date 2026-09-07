# -*- coding: utf-8 -*-
"""掃過展表，把進度算成 data/progress.json。

展表可能在三種地方，這支都吃：
  1. 資料夾裡的 .xlsx（本機匯出、或 runner 掛到的共用資料夾）
  2. 資料夾裡的 .gsheet（Google Drive 桌面版掛出來的捷徑檔——內容其實在 Google 試算表上）
  3. Google 雲端硬碟「資料夾」（給網址或 ID；CI 用這種，因為 runner 看不到 G:）
資料夾都是遞迴掃的：展表散在版本資料夾第一層和 @活動、@BM 底下各一層。
2 和 3 要透過 Google Sheets API 讀，需要一把「唯讀」service account 金鑰
（環境變數 GOOGLE_APPLICATION_CREDENTIALS 指到 json 檔）。金鑰不進 git、不印進 log。

數法只有一份（scan_sheet）。xlsx 用 openpyxl 讀、Google 試算表用 Sheets API 讀，
兩邊都包成同一個最小介面（title / ws[4] / cell() / max_row）再丟進去數，
這樣兩種來源不會各算各的。scripts/selftest.py 會拿同一份資料走兩條路比對結果。

進度不是抄試算表裡 COUNTIF 的結果，是自己一格一格數判定格。展表「測試進度」分頁的
統計只拿來「對帳」：跟自己數的不一致就列在看板上，不拿它當答案。
但「哪些分頁算」是照它的：測試進度第 3 列公式點到哪些分頁，就掃那些（scope_from_formulas）。

母數的算法有一個一定要照做的地方：表頭要照「欄的順序」讀成 list，
不能塞進 dict。四輪覆測的判定欄名字一模一樣，用 dict 只會留下最後
一輪那一欄，母數會整個算錯——實測差了 2.4 倍。

掃到 0 份展表會直接讓 job 失敗（除非 --allow-empty）。資料夾在、但裡面
沒有掃得到的展表，跟資料夾不在一樣危險：都會產出一頁全零、卻長得像真的看板。
"""
from __future__ import annotations
import argparse, datetime, io, json, os, re, sys, time, zipfile

for _stream in (sys.stdout, sys.stderr):      # Windows 導向到檔案時是 cp950，中文以外的符號別讓它炸
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")

import openpyxl
import yaml

# 判定欄：表頭以「確認」開頭。原本只認「確認內容…」，2026-09-07 拿 0826 版 15 份實掃後放寬——
# 序號兌換（確認兌換獲得品項數量正確）、怪物設定（確認消耗正常消耗殷海薩…五欄）、
# 製作清單（確認製作時間區間設定正確…三欄）都是判定欄，只認「確認內容」會整個分頁漏掉。
# 15 份掃過一輪，以「確認」開頭的表頭全部都是判定欄，沒有例外。
JUDGE = re.compile(r"^確認")
JUDGE_TAIL = re.compile(r"(正確|正常)$")     # 同一輪區塊裡「機率正確」「payback正確」這種也是判定欄
OKV = ("Y", "N", "X", "未測試")
SKIP_SHEETS = ("測試進度", "圖片資料庫", "圖片資料庫(勿動)")
NAME_KEYS = ("道具名稱", "道具名稱正確", "企畫書道具名", "NPC名", "道具名", "製作結果道具",
             "怪物名稱", "項目", "區分", "福袋")
ID_KEYS = ("ItemID", "ItemClassID", "NPC ID", "CraftID")
FLAG_HINTS = ("一致", "對照", "檢查結果", "狀態")

DEFAULT_PREFIX = "QA - "
SKIP_FILE_PREFIX = ("~$", "提問清單")

MIME_FOLDER = "application/vnd.google-apps.folder"
MIME_GSHEET = "application/vnd.google-apps.spreadsheet"
MIME_SHORTCUT = "application/vnd.google-apps.shortcut"
DRIVE = "https://www.googleapis.com/drive/v3"
SHEETS = "https://sheets.googleapis.com/v4/spreadsheets"
SCOPES = ("https://www.googleapis.com/auth/spreadsheets.readonly",
          "https://www.googleapis.com/auth/drive.readonly")


# ---------------------------------------------------------------- 數法
# 口徑照展表自己的算法（2026-09-04 跟 0401 版 6 份展表的「測試進度」分頁對過）：
#   ・表頭列不固定：多數模板在第 4 列，BM商城／製作／binary 測試在第 5 列（第 4 列是群組標題）。
#     取前 15 列裡第一個出現「確認內容…」判定欄的那一列。
#   ・第一輪是哪些欄：用上方「第N次測試」標籤切——第一個標籤到第二個標籤之間的判定欄。
#     binary 類一輪有 AOS、iOS 兩欄判定，中間隔著備註／人員／日期，用「連續判定欄」會漏掉一欄。
#     找不到標籤才退回連續法。
#   ・每個判定欄各自的備註欄，是它右邊、下一個判定欄之前的第一個「異常狀況備註／測試備註」。
#   ・掃哪些分頁（2026-09-07 起）：照展表自己「測試進度」分頁第 3 列公式點到的分頁。
#     那串公式就是測試人員宣告的範圍——空模板分頁（哈爾巴斯留著 BM商城／現金商城／製作／Buff
#     四個沒用的模板，131 列未測試）不在公式裡，就不算進母數；反過來有判定欄但沒掛進公式的分頁
#     會列成警告，讓人決定是補公式還是本來就不用測。沒有「測試進度」或公式抓不到分頁名時退回
#     「所有有判定欄的分頁都算」。
#   ・一格一項：一列有五個判定欄就是五項（怪物設定 118 列 × 5 欄 = 590），跟展表 COUNTIF 一樣。
HEADER_SEARCH_ROWS = 15
ROUND_LABEL = re.compile(r"^(第\s*[一二三四五六七八九十\d]+\s*次(測試)?|複測|覆測)")
NOTE_HEADS = ("異常狀況備註", "測試備註")
# 公式裡的分頁參照：'道具說明使用'!AB3（名字裡的 ' 會寫成 ''）或 binary基本測試!F3（沒引號）
SHEET_REF = re.compile(r"'((?:[^']|'')+)'!|([^\s'!+\-*/(),=:;&^<>\"%\[\]{}]+)!")
EXCEL_SHEET_NAME_MAX = 31


def _clean(v):
    return str(v).replace("\n", "").strip()


def header_row(ws):
    """表頭列：前 15 列裡第一個「有『確認…』判定欄、而且同列有備註欄或測試人員欄」的列。
    第二個條件是擋「確認常發生問題的變身，有時間再抽查其他變身」這種寫在表頭上方的說明文字
    （fix測試的變身確認分頁）。都沒有符合的，退回第一個有「確認…」的列。"""
    first = None
    for r in range(1, min(HEADER_SEARCH_ROWS, ws.max_row) + 1):
        vals = [_clean(c.value) for c in ws[r] if c.value not in (None, "")]
        if any(JUDGE.match(v) for v in vals):
            if first is None:
                first = r
            if any(v in NOTE_HEADS or v.startswith("測試人員") for v in vals):
                return r
    return first


def sheet_formula_rows(ws):
    """一張分頁前 15 列的原始值（呼叫端保證是公式字串，不是算好的值）。
    openpyxl 沒開 data_only 時 cell.value 就是公式；Sheets API 用 valueRenderOption=FORMULA。"""
    out = []
    for r in range(1, min(HEADER_SEARCH_ROWS, ws.max_row) + 1):
        row = {}
        for c in ws[r]:
            if c.value not in (None, ""):
                row[c.column] = c.value
        out.append(row)
    return out


COUNTIF_RANGE = re.compile(r"COUNTIFS?\(\s*(?:'[^']*'|[^\s'!(,]+)?(!?)\$?([A-Z]{1,3})\$?\d+\s*:\s*\$?([A-Z]{1,3})\$?\d+", re.I)


def col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def round1_bounds(frows):
    """分頁自己第一輪「所有已測數」底下那格 COUNTIF 的欄範圍，例如 =COUNTIF(F6:M99,"Y") → (6, 13)。
    這是展表自己宣告的「第一輪有哪些欄」，比猜還準：變身確認左右兩張表（F 與 P 兩個判定欄）、
    製作清單五個判定欄，都靠這個。frows 是 sheet_formula_rows 的結果；拿不到就回 None。"""
    for i, row in enumerate(frows):
        for col, v in row.items():
            if isinstance(v, str) and _clean(v) == "所有已測數":
                below = frows[i + 1].get(col) if i + 1 < len(frows) else None
                if not isinstance(below, str):
                    return None
                cols = []
                for bang, a, b in COUNTIF_RANGE.findall(below):
                    if bang:
                        return None      # 點到別張分頁的範圍，不能拿來當這張的欄範圍
                    cols += [col_index(a), col_index(b)]
                # 多段相加（=COUNTIF(F6:F99,"Y")+COUNTIF(J6:J99,"Y")）取聯集
                return (min(cols), max(cols)) if cols else None
    return None


def round_label_cols(ws, hrow):
    """表頭列上方第一個有「第N次測試」標籤的列，回它的標籤欄位（由左到右）。沒有就回 []。"""
    for r in range(1, hrow):
        cols = sorted(c.column for c in ws[r]
                      if c.value not in (None, "") and ROUND_LABEL.match(str(c.value).strip()))
        if cols:
            return cols
    return []


def col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def scan_sheet(ws, frows=None):
    """數一張分頁。frows：這張分頁前 15 列的公式（sheet_formula_rows）；不給就從 ws 本身取
    （openpyxl 沒開 data_only 時 ws 裡就是公式——selftest 走這條）。"""
    hrow = header_row(ws)
    if hrow is None:
        return None
    if frows is None:
        frows = sheet_formula_rows(ws)
    # 照欄順序讀表頭成 list。heads 只給「用名字找欄」用，取第一個出現的。
    heads_row = [(c.column, _clean(c.value)) for c in ws[hrow] if c.value not in (None, "")]
    heads = {}
    for col, h in heads_row:
        heads.setdefault(h, col)

    jcols = [col for col, h in heads_row if JUDGE.match(h)]
    if not jcols:
        return None
    labels = round_label_cols(ws, hrow)
    bounds = round1_bounds(frows)
    if bounds and len(labels) >= 2 and not (labels[0] <= bounds[0] < labels[1]):
        bounds = None       # COUNTIF 的欄範圍跟「第N次測試」標籤對不上，不信它
    # 第一輪區塊的右界：第二個「第N次測試」標籤（沒有就到底）。備註欄找不到就不會往第二輪借。
    block_end = labels[1] if len(labels) >= 2 else 10 ** 9
    if bounds:
        block_end = min(block_end, bounds[1] + 1)
    # 「機率正確」「payback正確」（製作清單）這種沒寫「確認」但一樣填 Y/N 的判定欄：
    # 只認第一個「確認…」欄右邊、同一輪區塊裡、以「正確／正常」結尾的表頭。
    # 區塊左邊的「怪物名稱正確」是資料欄，不會被撈進來。
    first_j = min(jcols)
    jcols = sorted(set(jcols) | {col for col, h in heads_row
                                 if first_j < col < block_end and JUDGE_TAIL.search(h)})

    # 第一輪是哪些欄，三招按可信度排：
    #   1. 分頁自己「所有已測數」的 COUNTIF 欄範圍（展表自己宣告的第一輪範圍）
    #   2. 「第N次測試」標籤：第一個到第二個標籤之間。要有兩個以上，且第一個標籤要在第一個判定欄左邊或同一欄
    #   3. 都沒有：從第一個判定欄往右取連續的一段，遇到非判定欄就停
    r1 = []
    if bounds:
        r1 = [c for c in jcols if bounds[0] <= c <= bounds[1]]
    if not r1 and len(labels) >= 2 and labels[0] <= min(jcols):
        r1 = [c for c in jcols if labels[0] <= c < labels[1]]
    if not r1:
        first = min(jcols)
        for col, h in heads_row:
            if col < first:
                continue
            if col in jcols:
                r1.append(col)
            elif r1:
                break
    head_of = dict(heads_row)

    def note_after(c):
        # 先找它跟下一個判定欄之間的備註欄；五個判定欄並排、備註只有一欄在最右邊時
        # （怪物設定那種），退回到這一輪區塊裡它右邊第一個備註欄。
        nxt = min([j for j in jcols if j > c], default=10 ** 9)
        found = next((col for col, h in heads_row if c < col < nxt and h in NOTE_HEADS), None)
        if found is None:
            found = next((col for col, h in heads_row if c < col < block_end and h in NOTE_HEADS), None)
        return found

    notecols = {c: note_after(c) for c in r1}
    flagcols = [col for h, col in heads.items()
                if any(k in h for k in FLAG_HINTS)]
    namecol = next((heads[k] for k in NAME_KEYS if k in heads), 1)
    idcol = next((heads[k] for k in ID_KEYS if k in heads), None)

    cnt = {"Y": 0, "N": 0, "X": 0, "未測試": 0, "其他": 0}
    stars, blanks, others = [], [], []

    def cell(r, c):
        return str(ws.cell(row=r, column=c).value or "").strip()

    for r in range(hrow + 1, ws.max_row + 1):
        n_without_note = False
        for c in r1:
            v = cell(r, c)
            if not v:
                continue
            if _clean(v) == head_of.get(c):
                continue        # 分頁中段又出現一次表頭（變身確認一張表分好幾段），不是判定值
            if v.startswith("/"):
                continue        # 段落上方寫的遊戲指令（/1 157591 1 需確認…），不是判定值
            if v in OKV:
                cnt[v] += 1
                if v == "N" and notecols[c] and not cell(r, notecols[c]):
                    n_without_note = True
            else:
                cnt["其他"] += 1
                if len(others) < 50:
                    others.append({"列": r, "值": v[:30],
                                   "名": cell(r, namecol)[:40]})
        if n_without_note:
            blanks.append({"列": r, "名": cell(r, namecol)[:40],
                           "id": cell(r, idcol) if idcol else ""})
        for c in flagcols:
            v = cell(r, c)
            if v.startswith("★"):
                stars.append({"列": r, "名": cell(r, namecol)[:40],
                              "id": cell(r, idcol) if idcol else "",
                              "說明": v[:150]})
                break

    total = sum(cnt.values())
    if total == 0 and not stars:
        return None
    # 母數＝Y＋N＋未測試。X 是不適用；「其他」是認不得的值（會列在判定值異常），兩者都不算進母數——
    # 跟展表自己的 COUNTIF 一樣，看板文字也是這麼寫的。
    return {"分頁": ws.title, "計": cnt, "總": total,
            "母數": total - cnt["X"] - cnt["其他"],
            "已測": cnt["Y"] + cnt["N"],
            "表頭列": hrow, "第一輪欄": [col_letter(c) for c in r1],
            "star": stars, "blank": blanks, "其他值": others}


def self_report(rows):
    """展表自己「測試進度」分頁的第一次測試統計：第 2 列找「所有項目數」，第 3 列往右取 4 格。
    rows 是前三列的值。拿不到就回 None（不擋，只是少一個對照）。"""
    r2 = list(rows[1]) if len(rows) > 1 else []
    r3 = list(rows[2]) if len(rows) > 2 else []
    for i, v in enumerate(r2):
        if _clean(v) == "所有項目數":
            vals = [r3[j] if j < len(r3) else None for j in range(i, i + 4)]
            try:
                n = [int(float(x)) for x in vals]
            except (TypeError, ValueError):
                return None
            return {"項目": n[0], "Y": n[1], "未測試": n[2], "N": n[3]}
    return None


def compare_report(report, agg):
    """展表自報 vs 看板算的（第一輪）。展表的「所有項目數」= Y+未測+N，不含 X、也不含怪值。"""
    mine = {"項目": agg["Y"] + agg["N"] + agg["未測試"], "Y": agg["Y"],
            "未測試": agg["未測試"], "N": agg["N"]}
    labels = {"項目": "項目數", "Y": "已測（Y）", "未測試": "未測試", "N": "異常（N）"}
    return [{"欄位": labels[k], "自報": report[k], "看板": mine[k]}
            for k in ("項目", "Y", "未測試", "N") if report[k] != mine[k]]


def scan_worksheets(worksheets):
    """全部有判定欄的分頁都數（沒有範圍資訊時的退路；selftest 也直接用這個）。"""
    out = []
    for ws in worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        s = scan_sheet(ws)
        if s:
            out.append(s)
    return out


def scope_from_formulas(rows):
    """從「測試進度」前三列（公式字串，不是算好的值）讀出第一輪「所有項目數」等四格公式
    點到的分頁名。回 list（照出現順序、不重複）；找不到公式就回 []，呼叫端退回全掃。
    公式範例：='道具說明使用'!AB3+'序號兌換測試'!G3"""
    r2 = list(rows[1]) if len(rows) > 1 else []
    r3 = list(rows[2]) if len(rows) > 2 else []
    start = next((i for i, v in enumerate(r2) if v is not None and _clean(v) == "所有項目數"), None)
    if start is None:
        return []
    names = []
    for j in range(start, start + 4):
        f = r3[j] if j < len(r3) else None
        if not isinstance(f, str) or not f.startswith("="):
            continue
        for quoted, bare in SHEET_REF.findall(f):
            n = quoted.replace("''", "'") if quoted else bare
            if n and n not in names:
                names.append(n)
    return names


def scan_book(worksheets, formulas):
    """一本展表。formulas：{分頁名: 前 15 列公式}（含「測試進度」），從中拿兩件事：
      ・掃描範圍：測試進度第 3 列公式點到哪些分頁（scope_from_formulas）
      ・每張分頁第一輪的欄範圍：它自己 COUNTIF 的欄（round1_bounds，在 scan_sheet 裡用）
    回 (算進去的分頁 list, 範圍資訊 dict)。範圍抓不到就全掃、記「自動」。"""
    formulas = formulas or {}
    scope = scope_from_formulas(formulas.get("測試進度", []))
    by_title = {ws.title: ws for ws in worksheets}

    def frows_of(ws):
        f = formulas.get(ws.title)
        return [{c: v for c, v in enumerate(row, 1) if v not in (None, "")} for row in f] if f is not None else None

    def scan_all():
        out = []
        for ws in worksheets:
            if ws.title in SKIP_SHEETS:
                continue
            s = scan_sheet(ws, frows_of(ws))
            if s:
                out.append(s)
        return out

    info = {"範圍": "自動", "未掛進進度表": [], "進度表點到但沒這分頁": [], "進度表點到但掃不到": []}
    if not scope:
        return scan_all(), info

    # xlsx 的分頁名最長 31 字，Google 匯出時會截斷；公式裡是全名，對不到就拿前 31 字再對一次
    def find(name):
        return by_title.get(name) or by_title.get(name[:EXCEL_SHEET_NAME_MAX])
    counted, absent, unscannable, extra = [], [], [], []
    used = set()
    for name in scope:
        ws = find(name)
        if ws is None or ws.title in SKIP_SHEETS:
            absent.append(name)
            continue
        if ws.title in used:
            continue
        used.add(ws.title)
        s = scan_sheet(ws, frows_of(ws))
        if s:
            counted.append(s)
        else:
            unscannable.append(name)
    for ws in worksheets:
        if ws.title in SKIP_SHEETS or ws.title in used:
            continue
        s = scan_sheet(ws, frows_of(ws))
        if s and s["總"]:
            extra.append({"分頁": ws.title, "母數": s["母數"], "已測": s["已測"]})
    info.update({"進度表點到但沒這分頁": absent, "進度表點到但掃不到": unscannable, "未掛進進度表": extra})
    if not counted and extra:
        # 公式點到的分頁一張都對不到（多半是分頁改名、公式沒跟著改），但別的分頁有判定欄：
        # 不能讓整本展表從看板上消失，退回全掃，並在範圍註記，呼叫端會出警告。
        info["範圍"] = "自動（進度表公式對不到任何分頁，退回全掃）"
        info["未掛進進度表"] = []
        return scan_all(), info
    info["範圍"] = "進度表"
    return counted, info


# ------------------------------------------- Google 試算表包成 openpyxl 的最小介面
class _Cell:
    __slots__ = ("column", "value")

    def __init__(self, column, value):
        self.column, self.value = column, value


class GridSheet:
    """Sheets API 回來的二維 values（list of rows）。空格是 ""、列尾空格會被 API 省略，
    所以超出範圍一律回 None，跟 openpyxl 讀到空格的行為一致。"""

    def __init__(self, title, values):
        self.title = title
        self._rows = values or []

    @property
    def max_row(self):
        return len(self._rows)

    def __getitem__(self, row):           # ws[4] → 第 4 列所有有值的格
        vals = self._rows[row - 1] if 0 < row <= len(self._rows) else []
        return [_Cell(i + 1, v) for i, v in enumerate(vals)]

    def cell(self, row, column):
        vals = self._rows[row - 1] if 0 < row <= len(self._rows) else []
        v = vals[column - 1] if 0 < column <= len(vals) else None
        return _Cell(column, None if v == "" else v)


class Google:
    """唯讀 service account。金鑰路徑放 GOOGLE_APPLICATION_CREDENTIALS。"""

    def __init__(self):
        key = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")
        if not key or not os.path.isfile(key):
            # 絕對不要把變數的值印出來：如果有人把金鑰內容（而不是路徑）貼進一般變數，印出來就進 log 了。
            if key.lstrip().startswith("{"):
                state = "看起來是金鑰的「內容」而不是路徑——GitLab 變數請改成 File 類型"
            elif key:
                state = f"指到 {key!r}，但那裡沒有檔案"
            else:
                state = "沒設"
            sys.exit("要讀 Google 試算表需要 service account 金鑰：\n"
                     "  把金鑰 json 的『路徑』放進環境變數 GOOGLE_APPLICATION_CREDENTIALS\n"
                     "  （GitLab 上用 File 類型的 CI/CD 變數，名字就叫這個）。\n"
                     "  目前這個變數：" + state)
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
        creds = service_account.Credentials.from_service_account_file(key, scopes=SCOPES)
        self.email = creds.service_account_email
        self.s = AuthorizedSession(creds)
        self.calls = 0

    RETRIES = 6          # 5、10、20、40、60 秒，總共約兩分多鐘：08:00 排程撞上配額時等得過去

    def get(self, url, **params):
        import requests
        for i in range(self.RETRIES):
            wait = min(5 * 2 ** i, 60)
            try:
                r = self.s.get(url, params=params, timeout=120)
            except requests.RequestException as ex:
                if i == self.RETRIES - 1:
                    raise
                print(f"連線失敗（{ex.__class__.__name__}），{wait} 秒後重試", file=sys.stderr)
                time.sleep(wait)
                continue
            self.calls += 1
            detail = ""
            if r.status_code >= 400:
                try:
                    detail = r.json()["error"]["message"]
                except Exception:
                    detail = r.text[:200]
            rate_limited = r.status_code == 429 or (
                r.status_code == 403 and "ratelimit" in detail.replace(" ", "").lower())
            if (rate_limited or r.status_code in (500, 502, 503, 504)) and i < self.RETRIES - 1:
                print(f"Google 回 {r.status_code}（{detail[:60]}），{wait} 秒後重試", file=sys.stderr)
                time.sleep(wait)
                continue
            if r.status_code in (403, 404):
                raise PermissionError(
                    f"HTTP {r.status_code}：{detail}\n"
                    f"  service account 是 {self.email}，那個資料夾／試算表有分享給它嗎？"
                    "（在共用雲端硬碟把它加成成員，「檢視者」就夠）")
            r.raise_for_status()
            return r.json()

    def meta(self, fid):
        return self.get(f"{DRIVE}/files/{fid}", fields="id,name,mimeType,modifiedTime",
                        supportsAllDrives="true")

    def children(self, folder_id):
        token, out = None, []
        while True:
            j = self.get(f"{DRIVE}/files",
                         q=f"'{folder_id}' in parents and trashed = false",
                         fields="nextPageToken,files(id,name,mimeType,modifiedTime,"
                                "shortcutDetails(targetId,targetMimeType))",
                         pageSize=1000, orderBy="name",
                         supportsAllDrives="true", includeItemsFromAllDrives="true",
                         **({"pageToken": token} if token else {}))
            out += j.get("files", [])
            token = j.get("nextPageToken")
            if not token:
                return out

    def worksheets(self, sheet_id):
        """回傳 (GridSheet list, 測試進度前三列的值, 測試進度前三列的公式)。分頁順序照試算表；
        SKIP_SHEETS 不抓內容，只有「測試進度」多抓前三列：算好的值給自報對照、公式給掃描範圍。"""
        meta = self.get(f"{SHEETS}/{sheet_id}", fields="sheets.properties.title")
        titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
        want = [t for t in titles if t not in SKIP_SHEETS]
        ranges = ["'" + t.replace("'", "''") + "'" for t in want]
        if "測試進度" in titles:
            ranges.append("'測試進度'!A1:Z3")
        values = {}
        for i in range(0, len(ranges), 20):
            chunk = ranges[i:i + 20]
            j = self.get(f"{SHEETS}/{sheet_id}/values:batchGet",
                         ranges=chunk,
                         valueRenderOption="UNFORMATTED_VALUE",
                         dateTimeRenderOption="FORMATTED_STRING")
            for rng, vr in zip(chunk, j.get("valueRanges", [])):
                values[rng] = vr.get("values", [])
        sheets = [GridSheet(t, values.get(rng, [])) for t, rng in zip(want, ranges)]
        # 每張分頁前 15 列的「公式」（不是值）：測試進度的給掃描範圍、其他分頁的給第一輪欄範圍
        frng = ["'" + t.replace("'", "''") + "'!1:" + str(HEADER_SEARCH_ROWS) for t in titles]   # 整列，欄數不設上限
        formulas = {}
        for i in range(0, len(frng), 20):
            chunk = frng[i:i + 20]
            j = self.get(f"{SHEETS}/{sheet_id}/values:batchGet", ranges=chunk, valueRenderOption="FORMULA")
            for t, vr in zip(titles[i:i + 20], j.get("valueRanges", [])):
                formulas[t] = vr.get("values", [])
        return sheets, values.get("'測試進度'!A1:Z3", []), formulas


# ----------------------------------------------------------------- 找展表
def gsheet_id(path):
    """Google Drive 桌面版的 .gsheet 是一小段 json，裡面有 doc_id 或 url。"""
    raw = open(path, encoding="utf-8", errors="replace").read()
    try:
        j = json.loads(raw)
        for k in ("doc_id", "resource_id"):
            if j.get(k):
                return str(j[k]).split(":")[-1]
        raw = str(j.get("url", raw))
    except ValueError:
        pass
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", raw)
    return m.group(1) if m else None


def drive_folder_id(src):
    m = re.search(r"drive\.google\.com/.*?/folders/([A-Za-z0-9_-]+)", src)
    if m:
        return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", src):
        return src
    return None


def wanted(basename, prefix):
    return not basename.startswith(SKIP_FILE_PREFIX) and basename.startswith(prefix)


def local_time(rfc3339):
    dt = datetime.datetime.fromisoformat(rfc3339.replace("Z", "+00:00"))
    return dt.astimezone().strftime("%Y-%m-%d %H:%M")


def discover_local(src, prefix, stat):
    """遞迴找 .xlsx 與 .gsheet。回 [(顯示名, 種類, 位置)]，種類是 xlsx 或 gsheet。"""
    found = []
    for root, dirs, files in os.walk(src):
        dirs.sort()
        for f in sorted(files):
            low = f.lower()
            if not low.endswith((".xlsx", ".gsheet")):
                continue
            stat["候選"] += 1
            if not wanted(f, prefix):
                stat["非前綴略過"] += 1
                continue
            found.append((f, "xlsx" if low.endswith(".xlsx") else "gsheet",
                          os.path.join(root, f)))
    return found


def discover_zip(path, prefix, stat):
    """Google 雲端硬碟「下載」出來的 zip：試算表會被轉成 xlsx 放在裡面，可能有子資料夾。
    不解壓，直接讀。回 [(顯示名, "zip", (zip路徑, 成員名))]。"""
    found = []
    with zipfile.ZipFile(path) as z:
        for info in sorted(z.infolist(), key=lambda i: i.filename):
            if info.is_dir() or info.filename.startswith("__MACOSX/"):
                continue
            base = info.filename.rsplit("/", 1)[-1]
            if not base.lower().endswith(".xlsx"):
                continue
            stat["候選"] += 1
            if not wanted(base, prefix):
                stat["非前綴略過"] += 1
                continue
            found.append((base, "zip", (path, info.filename)))
    return found


def zip_top_folder(path):
    """zip 裡如果只有一個最上層資料夾（下載整個版本資料夾時會這樣），回它的名字給版本檢查用。"""
    tops = set()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("__MACOSX/"):
                continue
            if "/" in n:
                tops.add(n.split("/", 1)[0])
            else:
                return None          # 有檔案直接在根目錄：多選檔案下載的 zip，沒有資料夾名可對
    return tops.pop() if len(tops) == 1 else None


MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def discover_drive(g, folder_id, prefix, stat):
    found, queue, seen = [], [folder_id], {folder_id}
    while queue:
        for it in g.children(queue.pop(0)):
            mime, name = it["mimeType"], it["name"]
            if mime == MIME_SHORTCUT:
                sd = it.get("shortcutDetails") or {}
                if not sd.get("targetId"):
                    continue
                mime, it = sd.get("targetMimeType", ""), {**it, "id": sd["targetId"]}
            if mime == MIME_FOLDER:
                if it["id"] not in seen:          # 捷徑指回上層會繞圈，看過的不再進
                    seen.add(it["id"])
                    queue.append(it["id"])
                continue
            # 用 MIME 判斷，不看檔名：上傳後轉成 Google 試算表的檔名字仍帶 .xlsx，那是原生試算表
            is_sheet = mime == MIME_GSHEET
            is_xlsx = mime == MIME_XLSX
            if not (is_sheet or is_xlsx):
                continue
            stat["候選"] += 1
            if not wanted(name, prefix):
                stat["非前綴略過"] += 1
                continue
            if is_xlsx:
                # 上傳成檔案的 xlsx 不在這條路上讀（會整檔下載，動輒幾十 MB）。先喊一聲。
                stat["雲端xlsx未讀"] += 1
                print(f"略過雲端硬碟上的 xlsx 檔（只讀 Google 試算表）：{name}", file=sys.stderr)
                continue
            found.append((name, "drive", it["id"]))
    return found


def xlsx_progress_rows(path):
    """xlsx 的「測試進度」前三列的「值」。data_only 才拿得到公式的快取值（Google 匯出的 xlsx 有）。
    用 read_only 串流讀，不會把整本圖片再載一次。path 可以是路徑或 BytesIO；讀完 seek(0) 讓人能再讀。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if "測試進度" not in wb.sheetnames:
            return []
        return [list(r) for r in wb["測試進度"].iter_rows(min_row=1, max_row=3, values_only=True)]
    finally:
        wb.close()
        if hasattr(path, "seek"):
            path.seek(0)


def xlsx_formula_rows(path):
    """每張分頁前 15 列的「公式」：{分頁名: [[...], ...]}。read_only 串流、只讀前 15 列，很快。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        out = {}
        for ws in wb.worksheets:
            out[ws.title] = [list(r) for r in ws.iter_rows(min_row=1, max_row=HEADER_SEARCH_ROWS, values_only=True)]
        return out
    finally:
        wb.close()
        if hasattr(path, "seek"):
            path.seek(0)


def read_book(kind, where, g):
    """回傳 (已數好的分頁 list, 更新時間字串, 展表自報 or None, 範圍資訊 dict)。"""
    if kind == "zip":
        zpath, member = where
        with zipfile.ZipFile(zpath) as z:
            raw = z.read(member)
            dt = datetime.datetime(*z.getinfo(member).date_time)
        buf = io.BytesIO(raw)
        formulas = xlsx_formula_rows(buf)
        report = self_report(xlsx_progress_rows(buf))
        wb = openpyxl.load_workbook(buf, data_only=True, read_only=False)
        sheets, info = scan_book(wb.worksheets, formulas)
        wb.close()
        return sheets, dt.strftime("%Y-%m-%d %H:%M"), report, info
    if kind == "xlsx":
        formulas = xlsx_formula_rows(where)
        report = self_report(xlsx_progress_rows(where))
        # data_only：判定格若是公式，讀算好的值（跟 Sheets API 回的一樣）。Google 匯出的 xlsx 有快取值。
        wb = openpyxl.load_workbook(where, data_only=True, read_only=False)
        sheets, info = scan_book(wb.worksheets, formulas)
        wb.close()
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(where)).strftime("%Y-%m-%d %H:%M")
        return sheets, mtime, report, info
    if kind == "gsheet":
        sid = gsheet_id(where)
        if not sid:
            raise ValueError("這個 .gsheet 裡找不到試算表 ID")
    else:
        sid = where
    meta = g.meta(sid)
    worksheets, progress_rows, formulas = g.worksheets(sid)
    sheets, info = scan_book(worksheets, formulas)
    return sheets, local_time(meta["modifiedTime"]), self_report(progress_rows), info


# ------------------------------------------------------------------- 主程式
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.yml")
    ap.add_argument("--out", default="data/progress.json")
    ap.add_argument("--source", action="append", default=None,
                    help="蓋掉設定檔裡的掃描來源。可給多次。可以是：Google 雲端硬碟下載的 .zip、"
                         "資料夾路徑、或 Google 雲端硬碟資料夾網址／ID")
    ap.add_argument("--allow-empty", action="store_true",
                    help="來源讀不到、或一份展表都沒掃到時照樣產出（只給第一次接線用）")
    a = ap.parse_args()

    cfg = yaml.safe_load(open(a.config, encoding="utf-8")) or {}
    sources = a.source or [os.environ.get("SCAN_SOURCE") or cfg.get("掃描來源")]
    sources = [str(x).strip() for x in sources if x]
    if not sources:
        sys.exit("設定檔沒有『掃描來源』，也沒有給 --source 或 SCAN_SOURCE。")
    prefix = cfg.get("展表檔名前綴", DEFAULT_PREFIX)
    prefix = "" if prefix is None else str(prefix)
    version = str(cfg.get("版本", ""))

    stat = {"候選": 0, "非前綴略過": 0, "讀不開": 0, "沒有判定欄": 0, "雲端xlsx未讀": 0, "無自報可對": 0}
    warnings = []
    g = None
    found, folder_names = [], []

    def fail_or_warn(msg):
        if not a.allow_empty:
            # 故意讓 pipeline 失敗。看板寧可不更新，也不要顯示過期或空白的數字
            # 卻長得像真的——有人會照著它做決定。
            sys.exit(msg)
        print("警告：" + msg, file=sys.stderr)

    for src in sources:
        if os.path.isfile(src) and src.lower().endswith(".zip"):
            try:
                found += discover_zip(src, prefix, stat)
                top = zip_top_folder(src)
            except zipfile.BadZipFile:
                fail_or_warn(f"這不是能打開的 zip：{src}")
                continue
            folder_names.append(top or os.path.basename(src))
        elif os.path.isdir(src):
            found += discover_local(src, prefix, stat)
            folder_names.append(os.path.basename(os.path.normpath(src)))
            if any(k == "gsheet" for _, k, _ in found) and g is None:
                g = Google()
        else:
            fid = drive_folder_id(src)
            if not fid:
                fail_or_warn(f"讀不到展表來源：{src}\n"
                             "  不是存在的資料夾或 zip，也不像 Google 雲端硬碟資料夾的網址或 ID。\n"
                             "  runner 有掛到那個資料夾嗎？或是 SCAN_SOURCE 有填對嗎？")
                continue
            if g is None:
                g = Google()
            try:
                name = g.meta(fid)["name"]
            except PermissionError as ex:
                fail_or_warn(f"讀不到 Google 雲端硬碟資料夾 {fid}：\n  {ex}")
                continue
            folder_names.append(name)
            found += discover_drive(g, fid, prefix, stat)

    for fname in folder_names:
        # zip 檔名（drive-download-…）不是資料夾名，不拿來對版本
        if version and fname and not fname.lower().endswith(".zip") and not fname.startswith(version):
            w = f"設定檔的版本是「{version}」，但掃的資料夾叫「{fname}」——兩邊對不上，是不是改了一邊忘了另一邊？"
            warnings.append(w)
            print("警告：" + w, file=sys.stderr)
    src = "；".join(sources)
    folder_name = folder_names[0] if len(folder_names) == 1 else None

    books, mismatches = [], []
    for name, kind, where in found:
        try:
            sheets, mtime, report, info = read_book(kind, where, g)
        except Exception as ex:
            stat["讀不開"] += 1
            print(f"跳過（讀不開）{name}：{ex}", file=sys.stderr)
            continue
        title = (re.sub(r"^" + re.escape(prefix) + r"|\.(xlsx|gsheet)$", "", name) if prefix
                 else re.sub(r"\.(xlsx|gsheet)$", "", name))
        # 範圍是照展表「測試進度」公式來的。這幾種狀況都要上看板讓人看，而且要在「這本沒東西就跳過」
        # 之前就記下來——不然公式全對不到的那本會從看板上安靜消失，連警告都沒有。
        def warn(w):
            warnings.append(w)
            print("警告：" + w, file=sys.stderr)
        if info["範圍"].startswith("自動（"):
            warn(f"「{title}」的測試進度公式點到的分頁一張都對不到（{info['進度表點到但沒這分頁'] + info['進度表點到但掃不到']}），"
                 "看板改成把這本所有有判定欄的分頁都算——多半是分頁改了名、公式沒跟著改，請展表負責人確認。")
        # 公式沒點到、卻有人在裡面測（已測 > 0）的分頁：看板沒算，這是真的會少算的風險。
        # 已測 0 的（空模板）只留在 progress.json 的展表資料裡，不吵。
        tested_extra = [x for x in info["未掛進進度表"] if x["已測"]]
        if tested_extra:
            desc = "、".join(f"{x['分頁']}（已測 {x['已測']}／母數 {x['母數']}）" for x in tested_extra)
            warn(f"「{title}」有 {len(tested_extra)} 個分頁有人測了、但沒掛進「測試進度」的公式，看板沒算：{desc}。"
                 "要算的話請展表負責人把它加進測試進度的公式。")
        if info["進度表點到但沒這分頁"] and not info["範圍"].startswith("自動（"):
            warn(f"「{title}」的測試進度公式點到分頁 {info['進度表點到但沒這分頁']}，但這本裡沒有這張分頁——改名了？")
        if info["進度表點到但掃不到"] and not info["範圍"].startswith("自動（"):
            warn(f"「{title}」的測試進度公式點到分頁 {info['進度表點到但掃不到']}，"
                 "但掃描器在裡面找不到判定欄（前 15 列沒有「確認…」表頭），這張沒算。")
        if not sheets:
            stat["沒有判定欄"] += 1
            print(f"跳過（沒有『確認…』判定欄的分頁）{name}", file=sys.stderr)
            continue
        agg = {"Y": 0, "N": 0, "X": 0, "未測試": 0, "其他": 0}
        for s in sheets:
            for k in agg:
                agg[k] += s["計"][k]
        # 同一件事展表自己也在算（測試進度分頁）。對一次，不一致就上看板吵，不要安靜地少算。
        if report is None:
            stat["無自報可對"] += 1
        else:
            for d in compare_report(report, agg):
                mismatches.append({"檔": title, **d})
        books.append({
            "檔名": name,
            "標題": title,
            "分頁": sheets,
            "計": agg,
            "母數": sum(agg.values()) - agg["X"] - agg["其他"],
            "已測": agg["Y"] + agg["N"],
            "更新時間": mtime,
            "自報": report,
            "範圍": info["範圍"],
            "未掛進進度表": info["未掛進進度表"],
            "進度表點到但沒這分頁": info["進度表點到但沒這分頁"],
            "進度表點到但掃不到": info["進度表點到但掃不到"],
        })

    if stat["讀不開"] and not a.allow_empty:
        # 少一份展表的看板跟全零的看板一樣危險：數字看起來正常，其實缺了一塊。
        sys.exit(f"有 {stat['讀不開']} 份展表讀不開（上面有逐份原因），不產出看板。\n"
                 "  多半是那份沒分享給 service account、Google 配額、或檔案壞了。修好再跑；"
                 "第一次接線可以先加 --allow-empty。")
    seen_titles = {}
    for bk in books:
        seen_titles.setdefault(bk["標題"], []).append(bk["檔名"])
    for t, names in seen_titles.items():
        if len(names) > 1:
            w = f"有 {len(names)} 份展表標題都叫「{t}」，會重複計算：{names}"
            warnings.append(w)
            print("警告：" + w, file=sys.stderr)

    if not books:
        msg = (f"一份展表都沒掃到：{src}\n"
               f"  候選檔 {stat['候選']}（檔名不是「{prefix}」開頭而略過 {stat['非前綴略過']}、"
               f"讀不開 {stat['讀不開']}、沒有判定欄 {stat['沒有判定欄']}、雲端 xlsx 未讀 {stat['雲端xlsx未讀']}）\n"
               "  資料夾是對的嗎？展表是不是還沒建、或是檔名前綴改了？")
        if not a.allow_empty:
            # 資料夾在、裡面卻沒有展表，跟資料夾不在一樣：不產出全零的看板。
            sys.exit(msg)
        print("警告：" + msg, file=sys.stderr)

    tot = {"Y": 0, "N": 0, "X": 0, "未測試": 0, "其他": 0}
    for bk in books:
        for k in tot:
            tot[k] += bk["計"][k]
    base = sum(tot.values()) - tot["X"] - tot["其他"]

    data = {
        "版本": version,
        "站台標題": cfg.get("站台標題", "測試進度"),
        "產出時間": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "來源": src if not folder_name or folder_name in src else f"{folder_name}（{src}）",
        "展表數": len(books),
        "總計": tot,
        "母數": base,
        "已測": tot["Y"] + tot["N"],
        "進度": round((tot["Y"] + tot["N"]) * 100 / base, 1) if base else 0,
        "展表": books,
        "風險": {
            "判N無備註": [{"檔": bk["標題"], "分頁": s["分頁"], **x}
                          for bk in books for s in bk["分頁"] for x in s["blank"]],
            "要先問": [{"檔": bk["標題"], "分頁": s["分頁"], **x}
                       for bk in books for s in bk["分頁"] for x in s["star"]],
            "判定值異常": [{"檔": bk["標題"], "分頁": s["分頁"], **x}
                           for bk in books for s in bk["分頁"] for x in s["其他值"]],
            "自報不符": mismatches,
        },
        "掃描統計": stat,
        "警告": warnings,
    }
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1)
    r = data["風險"]
    print(f"{len(books)} 份展表｜母數 {base}｜已測 {data['已測']}（{data['進度']}%）"
          f"｜異常 {tot['N']}｜判N無備註 {len(r['判N無備註'])}"
          f"｜★要先問 {len(r['要先問'])}｜判定值異常 {len(r['判定值異常'])}"
          f"｜分頁 {sum(len(bk['分頁']) for bk in books)}"
          f"｜未掛進進度表 {sum(len(bk['未掛進進度表']) for bk in books)} 分頁"
          f"｜自報不符 {len(mismatches)} 筆（{stat['無自報可對']} 份沒自報可對）"
          + (f"｜Google API 呼叫 {g.calls} 次" if g else ""))


if __name__ == "__main__":
    main()
