"""每週重抓運動績優單獨招生資訊，有變動才重建 sport.html。

各校的單招簡章是陸續公告的——12 月中旬第一批出來，之後一路補到 3 月。
教育部沒有統一的簡章入口，只有「大專校院單獨招生試務資訊查詢系統」這個
由各校自行填報的彙整頁，所以這支程式就盯著它。

做四件事：

1. 抓查詢系統的每一頁，解析成跟 sport_src_single.json 同樣形狀的紀錄。
2. 跟上一次的結果比對，列出新增／消失／欄位變動的校系。
3. 順便看一眼升學輔導網的公告列表，新公告（例如「116 學年度甄審甄試簡章」）
   會被標出來——那份是 PDF，要手動更新 sport_src_quota.json，程式只負責提醒。
4. 有變動就覆寫 sport_src_single.json、重跑 build_sport.py，並把摘要寫進
   sport_watch_log.md。沒變動就什麼都不動，只在 log 記一行。

抓不到網頁時**不會**清空既有資料——寧可維持舊資料並在 log 標紅，
也不要讓一次網路故障把整張表洗掉。

用法：
    python fetch_sport.py            照常執行
    python fetch_sport.py --dry-run  只比對、只印報告，不寫任何檔案
"""
import json
import re
import ssl
import subprocess
import sys
import time
import unicodedata
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

BUILD = Path(__file__).parent
SRC = BUILD / "sport_src_single.json"
NOTICES = BUILD / "sport_notices.json"
LOG = BUILD / "sport_watch_log.md"
STATE = BUILD / "sport_watch_state.json"
REPORTS = BUILD / "sport_reports"

# 檢核頻率：各校的單招簡章 12 月中旬才開始陸續公告，9-10 月幾乎不會動，
# 所以那兩個月每兩週看一次就夠；11 月起進入公告密集期，改成每週。
# 工作排程固定每週觸發，由這裡決定這一次要不要真的去抓。
def interval_days(today):
    return 14 if today.month in (9, 10) else 7

LIST_URL = "https://iss.ntus.edu.tw/open/sexam"
NOTICE_URL = "https://lulu.ntus.edu.tw/"
UA = "Mozilla/5.0 (path-map sport watcher; +https://github.com/mathmap-kitty)"
MAX_PAGES = 60          # 目前 18 頁，留很大的餘裕；靠「沒有資料列」自然停下
DRY = "--dry-run" in sys.argv


# 這兩個站的憑證少了 Subject Key Identifier 這個擴充欄位，Python 3.13 起預設開啟的
# VERIFY_X509_STRICT 會因此拒絕連線（curl、瀏覽器都連得上）。這裡只關掉「嚴格模式」，
# 憑證鏈本身仍照常驗證，不是 verify_mode = CERT_NONE。
def context():
    ctx = ssl.create_default_context()
    ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return ctx


SSL_CTX = context()


def get(url, tries=3):
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=45, context=SSL_CTX) as r:
                return r.read().decode("utf-8", "replace")
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if attempt == tries - 1:
                raise
            time.sleep(3 * (attempt + 1))


def text(fragment):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", fragment)).strip()


def norm(s):
    s = unicodedata.normalize("NFKC", str(s or ""))
    for bad, good in (("⽻", "羽"), ("⽥", "田"), ("⺠", "民"), ("⼿", "手"), ("壼", "壺")):
        s = s.replace(bad, good)
    return re.sub(r"\s+", " ", s).strip()


def clip(s, n):
    s = norm(s)
    return s[:n] + ("…" if len(s) > n else "")


def digits(s):
    m = re.search(r"\d+", str(s))
    return int(m.group()) if m else 0


def scrape():
    """把查詢系統整站讀成紀錄清單。每一頁 50 列，讀到空頁就停。"""
    records = []
    for page in range(1, MAX_PAGES + 1):
        html = get(f"{LIST_URL}?page={page}")
        before = len(records)
        current = None
        for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
            raw = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S)
            cells = [text(c) for c in raw]
            if len(cells) == 9 and re.match(r"^\d{4}$", cells[0]):
                current = {
                    "sch": norm(cells[1]), "dep": norm(cells[2]),
                    # 各校自填的運動種類欄位常混著代碼，例如「001籃球、002排球」
                    "sp": clip(re.sub(r"(?<![0-9])0*(\d{3})(?=[一-鿿])", "", norm(cells[3])), 120),
                    "sex": norm(cells[4]), "n": digits(cells[5]),
                    "reg": norm(cells[7]),
                    "exam": "", "art": "", "memo": "", "url": "",
                }
                records.append(current)
            elif current is not None and len(cells) >= 2:
                key, value = cells[0], " ".join(cells[1:]).strip()
                links = re.findall(r'href="(https?://[^"]+)"', " ".join(raw))
                if "學科考試日期" in key:
                    current["exam"] = clip(value, 70)
                elif "術科檢定日期" in key:
                    current["art"] = clip(value, 70)
                elif "備註" in key:
                    current["memo"] = clip(value, 170)
                elif "試務網頁" in key or "連結" in key:
                    current["url"] = links[0] if links else value
        if len(records) == before:
            break
    return records


def exam_digest(records):
    """「學科考試」欄位的分布。

    這一欄每年逐校都會變，而且它只講「學校有沒有自辦筆試」——顯示「無」不代表
    不採計學測。所以每次都把分布印出來，新年度簡章一到就看得出哪些學校改了做法，
    再逐校去翻簡章確認是採計學測還是自辦考試。
    """
    gsat = sum(1 for r in records if "學測" in r["exam"])
    none = sum(1 for r in records if not r["exam"] or r["exam"] in ("無", "-", "—"))
    own = len(records) - gsat - none
    return (f"提到學測 {gsat} 筆、自辦或另訂日期 {own} 筆、填「無」{none} 筆"
            "（「無」只表示沒有自辦筆試，仍可能採計學測，須逐校翻簡章）")


def key(r):
    return (r["sch"], r["dep"], r["sp"], r["sex"])


def diff(old, new):
    """回傳 (新增, 消失, 變動)。變動只看真的會影響選填的欄位。"""
    watched = ("n", "reg", "exam", "art", "url")
    a = {key(r): r for r in old}
    b = {key(r): r for r in new}
    added = [b[k] for k in b.keys() - a.keys()]
    gone = [a[k] for k in a.keys() - b.keys()]
    changed = []
    for k in a.keys() & b.keys():
        deltas = [(f, a[k][f], b[k][f]) for f in watched if a[k][f] != b[k][f]]
        if deltas:
            changed.append((b[k], deltas))
    added.sort(key=key)
    gone.sort(key=key)
    changed.sort(key=lambda x: key(x[0]))
    return added, gone, changed


def notices():
    """升學輔導網的公告標題，用來發現新一年度的甄審甄試簡章。"""
    try:
        html = get(NOTICE_URL)
    except Exception:
        return []
    seen, out = set(), []
    # 最新消息表格：第一欄是標題、<time> 是發布日期、按鈕的 href 是內文網址
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S):
        link = re.search(r'href="(https://lulu\.ntus\.edu\.tw/new/\d+/show)"', tr)
        if not link:
            continue
        url = link.group(1)
        if url in seen:
            continue
        cells = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)
        title = text(cells[0]) if cells else ""
        date = re.search(r'<time datetime="([\d-]+)"', tr)
        if title:
            seen.add(url)
            out.append({"title": title, "date": date.group(1) if date else "",
                        "url": url})
    return out


def write_log(lines):
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    body = f"\n## {stamp}\n\n" + "\n".join(lines) + "\n"
    if DRY:
        print(body)
        return
    if not LOG.exists():
        LOG.write_text("# 運動績優單招 每週檢核紀錄\n", encoding="utf-8")
    with LOG.open("a", encoding="utf-8") as f:
        f.write(body)


def due(today):
    """這一次該不該真的去抓。工作排程每週觸發，9-10 月改成兩週一次。"""
    if "--force" in sys.argv:
        return True, None
    state = json.loads(STATE.read_text(encoding="utf-8")) if STATE.exists() else {}
    last = state.get("last_checked")
    if not last:
        return True, None
    gap = (today - datetime.strptime(last, "%Y-%m-%d").date()).days
    need = interval_days(today)
    return gap >= need, gap


def stamp_state(today, summary):
    if DRY:
        return
    STATE.write_text(json.dumps({
        "last_checked": today.strftime("%Y-%m-%d"),
        "cadence": f"{interval_days(today)} 天",
        "summary": summary,
    }, ensure_ascii=False, indent=1), encoding="utf-8")


def write_report(today, lines, changed):
    """每次實際檢核都留一份書面統計，有沒有變動都要寫——這樣才知道它有在跑。"""
    if DRY:
        return None
    REPORTS.mkdir(exist_ok=True)
    f = REPORTS / f"{today.strftime('%Y-%m-%d')}_單招檢核.md"
    head = [f"# 運動績優單獨招生 檢核報告　{today.strftime('%Y-%m-%d')}", "",
            f"檢核頻率：目前每 {interval_days(today)} 天一次"
            f"（9–10 月每兩週，11 月起每週）", "",
            ("## 有變動，資料已更新" if changed else "## 本次無變動"), ""]
    f.write_text("\n".join(head + lines) + "\n", encoding="utf-8")
    return f


def main():
    today = datetime.now().date()
    ok, gap = due(today)
    if not ok:
        print(f"距離上次檢核只有 {gap} 天，未達 {interval_days(today)} 天，本次跳過。")
        return 0

    old = json.loads(SRC.read_text(encoding="utf-8")) if SRC.exists() else []

    try:
        new = scrape()
    except Exception as e:
        write_log([f"- ⚠️ **抓取失敗**：{e}", "- 既有資料未更動，下週再試。"])
        print(f"抓取失敗：{e}", file=sys.stderr)
        return 1

    # 一次抓到的筆數掉超過三成，多半是網站改版或被擋，不是真的少了那麼多校系。
    if old and len(new) < len(old) * 0.7:
        write_log([f"- ⚠️ **筆數異常**：抓到 {len(new)} 筆，上次 {len(old)} 筆，落差過大。",
                   "- 已中止更新以免洗掉正確資料，請人工確認查詢系統是否改版。"])
        print(f"筆數異常 {len(old)} → {len(new)}，已中止", file=sys.stderr)
        return 1

    added, gone, changed_rows = diff(old, new)

    seen_notices = json.loads(NOTICES.read_text(encoding="utf-8")) if NOTICES.exists() else []
    seen_urls = {n["url"] for n in seen_notices}
    fresh = [n for n in notices() if n["url"] not in seen_urls]

    lines = [f"- 查詢系統 {len(new)} 筆、名額合計 {sum(r['n'] for r in new)}"
             f"（上次 {len(old)} 筆）",
             "- 學科要求分布：" + exam_digest(new)]

    if added:
        lines.append(f"- **新增 {len(added)} 筆**")
        for r in added[:40]:
            lines.append(f"    - {r['sch']}　{r['dep']}　{r['sp']}　{r['n']} 名　"
                         f"報名 {r['reg'] or '—'}　學科 {r['exam'] or '無'}")
        if len(added) > 40:
            lines.append(f"    - …另有 {len(added)-40} 筆，詳見 sport.html")
    if gone:
        lines.append(f"- **消失 {len(gone)} 筆**")
        for r in gone[:20]:
            lines.append(f"    - {r['sch']}　{r['dep']}　{r['sp']}")
        if len(gone) > 20:
            lines.append(f"    - …另有 {len(gone)-20} 筆")
    if changed_rows:
        lines.append(f"- **異動 {len(changed_rows)} 筆**")
        labels = {"n": "名額", "reg": "報名日期", "exam": "學科考試", "art": "術科檢定", "url": "試務網頁"}
        for r, deltas in changed_rows[:30]:
            for f, o, n in deltas:
                lines.append(f"    - {r['sch']}　{r['dep']}：{labels[f]} {o or '—'} → {n or '—'}")
        if len(changed_rows) > 30:
            lines.append(f"    - …另有 {len(changed_rows)-30} 筆")
    if fresh:
        lines.append(f"- 📣 **升學輔導網有 {len(fresh)} 則新公告**（甄審甄試簡章是 PDF，需手動更新名額表）")
        for n in fresh:
            lines.append(f"    - {n.get('date','')}　[{n['title']}]({n['url']})")
    if not (added or gone or changed_rows or fresh):
        lines.append("- 無變動。")

    print("\n".join(lines))

    if DRY:
        write_log(lines)
        return 0

    if added or gone or changed_rows:
        SRC.write_text(json.dumps(new, ensure_ascii=False, indent=0), encoding="utf-8")
        result = subprocess.run([sys.executable, str(BUILD / "build_sport.py")],
                                capture_output=True, text=True, encoding="utf-8")
        print(result.stdout or "", end="")
        if result.returncode == 0:
            lines.append(f"- ✅ 已重建 sport.html")
        else:
            lines.append(f"- ⚠️ 重建 sport.html 失敗：{(result.stderr or '').strip()[:300]}")
    if fresh:
        NOTICES.write_text(json.dumps(seen_notices + fresh, ensure_ascii=False, indent=1),
                           encoding="utf-8")

    changed = bool(added or gone or changed_rows)
    report = write_report(today, lines, changed)
    stamp_state(today, ("有變動：新增 %d、消失 %d、異動 %d" % (len(added), len(gone), len(changed_rows)))
                if changed else "無變動")
    write_log(lines + ([f"- 📄 書面統計：`{report.name}`"] if report else []))
    return 0


if __name__ == "__main__":
    # 排程用 pythonw 之類的無主控台方式執行時 sys.stdout 會是 None，
    # 這時把輸出導到檔案，才不會一 print 就整支掛掉。
    if sys.stdout is None or sys.stderr is None:
        sink = (BUILD / "sport_watch_stdout.log").open("a", encoding="utf-8")
        sys.stdout = sys.stderr = sink
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception as fatal:  # 排程跑的時候沒人看得到 traceback，寫進 log 才找得回來
        import traceback
        write_log([f"- ⚠️ **執行失敗**：{fatal}",
                   "```", traceback.format_exc().strip()[-1200:], "```"])
        raise
