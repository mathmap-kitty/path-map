"""Compress the three 體育班 sources into one payload and inline it into sport.html.

Sources (all in this folder, all plain JSON lists):

  sport_src_quota.json   甄審甄試名額彙整表 —— 從《運動成績優良學生升學輔導甄審甄試簡章》
                         PDF 的名額彙整表抽出。一年換一次，手動更新。
  sport_src_single.json  各校運動績優單獨招生 —— 從大專校院單獨招生試務資訊查詢系統抓。
                         各校陸續填報，所以由 fetch_sport.py 每週重抓。
  sport_src_apply.json   個人申請中門檻寬鬆的校系 —— 從 ../歷年個申簡章.xlsx 篩出。

輸出 ../sport.html：單一檔案、資料內嵌、可直接上傳。

字串（校名、系名、運動種類、備註、網址）重複度很高，所以做成字串表用索引參照，
整個 payload 從 800 KB 降到約 320 KB。
"""
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

BUILD = Path(__file__).parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else BUILD.parent / "sport.html"


class Table:
    """字串表：同一個字串只存一次，其他地方用索引指過去。"""

    def __init__(self):
        self.items, self.index = [], {}

    def __call__(self, value):
        value = value or ""
        if value not in self.index:
            self.index[value] = len(self.items)
            self.items.append(value)
        return self.index[value]


def load(name):
    return json.loads((BUILD / name).read_text(encoding="utf-8"))


def main():
    quota = load("sport_src_quota.json")
    single = load("sport_src_single.json")
    apply_ = load("sport_src_apply.json")

    sch, dep, sp, memo, url, sex, reg, ex, art = (Table() for _ in range(9))

    Q = [[0 if r["rt"] == "甄審" else 1, sp(r["sp"]), sch(r["sch"]), dep(r["dep"]),
          r["code"], sex(r["sex"]), r["n"], memo(r["memo"])] for r in quota]
    S = [[sch(r["sch"]), dep(r["dep"]), sp(r["sp"]), sex(r["sex"]), r["n"],
          reg(r["reg"]), ex(r["exam"]), art(r["art"]), memo(r["memo"]), url(r["url"])]
         for r in single]

    # 運動種類清單以甄審甄試表的官方名稱為準（單招欄位是各校自填的自由文字），
    # 依名額排序，讓下拉選單一開始就是熱門項目。
    by_quota = Counter()
    for r in quota:
        by_quota[r["sp"]] += r["n"]
    sports = [name for name, _ in by_quota.most_common() if name]

    # 跨管道對照：一個運動種類 / 一所學校在四條路上各有多少。
    # 單招欄位一列可能同時列好幾種運動且共用名額，加總會重複計算，所以只報「校系數」。
    SP = []
    for s in sports:
        hits = [r for r in single if s in r["sp"]]
        schools = sorted({r["sch"] for r in quota if r["sp"] == s})
        SP.append([s,
                   sum(r["n"] for r in quota if r["rt"] == "甄審" and r["sp"] == s),
                   sum(r["n"] for r in quota if r["rt"] == "甄試" and r["sp"] == s),
                   len(hits), len(schools),
                   "、".join(schools[:12]) + ("…" if len(schools) > 12 else "")])

    names = sorted({r["sch"] for r in quota} | {r["sch"] for r in single}
                   | {r["sch"] for r in apply_})
    SC = []
    for n in names:
        kinds = sorted({r["sp"] for r in quota if r["sch"] == n})
        SC.append([n,
                   sum(r["n"] for r in quota if r["rt"] == "甄審" and r["sch"] == n),
                   sum(r["n"] for r in quota if r["rt"] == "甄試" and r["sch"] == n),
                   sum(r["n"] for r in single if r["sch"] == n),
                   sum(r["quota"] or 0 for r in apply_ if r["sch"] == n),
                   "、".join(kinds[:14]) + ("…" if len(kinds) > 14 else "")])

    payload = {
        "Q": Q, "S": S, "A": apply_, "SP": SP, "SC": SC, "sports": sports,
        "T": {"sch": sch.items, "dep": dep.items, "sp": sp.items, "memo": memo.items,
              "url": url.items, "sex": sex.items, "reg": reg.items,
              "ex": ex.items, "art": art.items},
        "built": datetime.now().strftime("%Y-%m-%d"),
    }

    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    # payload 住在 <script type="application/json"> 裡，唯一能跳出來的只有字面的 </script>
    data = data.replace("</", "<\\/")

    template = (BUILD / "sport_template.html").read_text(encoding="utf-8")
    if "__DATA__" not in template:
        sys.exit("sport_template.html 找不到 __DATA__ 佔位字串")
    OUT.write_text(template.replace("__DATA__", data), encoding="utf-8")

    print(f"wrote {OUT}  {OUT.stat().st_size/1e6:.2f} MB")
    print(f"  甄審 {sum(r['n'] for r in quota if r['rt']=='甄審')} 名額 / "
          f"甄試 {sum(r['n'] for r in quota if r['rt']=='甄試')} 名額 / "
          f"單招 {sum(r['n'] for r in single)} 名額 / "
          f"個申 {sum(r['quota'] or 0 for r in apply_)} 名額")


if __name__ == "__main__":
    main()
