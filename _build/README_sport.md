# 體育班升學地圖（sport.html / sport-a4.html）

產出兩個檔：

| 檔案 | 是什麼 |
|---|---|
| `../sport.html` | 查詢頁。四條管道的名額、條件、時程，2,493 筆逐筆志願可依運動種類跨管道對照。單一檔案、資料內嵌。 |
| `../sport-a4.html` | 一張 A4 的班會決策單，按瀏覽器列印即為單頁。 |
| `../sport-a4.pdf` | 上面那張的 PDF，直接印給學生。 |

## 為什麼要做這頁

原本的 `index.html` 只把「單獨招生」當成一個小管道帶過，但體育班真正的名額分布是：

| 管道 | 志願數 | 名額 |
|---|---:|---:|
| 運動績優甄審 | 1,292 | 2,384 |
| 運動績優甄試 | 205 | 240 |
| 各校運動績優單獨招生 | 896 | 3,263 |
| 個人申請（門檻寬鬆校系） | 100 | 1,514 |

個人申請是四條裡名額最少、學科要求最高的一條。這頁的存在就是把順序講清楚。

## 資料從哪來

| 來源檔 | 從哪抓 | 多久換一次 |
|---|---|---|
| `sport_src_quota.json` | 《運動成績優良學生升學輔導甄審甄試簡章》PDF 第 48–112 頁的名額彙整表 | 一年一次，**手動** |
| `sport_src_single.json` | [大專校院單獨招生試務資訊查詢系統](https://iss.ntus.edu.tw/open/sexam) | **每週自動** |
| `sport_src_apply.json` | `../歷年個申簡章.xlsx` 的 115 分頁，篩「校系種類＝體育」＋系名含運動／休閒／觀光且門檻寬鬆者 | 一年一次，手動 |

三個都是同一種形狀的 JSON list，欄位名稱直接對應 `build_sport.py` 讀的鍵。

## 指令

```
python build_sport.py            # 重建 ../sport.html
python fetch_sport.py            # 重抓單招資料，有變動才重建
python fetch_sport.py --dry-run  # 只比對、只印報告，不寫檔
```

A4 那張改完之後重出 PDF（需要 Chrome）：

```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --headless=new --disable-gpu ^
  --no-pdf-header-footer --print-to-pdf="D:\path-map\sport-a4.pdf" "file:///D:/path-map/sport-a4.html"
```

## 每週自動檢核

Windows 工作排程器裡有一個工作：

* 名稱：**path-map 體育班單招每週檢核**
* 時間：每週日 09:00（沒開機就等下次開機補跑）
* 動作：`pythonw.exe fetch_sport.py`（工作目錄 `D:\path-map\_build`）

它做四件事：

1. 抓單招查詢系統的每一頁，解析成紀錄。
2. 跟 `sport_src_single.json` 比對，列出**新增／消失／名額或日期異動**的校系。
3. 看一眼升學輔導網的公告列表，有新公告就列出來——例如「116 學年度甄審甄試簡章」。
   那份是 PDF，程式只負責提醒，名額表仍要人工更新 `sport_src_quota.json`。
4. 有變動就覆寫來源檔、重跑 `build_sport.py`，摘要寫進 `sport_watch_log.md`。

看結果就開 `sport_watch_log.md`，最新的在最下面。

兩道保險：

* 抓不到網頁時**不覆寫**既有資料，只在 log 標記，下週再試。
* 抓到的筆數比上次少三成以上就中止更新——那通常代表網站改版或被擋，不是真的少了那麼多校系。

### 重灌或換電腦之後怎麼裝回來

工作排程本身不在 repo 裡（它是 Windows 的設定，不是檔案），但註冊腳本在。
clone 完 repo、確認 Python 裝好且在 PATH 上，然後在 repo 根目錄跑：

```
powershell -ExecutionPolicy Bypass -File _build\register_sport_watch.ps1
```

它會自己找 `pythonw.exe`、用 `_build` 當工作目錄註冊每週日 09:00 的工作，
並印出下次執行時間與「先手動跑一次確認」的指令。想換時間就加參數：

```
powershell -ExecutionPolicy Bypass -File _build\register_sport_watch.ps1 -DayOfWeek Monday -At 07:30
```

停掉排程：

```
powershell -ExecutionPolicy Bypass -File _build\register_sport_watch.ps1 -Unregister
```

或直接下：

```
Unregister-ScheduledTask -TaskName 'path-map 體育班單招每週檢核' -Confirm:$false
```

### 註冊腳本裡三個「別拿掉」的設定

都是實測踩出來的坑，改腳本前先看一眼：

* **用 `pythonw.exe`，不要用 `run_sport_watch.bat`。** Task Scheduler 執行 .bat 時，
  cmd 的主控台一關就回 `0xC000013A`（Ctrl-C 結束），任務等於沒跑成。
  `pythonw` 沒有主控台就沒這問題。`run_sport_watch.bat` 留著給你手動雙擊用。
* **`-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries`。**
  `New-ScheduledTaskSettingsSet` 預設「使用電池時不啟動」，筆電沒插電時工作會一直卡在
  `Queued` 不執行——而且 `LastTaskResult` 還是顯示 `0`，非常難發現。
* **不要加 `-RunOnlyIfNetworkAvailable`。** 同樣會卡在 `Queued`。
  抓不到網頁時 `fetch_sport.py` 本來就會記 log 並保留舊資料，讓它跑下去再失敗比較好。

### 怎麼確認它真的有在跑

不要只看 `LastTaskResult`——卡在 `Queued` 的時候它也是 0。要看兩個地方：

```
Get-ScheduledTaskInfo -TaskName 'path-map 體育班單招每週檢核' | Select-Object LastRunTime,LastTaskResult
(Get-ScheduledTask -TaskName 'path-map 體育班單招每週檢核').State
```

`State` 應該是 `Ready`（執行中會短暫變 `Running`）。**真正的證據是 `sport_watch_log.md`
最下面有沒有多出一段當天的紀錄**——有寫進去才代表程式真的跑完了。

## 一個環境細節

`iss.ntus.edu.tw` 和 `lulu.ntus.edu.tw` 的憑證少了 Subject Key Identifier 這個擴充欄位，
Python 3.13 起預設開啟的 `VERIFY_X509_STRICT` 會直接拒絕連線（瀏覽器和 curl 都連得上）。
`fetch_sport.py` 只關掉這個嚴格模式旗標，憑證鏈本身照常驗證。

## 「學科考試」這一欄有陷阱

單招的學科要求有**兩種完全不同的東西**：

1. **採計學測** —— 不另外考，但要達到該校系的學測檢定標準。
2. **學校自辦學科考試** —— 另外訂日期到校應考。

查詢系統的「學科考試日期」欄位講的是第 2 種。**顯示「無」只代表沒有自辦筆試，不代表不採計學測**——成功大學 14 個系該欄都是「無」，簡章卻要求學測檢定。

所以 `fetch_sport.py` 每次都會印一行「學科要求分布」（提到學測 N 筆／自辦或另訂日期 N 筆／填「無」N 筆）。新年度簡章一到，先看這行有沒有大變動，再逐校翻簡章確認到底是哪一種。**這一欄不能直接照抄給學生。**

## 每年要手動做的事

1. 新年度甄審甄試簡章公告後（約 1 月底，log 會提醒），從 PDF 的名額彙整表重建 `sport_src_quota.json`。
2. 新年度個申簡章進 `../歷年個申簡章.xlsx` 後，重出 `sport_src_apply.json`。
3. **逐校核對單招的「學科考試」與「學測檢定」**（見上一節），頁面上的例子（中興電機、清大運科、成大）要跟著更新。
4. `sport_template.html` 裡的時程區塊與 `sport-a4.html` 的日期要跟著換年度。
5. 跑 `python build_sport.py`，並重出 A4 的 PDF。

## 條文上容易寫錯的兩點

寫這頁時查證出來的，改內容前先看一眼：

* **甄審不是「有國際賽成績就行」。** 辦法第 4 條的前提句是「依國家代表隊教練與選手選拔培訓及參賽處理辦法規定，**代表國家參加國際運動賽會**」——國家代表隊身分、清單內賽會、名次達標，三者缺一不可。另有第 5 條的參賽國（地區）與隊（人）數門檻。
* **全國賽成績不能甄審。** 全運會、全民運、全中運、中等學校運動聯賽都是第 6 條的國內賽會，只能走**甄試**。體育班學生大多數落在這裡。

## 資料年度怎麼標

頁面上每一塊都標了自己的年度，因為三個來源不是同時換的：

| 來源 | 檔案 | 誰更新 |
|---|---|---|
| 甄審甄試名額 | `sport_src_quota.json` | 手動（每年 1–2 月簡章 PDF 出來後） |
| 各校單獨招生 | `sport_src_single.json` | `fetch_sport.py` 自動 |
| 個人申請校系 | `sport_src_apply.json` | 手動（每年個申簡章出來後） |

年度值放在 `sport_years.json`：

```json
{ "quota": "115", "apply": "115", "single": "115" }
```

- **`single` 不用手動改** —— `build_sport.py` 會從報名日期推算（115 學年度的報名橫跨
  114 年 12 月到 115 年 5 月，取出現過的最大民國年），推不出來才沿用設定檔的值。
- **`quota` 和 `apply` 換年度時要自己改這裡**，改完跑 `python build_sport.py`。

三份年度不一致時（例如單招已經是 116、另兩份還是 115），較新的那一份徽章會變成
青綠色，一眼看得出哪一塊已經換年度、哪一塊還是舊資料。頁尾的資料來源引註也會跟著變。
