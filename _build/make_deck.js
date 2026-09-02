/* 產生「體育班升學說明」簡報（.pptx），班會與親師座談用。 */
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const BUILD = __dirname;
const tiers = JSON.parse(fs.readFileSync(path.join(BUILD, 'doc_tiers.json'), 'utf8'));

const BERRY = '8C2740';      // 主色，沿用 path-map 網站
const DEEP = '5A1526';       // 深色底
const CREAM = 'F7F2EE';
const INK = '2B2B2B';
const MUT = '7A6B65';
const TEAL = '1F6F78';
const GOLD = 'B8862B';
const WHITE = 'FFFFFF';

const HEAD = '微軟正黑體';
const BODY = '微軟正黑體';

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';           // 13.3 × 7.5
const W = 13.3, H = 7.5, M = 0.7;

pres.defineSlideMaster({
  title: 'LIGHT',
  background: { color: WHITE },
});

function titleOf(s, text, sub) {
  s.addText(text, {
    x: M, y: 0.42, w: W - 2 * M, h: 0.75, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 34, bold: true, color: BERRY, align: 'left',
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.18, w: W - 2 * M, h: 0.42, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 14, color: MUT,
    });
  }
}

/* 卡片：淡色底，不用邊條 */
function card(s, o) {
  s.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.08,
    fill: { color: o.fill || CREAM }, line: { color: o.line || 'E7DCD6', width: 1 },
    shadow: { type: 'outer', color: '8C7A72', blur: 8, offset: 2, angle: 90, opacity: 0.12 },
  });
}

function stat(s, o) {
  card(s, { x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill });
  s.addText(o.big, {
    x: o.x, y: o.y + 0.22, w: o.w, h: 0.95, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: o.size || 50, bold: true, color: o.color || BERRY, align: 'center',
  });
  s.addText(o.label, {
    x: o.x + 0.15, y: o.y + o.h - 0.92, w: o.w - 0.3, h: 0.78, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 13, color: o.labelColor || INK, align: 'center', valign: 'top',
  });
}

/* ── 1 封面 ── */
let s = pres.addSlide();
s.background = { color: DEEP };
s.addText('體育班升學說明', {
  x: M, y: 2.15, w: W - 2 * M, h: 1.1, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 54, bold: true, color: WHITE,
});
s.addText('以「個人申請」為主要管道', {
  x: M, y: 3.3, w: W - 2 * M, h: 0.6, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 26, color: 'E7B9C6',
});
s.addText('115 學年度資料　｜　給學生與家長', {
  x: M, y: 4.15, w: W - 2 * M, h: 0.4, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 15, color: 'C9A6B0',
});
s.addNotes('這份說明針對沒有比賽成績的體育班學生。先講三條運動管道為什麼走不通，再講個人申請該怎麼準備。');

/* ── 2 四條路，三條關著 ── */
s = pres.addSlide();
titleOf(s, '體育班有四條路，三條需要比賽成績');
const routes = [
  ['運動績優甄審', '國家代表隊\n國際賽會名次', 'X', BERRY],
  ['運動績優甄試', '全國賽前八名\n全運會、全中運', 'X', BERRY],
  ['各校單獨招生', '多數要縣市級以上\n成績或校隊資歷', 'X', BERRY],
  ['大學個人申請', '學測 + 備審面試\n不看比賽成績', 'O', TEAL],
];
routes.forEach((r, i) => {
  const x = M + i * 3.05, w = 2.8;
  const open = r[2] === 'O';
  card(s, { x, y: 1.85, w, h: 3.6, fill: open ? 'E3F0F0' : CREAM, line: open ? TEAL : 'E7DCD6' });
  s.addText(r[2] === 'O' ? '可以走' : '走不通', {
    x, y: 2.1, w, h: 0.42, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 15, bold: true, color: open ? TEAL : 'B08A94', align: 'center',
  });
  s.addText(r[0], {
    x: x + 0.15, y: 2.65, w: w - 0.3, h: 0.6, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 21, bold: true, color: open ? TEAL : BERRY, align: 'center',
  });
  s.addText(r[1], {
    x: x + 0.2, y: 3.4, w: w - 0.4, h: 1.5, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, color: INK, align: 'center', lineSpacing: 22,
  });
});
s.addText('沒有比賽成績時，能走的是第四條。', {
  x: M, y: 5.75, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 20, bold: true, color: DEEP,
});
s.addNotes('甄審要的是國家代表隊身分，甄試要全國賽前八名，單招各校自訂但多數也要成績。');

/* ── 3 實際可填的範圍 ── */
s = pres.addSlide();
titleOf(s, '先搞清楚：實際能填的沒有想像中多', '個人申請裡體育班打得到的落點');
stat(s, { x: M, y: 2.0, w: 3.5, h: 2.5, big: '100', label: '個校系\n看起來打得到' });
s.addText('－', { x: 4.35, y: 2.85, w: 0.6, h: 0.8, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 40, color: MUT, align: 'center' });
stat(s, { x: 4.95, y: 2.0, w: 3.5, h: 2.5, big: '17', color: 'A03D2E',
  label: '個填不了\n15 個青年儲蓄帳戶組\n2 個公費生', fill: 'F7E9E4' });
s.addText('＝', { x: 8.55, y: 2.85, w: 0.6, h: 0.8, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 36, color: MUT, align: 'center' });
stat(s, { x: 9.15, y: 2.0, w: 3.45, h: 2.5, big: '83', color: TEAL,
  label: '個系、1,495 個名額\n應屆生真正能填的', fill: 'E3F0F0' });
s.addText('「青年儲蓄帳戶組」是高中畢業後先去工作或體驗 2～3 年才能用的管道，應屆生填不了，而且每系名額多半只有 1 名。', {
  x: M, y: 5.0, w: W - 2 * M, h: 0.9, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 15, color: INK, lineSpacing: 26,
});
s.addNotes('這一點很多資料會漏掉，把青儲組當成「學測佔 0% 的好機會」介紹，其實應屆生根本填不了。');

/* ── 4 主力科目 chart ── */
s = pres.addSlide();
titleOf(s, '主力科目：國文第一，英文第二', '82 個校系中，各科在一階篩選出現的比例');
s.addChart(pres.ChartType.bar, [{
  name: '一階篩選會用到',
  labels: ['國文', '英文', '社會', '數學 B', '自然', '數學 A'],
  values: [81, 65, 28, 13, 12, 2],
}], {
  x: M, y: 1.75, w: 7.4, h: 4.6,
  barDir: 'bar', chartColors: [BERRY], showValue: true, dataLabelPosition: 'outEnd',
  dataLabelColor: INK, dataLabelFontSize: 13, dataLabelFontFace: BODY,
  dataLabelFormatCode: '0"%"',
  catAxisLabelColor: INK, catAxisLabelFontSize: 15, catAxisLabelFontFace: BODY,
  valAxisLabelColor: MUT, valAxisLabelFontSize: 11, valAxisMaxVal: 100,
  valGridLine: { color: 'EDE4E0', size: 1 }, catGridLine: { style: 'none' },
  showLegend: false, barGapWidthPct: 45,
});
card(s, { x: 8.4, y: 1.9, w: 4.2, h: 4.3 });
s.addText('怎麼用這張圖', {
  x: 8.65, y: 2.15, w: 3.7, h: 0.42, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 19, bold: true, color: BERRY,
});
s.addText([
  { text: '國文和英文是主力。', options: { bold: true, breakLine: true } },
  { text: '五成的系設國文檢定，八成的系一階會篩國文。', options: { breakLine: true } },
  { text: '', options: { breakLine: true } },
  { text: '社會排第三。', options: { bold: true, breakLine: true } },
  { text: '觀光、休閒、餐旅類常拿它當篩選科目。', options: { breakLine: true } },
  { text: '', options: { breakLine: true } },
  { text: '數學 A 只有 2 個系會用到。', options: { bold: true, color: 'A03D2E', breakLine: true } },
  { text: '數學 B 也只有 13%，集中在幾所公立體育系。' },
], {
  x: 8.65, y: 2.7, w: 3.7, h: 3.3, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 14, color: INK, lineSpacing: 24, valign: 'top',
});
s.addNotes('對體育班來說這是好消息：要顧的科目其實只有國文和英文，社會加減看。數學可以放掉。');

/* ── 5 級分對照 chart ── */
s = pres.addSlide();
titleOf(s, '每科多 1 級分，可填的校系差很多', '115 學年度實際篩選結果推算，共 81 個校系');
const acc = JSON.parse(fs.readFileSync(path.join(BUILD, 'ppt_acc.json'), 'utf8'));
s.addChart(pres.ChartType.line, [{
  name: '過得了一階的校系數',
  labels: acc.map((a) => a[0] + ' 級分'),
  values: acc.map((a) => a[1]),
}], {
  x: M, y: 1.85, w: 7.6, h: 4.4,
  chartColors: [BERRY], lineSize: 4, lineSmooth: false,
  showValue: true, dataLabelPosition: 't', dataLabelColor: BERRY,
  dataLabelFontSize: 14, dataLabelFontFace: BODY, dataLabelFontBold: true,
  catAxisLabelColor: INK, catAxisLabelFontSize: 13, catAxisLabelFontFace: BODY,
  valAxisLabelColor: MUT, valAxisLabelFontSize: 11,
  valGridLine: { color: 'EDE4E0', size: 1 }, catGridLine: { style: 'none' },
  showLegend: false,
});
card(s, { x: 8.55, y: 2.0, w: 4.05, h: 3.1, fill: 'F7E9E4', line: 'E4C8C0' });
s.addText('目標訂在\n每科 8 級分', {
  x: 8.75, y: 2.25, w: 3.65, h: 1.0, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 26, bold: true, color: BERRY, align: 'center', lineSpacing: 34,
});
s.addText('7 級分能填 25 系\n8 級分能填 43 系\n9 級分能填 60 系', {
  x: 8.75, y: 3.45, w: 3.65, h: 1.4, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 16, color: INK, align: 'center', lineSpacing: 28,
});
s.addText('「每科 8 級分」是一個看得見的目標，不是「盡力就好」。', {
  x: 8.55, y: 5.35, w: 4.05, h: 0.8, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 14, color: MUT, lineSpacing: 22,
});
s.addNotes('這條線不是平的。7 到 8 之間多 18 個系，8 到 9 之間多 17 個系。跟學生講具體數字比講「要努力」有用。');

/* ── 6 專長對照 ── */
s = pres.addSlide();
titleOf(s, '四個專長，各自能走哪裡', '第二階段「自辦專長術科」的系會限定收哪些運動項目');
const tal = [
  ['田徑', '臺中教大 運動專長組', '男 200 公尺 1 名\n女 100 公尺跨欄 1 名', '其餘田徑項目不收', GOLD],
  ['滑輪溜冰', '沒有任何一系收', '—', '只能走術科考試或一般組', 'A03D2E'],
  ['足球', '臺體大 運動專長組', '8 個名額', '四個專長裡名額最多的', TEAL],
  ['羽球', '臺中教大 運動專長組', '男 2 名、女 1 名', '名額極少', GOLD],
];
tal.forEach((r, i) => {
  const y = 1.85 + i * 1.16;
  card(s, { x: M, y, w: W - 2 * M, h: 1.0 });
  s.addText(r[0], { x: M + 0.25, y: y + 0.28, w: 1.7, h: 0.45, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 20, bold: true, color: r[4] });
  s.addText(r[1], { x: M + 2.05, y: y + 0.28, w: 3.3, h: 0.45, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 15, color: INK });
  s.addText(r[2], { x: M + 5.45, y: y + 0.14, w: 3.0, h: 0.72, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, color: INK, lineSpacing: 20 });
  s.addText(r[3], { x: M + 8.55, y: y + 0.28, w: 3.2, h: 0.45, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, color: MUT });
});
s.addText('專長對不上，不代表沒有路——下一頁那張牌對所有專長都開放。', {
  x: M, y: 6.55, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 17, bold: true, color: DEEP,
});
s.addNotes('彰化師大運動學系競技組雖然門檻低，但它要全運會或全中運前八名，沒有成績的學生不符合資格，已排除。');

/* ── 7 術科考試 ── */
s = pres.addSlide();
s.background = { color: DEEP };
s.addText('大學術科考試（體育組）', {
  x: M, y: 0.75, w: W - 2 * M, h: 0.8, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 36, bold: true, color: WHITE,
});
s.addText('不看比賽成績，只看你當天的體能。這是這份名單裡唯一完全由自己決定的分數。', {
  x: M, y: 1.6, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 17, color: 'E7B9C6',
});
const five = ['60 公尺\n立姿快跑', '20 秒\n反覆側步', '一分鐘\n屈膝仰臥起坐', '立定\n連續三次跳', '1600 公尺\n跑走'];
five.forEach((v, i) => {
  const x = M + i * 2.44;
  s.addShape(pres.ShapeType.roundRect, {
    x, y: 2.5, w: 2.24, h: 1.5, rectRadius: 0.1,
    fill: { color: '73203A' }, line: { color: '9E4560', width: 1 },
  });
  s.addText(v, { x: x + 0.1, y: 2.62, w: 2.04, h: 1.26, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 15, color: WHITE, align: 'center', valign: 'middle', lineSpacing: 22 });
});
s.addText([
  { text: '報名：', options: { bold: true, color: 'E7B9C6' } },
  { text: '高三上學期 10 月底至 11 月初，由學校集體報名，錯過沒有補救管道。', options: { color: WHITE, breakLine: true } },
  { text: '考試：', options: { bold: true, color: 'E7B9C6' } },
  { text: '1 月在國立體育大學，五項 T 分數加總，男女分開評比，以百分等級（PR）呈現。', options: { color: WHITE, breakLine: true } },
  { text: '用途：', options: { bold: true, color: 'E7B9C6' } },
  { text: '這張成績繁星推薦、個人申請、分發入學三個管道都能用。', options: { color: WHITE } },
], {
  x: M, y: 4.3, w: W - 2 * M, h: 1.5, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 15, lineSpacing: 28,
});
s.addText('採計術科的 6 個系：臺師大 25%、北市大 20%、國體大運保 20%、高師大 10%、彰師大運動健康組 10%、東華 10%', {
  x: M, y: 6.1, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 13, color: 'C9A6B0',
});
s.addNotes('體能是體育班的強項。這一場一定要報，10 月底就要交名單。');

/* ── 8-9 校系明細 ── */
function deptSlide(tier, note) {
  const sl = pres.addSlide();
  titleOf(sl, `${tier.title}`, `${tier.n} 個系、${tier.quota} 個名額　｜　「一階通過分數」為 115 學年度實際篩選結果`);
  const cols = [3.5, 0.85, 1.5, 2.8, 1.15, 2.1];
  const head = ['學校・學系', '名額', '檢定', '115 一階通過分數', '學測佔比', '第二階段'];
  let x0 = M;
  const rows = tier.rows.slice(0, 10);
  sl.addShape(pres.ShapeType.rect, { x: M, y: 1.75, w: W - 2 * M, h: 0.42, fill: { color: 'F0DDE3' }, line: { color: 'F0DDE3' } });
  head.forEach((hh, i) => {
    sl.addText(hh, { x: x0 + 0.1, y: 1.79, w: cols[i] - 0.2, h: 0.34, isTextBox: true, margin: 0,
      fontFace: HEAD, fontSize: 12, bold: true, color: DEEP,
      align: (i === 1 || i === 4) ? 'center' : 'left' });
    x0 += cols[i];
  });
  rows.forEach((r, i) => {
    const y = 2.2 + i * 0.44;
    if (i % 2 === 0) {
      sl.addShape(pres.ShapeType.rect, { x: M, y, w: W - 2 * M, h: 0.44, fill: { color: 'FBF6F3' }, line: { color: 'FBF6F3' } });
    }
    const vals = [r.name, String(r.quota ?? '—'), r.gate, r.stage1, r.gsat + '%', r.items];
    let x = M;
    vals.forEach((v, j) => {
      sl.addText(v, { x: x + 0.1, y: y + 0.04, w: cols[j] - 0.2, h: 0.36, isTextBox: true, margin: 0,
        fontFace: BODY, fontSize: 10.5, bold: j === 0, color: j === 2 || j === 5 ? MUT : INK,
        align: (j === 1 || j === 4) ? 'center' : 'left', valign: 'middle' });
      x += cols[j];
    });
  });
  if (tier.rows.length > 10) {
    sl.addText(`另有 ${tier.rows.length - 10} 個系，完整名單見書面資料與升學地圖網頁。`, {
      x: M, y: 2.2 + rows.length * 0.44 + 0.15, w: W - 2 * M, h: 0.4, isTextBox: true, margin: 0,
      fontFace: BODY, fontSize: 12, color: MUT,
    });
  }
  sl.addNotes(note);
}
deptSlide(tiers[0], '這一群是門檻最低的。注意最後一欄，第二階段大多佔六成到十成，備審和面試才是決勝點。');
deptSlide(tiers[1], '中段這一群多半只卡國文或社會一科，門檻在 7 級分上下。');
deptSlide(tiers[2], '這一群名額最多，但要每科 8 到 9 級分。');

/* ── 10 二階才是決勝點 ── */
s = pres.addSlide();
titleOf(s, '學測只是入場券，決勝點在第二階段', '門檻最低那 13 個系的第二階段佔比');
const pcts = [
  ['大葉 運動健康管理', 100], ['實踐 觀光旅運暨運動觀光', 100],
  ['臺體大 體育學系運動專長組', 90], ['臺中教大 體育學系運動專長組', 90],
  ['臺東 體育學系', 80], ['東華 體育與運動科學系', 70],
  ['臺北市立 體育學系', 60], ['臺南 體育學系', 60],
];
pcts.forEach((r, i) => {
  const y = 1.95 + i * 0.6;
  s.addText(r[0], { x: M, y, w: 4.2, h: 0.42, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 14, color: INK, valign: 'middle' });
  s.addShape(pres.ShapeType.rect, { x: M + 4.3, y: y + 0.1, w: 6.4, h: 0.24, fill: { color: 'EDE4E0' }, line: { color: 'EDE4E0' } });
  s.addShape(pres.ShapeType.rect, { x: M + 4.3, y: y + 0.1, w: 6.4 * r[1] / 100, h: 0.24, fill: { color: BERRY }, line: { color: BERRY } });
  s.addText(r[1] + '%', { x: M + 10.85, y, w: 0.9, h: 0.42, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 15, bold: true, color: BERRY, valign: 'middle' });
});
s.addText('第二階段＝審查資料 + 面試。這是現在就能開始練的，不必等學測考完。', {
  x: M, y: 6.75, w: W - 2 * M, h: 0.5, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 17, bold: true, color: DEEP,
});
s.addNotes('這一頁是要講給家長聽的：孩子學測考差不代表沒機會，但備審和面試沒準備就真的沒機會。');

/* ── 11 時程 ── */
s = pres.addSlide();
titleOf(s, '這一年的關卡順序');
const tl = [
  ['10 月底–11 月初', '大學術科考試報名', '由學校集體辦理，錯過沒有補救', BERRY],
  ['1 月', '學測 + 大學術科考試', '術科在國立體育大學，考五項體能', BERRY],
  ['3 月', '個人申請報名、一階放榜', '最多 6 個志願', TEAL],
  ['5 月', '第二階段：審查資料與面試', '決勝點', TEAL],
  ['6 月', '登記就讀志願序、統一分發', '沒登記等於放棄全部錄取校系', GOLD],
];
tl.forEach((r, i) => {
  const y = 1.9 + i * 1.02;
  s.addShape(pres.ShapeType.ellipse, { x: M, y: y + 0.14, w: 0.42, h: 0.42, fill: { color: r[3] }, line: { color: r[3] } });
  s.addText(String(i + 1), { x: M, y: y + 0.14, w: 0.42, h: 0.42, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 15, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  s.addText(r[0], { x: M + 0.7, y: y + 0.1, w: 2.5, h: 0.5, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 17, bold: true, color: r[3], valign: 'middle' });
  s.addText(r[1], { x: M + 3.3, y: y + 0.1, w: 5.0, h: 0.5, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 16, color: INK, valign: 'middle' });
  s.addText(r[2], { x: M + 8.4, y: y + 0.1, w: 3.4, h: 0.5, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 13, color: MUT, valign: 'middle' });
});
s.addNotes('第一關就在下個月。術科報名是體育班最不該漏掉的一件事。');

/* ── 12 結尾三句話 ── */
s = pres.addSlide();
s.background = { color: DEEP };
s.addText('給學生的三句話', {
  x: M, y: 0.9, w: W - 2 * M, h: 0.8, isTextBox: true, margin: 0,
  fontFace: HEAD, fontSize: 36, bold: true, color: WHITE,
});
const three = [
  ['國文和英文是主力', '每科多 1 級分，可填的校系會多十幾個。這是差距最有感的地方。'],
  ['備審和面試現在就能練', '門檻最低那一群，第二階段佔總成績六成到十成。'],
  ['術科考試不看比賽成績', '只看你當天的體能。這是唯一完全由自己決定的分數。'],
];
three.forEach((r, i) => {
  const y = 2.1 + i * 1.45;
  s.addShape(pres.ShapeType.ellipse, { x: M, y: y + 0.08, w: 0.62, h: 0.62, fill: { color: 'E7B9C6' }, line: { color: 'E7B9C6' } });
  s.addText(String(i + 1), { x: M, y: y + 0.08, w: 0.62, h: 0.62, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 22, bold: true, color: DEEP, align: 'center', valign: 'middle' });
  s.addText(r[0], { x: M + 0.95, y, w: 11.0, h: 0.5, isTextBox: true, margin: 0,
    fontFace: HEAD, fontSize: 24, bold: true, color: WHITE });
  s.addText(r[1], { x: M + 0.95, y: y + 0.55, w: 11.0, h: 0.6, isTextBox: true, margin: 0,
    fontFace: BODY, fontSize: 16, color: 'D9BBC4', lineSpacing: 24 });
});
s.addText('完整名單與逐系條件：path-map 升學地圖　→　體育班升學地圖', {
  x: M, y: 6.6, w: W - 2 * M, h: 0.4, isTextBox: true, margin: 0,
  fontFace: BODY, fontSize: 14, color: 'C9A6B0',
});
s.addNotes('收尾。把書面資料發下去，讓家長帶回家看完整名單。');

const out = path.join(BUILD, '..', '體育班升學說明_115.pptx');
pres.writeFile({ fileName: out }).then(() => console.log('wrote', out));
