/* 產生「運動績優單獨招生整理」書面資料（.docx）。
   資料為 115 學年度（去年）各校填報內容，用來抓 116 的方向。 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageOrientation, LevelFormat, PageBreak,
} = require('docx');

const BUILD = __dirname;
const byTalent = JSON.parse(fs.readFileSync(path.join(BUILD, 'single_by_talent.json'), 'utf8'));

const MAROON = '8C2740', DARK = '6F1F33', INK = '2B2B2B', MUT = '6B5F5A';
const ROSE = 'F6E7EC', CREAM = 'FBF6F3', TEAL = '1F6F78';
const FONT = '微軟正黑體';

const PAGE = { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
               margin: { top: 850, right: 850, bottom: 850, left: 850 } };
const CONTENT_W = 16838 - 1700;

const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 20,
  bold: o.bold, color: o.color || INK });
const p = (text, o = {}) => new Paragraph({
  alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: 300 },
  children: Array.isArray(text) ? text : [t(text, o)],
});
const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 150 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: MAROON, space: 4 } },
  children: [t(text, { size: 30, bold: true, color: MAROON })],
});
const h2 = (text, sub) => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 90 },
  children: [t(text, { size: 24, bold: true, color: DARK }),
             ...(sub ? [t('　' + sub, { size: 16, color: MUT })] : [])],
});
const cell = (children, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  margins: { top: 55, bottom: 55, left: 85, right: 85 }, verticalAlign: 'center',
  children: Array.isArray(children) ? children : [children],
});
const table = (widths, rows) => new Table({
  columnWidths: widths,
  width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: 'D8CCC6' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D8CCC6' },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'E7DCD6' },
    insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'E7DCD6' },
  },
  rows,
});
const headRow = (widths, labels) => new TableRow({
  tableHeader: true,
  children: labels.map((L, i) => cell(p(L, { size: 18, bold: true, color: DARK, after: 0 }),
    { w: widths[i], fill: ROSE })),
});
const callout = (lines, color) => new Table({
  columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE },
    insideVertical: { style: BorderStyle.NONE },
    left: { style: BorderStyle.SINGLE, size: 18, color: color || MAROON },
  },
  rows: [new TableRow({ children: [new TableCell({
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: color === TEAL ? 'E3F0F0' : ROSE, color: 'auto' },
    margins: { top: 130, bottom: 130, left: 190, right: 190 },
    children: lines,
  })] })],
});
const bullet = (text) => new Paragraph({
  numbering: { reference: 'dots', level: 0 }, spacing: { after: 70, line: 300 },
  children: Array.isArray(text) ? text : [t(text)],
});

/* ── 各專長的學校表 ── */
const C = [2750, 800, 700, 1950, 2500, 2200, 4138];
function talentTable(rows) {
  const trs = [headRow(C, ['學校', '名額', '系數', '學科要求', '報名期間', '術科檢定', '有開名額的體育／運動相關系'])];
  rows.forEach((r, i) => {
    const fill = i % 2 ? CREAM : undefined;
    trs.push(new TableRow({ children: [
      cell(p(r.sch, { size: 18, bold: true, after: 0 }), { w: C[0], fill }),
      cell(p(String(r.quota), { size: 18, align: AlignmentType.CENTER, after: 0, bold: true, color: MAROON }), { w: C[1], fill }),
      cell(p(String(r.ndep), { size: 17, align: AlignmentType.CENTER, after: 0 }), { w: C[2], fill }),
      cell(p(r.exam, { size: 17, after: 0, color: r.exam === '採計學測' ? MAROON : INK }), { w: C[3], fill }),
      cell(p(r.reg, { size: 16, after: 0, color: MUT }), { w: C[4], fill }),
      cell(p(r.art || '—', { size: 16, after: 0, color: MUT }), { w: C[5], fill }),
      cell(p(r.sportdep || (r.other + ' 個系全是一般科系'), { size: 16, after: 0, color: MUT }), { w: C[6], fill }),
    ] }));
  });
  return table(C, trs);
}

const talents = ['田徑', '滑輪溜冰', '足球', '羽球'];
const summaryW = [2400, 2400, 2400, 2400, 4838];
const summary = table(summaryW, [
  headRow(summaryW, ['專長', '開名額的學校', '校系數', '名額合計', '名額最多的三所']),
  ...talents.map((name, i) => {
    const rows = byTalent[name];
    const top3 = rows.slice(0, 3).map((r) => `${r.sch.replace('國立', '')} ${r.quota}`).join('、');
    return new TableRow({ children: [
      cell(p(name, { size: 20, bold: true, color: MAROON, after: 0 }), { w: summaryW[0], fill: i % 2 ? CREAM : undefined }),
      cell(p(rows.length + ' 所', { size: 19, align: AlignmentType.CENTER, after: 0 }), { w: summaryW[1], fill: i % 2 ? CREAM : undefined }),
      cell(p(rows.reduce((a, b) => a + b.ndep, 0) + ' 系', { size: 19, align: AlignmentType.CENTER, after: 0 }), { w: summaryW[2], fill: i % 2 ? CREAM : undefined }),
      cell(p(rows.reduce((a, b) => a + b.quota, 0) + ' 名', { size: 19, align: AlignmentType.CENTER, bold: true, color: MAROON, after: 0 }), { w: summaryW[3], fill: i % 2 ? CREAM : undefined }),
      cell(p(top3, { size: 17, after: 0, color: MUT }), { w: summaryW[4], fill: i % 2 ? CREAM : undefined }),
    ] });
  }),
]);

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20, color: INK } } } },
  numbering: { config: [{ reference: 'dots', levels: [{ level: 0, format: LevelFormat.BULLET,
    text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 340, hanging: 200 } } } }] }] },
  sections: [{
    properties: { page: PAGE },
    children: [
      p([t('運動績優單獨招生　整理', { size: 40, bold: true, color: MAROON })],
        { align: AlignmentType.CENTER, after: 60 }),
      p([t('115 學年度（去年）資料　—　田徑・滑輪溜冰・足球・羽球', { size: 22, color: MUT })],
        { align: AlignmentType.CENTER, after: 220 }),

      callout([
        p([t('先讀這一段：', { bold: true, size: 22, color: DARK })], { after: 70 }),
        p([t('單招的「報考資格」由各校自己訂，'),
           t('全國單招查詢系統上看不到', { bold: true, color: MAROON }),
           t('——多數學校要求縣市級以上比賽成績或校隊資歷，也有學校（例如成功大學）把「經主管機關核定的高中體育班畢業生」列為可報考資格。'),
           t('所以這份表只能告訴你「哪些學校有開這個項目的名額、什麼時候報名」，不能告訴你「你的學生有沒有資格」。', { bold: true })], { after: 70 }),
        p([t('資格要逐校翻簡章確認。表格最右邊沒有列出的部分，以及每一校的完整條件，都在各校試務網頁上（見最後一節）。', {})], { after: 0 }),
      ]),

      h1('一、單招和個人申請差在哪'),
      bullet([t('管道性質：', { bold: true }), t('單招是各大學經教育部核定自行辦理的招生，不經甄選會、不經分發會，報名、考試、放榜全部由學校自己來。')]),
      bullet([t('時間最早：', { bold: true }), t('115 學年度多數學校在 '), t('114 年 12 月至 115 年 3 月', { bold: true, color: MAROON }), t(' 報名，術科檢定集中在 3 月中下旬——比個人申請的二階早了兩個月。')]),
      bullet([t('學測不一定要：', { bold: true }), t('有些學校採計學測成績（國體大、臺東大、輔大部分系、臺體大），有些完全不採計、只看術科與書審面試（文化、明新、開南、元培等）。這一欄在下面每張表都有標。')]),
      bullet([t('名額比個申多：', { bold: true }), t('115 學年度全國單招共 896 個志願、3,263 個名額，是四條路裡名額最多的。')]),
      bullet([t('要自己盯：', { bold: true }), t('沒有統一的報名系統，每一所學校各自公告、各自收件，'), t('錯過單一學校的期程無法補救', { bold: true, color: MAROON }), t('。')]),

      h1('二、四個專長的規模對照'),
      summary,
      p('', { after: 130 }),
      callout([
        p([t('兩個要注意的地方：', { bold: true, size: 21, color: DARK })], { after: 70 }),
        p([t('1. '), t('滑輪溜冰只有 10 所學校開名額', { bold: true, color: MAROON }),
           t('，而且其中 8 所是私立或科大；田徑、足球、羽球則有 33～56 所。專長不同，選擇的寬度差很多。')], { after: 70 }),
        p([t('2. '), t('開名額的不只體育系。', { bold: true, color: MAROON }),
           t('輔仁大學一次開 25 個系收運動績優生，包含企業管理 15 名、營養科學 8 名、金融與國際企業 6 名、影像傳播 6 名；元培、開南、明新也都是全校性開放。'),
           t('這代表體育班學生不是只能念體育系。', { bold: true })], { after: 0 }),
      ], TEAL),

      new Paragraph({ children: [new PageBreak()] }),

      ...talents.flatMap((name) => {
        const rows = byTalent[name];
        return [
          h1(`三、${name}`),
          p([t(`${rows.length} 所學校、${rows.reduce((a, b) => a + b.ndep, 0)} 個校系、${rows.reduce((a, b) => a + b.quota, 0)} 個名額。依名額多寡排序。`,
               { color: MUT, size: 18 })]),
          talentTable(rows),
          p('', { after: 160 }),
        ];
      }),

      h1('四、現在該做的事'),
      bullet([t('先確認資格。', { bold: true, color: MAROON }), t('挑出想去的學校，一所一所去翻「運動績優單獨招生簡章」，看它接受哪些成績或資歷。沒有比賽成績的學生，重點看有沒有「體育班畢業」或「校隊一年以上」這類條件。')]),
      bullet([t('把報名日期抄進行事曆。', { bold: true }), t('最早的 12 月中旬就開跑，最晚的到 5 月。表格裡的日期是 115 學年度的，116 會微調但落點差不多。')]),
      bullet([t('學測還是要考。', { bold: true }), t('有相當比例的學校採計學測；就算單招不採計，個人申請那條路也需要。')]),
      bullet([t('術科檢定日要對時間。', { bold: true }), t('多數集中在 3 月中下旬，同一天可能撞校。先排優先順序。')]),
      p('', { after: 130 }),
      callout([
        p([t('116 學年度的資料會自動更新', { bold: true, size: 21, color: DARK })], { after: 70 }),
        p([t('這份是 115 學年度（去年）的資料，用來抓方向。116 學年度各校的簡章從 12 月中旬開始陸續公告，'),
           t('系統設定為 9–10 月每兩週、11 月起每週自動檢核一次', { bold: true }),
           t('，有新簡章或名額異動就更新資料並產出書面統計。')], { after: 0 }),
      ], TEAL),
      p('', { after: 160 }),
      p([t('資料來源：大專校院單獨招生試務資訊查詢系統（各校自行填報，115 學年度共 896 筆）。各項招生考試日期、方式及內容仍以各校正式招生簡章為準；本表與簡章牴觸時，一律以簡章為準。',
           { size: 16, color: MUT })]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(BUILD, '..', '運動績優單獨招生整理_115.docx');
  fs.writeFileSync(out, buf);
  console.log('wrote', out, (buf.length / 1024).toFixed(0) + ' KB');
});
