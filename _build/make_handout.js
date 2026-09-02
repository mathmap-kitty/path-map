/* 產生「體育班升學說明」書面資料（.docx），給學生與家長。
   資料來自 doc_tiers.json（由 sport_apply_core.json 分層而來）。 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageOrientation, LevelFormat, PageBreak,
} = require('docx');

const BUILD = __dirname;
const tiers = JSON.parse(fs.readFileSync(path.join(BUILD, 'doc_tiers.json'), 'utf8'));

const MAROON = '8C2740';
const DARK = '6F1F33';
const INK = '2B2B2B';
const MUT = '6B5F5A';
const ROSE = 'F6E7EC';
const CREAM = 'FBF6F3';

const FONT = '微軟正黑體';

// 橫向 A4：傳直向尺寸，docx-js 自己對調
const PAGE = { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
               margin: { top: 900, right: 900, bottom: 900, left: 900 } };
const CONTENT_W = 16838 - 1800;   // 橫向可用寬度

const t = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 20,
  bold: o.bold, color: o.color || INK, italics: o.italics });

const p = (text, o = {}) => new Paragraph({
  alignment: o.align, spacing: { before: o.before ?? 0, after: o.after ?? 100, line: 300 },
  indent: o.indent, border: o.border,
  children: Array.isArray(text) ? text : [t(text, o)],
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: MAROON, space: 4 } },
  children: [t(text, { size: 30, bold: true, color: MAROON })],
});
const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 100 },
  children: [t(text, { size: 24, bold: true, color: DARK })],
});

const cell = (children, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  margins: { top: 60, bottom: 60, left: 90, right: 90 },
  verticalAlign: 'center',
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

/* ── 校系表 ── */
const COLS = [2900, 800, 1900, 3100, 2700, 900, 2738];
function deptTable(tier) {
  const rows = [headRow(COLS, ['學校・學系', '名額', '檢定標準', '115 一階通過分數', '二階學測採計', '學測佔比', '第二階段項目'])];
  tier.rows.forEach((r, i) => {
    const fill = i % 2 ? CREAM : undefined;
    rows.push(new TableRow({
      children: [
        cell(p([t(r.name, { size: 18, bold: true })], { after: 0 }), { w: COLS[0], fill }),
        cell(p(String(r.quota ?? '—'), { size: 18, align: AlignmentType.CENTER, after: 0 }), { w: COLS[1], fill }),
        cell(p(r.gate, { size: 17, color: MUT, after: 0 }), { w: COLS[2], fill }),
        cell(p(r.stage1, { size: 17, after: 0 }), { w: COLS[3], fill }),
        cell(p(r.cnt, { size: 17, color: MUT, after: 0 }), { w: COLS[4], fill }),
        cell(p(r.gsat + '%', { size: 18, align: AlignmentType.CENTER, after: 0 }), { w: COLS[5], fill }),
        cell(p(r.items, { size: 17, color: MUT, after: 0 }), { w: COLS[6], fill }),
      ],
    }));
  });
  return table(COLS, rows);
}

/* ── 主力科目表 ── */
const SW = [2200, 2600, 2600, 2600];
const subjects = [
  ['國文', '41 系（50%）', '67 系（81%）', '63 系（76%）'],
  ['英文', '28 系（34%）', '54 系（65%）', '56 系（68%）'],
  ['社會', '13 系（15%）', '23 系（28%）', '25 系（30%）'],
  ['數學 B', '10 系（12%）', '11 系（13%）', '12 系（14%）'],
  ['自然', '5 系（6%）', '10 系（12%）', '9 系（10%）'],
  ['數學 A', '8 系（9%）', '2 系（2%）', '2 系（2%）'],
];
const subjTable = table(SW, [
  headRow(SW, ['科目', '設檢定標準', '一階篩選會用到', '二階採計學測']),
  ...subjects.map((r, i) => new TableRow({
    children: r.map((v, j) => cell(
      p(v, { size: 19, bold: j === 0, align: j ? AlignmentType.CENTER : undefined, after: 0,
             color: j === 0 ? (i < 2 ? MAROON : INK) : INK }),
      { w: SW[j], fill: i < 2 ? ROSE : (i % 2 ? CREAM : undefined) })),
  })),
]);

/* ── 專長對照表 ── */
const TW = [1700, 4200, 4300, 3838];
const talents = [
  ['田徑', '臺中教大 運動專長組（男 200 公尺 1 名、女 100 公尺跨欄 1 名）', '限這兩個項目，其餘田徑項目不收', '走大學術科考試或一般組'],
  ['羽球', '臺中教大 運動專長組（男 2 名、女 1 名）', '名額極少', '走大學術科考試或一般組'],
  ['足球', '臺體大 運動專長組（8 名）', '名額最多的一項，但只有這一系', '同時可走大學術科考試'],
  ['滑輪溜冰', '無', '兩系都沒有這個項目的名額', '只能走大學術科考試或一般組'],
];
const talentTable = table(TW, [
  headRow(TW, ['專長', '自辦專長術科收不收', '說明', '沒對上的話走哪裡']),
  ...talents.map((r, i) => new TableRow({
    children: r.map((v, j) => cell(
      p(v, { size: 18, bold: j === 0, after: 0, color: j === 0 ? MAROON : INK }),
      { w: TW[j], fill: i % 2 ? CREAM : undefined })),
  })),
]);

/* ── 術科考試表 ── */
const AW = [4600, 2400, 2400, 4638];
const artRows = [
  ['臺灣師範大學 體育與運動科學系', '25%', 'PR ≥ 75.22', '一階另需 國英社自 ≥ 31/60，檢定四科'],
  ['臺北市立大學 體育學系', '20%', 'PR ≥ 73.08', '一階另需 國英 ≥ 10/30，無檢定'],
  ['體育大學 運動保健學系（體育術科組）', '20%', '二階才採計', '一階 國英自 ≥ 24/45'],
  ['高雄師範大學 體育學系', '10%', 'PR ≥ 14.19', '術科幾乎是形式，卡在國英數B ≥ 21/45'],
  ['彰化師範大學 運動學系（運動健康組）', '10%', '二階才採計', '一階 國英 ≥ 15/30、英文 ≥ 6/15'],
  ['東華大學 體育與運動科學系', '10%', '二階才採計', '一階 國文 ≥ 5、數學B ≥ 2、英文 ≥ 3'],
];
const artTable = table(AW, [
  headRow(AW, ['學校・學系', '術科佔甄選總成績', '115 一階術科門檻', '備註']),
  ...artRows.map((r, i) => new TableRow({
    children: r.map((v, j) => cell(
      p(v, { size: 18, bold: j === 0, after: 0,
             align: (j === 1 || j === 2) ? AlignmentType.CENTER : undefined }),
      { w: AW[j], fill: i % 2 ? CREAM : undefined })),
  })),
]);

const bullet = (text, o = {}) => new Paragraph({
  numbering: { reference: 'dots', level: 0 },
  spacing: { after: 80, line: 300 },
  children: Array.isArray(text) ? text : [t(text, o)],
});

const callout = (lines) => new Table({
  columnWidths: [CONTENT_W],
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE },
    insideVertical: { style: BorderStyle.NONE },
    left: { style: BorderStyle.SINGLE, size: 18, color: MAROON },
  },
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: ROSE, color: 'auto' },
      margins: { top: 140, bottom: 140, left: 200, right: 200 },
      children: lines,
    })],
  })],
});

const tierSection = (tier) => [
  h2(`${tier.tag}　${tier.title}　—　${tier.n} 系、${tier.quota} 個名額`),
  deptTable(tier),
  p('', { after: 160 }),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 20, color: INK } } } },
  numbering: {
    config: [{
      reference: 'dots',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 340, hanging: 200 } } } }],
    }],
  },
  sections: [{
    properties: { page: PAGE },
    children: [
      p([t('115 學年度　體育班升學說明', { size: 40, bold: true, color: MAROON })],
        { align: AlignmentType.CENTER, after: 60 }),
      p([t('以「個人申請」為主要管道　—　給學生與家長', { size: 22, color: MUT })],
        { align: AlignmentType.CENTER, after: 240 }),

      callout([
        p([t('一句話結論：', { bold: true, size: 22, color: DARK }),
           t('沒有比賽成績，運動績優的三條路都關著；主戰場是個人申請，而學科要押的是', { size: 22 }),
           t('國文和英文', { bold: true, size: 22, color: MAROON }),
           t('，不是數學。體能好的人還有一張自己造得出來的牌：', { size: 22 }),
           t('大學術科考試（體育組）', { bold: true, size: 22, color: MAROON }),
           t('，它不看比賽成績。', { size: 22 })], { after: 0 }),
      ]),

      h1('一、為什麼是個人申請'),
      p('體育班學生升大學有四條路，前三條都以運動賽會成績為門票：'),
      bullet([t('運動績優甄審', { bold: true }), t('：須以'), t('國家代表隊', { bold: true }),
              t('身分參加國際賽會並達規定名次（奧運、亞運、世錦賽、亞錦賽等）。')]),
      bullet([t('運動績優甄試', { bold: true }), t('：須有'), t('全國性賽會', { bold: true }),
              t('成績——全國運動會、全民運動會、全國中等學校運動會前八名，或中等學校運動聯賽最優級組前八名。')]),
      bullet([t('各校運動績優單獨招生', { bold: true }), t('：各校自訂，多數仍要求縣市級以上比賽成績或校隊資歷。')]),
      p([t('這三條路都需要「拿得出來的名次」。沒有比賽成績時，能走的是第四條：', { }),
         t('大學個人申請', { bold: true, color: MAROON }), t('。')], { before: 100 }),

      h1('二、應屆生實際能填的範圍'),
      p([t('115 學年度全部 2,206 個校系中，體育班打得到的落點有 100 個。但其中'),
         t('15 個是「青年儲蓄帳戶組」', { bold: true, color: MAROON }),
         t('——那是高中畢業後先去工作或體驗 2～3 年才能用的管道，'),
         t('應屆生填不了', { bold: true }),
         t('，而且每系名額多半只有 1 名；另有 2 個公費生名額另有條件。')]),
      p([t('扣掉之後，'), t('應屆生實際可填的是 83 系、1,495 個名額', { bold: true, color: MAROON }),
         t('。其中 17 系不設任何檢定標準，11 系學測完全不佔甄選總成績（大葉、實踐、開南、佛光）。')]),

      h1('三、主力科目：國文第一，英文第二'),
      p('把 82 個校系（已扣除舞蹈類）攤開，統計每一科在三個關卡出現的次數：'),
      subjTable,
      p('', { after: 120 }),
      callout([
        p([t('怎麼讀這張表：', { bold: true, size: 21, color: DARK })], { after: 60 }),
        p([t('國文', { bold: true }), t(' 是絕對主力——五成的系設國文檢定、八成的系一階會篩國文。'),
           t('英文', { bold: true }), t(' 是第二主力。'),
           t('社會', { bold: true }), t(' 排第三，觀光、休閒、餐旅類的系常用它當篩選科目。')], { after: 60 }),
        p([t('數學 A 只有 2 個系會在一階用到', { bold: true, color: MAROON }),
           t('。數學 B 也只有 13%，集中在幾所公立體育系（臺南、屏東、高師大、臺中教大）。'),
           t('把有限的讀書時間押在國文和英文，效益遠高於數學。', { bold: true })], { after: 0 }),
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      h1('四、四個專長各自能走哪裡'),
      p('有些校系的第二階段是「自辦專長術科測驗」，佔分很高（50–80%），但這類系會限定收哪些運動項目：'),
      talentTable,
      p('', { after: 120 }),
      callout([
        p([t('要注意兩件事：', { bold: true, size: 21, color: DARK })], { after: 60 }),
        p([t('1. 彰化師大運動學系（競技運動與表演藝術組）雖然一階門檻很低，但它'),
           t('要求全運會或全中運前八名等成績', { bold: true, color: MAROON }),
           t('，沒有比賽成績的學生不符合資格，本說明已將它排除。')], { after: 60 }),
        p([t('2. 自辦專長術科的名額少又限項目。'),
           t('真正對所有專長都開放的是「大學術科考試（體育組）」', { bold: true, color: MAROON }),
           t('——它考的是五項體能，不看比賽成績，任何高三生都能報名。')], { after: 0 }),
      ]),

      h1('五、大學術科考試（體育組）：不看比賽成績的那張牌'),
      p([t('每年 1 月在國立體育大學舉行，考五項：'),
         t('60 公尺立姿快跑、20 秒反覆側步、一分鐘屈膝仰臥起坐、立定連續三次跳、1600 公尺跑走',
           { bold: true }),
         t('。五項成績各自轉為 T 分數後加總，男女分開評比，以百分等級（PR）呈現。')]),
      p([t('報名在高三上學期（115 學年度為 10 月 27 日至 11 月 10 日），'),
         t('由學校集體報名，錯過沒有補救管道', { bold: true, color: MAROON }),
         t('。這張成績繁星推薦、個人申請、分發入學三個管道都能用。')]),
      p('', { after: 60 }),
      artTable,

      new Paragraph({ children: [new PageBreak()] }),

      h1('六、校系明細'),
      p([t('以下依「一階通過難度」分三群。'),
         t('「115 一階通過分數」是 115 學年度實際的篩選結果', { bold: true }),
         t('，不是預估——例如「國英 ≥ 10/30」代表國文加英文兩科級分合計要 10 級分以上才通過該關。'),
         t('每一關都要通過才進第二階段。116 學年度數字會變動，僅供抓方向。', { color: MUT })]),
      p('', { after: 100 }),
      ...tiers.flatMap(tierSection),

      h1('七、時程與現在該做的事'),
      bullet([t('10 月 27 日–11 月 10 日　', { bold: true, color: MAROON }),
              t('大學術科考試（體育組）報名，由學校集體辦理。體能是體育班的強項，這一場一定要去。')]),
      bullet([t('1 月　', { bold: true, color: MAROON }),
              t('學科能力測驗；同月術科考試在國立體育大學舉行。')]),
      bullet([t('3 月　', { bold: true, color: MAROON }),
              t('個人申請報名、一階篩選放榜。')]),
      bullet([t('5 月　', { bold: true, color: MAROON }),
              t('第二階段：審查資料與面試。')]),
      p('', { after: 100 }),
      callout([
        p([t('給學生的三句話：', { bold: true, size: 21, color: DARK })], { after: 60 }),
        p([t('1. 國文和英文是主力。'), t('每科多 1 級分，可填的校系會多十幾個', { bold: true }),
           t('——這是差距最有感的地方。')], { after: 60 }),
        p([t('2. 表格最後一欄「第二階段項目」大多是審查資料加面試，佔總成績五成到十成。'),
           t('備審和面試是現在就能開始練的，不必等學測考完。', { bold: true })], { after: 60 }),
        p([t('3. 術科考試不看比賽成績，只看你當天的體能。'),
           t('這是這份名單裡，唯一完全由自己決定的分數。', { bold: true, color: MAROON })], { after: 0 }),
      ]),
      p('', { after: 160 }),
      p([t('資料來源：115 學年度大學個人申請招生簡章校系分則、大學甄選入學委員會 111–115 學年度篩選結果、115 學年度大學術科考試簡章、中等以上學校運動成績優良學生升學輔導辦法。116 學年度簡章公告後數字會調整，選填前請以當年度正式簡章為準。',
         { size: 16, color: MUT })]),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = path.join(BUILD, '..', '體育班升學說明_115.docx');
  fs.writeFileSync(out, buf);
  console.log('wrote', out, (buf.length / 1024).toFixed(0) + ' KB');
});
