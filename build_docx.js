const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, TableOfContents, Header, Footer, PageNumber, convertMillimetersToTwip
} = require('docx');

const D = JSON.parse(fs.readFileSync(__dirname + '/catalogue.json', 'utf8'));
const cat = D.catalogue, sets = D.datasets, fields = D.field_definitions;

// The gaps register is internal review material and is not committed. It lives in
// gaps.json alongside this script; without that file the DOCX builds without the
// gaps section, which is what a clone of the public repository gets.
const GAPS_SRC = __dirname + '/gaps.json';
const gaps = fs.existsSync(GAPS_SRC) ? JSON.parse(fs.readFileSync(GAPS_SRC, 'utf8')).gaps : [];

const W = 9638;               // usable width in DXA (A4 portrait, 2 cm margins)
const INK = '16212B', GREY = '6B7A85', RULE = 'CBD5DB';
// Declared in catalogue.json so this document and the web page cannot drift
// apart; Word wants the hex without the leading '#'.
const DOM = Object.fromEntries(Object.entries(cat.domains).map(([k, v]) => [k, v.replace('#', '')]));
const PROJ = {}; cat.projects.forEach(p => PROJ[p.code] = p);

const eyebrow = [cat.institution, cat.member].filter(Boolean).join(' · ')
  + (cat.unit ? ' — ' + cat.unit : '');

// Every record key this document accounts for: the ones the per-dataset table
// prints, the ones its header prints, and the two granularity ranks, which are a
// web-page device and are deliberately not shown here. Anything else would be
// dropped without a trace, so say so instead.
const COVERED = new Set([
  'id', 'title', 'title_pt', 'domain', 'projects',
  'description', 'key_variables', 'custodian', 'producer', 'contact',
  'spatial_coverage', 'native_resolution', 'analysis_unit', 'temporal_coverage',
  'temporal_resolution', 'update_frequency', 'latency', 'access', 'access_url',
  'formats', 'licence', 'sharing_conditions', 'personal_data', 'linkage_key',
  'limitations', 'holding_status', 'steward', 'record_status',
  'spatial_grain', 'temporal_grain'
]);
const uncovered = [...new Set(sets.flatMap(Object.keys))].filter(k => !COVERED.has(k));
if (uncovered.length) {
  console.warn('warning: record field(s) not shown in the DOCX: ' + uncovered.join(', '));
}

const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const thin = c => ({ style: BorderStyle.SINGLE, size: 4, color: c || RULE });

const p = (text, o = {}) => new Paragraph({
  spacing: { before: o.before ?? 0, after: o.after ?? 100, line: 264 },
  alignment: o.align, keepNext: o.keepNext,
  border: o.rule ? { bottom: thin(INK) } : undefined,
  children: [new TextRun({
    text, bold: o.bold, italics: o.italics, size: o.size ?? 20,
    color: o.color ?? INK, font: o.font ?? 'Calibri',
    allCaps: o.caps, characterSpacing: o.caps ? 20 : undefined
  })]
});

const label = (t, color) => p(t, { caps: true, bold: true, size: 15, color: color || GREY, after: 60, keepNext: true });

const cell = (children, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA },
  shading: o.fill ? { type: ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
  margins: { top: 70, bottom: 70, left: 110, right: 110 },
  borders: {
    top: o.noTop ? NONE : thin(o.bc), bottom: thin(o.bc),
    left: o.open ? NONE : thin(o.bc), right: o.open ? NONE : thin(o.bc)
  },
  children
});

// ---------- front matter ----------
const body = [
  p(eyebrow, { caps: true, size: 15, color: GREY, after: 160 }),
  new Paragraph({
    spacing: { after: 120 }, heading: HeadingLevel.TITLE,
    children: [new TextRun({ text: cat.title, bold: true, size: 44, color: INK, font: 'Calibri' })]
  }),
  p(cat.subtitle, { size: 22, color: GREY, after: 200 }),
  p(`Version ${cat.version}   ·   ${cat.date}   ·   ${sets.length} datasets   ·   ${cat.projects.length} research lines`,
    { font: 'Consolas', size: 17, color: GREY, after: 60 }),
  p('Maintainer: ' + cat.maintainer, { font: 'Consolas', size: 17, color: GREY, after: 320, rule: true }),

  p('Why this catalogue exists', { bold: true, size: 24, before: 240, after: 120 }),
  p(cat.purpose, { size: 21, after: 220 }),

  p('How to use and maintain it', { bold: true, size: 24, before: 160, after: 120 }),
  p('This document is the printable, editable version of the catalogue. The machine-readable register (climate_health_catalogue_register.csv) is the source: one row per dataset, one column per field. The browsable web page is generated from that same file.', { size: 21, after: 120 }),
  p('To add or change a dataset, edit the register, then rebuild the web page and this document. If you edit this document directly instead, copy the change back into the register so the three versions do not drift apart.', { size: 21, after: 120 }),
  p('Records marked "Draft" or "Action needed" are entries where the description still has to be confirmed with the organisation that holds the data. They are kept in the catalogue on purpose: an unconfirmed description that names the right contact is more useful than a blank.', { size: 21, after: 240 }),

  p('Contents', { bold: true, size: 24, before: 160, after: 120 }),
  new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-2' }),
  p('If the list above is blank, click it in Word and press F9 to build it.', { italics: true, size: 17, color: GREY, before: 80 }),
  new Paragraph({ children: [new PageBreak()] })
];

// ---------- research lines ----------
body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 160 },
  children: [new TextRun({ text: 'Research lines', bold: true, size: 28, color: INK, font: 'Calibri' })] }));

cat.projects.forEach(pr => {
  body.push(new Table({
    columnWidths: [W], width: { size: W, type: WidthType.DXA },
    rows: [new TableRow({ children: [cell([
      p(pr.code, { font: 'Consolas', bold: true, size: 17, color: pr.colour, after: 60 }),
      p(pr.name, { bold: true, size: 24, after: 80 }),
      p(pr.question, { italics: true, size: 20, color: GREY, after: 100 }),
      p('Unit of analysis: ' + pr.unit, { size: 19, after: 40 }),
      p('Extent: ' + pr.extent, { size: 19, after: 0 })
    ], { w: W, fill: 'F7F9FA', bc: RULE })] })]
  }));
  body.push(p('', { after: 120 }));
});

// ---------- summary register ----------
body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 100 },
  children: [new TextRun({ text: 'Register at a glance', bold: true, size: 28, color: INK, font: 'Calibri' })] }));
body.push(p('Every dataset, with the three questions asked most often: how fine is it, how often does it change, and can we have it?', { size: 20, color: GREY, after: 160 }));

const SW = [980, 2560, 2000, 1480, 1250, 1368];
const clip = (t, n) => t.length <= n ? t : t.slice(0, t.lastIndexOf(' ', n)) + '…';
const hdr = ['ID', 'Dataset', 'Held by', 'Spatial', 'Temporal', 'Access'];
const sumRows = [new TableRow({ tableHeader: true, children: hdr.map((h, i) =>
  cell([p(h, { caps: true, bold: true, size: 15, color: 'FFFFFF', after: 0 })], { w: SW[i], fill: INK, bc: INK })) })];

sets.forEach((d, i) => {
  const fill = i % 2 ? 'F4F7F8' : undefined;
  const vals = [d.id, d.title, clip(d.custodian.split(';')[0].split(',')[0], 40),
    clip(d.native_resolution, 42), clip(d.temporal_resolution, 40), d.access];
  sumRows.push(new TableRow({ children: vals.map((v, j) =>
    cell([p(v, { size: 16, after: 0, font: j === 0 ? 'Consolas' : 'Calibri', bold: j === 0, color: j === 0 ? DOM[d.domain] : INK })],
      { w: SW[j], fill, bc: RULE })) }));
});
body.push(new Table({ columnWidths: SW, width: { size: W, type: WidthType.DXA }, rows: sumRows }));
body.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- full records ----------
body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 100 },
  children: [new TextRun({ text: 'Dataset records', bold: true, size: 28, color: INK, font: 'Calibri' })] }));
body.push(p('One record per dataset. Fields are defined at the end of the document.', { size: 20, color: GREY, after: 200 }));

const order = ['Health', 'Climate', 'Environment', 'Reference', 'Derived'];
const KW = 2300, VW = W - KW;

order.forEach(domain => {
  const group = sets.filter(d => d.domain === domain);
  if (!group.length) return;
  body.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 260, after: 140 },
    children: [new TextRun({ text: domain + ' data', bold: true, size: 24, color: DOM[domain], font: 'Calibri' })] }));

  group.forEach(d => {
    const rows = [];
    // title band
    rows.push(new TableRow({ children: [new TableCell({
      columnSpan: 2, width: { size: W, type: WidthType.DXA },
      margins: { top: 90, bottom: 90, left: 110, right: 110 },
      shading: { type: ShadingType.CLEAR, fill: DOM[domain], color: 'auto' },
      borders: { top: thin(DOM[domain]), bottom: thin(DOM[domain]), left: thin(DOM[domain]), right: thin(DOM[domain]) },
      children: [
        p(d.id + '   ·   ' + d.projects.join(', '), { font: 'Consolas', size: 15, color: 'FFFFFF', after: 50, caps: true }),
        p(d.title, { bold: true, size: 24, color: 'FFFFFF', after: 30 }),
        p(d.title_pt, { italics: true, size: 18, color: 'FFFFFF', after: 0 })
      ]
    })] }));

    const F = [
      ['What it contains', d.description],
      ['Key variables', d.key_variables],
      ['Who holds it', d.custodian],
      ['Producer', d.producer],
      ['Contact', d.contact],
      ['Spatial coverage', d.spatial_coverage],
      ['Native spatial resolution', d.native_resolution],
      ['Analysis unit', d.analysis_unit],
      ['Temporal coverage', d.temporal_coverage],
      ['Temporal resolution', d.temporal_resolution],
      ['Update frequency', d.update_frequency],
      ['Latency', d.latency],
      ['Access', d.access],
      ['Where to get it', d.access_url],
      ['Formats', d.formats],
      ['Licence', d.licence],
      ['Conditions for sharing', d.sharing_conditions],
      ['Personal data', d.personal_data],
      ['Linkage key', d.linkage_key],
      ['Known limitations', d.limitations],
      ['Holding status', d.holding_status],
      ['Steward', d.steward],
      ['Record status', d.record_status]
    ];
    F.forEach(([k, v]) => {
      const flag = /^(Action needed|Draft)/.test(v) && k === 'Record status';
      rows.push(new TableRow({ children: [
        cell([p(k, { caps: true, bold: true, size: 15, color: GREY, after: 0 })], { w: KW, fill: 'F7F9FA', bc: RULE }),
        cell([p(v, { size: 19, after: 0, color: flag ? 'C00000' : INK, bold: flag })], { w: VW, bc: RULE })
      ] }));
    });
    rows.forEach(r => { r.cantSplit = true; });
    body.push(new Table({ columnWidths: [KW, VW], width: { size: W, type: WidthType.DXA }, rows }));
    body.push(new Paragraph({ spacing: { after: 0 }, children: [new PageBreak()] }));
  });
});

// ---------- gaps ----------
if (gaps.length) {
  body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 100 },
    children: [new TextRun({ text: 'Open gaps and next steps', bold: true, size: 28, color: INK, font: 'Calibri' })] }));
  body.push(p('What the catalogue makes visible. Each item names the missing description or dataset, who owns it and when it closes. ' + gaps.filter(g => g.status === 'Resolved').length + ' of ' + gaps.length + ' are now resolved.', { size: 20, color: GREY, after: 180 }));

  const GW = [620, W - 620];
  const TONE = { 'Resolved': '2E6B3E', 'In progress': '9A6B15', 'Planned': '6B7A85' };
  gaps.forEach((g, i) => {
    body.push(new Table({ columnWidths: GW, width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
      cell([p(String(i + 1).padStart(2, '0'), { font: 'Consolas', bold: true, size: 20, color: TONE[g.status], after: 0 })],
        { w: GW[0], fill: g.status === 'Resolved' ? 'F0F6F1' : 'FBF7EF', bc: RULE }),
      cell([
        p(g.gap, { bold: true, size: 21, after: 30 }),
        p(g.status.toUpperCase() + '   \u00b7   ' + g.owner + '   \u00b7   ' + g.when,
          { font: 'Consolas', size: 15, color: TONE[g.status], after: 70 }),
        p(g.detail, { size: 19, color: GREY, after: 80 }),
        label(g.status === 'Resolved' ? 'What was done' : 'Next step', TONE[g.status]),
        p(g.action, { size: 19, after: 0 })
      ], { w: GW[1], bc: RULE })
    ] })] }));
    body.push(p('', { after: 100 }));
  });
}

// ---------- field definitions ----------
body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 100 },
  children: [new TextRun({ text: 'What each field means', bold: true, size: 28, color: INK, font: 'Calibri' })] }));
body.push(p('The agreed description standard. Any organisation joining the catalogue fills in these fields for its own datasets — that is all the coordination it requires.', { size: 20, color: GREY, after: 160 }));

const DW = [2500, W - 2500];
const defRows = [new TableRow({ tableHeader: true, children: ['Field', 'What goes in it'].map((h, i) =>
  cell([p(h, { caps: true, bold: true, size: 15, color: 'FFFFFF', after: 0 })], { w: DW[i], fill: INK, bc: INK })) })];
fields.forEach(([k, v]) => defRows.push(new TableRow({ children: [
  cell([p(k, { bold: true, size: 18, after: 0 })], { w: DW[0], fill: 'F7F9FA', bc: RULE }),
  cell([p(v, { size: 18, after: 0 })], { w: DW[1], bc: RULE })
] })));
body.push(new Table({ columnWidths: DW, width: { size: W, type: WidthType.DXA }, rows: defRows }));

// ---------- change log ----------
body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 100 },
  children: [new TextRun({ text: 'Change log', bold: true, size: 28, color: INK, font: 'Calibri' })] }));
body.push(p('Record every revision here, so partners can see what changed since the version they read.', { size: 20, color: GREY, after: 160 }));

const CW = [1300, 1300, W - 4100, 1500];
const logRows = [new TableRow({ tableHeader: true, children: ['Date', 'Version', 'What changed', 'By'].map((h, i) =>
  cell([p(h, { caps: true, bold: true, size: 15, color: 'FFFFFF', after: 0 })], { w: CW[i], fill: INK, bc: INK })) })];
logRows.push(new TableRow({ children: [
  cell([p('2026-08-19', { size: 18, after: 0 })], { w: CW[0], bc: RULE }),
  cell([p('0.1', { size: 18, after: 0 })], { w: CW[1], bc: RULE }),
  cell([p('First draft for review. 24 datasets across three research lines; seven open gaps recorded.', { size: 18, after: 0 })], { w: CW[2], bc: RULE }),
  cell([p('MCCM', { size: 18, after: 0 })], { w: CW[3], bc: RULE })
] }));
logRows.push(new TableRow({ children: [
  cell([p('2026-08-19', { size: 18, after: 0 })], { w: CW[0], bc: RULE }),
  cell([p('0.2', { size: 18, after: 0 })], { w: CW[1], bc: RULE }),
  cell([p('Seven gaps reviewed: three resolved (Para humidity, HIBR-10 provenance, and one internal data-access item), one in progress (inundation layer, with CLI-006 and ENV-008 added as candidate sources), three planned with named owners and dates. ENV-007 re-described in full.', { size: 18, after: 0 })], { w: CW[2], bc: RULE }),
  cell([p('MCCM', { size: 18, after: 0 })], { w: CW[3], bc: RULE })
] }));
for (let i = 0; i < 5; i++) logRows.push(new TableRow({ children: CW.map(w =>
  cell([p('', { size: 18, after: 0 })], { w, bc: RULE })) }));
body.push(new Table({ columnWidths: CW, width: { size: W, type: WidthType.DXA }, rows: logRows }));

// ---------- assemble ----------
const doc = new Document({
  creator: cat.maintainer, title: cat.title,
  styles: { default: { document: { run: { font: 'Calibri', size: 20, color: INK } } } },
  sections: [{
    properties: { page: { margin: {
      top: convertMillimetersToTwip(20), bottom: convertMillimetersToTwip(20),
      left: convertMillimetersToTwip(20), right: convertMillimetersToTwip(20)
    } } },
    headers: { default: new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { after: 60 },
      children: [new TextRun({ text: cat.title + '  ·  v' + cat.version, size: 15, color: GREY, font: 'Consolas' })]
    })] }) },
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [new TextRun({ children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], size: 15, color: GREY, font: 'Consolas' })]
    })] }) },
    children: body
  }]
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(__dirname + '/climate_health_catalogue.docx', b);
  console.log('docx written', b.length);
});
