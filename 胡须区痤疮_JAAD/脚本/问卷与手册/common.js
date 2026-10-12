// Shared helpers for the questionnaire and investigator manual (docx-js).
const d = require('docx');
const { Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
        AlignmentType, VerticalAlign, Footer, PageNumber } = d;

const F = { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '宋体', cs: 'Arial' };
const FH = { ascii: 'Arial', hAnsi: 'Arial', eastAsia: '黑体', cs: 'Arial' };
const PAGE = { width: 11906, height: 16838 };
const MARGIN = 1020;
const W = PAGE.width - 2 * MARGIN; // 9866 DXA content width
const GRAY = '595959';

function runs(spec, base = {}) {
  const arr = Array.isArray(spec) ? spec : [spec];
  return arr.map(s => (typeof s === 'string')
    ? new TextRun({ text: s, ...base })
    : new TextRun({ text: s.t, bold: s.b, italics: s.i, color: s.color, size: s.size ?? base.size,
                    font: s.h ? FH : base.font, underline: s.u ? {} : undefined }));
}

function P(spec, o = {}) {
  return new Paragraph({
    children: runs(spec, o.run || {}),
    spacing: { before: o.before ?? 0, after: o.after ?? 60, line: o.line ?? 300 },
    indent: (o.indent || o.hanging) ? { left: o.indent || 0, hanging: o.hanging } : undefined,
    alignment: o.align,
    keepNext: o.keepNext,
    keepLines: true,
    pageBreakBefore: o.pageBreakBefore,
    border: o.border,
    numbering: o.numbering,
  });
}

// Section heading with a bottom rule.
function H(text, o = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, font: FH, size: o.size ?? 26 })],
    spacing: { before: o.before ?? 280, after: 120 },
    keepNext: true,
    pageBreakBefore: o.pageBreakBefore,
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '7F7F7F', space: 2 } },
  });
}
function H2(text, o = {}) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, font: FH, size: o.size ?? 22 })],
    spacing: { before: o.before ?? 200, after: 80 }, keepNext: true,
  });
}

// Question stem: bold number + text (+ optional gray note on the same line).
function Q(num, spec, note) {
  const ch = [new TextRun({ text: `${num}．`, bold: true, font: FH })];
  ch.push(...runs(spec));
  if (note) ch.push(new TextRun({ text: `　${note}`, color: GRAY, size: 19 }));
  return new Paragraph({ children: ch, spacing: { before: 160, after: 60, line: 300 },
                         indent: { left: 420, hanging: 420 }, keepNext: true, keepLines: true });
}
// Option line: □ a　□ b ...
function O(opts, o = {}) {
  const text = opts.map(s => `□ ${s}`).join('　　');
  return new Paragraph({ children: runs([o.prefix ? { t: o.prefix, b: true } : '', text].filter(Boolean)),
                         spacing: { before: 0, after: 60, line: 300 }, indent: { left: o.indent ?? 420 },
                         keepNext: o.keepNext, keepLines: true });
}
// Small gray note line.
function N(text, o = {}) {
  return new Paragraph({ children: [new TextRun({ text, color: GRAY, size: o.size ?? 19 })],
                         spacing: { before: 0, after: 60, line: 280 }, indent: { left: o.indent ?? 420 },
                         keepNext: o.keepNext, keepLines: true });
}

const thin = { style: BorderStyle.SINGLE, size: 4, color: '8C8C8C' };
const borders = { top: thin, bottom: thin, left: thin, right: thin, insideHorizontal: thin, insideVertical: thin };

function cell(content, width, o = {}) {
  const paras = (Array.isArray(content) && content.length && content[0] instanceof Paragraph) ? content
    : [new Paragraph({ children: runs(content, { size: o.size, bold: o.bold }), alignment: o.align ?? AlignmentType.CENTER,
                       spacing: { before: 30, after: 30, line: 260 }, keepNext: o.keepNext, keepLines: true })];
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    columnSpan: o.span, rowSpan: o.rowSpan,
    verticalAlign: VerticalAlign.CENTER,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
  });
}

// Grid table. header: array of header rows; each header cell = string | {t, span, rowSpan}.
// body: array of rows (arrays of strings or Paragraph arrays). First column left-aligned unless o.centerFirst.
function T(header, body, widths, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const size = o.size ?? 19;
  const rows = [];
  for (const hr of header) {
    let col = 0;
    const cells = hr.map(hc => {
      const c = typeof hc === 'string' ? { t: hc } : hc;
      const span = c.span || 1;
      const w = widths.slice(col, col + span).reduce((a, b) => a + b, 0);
      col += span;
      return cell([{ t: c.t, b: true, size }], w, { span: c.span, rowSpan: c.rowSpan, fill: 'EDEDED', size, keepNext: !o.split });
    });
    rows.push(new TableRow({ children: cells, tableHeader: true, cantSplit: true }));
  }
  body.forEach((r, ri) => {
    const keep = !o.split && ri < body.length - 1;
    rows.push(new TableRow({ cantSplit: true, children: r.map((v, i) =>
      cell(typeof v === 'string' ? [{ t: v, size }] : v, widths[i],
           { size, keepNext: keep, align: (i === 0 && !o.centerFirst) || (o.leftCols || []).includes(i) ? AlignmentType.LEFT : AlignmentType.CENTER })) }));
  });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows, borders });
}

// Bordered box (single-cell table) for instructions.
function Box(paras, o = {}) {
  return new Table({
    width: { size: o.width ?? W, type: WidthType.DXA }, columnWidths: [o.width ?? W], borders,
    rows: [new TableRow({ cantSplit: o.cantSplit ?? true, children: [new TableCell({
      children: paras, width: { size: o.width ?? W, type: WidthType.DXA },
      shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
      margins: { top: 100, bottom: 100, left: 160, right: 160 } })] })],
  });
}

function spacer(after = 60) { return new Paragraph({ children: [], spacing: { before: 0, after } }); }

function footer(leftText) {
  return new Footer({ children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: leftText, size: 17, color: GRAY }),
               new TextRun({ children: ['第 ', PageNumber.CURRENT, ' 页 / 共 ', PageNumber.TOTAL_PAGES, ' 页'], size: 17, color: GRAY })],
  })] });
}

const docDefaults = {
  styles: { default: { document: { run: { font: F, size: 21 } } } },
};

const SITES = ['① 前额', '② 左面颊', '③ 右面颊', '④ 鼻部', '⑤ 人中（鼻下至上唇）', '⑥ 下巴', '⑦ 下颌缘'];

module.exports = { d, F, FH, PAGE, MARGIN, W, GRAY, runs, P, H, H2, Q, O, N, T, Box, cell, spacer, footer, docDefaults, SITES };
