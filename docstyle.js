// Shared Word styles for The Constrictor docs (navy and gold, Arial). Used by csj.js, second_chance.js and athlete.js.
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Footer, PageNumber, ShadingType, LevelFormat, ImageRun,
  Table, TableRow, TableCell, WidthType, VerticalAlign, TableLayoutType } = require("docx");

module.exports = function style(opt = {}) {
  const NAVY = opt.navy || "081C2B", GOLD = opt.gold || "C8862E", GOLDL = opt.goldLight || "FBEBD3", PALE = "F2F4F6", FONT = "Arial", W = 9720;
  const R = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 20, ...o });
  // [Cody: ...] and other [brackets] are highlighted yellow so the gaps are easy to spot
  const split = (t, o = {}) => String(t).split(/(\[[^\]]+\])/).filter(Boolean).map(s => s.startsWith("[") ? R(s, { ...o, highlight: "yellow" }) : R(s, o));
  const runs = (x) => (Array.isArray(x) ? x : [x]).flatMap(y => typeof y === "string" ? split(y) : (Array.isArray(y) ? y : [y]));
  const P = (x, o = {}) => new Paragraph({ spacing: { after: 90 }, ...o, children: runs(x) });
  const Bd = (t, o = {}) => R(t, { bold: true, ...o });
  let SEC = 0;
  const H1 = (t, sub, brk) => { SEC++; return [new Paragraph({ keepNext: true, pageBreakBefore: !!brk, spacing: { before: 280, after: sub ? 20 : 100 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: GOLD, space: 3 } },
      children: [new TextRun({ text: `${SEC}. ${t}`.toUpperCase(), font: FONT, size: 27, bold: true, color: NAVY, characterSpacing: 10 })] }),
    ...(sub ? [new Paragraph({ keepNext: true, spacing: { after: 110 }, children: [R(sub, { italics: true, color: "5A6672", size: 19 })] })] : [])]; };
  const H2 = (t) => new Paragraph({ keepNext: true, spacing: { before: 160, after: 60 }, children: [R(t, { bold: true, size: 22, color: GOLD })] });
  const bullet = (x) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 45 }, children: runs(x) });
  const num = (x, ref = "n") => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 55 }, children: runs(x) });
  const chk = (x) => new Paragraph({ keepNext: true, spacing: { after: 60 }, indent: { left: 360, hanging: 360 }, children: [new TextRun({ text: "☐  ", font: "Segoe UI Symbol", size: 22 }), ...runs(x)] });
  const line = { style: BorderStyle.SINGLE, size: 4, color: "D7DDE3" };
  const borders = { top: line, bottom: line, left: line, right: line };
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const cellParas = (x) => (x && x._lines)
    ? x.map(l => new Paragraph({ spacing: { after: 30 }, children: runs(l) })) : [new Paragraph({ spacing: { after: 20 }, children: runs(x) })];
  const cell = (x, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: o.borders || borders, verticalAlign: o.v || VerticalAlign.CENTER,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, margins: { top: 60, bottom: 60, left: 110, right: 110 },
    children: cellParas(x) });
  const lines = (...l) => { const a = [...l]; a._lines = true; return a; };
  function table(cols, rows, o = {}) {
    return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: cols, layout: TableLayoutType.FIXED,
      rows: rows.map((r, i) => new TableRow({ cantSplit: !o.split, tableHeader: i === 0,
        children: r.map((c, j) => i === 0 ? cell([Bd(c, { color: "FFFFFF", size: 18 })], cols[j], { fill: NAVY })
          : cell(c === "☐" ? [new TextRun({ text: "☐", font: "Segoe UI Symbol", size: 22 })] : (typeof c === "string" && j === 0 && o.label ? [Bd(c, { size: 19 })] : c), cols[j],
            { fill: o.label && j === 0 ? GOLDL : (i % 2 === 0 ? PALE : undefined), v: o.top ? VerticalAlign.TOP : undefined })) })) });
  }
  const callout = (title, ls, fill = GOLDL, bar = GOLD) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" },
      borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: bar } }, margins: { top: 110, bottom: 110, left: 200, right: 160 },
      children: [new Paragraph({ spacing: { after: 50 }, children: [Bd(title, { size: 21, color: NAVY })] }), ...ls.map(l => P(l, { spacing: { after: 50 } }))] })] })] });
  const sp = (a = 100) => new Paragraph({ spacing: { after: a }, children: [] });
  // A "draft answer" box: the text Cody pastes into the form
  const answer = (label, text) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: false, children: [new TableCell({ width: { size: W, type: WidthType.DXA }, borders,
      margins: { top: 90, bottom: 90, left: 160, right: 160 },
      children: [new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [Bd(label, { size: 18, color: GOLD, characterSpacing: 10 })] }),
        ...(Array.isArray(text) ? text : [text]).map(t => P(t, { spacing: { after: 50 } }))] })] })] });
  function cover(eyebrow, title, subtitle, logos = []) {
    const lw = logos.length ? 3000 : 0;
    return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: logos.length ? [W - lw, lw] : [W], layout: TableLayoutType.FIXED,
      rows: [new TableRow({ children: [
        new TableCell({ width: { size: W - lw, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" }, borders: { top: none, bottom: none, left: none, right: none },
          margins: { top: 300, bottom: 280, left: 340, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [
            new Paragraph({ spacing: { after: 40 }, children: [R(eyebrow, { size: 20, bold: true, color: opt.goldOnNavy || "EAA959", characterSpacing: 40 })] }),
            new Paragraph({ spacing: { after: 60 }, children: [R(title, { font: "Arial Black", size: 46, bold: true, color: "F4F1EA" })] }),
            new Paragraph({ children: [R(subtitle, { size: 20, color: "C9D3DC" })] })] }),
        ...(logos.length ? [new TableCell({ width: { size: lw, type: WidthType.DXA }, shading: { fill: NAVY, type: ShadingType.CLEAR, color: "auto" }, borders: { top: none, bottom: none, left: none, right: none },
          margins: { top: 160, bottom: 160, left: 100, right: 260 }, verticalAlign: VerticalAlign.CENTER, children: [
            new Paragraph({ alignment: AlignmentType.RIGHT, children: logos.flatMap((l, i) => [...(i ? [R("  ")] : []), new ImageRun({ type: "png", data: fs.readFileSync(l), transformation: { width: 84, height: 84 } })]) })] })] : []),
      ] })] });
  }
  function save(out, children, footerText) {
    const doc = new Document({
      styles: { default: { document: { run: { font: FONT, size: 20 } } } },
      numbering: { config: [
        { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } }] },
        ...["n", "n2", "n3", "n4", "n5", "n6"].map(ref => ({ reference: ref, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 300 } } } }] })),
      ] },
      sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 900, bottom: 900, left: 1260, right: 1260 } } },
        footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [R(footerText + "  ·  page ", { size: 16, color: "8A949E" }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: "8A949E" })] })] }) },
        children }],
    });
    return Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log("saved", out); });
  }
  return { R, P, Bd, H1, H2, bullet, num, chk, table, callout, sp, answer, cover, save, lines, W, NAVY, GOLD };
};
