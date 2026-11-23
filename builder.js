(function (root) {
'use strict';
const A = root.APPDATA;
const X = () => root.docx;
const GREEN = '286B35', BLUE = '5B9BD5', GRAY = 'D9D9D9';
const DAYS = ['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
const MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];

/* ---------- helpers ---------- */
function t12(t) { if (!t) return ''; const [h0, m] = t.split(':').map(Number); const ap = h0 >= 12 ? 'م' : 'ص'; const h = h0 % 12 || 12; return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0') + ' ' + ap; }
function dmy(s) { const [y, m, d] = s.split('-'); return d + '-' + m + '-' + y; }
function dayName(s) { const [y, m, d] = s.split('-').map(Number); return DAYS[new Date(y, m - 1, d).getDay()]; }
function schoolFull(n) { n = (n || '').trim(); return /^المدرسة/.test(n) ? n : 'المدرسة المصرية اليابانية ' + n.split(' - ')[0]; }

function R(text, o) {
  o = o || {}; const x = X();
  return new x.TextRun({ text: String(text == null ? '' : text), font: 'Arial', size: o.size || 22, bold: o.bold !== false, color: o.color,
    underline: o.underline ? { type: x.UnderlineType.SINGLE } : undefined, rightToLeft: o.ltr ? false : true });
}
function P(text, o) {
  o = o || {}; const x = X();
  return new x.Paragraph({ bidirectional: true, alignment: o.align || x.AlignmentType.START,
    spacing: { before: o.before == null ? 30 : o.before, after: o.after == null ? 30 : o.after }, keepNext: o.keepNext,
    children: [R(text, o)] });
}
function PC(text, o) { o = Object.assign({}, o, { align: X().AlignmentType.CENTER }); return P(text, o); }
function numbered(n, text, o) {
  const lines = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  return lines.map((l, i) => P(i === 0 ? n + '- ' + l : l, o));
}
const NONE = () => ({ style: X().BorderStyle.NONE, size: 0, color: 'FFFFFF' });
function C(children, w, o) {
  o = o || {}; const x = X();
  return new x.TableCell({
    width: { size: w, type: x.WidthType.DXA }, columnSpan: o.span, verticalAlign: o.va || x.VerticalAlign.CENTER,
    shading: o.fill ? { type: x.ShadingType.CLEAR, fill: o.fill, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: o.ml == null ? 100 : o.ml, right: o.ml == null ? 100 : o.ml },
    borders: o.noBorder ? { top: NONE(), bottom: NONE(), left: NONE(), right: NONE() } : undefined,
    children: children.length ? children : [new x.Paragraph({ children: [] })]
  });
}
function ROW(cells, o) {
  o = o || {}; const x = X();
  return new x.TableRow({ children: cells, tableHeader: o.header, cantSplit: o.cantSplit,
    height: o.h ? { value: o.h, rule: x.HeightRule.ATLEAST } : undefined });
}
function TBL(rows, cols, o) {
  o = o || {}; const x = X(); const w = cols.reduce((a, b) => a + b, 0);
  return new x.Table({ visuallyRightToLeft: true, width: { size: w, type: x.WidthType.DXA }, columnWidths: cols, rows,
    borders: o.noBorder ? { top: NONE(), bottom: NONE(), left: NONE(), right: NONE(), insideHorizontal: NONE(), insideVertical: NONE() } : undefined });
}
const SP = (n) => new (X().Paragraph)({ spacing: { before: 0, after: n || 80 }, children: [] });

/* ---------- header / footer ---------- */
function header(assets, widthPx) {
  const x = X(); const h = Math.round(widthPx * 230 / 1800);
  return new x.Header({ children: [new x.Paragraph({ alignment: x.AlignmentType.CENTER, spacing: { before: 0, after: 0 },
    children: [new x.ImageRun({ type: 'png', data: assets.banner, transformation: { width: widthPx, height: h } })] })] });
}
function footer(assets) {
  const x = X(), cols = [4700, 2500, 2400, 866];
  const line = { style: x.BorderStyle.SINGLE, size: 18, color: GREEN };
  const txt = (t, ltr, al) => new x.Paragraph({ bidirectional: !ltr, alignment: al || x.AlignmentType.CENTER, spacing: { before: 40, after: 0 },
    children: [new x.TextRun({ text: ltr ? '\u202A' + t + '\u202C' : t, font: 'Arial', size: 15, bold: true, color: GREEN, rightToLeft: !ltr })] });
  const cell = (ch, w) => new x.TableCell({ width: { size: w, type: x.WidthType.DXA }, verticalAlign: x.VerticalAlign.CENTER,
    margins: { top: 20, bottom: 0, left: 20, right: 20 }, borders: { top: NONE(), bottom: NONE(), left: NONE(), right: NONE() }, children: ch });
  const logo = new x.Paragraph({ alignment: x.AlignmentType.CENTER, spacing: { before: 20, after: 0 },
    children: [new x.ImageRun({ type: 'png', data: assets.footer, transformation: { width: 44, height: 21 } })] });
  const t = new x.Table({ visuallyRightToLeft: true, width: { size: 10466, type: x.WidthType.DXA }, columnWidths: cols,
    borders: { top: line, bottom: NONE(), left: NONE(), right: NONE(), insideHorizontal: NONE(), insideVertical: NONE() },
    rows: [new x.TableRow({ children: [
      cell([txt('7 شارع تاج الدين السبكي، أرض الجولف، مدينة نصر، القاهرة')], cols[0]),
      cell([txt('+02 (2) 24141192 / 24180313', true)], cols[1]),
      cell([txt('info@emerald.education', true)], cols[2]),
      cell([logo], cols[3])] })] });
  return new x.Footer({ children: [t] });
}

/* ---------- daily report ---------- */
function attTable(att) {
  const labels = A.ATT, cw = 420, cols = labels.map(() => cw), x = X();
  const r1 = ROW(labels.map(l => C([PC(l, { size: 13, before: 0, after: 0 })], cw, { ml: 20 })));
  const r2 = ROW(labels.map((l, i) => C([PC(att && att[i] ? att[i] : '0', { size: 20, before: 20, after: 20 })], cw, { ml: 20 })));
  return TBL([r1, r2], cols);
}
function mainTable(d) {
  const x = X(), cols = [1300, 4900, 4266];
  const day = d.date ? dayName(d.date) : '', dt = d.date ? '\u202A' + dmy(d.date) + '\u202C' : '';
  const head = ROW([
    C([PC('البند', { size: 22 })], cols[0], { fill: GRAY }),
    C([PC('اليوم ' + day + ' ' + dt, { size: 22 }), PC(t12(d.from) + '  الى  ' + t12(d.to), { size: 22 })], cols[1], { fill: GRAY }),
    C([PC('الردود والملاحظات', { size: 22 })], cols[2], { fill: GRAY })
  ], { header: true, cantSplit: true });
  const rows = [head];
  A.SECTIONS.forEach(sec => {
    const items = [], resp = [];
    sec.items.forEach((it, i) => {
      items.push(...numbered(i + 1, it, { size: 21 }));
      if (sec.id === 'hr' && i === 0) { items.push(attTable(d.att)); items.push(SP(60)); }
      const r = (d.f || {})[sec.id + '_' + i];
      if (r && r.trim()) resp.push(...numbered(i + 1, r, { size: 21 }));
    });
    rows.push(ROW([C([PC(sec.title, { size: 22 })], cols[0]), C(items, cols[1], { va: x.VerticalAlign.TOP }),
      C(resp, cols[2], { va: x.VerticalAlign.TOP })]));
  });
  const ev = String(d.events || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean).map(s => P(s, { size: 21 }));
  rows.push(ROW([C([PC('أحداث وما يستجد خلال الزيارة', { size: 20 })], cols[0]), C(ev, cols[1] + cols[2], { span: 2, va: x.VerticalAlign.TOP })], { h: 1300 }));
  return TBL(rows, cols);
}
function infoBlock(d) {
  const w = [6000, 4466];
  return [
    P('*اسم المدرسة:- ' + schoolFull(d.school), { size: 22, before: 120 }),
    TBL([
      ROW([C([P('*اسم الموظف :- ' + (d.emp || ''), { size: 22 })], w[0], { noBorder: true }), C([P('*الوظيفة:- ' + (d.job || ''), { size: 22 })], w[1], { noBorder: true })]),
      ROW([C([P('اليوم :- ' + (d.date ? dayName(d.date) : ''), { size: 22 })], w[0], { noBorder: true }), C([P('*التاريخ :- ' + (d.date ? '\u202A' + d.date.replace(/-/g, '/') + '\u202C' : ''), { size: 22 })], w[1], { noBorder: true })])
    ], w, { noBorder: true }), SP(100)
  ];
}
function album(photos) {
  if (!photos || !photos.length) return [];
  const x = X(), cols = [2200, 8266];
  const rows = [ROW([C([PC('ألبوم الصور من الزيارة', { size: 28, color: 'FF0000', bold: false, keepNext: true })], cols[0] + cols[1], { span: 2, fill: BLUE })], { cantSplit: true })];
  photos.forEach((p, i) => {
    const s = Math.min(500 / p.w, 340 / p.h, 1), w = Math.round(p.w * s), h = Math.round(p.h * s);
    rows.push(ROW([
      C([P((i + 1) + ' – ' + (p.caption && p.caption.trim() ? p.caption.trim() : 'متابعه'), { size: 22, underline: true })], cols[0], { fill: BLUE, va: x.VerticalAlign.BOTTOM }),
      C([new x.Paragraph({ alignment: x.AlignmentType.CENTER, spacing: { before: 60, after: 60 },
        children: [new x.ImageRun({ type: 'jpg', data: p.data, transformation: { width: w, height: h } })] })], cols[1], { fill: BLUE })
    ], { cantSplit: true }));
  });
  return [SP(160), TBL(rows, cols)];
}
function signatures(d, cfg) {
  const w = [5233, 5233];
  return [SP(300), TBL([
    ROW([C([P('مقدم التقرير :- ' + (d.emp || ''), { size: 22 })], w[0], { noBorder: true }), C([P('يعتمد :- ' + (cfg.approver || ''), { size: 22 })], w[1], { noBorder: true })]),
    ROW([C([P('مسئول قطاع :- ' + (cfg.sector || ''), { size: 22 })], w[0], { noBorder: true }), C([P('مدير ادارة المنسق', { size: 22 })], w[1], { noBorder: true })]),
    ROW([C([P('التوقيع/', { size: 22, before: 160 })], w[0], { noBorder: true }), C([P('التوقيع/', { size: 22, before: 160 })], w[1], { noBorder: true })])
  ], w, { noBorder: true })];
}
function daily(d, photos, assets, cfg) {
  const x = X(); cfg = cfg || {};
  const children = [
    new x.Paragraph({ alignment: x.AlignmentType.CENTER, spacing: { before: 120, after: 60 }, children: [R('التقرير اليومي', { size: 32, underline: true })] }),
    ...infoBlock(d), mainTable(d), ...album(photos), ...signatures(d, cfg)
  ];
  return new x.Document({
    creator: 'Emerald', title: 'التقرير اليومي',
    styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1750, bottom: 1250, left: 720, right: 720, header: 280, footer: 380 } } },
      headers: { default: header(assets, 690) }, footers: { default: footer(assets) }, children }]
  });
}

/* ---------- monthly report ---------- */
function monthly(visits, ym, assets, cfg) {
  const x = X(); cfg = cfg || {};
  const [yy, mm] = ym.split('-').map(Number), cols = [700, 5800, 8898], W = cols[0] + cols[1] + cols[2];
  const v = visits.slice().sort((a, b) => a.date.localeCompare(b.date));
  const schools = [...new Set(v.map(s => s.school))];
  const rows = [ROW([C([PC('م', { size: 22 })], cols[0], { fill: GRAY }), C([PC('بنود المتابعة', { size: 22 })], cols[1], { fill: GRAY }),
    C([PC('ما تم رصده خلال الزيارات (المدرسة – التاريخ : الرد)', { size: 22 })], cols[2], { fill: GRAY })], { header: true, cantSplit: true })];
  A.SECTIONS.forEach(sec => {
    rows.push(ROW([C([PC(sec.title, { size: 24, color: 'FFFFFF' })], W, { span: 3, fill: GREEN })], { cantSplit: true }));
    sec.items.forEach((it, i) => {
      const lines = [];
      v.forEach(s => { const r = (s.f || {})[sec.id + '_' + i]; if (r && r.trim()) lines.push(P(schoolFull(s.school).replace('المدرسة المصرية اليابانية ', '') + ' – \u202A' + dmy(s.date) + '\u202C : ' + r.trim().replace(/\s*\n\s*/g, ' / '), { size: 20 })); });
      rows.push(ROW([C([PC(String(i + 1), { size: 20 })], cols[0]), C([P(it, { size: 20 })], cols[1]), C(lines.length ? lines : [PC('-', { size: 20 })], cols[2], { va: x.VerticalAlign.TOP })]));
    });
  });
  const ev = [];
  v.forEach(s => { if (s.events && s.events.trim()) ev.push(P(schoolFull(s.school).replace('المدرسة المصرية اليابانية ', '') + ' – \u202A' + dmy(s.date) + '\u202C : ' + s.events.trim().replace(/\s*\n\s*/g, ' / '), { size: 20 })); });
  rows.push(ROW([C([PC('أحداث وإجراءات خلال الزيارات', { size: 24, color: 'FFFFFF' })], W, { span: 3, fill: GREEN })], { cantSplit: true }));
  rows.push(ROW([C(ev.length ? ev : [PC('-', { size: 20 })], W, { span: 3, va: x.VerticalAlign.TOP })]));
  const sw = [7699, 7699];
  const children = [
    new x.Paragraph({ alignment: x.AlignmentType.CENTER, spacing: { before: 120, after: 60 }, children: [R('التقرير الشهري لفريق المنسقين عن شهر ' + MONTHS[mm - 1] + ' ' + yy + 'م', { size: 30, underline: true })] }),
    P('عدد الزيارات : ' + v.length + '      المدارس : ' + schools.map(s => schoolFull(s).replace('المدرسة المصرية اليابانية ', '')).join('، '), { size: 21, before: 80, after: 100 }),
    TBL(rows, cols), SP(300),
    TBL([ROW([C([P('مقدم التقرير :- ' + (v[0] && v[0].emp || ''), { size: 22 })], sw[0], { noBorder: true }), C([P('يعتمد :- ' + (cfg.approver || ''), { size: 22 })], sw[1], { noBorder: true })]),
      ROW([C([P('التوقيع/', { size: 22, before: 160 })], sw[0], { noBorder: true }), C([P('التوقيع/', { size: 22, before: 160 })], sw[1], { noBorder: true })])], sw, { noBorder: true })
  ];
  return new x.Document({
    creator: 'Emerald', title: 'التقرير الشهري', styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: x.PageOrientation.LANDSCAPE }, margin: { top: 1750, bottom: 1250, left: 720, right: 720, header: 280, footer: 380 } } },
      headers: { default: header(assets, 1000) }, footers: { default: footer(assets) }, children }]
  });
}

root.Builder = { daily, monthly, MONTHS };
})(typeof window !== 'undefined' ? window : globalThis);
