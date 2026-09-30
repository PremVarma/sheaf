#!/usr/bin/env node
/**
 * Generates sample files in ./samples for trying the viewer:
 *
 *   npm run samples             sales-report.xlsx, cities.csv, legacy.xls, corrupted.xlsx
 *   npm run samples -- --large  also large-100k.xlsx and large-500k.csv
 *
 * SheetJS Community Edition can't write cell styles or freeze panes, so the
 * .xlsx is post-processed: fonts, fills, borders and a frozen header row are
 * patched into the package XML, the same way Excel stores them.
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';

const outDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'samples');
mkdirSync(outDir, { recursive: true });
const large = process.argv.includes('--large');

function random(seed) {
  return () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed / 2 ** 32;
  };
}

const CITIES = ['Mumbai', 'Delhi', 'Bengaluru', 'Chennai', 'Kolkata', 'Pune', 'Hyderabad', 'Ahmedabad', 'Jaipur', 'Lucknow'];
const PRODUCTS = { Laptop: 68999, Monitor: 15499, Keyboard: 2499, Mouse: 999, Headset: 3499, Webcam: 4299, Dock: 8999, Tablet: 32999 };

/** Adds styles to a SheetJS-written package. `styles` maps sheet index → { A1: styleName }. */
function patchPackage(buffer, { styles, frozen, hideGridlines }) {
  const cfb = XLSX.CFB.read(buffer, { type: 'buffer' });
  const readPart = (path) => Buffer.from(XLSX.CFB.find(cfb, `/${path}`).content).toString('utf8');
  const writePart = (path, text) => {
    const entry = XLSX.CFB.find(cfb, `/${path}`);
    entry.content = Buffer.from(text, 'utf8');
    entry.size = entry.content.length;
  };

  // Append fonts, fills, borders and cell formats after the ones SheetJS wrote.
  let stylesXml = readPart('xl/styles.xml');
  const count = (tag) => Number(new RegExp(`<${tag} count="(\\d+)"`).exec(stylesXml)[1]);
  const append = (tag, items) => {
    const start = count(tag);
    stylesXml = stylesXml
      .replace(new RegExp(`<${tag} count="\\d+"`), `<${tag} count="${start + items.length}"`)
      .replace(`</${tag}>`, `${items.join('')}</${tag}>`);
    return start;
  };
  const fontBase = append('fonts', [
    '<font><b/><sz val="16"/><color rgb="FF14532D"/><name val="Calibri"/></font>',
    '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
    '<font><b/><sz val="12"/><color theme="1"/><name val="Calibri"/></font>',
    '<font><i/><sz val="12"/><color rgb="FF6B7280"/><name val="Calibri"/></font>',
    '<font><sz val="12"/><color rgb="FFB91C1C"/><name val="Calibri"/></font>',
  ]);
  const fillBase = append('fills', [
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1A7A4C"/><bgColor indexed="64"/></patternFill></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE2F3EA"/><bgColor indexed="64"/></patternFill></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/><bgColor indexed="64"/></patternFill></fill>',
  ]);
  const borderBase = append('borders', [
    '<border><left/><right/><top style="thin"><color rgb="FF1A7A4C"/></top><bottom style="double"><color rgb="FF1A7A4C"/></bottom><diagonal/></border>',
    '<border><left style="thin"><color rgb="FFB7C4BD"/></left><right style="thin"><color rgb="FFB7C4BD"/></right><top style="thin"><color rgb="FFB7C4BD"/></top><bottom style="thin"><color rgb="FFB7C4BD"/></bottom><diagonal/></border>',
  ]);
  const xf = (font, fill, border, align = '', numFmt = 0) =>
    `<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0" applyFont="1" applyFill="1" applyBorder="1"${
      numFmt ? ' applyNumberFormat="1"' : ''
    }${align ? ` applyAlignment="1"><alignment ${align}/></xf>` : '/>'}`;
  const xfBase = append('cellXfs', [
    xf(fontBase, 0, 0, 'horizontal="center" vertical="center"'), // title
    xf(fontBase + 1, fillBase, 0, 'horizontal="center" vertical="center"'), // header
    xf(fontBase + 2, 0, borderBase), // total label
    xf(fontBase + 2, 0, borderBase, '', 4), // total amount #,##0.00
    xf(fontBase + 3, 0, 0, 'wrapText="1" vertical="top"'), // note
    xf(0, fillBase + 1, borderBase + 1), // band
    xf(fontBase + 4, fillBase + 2, 0), // warning text
    xf(fontBase + 2, 0, borderBase, '', 10), // total percent 0.00%
  ]);
  const styleIds = { title: 0, header: 1, totalLabel: 2, totalAmount: 3, note: 4, band: 5, warning: 6, totalPercent: 7 };
  writePart('xl/styles.xml', stylesXml);

  Object.entries(styles).forEach(([sheetIndex, cells]) => {
    const path = `xl/worksheets/sheet${Number(sheetIndex) + 1}.xml`;
    let xml = readPart(path);
    for (const [ref, name] of Object.entries(cells)) {
      const s = xfBase + styleIds[name];
      const withStyle = new RegExp(`<c r="${ref}"([^>]*?) s="\\d+"`);
      xml = withStyle.test(xml) ? xml.replace(withStyle, `<c r="${ref}"$1 s="${s}"`) : xml.replace(`<c r="${ref}"`, `<c r="${ref}" s="${s}"`);
    }
    const view = [];
    if (hideGridlines?.includes(Number(sheetIndex))) view.push(' showGridLines="0"');
    const pane = frozen?.[sheetIndex];
    const paneXml = pane ? `<pane ySplit="${pane.rows}" topLeftCell="A${pane.rows + 1}" activePane="bottomLeft" state="frozen"/>` : '';
    xml = xml.replace(/<sheetView([^>]*?)\/>/, (_, attrs) => (paneXml ? `<sheetView${attrs}${view.join('')}>${paneXml}</sheetView>` : `<sheetView${attrs}${view.join('')}/>`));
    writePart(path, xml);
  });

  return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'buffer', compression: true });
}

function salesReport() {
  const rand = random(2024);
  const rows = [['Order ID', 'Date', 'City', 'Product', 'Quantity', 'Unit Price', 'Total', 'Paid', 'Notes']];
  const byCity = new Map(CITIES.map((c) => [c, { orders: 0, units: 0, revenue: 0 }]));
  const N = 5000;
  for (let i = 0; i < N; i++) {
    const r = i + 2;
    const product = Object.keys(PRODUCTS)[Math.floor(rand() * 8)];
    const city = CITIES[Math.floor(rand() * CITIES.length)];
    const qty = 1 + Math.floor(rand() * 12);
    const price = PRODUCTS[product];
    const stats = byCity.get(city);
    stats.orders++;
    stats.units += qty;
    stats.revenue += qty * price;
    rows.push([
      `ORD-${10001 + i}`,
      { t: 'n', v: 45292 + Math.floor(rand() * 366), z: 'yyyy-mm-dd' },
      city,
      product,
      qty,
      { t: 'n', v: price, z: '#,##0.00' },
      { t: 'n', v: qty * price, f: `E${r}*F${r}`, z: '#,##0.00' },
      rand() > 0.12,
      i % 97 === 0 ? 'Priority delivery requested by the customer; call before dispatch.' : '',
    ]);
  }
  const transactions = XLSX.utils.aoa_to_sheet(rows);
  transactions['!cols'] = [{ wch: 11 }, { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 8 }, { wch: 11 }, { wch: 13 }, { wch: 6 }, { wch: 28 }];

  const total = [...byCity.values()].reduce((sum, c) => sum + c.revenue, 0);
  const summaryRows = [
    ['Quarterly Sales Summary — FY 2024'],
    ['Revenue by city, generated from the Transactions sheet.'],
    ['City', 'Orders', 'Units', 'Revenue', 'Share'],
    ...CITIES.map((city, k) => {
      const c = byCity.get(city);
      const r = k + 4;
      return [
        city,
        { t: 'n', v: c.orders, f: `COUNTIF(Transactions!C:C,A${r})` },
        { t: 'n', v: c.units, f: `SUMIF(Transactions!C:C,A${r},Transactions!E:E)` },
        { t: 'n', v: c.revenue, f: `SUMIF(Transactions!C:C,A${r},Transactions!G:G)`, z: '#,##0.00' },
        { t: 'n', v: c.revenue / total, f: `D${r}/D$${CITIES.length + 4}`, z: '0.0%' },
      ];
    }),
    ['Total', { t: 'n', v: N, f: `SUM(B4:B${CITIES.length + 3})` }, null, { t: 'n', v: total, f: `SUM(D4:D${CITIES.length + 3})`, z: '#,##0.00' }, { t: 'n', v: 1, z: '0.0%' }],
    [],
    ['Checks'],
    ['Division by zero', { t: 'e', v: 0x07, f: '1/0' }],
    ['Lookup miss', { t: 'e', v: 0x2a, f: 'VLOOKUP("Goa",A4:D13,4,FALSE)' }],
    ['Report approved', true],
    ['Generated', { t: 'n', v: 45565.5, z: 'dd mmm yyyy hh:mm' }],
  ];
  const summary = XLSX.utils.aoa_to_sheet(summaryRows);
  summary['!merges'] = [XLSX.utils.decode_range('A1:E1'), XLSX.utils.decode_range('A2:E2')];
  summary['!cols'] = [{ wch: 18 }, { wch: 10 }, { wch: 10 }, { wch: 16 }, { wch: 9 }];
  summary['!rows'] = [{ hpt: 30 }, { hpt: 18 }, { hpt: 20 }];

  const wide = XLSX.utils.aoa_to_sheet(
    Array.from({ length: 60 }, (_, r) => Array.from({ length: 60 }, (_, c) => (r === 0 ? `Metric ${c + 1}` : Math.round(random(r * 100 + c)() * 1000) / 10))),
  );
  const lookup = XLSX.utils.aoa_to_sheet([['Code', 'City'], ...CITIES.map((c, k) => [k + 1, c])]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, summary, 'Summary');
  XLSX.utils.book_append_sheet(wb, transactions, 'Transactions');
  XLSX.utils.book_append_sheet(wb, wide, 'Wide Metrics');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), 'Notes');
  XLSX.utils.book_append_sheet(wb, lookup, 'Lookup');
  wb.Workbook = { Sheets: [{}, {}, {}, {}, { Hidden: 1 }] };
  wb.Props = { Title: 'Sales Report FY 2024', Author: 'Finance Team', LastAuthor: 'Sheaf', CreatedDate: new Date('2024-10-01T09:30:00Z') };

  const summaryStyles = { A1: 'title', A2: 'note', A3: 'header', B3: 'header', C3: 'header', D3: 'header', E3: 'header' };
  const totalRow = CITIES.length + 4;
  summaryStyles[`A${totalRow}`] = 'totalLabel';
  summaryStyles[`B${totalRow}`] = 'totalLabel';
  summaryStyles[`C${totalRow}`] = 'totalLabel';
  summaryStyles[`D${totalRow}`] = 'totalAmount';
  summaryStyles[`E${totalRow}`] = 'totalPercent';
  CITIES.forEach((_, k) => k % 2 === 1 && (summaryStyles[`A${k + 4}`] = 'band'));
  summaryStyles[`A${totalRow + 3}`] = 'warning';
  const txStyles = Object.fromEntries('ABCDEFGHI'.split('').map((col) => [`${col}1`, 'header']));

  const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true });
  return patchPackage(buffer, {
    styles: { 0: summaryStyles, 1: txStyles, 2: { A1: 'header' } },
    frozen: { 1: { rows: 1 }, 2: { rows: 1 } },
    hideGridlines: [0],
  });
}

function citiesCsv() {
  const lines = [
    'City,State,Population,Area (km²),Founded,Capital,Notes',
    'Mumbai,Maharashtra,12442373,603.4,1507,FALSE,"Financial capital, ""City of Dreams"""',
    'Delhi,Delhi,11034555,1484,-,TRUE,Seat of the national government',
    'Bengaluru,Karnataka,8443675,741,1537,FALSE,"Known as the ""Garden City""\nand the Silicon Valley of India"',
    'Chennai,Tamil Nadu,4646732,426,1639,FALSE,Gateway to South India',
    'Kolkata,West Bengal,4496694,206.1,1690,FALSE,City of Joy',
    'Zürich (test),—,0415,,,,Leading zero and unicode kept as text',
    'मुंबई,महाराष्ट्र,,,,,Devanagari text',
  ];
  return '﻿' + lines.join('\r\n') + '\r\n';
}

function legacyXls() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Year'],
      ['North', 120, 135, 128, 160, { t: 'n', v: 543, f: 'SUM(B2:E2)' }],
      ['South', 98, 110, 121, 140, { t: 'n', v: 469, f: 'SUM(B3:E3)' }],
      ['East', 76, 81, 95, 99, { t: 'n', v: 351, f: 'SUM(B4:E4)' }],
      ['West', 142, 150, 149, 171, { t: 'n', v: 612, f: 'SUM(B5:E5)' }],
    ]),
    'Regions',
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Updated', { t: 'n', v: 45000, z: 'mm/dd/yyyy' }]]), 'About');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xls' });
}

function largeXlsx(rows) {
  const rand = random(7);
  const data = [['ID', 'Date', 'City', 'Product', 'Qty', 'Price', 'Amount', 'Channel', 'Rep', 'Region', 'Discount', 'Status']];
  const channels = ['Online', 'Retail', 'Partner'];
  const statuses = ['Delivered', 'Shipped', 'Pending', 'Returned'];
  for (let i = 0; i < rows; i++) {
    const qty = 1 + Math.floor(rand() * 20);
    const price = Math.round(rand() * 50000) / 100;
    data.push([
      i + 1,
      { t: 'n', v: 44927 + Math.floor(rand() * 730), z: 'yyyy-mm-dd' },
      CITIES[Math.floor(rand() * 10)],
      Object.keys(PRODUCTS)[Math.floor(rand() * 8)],
      qty,
      price,
      Math.round(qty * price * 100) / 100,
      channels[Math.floor(rand() * 3)],
      `Rep ${1 + Math.floor(rand() * 250)}`,
      ['North', 'South', 'East', 'West'][Math.floor(rand() * 4)],
      Math.round(rand() * 30) / 100,
      statuses[Math.floor(rand() * 4)],
    ]);
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data, { dense: true }), 'Orders');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true });
}

function largeCsv(rows) {
  const rand = random(11);
  const out = ['id,timestamp,city,sensor,temperature,humidity,pressure,status,battery,firmware'];
  for (let i = 0; i < rows; i++) {
    out.push(
      [
        i + 1,
        new Date(Date.UTC(2025, 0, 1) + i * 60_000).toISOString(),
        CITIES[i % 10],
        `S-${1000 + (i % 400)}`,
        (15 + rand() * 25).toFixed(2),
        (30 + rand() * 60).toFixed(1),
        (990 + rand() * 40).toFixed(1),
        rand() > 0.97 ? 'ALERT' : 'OK',
        Math.floor(rand() * 100),
        `v${1 + (i % 3)}.${i % 10}`,
      ].join(','),
    );
  }
  return out.join('\n') + '\n';
}

const write = (name, data) => {
  writeFileSync(join(outDir, name), data);
  console.log(`  samples/${name}`);
};

console.log('Writing samples:');
const report = salesReport();
write('sales-report.xlsx', report);
write('cities.csv', citiesCsv());
write('legacy.xls', legacyXls());
write('corrupted.xlsx', report.subarray(0, Math.floor(report.length * 0.6)));
if (large) {
  write('large-100k.xlsx', largeXlsx(100_000));
  write('large-500k.csv', largeCsv(500_000));
}
