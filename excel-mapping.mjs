import XLSX from 'xlsx';

// Prefer exported technical column IDs to an AI guess, regardless of customer/sheet name.
export function technicalMappings(wb) {
  const result = {};
  const specs = [
    ['locationSheet', ['LOCID', 'LOCDESCR'], { idCol: 'LOCID', nameCol: 'LOCDESCR', cityCol: 'LOCCITY', countryCol: 'LOCCOUNTRY', regionCol: 'LOCREGION', typeCol: 'LOCTYPE', latCol: 'GEOLATITUDE', lonCol: 'GEOLONGITUDE' }],
    ['customerSheet', ['CUSTID', 'CUSTDESCR'], { idCol: 'CUSTID', nameCol: 'CUSTDESCR', countryCol: 'CUSTCOUNTRY', regionCol: 'CUSTREGION', channelCol: 'CUSTCHANNEL', salesAreaCol: 'SALESAREA', latCol: 'GEOLATITUDE', lonCol: 'GEOLONGITUDE' }],
    ['productSheet', ['PRDID', 'PRDDESCR'], { idCol: 'PRDID', nameCol: 'PRDDESCR', familyCol: 'PRDFAMILY', uomCol: 'UOMID', typeCol: 'MATTYPEID', locationCol: 'LOCID' }],
    ['resourceSheet', ['LOCID', 'RESID'], { locationCol: 'LOCID', resourceIdCol: 'RESID', resourceTypeCol: 'RESOURCETYPE', restypeCol: 'RESTYPE' }],
    ['locationProductSheet', ['LOCID', 'PRDID', 'PLUNITID'], { locationCol: 'LOCID', productCol: 'PRDID', planningUnitCol: 'PLUNITID', distributionTypeCol: 'DISTRIBUTIONTYPE', stockingNodeTypeCol: 'STOCKINGNODETYPE' }],
    ['actualsSheet', ['PRDID', 'LOCID', 'CUSTID', 'KEYFIGUREDATE', 'ACTUALSQTY'], { productCol: 'PRDID', locationCol: 'LOCID', customerCol: 'CUSTID', dateCol: 'KEYFIGUREDATE', qtyCol: 'ACTUALSQTY', salesAreaCol: 'SALESAREA' }],
  ];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      const cols = rows[i].map(v => String(v).trim().toUpperCase());
      for (const [key, required, fields] of specs) {
        if (result[key] || !required.every(c => cols.includes(c))) continue;
        const next = rows[i + 1] || [];
        const descriptionRow = next.some(v => /^(#|Version|.* ID\*)$/i.test(String(v)));
        result[key] = { sheetName: name, headerRows: i + 1 + Number(descriptionRow),
          ...Object.fromEntries(Object.entries(fields).map(([field, column]) => [field, cols.includes(column) ? cols.indexOf(column) : null])) };
      }
    }
  }
  return result;
}

export function extractActuals(wb, mapping) {
  const m = mapping.actualsSheet;
  if (!m?.sheetName) return [];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[m.sheetName], { header: 1, defval: '' }).slice(m.headerRows);
  return rows.filter(r => r[m.productCol] && r[m.qtyCol] !== '').map(r => {
    const value = r[m.dateCol];
    const d = typeof value === 'number' ? XLSX.SSF.parse_date_code(value) : null;
    const KEYFIGUREDATE = d ? `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}` : String(value).trim();
    const loc = String(r[m.locationCol]).trim();
    return { PRDID: String(r[m.productCol]).trim(), LOCID: loc.startsWith('LOC-') ? loc : `LOC-${loc}`,
      CUSTID: String(r[m.customerCol]).trim(), KEYFIGUREDATE, ACTUALSQTY: Number(r[m.qtyCol]),
      SALESAREA: m.salesAreaCol != null ? String(r[m.salesAreaCol] ?? '').trim() : '' };
  });
}

export function extractTechnicalSources(wb) {
  const tables = [];
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    for (let i = 0; i < Math.min(rows.length, 5); i++) {
      const cols = rows[i].map(v => String(v).trim().toUpperCase());
      if (!cols.includes('SOURCEID') || !cols.includes('LOCID')) continue;
      const skip = (rows[i + 1] || []).some(v => /^(#|Version|.* ID\*)$/i.test(String(v))) ? 2 : 1;
      tables.push({ cols, data: rows.slice(i + skip).map(r => Object.fromEntries(cols.map((k, c) => [k, r[c] ?? '']))) });
      break;
    }
  }
  const sources = new Map();
  for (const t of tables) if (t.cols.includes('PRDID') && t.cols.includes('OUTPUTCOEFFICIENT')) {
    for (const r of t.data) if (r.SOURCEID && r.PRDID) sources.set(r.SOURCEID, r);
  }
  const bom = [], routing = [];
  for (const t of tables) for (const r of t.data) {
    const parent = sources.get(r.SOURCEID);
    if (!parent) continue;
    if (t.cols.includes('COMPONENTCOEFFICIENT') && r.PRDID) bom.push({ parent: parent.PRDID,
      component: r.PRDID, plant: r.LOCID, qty: Number(r.COMPONENTCOEFFICIENT), seq: String(r.SOURCEITMID || r.PRDID) });
    if (t.cols.includes('RESID') && r.RESID) routing.push({ product: parent.PRDID,
      plant: r.LOCID, locationId: r.LOCID, workCenter: r.RESID });
  }
  const productionMappings = [...sources.values()].map(r => ({ product: r.PRDID, locationId: r.LOCID, plant: r.LOCID }));
  return { bom, routing, productionMappings };
}

export function applyLocationProductMappings(wb, mapping, families, locations) {
  const m = mapping.locationProductSheet;
  if (!m?.sheetName) return {};
  const products = new Map(families.flatMap(f => [...f.fg, ...f.semi, ...f.raw]).map(p => [p.id, p]));
  const units = new Set();
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[m.sheetName], { header: 1, defval: '' }).slice(m.headerRows);
  for (const r of rows) {
    const p = products.get(String(r[m.productCol]).trim());
    const loc = String(r[m.locationCol] ?? '').trim();
    if (!p || !loc) continue;
    const locId = loc.startsWith('LOC-') ? loc : `LOC-${loc}`;
    (p.stockLocationIds ||= []).push(locId);
    if (locations.some(l => l.id === locId && l.locType === 'P')) (p.productionLocationIds ||= []).push(locId);
    const unit = String(r[m.planningUnitCol] ?? '').trim();
    if (unit) units.add(unit);
    p.distributionType = m.distributionTypeCol != null ? String(r[m.distributionTypeCol] ?? '') : '';
    p.stockingNodeType = m.stockingNodeTypeCol != null ? String(r[m.stockingNodeTypeCol] ?? '') : '';
  }
  // Keep multiple units on the product instead of silently collapsing the workbook to IBP1.
  for (const r of rows) {
    const p = products.get(String(r[m.productCol]).trim());
    const loc = String(r[m.locationCol] ?? '').trim();
    const unit = String(r[m.planningUnitCol] ?? '').trim();
    if (p && unit) (p.planningUnitsByLocation ||= {})[loc.startsWith('LOC-') ? loc : `LOC-${loc}`] = unit;
  }
  return units.size === 1 ? { planningUnitId: [...units][0] } : {};
}
