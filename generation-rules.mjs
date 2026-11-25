import { TEMPLATE_HEADERS } from './template-headers.mjs';

export const RULES_PROMPT = `IPAV2 safety rules (all customers): LOCVALID=X; LOCTYPE=P or DC.
Preserve uploaded IDs, UOM, customer region, sales area and product-to-production-location mappings.
Every location-product needs PLANNER=1 and a registered PLUNITID. Include BOM components as products.
Never invent circular BOMs, self-loop transport routes or all-to-all plant production assignments.
RESTYPE text does NOT prove IBP Production Category. Do not claim Resource Type 1 is production.
Production Category and model-specific Capacity KF definitions must be verified by the user.
Use DC (not D). Give each product productionLocationIds and uomId; customers need region and salesArea.
Company-name analysis is synthetic demo data, never actual customer transactions.`;

export class GenerationValidationError extends Error {
  constructor(errors) {
    super(errors.join('\n'));
    this.name = 'GenerationValidationError';
    this.errors = errors;
  }
}

export const csvRow = values => values.map(v => {
  const s = String(v ?? '');
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}).join(',');

export function readCsv(text) {
  const clean = text.replace(/^\uFEFF/, '');
  const sep = clean.match(/^sep=(.)\r?\n/i);
  const body = sep ? clean.slice(sep[0].length) : clean;
  const delimiter = sep?.[1] || (clean.split('\n')[0].includes(';') ? ';' : ',');
  const rows = []; let row = [], cell = '', quoted = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      if (quoted && body[i + 1] === '"') { cell += '"'; i++; }
      else quoted = !quoted;
    } else if (!quoted && ch === delimiter) { row.push(cell); cell = ''; }
    else if (!quoted && (ch === '\n' || ch === '\r')) {
      if (ch === '\r' && body[i + 1] === '\n') i++;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += ch;
  }
  if (quoted) throw new GenerationValidationError(['CSV contains an unterminated quoted field']);
  if (row.length || cell) { row.push(cell); rows.push(row); }
  const [header = [], ...data] = rows;
  return { header, rows: data.filter(r => r.some(v => v !== '')).map(r => {
    if (r.length !== header.length) throw new GenerationValidationError(['CSV column count mismatch']);
    return Object.fromEntries(header.map((k, i) => [k, r[i]]));
  }) };
}

export const writeCsv = (header, rows) => 'sep=,\n' + csvRow(header) + '\n' +
  rows.map(r => csvRow(header.map(k => r[k]))).join('\n') + (rows.length ? '\n' : '');

export function prepareGeneration(analysis, selections) {
  if (!analysis || !selections) throw new GenerationValidationError(['analysis and selections are required']);
  analysis = structuredClone(analysis);
  selections = structuredClone(selections);
  const errors = [], warnings = [];
  const options = { ...analysis.generationOptions, ...selections.generationOptions };
  const families = (analysis.productFamilies || []).filter(f => selections.selectedFamilyIds?.includes(f.id));
  const locs = (analysis.locations || []).filter(l => selections.selectedLocationIds?.includes(l.id));
  const customers = selections.customers || [];
  if (!families.length || !locs.length || !customers.length) errors.push('Select products, locations and customers before generating.');
  const id = (value, label) => {
    if (!value || /[|,"\r\n]/.test(String(value))) errors.push(`${label}: missing or invalid ID`);
  };
  const unique = (items, key, label) => {
    const seen = new Set();
    for (const item of items) {
      id(item[key], label);
      if (seen.has(item[key])) errors.push(`${label}: duplicate ${item[key]}`);
      seen.add(item[key]);
    }
  };
  unique(locs, 'id', 'LOCID'); unique(customers, 'custId', 'CUSTID');
  for (const l of locs) {
    const t = String(l.locType || 'P').trim().toUpperCase();
    l.locType = /^(P|PLANT|FACTORY|MFG|PROD.*)$/.test(t) ? 'P' :
      /^(D|DC|DIST.*|WAREHOUSE|WH|WHSE|LOGISTICS)$/.test(t) ? 'DC' : t;
    if (!['P', 'DC', 'S', 'C'].includes(l.locType)) errors.push(`${l.id}: unknown LOCTYPE ${t}`);
  }
  for (const c of customers) {
    if (!c.region) {
      c.region = c.country || analysis.company?.hq?.region;
      warnings.push(`${c.custId}: CUSTREGION inferred; verify the DBP grouping.`);
    }
    if (!c.region) errors.push(`${c.custId}: CUSTREGION required for DBP`);
    c.channel ||= 'Direct'; c.group ||= 'Industry';
  }
  const products = new Map();
  for (const f of families) {
    if (!f.code) errors.push(`${f.id}: product family code is required for DBP`);
    for (const kind of ['fg', 'semi', 'raw']) for (const p of f[kind] || []) {
      id(p.id, 'PRDID');
      if (products.has(p.id)) errors.push(`Duplicate PRDID: ${p.id}`);
      p.uomId ||= options.baseUom || 'PC';
      id(p.uomId, 'UOMID');
      products.set(p.id, { p, f, kind });
    }
  }
  // Uploaded BOM components must be registered before location-product / BOM upload.
  const bom = (analysis.bom || []).filter(b => products.has(b.parent));
  for (const b of bom) {
    id(b.component, 'BOM component');
    if (!products.has(b.component)) {
      const f = products.get(b.parent).f;
      const p = { id: b.component, desc: b.component, uomId: b.uom || 'PC' };
      (f.raw ||= []).push(p); products.set(p.id, { p, f, kind: 'raw' });
      warnings.push(`${p.id}: registered missing BOM component; verify description and UOM.`);
    }
    if (!Number.isFinite(Number(b.qty)) || !(Number(b.qty) > 0)) errors.push(`BOM ${b.parent} → ${b.component}: quantity must be positive`);
  }
  const edges = new Map();
  for (const b of bom) (edges.get(b.parent) || edges.set(b.parent, []).get(b.parent)).push(b.component);
  const active = new Set(), done = new Set();
  const visit = (p, trail = []) => {
    if (active.has(p)) { errors.push(`1020 circular BOM: ${[...trail, p].join(' → ')}`); return; }
    if (done.has(p)) return;
    active.add(p); for (const c of edges.get(p) || []) visit(c, [...trail, p]);
    active.delete(p); done.add(p);
  };
  for (const p of edges.keys()) visit(p);
  const plants = locs.filter(l => l.locType === 'P');
  const resolveLoc = value => locs.find(l => l.id === value || l.id === `LOC-${value}`)?.id;
  const production = new Set(), stocked = new Set();
  const pair = (p, l) => `${p}|${l}`;
  for (const { p, kind } of products.values()) {
    let declared = options.productionLocations?.[p.id] || p.productionLocationIds || (p.locationId ? [p.locationId] : []);
    if (!Array.isArray(declared)) { errors.push(`${p.id}: productionLocations must be an array`); declared = []; }
    if (!declared.length) declared = [...new Set([
      ...(analysis.routing || []).filter(r => r.product === p.id).map(r => r.locationId || r.plant),
      ...bom.filter(b => b.parent === p.id || b.component === p.id).map(b => b.plant).filter(Boolean),
    ])];
    let mapped = declared.map(resolveLoc).filter(Boolean);
    for (const value of declared) if (!resolveLoc(value)) errors.push(`${p.id}: mapped location ${value} is not selected/registered`);
    if (!mapped.length && kind !== 'raw') {
      if (plants.length === 1 || (!analysis._importedFromExcel && plants.length)) {
        mapped = [plants[0].id];
        if (plants.length > 1) warnings.push(`${p.id}: proposed production plant ${mapped[0]}; confirm before upload.`);
      } else if (plants.length > 1) errors.push(`${p.id}: multiple plants; provide productionLocationIds or BOM/routing plant mapping`);
    }
    for (const l of mapped) {
      stocked.add(pair(p.id, l));
      if (kind !== 'raw') {
        if (!plants.some(pl => pl.id === l)) errors.push(`${p.id}: production location ${l} is not a plant`);
        else production.add(pair(p.id, l));
      }
    }
    for (const l of p.stockLocationIds || []) {
      if (!resolveLoc(l)) errors.push(`${p.id}: unknown stocking location ${l}`);
      else stocked.add(pair(p.id, resolveLoc(l)));
    }
    if (kind === 'fg' && !p.stockLocationIds) for (const l of locs.filter(l => l.locType === 'DC')) stocked.add(pair(p.id, l.id));
  }
  for (const b of bom) {
    const parentLocs = [...production].filter(k => k.startsWith(`${b.parent}|`)).map(k => k.split('|')[1]);
    if (b.plant && !parentLocs.includes(resolveLoc(b.plant))) errors.push(`BOM ${b.parent}: plant ${b.plant} conflicts with its production mapping`);
    for (const l of parentLocs) stocked.add(pair(b.component, l));
  }
  if (!analysis._importedFromExcel && !bom.length) {
    // Keep synthetic demo components consistent with the synthetic BOM generator.
    for (const f of families) {
      const links = [
        ...(f.fg || []).map(p => [p, f.semi?.[0] || f.raw?.[0]]),
        ...(f.semi || []).map(p => [p, f.raw?.[1] || f.raw?.[0]]),
      ];
      for (const [parent, component] of links) if (component && parent.id !== component.id) {
        for (const k of production) if (k.startsWith(`${parent.id}|`)) stocked.add(pair(component.id, k.split('|')[1]));
      }
    }
  }
  const planningUnit = options.planningUnitId || 'IBP1'; id(planningUnit, 'PLUNITID');
  const planningUnits = new Set([planningUnit]);
  for (const { p } of products.values()) for (const unit of Object.values(p.planningUnitsByLocation || {})) {
    id(unit, 'PLUNITID'); planningUnits.add(unit);
  }
  const confirmed = new Set(options.confirmedProductionResourceIds || []);
  const typeCode = String(options.productionResourceTypeCode ?? '');
  if (confirmed.size && (!typeCode || typeCode === '1')) errors.push('1080: confirmed resources require the verified model Production RESOURCETYPE code (not 1).');
  const generatedResources = new Set(locs.flatMap(l => ['ASSEMBLY', 'PACK', 'SUBASSEMBLY', 'STORAGE', 'TRANSPORT', 'HANDLING'].map(t => `${l.id} ${t}`)));
  for (const r of confirmed) if (!generatedResources.has(r) || !/ (ASSEMBLY|PACK|SUBASSEMBLY)$/.test(r))
    errors.push(`${r}: confirmation must name an exported production Resource ID`);
  for (const r of analysis.resources || []) {
    if (String(r.resourceType ?? r.resType ?? '').trim() === '1') {
      if (['STORAGE', 'TRANSPORT', 'HANDLING'].includes(String(r.resType || '').toUpperCase()))
        warnings.push(`${r.resourceId}: non-production Type 1 is not connected to production sources.`);
      else errors.push(`${r.resourceId}: Resource Type 1 cannot be used as production (1080)`);
    }
  }
  const actuals = (analysis.actuals || []).filter(r => products.has(r.PRDID));
  const actualKeys = new Set();
  for (const r of actuals) {
    if (!stocked.has(pair(r.PRDID, r.LOCID)) || !customers.some(c => c.custId === r.CUSTID))
      errors.push(`Actuals ${r.PRDID}/${r.LOCID}/${r.CUSTID}: missing selected master-data combination`);
    const date = new Date(r.KEYFIGUREDATE);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(r.KEYFIGUREDATE) || !Number.isFinite(date.getTime()) || date.toISOString().slice(0, 10) !== r.KEYFIGUREDATE)
      errors.push(`Actuals ${r.PRDID}: invalid KEYFIGUREDATE ${r.KEYFIGUREDATE}`);
    if (!Number.isFinite(Number(r.ACTUALSQTY)) || Number(r.ACTUALSQTY) < 0) errors.push(`Actuals ${r.PRDID}: invalid quantity`);
    const key = `${r.PRDID}|${r.LOCID}|${r.CUSTID}|${r.KEYFIGUREDATE}`;
    if (actualKeys.has(key)) errors.push(`Duplicate actuals key: ${key}`);
    actualKeys.add(key);
    const c = customers.find(c => c.custId === r.CUSTID);
    if (r.SALESAREA && c) {
      if (c.salesArea && c.salesArea !== r.SALESAREA) errors.push(`${c.custId}: actuals Sales Area differs from customer mapping`);
      else c.salesArea = r.SALESAREA;
    }
  }
  if (actuals.length) warnings.push('Uploaded actuals replace synthetic history; verify a minimum two-year history per product/location/customer. No synthetic history is backfilled. Future forecasts remain demo proposals.');
  warnings.push('RESTYPE strings do not prove IBP Production Category. Unconfirmed production-resource links are omitted (1080 prevention).');
  warnings.push('Capacity formula in the source memo differs from the supplied template: ZSRESOURCEWORKDAYS/ZSCAPAMULTIPLIER and ZSCAPASUPLY* are not assumed aliases of ZSSHIFTHRS/ZSRESOURCE/ZSCAPASUPPLY*. Verify the active planning model and Base Version.');
  warnings.push('DBP: create Driver/ROP master data and set Include in Plan=1 in IBP; run Statistical Forecast then confirm Consensus Demand. Generated history is synthetic demo data, not uploaded customer actuals.');
  warnings.push('CP1BU/CP1PRODUCTBU and planning-unit description/UOM need tenant-exported templates; no unverified columns are generated. Ensure UOM/Currency reference masters exist before loading.');
  if (errors.length) throw new GenerationValidationError([...new Set(errors)]);
  return { analysis, families, locs, customers, bom, actuals, options, planningUnit, planningUnits, products, production, stocked, confirmed, typeCode, warnings, files: [] };
}

export function applyRules(filePath, content, ctx, salesArea) {
  if (!content) return content;
  const { header, rows } = readCsv(content);
  const base = filePath.split('/').pop();
  const templateName = base.replace(/_NEW\.csv$/i, '');
  const templateHeader = ctx.options.templateHeaders?.[templateName] || TEMPLATE_HEADERS[templateName];
  if (templateHeader) {
    const columns = Array.isArray(templateHeader) ? templateHeader : templateHeader.split(',');
    if (new Set(columns).size !== columns.length || columns.some(k => !/^[A-Z][A-Z0-9_]*$/.test(k)))
      throw new GenerationValidationError([`${base}: invalid/duplicate template columns`]);
    for (const k of header) if (!columns.includes(k)) throw new GenerationValidationError([`${base}: ${k} is absent from the supplied template`]);
    header.splice(0, header.length, ...columns);
  }
  const isProduction = base.includes('SOURCEPRODUCTIONHEADER') || base.includes('PRODUCTIONSOURCERESOURCE') ||
    base.includes('WKPRODLOCSRC') || base.includes('WKRESPRODLOCSRC');
  const pair = r => `${r.PRDID}|${r.LOCID}`;
  let filtered = rows.filter(r => {
    if (r.PRDID && r.LOCID && !ctx.stocked.has(pair(r))) return false;
    if (isProduction && r.PRDID && !ctx.production.has(pair(r))) return false;
    if (r.LOCFR && (r.LOCFR === r.LOCID || !ctx.stocked.has(`${r.PRDID}|${r.LOCFR}`) || ctx.locs.find(l => l.id === r.LOCID)?.locType !== 'DC')) return false;
    if (r.RESID && (base.includes('PRODUCTIONSOURCERESOURCE') || base.includes('RESOURCELOCATIONPRODUCT') || base.includes('WKRESPRODLOCSRC')) && !ctx.confirmed.has(r.RESID)) return false;
    return true;
  });
  if (base.includes('PLANNINGUNIT')) filtered = [...ctx.planningUnits].map(PLUNITID => ({ PLUNITID }));
  if (base.includes('WKPRODLOCSACUST') && ctx.actuals.length) {
    const key = r => `${r.PRDID}|${r.LOCID}|${r.CUSTID}|${r.KEYFIGUREDATE}`;
    const keys = new Set(ctx.actuals.map(key));
    filtered = [...filtered.filter(r => r.ACTUALSQTY === '' && !keys.has(key(r))), ...ctx.actuals];
  }
  if (base.includes('LOCATIONPRODUCT_SOP')) {
    // Replace the old partial/reordered header with the tenant-exported template.
    header.splice(0, header.length, ...LOCATION_PRODUCT_HEADER.split(','));
    filtered = [...ctx.stocked].map(k => {
      const [PRDID, LOCID] = k.split('|');
      const p = ctx.products.get(PRDID).p;
      for (const field of ['distributionType', 'stockingNodeType']) if (String(p[field] || '').length > 1)
        throw new GenerationValidationError([`${PRDID}: ${field} has a one-character limit (010)`]);
      return { PRDID, LOCID, PLANNER: '1', PLUNITID: p.planningUnitsByLocation?.[LOCID] || ctx.planningUnit,
        MRPTYPE: ctx.locs.find(l => l.id === LOCID).locType === 'DC' ? 'X' : '',
        DISTRIBUTIONTYPE: p.distributionType || '', STOCKINGNODETYPE: p.stockingNodeType || '', NAGINDEPNETWORK: '' };
    });
  }
  for (const r of filtered) {
    if (base.includes('WKPRODLOCSACUST')) {
      const factor = ctx.products.get(r.PRDID).p.disaggregationFactor ?? 1;
      if (!Number.isFinite(Number(factor)) || !(Number(factor) > 0)) throw new GenerationValidationError([`${r.PRDID}: disaggregationFactor must be positive`]);
      r.ZXDISAGGREGATIONFACTOR = factor;
    }
    if (header.includes('SALESAREA') && r.CUSTID) {
      const c = ctx.customers.find(c => c.custId === r.CUSTID);
      r.SALESAREA = c?.salesArea || salesArea(c?.channel);
    }
    if (base === '030_PRODUCT_NEW.CSV') { r.UOMID = ctx.products.get(r.PRDID).p.uomId; r.UOMDESCR = r.UOMID; }
    if (base.includes('SOURCEPRODUCTIONHEADER')) r.CPPRODUCTIONLINE = ''; // never trust unverified routing work centers
    if (base.includes('RESOURCELOCATION_NEW') && ctx.confirmed.has(r.RESID)) r.RESOURCETYPE = ctx.typeCode;
    if (base.includes('WKRESLOC') && r.RESID.endsWith(' HANDLING')) r.ZSCAPASUPPLYTRANSPORT = r.ZSHANDLINGCAPACITY;
    if (base.includes('WKRESLOC') && !ctx.options.capacityProfiles?.[r.RESID]) {
      r.ZSSHIFTHRS = ''; r.ZSRESOURCE = ''; // semantics not proven by memo/template equivalence
    }
    if (base.includes('WKRESLOC')) {
      for (const [kf, value] of Object.entries(ctx.options.capacityProfiles?.[r.RESID] || {})) {
        if (!header.includes(kf) || kf === 'CAPASUPPLY' || ['LOCID', 'KEYFIGUREDATE', 'RESID'].includes(kf))
          throw new GenerationValidationError([`${r.RESID}: unsupported capacity input ${kf}; export the active tenant template`]);
        if (!Number.isFinite(Number(value)) || Number(value) < 0 || (kf === 'ZSEFFICIENCY' && Number(value) > 1))
          throw new GenerationValidationError([`${r.RESID}: invalid ${kf} input`]);
        r[kf] = value;
      }
      r.CAPASUPPLY = '';
    }
    for (const [key, value] of Object.entries(r)) if (['undefined', 'NaN', 'Infinity'].includes(String(value)))
      throw new GenerationValidationError([`${base}: invalid ${key} value ${value}`]);
  }
  // Re-normalize transport/customer quotas after filtering infeasible locations.
  if (base.includes('SOURCELOCATION') || base.includes('SOURCECUSTOMER')) {
    const groups = new Map();
    for (const r of filtered) {
      const key = base.includes('SOURCELOCATION') ? `${r.PRDID}|${r.LOCID}` : `${r.PRDID}|${r.CUSTID}`;
      (groups.get(key) || groups.set(key, []).get(key)).push(r);
    }
    for (const group of groups.values()) for (const r of group) r[base.includes('SOURCELOCATION') ? 'TRATIO' : 'CRATIO'] = 1 / group.length;
  }
  if (header.includes('SOURCEID') && !base.includes('SOURCEPRODUCTIONHEADER')) {
    filtered = filtered.filter(r => ctx.sourceIds?.has(r.SOURCEID));
  }
  if (base.includes('SOURCEPRODUCTIONHEADER')) ctx.sourceIds = new Set(filtered.map(r => r.SOURCEID));
  if (base === '100_UOMCONVERSIONFACTOR_NEW.CSV') {
    // Never invent 1:1 conversions between physical units (KG, M2, EA, ...).
    filtered = [...ctx.products.values()].flatMap(({ p }) => [
      { PRDID: p.id, UOMTOID: p.uomId, UOMCONVERSIONFACTOR: 1 },
      ...Object.entries(p.uomConversions || {}).map(([UOMTOID, factor]) => {
        if (!Number.isFinite(Number(factor)) || !(Number(factor) > 0)) throw new GenerationValidationError([`${p.id}: invalid UOM conversion to ${UOMTOID}`]);
        return { PRDID: p.id, UOMTOID, UOMCONVERSIONFACTOR: factor };
      }),
    ]);
  }
  if (base === '200_SA_NEW.CSV') {
    filtered = [...new Set(ctx.customers.map(c => c.salesArea || salesArea(c.channel)))].map(SALESAREA => {
      const [SALESORG, DISCTCHANNEL, DIVISION, extra] = SALESAREA.split('/');
      if (!SALESORG || !DISCTCHANNEL || !DIVISION || extra || DISCTCHANNEL.length > 2)
        throw new GenerationValidationError([`Invalid salesArea ${SALESAREA}; expected SALESORG/CHANNEL/DIVISION`]);
      return { SALESAREA, SALESORG, DISCTCHANNEL, DIVISION, CPCOMPANYCODE: SALESORG };
    });
  }
  ctx.files.push({ file: filePath, rows: filtered.length });
  return writeCsv(header, filtered);
}

const LOCATION_PRODUCT_HEADER = 'LOCID,PRDID,CPLOCPRDATT01,CPLOCPRDATT02,CPLOCPRDATT03,CRITICALOPERATION,DECOUPLINGPOINTREASON,DISCINVFLAG,DISTRIBUTIONTYPE,FOOTPRINTPERUNIT,FROZENHORIZON,GREENZONECALCULATIONTYPE,INVENTORYHOLDINGCOSTRATE,INVENTORYTARGETVIOLATIONCOST,INVHOLDINGCOSTPOLICY,INVTARGETVIOLCOSTPOLICY,IOECHELONID,LEADTIMERANK,LOTSIZEPOLICY,LPQUOTACOVERAGE,MAXINTERNALSERVICELEVEL,MAXINVENTORYVIOLATIONCOST,MININTERNALSERVICELEVEL,MINSHELFLIFEINDEPDEM,MINSHELFLIFEINVTARGET,MINSHELFLIFEINVTARGETTRIGGER,MOQ,MRPTYPE,NAGINDEPNETWORK,NETWORKID,NONDELIVERYCOSTPOLICY,PBR,PLANNER,PLUNITID,RECDECOUPPOINTIND,SAFETYDAYSIND,SAFETYDAYSSUPPLY,SAFETYSTOCK,SAFETYSTOCKPOLICY,SERVICELEVELTYPE,STOCKINGNODETYPE,TRANSPORTSUBSTPOLICY,VARIABILITYRANK,ZXTSPROPPRDLOC';
