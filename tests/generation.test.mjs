import { test } from 'node:test';
import assert from 'node:assert/strict';
import AdmZip from 'adm-zip';
import XLSX from 'xlsx';
import { generateZip, parseExcelToIbpData } from '../server.mjs';
import { csvRow, readCsv, prepareGeneration } from '../generation-rules.mjs';
import { TEMPLATE_HEADERS } from '../template-headers.mjs';
import fs from 'node:fs';
import path from 'node:path';

function fixture(imported = false) {
  const analysis = {
    company: { name: 'Other Customer', hq: { region: 'EMEA' } }, _importedFromExcel: imported,
    locations: [
      { id: 'LOC-A', desc: 'Factory, "A"\nKorea', locType: 'Plant', region: 'APJ', country: 'KR', lat: 0, lon: 127 },
      { id: 'LOC-B', desc: 'Factory B', locType: 'P', region: 'APJ', country: 'KR' },
      { id: 'LOC-DC', desc: 'Warehouse', locType: 'DC', region: 'APJ', country: 'KR' },
    ],
    productFamilies: [{ id: 'F1', code: 'FAMILY 1', brand: 'BRAND', group: 'G',
      fg: [{ id: 'FG', desc: 'Product, "one"', uomId: 'M2', productionLocationIds: ['LOC-B'] }], semi: [], raw: [] }],
    bom: [{ parent: 'FG', component: 'RAW', plant: 'B', qty: 2, uom: 'KG', seq: '0010' }],
  };
  const selections = { customers: [{ custId: 'C1', custdescr: 'Customer, Ltd.', country: 'KR', region: 'APJ', channel: 'Direct', group: 'G', salesArea: '2020/12/01' }],
    selectedFamilyIds: ['F1'], selectedLocationIds: ['LOC-A', 'LOC-B', 'LOC-DC'] };
  return { analysis, selections };
}
const output = f => new AdmZip(generateZip(f.analysis, f.selections));
const csv = (zip, name) => {
  const e = zip.getEntries().find(e => e.entryName.endsWith(name));
  assert.ok(e, `Missing ${name}`);
  assert.ok(e.getData().toString('utf8').startsWith('\uFEFFsep=,'));
  return readCsv(e.getData().toString('utf8'));
};

test('CSV roundtrip preserves Korean, commas, quotes, newlines and empty cells', () => {
  const values = ['한글, "설명"\n두번째 줄', '', '0'];
  const parsed = readCsv('sep=,\nA,B,C\n' + csvRow(values));
  assert.deepEqual(Object.values(parsed.rows[0]), values);
});

test('common ZIP rules preserve UOM, add components, scope production and keep templates', () => {
  const f = fixture(true), zip = output(f);
  const products = csv(zip, '030_PRODUCT_NEW.CSV').rows;
  assert.equal(products.find(p => p.PRDID === 'FG').UOMID, 'M2');
  assert.equal(products.find(p => p.PRDID === 'FG').PRDDESCR, 'Product, "one"');
  assert.equal(products.find(p => p.PRDID === 'RAW').UOMID, 'KG');
  assert.equal(csv(zip, '020_LOCATION_NEW.CSV').rows[0].GEOLATITUDE, '0');
  const lp = csv(zip, '120_LOCATIONPRODUCT_SOP_DP_NEW.CSV');
  assert.deepEqual(lp.header.slice(0, 2), ['LOCID', 'PRDID']);
  assert.ok(lp.rows.every(r => r.PLANNER === '1' && r.PLUNITID === 'IBP1'));
  assert.ok(!lp.rows.some(r => r.PRDID === 'FG' && r.LOCID === 'LOC-A'));
  assert.ok(lp.rows.some(r => r.PRDID === 'RAW' && r.LOCID === 'LOC-B'));
  assert.equal(lp.rows.find(r => r.LOCID === 'LOC-DC').MRPTYPE, 'X');
  assert.deepEqual(csv(zip, '160_SOURCEPRODUCTIONHEADER_NEW.CSV').rows.map(r => r.LOCID), ['LOC-B']);
  assert.ok(csv(zip, '170_PRODUCTIONSOURCEITM_NEW.CSV').rows.every(r => r.PRDID === 'RAW' && r.LOCID === 'LOC-B'));
  assert.equal(csv(zip, '180_PRODUCTIONSOURCERESOURCE_NEW.CSV').rows.length, 0);
  assert.equal(csv(zip, 'IPAV2__WKRESPRODLOCSRC_NEW.csv').rows.length, 0);
  assert.ok(csv(zip, '150_SOURCELOCATION_SOP_DP_NEW.CSV').rows.every(r => r.LOCFR === 'LOC-B' && r.LOCID === 'LOC-DC'));
  assert.deepEqual(csv(zip, '100_UOMCONVERSIONFACTOR_NEW.CSV').rows.map(r => r.UOMTOID).sort(), ['KG', 'M2']);
  const demand = csv(zip, 'IPAV2__WKPRODLOCSACUST_NEW.csv').rows;
  assert.ok(demand.every(r => r.SALESAREA === '2020/12/01' && Number(r.ZXDISAGGREGATIONFACTOR) > 0));
  const dates = new Set(demand.filter(r => r.ACTUALSQTY !== '').map(r => r.KEYFIGUREDATE));
  assert.equal(dates.size, 105);
  const report = JSON.parse(zip.readAsText('VALIDATION_REPORT.json'));
  assert.ok(report.warnings.some(w => w.includes('synthetic')));
  assert.ok(zip.readAsText('UPLOAD_ORDER.txt').split('\n')[0].includes('PLANNINGUNIT'));
  assert.ok(zip.getEntry('GENERATION_GUIDELINES.md'));
  assert.equal(csv(zip, '200_SA_NEW.CSV').rows[0].SALESAREA, '2020/12/01');
  assert.ok(csv(zip, '140_SOURCECUSTOMER_NEW.CSV').rows.some(r => r.PRDID === 'FG' && r.LOCID === 'LOC-DC'));
  assert.ok(csv(zip, 'IPAV2__WKPRODLOCCUST_NEW.csv').rows.length > 0);
  assert.ok(csv(zip, 'IPAV2__WKRESLOC_NEW.csv').rows.filter(r => r.RESID.endsWith(' HANDLING')).every(r => r.ZSCAPASUPPLYTRANSPORT !== ''));
});

test('all expanded KF schemas exactly match snapshot and optional local tenant templates', () => {
  const zip = output(fixture());
  for (const [name, header] of Object.entries(TEMPLATE_HEADERS)) {
    assert.equal(csv(zip, `${name}_NEW.csv`).header.join(','), header);
    if (process.env.IPAV2_TEMPLATE_DIR) {
      const text = fs.readFileSync(path.join(process.env.IPAV2_TEMPLATE_DIR, `${name}_TEMPLATE.csv`), 'utf8');
      assert.equal(header, readCsv(text).header.join(','));
    }
  }
});

test('uploaded actuals replace demo history without inventing missing weeks', () => {
  const f = fixture(true);
  f.analysis.actuals = [{ PRDID: 'FG', LOCID: 'LOC-DC', CUSTID: 'C1', KEYFIGUREDATE: '2024-11-25', ACTUALSQTY: 123, SALESAREA: '2020/12/01' }];
  const zip = output(f);
  const history = csv(zip, 'IPAV2__WKPRODLOCSACUST_NEW.csv').rows.filter(r => r.ACTUALSQTY !== '');
  assert.equal(history.length, 1);
  assert.equal(history[0].ACTUALSQTY, '123');
  assert.equal(JSON.parse(zip.readAsText('VALIDATION_REPORT.json')).actualHistorySource, 'uploaded-excel');
  f.analysis.actuals[0].KEYFIGUREDATE = '2024-02-30';
  assert.throws(() => output(f), /invalid KEYFIGUREDATE/);
});

test('BOM-less uploaded products do not get fabricated circular or synthetic components', () => {
  const f = fixture(true); f.analysis.bom = [];
  assert.equal(csv(output(f), '170_PRODUCTIONSOURCEITM_NEW.CSV').rows.length, 0);
});

test('1020 self and multi-level BOM cycles are blocked', () => {
  const f = fixture(); f.analysis.bom[0].component = 'FG';
  assert.throws(() => output(f), /1020 circular/);
  const g = fixture(); g.analysis.productFamilies[0].semi.push({ id: 'RAW', productionLocationIds: ['LOC-B'] });
  g.analysis.bom.push({ parent: 'RAW', component: 'FG', plant: 'B', qty: 1 });
  assert.throws(() => output(g), /1020 circular/);
});

test('multi-plant Excel ambiguity, one-character fields and missing selections are blocked', () => {
  const f = fixture(true); f.analysis.bom = []; delete f.analysis.productFamilies[0].fg[0].productionLocationIds;
  assert.throws(() => output(f), /multiple plants/);
  const g = fixture(); g.analysis.productFamilies[0].fg[0].stockingNodeType = 'DC';
  assert.throws(() => output(g), /one-character/);
  g.selections.selectedFamilyIds = [];
  assert.throws(() => output(g), /Select products/);
});

test('production Resource Type 1 is blocked, explicit verified category enables only confirmed links', () => {
  const f = fixture(); f.selections.generationOptions = { confirmedProductionResourceIds: ['LOC-B ASSEMBLY'], productionResourceTypeCode: '1' };
  assert.throws(() => output(f), /1080/);
  f.selections.generationOptions.productionResourceTypeCode = 'VERIFIED_PRODUCTION';
  const zip = output(f);
  assert.equal(csv(zip, '180_PRODUCTIONSOURCERESOURCE_NEW.CSV').rows.length, 1);
  assert.equal(csv(zip, '130_RESOURCELOCATION_NEW.CSV').rows.find(r => r.RESID === 'LOC-B ASSEMBLY').RESOURCETYPE, 'VERIFIED_PRODUCTION');
  f.analysis.resources = [{ resourceId: 'BAD', resourceType: '1' }];
  assert.throws(() => output(f), /1080/);
});

test('capacity aliases are not inferred; unknown KF/efficiency out of range are blocked', () => {
  const f = fixture(); f.selections.generationOptions = { capacityProfiles: { 'LOC-B ASSEMBLY': { ZSRESOURCEWORKDAYS: 5 } } };
  assert.throws(() => output(f), /unsupported capacity input/);
  f.selections.generationOptions.capacityProfiles['LOC-B ASSEMBLY'] = { ZSEFFICIENCY: 85 };
  assert.throws(() => output(f), /invalid ZSEFFICIENCY/);
});

test('offline technical Excel import preserves exported IDs/types/regions/UOM and mapping across customers', async () => {
  const wb = XLSX.utils.book_new();
  const add = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  add('Location', [['SCNID', 'LOCID', 'LOCDESCR', 'LOCTYPE', 'LOCREGION', 'LOCCOUNTRY', 'GEOLATITUDE', 'GEOLONGITUDE'],
    ['Version', 'Location ID*', 'Location Desc*', 'Type', 'Region', 'Country', 'Lat', 'Lon'],
    ['Base Version', 'LOC-DE', 'Berlin, Plant', 'P', 'DE-BE', 'DE', 0, 13.4]]);
  add('Product', [['PRDID', 'PRDDESCR', 'UOMID', 'PRDFAMILY', 'MATTYPEID'], ['SKU', 'Panel', 'M2', 'Panels', 'FERT']]);
  add('Customer', [['CUSTID', 'CUSTDESCR', 'CUSTREGION', 'CUSTCOUNTRY', 'SALESAREA'], ['BUYER', 'Buyer', 'DE-BE', 'DE', '3030/10/00']]);
  add('Location Product', [['LOCID', 'PRDID', 'PLUNITID'], ['LOC-DE', 'SKU', 'DE_PLAN']]);
  add('Actuals', [['PRDID', 'LOCID', 'CUSTID', 'KEYFIGUREDATE', 'ACTUALSQTY', 'SALESAREA'], ['SKU', 'LOC-DE', 'BUYER', '2024-11-25', 77, '3030/10/00']]);
  const result = await parseExcelToIbpData(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), 'Other.xlsx', { companyCountry: 'DE', companyName: 'Other' });
  assert.equal(result.locations[0].id, 'LOC-DE');
  assert.equal(result.locations[0].lat, 0);
  assert.equal(result.customers[0].region, 'DE-BE');
  assert.equal(result.productFamilies[0].fg[0].uomId, 'M2');
  assert.equal(result.generationOptions.planningUnitId, 'DE_PLAN');
  assert.equal(result.actuals[0].ACTUALSQTY, 77);
  const selections = { customers: result.customers, selectedFamilyIds: result.productFamilies.map(f => f.id), selectedLocationIds: ['LOC-DE'] };
  const ctx = prepareGeneration(result, selections);
  assert.ok(ctx.production.has('SKU|LOC-DE'));
  assert.equal(csv(new AdmZip(generateZip(result, selections)), '205_PLANNINGUNIT_NEW.CSV').rows[0].PLUNITID, 'DE_PLAN');
});

test('technical Source Header and BOM sheets resolve parent/component relationships for any customer', async () => {
  const wb = XLSX.utils.book_new();
  const add = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  add('Location', [['LOCID', 'LOCDESCR', 'LOCTYPE'], ['LOC-A', 'Factory', 'P']]);
  add('Product', [['PRDID', 'PRDDESCR', 'MATTYPEID', 'PRDFAMILY'], ['FG', 'Panel', 'FERT', 'Panels'], ['RAW', 'Resin', 'ROH', 'Panels']]);
  add('Customer', [['CUSTID', 'CUSTDESCR'], ['C1', 'Customer']]);
  add('Production Source Header', [['LOCID', 'PRDID', 'SOURCEID', 'OUTPUTCOEFFICIENT'], ['LOC-A', 'FG', 'S1', 1]]);
  add('Production Source Item', [['PRDID', 'SOURCEID', 'LOCID', 'COMPONENTCOEFFICIENT', 'SOURCEITMID'], ['RAW', 'S1', 'LOC-A', 2.5, '10']]);
  const result = await parseExcelToIbpData(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), 'Customer B.xlsx', { companyName: 'Customer B', companyCountry: 'KR' });
  assert.equal(result.bom[0].parent, 'FG');
  assert.equal(result.bom[0].component, 'RAW');
  const zip = new AdmZip(generateZip(result, { customers: result.customers, selectedFamilyIds: ['F001'], selectedLocationIds: ['LOC-A'] }));
  assert.equal(csv(zip, '170_PRODUCTIONSOURCEITM_NEW.CSV').rows[0].COMPONENTCOEFFICIENT, '2.5');
});
