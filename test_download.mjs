import fs from 'fs';

// 1. Import Excel
const buf = fs.readFileSync(String.raw`C:\Users\I544201\Downloads\(LX하우시스) PoC 기준정보 요청서(IBP)_v1.0의 복사본.xlsx`);
const fileData = buf.toString('base64');

const imp = await fetch('http://localhost:3001/api/import-excel', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ fileData, fileName: 'LX하우시스.xlsx' }),
});
const data = await imp.json();
console.log('Import OK — bom:', data.bom?.length, 'routing:', data.routing?.length, 'lt:', JSON.stringify(data.leadTimes));

// 2. Download ZIP
const families = data.productFamilies.map(f => f.id);
const locs = data.locations.map(l => l.id);
const dlRes = await fetch('http://localhost:3001/api/download', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    analysis: data,
    selections: { customers: data.customers, selectedFamilyIds: families, selectedLocationIds: locs },
  }),
});
if (!dlRes.ok) { console.error('Download failed:', await dlRes.text()); process.exit(1); }

const zipBuf = Buffer.from(await dlRes.arrayBuffer());
fs.writeFileSync('C:/tmp/test_output.zip', zipBuf);
console.log(`ZIP saved: ${Math.round(zipBuf.length / 1024)} KB`);

// 3. Check key CSVs inside ZIP
import AdmZip from 'adm-zip';
const zip = new AdmZip('C:/tmp/test_output.zip');
const check = (name) => {
  const e = zip.getEntries().find(e => e.entryName.includes(name));
  if (!e) { console.log(`  [MISSING] ${name}`); return; }
  const lines = e.getData().toString('utf8').split('\n');
  console.log(`\n--- ${name} (${lines.length-1} data rows) ---`);
  lines.slice(0, 4).forEach(l => console.log(' ', l.slice(0, 100)));
};

check('130_RESOURCE');
check('150_SOURCE');
check('160_SOURCEPROD');
check('170_PRODUCTION');
check('180_PRODUCTIONRES');
