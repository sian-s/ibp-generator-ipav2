import fs from 'fs';

const filePath = String.raw`C:\Users\I544201\Downloads\(LX하우시스) PoC 기준정보 요청서(IBP)_v1.0의 복사본.xlsx`;
const buf = fs.readFileSync(filePath);
const fileData = buf.toString('base64');

console.log(`File: ${Math.round(fileData.length * 0.75 / 1024)} KB — calling /api/import-excel ...`);

const res = await fetch('http://localhost:3001/api/import-excel', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ fileData, fileName: '(LX하우시스) PoC 기준정보 요청서(IBP)_v1.0의 복사본.xlsx' }),
});

if (!res.ok) { console.error(await res.json()); process.exit(1); }
const data = await res.json();

console.log('\n=== Company ===');
console.log(data.company?.name, '/', data.company?.hq?.country);

console.log('\n=== Locations ===', data.locations?.length);
data.locations?.forEach(l => console.log(' ', l.id, l.desc));

console.log('\n=== Customers ===', data.customers?.length);
console.log(' ', data.customers?.slice(0,3).map(c => `${c.custId} ${c.custdescr}`).join(', '));

console.log('\n=== Product Families ===', data.productFamilies?.length);
data.productFamilies?.forEach(f => console.log(' ', f.id, f.label, `(${f.fg?.length} FG)`));

console.log('\n=== BOM ===', data.bom?.length, 'entries');
data.bom?.slice(0, 3).forEach(b => console.log(`  Plant:${b.plant} Parent:${b.parent} Comp:${b.component} Qty:${b.qty} ${b.uom}`));

console.log('\n=== Resources ===', data.resources?.length, 'entries');
data.resources?.slice(0, 5).forEach(r => console.log(`  ${r.locationId} / ${r.resourceId} cap:${r.dailyCap}`));

console.log('\n=== Routing ===', data.routing?.length, 'entries');
data.routing?.slice(0, 3).forEach(r => console.log(`  ${r.product} @ ${r.locationId} → ${r.workCenter}`));

console.log('\n=== Lead Times ===');
console.log(' ', JSON.stringify(data.leadTimes));
