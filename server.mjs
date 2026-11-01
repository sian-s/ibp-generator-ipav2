import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import Anthropic from '@anthropic-ai/sdk';
import AdmZip from 'adm-zip';
import XLSX from 'xlsx';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = process.env.PORT || 3001;

// Load .env if present
try {
  const env = fs.readFileSync(path.join(__dirname, '.env'), 'utf8');
  for (const line of env.split('\n')) {
    const m = line.match(/^([^#=\s][^=]*)=(.+)$/);
    if (m) process.env[m[1].trim()] = m[2].trim().replace(/^["']|["']$/g, '');
  }
} catch {}

const client = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY,
  ...(process.env.ANTHROPIC_BASE_URL && { baseURL: process.env.ANTHROPIC_BASE_URL }),
});

const BP_ZIP_PATH = process.env.BP_ZIP_PATH ||
  path.join(process.env.USERPROFILE || process.env.HOME || '',
    'Downloads', 'IBP2608_Technical_Assets_SAPIBP1 (1).zip');

// ─── BP ZIP Reference Data ────────────────────────────────────────────────────

let bpData = {};   // filename → raw CSV string
let bpSeeds = {};  // extracted seed values

function parseCsvRows(text) {
  const lines = text.trim().split('\n').filter(Boolean);
  if (!lines.length) return { header: '', rows: [] };
  return { header: lines[0], rows: lines.slice(1) };
}

function filterByProduct(csvText, productIds) {
  const { header, rows } = parseCsvRows(csvText);
  const filtered = rows.filter(r => productIds.some(id => r.startsWith(id + ';')));
  return filtered.length ? header + '\n' + filtered.join('\n') : header;
}

function loadBpZip() {
  try {
    const zip = new AdmZip(BP_ZIP_PATH);
    for (const e of zip.getEntries()) {
      if (!e.isDirectory) {
        const base = path.basename(e.entryName);
        bpData[base] = e.getData().toString('utf8');
      }
    }
    console.log(`[bp] Loaded ${Object.keys(bpData).length} files from ZIP`);
    extractSeeds();
  } catch (err) {
    console.warn(`[bp] Could not load ZIP: ${err.message}`);
    console.warn(`[bp] Expected at: ${BP_ZIP_PATH}`);
  }
}

function extractSeeds() {
  // Monthly AOP values for IBP-100 (12 months of actuals)
  const mth = parseCsvRows(bpData['MTHPRODCUST_2608.csv'] || '');
  const mthIbp = mth.rows.filter(r => r.startsWith('IBP-100;'));
  const mthAop = mthIbp.slice(0, 12).map(r => parseInt(r.split(';')[3]) || 1050);
  while (mthAop.length < 12) mthAop.push(1050);

  // Monthly revenue/profit for IBP-100
  const mthCurr = parseCsvRows(bpData['MTHPRODCUSTCURR_2608.csv'] || '');
  const mthCurrIbp = mthCurr.rows.filter(r => r.startsWith('IBP-100;'));
  const mthRev = mthCurrIbp.slice(0, 12).map(r => parseInt(r.split(';')[4]) || 55000);
  const mthPft = mthCurrIbp.slice(0, 12).map(r => parseInt(r.split(';')[5]) || 8000);
  while (mthRev.length < 12) mthRev.push(55000);
  while (mthPft.length < 12) mthPft.push(8000);

  // Weekly WKPRODLOC cost parameters (IBP-100)
  const wkLoc = parseCsvRows(bpData['WKPRODLOC_2608.csv'] || '');
  const wkLocRow = wkLoc.rows.find(r => r.startsWith('IBP-100;')) || '';
  const wkLocParts = wkLocRow.split(';');

  // Weekly capacity (WKRESLOC)
  const wkRes = parseCsvRows(bpData['WKRESLOC_2608.csv'] || '');
  const wkResRow = wkRes.rows.find(r => r.startsWith('1010;')) || '';
  const wkResParts = wkResRow.split(';');

  // WKPRODLOCCUSTGROUP IO forecast (IBP-100)
  const wkGrp = parseCsvRows(bpData['WKPRODLOCCUSTGROUP_2608.csv'] || '');
  const wkGrpIbp = wkGrp.rows.filter(r => r.startsWith('IBP-100;'));
  const ioForecastVals = wkGrpIbp.slice(0, 8).map(r => parseInt(r.split(';')[4]) || 200);

  // WKPRODLOCCUST actuals
  const wkLcc = parseCsvRows(bpData['WKPRODLOCCUST_2608.csv'] || '');
  const wkLccIbp = wkLcc.rows.filter(r => r.startsWith('IBP-100;'));
  const wkActuals = wkLccIbp.slice(0, 8).map(r => parseInt(r.split(';')[4]) || 0);
  const wkSalesFcast = wkLccIbp.slice(0, 8).map(r => parseInt(r.split(';')[7]) || 242);
  // fallback: WKPRODLOCCUST_2608.csv not in BP ZIP → derive weekly actuals from monthly AOP
  if (wkActuals.every(v => v === 0)) {
    const weeklyBase = Math.round((mthAop[0] || 1081) / 4.3);
    for (let i = 0; i < 8; i++) wkActuals[i] = weeklyBase;
  }

  // WKPRODLOCSRC production costs
  const wkSrc = parseCsvRows(bpData['WKPRODLOCSRC_2608.csv'] || '');
  const wkSrcRow = wkSrc.rows.find(r => r.startsWith('IBP-100;')) || '';
  const wkSrcParts = wkSrcRow.split(';');

  // WKPRODLOCCUSTCURR revenue
  const wkCustCurr = parseCsvRows(bpData['WKPRODLOCCUSTCURR_2608.csv'] || '');
  const wkCustCurrIbp = wkCustCurr.rows.filter(r => r.startsWith('IBP-100;'));
  const wkActualsRev = wkCustCurrIbp.slice(0, 8).map(r => parseInt(r.split(';')[5]) || 14000);
  const wkOpenOrdersRev = wkCustCurrIbp.slice(0, 8).map(r => parseInt(r.split(';')[6]) || 8000);

  bpSeeds = {
    mthAop,
    mthRev,
    mthPft,
    wkInitialInventory:                    parseInt(wkLocParts[3]) || 100,
    wkInventoryHoldingCostRate:            parseFloat(wkLocParts[4]) || 1,
    wkInventoryTargetViolationCostRate:    parseFloat(wkLocParts[6]) || 1,
    wkNagCstInitialAggrUnitCost:           parseFloat(wkLocParts[9]) || 8,
    wkNagCstInventoryCostRate:             parseFloat(wkLocParts[10]) || 1,
    wkNagCo2InitialAggrUnitEmis:           parseFloat(wkLocParts[13]) || 13647,
    wkNagCo2InventoryEmisRate:             parseFloat(wkLocParts[14]) || 1,
    wkCapaSupply:                          parseInt(wkResParts[3]) || 3000,
    ioForecastVals,
    wkActuals,
    wkSalesFcast,
    wkActualsRev,
    wkOpenOrdersRev,
    wkProductionCostRate:    parseFloat(wkSrcParts[4]) || 1,
    wkNagCstProdCostRate:    parseFloat(wkSrcParts[5]) || 1,
    wkNagCstFixedProdCost:   parseFloat(wkSrcParts[6]) || 1,
    wkNagCo2ProdEmisRate:    parseFloat(wkSrcParts[7]) || 5000,
    wkNagCo2FixedProdEmis:   parseFloat(wkSrcParts[8]) || 10000,
    plannedPrice: 400,
    costPerUnit: 100,
  };

  console.log(`[bp] Seeds extracted — mthAop[0]=${bpSeeds.mthAop[0]}, wkInitInv=${bpSeeds.wkInitialInventory}`);
}

loadBpZip();

// ─── Date helpers ──────────────────────────────────────────────────────────────

function monthDates(count) {
  const now = new Date();
  const yr = now.getFullYear(), mo = now.getMonth();
  return Array.from({ length: count }, (_, i) => {
    const m = mo + i;
    return `${yr + Math.floor(m / 12)}-${String((m % 12) + 1).padStart(2, '0')}-01`;
  });
}

function pastMonthDates(count) {
  const now = new Date();
  const yr = now.getFullYear(), mo = now.getMonth();
  return Array.from({ length: count }, (_, i) => {
    const m = mo - count + i;
    const absM = ((m % 12) + 12) % 12;
    const y = yr + Math.floor(m / 12);
    return `${y}-${String(absM + 1).padStart(2, '0')}-01`;
  });
}

function weekDates(count) {
  const d = new Date();
  d.setDate(d.getDate() + (1 + 7 - d.getDay()) % 7);
  return Array.from({ length: count }, (_, i) => {
    const w = new Date(d); w.setDate(d.getDate() + i * 7);
    return w.toISOString().slice(0, 10);
  });
}

function pastWeekDates(count) {
  const d = new Date();
  d.setDate(d.getDate() - count * 7);
  return Array.from({ length: count }, (_, i) => {
    const w = new Date(d); w.setDate(d.getDate() + i * 7);
    return w.toISOString().slice(0, 10);
  });
}

function seedVal(arr, i) { return arr[i % arr.length] || arr[0] || ''; }
function jitter(base, pct = 0.15) {
  return Math.round(base * (1 + (Math.random() * 2 - 1) * pct));
}

// ─── CSV generation helpers ───────────────────────────────────────────────────

const csv = (header, rows) => rows.length ? 'sep=,\n' + header + '\n' + rows.join('\n') : null;

function sourceId(prdId, locId) {
  return `${prdId.padEnd(18)}${locId.padEnd(18)}S`;
}

// ─── File generators ───────────────────────────────────────────────────────────

function gen010Customer(customers) {
  // CP1CUSTOMER_TEMPLATE (1).csv — 39 columns
  const h = 'CUSTID,COSTCENTER,COSTCENTERDESCR,COSTCENTEROWNER,CPCUSTATT01,CPCUSTATT02,CPCUSTATT03,CUSTBUPAID,CUSTCHANNEL,CUSTCOUNTRY,CUSTDESCR,CUSTGROUP,CUSTGROUPDESC,CUSTREGION,CUSTREGIONDESCR,CUSTSALESREP,CUSTSALESVP,CUSTVALID,EQUIPNR,FLOCDESCR,FLOCID,FUNCTHIER1,FUNCTHIER2,GEOLATITUDE,GEOLONGITUDE,KEYACCOUNT,PLANT,ZXCUSTPRIORITY,ZXKAM,ZXSALESTER,ZXSEGCUSTCHANABCID,ZXSEGCUSTCHANABCLOCKED,ZXSEGCUSTCHANXYZID,ZXSEGCUSTCHANXYZLOCKED,ZXSEGCUSTOMERABCID,ZXSEGCUSTOMERABCLOCKED,ZXSEGCUSTOMERXYZID,ZXSEGCUSTOMERXYZLOCKED,ZXTSM';
  return csv(h, customers.map(c => [
    c.custId, `CC-${c.custId}`, `Cost Center ${c.custId}`, '', '', '', '', '',
    c.channel, c.country, c.custdescr, c.group, '', '', '',
    '', '', '', '', '', '', '', '',
    c.lat || '', c.lon || '',
    '', '', '', '', '', '', '', '', '', '', '', '', '', '',
  ].join(',')));
}

function gen020Location(locs) {
  const h = 'LOCID,CPLOCATT01,CPLOCATT02,CPLOCATT03,GEOLATITUDE,GEOLONGITUDE,HLOCATIONCOUNTER,HOLDINGCOSTPCT,LOCBUPAID,LOCCITY,LOCCOUNTRY,LOCDESCR,LOCREGION,LOCTYPE,LOCVALID';
  return csv(h, locs.map(l =>
    `${l.id},,,,${l.lat||''},${l.lon||''},,,,${l.city||''},${l.country||''},${l.desc},${l.region},${l.locType||'P'},`
  ));
}

function gen030Product(families) {
  // CP1PRODUCT_TEMPLATE.csv — 48 columns
  const h = 'PRDID,ABCID,ABCLOCKED,AGGPRDATT01,AGGPRDATT02,BRAND,CATEGORY,CPBLOCKID,CPPRDATT01,CPPRDATT02,CPPRDATT03,EOLDATE,HPRODUCTCOUNTER,ITEMTYPECATEGORY,MATTYPEDESCR,MATTYPEID,PLANNEDACTIVITYDESCR,PLANNEDACTIVITYID,PLMSTATUS,PRDBU,PRDDESCR,PRDFAMILY,PRDFAMILYDESCR,PRDGROUP,PRDSERIES,PRDSUBFAMILY,PRDVALID,QUICKTURNMARKET,RELEASEDATE,REPLACEDPRD,UOMDESCR,UOMID,XYZID,XYZLOCKED,ZXCLASS,ZXGRADE,ZXPRODFORMULAFLAG,ZXPRODTYPE,ZXSEGPRODBRANDABCID,ZXSEGPRODBRANDABCLOCKED,ZXSEGPRODBRANDXYZID,ZXSEGPRODBRANDXYZLOCKED,ZXSEGPRODUCTABCID,ZXSEGPRODUCTABCLOCKED,ZXSEGPRODUCTXYZID,ZXSEGPRODUCTXYZLOCKED,ZXTSPROPPRD,ZXVALUESTREAM';
  const rows = [];
  for (const f of families) {
    const subfam = f.code.replace('FAMILY ', 'SUB-FAMILY ');
    const allPrds = [
      ...(f.fg   || []).map(p => ({ ...p, type: 'FERT', abc: p.abc || '', xlk: p.abc === 'A' ? 'X' : '', xyz: 'Z' })),
      ...(f.semi || []).map(p => ({ ...p, type: 'HALB', abc: '', xlk: '', xyz: '' })),
      ...(f.raw  || []).map(p => ({ ...p, type: 'ROH',  abc: '', xlk: '', xyz: '' })),
    ];
    for (const p of allPrds) {
      rows.push([
        p.id, p.abc, p.xlk,
        '', '', f.brand, '', '', '', '', '', '', '', '',
        '', p.type, '', '', 'EXISTING', '',
        p.desc || p.id, f.code, '', f.group, '', subfam, '', '', '', '',
        'PIECE', 'PC', p.xyz, '',
        '', '', '', '', '', '', '', '', '', '', '', '', '', '',
      ].join(','));
    }
  }
  return csv(h, rows);
}

function gen040Resource(locs, routing) {
  const h = 'RESID,RESDESCR';
  const rows = [];
  for (const l of locs) {
    rows.push(
      `${l.id} ASSEMBLY,${l.desc} Assembly`,
      `${l.id} PACK,${l.desc} Pack`,
      `${l.id} SUBASSEMBLY,${l.desc} Subassembly`,
    );
  }
  return csv(h, rows);
}

function gen100UomConversion(families) {
  // CP1UOMCONVERSIONFACTOR_TEMPLATE.csv — 3 columns
  const h = 'PRDID,UOMTOID,UOMCONVERSIONFACTOR';
  const rows = [];
  for (const f of families) {
    for (const p of [...(f.fg || []), ...(f.semi || []), ...(f.raw || [])]) {
      rows.push(`${p.id},EA,1`, `${p.id},PC,1`);
    }
  }
  return csv(h, rows);
}

function gen110CustomerProduct(customers, families) {
  // CP1CUSTOMERPRODUCT_TEMPLATE (1).csv — 16 columns
  const h = 'CUSTID,PRDID,HPRODCUSTCOUNTER,LATEDELIVERYCOSTPOLICY,MARKETFAMILY,MARKETSEGMENT,MINRQDSHLFLIFE,MINSHELFLIFECONSDEM,NONDELIVERYCOSTPOLICY,ZXSEGATTKFFORABCINPUT,ZXSEGCUSTPRODABCID,ZXSEGCUSTPRODABCLOCKED,ZXSEGCUSTPRODXYZID,ZXSEGCUSTPRODXYZLOCKED,ZXTSPROPPRDCUST,ZZZBSTFITFRCSTMOD';
  const rows = [];
  for (const c of customers) {
    const fams = c.recommendedFamilies?.length ? families.filter(f => c.recommendedFamilies.includes(f.id)) : families;
    for (const f of fams) {
      const famNum = (f.id.replace('F', '').padStart(3, '0'));
      const seg = (c.custId.slice(2, 4) + famNum).slice(0, 5);
      for (const p of (f.fg || [])) {
        rows.push([c.custId, p.id, '', '', '', seg, '', '', '', '', '', '', '', '', '', ''].join(','));
      }
    }
  }
  return csv(h, rows);
}

function gen120LocationProduct(locs, families) {
  const h = 'PRDID,LOCID,PLANNER,PLUNITID,FOOTPRINTPERUNIT,STOCKINGNODETYPE';
  const rows = [];
  for (const l of locs) {
    for (const f of families) {
      for (const p of [...(f.fg || []), ...(f.semi || [])]) {
        const fp = f.fg?.find(x => x.id === p.id) ? '100' : '';
        rows.push(`${p.id},${l.id},1,IBP1,${fp},`);
      }
    }
  }
  return csv(h, rows);
}

function gen130ResourceLocation(locs, resources, routing) {
  // CP1RESOURCELOCATION_TEMPLATE.csv — 9 columns
  const h = 'LOCID,RESID,CAPASUPPLY,CAPASUPPLYEXPANSION,CAPASUPPLYEXPANSIONTS,CAPASUPPLYTS,CONSTRAINTTYPE,RESOURCETYPE,ZXCO2ACTIVITY';
  const rows = [];
  for (const l of locs) {
    rows.push(
      [l.id, `${l.id} ASSEMBLY`,    '', '', '', '', 'F', '', ''].join(','),
      [l.id, `${l.id} PACK`,        '', '', '', '', 'F', '', ''].join(','),
      [l.id, `${l.id} SUBASSEMBLY`, '', '', '', '', 'F', '', ''].join(','),
    );
  }
  return csv(h, rows);
}

function gen135ResourceLocationProduct(locs, families) {
  // IPAV2 template CP1RESOURCELOCATIONPRODUCT: product-level resource capacity consumption
  const h = 'LOCID,PRDID,RESID,CAPACONSUMPTION,CONSUMPTIONTS';
  const rows = [];
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const workLocs = plants.length ? plants : locs;
  for (const l of workLocs) {
    for (const f of families) {
      for (const p of [...(f.fg || []), ...(f.semi || [])]) {
        rows.push(`${l.id},${p.id},${l.id} ASSEMBLY,1,`);
      }
    }
  }
  return csv(h, rows);
}

function gen140SourceCustomer(customers, locs, families) {
  // CP1SOURCECUSTOMER_TEMPLATE.csv — 11 columns
  const h = 'CUSTID,LOCID,PRDID,CBALANCERECEIPTSCOPE,CINVALID,CRATIO,CROUNDING,FCSTCONSMODE,LEADTIME,MATTRSOURCECUSTOMER,RATIOTS';
  const rows = [];
  for (const c of customers) {
    const nearLocs = locs.filter(l => l.region === c.region).slice(0, 2);
    const servLocs = nearLocs.length ? nearLocs : locs.slice(0, 1);
    const fams = c.recommendedFamilies?.length ? families.filter(f => c.recommendedFamilies.includes(f.id)) : families;
    for (const f of fams) {
      for (const p of (f.fg || [])) {
        const ratio = servLocs.length > 1 ? 0.5 : 1;
        for (const l of servLocs) rows.push(`${c.custId},${l.id},${p.id},,,${ratio},,CONTINUOUS,,,`);
      }
    }
  }
  return csv(h, rows);
}

function gen150SourceLocation(locs, families, leadTimes) {
  // CP1SOURCELOCATION_TEMPLATE.csv — 23 columns
  const h = 'LOCFR,LOCID,MOTID,PRDID,FOOTPRINTSUPPLIERPERUNIT,LEADTIME,MATTRSOURCELOCATION,MAXLOCATIONQUOTA,MINLOCATIONQUOTA,RATIOTS,TBALANCERECEIPTSCOPE,TCAPACONSPOLICY,TDELIVERYTYPE,TFREEZEHORIZON,TINCLOTSIZE,TINVALID,TLEADTIME,TLEADTIMEVARIABILITY,TMAXLOTSIZE,TMINLOTSIZE,TRATIO,TROUNDING,TSHIPMENTFREQUENCY';
  const rows = [];
  const distLt = Math.round(Number(leadTimes?.distribution) || 2);
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const dcs    = locs.filter(l => l.locType !== 'P' && l.locType !== 'Plant');
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const dc of (dcs.length ? dcs : locs)) {
        const sourcePlants = plants.filter(pl => pl.id !== dc.id && pl.region === dc.region);
        const targets = sourcePlants.length ? sourcePlants : plants.filter(pl => pl.id !== dc.id).slice(0, 1);
        if (!targets.length) continue;
        const ratio = targets.length > 1 ? 0.5 : 1;
        for (const pl of targets) {
          rows.push([pl.id, dc.id, '', p.id, '', distLt, '', '', '', '', '', '', '', '', '', '', distLt, '', '', '', ratio, '', ''].join(','));
        }
      }
    }
  }
  return csv(h, rows);
}

function gen160SourceProductionHeader(locs, families, routing, leadTimes) {
  // CP1SOURCEPRODUCTION_TEMPLATE.csv — 25 columns
  const h = 'LOCID,PRDID,SOURCEID,CPPRODUCTIONLINE,MATTRSOURCEPRODUCTION,OUTPUTCOEFFICIENT,OUTPUTCOEFFICIENTTS,PBALANCERECEIPTSCOPE,PCAPACONSPOLICY,PDELIVERYTYPE,PFREEZEHORIZON,PINCLOTSIZE,PINVALID,PLEADTIME,PLEADTIMEVARIABILITY,PMAXLOTSIZE,PMINLOTSIZE,PPERIODSOFCOVERAGE,PRATIO,PRATIOTS,PREMSHELFLIFE,PRODUCTIONCOSTRATE,PROUNDING,PSHIPMENTFREQUENCY,SOURCETYPE';
  const rows = [];
  const prodLt = Math.round(Number(leadTimes?.production) || 1);
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const workLocs = plants.length ? plants : locs;

  // Build workCenter lookup from routing: product → workCenter (first match)
  const wcByProduct = new Map();
  if (routing && routing.length) {
    for (const r of routing) {
      if (r.product && !wcByProduct.has(r.product)) wcByProduct.set(r.product, r.workCenter || '');
    }
  }

  // families 기준으로 모든 FG/SEMI 제품 행 생성, routing은 workCenter 조회에만 사용
  for (const l of workLocs) {
    for (const f of families) {
      for (const p of [...(f.fg || []), ...(f.semi || [])]) {
        if (!p.id) continue;
        const wc = wcByProduct.get(p.id) || '';
        const sid = sourceId(p.id, l.id);
        rows.push([l.id, p.id, sid, wc, '', '100', '', '', '', '', '', '', '', prodLt, '', '', '', '', '1', '', '', '', '', '', 'P'].join(','));
      }
    }
  }
  return csv(h, rows);
}

function gen170ProductionSourceItm(locs, families, bom) {
  // CP1PRODUCTIONSOURCEITM_TEMPLATE.csv — 8 columns
  const h = 'PRDID,SOURCEID,COMPONENTCOEFFICIENT,COMPONENTCOEFFICIENTTS,COMPONENTOFFSET,LOCID,MINSHELFLIFECOMP,SOURCEITMID';
  const rows = [];
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const workLocs = plants.length ? plants : locs;

  // Build set of all known product IDs from families (FG + SEMI + RAW)
  const knownProducts = new Set();
  for (const f of families) {
    for (const p of [...(f.fg || []), ...(f.semi || []), ...(f.raw || [])]) {
      if (p.id) knownProducts.add(p.id);
    }
  }

  // Build set of FG+SEMI IDs (these have SOURCEID entries from gen160)
  const sourceProducts = new Set();
  for (const f of families) {
    for (const p of [...(f.fg || []), ...(f.semi || [])]) {
      if (p.id) sourceProducts.add(p.id);
    }
  }

  // Try BOM data — only rows where parent is a source product AND component is a known product
  let usedBom = false;
  if (bom && bom.length) {
    const locByPlant = {};
    for (const l of workLocs) locByPlant[l.id.replace('LOC-', '')] = l;
    const seen = new Set();
    for (const b of bom) {
      if (!sourceProducts.has(b.parent)) continue;
      if (!knownProducts.has(b.component)) continue;
      const loc = locByPlant[b.plant] || workLocs[0];
      if (!loc) continue;
      const sid = sourceId(b.parent, loc.id);
      const key = `${b.component}|${sid}`;
      if (seen.has(key)) continue;
      seen.add(key);
      rows.push([b.component, sid, b.qty, '', '0', loc.id, '', b.seq || b.component].join(','));
      usedBom = true;
    }
  }

  // Fallback to families-based logic when BOM had no matching entries
  if (!usedBom) {
    let seq = 10;
    for (const l of workLocs) {
      for (const f of families) {
        const fgs   = f.fg   || [];
        const semis = f.semi || [];
        const raws  = f.raw  || [];
        for (const fg of fgs) {
          const sid = sourceId(fg.id, l.id);
          const comp = semis[0] || raws[0] || fg; // FG이라도 유효한 컴포넌트로 사용
          rows.push([comp.id, sid, '1', '', '0', l.id, '', String(seq).padStart(4, '0')].join(','));
          seq += 10;
        }
        for (const semi of semis) {
          const sid = sourceId(semi.id, l.id);
          const raw = raws[1] || raws[0] || semi;
          rows.push([raw.id, sid, '1', '', '0', l.id, '', String(seq).padStart(4, '0')].join(','));
          seq += 10;
        }
      }
    }
  }
  return csv(h, rows);
}

function gen180ProductionSourceResource(locs, families, routing, resources) {
  // CP1PRODUCTIONRESOURCE_TEMPLATE.csv — 9 columns
  const h = 'RESID,SOURCEID,CAPACONSUMPTION,CONSUMPTIONTS,FIXEDPCAPACONSUMPTION,ISALTRES,LOCID,PCAPACONSEND,PCAPACONSSTART';
  const rows = [];
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const workLocs = plants.length ? plants : locs;

  // families 기준으로 FG/SEMI 순회, RESID = LOC-XXXX ASSEMBLY (gen040/gen130과 일치)
  for (const l of workLocs) {
    for (const f of families) {
      for (const p of [...(f.fg || []), ...(f.semi || [])]) {
        if (!p.id) continue;
        const sid = sourceId(p.id, l.id);
        rows.push([`${l.id} ASSEMBLY`, sid, '1', '', '', '', l.id, '', ''].join(','));
      }
    }
  }
  return csv(h, rows);
}

function gen190SourceCustGroup(customers, locs, families) {
  // CP1SOURCECUSTGROUP_TEMPLATE.csv — 6 columns
  const h = 'CUSTGROUP,LOCID,PRDID,CUSTOMERTOLERANCETIME,SALESORDERVISIBILITYHORIZON,TARGETSERVICELEVEL';
  const rows = [];
  const groups = [...new Set(customers.map(c => c.group))];
  for (const l of locs) {
    for (const f of families) {
      for (const p of (f.fg || [])) {
        for (const g of groups) rows.push([g, l.id, p.id, '', '', ''].join(','));
      }
    }
  }
  return csv(h, rows);
}

// DISCTCHANNEL max 2 chars; SALESAREA = SALESORG/CHAN/DIVISION
const CHAN_CODE = { Direct:'10', Indirect:'11', Online:'12', Export:'13', Default:'10' };
function salesArea(channel) {
  const ch = CHAN_CODE[channel] || '10';
  return `1010/${ch}/00`;
}

function gen200SalesArea(customers) {
  // CP1SA_TEMPLATE.csv — 12 columns; DISCTCHANNEL max 2 chars
  const h = 'SALESAREA,CPCOMPANYCODE,CPCOMPANYCODEDESCR,CPSAATT01,CPSAATT02,CPSAATT03,DISCTCHANNEL,DISTCHANNELDESCR,DIVISION,DIVISIONDESC,SALESORG,SALESORGDESCR';
  const channels = [...new Set(customers.map(c => c.channel || 'Direct'))];
  const rows = channels.map(ch => {
    const code = CHAN_CODE[ch] || '10';
    const sa   = `1010/${code}/00`;
    return [sa, '1010', '', '', '', '', code, `${ch} Sales`, '00', 'Product Division', '1010', 'Sales Org'].join(',');
  });
  return csv(h, rows);
}

// ── KF generators ─────────────────────────────────────────────────────────────

function genMthProdCust(customers, families) {
  const h = 'PRDID,CUSTID,KEYFIGUREDATE,AOPQTY';
  const dates = monthDates(12);
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        for (let i = 0; i < dates.length; i++) {
          rows.push(`${p.id},${c.custId},${dates[i]},${jitter(seedVal(bpSeeds.mthAop, i))}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genMthProdCustCurr(customers, families) {
  const h = 'PRDID,CUSTID,CURRID,KEYFIGUREDATE,AOPREV,AOPPROFIT';
  const dates = monthDates(12);
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        for (let i = 0; i < dates.length; i++) {
          rows.push(`${p.id},${c.custId},USD,${dates[i]},${jitter(seedVal(bpSeeds.mthRev, i))},${jitter(seedVal(bpSeeds.mthPft, i))}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLoc(locs, families) {
  const h = 'PRDID,LOCID,KEYFIGUREDATE,INVENTORYHOLDINGCOSTRATE,INVENTORYTARGETVIOLATIONCOSTRATE,NAGCSTINITIALAGGRUNITCOST,NAGCSTINVENTORYCOSTRATE,NAGCO2INITIALAGGRUNITEMIS,NAGCO2INVENTORYEMISRATE,ZSINITIALINVENTORYRESET';
  const dates = weekDates(26);
  const s = bpSeeds;
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const l of locs) {
        for (const d of dates) {
          rows.push(`${p.id},${l.id},${d},${s.wkInventoryHoldingCostRate},${s.wkInventoryTargetViolationCostRate},${s.wkNagCstInitialAggrUnitCost},${s.wkNagCstInventoryCostRate},${s.wkNagCo2InitialAggrUnitEmis},${s.wkNagCo2InventoryEmisRate},${s.wkInitialInventory}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocSrc(locs, families) {
  const h = 'PRDID,LOCID,KEYFIGUREDATE,SOURCEID,PRODUCTIONCOSTRATE,NAGCSTPRODUCTIONCOSTRATE,NAGCSTFIXEDPRODUCTIONCOST,NAGCO2PRODUCTIONEMISRATE,NAGCO2FIXEDPRODUCTIONEMIS';
  const dates = weekDates(26);
  const s = bpSeeds;
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const workLocs = plants.length ? plants : locs;
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const l of workLocs) {
        const sid = sourceId(p.id, l.id);
        for (const d of dates) {
          rows.push(`${p.id},${l.id},${d},${sid},${s.wkProductionCostRate},${s.wkNagCstProdCostRate},${s.wkNagCstFixedProdCost},${s.wkNagCo2ProdEmisRate},${s.wkNagCo2FixedProdEmis}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocCust(customers, locs, families) {
  // IPAV2 WKPRODLOCCUST: ratio/cost data only (forecast/actuals moved to WKPRODLOCSACUST)
  const h = 'PRDID,LOCID,KEYFIGUREDATE,CUSTID,CRATIO,CUSTOMERRATIO,CUSTOMERTRANSPORTATIONCOSTRATE';
  const dates = weekDates(26);
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        const nearLocs = locs.filter(l => l.region === c.region).slice(0, 1);
        const servLocs = nearLocs.length ? nearLocs : locs.slice(0, 1);
        for (const l of servLocs) {
          for (const d of dates) {
            rows.push(`${p.id},${l.id},${d},${c.custId},1,1,`);
          }
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocSaCust(customers, locs, families) {
  // IPAV2 WKPRODLOCSACUST: all products × all locs × all customers (52wk past + 26wk future)
  const h = 'PRDID,LOCID,KEYFIGUREDATE,CUSTID,SALESAREA,ACTUALSQTY,SALESFORECASTQTY,DEMANDPLANNINGQTY,WEEKLYOPENORDER';
  const pastDates = pastWeekDates(52);
  const futDates  = weekDates(26);
  const allDates  = [...pastDates, ...futDates];
  const s = bpSeeds;
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        for (const l of locs) {
          for (let i = 0; i < allDates.length; i++) {
            const isPast = i < pastDates.length;
            const actuals = isPast ? jitter(seedVal(s.wkActuals, i)) : '';
            const fcast   = isPast ? '' : jitter(seedVal(s.wkSalesFcast, i));
            const sop     = isPast ? '' : jitter(seedVal(s.wkSalesFcast, i), 0.1);
            const orders  = isPast ? jitter(seedVal(s.wkActuals, i), 0.2) : '';
            rows.push(`${p.id},${l.id},${allDates[i]},${c.custId},${salesArea(c.channel)},${actuals},${fcast},${sop},${orders}`);
          }
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocCustGroup(locs, families) {
  const h = 'PRDID,LOCID,CUSTGROUP,KEYFIGUREDATE,IOFORECAST';
  const dates = weekDates(26);
  const s = bpSeeds;
  const groups = ['Industry', 'Retail', 'Wholesale'];
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const l of locs) {
        for (const g of groups) {
          for (let i = 0; i < dates.length; i++) {
            rows.push(`${p.id},${l.id},${g},${dates[i]},${jitter(seedVal(s.ioForecastVals, i))}`);
          }
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocCustCurr(customers, locs, families) {
  const h = 'PRDID,LOCID,CUSTID,CURRID,KEYFIGUREDATE,ACTUALSREV,OPENORDERSREV,MARKETINGBUDGET';
  const pastDates = pastWeekDates(12);
  const s = bpSeeds;
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        const nearLocs = locs.filter(l => l.region === c.region).slice(0, 1);
        const servLocs = nearLocs.length ? nearLocs : locs.slice(0, 1);
        for (const l of servLocs) {
          for (let i = 0; i < pastDates.length; i++) {
            rows.push(`${p.id},${l.id},${c.custId},USD,${pastDates[i]},${jitter(seedVal(s.wkActualsRev, i))},${jitter(seedVal(s.wkOpenOrdersRev, i))},`);
          }
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocLocTo(locs, families) {
  const h = 'PRDID,LOCID,LOCTO,KEYFIGUREDATE,TRANSPORTATIONCOSTRATE';
  const dates = weekDates(13);
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const dcs = locs.filter(l => l.locType !== 'P' && l.locType !== 'Plant');
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const pl of plants) {
        const targetDcs = dcs.filter(dc => dc.region === pl.region);
        for (const dc of (targetDcs.length ? targetDcs : dcs.slice(0, 1))) {
          for (const d of dates) rows.push(`${p.id},${pl.id},${dc.id},${d},1`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkResLoc(locs) {
  // IPAV2 template: LOCID,KEYFIGUREDATE,RESID,CAPASUPPLY,...
  const h = 'LOCID,KEYFIGUREDATE,RESID,CAPASUPPLY,ALRTCAPACITYOVERLOADTHRESHOLD';
  const dates = weekDates(26);
  const rows = [];
  for (const l of locs) {
    rows.push(...dates.map(d =>
      `${l.id},${d},${l.id} ASSEMBLY,${bpSeeds.wkCapaSupply},1`
    ));
  }
  return csv(h, rows);
}

function genWkResProdLocSrc(locs, families) {
  // IPAV2 template: PRDID,LOCID,KEYFIGUREDATE,RESID,SOURCEID,PCAPACONSUMPTION
  const h = 'PRDID,LOCID,KEYFIGUREDATE,RESID,SOURCEID,PCAPACONSUMPTION';
  const dates = weekDates(26);
  const plants = locs.filter(l => l.locType === 'P' || l.locType === 'Plant');
  const workLocs = plants.length ? plants : locs;
  const rows = [];
  for (const l of workLocs) {
    for (const f of families) {
      for (const p of (f.fg || [])) {
        const sid = sourceId(p.id, l.id);
        rows.push(...dates.map(d =>
          `${p.id},${l.id},${d},${l.id} ASSEMBLY,${sid},1`
        ));
      }
    }
  }
  return csv(h, rows);
}

function genWkProdCustCurr(customers, families) {
  const h = 'PRDID,CUSTID,CURRID,KEYFIGUREDATE,PLANNEDPRICE,MARKETSIZEREV,OPPORTUNITIESREV,INDIRECTCOSTS';
  const dates = weekDates(52);
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        for (const d of dates) {
          rows.push(`${p.id},${c.custId},USD,${d},${bpSeeds.plannedPrice},,,`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdLocCurr(locs, families) {
  // IPAV2 template: PRDID,LOCID,KEYFIGUREDATE,CURRID,...,COSTPERUNIT
  const h = 'PRDID,LOCID,KEYFIGUREDATE,CURRID,COSTPERUNIT';
  const dates = weekDates(26);
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const l of locs) {
        for (const d of dates) rows.push(`${p.id},${l.id},${d},USD,${bpSeeds.costPerUnit}`);
      }
    }
  }
  return csv(h, rows);
}

function genWkProdFmlLocRgn(locs, families) {
  const h = 'LOCREGION,KEYFIGUREDATE,PRDFAMILY,INVENTORYBUDGETFLAG';
  const dates = weekDates(13);
  const regions = [...new Set(locs.map(l => l.region))];
  const rows = [];
  for (const r of regions) {
    for (const f of families) {
      for (const d of dates) rows.push(`${r},${d},${f.code},1`);
    }
  }
  return csv(h, rows);
}

function genDayProdLocCust(customers, locs, families) {
  // IPAV2 template DAYPRODLOCSACUST: all products × all locs × all customers (26wk past)
  const h = 'PRDID,LOCID,CUSTID,KEYFIGUREDATE,SALESAREA,DELIVQTY';
  const dates = pastWeekDates(26);
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        const weeklyBase = Math.round(seedVal(bpSeeds.mthAop, 0) / 4.3);
        for (const l of locs) {
          for (const d of dates) rows.push(`${p.id},${l.id},${c.custId},${d},${salesArea(c.channel)},${jitter(weeklyBase)}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genMthProdSaCustCurr(customers, families) {
  const h = 'PRDID,KEYFIGUREDATE,CUSTID,CURRID,SALESAREA,ZCPRICE';
  const dates = [...pastMonthDates(12), ...monthDates(12)];
  const price = bpSeeds.plannedPrice || 100;
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        const sa = salesArea(c.channel);
        for (const d of dates) {
          rows.push(`${p.id},${d},${c.custId},USD,${sa},${price}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genWkProdSaCustCurr(customers, families) {
  const h = 'PRDID,KEYFIGUREDATE,CUSTID,CURRID,SALESAREA,ZCPRICEW';
  const dates = [...pastWeekDates(52), ...weekDates(26)];
  const price = bpSeeds.plannedPrice || 100;
  const rows = [];
  for (const f of families) {
    for (const p of (f.fg || [])) {
      for (const c of customers) {
        const hasFam = !c.recommendedFamilies?.length || c.recommendedFamilies.includes(f.id);
        if (!hasFam) continue;
        const sa = salesArea(c.channel);
        for (const d of dates) {
          rows.push(`${p.id},${d},${c.custId},USD,${sa},${price}`);
        }
      }
    }
  }
  return csv(h, rows);
}

function genMthCostCenterCurr(customers) {
  const h = 'KEYFIGUREDATE,CURRID,COSTCENTER,ZFPMFINANCEBUDGETFACTOR,ZFPMFINANCEBUDGETADJUSTMENT';
  const dates = [...pastMonthDates(12), ...monthDates(12)];
  const rows = [];
  for (const c of customers) {
    const cc = `CC-${c.custId}`;
    for (const d of dates) {
      rows.push(`${d},USD,${cc},1,`);
    }
  }
  return csv(h, rows);
}

// Static files — copied from BP ZIP as-is
const STATIC_MASTER = [
  '060_LAG_2608.CSV', '070_SERVICELEVELCLASS_2608.CSV',
  '080_CURRENCY_2608.CSV', '090_UOMTO_2608.CSV',
  '105_EXCHANGE_RATE_2608.CSV', '220_FORECASTCONSUMPTIONMODE_2608.CSV',
  '240_DSFULFILLMENT_2608.CSV', '280_INTERVAL_2608.CSV',
];

// ─── Main generate function ───────────────────────────────────────────────────

function generateZip(analysis, selections) {
  const { customers, selectedFamilyIds, selectedLocationIds } = selections;
  const allFamilies = analysis.productFamilies || [];
  const allLocs = analysis.locations || [];

  const families = allFamilies.filter(f => selectedFamilyIds.includes(f.id));
  const locs = allLocs.filter(l => selectedLocationIds.includes(l.id));

  const bom       = analysis.bom       || [];
  const resources = analysis.resources || [];
  const routing   = analysis.routing   || [];
  const leadTimes = analysis.leadTimes || { production: 1, outsourcing: 7, distribution: 2 };

  const zip = new AdmZip();

  const addFile = (filePath, content) => {
    if (!content) return;
    zip.addFile(filePath, Buffer.from('﻿' + content, 'utf8'));
  };

  // ── Master Data ──────────────────────────────────────────────────────────
  addFile('Master_Data/010_CUSTOMER_NEW.CSV',               gen010Customer(customers));
  addFile('Master_Data/020_LOCATION_NEW.CSV',               gen020Location(locs));
  addFile('Master_Data/030_PRODUCT_NEW.CSV',                gen030Product(families));
  addFile('Master_Data/040_RESOURCE_NEW.CSV',               gen040Resource(locs, routing));
  addFile('Master_Data/100_UOMCONVERSIONFACTOR_NEW.CSV',    gen100UomConversion(families));
  addFile('Master_Data/110_CUSTOMERPRODUCT_NEW.CSV',        gen110CustomerProduct(customers, families));
  addFile('Master_Data/120_LOCATIONPRODUCT_SOP_DP_NEW.CSV', gen120LocationProduct(locs, families));
  addFile('Master_Data/130_RESOURCELOCATION_NEW.CSV',       gen130ResourceLocation(locs, resources, routing));
  addFile('Master_Data/135_RESOURCELOCATIONPRODUCT_NEW.CSV', gen135ResourceLocationProduct(locs, families));
  addFile('Master_Data/140_SOURCECUSTOMER_NEW.CSV',         gen140SourceCustomer(customers, locs, families));
  addFile('Master_Data/150_SOURCELOCATION_SOP_DP_NEW.CSV',  gen150SourceLocation(locs, families, leadTimes));
  addFile('Master_Data/160_SOURCEPRODUCTIONHEADER_NEW.CSV', gen160SourceProductionHeader(locs, families, routing, leadTimes));
  addFile('Master_Data/170_PRODUCTIONSOURCEITM_NEW.CSV',    gen170ProductionSourceItm(locs, families, bom));
  addFile('Master_Data/180_PRODUCTIONSOURCERESOURCE_NEW.CSV', gen180ProductionSourceResource(locs, families, routing, resources));
  addFile('Master_Data/190_SOURCECUSTGROUP_NEW.CSV',        gen190SourceCustGroup(customers, locs, families));
  addFile('Master_Data/200_SA_NEW.CSV',                     gen200SalesArea(customers));

  // Static master data from BP ZIP
  for (const fname of STATIC_MASTER) {
    if (bpData[fname]) {
      const newName = fname.replace('_2608.', '_REF.');
      addFile(`Master_Data/${newName}`, bpData[fname]);
    }
  }

  // ── KF Data ──────────────────────────────────────────────────────────────
  addFile('KF_Data/IPAV2__WKPRODLOC_NEW.csv',           genWkProdLoc(locs, families));
  addFile('KF_Data/IPAV2__WKPRODLOCSRC_NEW.csv',        genWkProdLocSrc(locs, families));
  addFile('KF_Data/IPAV2__WKPRODLOCCUST_NEW.csv',       genWkProdLocCust(customers, locs, families));
  addFile('KF_Data/IPAV2__WKPRODLOCSACUST_NEW.csv',     genWkProdLocSaCust(customers, locs, families));
  addFile('KF_Data/IPAV2__WKPRODLOCCURR_NEW.csv',       genWkProdLocCurr(locs, families));
  addFile('KF_Data/IPAV2__WKRESLOC_NEW.csv',            genWkResLoc(locs));
  addFile('KF_Data/IPAV2__WKRESPRODLOCSRC_NEW.csv',     genWkResProdLocSrc(locs, families));
  addFile('KF_Data/IPAV2__DAYPRODLOCSACUST_NEW.csv',    genDayProdLocCust(customers, locs, families));
  addFile('KF_Data/IPAV2__MTHCOSTCENTERCURR_NEW.csv',   genMthCostCenterCurr(customers));
  addFile('KF_Data/IPAV2__MTHPRODSACUSTCURR_NEW.csv',   genMthProdSaCustCurr(customers, families));
  addFile('KF_Data/IPAV2__WKPRODSACUSTCURR_NEW.csv',    genWkProdSaCustCurr(customers, families));

  // README
  const company = analysis.company || {};
  const totalFiles = 15 + STATIC_MASTER.filter(f => bpData[f]).length + 11;
  addFile('README.txt', `IBP Configuration Package — Planning Area IPAV2
Company: ${company.name || ''}
Industry: ${company.industry || ''}
Generated: ${new Date().toISOString()}

Master Data files: 15 generated + ${STATIC_MASTER.filter(f => bpData[f]).length} static (from BP 2608)
KF Data files: 11 generated (IPAV2 KF sets)
Total: ${totalFiles} files

Import Order:
  1. 010_CUSTOMER_NEW.CSV
  2. 020_LOCATION_NEW.CSV
  3. 030_PRODUCT_NEW.CSV
  4. 040_RESOURCE_NEW.CSV
  5. 060_LAG_REF.CSV
  6. 070_SERVICELEVELCLASS_REF.CSV
  7. 080_CURRENCY_REF.CSV
  8. 090_UOMTO_REF.CSV
  9. 100_UOMCONVERSIONFACTOR_NEW.CSV
 10. 105_EXCHANGE_RATE_REF.CSV
 11. 110_CUSTOMERPRODUCT_NEW.CSV
 12. 120_LOCATIONPRODUCT_SOP_DP_NEW.CSV
 13. 130_RESOURCELOCATION_NEW.CSV
 14. 140_SOURCECUSTOMER_NEW.CSV
 15. 150_SOURCELOCATION_SOP_DP_NEW.CSV
 16. 160_SOURCEPRODUCTIONHEADER_NEW.CSV
 17. 170_PRODUCTIONSOURCEITM_NEW.CSV
 18. 180_PRODUCTIONSOURCERESOURCE_NEW.CSV
 19. 190_SOURCECUSTGROUP_NEW.CSV
 20. 200_SA_NEW.CSV
 21. 220_FORECASTCONSUMPTIONMODE_REF.CSV
 21. 240_DSFULFILLMENT_REF.CSV
 22. 280_INTERVAL_REF.CSV
 Then KF_Data files (any order).

IPAV2 KF Sets:
  IPAV2__WKPRODLOC          — Inventory, cost, CO2 (weekly / product-location)
  IPAV2__WKPRODLOCSRC       — Production cost & CO2 (weekly / product-location-source)
  IPAV2__WKPRODLOCCUST      — Customer ratios (weekly / product-location-customer)
  IPAV2__WKPRODLOCSACUST    — Demand planning actuals & forecast (weekly / product-location-salesarea-customer)
  IPAV2__WKPRODLOCCURR      — Cost per unit (weekly / product-location-currency)
  IPAV2__WKRESLOC           — Resource capacity supply (weekly / location)
  IPAV2__WKRESPRODLOCSRC    — Resource capacity consumption (weekly / product-location-resource-source)
  IPAV2__DAYPRODLOCSACUST   — Delivered quantity actuals (weekly / product-location-salesarea-customer)
  IPAV2__MTHCOSTCENTERCURR  — Finance budget factor (monthly / cost-center-currency)
  IPAV2__MTHPRODSACUSTCURR  — Monthly price (monthly / product-salesarea-customer-currency)
  IPAV2__WKPRODSACUSTCURR   — Weekly price (weekly / product-salesarea-customer-currency)

KF seed values based on SAP IBP Best Practice 2608 reference data.
`);

  return zip.toBuffer();
}

// ─── Web fetch helpers (for company analysis) ─────────────────────────────────

async function safeFetch(url, ms = 8000) {
  try {
    const r = await fetch(url, {
      signal: AbortSignal.timeout(ms),
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; IBP-Generator/1.0)', 'Accept': 'text/html,*/*' },
    });
    if (!r.ok) return null;
    return await r.text();
  } catch { return null; }
}

function stripHtml(html, maxLen = 6000) {
  return html
    .replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<nav[\s\S]*?<\/nav>/gi, '').replace(/<footer[\s\S]*?<\/footer>/gi, '')
    .replace(/<[^>]+>/g, ' ').replace(/&[a-z]+;/g, ' ').replace(/\s{2,}/g, ' ')
    .trim().slice(0, maxLen);
}

async function searchCompanyInfo(name) {
  const q = encodeURIComponent(name);
  let officialUrl = null, ddgSummary = '';
  const ddg = await safeFetch(`https://api.duckduckgo.com/?q=${q}&format=json&no_html=1&skip_disambig=1`, 5000);
  if (ddg) {
    try {
      const d = JSON.parse(ddg);
      if (d.AbstractURL) officialUrl = d.AbstractURL;
      if (d.AbstractText) ddgSummary = d.AbstractText;
      if (!officialUrl && d.Results?.length) officialUrl = d.Results[0].FirstURL;
    } catch {}
  }
  const pagesContent = [];
  if (officialUrl?.startsWith('http')) {
    const base = officialUrl.replace(/\/$/, '');
    for (const url of [base, `${base}/about`, `${base}/about-us`, `${base}/global`]) {
      if (pagesContent.length >= 2) break;
      const html = await safeFetch(url);
      if (html && html.length > 500) {
        const text = stripHtml(html);
        if (text.length > 300) pagesContent.push(`[${url}]\n${text}`);
      }
    }
  }
  return { officialUrl, ddgSummary, pagesContent };
}

// ─── SAP AI Core OAuth2 + Inference ──────────────────────────────────────────

let _aiToken = null, _aiTokenExp = 0, _deploymentId = null;

async function getSapToken() {
  if (_aiToken && Date.now() < _aiTokenExp) return _aiToken;
  const { AI_AUTH_URL, AI_CLIENT_ID, AI_CLIENT_SECRET } = process.env;
  if (!AI_AUTH_URL || !AI_CLIENT_ID || !AI_CLIENT_SECRET)
    throw new Error('SAP AI Core credentials missing in .env (AI_AUTH_URL, AI_CLIENT_ID, AI_CLIENT_SECRET)');
  const creds = Buffer.from(`${AI_CLIENT_ID}:${AI_CLIENT_SECRET}`).toString('base64');
  const r = await fetch(`${AI_AUTH_URL}/oauth/token`, {
    method: 'POST',
    headers: { 'Authorization': `Basic ${creds}`, 'Content-Type': 'application/x-www-form-urlencoded' },
    body: 'grant_type=client_credentials',
  });
  if (!r.ok) throw new Error(`SAP Auth failed: ${r.status} ${await r.text()}`);
  const d = await r.json();
  _aiToken = d.access_token;
  _aiTokenExp = Date.now() + (d.expires_in - 60) * 1000;
  return _aiToken;
}

async function getClaudeDeployment(token) {
  if (_deploymentId) return _deploymentId;
  const apiUrl = process.env.AI_API_URL;

  // If explicitly set in .env, use it directly
  if (process.env.AI_DEPLOYMENT_ID) {
    _deploymentId = process.env.AI_DEPLOYMENT_ID;
    console.log(`[ai-core] Using deployment from .env: ${_deploymentId}`);
    return _deploymentId;
  }

  // Discover resource groups first
  const rgRes = await fetch(`${apiUrl}/v2/resourceGroups`, {
    headers: { 'Authorization': `Bearer ${token}` },
  });
  const rgGroups = rgRes.ok ? (await rgRes.json()).resources || [] : [];
  const groupNames = rgGroups.map(g => g.resourceGroupId || g.id).filter(Boolean);
  if (!groupNames.length) groupNames.push('default');
  console.log(`[ai-core] Resource groups: ${groupNames.join(', ')}`);

  // Search each group for a running Claude deployment
  for (const rg of groupNames) {
    const r = await fetch(`${apiUrl}/v2/inference/deployments`, {
      headers: { 'Authorization': `Bearer ${token}`, 'AI-Resource-Group': rg },
    });
    if (!r.ok) { const body = await r.text(); console.warn(`[ai-core] Group "${rg}" list failed: ${r.status} ${body}`); continue; }
    const d = await r.json();
    const dep = (d.resources || []).find(x =>
      x.status === 'RUNNING' && JSON.stringify(x).toLowerCase().includes('claude')
    );
    if (dep) {
      _deploymentId = dep.id;
      console.log(`[ai-core] Claude deployment "${_deploymentId}" in group "${rg}"`);
      return _deploymentId;
    }
  }
  throw new Error(`No running Claude deployment found. Check AI Launchpad. Resource groups checked: ${groupNames.join(', ')}`);
}

async function callSapAiCore(system, userMsg) {
  const token = await getSapToken();
  const inferenceUrl = process.env.AI_INFERENCE_URL ||
    `${process.env.AI_API_URL}/v2/inference/deployments/${process.env.AI_DEPLOYMENT_ID}`;
  const rg = process.env.AI_RESOURCE_GROUP;
  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
  };
  if (rg) headers['AI-Resource-Group'] = rg;

  console.log(`[ai-core] POST ${inferenceUrl}/invoke (rg=${rg})`);
  const r = await fetch(`${inferenceUrl}/invoke`, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      anthropic_version: 'bedrock-2023-05-31',
      max_tokens: 4096,
      system,
      messages: [{ role: 'user', content: userMsg }],
    }),
  });
  if (!r.ok) throw new Error(`SAP AI Core ${r.status}: ${await r.text()}`);
  return await r.json();
}

// ─── Excel Import ─────────────────────────────────────────────────────────────

const CITY_COORDS = {
  // Korea
  '서울': { lat: 37.57, lon: 126.98 }, 'Seoul': { lat: 37.57, lon: 126.98 },
  '부산': { lat: 35.18, lon: 129.08 }, 'Busan': { lat: 35.18, lon: 129.08 },
  '청주': { lat: 36.64, lon: 127.49 }, 'Cheongju': { lat: 36.64, lon: 127.49 },
  '울산': { lat: 35.54, lon: 129.31 }, 'Ulsan': { lat: 35.54, lon: 129.31 },
  '대구': { lat: 35.87, lon: 128.60 }, 'Daegu': { lat: 35.87, lon: 128.60 },
  '인천': { lat: 37.46, lon: 126.70 }, 'Incheon': { lat: 37.46, lon: 126.70 },
  '광주': { lat: 35.16, lon: 126.85 }, 'Gwangju': { lat: 35.16, lon: 126.85 },
  '대전': { lat: 36.35, lon: 127.38 }, 'Daejeon': { lat: 36.35, lon: 127.38 },
  '수원': { lat: 37.27, lon: 127.01 }, 'Suwon': { lat: 37.27, lon: 127.01 },
  '창원': { lat: 35.23, lon: 128.68 }, 'Changwon': { lat: 35.23, lon: 128.68 },
  '안산': { lat: 37.32, lon: 126.83 }, '평택': { lat: 36.99, lon: 127.11 },
  // Asia
  'Tokyo': { lat: 35.68, lon: 139.69 }, '東京': { lat: 35.68, lon: 139.69 },
  'Osaka': { lat: 34.69, lon: 135.50 }, 'Nagoya': { lat: 35.18, lon: 136.91 },
  'Shanghai': { lat: 31.23, lon: 121.47 }, '上海': { lat: 31.23, lon: 121.47 },
  'Beijing': { lat: 39.91, lon: 116.39 }, '北京': { lat: 39.91, lon: 116.39 },
  'Guangzhou': { lat: 23.13, lon: 113.26 }, 'Shenzhen': { lat: 22.54, lon: 114.06 },
  'Singapore': { lat: 1.35, lon: 103.82 },
  'Bangkok': { lat: 13.75, lon: 100.52 },
  'Mumbai': { lat: 19.08, lon: 72.88 }, 'Delhi': { lat: 28.61, lon: 77.21 },
  'Sydney': { lat: -33.87, lon: 151.21 }, 'Melbourne': { lat: -37.81, lon: 144.96 },
  'Dubai': { lat: 25.20, lon: 55.27 },
  'Taipei': { lat: 25.04, lon: 121.56 },
  // Americas
  'New York': { lat: 40.71, lon: -74.01 }, 'Los Angeles': { lat: 34.05, lon: -118.24 },
  'Chicago': { lat: 41.88, lon: -87.63 }, 'Houston': { lat: 29.76, lon: -95.37 },
  'São Paulo': { lat: -23.55, lon: -46.63 }, 'Mexico City': { lat: 19.43, lon: -99.13 },
  // Europe
  'London': { lat: 51.51, lon: -0.13 }, 'Paris': { lat: 48.85, lon: 2.35 },
  'Frankfurt': { lat: 50.11, lon: 8.68 }, 'Berlin': { lat: 52.52, lon: 13.40 },
  'Munich': { lat: 48.14, lon: 11.58 }, 'Amsterdam': { lat: 52.37, lon: 4.90 },
  'Milan': { lat: 45.46, lon: 9.19 }, 'Madrid': { lat: 40.42, lon: -3.70 },
};

const COUNTRY_HQ = {
  KR: { city: 'Seoul',     lat: 37.57,  lon: 126.98 },
  JP: { city: 'Tokyo',     lat: 35.68,  lon: 139.69 },
  CN: { city: 'Shanghai',  lat: 31.23,  lon: 121.47 },
  US: { city: 'New York',  lat: 40.71,  lon: -74.01 },
  DE: { city: 'Frankfurt', lat: 50.11,  lon: 8.68   },
  GB: { city: 'London',    lat: 51.51,  lon: -0.13  },
  FR: { city: 'Paris',     lat: 48.85,  lon: 2.35   },
  SG: { city: 'Singapore', lat: 1.35,   lon: 103.82 },
  IN: { city: 'Mumbai',    lat: 19.08,  lon: 72.88  },
  AU: { city: 'Sydney',    lat: -33.87, lon: 151.21 },
  AE: { city: 'Dubai',     lat: 25.20,  lon: 55.27  },
};

const COUNTRY_REGION = {
  KR: 'APJ', JP: 'APJ', CN: 'APJ', SG: 'APJ', IN: 'APJ', AU: 'APJ', TW: 'APJ', VN: 'APJ', TH: 'APJ', MY: 'APJ',
  US: 'AMER', CA: 'AMER', MX: 'AMER', BR: 'AMER', CL: 'AMER', CO: 'AMER',
  DE: 'EMEA', GB: 'EMEA', FR: 'EMEA', NL: 'EMEA', IT: 'EMEA', ES: 'EMEA', AE: 'EMEA', SA: 'EMEA', PL: 'EMEA',
};

function getCityCoords(str) {
  if (!str) return null;
  const key = Object.keys(CITY_COORDS).find(c => str.includes(c));
  return key ? CITY_COORDS[key] : null;
}

function getCountryRegion(code) {
  return COUNTRY_REGION[String(code).toUpperCase()] || 'APJ';
}

function extractLocations(wb, mapping) {
  if (!mapping.locationSheet?.sheetName) return [];
  const ws = wb.Sheets[mapping.locationSheet.sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const m = mapping.locationSheet;
  const skip = m.headerRows ?? 1;
  return rows.slice(skip)
    .filter(r => String(r[m.idCol] ?? '').trim())
    .map((r, i) => {
      const id      = String(r[m.idCol]).trim();
      const name    = String(r[m.nameCol] ?? '').trim();
      const city    = m.cityCol    != null ? String(r[m.cityCol]    ?? '').trim() : '';
      const rawCtry = m.countryCol != null ? String(r[m.countryCol] ?? '').trim() : '';
      const country = (rawCtry.length === 2 ? rawCtry.toUpperCase() : null) || mapping.companyCountry || 'KR';
      const coords  = getCityCoords(city) || getCityCoords(name) || COUNTRY_HQ[country] || { lat: 37.57 + i * 0.1, lon: 126.98 + i * 0.1 };
      return { id: `LOC-${id}`, desc: name || id, locType: 'P', region: getCountryRegion(country), country, city: city || name, lat: coords.lat, lon: coords.lon };
    });
}

function extractCustomers(wb, mapping) {
  if (!mapping.customerSheet?.sheetName) return [];
  const ws = wb.Sheets[mapping.customerSheet.sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const m = mapping.customerSheet;
  const skip = m.headerRows ?? 1;
  return rows.slice(skip)
    .filter(r => String(r[m.idCol] ?? '').trim())
    .slice(0, 10)
    .map((r, i) => {
      const id      = String(r[m.idCol]).trim();
      const name    = String(r[m.nameCol] ?? '').trim();
      const rawCtry = m.countryCol != null ? String(r[m.countryCol] ?? '').trim() : '';
      const country = (rawCtry.length === 2 ? rawCtry.toUpperCase() : null) || mapping.companyCountry || 'KR';
      const city    = m.cityCol    != null ? String(r[m.cityCol]    ?? '').trim() : '';
      const channel = m.channelCol != null ? String(r[m.channelCol] ?? '').trim() : 'Direct';
      const coords  = getCityCoords(city) || COUNTRY_HQ[country] || { lat: 37.57 + (i % 5) * 0.2, lon: 126.98 + (i % 5) * 0.3 };
      return { custId: id, country, region: getCountryRegion(country), channel: channel || 'Direct', group: 'Industry', custdescr: (name && name !== id) ? name : `Customer ${id}`, lat: coords.lat, lon: coords.lon, recommendedFamilies: [] };
    });
}

function extractProductFamilies(wb, mapping) {
  if (!mapping.productSheet?.sheetName) return [];
  const ws = wb.Sheets[mapping.productSheet.sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const m = mapping.productSheet;
  const skip = m.headerRows ?? 1;
  const familyMap = {};
  for (const row of rows.slice(skip)) {
    const id  = String(row[m.idCol]   ?? '').trim();
    if (!id) continue;
    const desc = String(row[m.nameCol] ?? '').trim();
    const fam  = m.familyCol != null ? String(row[m.familyCol] ?? '').trim() : 'General';
    const key  = fam || 'General';
    if (!familyMap[key]) familyMap[key] = [];
    if (familyMap[key].length < 10) familyMap[key].push({ id, desc: desc || id });
  }
  const ICONS = ['📦', '🎨', '🏠', '⚙️', '🔧', '💡', '🚗', '🌿', '🧴', '🔩'];
  return Object.entries(familyMap).map(([famName, products], i) => {
    const fcode = `F${String(i + 1).padStart(3, '0')}`;
    const brand = famName.replace(/\s+/g, '').toUpperCase().slice(0, 8) || `BRAND${fcode}`;
    return { id: fcode, code: `FAMILY ${fcode}-${famName}`, label: famName, icon: ICONS[i % ICONS.length], brand, group: brand.slice(0, 6), fg: products.map(p => ({ id: p.id, desc: p.desc, abc: 'A', plm: 'EXISTING' })), semi: [], raw: [] };
  });
}

function buildFallbackMapping(sheetNames) {
  const find = (...cands) => sheetNames.find(s => cands.some(c => s.toLowerCase().includes(c.toLowerCase()))) || null;
  return {
    companyName: 'Unknown Company',
    companyCountry: 'KR',
    locationSheet:  { sheetName: find('공장','plant','factory','location','창고','warehouse'), idCol: 0, nameCol: 1, cityCol: null, countryCol: null, headerRows: 1 },
    customerSheet:  { sheetName: find('거래처','customer','client','account','고객'), idCol: 0, nameCol: 1, countryCol: null, cityCol: null, channelCol: null, headerRows: 1 },
    productSheet:   { sheetName: find('품목','product','item','material','제품','상품'), idCol: 0, nameCol: 1, familyCol: 3, subFamilyCol: 4, headerRows: 3 },
    productHierarchySheet: null,
    bomSheet:       { sheetName: find('BOM','bom','자재명세'), plantCol: 0, parentCol: 1, seqCol: 6, componentCol: 7, qtyCol: 8, uomCol: 9, headerRows: 2 },
    resourceSheet:  { sheetName: find('BOR','생산성','설비 능력','resource'), locationCol: 0, resourceIdCol: 1, dailyCapCol: 2, hourlyCapCol: 3, headerRows: 3 },
    routingSheet:   { sheetName: find('공정','routing','LX_추가'), productCol: 0, plantCol: 2, groupCol: 3, operationCol: 6, workCenterCol: 7, headerRows: 2 },
    leadTimeSheet:  { sheetName: find('BOD','lead','리드타임'), plantCol: 0, ltCol: 3, headerRows: 2 },
  };
}

function extractBom(wb, mapping) {
  if (!mapping.bomSheet?.sheetName) return [];
  const ws = wb.Sheets[mapping.bomSheet.sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const m = mapping.bomSheet;
  const skip = m.headerRows ?? 2;
  const seen = new Set();
  const result = [];
  for (const row of rows.slice(skip)) {
    const plant     = String(row[m.plantCol     ?? 0] ?? '').trim();
    const parent    = String(row[m.parentCol    ?? 1] ?? '').trim();
    const seq       = String(row[m.seqCol       ?? 6] ?? '').trim();
    const component = String(row[m.componentCol ?? 7] ?? '').trim();
    const qty       = parseFloat(row[m.qtyCol   ?? 8]) || 1;
    const uom       = String(row[m.uomCol       ?? 9] ?? '').trim() || 'EA';
    if (!parent || !component) continue;
    const key = `${plant}|${parent}|${component}|${seq}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (result.length < 1000) result.push({ plant, parent, component, qty, uom, seq });
  }
  return result;
}

function extractResources(wb, mapping) {
  if (!mapping.resourceSheet?.sheetName) return [];
  const ws = wb.Sheets[mapping.resourceSheet.sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const m = mapping.resourceSheet;
  const skip = m.headerRows ?? 3;
  const seen = new Set();
  const result = [];
  for (const row of rows.slice(skip)) {
    const locCode  = String(row[m.locationCol   ?? 0] ?? '').trim();
    const resId    = String(row[m.resourceIdCol ?? 1] ?? '').trim();
    const dailyCap = parseFloat(row[m.dailyCapCol ?? 2]) || null;
    const hourlyCap = parseFloat(row[m.hourlyCapCol ?? 3]) || null;
    if (!locCode || !resId) continue;
    const key = `${locCode}|${resId}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push({ locationId: `LOC-${locCode}`, resourceId: resId, dailyCap, hourlyCap });
  }
  return result;
}

function extractRouting(wb, mapping) {
  if (!mapping.routingSheet?.sheetName) return [];
  const ws = wb.Sheets[mapping.routingSheet.sheetName];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const m = mapping.routingSheet;
  const skip = m.headerRows ?? 2;
  const seen = new Set();
  const result = [];
  for (const row of rows.slice(skip)) {
    const product    = String(row[m.productCol    ?? 0] ?? '').trim();
    const plant      = String(row[m.plantCol      ?? 2] ?? '').trim();
    const group      = String(row[m.groupCol      ?? 3] ?? '').trim();
    const operation  = String(row[m.operationCol  ?? 6] ?? '').trim();
    const workCenter = String(row[m.workCenterCol ?? 7] ?? '').trim();
    if (!product || !plant || !workCenter) continue;
    const key = `${product}|${plant}|${workCenter}`;
    if (seen.has(key)) continue;
    seen.add(key);
    if (result.length < 1000) result.push({ product, plant, locationId: `LOC-${plant}`, group, operation, workCenter });
  }
  return result;
}

function extractLeadTimes(wb, mapping) {
  const lt = { production: 1, outsourcing: 7, distribution: 2 };

  // From BOD-like sheet: LT column (e.g. "2일", "2 days")
  if (mapping.leadTimeSheet?.sheetName) {
    const ws = wb.Sheets[mapping.leadTimeSheet.sheetName];
    if (ws) {
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      const m = mapping.leadTimeSheet;
      for (const row of rows.slice(m.headerRows ?? 2)) {
        const val = String(row[m.ltCol ?? 3] ?? '').trim();
        const days = parseInt(val);
        if (!isNaN(days) && days > 0) { lt.distribution = days; break; }
      }
    }
  }

  // Scan all sheets for "Lead Time N" pattern
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    for (const row of rows.slice(0, 5)) {
      for (const cell of row) {
        const m = String(cell).match(/Lead\s*Time\s+(\d+)/i);
        if (m) { lt.outsourcing = parseInt(m[1]); break; }
      }
    }
    if (lt.outsourcing !== 7) break;
  }

  // Scan 문의/inquiry sheet for production lead time
  const inqWs = wb.Sheets['문의'] || wb.Sheets['문의사항'] || wb.Sheets['inquiry'];
  if (inqWs) {
    const rows = XLSX.utils.sheet_to_json(inqWs, { header: 1, defval: '' });
    for (const row of rows) {
      for (const cell of row) {
        const m = String(cell).match(/(\d+)\s*일\b/);
        if (m) { const d = parseInt(m[1]); if (d >= 1 && d <= 30) { lt.production = d; break; } }
      }
    }
  }

  return lt;
}

async function parseExcelToIbpData(bufOrPath, fileName = '') {
  const wb = typeof bufOrPath === 'string'
    ? XLSX.readFile(bufOrPath)
    : XLSX.read(bufOrPath, { type: 'buffer' });

  // ── Step 1: Sample up to 12 sheets, first 20 rows each ──
  const samples = {};
  for (const name of wb.SheetNames.slice(0, 12)) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    samples[name] = rows.slice(0, 20);
  }

  const sampleText = Object.entries(samples).map(([name, rows]) => {
    const preview = rows.map((r, i) =>
      `  row${i}: ${r.slice(0, 10).map(v => JSON.stringify(String(v ?? '').slice(0, 30))).join(' | ')}`
    ).join('\n');
    return `Sheet "${name}":\n${preview}`;
  }).join('\n\n');

  // ── Step 2: Ask Claude to identify column mapping ──
  const prompt = `You are analyzing an Excel file to extract master data for SAP IBP.
File: ${fileName || 'unknown.xlsx'}

Sheet sample data:
${sampleText}

Return ONLY a JSON object (no markdown, no explanation) with this structure:
{
  "companyName": "<infer from data or filename>",
  "companyCountry": "<ISO-2 country code, e.g. KR, US, DE, JP>",
  "locationSheet": {
    "sheetName": "<sheet with factories/plants/warehouses>",
    "idCol": <0-based column for location code>,
    "nameCol": <0-based column for location name>,
    "cityCol": <column for city or null>,
    "countryCol": <column for country or null>,
    "headerRows": <rows to skip, default 1>
  },
  "customerSheet": {
    "sheetName": "<sheet with customers/accounts>",
    "idCol": <0-based column for customer ID>,
    "nameCol": <0-based column for customer name>,
    "countryCol": <column for country or null>,
    "cityCol": <column for city or null>,
    "channelCol": <column for sales channel or null>,
    "headerRows": <rows to skip>
  },
  "productSheet": {
    "sheetName": "<sheet with products/items/SKUs>",
    "idCol": <0-based column for product code>,
    "nameCol": <0-based column for product name>,
    "familyCol": <column for product family/group or null>,
    "subFamilyCol": <column for sub-category or null>,
    "headerRows": <rows to skip>
  },
  "productHierarchySheet": null,
  "bomSheet": {
    "sheetName": "<sheet with BOM / bill of materials (parent-component structure)>",
    "plantCol": <0-based column for plant code>,
    "parentCol": <column for parent/finished product code>,
    "seqCol": <column for sequence/line number>,
    "componentCol": <column for component/material code>,
    "qtyCol": <column for component quantity per parent>,
    "uomCol": <column for unit of measure>,
    "headerRows": <rows to skip>
  },
  "resourceSheet": {
    "sheetName": "<sheet with resources/machines/equipment capacity (BOR or similar)>",
    "locationCol": <0-based column for plant/location code>,
    "resourceIdCol": <column for resource/machine ID>,
    "dailyCapCol": <column for daily capacity or null>,
    "hourlyCapCol": <column for hourly capacity or null>,
    "headerRows": <rows to skip>
  },
  "routingSheet": {
    "sheetName": "<sheet with production routing / work center assignments per product>",
    "productCol": <0-based column for product/material code>,
    "plantCol": <column for plant code>,
    "groupCol": <column for routing group or null>,
    "operationCol": <column for operation number>,
    "workCenterCol": <column for work center / machine ID>,
    "headerRows": <rows to skip>
  },
  "leadTimeSheet": {
    "sheetName": "<sheet with lead time data (BOD, sourcing, distribution lead times)>",
    "plantCol": <0-based column for source plant>,
    "ltCol": <column containing lead time value (days)>,
    "headerRows": <rows to skip>
  }
}
Set any sheet key to null if the sheet does not exist.`;

  let mapping;
  try {
    const r = await callSapAiCore(
      'You are an SAP IBP data analyst. Return ONLY valid JSON, no explanation, no markdown.',
      prompt
    );
    const text = (r.content?.[0]?.text || '').trim();
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    mapping = JSON.parse(jsonMatch ? jsonMatch[0] : text);
    console.log('[import-excel] AI mapping:', JSON.stringify(mapping).slice(0, 300));
  } catch (err) {
    console.warn('[import-excel] AI mapping failed, using heuristics:', err.message);
    mapping = buildFallbackMapping(wb.SheetNames);
  }

  // ── Step 3: Extract data using the mapping ──
  const locations       = extractLocations(wb, mapping);
  const customers       = extractCustomers(wb, mapping);
  const productFamilies = extractProductFamilies(wb, mapping);
  const bom             = extractBom(wb, mapping);
  const resources       = extractResources(wb, mapping);
  const routing         = extractRouting(wb, mapping);
  const leadTimes       = extractLeadTimes(wb, mapping);

  console.log(`[import-excel] bom=${bom.length}, resources=${resources.length}, routing=${routing.length}, lt=${JSON.stringify(leadTimes)}`);

  const country = mapping.companyCountry || 'KR';
  const hq = COUNTRY_HQ[country] || COUNTRY_HQ['KR'];

  return {
    company: {
      name: mapping.companyName || 'Unknown Company',
      industry: 'Manufacturing',
      hq: { country, city: hq.city, region: getCountryRegion(country), lat: hq.lat, lon: hq.lon },
    },
    productFamilies,
    locations,
    customers,
    bom,
    resources,
    routing,
    leadTimes,
    _importedFromExcel: true,
    _aiMapped: true,
    _meta: { source: fileName },
  };
}

// ─── Analysis ─────────────────────────────────────────────────────────────────

const SYSTEM = `You are an SAP IBP Best Practice configuration expert. Return ONLY valid JSON, no explanation.`;

async function analyzeWithClaude(companyName, webInfo) {
  const webCtx = webInfo.pagesContent.length
    ? `Website:\n${webInfo.pagesContent.join('\n\n').slice(0, 10000)}`
    : webInfo.ddgSummary ? `Summary: ${webInfo.ddgSummary}` : `Use your knowledge about "${companyName}".`;

  const r = await callSapAiCore(SYSTEM, `Analyze "${companyName}" and return IBP config JSON.

${webCtx}

Return exactly this JSON structure (no markdown):
{
  "company": { "name": "...", "industry": "...", "hq": { "country": "ISO2", "city": "...", "region": "EMEA|AMERICAS|APJ", "lat": 0.0, "lon": 0.0 } },
  "productFamilies": [
    { "id": "F100", "code": "FAMILY 100-PRODUCTNAME", "label": "Short Label", "icon": "emoji", "brand": "BRND100", "group": "GROUPNAME",
      "fg": [{"id":"PRD-100","desc":"desc","abc":"A","plm":"EXISTING"},{"id":"PRD-110","desc":"desc","abc":"B","plm":"EXISTING"},{"id":"PRD-120","desc":"desc","abc":"C","plm":"EXISTING"}],
      "semi": [{"id":"PRD-101-S"},{"id":"PRD-111-S"},{"id":"PRD-121-S"}],
      "raw": [{"id":"PRD-101-R"},{"id":"PRD-102-R"},{"id":"PRD-111-R"}] }
  ],
  "locations": [
    { "id": "LOC-XXXX", "desc": "Plant City Country", "locType": "P", "region": "EMEA|AMERICAS|APJ", "country": "ISO2", "city": "city", "lat": 0.0, "lon": 0.0 }
  ],
  "customers": [
    { "country": "ISO2", "region": "EMEA|AMERICAS|APJ", "channel": "Direct|Indirect|Online", "group": "Retail|Wholesale|Industry|Trading|PrivateCustomer", "custdescr": "...", "lat": 0.0, "lon": 0.0, "recommendedFamilies": ["F100"] }
  ]
}

Rules: 2-4 product families representing REAL product lines. 3-8 locations (Plants locType=P, DCs locType=D). 3-6 customers. Product IDs unique, use category prefix (e.g. FLOOR-100). Accurate lat/lon.`);
  const text = (r.content?.[0]?.text || r.choices?.[0]?.message?.content || '').trim();
  const m = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || [null, text];
  return JSON.parse(m[1]);
}

// ─── HTTP Server ──────────────────────────────────────────────────────────────

function cors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

const server = http.createServer(async (req, res) => {
  cors(res);
  if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }
  const url = new URL(req.url, `http://localhost:${PORT}`);

  if (req.method === 'GET' && url.pathname === '/') {
    try {
      const html = fs.readFileSync(path.join(__dirname, 'public', 'index.html'), 'utf8');
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); res.end(html);
    } catch { res.writeHead(404); res.end('Not found'); }
    return;
  }

  if (req.method === 'GET' && url.pathname === '/api/ai-test') {
    try {
      const token = await getSapToken();
      const apiUrl = process.env.AI_API_URL;
      const rg = process.env.AI_RESOURCE_GROUP || 'default';
      const [rgRes, depRes, depDirect] = await Promise.all([
        fetch(`${apiUrl}/v2/resourceGroups`, { headers: { 'Authorization': `Bearer ${token}` } }),
        fetch(`${apiUrl}/v2/inference/deployments`, { headers: { 'Authorization': `Bearer ${token}`, 'AI-Resource-Group': rg } }),
        fetch(`${apiUrl}/v2/inference/deployments/${process.env.AI_DEPLOYMENT_ID}`, { headers: { 'Authorization': `Bearer ${token}`, 'AI-Resource-Group': rg } }),
      ]);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        resourceGroups: rgRes.ok ? await rgRes.json() : { status: rgRes.status, body: await rgRes.text() },
        deploymentList: depRes.ok ? await depRes.json() : { status: depRes.status, body: await depRes.text() },
        deploymentDirect: depDirect.ok ? await depDirect.json() : { status: depDirect.status, body: await depDirect.text() },
      }, null, 2));
    } catch (e) { res.writeHead(500); res.end(JSON.stringify({ error: e.message })); }
    return;
  }


  if (req.method === 'GET' && url.pathname === '/api/seeds') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ loaded: Object.keys(bpData).length, seeds: bpSeeds }));
    return;
  }

  let body = '';
  for await (const c of req) body += c;

  if (req.method === 'POST' && url.pathname === '/api/import-excel') {
    try {
      const { fileData, fileName } = JSON.parse(body);
      if (!fileData) { res.writeHead(400); res.end(JSON.stringify({ error: 'fileData required' })); return; }
      console.log(`[import-excel] "${fileName}" (${Math.round(fileData.length * 0.75 / 1024)} KB)`);
      const buf = Buffer.from(fileData, 'base64');
      const result = await parseExcelToIbpData(buf, fileName || '');
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result));
    } catch (err) {
      console.error('[import-excel error]', err.message);
      res.writeHead(500); res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/analyze') {
    let companyName;
    try { companyName = JSON.parse(body).company?.trim(); } catch {}
    if (!companyName) { res.writeHead(400); res.end(JSON.stringify({ error: 'company required' })); return; }
    try {
      console.log(`[analyze] "${companyName}"`);
      const webInfo = await searchCompanyInfo(companyName);
      const result = await analyzeWithClaude(companyName, webInfo);
      result._meta = { officialUrl: webInfo.officialUrl, pagesFound: webInfo.pagesContent.length };
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(result));
    } catch (err) {
      console.error('[analyze error]', err.message, err.status, err.error);
      res.writeHead(500); res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  if (req.method === 'POST' && url.pathname === '/api/generate') {
    let payload;
    try { payload = JSON.parse(body); } catch { res.writeHead(400); res.end('bad json'); return; }
    try {
      console.log(`[generate] ${payload.analysis?.company?.name}`);
      const zipBuf = generateZip(payload.analysis, payload.selections);
      const name = (payload.analysis?.company?.name || 'IBP').replace(/[^a-zA-Z0-9]/g, '_').slice(0, 30);
      res.writeHead(200, {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="IBP_${name}_${new Date().toISOString().slice(0,10)}.zip"`,
        'Content-Length': zipBuf.length,
      });
      res.end(zipBuf);
    } catch (err) {
      console.error('[generate error]', err.message, err.stack);
      res.writeHead(500); res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  res.writeHead(404); res.end();
});

server.listen(PORT, () => {
  console.log(`\n  IBP Generator → http://localhost:${PORT}`);
  console.log(`  BP ZIP: ${Object.keys(bpData).length ? `${Object.keys(bpData).length} files loaded` : 'NOT FOUND — ' + BP_ZIP_PATH}\n`);
  if (!process.env.ANTHROPIC_API_KEY) console.warn('  ⚠  ANTHROPIC_API_KEY not set\n');
});
