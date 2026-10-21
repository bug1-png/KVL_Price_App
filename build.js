// Usage: node build.js <path-containing-node_modules/xlsx>
const X = require(process.argv[2] + '/node_modules/xlsx');
const fs = require('fs');
const wb = X.readFile('../KVL_Price_App.xlsx');
const rd = n => X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
const clean = v => (v == null ? '' : typeof v === 'string' ? v.trim() : v);
const dt = v => {
  if (typeof v === 'number' && v > 40000) {
    const d = new Date(Math.round((v - 25569) * 864e5));
    return String(d.getUTCDate()).padStart(2, '0') + '/' + String(d.getUTCMonth() + 1).padStart(2, '0') + '/' + String(d.getUTCFullYear()).slice(2);
  }
  return clean(v);
};
const nz = v => (typeof v === 'number' ? v : null);
const z = v => { v = clean(v); return v === 0 || v === '0' ? '' : v; };

const master = rd(wb.SheetNames.find(n => /^Pricelist_.*\(Master\)/.test(n))).slice(2).filter(r => r[2]).map(r => ({
  no: clean(r[0]), open: clean(r[1]), unit: clean(r[2]), house: clean(r[3]), code: z(r[4]),
  tower: clean(r[5]), floor: clean(r[6]), type: clean(r[7]), dir: z(r[8]), area: nz(r[9]), view: z(r[10]),
  status: clean(r[11]), detail: clean(r[12]),
  price: nz(r[13]), cb: nz(r[14]), ww1: nz(r[15]), ww2: nz(r[16]), onhand: nz(r[17]), cbAll: nz(r[18]),
  remark: clean(r[20]), freeAll: nz(r[21]), plan: clean(r[22]), date: dt(r[23]),
  base2: nz(r[24]), variance: nz(r[25]), proTotal: nz(r[26]), chk: nz(r[27]),
}));
const hot = rd(wb.SheetNames.find(n => /^HotDeal_\d{6}/.test(n))).slice(1).filter(r => r[3]).map(r => ({
  unit: clean(r[3]), plan: clean(r[1]), date: dt(r[2]), tower: clean(r[4]), floor: clean(r[5]), type: clean(r[6]),
  dir: z(r[7]), area: nz(r[8]), view: z(r[9]), status: clean(r[10]), detail: clean(r[11]),
  price: nz(r[12]), disc: nz(r[13]), pro: nz(r[14]), ontop: clean(r[15]),
  masBase: nz(r[34]),
}));
const slim = o => { const x = {}; for (const k in o) if (o[k] !== '' && o[k] !== null) x[k] = o[k]; return x; };
const fmt = d => String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
const hm = /HotDeal_(\d\d)(\d\d)(\d\d)/.exec(wb.SheetNames.find(n => /^HotDeal_\d{6}/.test(n)) || '');
const meta = { asof: fmt(fs.statSync('../KVL_Price_App.xlsx').mtime), hot: hm ? hm[1] + '/' + hm[2] + '/' + hm[3] : '' };
const data = { meta,  master: master.map(slim), hot: hot.map(slim) };
const html = fs.readFileSync('template.html', 'utf8').replace('__DATA__', () => JSON.stringify(data));
fs.writeFileSync('index.html', html);
console.log(master.length, hot.length, (html.length / 1024).toFixed(0) + 'KB');
