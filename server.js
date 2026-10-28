const http = require('http');
const fs = require('fs');
const path = require('path');
const XLSX = require('./xlsx.full.min.js');

const PORT = 3000;
const FILE = path.join(__dirname, 'job-applications.xlsx');
const zlib = require('zlib');
const APP = {
  fields: ['id', 'company', 'role', 'status', 'applied', 'followUp', 'link', 'referral', 'referredBy', 'description', 'notes'],
  headers: ['ID', 'COMPANY', 'ROLE', 'STATUS', 'APPLIED', 'FOLLOW-UP', 'LINK', 'REFERRAL', 'REFERRED BY', 'DESCRIPTION', 'NOTES'],
  widths: [38, 25, 25, 12, 12, 12, 30, 10, 22, 60, 40],
};
const PORTAL = {
  fields: ['id', 'company', 'emailId', 'portalLink', 'loginType', 'username', 'password', 'notes'],
  headers: ['ID', 'COMPANY', 'EMAIL ID', 'PORTAL LINK', 'LOGIN TYPE', 'USERNAME', 'PASSWORD', 'NOTES'],
  widths: [38, 25, 30, 40, 12, 25, 25, 40],
};
const STATIC = { '/': 'index.html', '/index.html': 'index.html', '/xlsx.full.min.js': 'xlsx.full.min.js' };
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };

function readSheet(index, def) {
  if (!fs.existsSync(FILE)) return [];
  const wb = XLSX.read(fs.readFileSync(FILE));
  const ws = wb.Sheets[wb.SheetNames[index]];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: false }).map(r => {
    const norm = {};
    Object.keys(r).forEach(k => { norm[k.toLowerCase().replace(/[^a-z]/g, '')] = r[k]; });
    const o = {};
    def.fields.forEach(f => { o[f] = String(norm[f.toLowerCase()] ?? ''); });
    return o;
  }).filter(o => o.company || o.role);
}
const readApps = () => readSheet(0, APP);
const readPortals = () => readSheet(1, PORTAL);

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function zip(files) {
  const locals = [], centrals = [];
  let offset = 0;
  files.forEach(([name, text]) => {
    const nameBuf = Buffer.from(name);
    const raw = Buffer.from(text, 'utf8');
    const data = zlib.deflateRawSync(raw);
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(8, 8);
    local.writeUInt16LE(33, 12); local.writeUInt32LE(crc, 14); local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(raw.length, 22); local.writeUInt16LE(nameBuf.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
    central.writeUInt16LE(8, 10); central.writeUInt16LE(33, 14); central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20); central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuf.length, 28); central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, data);
    centrals.push(central, nameBuf);
    offset += 30 + nameBuf.length + data.length;
  });
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

const xmlEsc = s => String(s ?? '').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const cell = (col, row, text, style) =>
  `<c r="${String.fromCharCode(65 + col)}${row}" t="inlineStr"${style ? ' s="1"' : ''}><is><t xml:space="preserve">${xmlEsc(text)}</t></is></c>`;

function sheetXml(def, rows) {
  const cols = def.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
  const head = `<row r="1">${def.headers.map((h, i) => cell(i, 1, h, true)).join('')}</row>`;
  const body = rows.map((a, r) => `<row r="${r + 2}">${def.fields.map((f, i) => cell(i, r + 2, a[f])).join('')}</row>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${cols}</cols><sheetData>${head}${body}</sheetData></worksheet>`;
}

function buildXlsx(apps, portals) {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return zip([
    ['[Content_Types].xml', `${xml}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
    ['_rels/.rels', `${xml}<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `${xml}<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Applications" sheetId="1" r:id="rId1"/><sheet name="Portals" sheetId="2" r:id="rId3"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `${xml}<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`],
    ['xl/styles.xml', `${xml}<styleSheet xmlns="${main}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>`],
    ['xl/worksheets/sheet1.xml', sheetXml(APP, apps)],
    ['xl/worksheets/sheet2.xml', sheetXml(PORTAL, portals)],
  ]);
}

function writeAll(apps, portals) {
  const tmp = FILE + '.tmp';
  fs.writeFileSync(tmp, buildXlsx(apps, portals));
  fs.renameSync(tmp, FILE);
}

function send(res, code, body, type) {
  res.writeHead(code, { 'Content-Type': type || 'application/json' });
  res.end(body);
}

http.createServer((req, res) => {
  if (req.url === '/api/applications' && req.method === 'GET') {
    try { return send(res, 200, JSON.stringify(readApps())); }
    catch (e) { return send(res, 500, JSON.stringify({ error: e.message })); }
  }
  if (req.url === '/api/portals' && req.method === 'GET') {
    try { return send(res, 200, JSON.stringify(readPortals())); }
    catch (e) { return send(res, 500, JSON.stringify({ error: e.message })); }
  }
  if (req.url === '/api/download' && req.method === 'GET') {
    try {
      if (!fs.existsSync(FILE)) writeAll([], []);
      res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Disposition': 'attachment; filename="job-applications.xlsx"' });
      return res.end(fs.readFileSync(FILE));
    } catch (e) { return send(res, 500, JSON.stringify({ error: e.message })); }
  }
  if ((req.url === '/api/applications' || req.url === '/api/portals') && req.method === 'POST') {
    let body = '';
    req.on('data', c => { body += c; });
    req.on('end', () => {
      try {
        const data = JSON.parse(body);
        if (req.url === '/api/portals') writeAll(readApps(), data);
        else writeAll(data, readPortals());
        send(res, 200, '{"ok":true}');
      } catch (e) {
        send(res, 500, JSON.stringify({ error: e.message }));
      }
    });
    return;
  }
  const name = STATIC[req.url.split('?')[0]];
  if (name && req.method === 'GET') {
    return send(res, 200, fs.readFileSync(path.join(__dirname, name)), TYPES[path.extname(name)]);
  }
  send(res, 404, 'Not found', 'text/plain');
}).listen(PORT, '127.0.0.1', () => {
  try { writeAll(readApps(), readPortals()); } catch (e) { console.log('Could not refresh ' + FILE + ': ' + e.message); }
  console.log('Job Application Tracker running at http://localhost:' + PORT + '\nSaving to ' + FILE);
});
