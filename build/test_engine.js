// ทดสอบ engine ของแชทบอทแบบไม่ต้องเปิดเบราว์เซอร์
// รัน:  node build/test_engine.js     (จาก root ของ repo)
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const code = html.slice(html.lastIndexOf('<script>') + 8, html.lastIndexOf('</script>'));

const stubEl = () => {
  const e = {
    textContent: '', src: '', value: '', innerHTML: '', style: {},
    addEventListener() {}, appendChild() {}, remove() {}, focus() {}, select() {},
    setAttribute() {}, getAttribute() { return null; },
    scrollTop: 0, scrollHeight: 0, disabled: false,
  };
  e.querySelector = () => stubEl();
  e.querySelectorAll = () => [];
  e.closest = () => null;
  return e;
};
const sandbox = {
  console,
  setTimeout: () => {},
  matchMedia: () => ({ matches: false }),
  document: {
    getElementById: stubEl,
    createElement: () => { const d = stubEl(); d.firstElementChild = stubEl(); return d; },
    addEventListener() {},
    documentElement: stubEl(),
    body: stubEl(),
    execCommand: () => true,
  },
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
vm.runInContext(
  code + '\n;globalThis.__answer=answer;globalThis.__ROWS=ROWS;globalThis.__CODES=CODES;',
  sandbox);

const ask = q => sandbox.__answer(q);
const strip = s => String(s).replace(/<[^>]+>/g, '');

let failed = 0;
function check(name, cond, detail) {
  if (cond) return;
  failed++;
  console.log('FAIL  ' + name + (detail ? '  ->  ' + detail : ''));
}

// --- ข้อมูลโหลดครบ ---
check('โหลด 200 แถว', sandbox.__ROWS.length === 200, sandbox.__ROWS.length);
check('มี 105 รหัสแร็ค', sandbox.__CODES.length === 105, sandbox.__CODES.length);

// --- ตอบคำถามได้ถูกประเภท ---
const cases = [
  ['สวัสดี', a => !a.cards && /สวัสดี/.test(strip(a.headline))],
  ['ขอบคุณครับ', a => /ยินดี/.test(strip(a.headline))],
  ['แร็ค vigo ราคาเท่าไหร่', a => a.cards.length > 0 && /VIGO/.test(strip(a.headline))],
  ['ราคาแร็ควีโก้เท่าไหร่', a => a.cards.length > 0 && /VIGO/.test(strip(a.headline))],
  ['อยากรู้ราคาแร็ครถอัลติสครับ', a => /ALTIS/.test(strip(a.headline))],
  ['altis 2010', a => a.rowsCount === 2],
  ['RT01 ใช้กับรถอะไรได้บ้าง', a => a.cards.length === 1 && a.cards[0].code === 'RT01'],
  ['rt 1', a => a.cards[0].code === 'RT01'],
  ['ดีแม็ก 4wd', a => a.cards.every(c => c.rows.some(r => /D-MAX 4WD/.test(r.model)))],
  ['d-max 2wd 2015', a => a.rowsCount === 1],
  ['แร็คไฟฟ้าไม่เกิน 2800', a => a.cards.every(c => c.typeIdx === 1 && c.price <= 2800)],
  ['เครื่อง 2KD ใช้แร็คตัวไหน', a => a.rowsCount > 0],
  ['อีซูซุมีกี่รายการ', a => a.rowsCount === 20],
  ['มีกี่รายการทั้งหมด', a => a.rowsCount === 200],
  ['แร็คโตโยต้าถูกสุด', a => a.cards[0].price <= a.cards[a.cards.length - 1].price],
  ['ranger แพงสุด', a => a.cards[0].price >= a.cards[a.cards.length - 1].price],
  ['พวงมาลัยซ้ายมีอะไรบ้าง', a => a.cards.every(c => c.lhd)],
  ['แร็คเบนซ์ราคาเท่าไหร่', a => !a.cards && /ยังไม่เจอ/.test(strip(a.headline))],
  ['ฟหกด', a => !a.cards],
];
for (const [q, ok] of cases) {
  let a;
  try { a = ask(q); } catch (e) { check(q, false, 'throw: ' + e.message); continue; }
  check(q, ok(a), strip(a.headline).slice(0, 90));
}

// --- ห้ามมี HTML ดิบหลุดจากสิ่งที่ผู้ใช้พิมพ์ ---
// สองเคสแรกวิ่งเข้า branch หลัก (เจอผลลัพธ์) ซึ่งเป็นทางที่เคยหลุด HTML ดิบออกมาจริง
const injections = [
  'vigo <b>x</b>',
  'altis <i>y</i>',
  'vigo <img src=x onerror=alert(1)>',
  '<script>alert(1)</scr' + 'ipt>',
  'd-max <svg onload=alert(1)>',
];
for (const q of injections) {
  let a;
  try { a = ask(q); } catch (e) { check('injection ' + q, false, 'throw: ' + e.message); continue; }
  const blob = [a.headline, a.hint || ''].join(' ');
  // แท็กที่อนุญาตในคำตอบมีแค่ <b> ที่ระบบใส่เอง ส่วนที่มาจากผู้ใช้ต้องกลายเป็น &lt;
  const bad = /<(script|img|svg|iframe|object|i)\b/i.test(blob) || /<b>[^<]*<b>/i.test(blob);
  check('escape: ' + q, !bad, blob.slice(0, 120));
}
check('เคส injection วิ่งเข้า branch หลักจริง', ask('vigo <b>x</b>').rowsCount > 0);

// --- ราคาและประเภทต้องคงที่ต่อรหัส ---
const byCode = new Map();
for (const r of sandbox.__ROWS) {
  const k = r.brand + '|' + r.code;
  if (!byCode.has(k)) byCode.set(k, r);
  const f = byCode.get(k);
  check('รหัส ' + r.code + ' ราคาเดียว', f.price === r.price, f.price + ' vs ' + r.price);
  check('รหัส ' + r.code + ' ประเภทเดียว', f.typeIdx === r.typeIdx);
}

console.log(failed === 0
  ? 'PASS  ทดสอบผ่านทั้งหมด (' + sandbox.__ROWS.length + ' แถว / ' + sandbox.__CODES.length + ' รหัส)'
  : failed + ' เคสไม่ผ่าน');
process.exit(failed === 0 ? 0 : 1);
