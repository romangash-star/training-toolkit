// Tests for data/assets/anon-engine.js — run: node tools/test_anon_engine.js
const E = require('../data/assets/anon-engine.js');
let fail = 0, n = 0;
const ok = (cond, msg) => { n++; if (!cond) { fail++; console.log('FAIL', msg); } };

// --- Israeli ID check digit, against an independent implementation of the official algorithm
function reference(id) {
  const s = String(id).padStart(9, '0');
  let sum = 0;
  for (let i = 0; i < 9; i++) { let d = Number(s[i]) * (i % 2 === 0 ? 1 : 2); if (d > 9) d = Math.floor(d / 10) + (d % 10); sum += d; }
  return sum % 10 === 0;
}
// published sample IDs that are valid
['000000018', '123456782', '012345674', '039337423', '12345674'].forEach(id =>
  ok(E.isValidIsraeliId(id) === reference(id) && reference(id), 'valid id ' + id));
// one digit off is invalid
['123456789', '012345678', '039337424', '311111112'].forEach(id => ok(!E.isValidIsraeliId(id), 'invalid id ' + id));
// edge cases
ok(!E.isValidIsraeliId('000000000'), 'all zeros is not an id');
ok(!E.isValidIsraeliId('1234'), 'too short'); ok(!E.isValidIsraeliId('18'), 'two digits are not an id even if the check digit fits');
ok(!E.isValidIsraeliId('1234567890'), 'too long');
ok(!E.isValidIsraeliId('12a456782'), 'letters');
ok(E.isValidIsraeliId(' 123456782 '), 'surrounding spaces');
// 200,000 random numbers: same answer as the reference, and about 10% pass (one in ten check digits)
let pass = 0;
for (let i = 0; i < 200000; i++) {
  const x = String(Math.floor(Math.random() * 1e9)).padStart(9, '0');
  if (/^0+$/.test(x)) continue;
  const a = E.isValidIsraeliId(x), b = reference(x);
  if (a !== b) { ok(false, 'mismatch ' + x); break; }
  if (a) pass++;
}
ok(pass > 18000 && pass < 22000, 'about 10% of random numbers pass: ' + pass);
// generated ids are always valid
for (let i = 0; i < 1000; i++) ok(E.isValidIsraeliId(E.withCheckDigit(Math.floor(Math.random() * 1e8))), 'generated id valid');

// --- phones and emails
['050-1234567', '0501234567', '052 1234567', '02-6543210', '03-1234567', '+972-50-1234567', '+972501234567', '054-123-4567'].forEach(p => ok(E.isPhone(p), 'phone ' + p));
['1234567', '050-123456', '06-1234567', '12345678901'].forEach(p => ok(!E.isPhone(p), 'not phone ' + p));
ok(E.isEmail('dana.levi@gmail.com'), 'email');
ok(!E.isEmail('dana at gmail'), 'not email');

// --- free text scrubbing keeps the sentence
const c = { phone: 0, email: 0, id: 0 };
const t = E.scrubText('פנה התושב בטל 052-7654321, ת"ז 123456782, מייל dana@walla.co.il. פנייה מספר 4512345 לגבי פינוי גזם', c);
ok(t.includes('[טלפון_הוסר]') && t.includes('[תז_הוסרה]') && t.includes('[מייל_הוסר]'), 'scrubbed: ' + t);
ok(t.includes('4512345'), 'a 7-digit number that is not a valid id stays: ' + t);
ok(t.includes('פינוי גזם') && t.startsWith('פנה התושב'), 'sentence kept');
ok(c.phone === 1 && c.id === 1 && c.email === 1, 'counts ' + JSON.stringify(c));
ok(E.scrubText('תקציב 2025: 1500000 ש"ח') === 'תקציב 2025: 1500000 ש"ח', 'budget sentence untouched');

// --- column detection
const det = (h, vals) => E.detectColumn(h, vals);
ok(det('ת"ז', ['123456782']).type === 'person', 'id header');
ok(det('עמודה 1', Array.from({ length: 40 }, () => E.withCheckDigit(Math.floor(Math.random() * 1e8)))).type === 'person', 'ids found by content');
ok(det('מספר פנייה', Array.from({ length: 100 }, () => String(Math.floor(1e8 + Math.random() * 9e8)))).type === 'safe', 'random 9-digit ticket numbers are not ids');
ok(det('סכום', [1500, 2300, 870, 120000]).type === 'safe', 'amounts are safe');
ok(det('שם רחוב', ['הרצל']).type === 'address', 'street name is an address, not a person');
ok(det('שכונה', ['רמת השרון']).type === 'safe', 'neighbourhood is safe');
ok(det('עובד מטפל', ['משה']).type === 'employee', 'employee');
ok(det('טלפון נייד', ['050-1234567']).type === 'phone', 'phone header');
ok(det('תיאור הפנייה', ['פינוי גזם ברחוב']).action === 'scrub', 'free text scrub');
ok(det('הערות', ['בעיה עם הפח, לחזור ל 050-1112223 בערב']).type === 'freetext', 'free text with phone');
ok(det('מספר נכס', ['1002345']).type === 'property', 'property');
ok(det('תזמון', ['בוקר']).type === 'safe', '"תזמון" is not an id');

// --- rewriting: consistent tokens, empty stays empty, drop removes the column
const sheet = { cols: [
  { idx: 0, name: 'שם פונה', type: 'person', action: 'token' },
  { idx: 1, name: 'טלפון', type: 'phone', action: 'mask' },
  { idx: 2, name: 'קטגוריה', type: 'safe', action: 'keep' },
  { idx: 3, name: 'כתובת', type: 'address', action: 'drop' },
  { idx: 4, name: 'תיאור', type: 'freetext', action: 'scrub' }],
  rows: [['דנה', '050-1234567', 'גזם', 'הגפן 3', 'חזרו אליי 050-1234567'], ['יוסי', '', 'תאורה', 'הזית 1', 'פנס שבור'], ['דנה', '052-1111111', 'גזם', 'הגפן 3', '']] };
const ctx = E.newContext(), out = E.applySheet(sheet, ctx);
ok(out.header.join('|') === 'שם פונה|טלפון|קטגוריה|תיאור', 'dropped column removed: ' + out.header);
ok(out.rows[0][0] === 'תושב_001' && out.rows[1][0] === 'תושב_002' && out.rows[2][0] === 'תושב_001', 'consistent tokens');
ok(out.rows[0][1] === '[מוסתר]' && out.rows[1][1] === '', 'mask, empty stays empty');
ok(out.rows[0][3] === 'חזרו אליי [טלפון_הוסר]', 'scrub in sheet');
ok(ctx.mapping.length === 2 && ctx.mapping[0].value === 'דנה', 'mapping file rows');

console.log(`${n - fail}/${n} checks passed`);
process.exit(fail ? 1 : 0);
