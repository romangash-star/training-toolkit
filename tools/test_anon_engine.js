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
ok(det('תיאור הפנייה', ['פינוי גזם ברחוב']).type === 'freetext' && det('תיאור הפנייה', ['x']).action === 'keep', 'free text is kept (and cleaned)');
ok(det('הערות', ['בעיה עם הפח, לחזור ל 050-1112223 בערב']).type === 'freetext', 'free text with phone');
ok(det('מספר נכס', ['1002345']).type === 'property', 'property');
ok(det('תזמון', ['בוקר']).type === 'safe', '"תזמון" is not an id');

// --- rewriting: random but consistent numbers, empty stays empty, drop removes the column, free text cleaned
const sheet = { cols: [
  { idx: 0, name: 'שם פונה', type: 'person', action: 'token' },
  { idx: 1, name: 'טלפון', type: 'phone', action: 'token' },
  { idx: 2, name: 'קטגוריה', type: 'safe', action: 'keep' },
  { idx: 3, name: 'כתובת', type: 'address', action: 'drop' },
  { idx: 4, name: 'תיאור', type: 'freetext', action: 'keep' }],
  rows: [['דנה', '050-1234567', 'גזם', 'הגפן 3', 'חזרו אליי 050-1234567'], ['יוסי', '', 'תאורה', 'הזית 1', 'פנס שבור'], ['דנה', '052-1111111', 'גזם', 'הגפן 3', '']] };
const ctx = E.newContext(), out = E.applySheet(sheet, ctx);
ok(out.header.join('|') === 'שם פונה|טלפון|קטגוריה|תיאור', 'dropped column removed: ' + out.header);
const a0 = out.rows[0][0], a1 = out.rows[1][0], a2 = out.rows[2][0];
ok(a0 === 'תושב 001' && a1 === 'תושב 002', 'running numbers: ' + a0 + ', ' + a1);
ok(a0 === a2, 'same value, same number');
ok(out.rows[0][1] === 'טלפון 001', 'each group counts on its own: ' + out.rows[0][1]);
ok(out.rows[1][1] === '', 'empty stays empty');
ok(out.rows[0][3] === 'חזרו אליי [טלפון_הוסר]' && out.rows[1][3] === 'פנס שבור', 'free text kept and cleaned');
ok(out.rows[0][2] === 'גזם', 'safe column untouched');
ok(ctx.mapping.length === 4 && ctx.mapping[0].value === 'דנה' && ctx.mapping[0].token === a0, 'mapping file rows');
const again = E.applySheet(sheet, E.newContext());
ok(JSON.stringify(again.rows) === JSON.stringify(out.rows), 'same input, same numbers (preview matches download)');
// generalising, as on slide 74
ok(E.generalize('age', 37) === '35–44' && E.generalize('age', 64) === '55–64' && E.generalize('age', 3) === '0–4', 'age to range');
ok(E.generalize('date', new Date(2025, 2, 21)) === '03/2025' && E.generalize('date', '21.3.2025') === '03/2025', 'date to month');
ok(E.generalize('birth', new Date(1988, 5, 1), new Date(2026, 9, 26)) === '35–44', 'birth date to age range');
ok(E.generalize('address', 'הגפן 12, מגדל העמק') === 'הגפן, מגדל העמק' && E.generalize('address', 'רבי עקיבא 43/5') === 'רבי עקיבא', 'address to street');
ok(det('ת"ז', ['1']).idcol === true && det('שם פונה', ['דנה']).idcol === false, 'ID columns get their own numbers');
ok(E.applySheet({cols:[{idx:0,name:'ת"ז',type:'person',action:'token',idcol:true}],rows:[['123']]}, E.newContext()).rows[0][0] === 'ת"ז 001', 'ID running number');
ok(det('גיל', [37, 64]).action === 'generalize' && det('תאריך לידה', ['1.1.1980']).gen === 'birth' && det('כתובת', ['הגפן 3']).action === 'generalize', 'generalise is suggested for age, birth date and address');
ok(det('תאריך פתיחה', [new Date(2025, 1, 1), new Date(2025, 2, 1)]).gen === 'date' && det('תאריך פתיחה', [new Date(2025, 1, 1)]).action === 'keep', 'dates can be generalised but are kept by default');
const big = { cols: [{ idx: 0, name: 'x', type: 'person', action: 'token' }], rows: Array.from({ length: 5000 }, (_, i) => ['p' + i]) };
const nums = E.applySheet(big, E.newContext()).rows.map(r => r[0]);
ok(new Set(nums).size === 5000, '5000 different values get 5000 different numbers');

console.log(`${n - fail}/${n} checks passed`);
process.exit(fail ? 1 : 0);
