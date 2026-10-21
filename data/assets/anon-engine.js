/* Anonymisation engine tool (data/anonymize).
   Pure functions, no DOM and no network: detects personal data in a table and rewrites it.
   Works in the browser (window.AnonEngine) and in Node (module.exports) for the tests. */
(function(root){
  'use strict';

  /* ---------- Israeli ID: 7–9 digits, left-padded to 9, official check digit ---------- */
  function isValidIsraeliId(raw){
    var s=String(raw==null?'':raw).trim();
    if(!/^\d{5,9}$/.test(s)) return false;
    s=('000000000'+s).slice(-9);
    if(/^0+$/.test(s)) return false;
    var sum=0;
    for(var i=0;i<9;i++){
      var d=(s.charCodeAt(i)-48)*((i%2)+1);
      sum+= d>9 ? d-9 : d;
    }
    return sum%10===0;
  }
  /* builds a valid ID from 8 digits (used by the demo data and the tests) */
  function withCheckDigit(eight){
    var s=('00000000'+eight).slice(-8), sum=0;
    for(var i=0;i<8;i++){ var d=(s.charCodeAt(i)-48)*((i%2)+1); sum+= d>9?d-9:d; }
    return s+String((10-sum%10)%10);
  }

  /* ---------- patterns ---------- */
  var PHONE_CORE='(?:\\+972[- ]?|0)(?:5\\d|[23489])[- ]?\\d{3}[- ]?\\d{4}';
  var PHONE_FULL=new RegExp('^'+PHONE_CORE+'$');
  var EMAIL_FULL=/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  function isPhone(v){ return PHONE_FULL.test(String(v).trim()); }
  function isEmail(v){ return EMAIL_FULL.test(String(v).trim()); }
  function isId(v){ return isValidIsraeliId(String(v).replace(/[\s-]/g,'')); }

  /* ---------- free text scrubbing: keeps the sentence, removes what identifies ---------- */
  function scrubText(text, counts){
    if(text==null||text==='') return text;
    var s=String(text);
    s=s.replace(/[^\s@<>()[\]"',;:]+@[^\s@<>()[\]"',;:]+\.[A-Za-z]{2,}/g,function(){ if(counts) counts.email++; return '[מייל_הוסר]'; });
    s=s.replace(new RegExp('(?<![\\d+])'+PHONE_CORE+'(?!\\d)','g'),function(){ if(counts) counts.phone++; return '[טלפון_הוסר]'; });
    s=s.replace(/(?<!\d)\d{7,9}(?!\d)/g,function(m){ if(isValidIsraeliId(m)){ if(counts) counts.id++; return '[תז_הוסרה]'; } return m; });
    return s;
  }
  function hasEmbeddedPii(text){
    var c={phone:0,email:0,id:0}; scrubText(text,c); return c.phone+c.email+c.id>0;
  }

  /* ---------- column types and what to suggest for each ---------- */
  var TYPES={
    person:   {label:'מזהה אישי (שם / ת״ז)', prefix:'תושב',  action:'token'},
    employee: {label:'שם עובד',              prefix:'עובד',  action:'token'},
    phone:    {label:'טלפון',                prefix:'טלפון', action:'token'},
    email:    {label:'דוא״ל',                prefix:'מייל',  action:'token'},
    address:  {label:'כתובת',                prefix:'כתובת', action:'token'},
    payer:    {label:'מספר משלם',            prefix:'משלם',  action:'token'},
    property: {label:'מספר נכס / חשבון',     prefix:'נכס',   action:'token'},
    caseno:   {label:'מספר תיק',             prefix:'תיק',   action:'token'},
    vehicle:  {label:'מספר רכב',             prefix:'רכב',   action:'token'},
    freetext: {label:'טקסט חופשי',           prefix:'טקסט',  action:'keep'},
    safe:     {label:'לא נמצא מידע מזהה',     prefix:'ערך',   action:'keep'}
  };
  /* checked in order: the more specific phrases first ("שם רחוב" is an address, not a name) */
  var HEADER_RULES=[
    ['freetext',['תוכן פנייה','תוכן הפנייה','פירוט','תיאור','הערות','הערה','טקסט חופשי','סיכום שיחה','מלל חופשי']],
    ['property',['מספר נכס','מס נכס','חשבון ארנונה','מספר חשבון']],
    ['payer',   ['מספר משלם','קוד משלם','מס משלם']],
    ['caseno',  ['מספר תיק','מס תיק']],
    ['vehicle', ['מספר רכב','לוחית רישוי','לוחית','מספר רישוי']],
    ['address', ['כתובת','שם רחוב','רחוב','מספר בית','דירה']],
    ['employee',['שם עובד','עובד מטפל','גורם מטפל','נציג','שם המטפל']],
    ['person',  ['ת.ז','ת"ז','תז','תעודת זהות','מספר זהות','ח.פ','ח"פ','דרכון','שם פרטי','שם משפחה','שם מלא','שם תושב','שם פונה','שם הפונה','שם הורה','שם תלמיד','שם משלם','שם המבקש']],
    ['phone',   ['טלפון','נייד','סלולרי','סלולארי','פקס','טל.']],
    ['email',   ['מייל','דוא"ל','דואל','אימייל','email','e-mail']]
  ];
  function normHeader(h){ return String(h==null?'':h).toLowerCase().replace(/[״“”]/g,'"').replace(/[׳’]/g,"'").replace(/\s+/g,' ').trim(); }
  function headerType(h){
    var n=normHeader(h);
    if(!n) return null;
    if(n==='שם'||n==='שם:') return {type:'person', why:'שם העמודה: "'+h+'"'};
    for(var r=0;r<HEADER_RULES.length;r++){
      var words=HEADER_RULES[r][1];
      for(var w=0;w<words.length;w++){
        var k=words[w].toLowerCase();
        /* short keywords (תז, טל.) only as a whole word, so "תזמון" is not an ID */
        var hit = k.length<=3 ? new RegExp('(^|[^\\u0590-\\u05FFa-z])'+k.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'($|[^\\u0590-\\u05FFa-z])').test(n) : n.indexOf(k)>-1;
        if(hit) return {type:HEADER_RULES[r][0], why:'שם העמודה: "'+h+'"'};
      }
    }
    return null;
  }

  /* looks at the header and at up to 100 non-empty values */
  function detectColumn(header, values){
    var sample=[], n=0;
    for(var i=0;i<values.length && sample.length<100;i++){ var v=values[i]; if(v!==''&&v!=null){ sample.push(v); } }
    n=sample.length;
    var byHead=headerType(header);
    var ids=0, phones=0, emails=0, embedded=0, longText=0;
    sample.forEach(function(v){
      var s=String(v instanceof Date ? '' : v).trim();
      if(isId(s)) ids++;
      else if(isPhone(s)) phones++;
      else if(isEmail(s)) emails++;
      if(s.length>15){ longText++; if(hasEmbeddedPii(s)) embedded++; }
    });
    var share=function(x){ return n ? x/n : 0; };
    var byContent=null;
    if(share(ids)>=0.15) byContent={type:'person', why:Math.round(share(ids)*100)+'% מהערכים הם מספרי ת״ז תקינים (כולל ספרת ביקורת)'};
    else if(share(phones)>=0.15) byContent={type:'phone', why:Math.round(share(phones)*100)+'% מהערכים הם מספרי טלפון'};
    else if(share(emails)>=0.15) byContent={type:'email', why:Math.round(share(emails)*100)+'% מהערכים הם כתובות דוא״ל'};
    else if(embedded>0 && share(longText)>=0.3) byContent={type:'freetext', why:'נמצאו פרטים אישיים בתוך הטקסט ב-'+embedded+' מתוך '+n+' תאים שנבדקו'};
    /* the header decides the type, unless the content clearly says something more specific */
    var pick = byHead || byContent || {type:'safe', why:n?'לא נמצאו שמות, ת״ז, טלפונים או דוא״ל':'העמודה ריקה'};
    if(byHead && byContent && byHead.type==='safe') pick=byContent;
    if(byHead && byContent && byContent.type!==byHead.type && byHead.type!=='freetext') pick={type:byHead.type, why:byHead.why+'; וגם '+byContent.why};
    if(byHead && byHead.type==='freetext' && embedded>0) pick={type:'freetext', why:byHead.why+'; נמצאו פרטים אישיים בתוך הטקסט'};
    return {type:pick.type, why:pick.why, action:TYPES[pick.type].action, samples:sample.slice(0,5).map(function(v){ return v instanceof Date ? v.toLocaleDateString('he-IL') : String(v); })};
  }

  /* ---------- rewriting ---------- */
  /* small seeded generator: the same file gives the same random numbers in the preview and in the download */
  function seeded(seed){ var t=seed>>>0; return function(){ t=(t+0x6D2B79F5)>>>0; var r=Math.imul(t^(t>>>15),1|t); r^=r+Math.imul(r^(r>>>7),61|r); return ((r^(r>>>14))>>>0)/4294967296; }; }
  function newContext(seed){ return {maps:{}, mapping:[], counts:{phone:0,email:0,id:0}, tokens:0, used:new Set(), rnd:seeded(seed==null?20261021:seed)}; }
  /* each distinct value gets its own random 6-digit number, the same one everywhere in the file.
     Six digits can never be a valid Israeli ID (those have 7–9), so the output is not mistaken for one. */
  function token(ctx, group, value, colName){
    var key=String(value).trim();
    var m=ctx.maps[group]||(ctx.maps[group]=new Map());
    if(!m.has(key)){
      var n; do{ n=100000+Math.floor(ctx.rnd()*900000); }while(ctx.used.has(n));
      ctx.used.add(n); m.set(key,n); ctx.tokens++;
      ctx.mapping.push({token:n, value:key, column:colName});
    }
    return m.get(key);
  }
  /* returns the new value for one cell (drop is handled by the caller) */
  function transform(col, v, ctx){
    if(v===''||v==null) return v;
    switch(col.action){
      case 'token': return token(ctx, TYPES[col.type] ? TYPES[col.type].prefix : 'ערך', v instanceof Date ? v.toISOString() : v, col.name);
      case 'keep':
        if(col.type!=='freetext' || v instanceof Date) return v;
        if(typeof v==='number') return isValidIsraeliId(String(v)) ? (ctx.counts.id++,'[תז_הוסרה]') : v;
        return scrubText(v, ctx.counts);
      default: return v;
    }
  }
  /* applies the column decisions to a whole sheet: {header, rows} -> {header, rows} */
  function applySheet(sheet, ctx, limit){
    var keep=sheet.cols.filter(function(c){ return c.action!=='drop'; });
    var rows=[], max=limit==null?sheet.rows.length:Math.min(limit,sheet.rows.length);
    for(var r=0;r<max;r++){
      var src=sheet.rows[r];
      rows.push(keep.map(function(c){ return transform(c, src[c.idx], ctx); }));
    }
    return {header:keep.map(function(c){ return c.name; }), rows:rows, kept:keep};
  }

  var api={isValidIsraeliId:isValidIsraeliId, withCheckDigit:withCheckDigit, isPhone:isPhone, isEmail:isEmail, scrubText:scrubText,
           headerType:headerType, detectColumn:detectColumn, newContext:newContext, transform:transform, applySheet:applySheet, TYPES:TYPES};
  if(typeof module!=='undefined' && module.exports) module.exports=api; else root.AnonEngine=api;
})(this);
