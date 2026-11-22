/* "עבודה עם נתונים בעזרת בינה מלאכותית": the section page and the five round pages.
   The facts about each practice file (columns, formulas, expected answers) come from files.js,
   which tools/dataai_build.py generates and checks in a real Excel. The wording of the tasks lives here.
   The deck shows the same tasks: when a task changes here, change the slide too. */
(function(){
'use strict';
var F=window.DATAAI_FILES, KC='dataai_choice', KI='toolkit_identity';
var ROUND=+document.body.dataset.round||0, PAGE=document.body.dataset.page||(ROUND?'r'+ROUND:''), UP=PAGE&&PAGE!=='id'?'../':'';
var KD='dataai_done', GEMINI='https://gemini.google.com/', OPEN='https://data.gov.il/he/organizations/beer-sheva';
var LEVELS=[['b','בסיסית'],['m','בינונית'],['a','מתקדמת']];
var TOOLS=[['gapp','Gemini'],['copilot','Copilot ב-Excel'],['gemini','Gemini ב-Sheets'],['chat','צ\'אט חינמי עם העלאת קובץ']];
/* the four work parts of the deck, in its order (div-a, div-b, div-c, div-d). Names and times are the deck's. */
var ROUNDS=[null,
  {t:'עוזר אישי לנוסחאות', s:'מתארים את העמודות, בלי הנתונים, ומקבלים נוסחה. גם בלי מנוי בתשלום.', min:10, tool:'Gemini או Copilot בצ\'אט'},
  {t:'דשבורד ב-Gemini Canvas', s:'מקובץ נתונים לדשבורד, בשיחה.', min:20, tool:'Gemini Canvas'},
  {t:'ניתוח נתונים בתוך הגיליון', s:'ניתוח, סידור ודשבורד בתוך הגיליון. אין מנוי? אותם פרומפטים בצ\'אט, עם העלאת קובץ.', min:25, tool:'Copilot ב-Excel · Gemini ב-Sheets'},
  {t:'NotebookLM', s:'מחברת שעונה מתוך הקבצים שלכם, עם הפניה למקור. במסך תראו את השם החדש שלה: Gemini Notebook.', min:15, tool:'Gemini Notebook'}];
/* the deck's slide "rules" */
var RULES=['מתממים לפני שמעלים','נוסחה, ולא תשובה','בודקים מול המקור','מתעקשים: התשובה הראשונה היא טיוטה','כלי מתאים לכל משימה'];
var AGG={sum:'סכום',avg:'ממוצע',count:'ספירה'};
/* The day's stages, numbered, coloured and shaped like the deck's section slides
   (div-work = 01, div-a = 02, div-b = 03, div-c = 04, div-d = 05, div-e = 06).
   The deck shows the same number, colour and shape: change both together. Moving on is the trainer's call, from the deck,
   so the pages have no "next" buttons; the bar at the top lets anyone go to any stage. */
var STAGES=[
  {n:'00',t:'לפני שמתחילים',c:'#5A4636',sh:'triangle',k:'id',h:''},
  {n:'01',t:'עכשיו עובדים',c:'#1F5F5B',sh:'circle',h:'file/',subs:[{k:'file',t:'הקובץ',h:'file/'},{k:'anon',t:'התממה',h:'anonymize/',opt:1}]},
  {n:'02',t:ROUNDS[1].t,c:'#B4472A',sh:'square',k:'r1',h:'r1/'},
  {n:'03',t:ROUNDS[2].t,c:'#2B579A',sh:'diamond',k:'r2',h:'r2/'},
  {n:'04',t:ROUNDS[3].t,c:'#6B4594',sh:'hexagon',k:'r3',h:'r3/'},
  {n:'05',t:ROUNDS[4].t,c:'#8A6500',sh:'pentagon',k:'r4',h:'r4/'},
  {n:'06',t:'מקורות מידע פתוחים',c:'#2E6B2C',sh:'leaf',k:'open',h:'open/'}];
/* the deck's slide "start" */
var LVTXT={b:'שאלות בשפה חופשית, פרומפטים מוכנים, צעד אחר צעד',
           m:'עמודות חדשות, טבלאות ציר, בדיקה מדגמית',
           a:'תהליכים שחוזרים, נוסחאות מערך, מוצר'};
/* the deck's slide "files": which file suits which department */
var FOR={'06':'גבייה, חשבונות, כספים, חברה כלכלית','05':'ועדה מקומית והנדסה','08':'פיקוח','02':'חינוך וגנים','09':'משאבי אנוש','07':'דוברות','01':'מזכירות, דיגיטל, תכנון, GIS'};
var APP={gapp:'Excel',copilot:'Excel',gemini:'Google Sheets',chat:'Excel'};

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function num(v,pct){
  if(typeof v!=='number') return esc(v);
  if(pct) return (v*100).toLocaleString('he-IL',{maximumFractionDigits:1})+'%';
  return v.toLocaleString('he-IL',{maximumFractionDigits:Math.abs(v)<100&&v%1?2:(v%1?1:0)});
}
function isPct(name){ return /אחוז|שיעור/.test(name||''); }
function choice(){
  var c={}; try{ c=JSON.parse(localStorage.getItem(KC)||'{}')||{}; }catch(e){}
  var ok=!!(F[c.file]&&c.level);
  return {file:F[c.file]?c.file:'01',
          level:LEVELS.some(function(l){return l[0]===c.level;})?c.level:'b',
          tool:TOOLS.some(function(t){return t[0]===c.tool;})?c.tool:'gapp',
          src:/^(own|open)$/.test(c.src)?c.src:'practice', set:ok};
}
function save(c){ try{ localStorage.setItem(KC,JSON.stringify({file:c.file,level:c.level,tool:c.tool,src:c.src})); }catch(e){} }
function done(){ try{ return JSON.parse(localStorage.getItem(KD)||'{}')||{}; }catch(e){ return {}; } }
function mark(k){ var d=done(); d[k]=1; try{ localStorage.setItem(KD,JSON.stringify(d)); }catch(e){} }
var KS='dataai_me';
function ident(){ try{ return JSON.parse(sessionStorage.getItem(KS)||'{}')||{}; }catch(e){ return {}; } }
function stageOf(k){ for(var i=0;i<STAGES.length;i++){ var S=STAGES[i]; if(S.k===k||(S.subs&&S.subs.some(function(x){ return x.k===k; }))) return i; } return -1; }
function badge(S,cls){ return '<b class="sg sg-'+S.sh+(cls?' '+cls:'')+'" style="--sc:'+S.c+'" aria-hidden="true"><em>'+S.n+'</em></b>'; }
function stepper(cur){
  var ci=stageOf(cur), S=STAGES[ci];
  var h='<nav class="stages" aria-label="שלבי היום" style="--n:'+STAGES.length+'">'+STAGES.map(function(x,i){
    return '<a'+(i===ci?' class="cur" aria-current="step"':'')+' href="'+UP+x.h+'" style="--sc:'+x.c+'">'+badge(x)+'<span>'+esc(x.t)+'</span></a>'; }).join('')+'</nav>';
  if(!S) return h;
  h+='<div class="stagehead" style="--sc:'+S.c+'">'+badge(S,'big')+'<div class="st-txt"><small>שלב '+S.n+'</small>'+(S.subs?'<p class="st-t">':'<h1 class="st-t">')+esc(S.t)+(S.subs?'</p>':'</h1>')+'</div>'+
    (S.subs?'<div class="subs">'+S.subs.map(function(x){ return '<a href="'+UP+x.h+'"'+(x.k===cur?' aria-current="page"':'')+'>'+esc(x.t)+(x.opt?' <i>לא חובה</i>':'')+'</a>'; }).join('')+'</div>':'')+(cur==='id'?'':'<a class="pbank-in" href="'+UP+'prompts/">מאגר הפרומפטים</a>')+'</div>';
  return h;
}
function nextCard(cur){
  var i=stageOf(cur), N=STAGES[i+1]; if(!N) return '';
  return '<div class="stagenext" style="--sc:'+N.c+'"><p><b>סיימתם את השלב?</b> ממשיכים כשנגיע במצגת לשלב '+N.n+':</p>'+
    '<a href="'+UP+N.h+'">'+badge(N)+'<span>'+esc(N.t)+'</span></a></div>';
}
function label(list,k){ for(var i=0;i<list.length;i++) if(list[i][0]===k) return list[i][1]; return ''; }
function raw(f){ return UP+'../data/practice-files/files/'+f.n+'.xlsx'; }
function dl(href,name,text){ return '<a class="btn-dl" href="'+href+'" download="'+esc(name)+'">'+text+'</a>'; }
function fx(f){ return '<div class="fx"><code>'+esc(f)+'</code><button type="button" class="copy" onclick="cpF(this)">העתקה</button></div>'; }
function prompt(l,t){ return '<div class="prompt"><div class="ph"><span>'+esc(l)+'</span><button type="button" class="copy" onclick="copyPrompt(this)">העתקת הפרומפט</button></div><pre>'+esc(t)+'</pre></div>'; }

function put(text,btn){
  function done(){ var o=btn.textContent; btn.classList.add('ok'); btn.textContent='הועתק ✓'; setTimeout(function(){ btn.classList.remove('ok'); btn.textContent=o; },1600); }
  function old(){ var ta=document.createElement('textarea'); ta.value=text; ta.style.position='fixed'; ta.style.opacity='0'; document.body.appendChild(ta); ta.select(); try{ document.execCommand('copy'); }catch(e){} ta.remove(); done(); }
  if(navigator.clipboard&&navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done,old); else old();
}
window.copyPrompt=function(b){ put(b.closest('.prompt').querySelector('pre').textContent,b); };
window.cpF=function(b){ put(b.parentNode.querySelector('code').textContent,b); };

/* ---------- stage 00: before we start. Three questions, each person for themselves. Nothing is kept: the page warns before
   it is closed, and everything can be exported at any moment, as a picture (drawn on a canvas, no library) or a Word file. ---------- */
function home(){
  var app=document.getElementById('app');
  var Q=[['me_name','שם','לא חובה','input'],
         ['me_role','1. מה התפקיד שלי ברשות?','תפקיד, מחלקה ורשות','input'],
         ['me_big','2. משימות גדולות ומשמעותיות','פרויקטים, דוחות תקופתיים, החלטות. שורה לכל משימה.','area'],
         ['me_small','3. משימות יומיומיות וחזרתיות','מה חוזר כל יום או כל שבוע. שורה לכל משימה.','area'],
         ['me_data','4. באילו נתונים אני משתמש?','קבצים, מערכות, דוחות וטבלאות. שורה לכל אחד.','area']];
  var h=stepper('id')+'<div class="rhead"><p>כל אחד לעצמו: רשימה אישית שממנה בוחרים על מה עובדים היום.</p></div>';
  h+='<p class="priv"><b>המידע לא נשמר ולא נשלח לשום מקום.</b> רוצים לשמור? מייצאים לתמונה או ל-Word.</p>';
  h+='<section class="stepcard meform">'+Q.map(function(q){
    return '<div class="mq'+(q[3]==='input'?' short':'')+'"><label for="'+q[0]+'">'+q[1]+' <small>'+q[2]+'</small></label>'+
      (q[3]==='input'?'<input type="text" id="'+q[0]+'" autocomplete="off">':'<textarea id="'+q[0]+'" rows="5"></textarea>')+'</div>'; }).join('')+
    '<div class="exbar"><span>ייצוא:</span><button type="button" id="exImg">תמונה</button><button type="button" id="exDoc">קובץ Word</button><span class="exmsg" id="exMsg" role="status"></span></div></section>';
  h+=nextCard('id');
  app.innerHTML=h;
  function val(id){ return (document.getElementById(id).value||'').trim(); }
  function lines(id){ return val(id).split(/\n+/).map(function(x){ return x.replace(/^[\s•*\-–]+/,'').trim(); }).filter(Boolean); }
  function data(){ return {name:val('me_name'),role:val('me_role'),parts:[['משימות גדולות ומשמעותיות',lines('me_big')],['משימות יומיומיות וחזרתיות',lines('me_small')],['הנתונים שאני משתמש בהם',lines('me_data')]]}; }
  function empty(d){ return !d.name&&!d.role&&d.parts.every(function(p){ return !p[1].length; }); }
  function today(){ try{ return new Date().toLocaleDateString('he-IL'); }catch(e){ return ''; } }
  function fname(d,ext){ return 'מה אני עושה ברשות'+(d.name?' - '+d.name.replace(/[\\\/:*?"<>|]/g,''):'')+'.'+ext; }
  function save(blob,name){ var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click(); setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); },1500); }
  function msg(t){ var m=document.getElementById('exMsg'); m.textContent=t; setTimeout(function(){ if(m.textContent===t) m.textContent=''; },3000); }
  var dirty=false;
  [].forEach.call(app.querySelectorAll('.meform input,.meform textarea'),function(x){ x.addEventListener('input',function(){ dirty=true; }); });
  window.addEventListener('beforeunload',function(e){ if(dirty&&!empty(data())){ e.preventDefault(); e.returnValue=''; } });

  document.getElementById('exDoc').onclick=function(){
    var d=data(); if(empty(d)){ msg('עוד לא כתבתם כלום.'); return; }
    var body='<h1>מה אני עושה ברשות?</h1>'+(d.name?'<p><b>'+esc(d.name)+'</b></p>':'')+(d.role?'<p>'+esc(d.role)+'</p>':'')+
      d.parts.map(function(p){ return '<h2>'+p[0]+'</h2>'+(p[1].length?'<ul>'+p[1].map(function(x){ return '<li>'+esc(x)+'</li>'; }).join('')+'</ul>':'<p>-</p>'); }).join('')+
      '<p style="color:#777;font-size:10pt">עבודה עם נתונים בעזרת בינה מלאכותית · '+today()+'</p>';
    var doc='<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>התפקיד שלי</title>'+
      '<style>body{font-family:Arial,sans-serif;direction:rtl;text-align:right;font-size:12pt}h1{font-size:18pt;color:#1F5F5B}h2{font-size:14pt;margin-top:14pt}</style></head><body dir="rtl" lang="he">'+body+'</body></html>';
    save(new Blob(['﻿',doc],{type:'application/msword'}),fname(d,'doc')); dirty=false; msg('הקובץ ירד ✓');
  };

  document.getElementById('exImg').onclick=function(){
    var d=data(); if(empty(d)){ msg('עוד לא כתבתם כלום.'); return; }
    var W=1200, P=70, F='Alef, Arial, sans-serif';
    function draw(ctx,measureOnly){
      var y=P; ctx.direction='rtl'; ctx.textAlign='right'; ctx.textBaseline='top';
      function wrap(t,size,weight,color,gap,indent){
        ctx.font=weight+' '+size+'px '+F; ctx.fillStyle=color;
        var words=t.split(/\s+/), line='', max=W-2*P-(indent||0), out=[];
        words.forEach(function(w){ var tst=line?line+' '+w:w; if(ctx.measureText(tst).width>max&&line){ out.push(line); line=w; } else line=tst; });
        if(line) out.push(line);
        out.forEach(function(l){ if(!measureOnly) ctx.fillText(l,W-P-(indent||0),y); y+=size*1.35; });
        y+=gap||0;
      }
      if(!measureOnly){ ctx.fillStyle='#1F5F5B'; ctx.fillRect(0,0,W,14); }
      y+=10;
      wrap('מה אני עושה ברשות?',44,'700','#2a211b',14);
      if(d.name) wrap(d.name,30,'700','#1F5F5B',2);
      if(d.role) wrap(d.role,28,'400','#5c4f45',10);
      d.parts.forEach(function(p){
        y+=18; wrap(p[0],32,'700','#B4472A',8);
        (p[1].length?p[1]:['-']).forEach(function(x){
          if(!measureOnly){ ctx.fillStyle='#B4472A'; ctx.beginPath(); ctx.arc(W-P-8,y+16,5,0,7); ctx.fill(); }
          wrap(x,27,'400','#2a211b',6,30); });
      });
      y+=30; wrap('עבודה עם נתונים בעזרת בינה מלאכותית · '+today(),22,'400','#8a7d72',0);
      return y+P;
    }
    var go=function(){
      var c=document.createElement('canvas'), ctx=c.getContext('2d'); c.width=W; c.height=10;
      var H=Math.ceil(draw(ctx,true)); c.width=W; c.height=H; ctx=c.getContext('2d');
      ctx.fillStyle='#faf5ec'; ctx.fillRect(0,0,W,H); draw(ctx,false);
      c.toBlob(function(b){ if(b){ save(b,fname(d,'png')); dirty=false; msg('התמונה ירדה ✓'); } else msg('הייצוא לא הצליח. נסו את Word.'); },'image/png');
    };
    if(document.fonts&&document.fonts.load) document.fonts.load('700 30px Alef').then(go,go); else go();
  };
}

/* ---------- the prompt bank (/data-ai/prompts/): content in prompts.js ---------- */
function bankTool(){ try{ var t=localStorage.getItem('dataai_tool'); return t==='copilot'||t==='gapp'?t:''; }catch(e){ return ''; } }
window.cpP=function(b){ put(b.closest('.bk-body').querySelector('pre').textContent,b); };
function promptsPage(){
  var app=document.getElementById('app'), B=window.DATAAI_PROMPTS;
  function pick(s,t){ return s.replace(/\{\{([\s\S]*?)\|\|([\s\S]*?)\}\}/g,function(m,g,c){ return t==='copilot'?c:g; }); }
  function marks(s){ return esc(s).replace(/\[[^\]\n]+\]/g,function(m){ return '<mark>'+m+'</mark>'; }); }
  function two(n){ return (n<10?'0':'')+n; }
  function render(){
    var t=bankTool(), T=t==='copilot'?'Copilot':'Gemini';
    var h='<header class="bk-h"><p class="bk-kicker">עבודה עם נתונים בעזרת בינה מלאכותית</p><h1>מאגר הפרומפטים</h1>'+
      '<p class="bk-lead">'+B.items.length+' פרומפטים שעובדים, לפי שלבי העבודה. בוחרים פרומפט, ממלאים את מה ש<mark>מסומן</mark>, ומעתיקים.</p></header>';
    h+='<div class="bk-tool'+(t?'':' ask')+'" role="group" aria-label="הכלי שלי"><span>'+(t?'מותאם ל':'עם איזה כלי אתם עובדים?')+'</span>'+
      [['gapp','Gemini'],['copilot','Copilot']].map(function(o){ return '<button type="button" data-t="'+o[0]+'" aria-pressed="'+(t===o[0])+'">'+o[1]+'</button>'; }).join('')+'</div>';
    if(t){
      h+=B.stages.map(function(st,si){
        var items=B.items.filter(function(p){ return p.stage===st.k; }); if(!items.length) return '';
        return '<section class="bk-st"><h2>'+esc(st.t)+'</h2>'+items.map(function(p){
          return '<details class="bk-it" name="bk"><summary><b>'+two(p.id)+'</b><span><strong>'+esc(p.title)+'</strong><small>'+esc(p.why)+'</small></span><i aria-hidden="true"></i></summary>'+
            '<div class="bk-body"><p class="bk-where"><b>איפה זה עובד הכי טוב ב-'+T+':</b> '+esc(p.where[t==='copilot'?'c':'g'])+'</p>'+
            '<pre>'+marks(pick(p.text,t))+'</pre>'+
            '<button type="button" class="bk-copy" onclick="cpP(this)">העתקת הפרומפט</button></div></details>';
        }).join('')+'</section>';
      }).join('');
      h+='<p class="bk-foot">קובץ אמיתי מהעבודה? <a class="ln" href="../../data/anonymize/">מתממים</a> לפני שמעלים. ובודקים מספר אחד מול הגיליון לפני שממשיכים.</p>';
    }
    var y=window.scrollY; app.innerHTML=h; window.scrollTo(0,y);
    [].forEach.call(app.querySelectorAll('[data-t]'),function(b){ b.onclick=function(){
      try{ localStorage.setItem('dataai_tool',b.dataset.t); }catch(e){} var c=choice(); c.tool=b.dataset.t; save(c); render(); }; });
    var all=[].slice.call(app.querySelectorAll('.bk-it'));
    all.forEach(function(d){ d.addEventListener('toggle',function(){ if(d.open) all.forEach(function(o){ if(o!==d) o.open=false; }); }); });
  }
  render();
}

/* ---------- step 2: the file ---------- */
function filePage(){
  var app=document.getElementById('app'), P=window.DATAAI_PREVIEW||{};
  var raw0={}; try{ raw0=JSON.parse(localStorage.getItem(KC)||'{}')||{}; }catch(e){}
  var st={src:/^(own|open|practice)$/.test(raw0.src)?raw0.src:'', file:F[raw0.file]?raw0.file:'', level:/^[bma]$/.test(raw0.level||'')?raw0.level:'', tool:choice().tool};
  function keep(){ try{ localStorage.setItem(KC,JSON.stringify({file:st.file||'01',level:st.level||undefined,tool:st.tool,src:st.src||'practice'})); }catch(e){} }
  function preview(f){
    var p=P[f.n]; if(!p) return '';
    return '<div class="pv"><div class="pv-h"><div><h3>קובץ '+f.n+': '+esc(f.short)+'</h3><p>'+num(f.st.raw)+' שורות · '+f.head.length+' עמודות · נתונים פיקטיביים</p></div>'+
      dl(raw(f),f.n+' '+f.short+' - נתונים פיקטיביים.xlsx','הורדת הקובץ ←')+'</div>'+
      '<p class="pv-q"><b>השאלה שמנהל שואל על הקובץ הזה:</b> '+esc(f.ask)+'</p>'+
      '<div class="pv-t" tabindex="0" role="region" aria-label="תצוגה מקדימה של הקובץ"><table><thead><tr>'+p.head.map(function(c){ return '<th>'+esc(c)+'</th>'; }).join('')+'</tr></thead><tbody>'+
      p.rows.map(function(r){ return '<tr>'+r.map(function(c){ return '<td>'+esc(c)+'</td>'; }).join('')+'</tr>'; }).join('')+'</tbody></table></div>'+
      '<p class="out">תצוגה מקדימה: '+p.rows.length+' השורות הראשונות. תאים ריקים וכתיבים שונים הם חלק מהתרגיל.</p></div>';
  }
  function grid(){
    return '<div class="fgrid">'+Object.keys(F).sort().map(function(n){ var f=F[n];
      return '<button type="button" class="fcard" data-file="'+n+'" aria-pressed="'+(st.file===n)+'"><b>'+n+'</b><strong>'+esc(f.short)+'</strong><small>'+esc(FOR[n]||f.who)+'</small></button>'; }).join('')+'</div>';
  }
  function render(){
    var f=F[st.file], prac=st.src==='practice';
    var h=stepper('file')+'<div class="rhead"><h1>בוחרים קובץ</h1><p>קובץ אחד ילווה אתכם בכל שלבי העבודה. העובדים על אותו קובץ יושבים יחד.</p></div>';
    h+='<section class="stepcard"><div class="sc-h"><span class="sc-n">א</span><div><h2>על איזה קובץ עובדים?</h2></div></div><div class="srcs">'+
      [['practice','קובץ תרגול','עשרה קבצים פיקטיביים מעולם הרשות. רואים אותם כאן לפני שבוחרים.'],
       ['own','קובץ שלי מהעבודה','קודם מתממים, ורק אז מעלים ל-AI.'],
       ['open','נתונים פתוחים של באר שבע','מאגרים ציבוריים של העירייה, מאתר data.gov.il.']].map(function(o){
        return '<button type="button" class="src" data-src="'+o[0]+'" aria-pressed="'+(st.src===o[0])+'"><strong>'+o[1]+'</strong><span>'+o[2]+'</span></button>'; }).join('')+'</div>';
    if(prac) h+=grid()+(f?preview(f):'<p class="out pad">לחצו על קובץ כדי לראות אותו.</p>');
    if(st.src==='own') h+='<ol class="how"><li><b>מתממים את הקובץ.</b> מסירים שמות, תעודות זהות, טלפונים וכתובות, לפני שהקובץ מגיע ל-AI.<br><a class="btn-dl" href="../../data/anonymize/" target="_blank" rel="noopener">לכלי ההתממה ←</a></li>'+
      '<li><b>שומרים את הקובץ המותמם</b> במחשב, וחוזרים לכאן.</li></ol>';
    if(st.src==='open') h+='<ol class="how"><li><b>בוחרים מאגר ומורידים אותו.</b> זה מידע שהעירייה פרסמה לציבור.<br><a class="btn-dl" href="'+OPEN+'" target="_blank" rel="noopener">למאגרים של באר שבע ב-data.gov.il ←</a></li>'+
      '<li><b>פותחים את הקובץ</b> ובודקים שיש בו טבלה אחת, עם שורת כותרות.</li></ol>';
    if(st.src&&!prac) h+='<div class="exfile"><label for="exf">המשימות זהות לכולם. הדוגמאות והפתרונות יהיו מקובץ התרגול הזה:</label><select id="exf">'+
      Object.keys(F).sort().map(function(n){ return '<option value="'+n+'"'+((st.file||'01')===n?' selected':'')+'>'+n+' · '+esc(F[n].short)+'</option>'; }).join('')+'</select></div>';
    h+='</section>';
    h+='<section class="stepcard"><div class="sc-h"><span class="sc-n">ב</span><div><h2>באיזו רמה מתחילים?</h2><p>אפשר לעלות או לרדת רמה בכל שלב.</p></div></div><div class="lvs">'+
      LEVELS.map(function(l){ return '<button type="button" class="lvc lv-'+l[0]+'" data-level="'+l[0]+'" aria-pressed="'+(st.level===l[0])+'"><strong>'+l[1]+'</strong><span>'+LVTXT[l[0]]+'</span></button>'; }).join('')+'</div></section>';
    h+='<p class="gem">בכל שלב עובדים בכלי אחר. הכלי כתוב בראש השלב.</p>';
    h+=st.src==='own'?'<div class="pn"><a class="nx" id="go" href="../anonymize/">המשך: התממה ←</a></div>':nextCard('file');
    var y=window.scrollY; app.innerHTML=h; window.scrollTo(0,y);
    [].forEach.call(app.querySelectorAll('[data-src]'),function(b){ b.onclick=function(){ st.src=b.dataset.src; keep(); render(); }; });
    [].forEach.call(app.querySelectorAll('[data-file]'),function(b){ b.onclick=function(){ st.file=b.dataset.file; keep(); render(); var pv=app.querySelector('.pv'); if(pv&&pv.getBoundingClientRect().top>window.innerHeight-160) pv.scrollIntoView({behavior:'smooth',block:'nearest'}); }; });
    [].forEach.call(app.querySelectorAll('[data-level]'),function(b){ b.onclick=function(){ st.level=b.dataset.level; keep(); render(); }; });
    var ex=document.getElementById('exf'); if(ex) ex.onchange=function(){ st.file=ex.value; keep(); };
    var go=document.getElementById('go'); if(go) go.onclick=function(){ if(!st.file) st.file='01'; if(!st.level) st.level='b'; keep(); if(st.src) mark('file'); };
  }
  render();
}

/* ---------- step 3: anonymization. The tool itself is /data/anonymize/, shown here inside the page (?embed=1 hides its own header). ---------- */
function anonPage(){
  var app=document.getElementById('app'), raw0={}; try{ raw0=JSON.parse(localStorage.getItem(KC)||'{}')||{}; }catch(e){}
  var own=raw0.src==='own';
  var h=stepper('anon')+'<div class="rhead"><h1>התממה <span class="opt">לא חובה</span></h1>'+
    '<p>רק למי שעובד על קובץ מהעבודה. נתונים פתוחים וקובצי התרגול לא צריכים התממה.</p></div>';
  h+='<details class="stepcard anon-tool"'+(own?' open':'')+'><summary>'+(own?'כלי ההתממה':'רוצים לתרגל התממה? פתיחת הכלי')+'</summary>'+
     '<p class="out">מסירים שמות, תעודות זהות, טלפונים וכתובות. הקובץ לא נשלח לשום מקום.</p>'+
     '<iframe id="anonf" title="כלי ההתממה"></iframe>'+
     '<p class="out">הכלי לא נטען? <a class="ln" href="../../data/anonymize/" target="_blank" rel="noopener">פתיחה בעמוד מלא ←</a></p></details>'+nextCard('anon');
  app.innerHTML=h;
  var fr=document.getElementById('anonf'), det=app.querySelector('.anon-tool'), last=0;
  /* fit the frame to its content. Measure the content itself, never the frame's viewport, so it cannot grow forever. */
  function fit(){ try{ var d=fr.contentDocument, m=d&&(d.querySelector('main')||d.body); if(!m) return;
    var hgt=Math.ceil(m.getBoundingClientRect().bottom+(d.defaultView.scrollY||0))+16;
    if(Math.abs(hgt-last)>4){ last=Math.max(360,hgt); fr.style.height=last+'px'; } }catch(e){} }
  function load(){ if(fr.src) return; fr.src='../../data/anonymize/?embed=1';
    fr.addEventListener('load',function(){ fit(); try{ new ResizeObserver(fit).observe(fr.contentDocument.querySelector('main')||fr.contentDocument.body); }catch(e){ setInterval(fit,1000); } }); }
  if(det.open) load(); det.addEventListener('toggle',function(){ if(det.open) load(); });
}

/* ---------- round pages: the four work parts of the deck (div-a, div-b, div-c, div-d). Prompts are the deck's wording,
   with the practice file's real column names put in where the deck has an example or a [placeholder]. ---------- */
var COPILOT='https://copilot.microsoft.com/', NLM='https://notebooklm.google/';
var T3=[['copilot','Copilot ב-Excel'],['gemini','Gemini ב-Sheets'],['chat','צ\'אט חינמי עם העלאת קובץ']];
function ext(u,t){ return '<a class="ln" href="'+u+'" target="_blank" rel="noopener">'+t+'</a>'; }
function letterOf(f,name){ return (f.letters||{})[name]||''; }
function colList(f){ return f.head.map(function(h){ return letterOf(f,h)+' '+h; }).join(', '); }
function plist(list){ return list.map(function(p,i){ return prompt((i+1)+'. '+p[0],p[1]); }).join(''); }
function solQ(q){ var p=isPct(q.q), frac=p&&Math.abs(q.v)<=1.5;   // a column of 0-1 fractions is shown as %, a column of 0-100 numbers as is
  return '<li>'+esc(q.q)+' <b>'+num(q.v,frac)+'</b>'+fx(q.f)+(p&&!frac&&q.v>100?'<p class="out">ממוצע מעל 100? בעמודה יש ערכים חריגים: גדולים מ-100 ומתחת לאפס. זה חלק מהתרגיל: מוצאים אותם, ומחשבים שוב.</p>':'')+'</li>'; }

function build(n,c,f){
  var L=c.level, own=c.src&&c.src!=='practice', cols=own?'[שמות העמודות והאותיות שלהן]':colList(f), r={lv:false,ex:'',tip:'',prompts:[],sol:'',check:''};
  var byHood=f.hood?'שכונה':f.piv[0].rows, col0=own?'[עמודה]':f.filters[0];
  if(n===1){
    r.lv=true;
    r.steps=['פותחים '+ext(GEMINI,'Gemini')+' או '+ext(COPILOT,'Copilot')+' בלשונית נפרדת.','מתארים את העמודות, בלי הנתונים.','מבקשים נוסחה ומדביקים בגיליון.','שגיאה? מדביקים אותה בחזרה.'];
    r.note='כותבים לצ\'אט את שמות העמודות ומה רוצים לחשב, בלי להדביק נתונים. כך אפשר לעבוד גם על טבלה שאסור להעלות.';
    var ask;
    if(L==='b'){ ask=f.q3[0].q; r.task='<p>בקשו נוסחה לכל אחת משלוש השאלות, הדביקו בגיליון ובדקו:</p><ol>'+f.q3.map(function(q){ return '<li>'+esc(q.q)+'</li>'; }).join('')+'</ol>';
      r.sol='<p>בקובץ התרגול אמור לצאת:</p><ul>'+f.q3.map(solQ).join('')+'</ul>'; }
    else if(L==='m'){ ask=f.mid.q; r.task='<p>בקשו נוסחה עם תנאי: <b>'+esc(f.mid.q)+'</b> אחר כך נסחו שאלה משלכם, ובקשו גם לה נוסחה.</p>';
      r.sol='<p>לשאלה הראשונה אמור לצאת <b>'+num(f.mid.v)+'</b>:</p>'+fx(f.mid.f); }
    else{ var E=f.cols[4]; ask='עמודה חדשה בשם "'+E.name+'": '+E.what+'. נוסחה לשורה 2, שאפשר לגרור למטה';
      r.task='<p>בנו עמודה חדשה בנוסחה אחת: <b>'+esc(E.name)+'</b>. '+esc(E.what)+'. גררו למטה, ובדקו ביד חמש שורות.</p>';
      r.sol='<p>נוסחה אפשרית לשורה 2:</p>'+fx(E.f)+'<p>נוסחה אחרת שנותנת אותן תוצאות גם היא נכונה.</p>'; }
    if(own){ r.task='<p>בקשו נוסחה אחת לקובץ שלכם, הדביקו בגיליון ובדקו.</p>'; r.sol=''; ask='[מה רוצים לחשב]'; }
    var keyed=!/^(02|03|10)$/.test(f.n)&&!own;
    r.prompts=[['נוסחה מתיאור העמודות','העמודות: '+cols+'. כתוב נוסחה ל-'+(c.tool==='gemini'?'Google Sheets':'Excel')+' שעונה על: '+ask],
      ['שגיאה','הנוסחה מחזירה ‎#N/A. הנה הנוסחה ושורה לדוגמה: [להדביק]. מה הבעיה ואיך מתקנים?'],
      ['כפילויות','כתוב נוסחה שמסמנת "כפול" כשאותו '+(keyed?f.id+' מופיע יותר מפעם אחת בעמודה '+letterOf(f,f.id):'[מזהה] מופיע יותר מפעם אחת בעמודה [אות]')],
      ['הסבר','הסבר צעד אחר צעד מה הנוסחה הזו עושה, ותן גרסה ל-Google Sheets: [להדביק]']];
    r.check='<p>נוסחה אחת לפחות עובדת בגיליון שלכם.</p>';
  }
  if(n===2){
    r.steps=['מעלים קובץ ל-'+ext(GEMINI,'Gemini')+': Excel או CSV.','בוחרים Canvas ומבקשים דשבורד.','מתקנים בשיחה: "הוסף מסנן", "שנה צבעים".','משתפים בקישור.'];
    r.note='מי שמקבל את הקישור רואה את הנתונים שבדשבורד. משתפים רק דשבורד על נתונים פתוחים או מותממים.';
    r.task='<p>בונים דשבורד בסיסי מהקובץ'+(own?' שלכם':' ('+esc(f.short)+')')+'. בסוף בודקים שמספר הרשומות תואם לקובץ.</p>';
    r.prompts=[['דשבורד','צור דשבורד אינטראקטיבי בעברית, מימין לשמאל, מהקובץ המצורף'],
      ['מדדים וגרף','הוסף 4 מדדים בראש העמוד וגרף לפי '+byHood],
      ['מסנן','הוסף מסנן לפי '+col0+' וכפתור איפוס'],
      ['עיצוב','השתמש בצבעים רגועים והגדל את הגופן'],
      ['בדיקה','בדוק שמספר הרשומות בדשבורד זהה לקובץ']];
    r.more=[['אתר','בנה אתר של עמוד אחד בעברית, מימין לשמאל, שמציג לתושבים את הנתונים מהקובץ'],
      ['אתר','הוסף לאתר כותרת, שלושה מספרים בולטים, וגרף לפי שכונה'],
      ['דשבורד','בנה דשבורד אינטראקטיבי מהקובץ: 4 מדדים בראש, גרף לפי שכונה ומסנן לפי שנה'],
      ['דשבורד','הוסף טבלה שאפשר למיין ולחפש בה, וכפתור להורדת הנתונים'],
      ['דשבורד','הגדל את הגופן לקריאה מהמקרן, ושנה לצבעים רגועים']];
    r.check='<p><b>כל הרשומות בדשבורד?</b> משווים את הסך הכול בדשבורד למספר השורות בקובץ'+(own?'':': בקובץ '+f.n+' יש <b>'+num(f.st.raw)+'</b> שורות')+'. תואם? ממשיכים. לא תואם? מבקשים לבנות מחדש על כל השורות.</p>';
    r.tip='<p><b>לא תואם?</b> בדרך כלל ה-AI קרא רק חלק מהקובץ. בחשבון חינמי זה קורה בקבצים גדולים. אפשר גם להעלות טבלת סיכום קטנה יותר, ולבדוק שוב.</p>';
  }
  if(n===3){
    r.lv=true; r.tools=true;
    var T=c.tool==='copilot'||c.tool==='gemini'?c.tool:'chat';
    r.steps={copilot:['פותחים את הקובץ ב-Excel, בחשבון העבודה.','פותחים את Copilot בלשונית "בית".'],
             gemini:['מעלים את הקובץ ל-Google Drive ופותחים אותו ב-Google Sheets. קובץ Excel שומרים קודם כ-Google Sheets (קובץ ← שמירה כ-Google Sheets).','פותחים את חלונית Gemini בגיליון.'],
             chat:['מעלים את הקובץ ל-'+ext(GEMINI,'Gemini')+' או לצ\'אט אחר. קובץ מהעבודה רק אחרי '+'<a class="ln" href="../anonymize/">התממה</a>.','את הגיליון משאירים פתוח. שם מדביקים נוסחאות ובודקים.']}[T]
      .concat(['מדביקים פרומפט אחד בכל פעם.','מאמתים שלושה מספרים מול הגיליון.']);
    r.note={copilot:'Copilot לא מגיב לקובץ? בדרך כלל הקובץ צריך להיות שמור ב-OneDrive או ב-SharePoint. אין רישיון? עובדים באותם פרומפטים בצ\'אט.',
            gemini:'אין חלונית Gemini בגיליון? היא תלויה במנוי של הארגון. עובדים באותם פרומפטים בצ\'אט.',
            chat:'הצ\'אט לא רואה את הגיליון שלכם, רק את הקובץ שהעליתם. כל מספר בודקים בגיליון.'}[T];
    var upto={b:2,m:3,a:5}[L];
    r.task='<p>עובדים לפי הסדר. ברמה '+label(LEVELS,L)+': '+(upto===5?'כל חמשת הפרומפטים':'פרומפטים 1 עד '+upto+'. סיימתם? ממשיכים הלאה')+'.</p>';
    var C0=f.cols[0];
    r.prompts=[['ניתוח','נתח את הטבלה: מה המגמות העיקריות? 3 תובנות עם מספרים'],
      ['שגיאות','מצא שגיאות: כפילויות, תאים ריקים, ערכים חריגים וכתיב שונה'],
      ['סידור','סדר את הנתונים: אחד כתיב, פצל עמודות ותקן תאריכים'],
      ['דשבורד','צור גיליון דשבורד עם מדדים וגרפים שמתעדכנים כשמוסיפים שורות'],
      ['עמודה מחושבת',own?'הוסף עמודה שמחשבת [מה לחשב], והסבר באיזו נוסחה השתמשת':'הוסף עמודה בשם "'+C0.name+'": '+C0.what+'. הסבר באיזו נוסחה השתמשת']];
    if(!own) r.sol='<p>מספרים לבדיקה בקובץ '+f.n+':</p><ul><li>שורות בקובץ: <b>'+num(f.st.raw)+'</b>. אחרי הסרת כפילויות: <b>'+num(f.st.clean)+'</b>.</li>'+
      (f.hood?'<li>כתיבים שונים בעמודת השכונה: <b>'+f.st.spell_before+'</b>. אחרי איחוד: <b>'+f.st.spell_after+'</b>.</li>':'')+
      (f.st.text?'<li>מספרים או תאריכים שנשמרו כטקסט: <b>'+f.st.text+'</b> תאים.</li>':'')+'</ul>'+
      '<p>'+esc(C0.name)+', נוסחה אפשרית לשורה 2:</p>'+fx(C0.f);
    r.check='<p>שלושה מספרים מהתשובה, שבדקתם מול הגיליון.</p>';
    r.pit=['טעה בתאריכים','בלבל בין עמודות','קרא רק חלק מהקובץ','חישב לא נכון','המציא מספר','תא ריק נספר כאפס'];
  }
  if(n===4){
    r.steps=['נכנסים ל-'+ext(NLM,'Gemini Notebook')+' (לשעבר NotebookLM), בחינם עם חשבון Google.','יוצרים מחברת ומעלים שניים או שלושה קבצים.','שואלים שאלה אחת. כל משפט בתשובה מפנה למקור.','ב-Studio יוצרים תוצר ויזואלי אחד.'];
    r.note='קובץ Excel לא עולה? פותחים אותו ב-Google Sheets, ומוסיפים אותו כמקור מ-Google Drive.';
    r.task='<p>מעלים את הקובץ שלכם ועוד קובץ או שניים: קובץ תרגול נוסף, מאגר פתוח של באר שבע, או מסמך רקע. שואלים שאלה אחת, ויוצרים תוצר ויזואלי אחד.</p>';
    r.prompts=[['מגמות','מה שלוש המגמות הבולטות בקבצים? ציין מקור לכל אחת'],
      ['השוואה','השווה בין השנים בטבלה'],
      ['מה חסר','אילו נתונים חסרים כדי לענות על השאלה?'],
      ['Studio','צור אינפוגרפיקה בעברית לתושבים'],
      ['Studio','צור מפת חשיבה של הנושאים']];
    r.check='<p>תשובה אחת עם הפניה למקור, ותוצר אחד מ-Studio.</p>';
  }
  return r;
}

function round(){
  var c=choice(), f=F[c.file], n=ROUND, R=ROUNDS[n], r=build(n,c,f), li=LEVELS.map(function(l){return l[0];}).indexOf(c.level);
  var own=c.src!=='practice';
  var h=stepper('r'+n)+'<div class="mine lv-'+c.level+'"><span class="f">'+(own?(c.src==='own'?'קובץ שלי מהעבודה':'נתונים פתוחים של באר שבע'):'קובץ '+f.n+': '+esc(f.short))+'</span>'+(r.lv?'<span class="lvtag">רמה '+label(LEVELS,c.level)+'</span>':'')+'<a class="ln" href="../file/">החלפת קובץ'+(r.lv?' או רמה':'')+'</a></div>';
  h+='<div class="rhead"><p>'+R.s+'</p><p class="rmeta"><span>'+R.min+' דקות</span><span>'+R.tool.split(' · ').map(function(x){ return '<bdi>'+x+'</bdi>'; }).join(' · ')+'</span></p></div>';
  if(own&&n!==4) h+='<section class="box tip"><p><b>אתם עובדים על קובץ משלכם.</b> במקום שמסומן בסוגריים מרובעים, כותבים את שמות העמודות שלכם.</p></section>';
  h+='<section class="box task lv-'+c.level+'"><h2>המשימה'+(r.lv?' <span class="lvtag">רמה '+label(LEVELS,c.level)+'</span>':'')+'</h2>'+r.task+
     (r.lv?'<div class="shift"><button type="button" id="easier"'+(li===0?' disabled':'')+'>קשה לי: רמה אחת למטה</button><button type="button" id="harder"'+(li===2?' disabled':'')+'>קל לי: רמה אחת למעלה</button></div>':'')+'</section>';
  h+='<section class="box"><h2>השלבים'+(r.tools?': '+esc(label(T3,c.tool==='copilot'||c.tool==='gemini'?c.tool:'chat')):'')+'</h2><ol>'+r.steps.map(function(s){ return '<li>'+s+'</li>'; }).join('')+'</ol><p class="out">'+r.note+'</p>'+
     (r.tools?'<details class="hint alt"'+(c.tool==='copilot'||c.tool==='gemini'||c.tool==='chat'?'':' open')+'><summary>באיזה כלי אתם עובדים?</summary><div class="in"><label for="sTool" class="out">השלבים יותאמו לכלי שתבחרו:</label> <select id="sTool">'+T3.map(function(o){ return '<option value="'+o[0]+'"'+(o[0]===c.tool?' selected':'')+'>'+esc(o[1])+'</option>'; }).join('')+'</select></div></details>':'')+'</section>';
  h+='<section class="box"><h2>פרומפטים להעתקה</h2>'+plist(r.prompts)+'</section>';
  if(r.more) h+='<section class="box"><h2>זמן יצירה: אתר או דשבורד</h2><p class="out">מעלים קובץ, בוחרים Canvas, ומדביקים אחד מהפרומפטים.</p>'+r.more.map(function(p){ return prompt(p[0],p[1]); }).join('')+'</section>';
  if(r.pit) h+='<section class="box"><h2>תקלות שקורות, ואיך תופסים אותן</h2><ul class="pit">'+r.pit.map(function(x){ return '<li>'+x+'</li>'; }).join('')+'</ul><p>שואלים: "הראה לי איך חישבת", ובודקים מול הגיליון.</p></section>';
  if(r.tip) h+='<section class="box tip">'+r.tip+'</section>';
  if(r.sol) h+='<details class="hint sol"><summary>פתרון: מה אמור לצאת בקובץ התרגול</summary><div class="in">'+r.sol+'</div></details>';
  h+='<section class="box check"><h2>נקודת ביקורת</h2>'+r.check+'</section>';
  if(n<=3) h+='<section class="box stuck"><span class="t">קובץ התרגול</span><span class="d">קובץ '+f.n+': '+esc(f.short)+' · '+num(f.st.raw)+' שורות</span>'+dl(raw(f),f.n+' '+f.short+' - נתונים פיקטיביים.xlsx','הורדה ←')+'</section>';
  h+='<details class="hint more"><summary>שני פרומפטים שעוזרים בכל שלב</summary><div class="in">'+
     prompt('תקלה','קיבלתי את השגיאה הזאת: [מדביקים כאן את השגיאה]. תקן ותן לי את הכל מחדש.')+
     prompt('בדיקה עצמית','אילו מספרים בתשובה חישבת בעצמך, ואילו לקחת מנוסחה? מה כדאי לאמת ידנית?')+'</div></details>';
  h+=nextCard('r'+n);
  var app=document.getElementById('app'); app.innerHTML=h;
  function set(k,v){ var x=choice(); x[k]=v; save(x); var y=window.scrollY; round(); window.scrollTo(0,y); }
  var st=document.getElementById('sTool'); if(st) st.onchange=function(){ set('tool',this.value); };
  var e1=document.getElementById('easier'), e2=document.getElementById('harder');
  if(e1) e1.onclick=function(){ set('level',LEVELS[li-1][0]); };
  if(e2) e2.onclick=function(){ set('level',LEVELS[li+1][0]); };
  if(!c.set) save(c);
}

/* ---------- open data (deck: div-e, e-list) ---------- */
function openPage(){
  var S=[['data.gov.il','המאגר הממשלתי הפתוח. כאן המאגרים של עיריית באר שבע.',OPEN],
    ['הלמ"ס','מפקד האוכלוסין ונתוני רשויות','https://www.cbs.gov.il/'],
    ['GovMap','מפות ושכבות מידע: מחפשים כתובת, גוש וחלקה, ומוסיפים שכבות.','https://www.govmap.gov.il/'],
    ['שקיפות בחינוך','נתוני מוסדות חינוך: הישגים, נשירה, אקלים ותקציב.','https://shkifut.education.gov.il/home'],
    ['נתונים מוניציפליים','דמוגרפיה והשוואה בין רשויות.','https://municipal-data.org/demographics'],
    ['אתר העירייה','באר שבע במספרים: הדשבורד הדמוגרפי.','https://www.beer-sheva.muni.il/City/OnTheCity/b7numbers/Documents/demographic-dash.html']];
  var h=stepper('open')+'<div class="rhead"><p>איפה מוצאים נתונים, גם בלי לבקש מאף אחד. מידע פתוח לא צריך התממה.</p></div>';
  h+='<ul class="srcl">'+S.map(function(s){ return '<li><a href="'+s[2]+'" target="_blank" rel="noopener"><strong>'+esc(s[0])+' ↗</strong><span>'+esc(s[1])+'</span></a></li>'; }).join('')+'</ul>';
  h+='<p class="out pad">עוד מקורות, לפי נושא: <a class="ln" href="../../data/sources/">כלי מיפוי מקורות המידע</a></p>';
  document.getElementById('app').innerHTML=h;
}

var APPEL=document.getElementById('app'); if(APPEL){ APPEL.classList.add('first'); setTimeout(function(){ APPEL.classList.remove('first'); },1100); }
if(ROUND) round();
else if(PAGE==='id') home();
else if(PAGE==='file') filePage();
else if(PAGE==='anon') anonPage();
else if(PAGE==='prompts') promptsPage();
else if(PAGE==='open') openPage();
else if(PAGE==='end'){ var sp=document.getElementById('stepper'); if(sp) sp.innerHTML=stepper('end'); }
if(PAGE&&PAGE!=='prompts'&&PAGE!=='id') document.body.insertAdjacentHTML('beforeend','<a class="pbank" href="'+UP+'prompts/"><span>מאגר הפרומפטים</span></a>');
window.DataAI={choice:choice,files:F,levels:LEVELS,tools:TOOLS,rounds:ROUNDS,rules:RULES,esc:esc,stepper:stepper,mark:mark,build:build};
})();
