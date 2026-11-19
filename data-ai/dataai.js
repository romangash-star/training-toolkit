/* "עבודה עם נתונים בעזרת בינה מלאכותית": the section page and the five round pages.
   The facts about each practice file (columns, formulas, expected answers) come from files.js,
   which tools/dataai_build.py generates and checks in a real Excel. The wording of the tasks lives here.
   The deck shows the same tasks: when a task changes here, change the slide too. */
(function(){
'use strict';
var F=window.DATAAI_FILES, KC='dataai_choice', KI='toolkit_identity';
var ROUND=+document.body.dataset.round||0, UP=ROUND?'../':'';
var LEVELS=[['b','בסיסית'],['m','בינונית'],['a','מתקדמת']];
var TOOLS=[['copilot','Copilot ב-Excel'],['gemini','Gemini ב-Sheets'],['chat','צ\'אט חינמי עם העלאת קובץ']];
var ROUNDS=[null,
  {t:'לשאול את הטבלה', s:'מדברים עם הקובץ בשפה חופשית, ומבקשים נוסחה במקום מספר.'},
  {t:'לנקות', s:'ה-AI מאבחן בעיות באיכות הנתונים ומציע תיקון. אתם מחליטים ומבצעים.'},
  {t:'להעשיר ולסווג', s:'מוסיפים עמודות חדשות: חישובים, חלקי תאריך וקטגוריות.'},
  {t:'לנתח ולמצוא תובנה', s:'ה-AI מציע הצלבות, מסביר איך בונים טבלת ציר, ומנסח תובנה.'},
  {t:'מתובנה למוצר', s:'הופכים את הממצא למשהו שמנהל פותח בלי לגעת בגיליון.'}];
var RULES=['מתממים לפני שמעלים','נוסחה ולא תשובה','מנקים לפני שמנתחים','בודקים מדגם','תובנה עוברת בדיקות'];
var AGG={sum:'סכום',avg:'ממוצע',count:'ספירה'};
var APP={copilot:'Excel',gemini:'Google Sheets',chat:'Excel'};

function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function num(v,pct){
  if(typeof v!=='number') return esc(v);
  if(pct) return (v*100).toLocaleString('he-IL',{maximumFractionDigits:1})+'%';
  return v.toLocaleString('he-IL',{maximumFractionDigits:Math.abs(v)<100&&v%1?2:(v%1?1:0)});
}
function isPct(name){ return /אחוז|שיעור/.test(name||''); }
function choice(){
  var c={}; try{ c=JSON.parse(localStorage.getItem(KC)||'{}')||{}; }catch(e){}
  var ok=!!(F[c.file]&&c.level&&c.tool);
  return {file:F[c.file]?c.file:'01',
          level:LEVELS.some(function(l){return l[0]===c.level;})?c.level:'b',
          tool:TOOLS.some(function(t){return t[0]===c.tool;})?c.tool:'chat', set:ok};
}
function save(c){ try{ localStorage.setItem(KC,JSON.stringify({file:c.file,level:c.level,tool:c.tool})); }catch(e){} }
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

/* ---------- section page ---------- */
function identity(){
  var zone=document.getElementById('idZone'); if(!zone) return;
  function get(){ try{ return JSON.parse(localStorage.getItem(KI)||'{}'); }catch(e){ return {}; } }
  function render(edit){
    var d=get();
    if(d.name&&!edit){
      zone.innerHTML='<div class="idbar"><span class="who">'+esc([d.name,d.org,d.dept,d.role].filter(Boolean).join(' · '))+'</span><button id="idEdit">שינוי</button></div>';
      document.getElementById('idEdit').onclick=function(){ render(true); }; return;
    }
    zone.innerHTML='<div class="idcard"><h3>מי אתם? <button type="button" class="help" aria-label="הסבר" data-help="השם ושם הרשות נשמרים <b>רק במחשב שלכם</b>, ומשמשים את הכלים באתר. הם לא נשלחים לשום מקום.">?</button></h3>'+
      '<div class="fields"><input class="full" id="idName" placeholder="שם מלא" value="'+esc(d.name||'')+'"><input class="full" id="idOrg" placeholder="שם הרשות (עירייה / מועצה)" value="'+esc(d.org||'')+'">'+
      '<input id="idDept" placeholder="אגף / מחלקה" value="'+esc(d.dept||'')+'"><input id="idRole" placeholder="תפקיד" value="'+esc(d.role||'')+'"></div><button class="save" id="idSave">שמירה ←</button></div>';
    document.getElementById('idSave').onclick=function(){
      var g=function(i){ return (document.getElementById(i).value||'').trim(); }, v={name:g('idName'),org:g('idOrg'),dept:g('idDept'),role:g('idRole')};
      if(!v.name){ document.getElementById('idName').focus(); return; }
      if(!v.org){ document.getElementById('idOrg').focus(); return; }
      try{ localStorage.setItem(KI,JSON.stringify(v)); }catch(e){}
      render(false);
    };
  }
  render(false);
}
function home(){
  var c=choice(), f=F[c.file];
  function chips(list,cur,key,cls){
    return '<div class="chips">'+list.map(function(o){
      return '<button type="button" class="chip '+(cls?cls(o):'')+'" data-k="'+key+'" data-v="'+o[0]+'" aria-pressed="'+(c.set&&o[0]===cur)+'">'+o[1]+'</button>'; }).join('')+'</div>';
  }
  var files=Object.keys(F).sort().map(function(n){ return [n,'<span class="num">'+n+'</span>'+esc(F[n].short)+'<small>'+esc(F[n].who)+'</small>']; });
  document.getElementById('pick').innerHTML=
    '<div><h3>1. הקובץ שלי</h3>'+chips(files,c.file,'file')+'</div>'+
    '<div><h3>2. הרמה שלי</h3>'+chips(LEVELS,c.level,'level',function(o){ return 'lv lv-'+o[0]; })+'</div>'+
    '<div><h3>3. הכלי שלי</h3>'+chips(TOOLS,c.tool,'tool')+'</div>';
  document.getElementById('start').innerHTML=c.set?
    '<div class="start"><span class="t">קובץ '+f.n+': '+esc(f.short)+'</span><span class="d">'+num(f.st.raw)+' שורות · '+f.head.length+' עמודות · נתונים פיקטיביים</span>'+
      dl(raw(f),f.n+' '+f.short+' - נתונים פיקטיביים.xlsx','הורדת הקובץ ←')+'</div>'+
    '<ul class="checks"><li><label><input type="checkbox">הקובץ נפתח?</label></li><li><label><input type="checkbox">כלי ה-AI נפתח בחשבון העבודה?</label></li></ul>'
    :'<p class="out">בחרו קובץ, רמה וכלי. הבחירה נשמרת במחשב הזה, ואפשר לשנות אותה בכל דף.</p>';
  [].forEach.call(document.querySelectorAll('#pick .chip'),function(b){
    b.onclick=function(){ var n=choice(); n[b.dataset.k]=b.dataset.v; save(n); home(); };
  });
}

/* ---------- round pages ---------- */
function pivHow(f,p){
  if(p.agg==='yes'||p.agg==='no') return 'מסננים את "'+p.val+'" ל"'+(p.agg==='yes'?'כן':'לא')+'". בשורות: "'+p.rows+'". בערכים: ספירה.';
  if(p.agg==='count') return 'בשורות: "'+p.rows+'". בערכים: ספירה של "'+f.id+'".';
  return 'בשורות: "'+p.rows+'". בערכים: '+AGG[p.agg]+' של "'+p.val+'".';
}
function pivTable(p){
  var pct=isPct(p.val)&&p.agg==='avg';
  return '<p><b>'+esc(p.title)+'</b>'+(p.n>p.top.length?' (חמש הגבוהות מתוך '+p.n+')':'')+'</p><table class="mini"><tr><th>'+esc(p.rows)+'</th><th>'+(AGG[p.agg]||'ספירה')+'</th></tr>'+
    p.top.map(function(r){ return '<tr><td>'+esc(r[0])+'</td><td class="num">'+num(r[1],pct)+'</td></tr>'; }).join('')+'</table>';
}
function kpiList(f){
  return '<ul>'+f.kpis.map(function(k){ return '<li><b>'+esc(k.name)+':</b> '+num(k.v,isPct(k.name))+fx(k.f)+'</li>'; }).join('')+'</ul>';
}
function colSol(c){ return '<p><b>'+esc(c.name)+'</b> ('+esc(c.what)+')</p>'+fx(c.f); }

function build(n,c,f){
  var L=c.level, T=c.tool, names=f.head.join(', '), A=f.cols[0], B=f.cols[1], C=f.cols[2], D=f.cols[3], E=f.cols[4];
  var hoodL=f.hood?f.letters['שכונה']:null, idL='A', app=APP[T];
  var r={ex:'',tip:'',prompts:[]};
  if(n===1){
    if(L==='b'){
      r.task='<p>בקשו מה-AI לתאר את הקובץ. אחר כך שאלו את שלוש השאלות האלה, ולכל אחת בקשו נוסחה:</p><ol>'+f.q3.map(function(q){ return '<li>'+esc(q.q)+'</li>'; }).join('')+'</ol>';
      r.prompts=[['היכרות עם הקובץ','זו טבלת '+f.topic+' של עירייה. תאר בשלושה משפטים מה יש בה, כמה שורות, ואיזה טווח תאריכים או תקופה.'],
                 ['נוסחה, לא תשובה','אל תחשב בעצמך. תן לי נוסחת '+app+' שעונה על: '+f.q3[0].q+' העמודות: '+names+'.']];
      r.hints=['ה-AI נתן מספר בלי נוסחה? כתבו לו: "אל תחשב בעצמך, תן נוסחה". מספר שה-AI חישב לבד אי אפשר לבדוק.',
               'כותבים את הנוסחה בתא ריק, משמאל לטבלה. אם יצאה שגיאה, מדביקים אותה ל-AI ומבקשים תיקון.'];
      r.sol='<p>בקובץ המקורי אמור לצאת:</p><ul>'+f.q3.map(function(q){ return '<li>'+esc(q.q)+' <b>'+num(q.v,isPct(q.q))+'</b>'+fx(q.f)+'</li>'; }).join('')+'</ul>';
    }else if(L==='m'){
      r.task='<p>נסחו שלוש שאלות משלכם על הקובץ. לכל אחת בקשו נוסחה, הדביקו אותה בגיליון, והשוו למספר שה-AI נתן.</p>';
      r.ex='<b>דוגמה לשאלה:</b> '+esc(f.mid.q);
      r.prompts=[['נוסחה, לא תשובה','אל תחשב בעצמך. תן לי נוסחת '+app+' שעונה על: '+f.mid.q+' העמודות: '+names+'.']];
      r.hints=['שאלה טובה כוללת שם של עמודה ותנאי אחד. את השאלה שבפרומפט מחליפים בשאלה שלכם.',
               'לשאלה עם תנאי אחד מתאימות הנוסחאות COUNTIF, SUMIF ו-AVERAGEIF.'];
      r.sol='<p>לשאלת הדוגמה, בקובץ המקורי אמור לצאת <b>'+num(f.mid.v)+'</b>:</p>'+fx(f.mid.f)+
        (/שכונ/.test(f.mid.q)?'<p>שימו לב: הנוסחה סופרת רק את הכתיב המדויק. שורות שבהן שם השכונה כתוב אחרת נשארות בחוץ. את זה מתקנים בסבב 2.</p>':'');
    }else{
      r.task='<p>בקשו מה-AI חמש שאלות ש'+esc(f.role)+' היה שואל על הנתונים, ולכל שאלה נוסחה או טבלת ציר. בצעו לפחות שתיים מהן בגיליון.</p>';
      r.ex='<b>דוגמה לשאלה:</b> '+esc(f.adv);
      r.prompts=[['לחשוב כמו מנהל','אילו 5 שאלות '+f.role+' היה שואל על הנתונים האלה? לכל שאלה, נוסחת '+app+' או טבלת ציר שעונה עליה. העמודות: '+names+'.']];
      r.hints=['השאלות יצאו כלליות מדי? הוסיפו לפרומפט את השאלה הניהולית: "'+f.ask+'"',
               'בקשו שליד כל שאלה יהיה כתוב באילו עמודות היא משתמשת. כך קל לבדוק.'];
      r.sol='<p>אין כאן תשובה אחת. הבדיקה: לפחות שתי שאלות נענו, ולאחת מהן יש נוסחה שעובדת בגיליון.</p><p>לשאלה "'+esc(f.mid.q)+'" אמור לצאת <b>'+num(f.mid.v)+'</b>:</p>'+fx(f.mid.f);
    }
    r.check='<p>שתי שאלות שנענו, ונוסחה אחת שעובדת בגיליון ומחזירה מספר.</p>';
    r.bonus='<p>שאלו את ה-AI: "אילו מספרים בתשובה חישבת בעצמך, ואילו לקחת מנוסחה? מה כדאי לאמת ידנית?"</p>';
    if(T==='chat'&&f.head.some(function(h){ return /תאריך/.test(h); })) r.tip='<p><b>תאריכים שנראים כמו מספרים:</b> אם שמרתם את הקובץ כ-CSV והעליתם לצ\'אט, תאריך יכול להופיע כמספר (למשל 46043.25). זה אותו תאריך בתצוגה אחרת. כתבו ל-AI שהעמודה היא תאריך.</p>';
    r.stuck=['מתחילים מהקובץ המקורי',num(f.st.raw)+' שורות',raw(f),f.n+' '+f.short+' - נתונים פיקטיביים.xlsx'];
  }
  if(n===2){
    var keyed=!/^(02|03|10)$/.test(f.n);
    var askHood=f.spell&&f.spell.some(function(s){ return s[1]==='לבירור'; });
    if(L==='b'){
      r.task='<p>בקשו מה-AI רשימה של בעיות באיכות הנתונים. אחר כך הסירו כפילויות'+(keyed?' לפי העמודה "'+esc(f.id)+'"':' (שורות שזהות בכל העמודות)')+', ושורות ריקות אם יש. רשמו כמה שורות היו לפני, וכמה אחרי.</p>';
      r.prompts=[['אבחון','אילו בעיות איכות נתונים יש בקובץ? רשימה ממוספרת, דוגמה לכל בעיה, ואיך לתקן אותה.']];
      r.hints=['כפילות היא שורה שמופיעה פעמיים. ה-AI מציע, ואתם מוחקים בעצמכם בגיליון: כך רואים מה נמחק.',
               'סופרים שורות לפני ואחרי עם הנוסחה =COUNTA(A:A)-1'];
      r.sol='<p>בקובץ המקורי <b>'+num(f.st.raw)+'</b> שורות. אחרי הסרת הכפילויות: <b>'+num(f.st.clean)+'</b>.</p>'+
        (keyed?'<p>יצא לכם מספר גבוה יותר? נשארו כפילויות שנבדלות רק בצורת הכתיבה של תאריך או מספר. הסירו כפילויות לפי העמודה "'+esc(f.id)+'" בלבד.</p>'
              :'<p>בקובץ הזה "'+esc(f.id)+'" חוזר בכוונה בכמה שורות, ולכן מסירים רק שורות שזהות בכל העמודות.</p>');
    }else if(L==='m'){
      if(f.hood){
        r.task='<p>אחרי הסרת הכפילויות: בקשו מה-AI טבלה של כל הכתיבים בעמודת השכונה מול השם התקני. שימו אותה בגיליון חדש בשם "כתיבים", והוסיפו לטבלה עמודה "שכונה מתוקנת" עם XLOOKUP.</p>';
        r.prompts=[['איחוד שמות','תן טבלה של כל הכתיבים בעמודת שכונה, וליד כל כתיב את השם התקני. רק טבלה, בלי הסברים.']];
        r.hints=['כדי לראות את כל הכתיבים: =UNIQUE('+hoodL+':'+hoodL+') בתא ריק. כמה ערכים יצאו?',
                 'יש כתיבים שנבדלים רק ברווח בסוף המילה. בעין לא רואים אותו, אבל הגיליון רואה.'];
        r.sol='<p>לפני: <b>'+f.st.spell_before+'</b> כתיבים. אחרי: <b>'+f.st.spell_after+'</b> ערכים.</p>'+fx('=XLOOKUP('+hoodL+'2,כתיבים!A:A,כתיבים!B:B,'+hoodL+'2)')+
          (askHood?'<p>"מכבים-רעות" הוא לא טעות כתיב: אי אפשר לדעת לאיזו שכונה השורה שייכת. בקובץ הנקי הוא מסומן "לבירור". זו החלטה של בעל הנתונים, לא של ה-AI.</p>':'');
      }else{
        r.task='<p>אחרי הסרת הכפילויות: מצאו את התאים שבהם מספר או תאריך נשמרו כטקסט, ותקנו אותם בעמודת עזר.</p>';
        r.prompts=[['מספרים שנשמרו כטקסט','באילו עמודות יש מספרים או תאריכים שנשמרו כטקסט? תן נוסחת '+app+' לעמודת עזר שממירה אותם לערך אמיתי.']];
        r.hints=['=ISTEXT(תא) מחזיר TRUE כשמספר נשמר כטקסט. סימן נוסף: המספר מיושר לצד אחר מכל השאר.',
                 '=VALUE(TRIM(תא)) הופך טקסט למספר. לתאריך שנכתב כמו 15.04.22 צריך לפרק את הטקסט לימים, חודשים ושנים.'];
        r.sol='<p>בקובץ הזה יש <b>'+f.st.text+'</b> תאים כאלה. בקובץ הנקי כולם הומרו למספר או לתאריך.</p>';
      }
    }else{
      var pq=T==='gemini'?'שאילתת QUERY ב-Google Sheets':(T==='copilot'?'שלבי Power Query':'שלבי Power Query (או שאילתת QUERY, אם אני ב-Google Sheets)');
      r.task='<p>בנו את כל הניקוי כתהליך שחוזר: '+(T==='gemini'?'נוסחת QUERY ב-Sheets':'Power Query ב-Excel')+'. אחר כך הוסיפו לקובץ המקורי כמה שורות מבולגנות, רעננו, ובדקו שהן נוקו.</p>';
      r.prompts=[['תהליך שחוזר','כתוב '+pq+' שמסירים שורות ריקות וכפילויות'+(f.hood?' ומאחדים את עמודת שכונה לפי טבלת כתיבים (כתיב מול שם תקני)':' וממירים מספרים ותאריכים שנשמרו כטקסט')+'. העמודות: '+names+'. הסבר כל שלב בשורה.']];
      r.hints=['המטרה: גיליון המקור נשאר כמו שהוא, והגיליון הנקי נבנה ממנו מחדש בכל רענון.',
               'בדיקה: מוסיפים למקור שורה כפולה'+(f.hood?' ושורה עם שכונה בכתיב שגוי':'')+', מרעננים, ומחפשים אותן בגיליון הנקי.'];
      r.sol='<p>אחרי רענון אמורות להישאר <b>'+num(f.st.clean)+'</b> שורות מהקובץ המקורי, והשורות המבולגנות שהוספתם לא מופיעות פעמיים.</p>';
    }
    if(L!=='b') r.prompts.unshift(['אבחון','אילו בעיות איכות נתונים יש בקובץ? רשימה ממוספרת, דוגמה לכל בעיה, ואיך לתקן אותה.']);
    r.check='<p>מספר השורות לפני ואחרי: <b>'+num(f.st.raw)+' ← '+num(f.st.clean)+'</b>.'+(f.hood?' מספר הערכים בעמודת השכונה לפני ואחרי: <b>'+f.st.spell_before+' ← '+f.st.spell_after+'</b> (בודקים עם UNIQUE).':'')+'</p>';
    r.bonus='<p>בקובץ יש תאים חסרים. בקשו מה-AI לספור אותם לפי עמודה, והחליטו מה עושים: משאירים ריק, משלימים ממקור אחר, או מוציאים מהניתוח. לא ממציאים ערכים.</p>';
    r.stuck=['מתחילים מהקובץ המקורי',num(f.st.raw)+' שורות',raw(f),f.n+' '+f.short+' - נתונים פיקטיביים.xlsx'];
  }
  if(n===3){
    if(L==='b'){
      r.task='<p>הוסיפו לטבלה שתי עמודות חדשות:</p><ol><li><b>'+esc(A.name)+'</b>: '+esc(A.what)+'.</li><li><b>'+esc(B.name)+'</b>: '+esc(B.what)+'.</li></ol>';
      r.prompts=[['עמודה מחושבת','תן נוסחת '+app+' לעמודה חדשה בשם "'+A.name+'": '+A.what+'. אם חסר ערך בשורה, התא נשאר ריק. העמודות: '+names+'.']];
      r.hints=['כותבים את הנוסחה בשורה 2, בעמודה חדשה בסוף הטבלה, וגוררים למטה.','בודקים שורה אחת ביד, עם מחשבון. אם יצא אותו מספר, ממשיכים.'];
      r.sol=colSol(A)+colSol(B);
    }else if(L==='m'){
      r.task='<p>הוסיפו לטבלה עמודת דגל ועמודת סיווג:</p><ol><li><b>'+esc(C.name)+'</b>: '+esc(C.what)+'.</li><li><b>'+esc(D.name)+'</b>: '+esc(D.what)+'.'+(f.map?' הסיווג נעשה עם טבלת מיפוי בגיליון נפרד בשם "מיפוי".':'')+'</li></ol>';
      r.prompts=[['עמודת דגל','תן נוסחת '+app+' לעמודה חדשה שמסמנת כן או לא: '+C.what+'. אם חסר ערך בשורה, התא נשאר ריק. העמודות: '+names+'.'],
        f.map?['סיווג לקבוצות','קבץ את הערכים בעמודת '+f.map.src+' ל-'+f.map.pairs.map(function(p){return p[1];}).filter(function(v,i,a){return a.indexOf(v)===i;}).length+' קבוצות שמשמעותיות למנהל. החזר טבלת מיפוי: ערך וקבוצה.']
             :['סיווג למדרגות','תן נוסחת '+app+' לעמודה חדשה בשם "'+D.name+'": '+D.what+'. אם חסר ערך בשורה, התא נשאר ריק. העמודות: '+names+'.']];
      r.hints=['בדגל יש רק שני ערכים: כן ולא. תא שחסר בו נתון נשאר ריק, ולא הופך ל"לא".',
               f.map?'טבלת המיפוי: בעמודה A הערך המקורי, בעמודה B הקבוצה. הנוסחה מחפשת את הערך בטבלה ומחזירה את הקבוצה.':'במדרגות מתחילים מהתנאי הראשון ויורדים: כל IF בודק מדרגה אחת.'];
      r.sol=colSol(C)+colSol(D)+(f.map?'<table class="mini"><tr><th>'+esc(f.map.src)+'</th><th>קבוצה</th></tr>'+f.map.pairs.map(function(p){ return '<tr><td>'+esc(p[0])+'</td><td>'+esc(p[1])+'</td></tr>'; }).join('')+'</table>':'');
    }else{
      r.task='<p>בנו את העמודה <b>'+esc(E.name)+'</b> בנוסחת LET אחת: '+esc(E.what)+'. אחר כך בדקו ביד 15 שורות.</p>';
      r.prompts=[['כמה תנאים בנוסחה אחת','כתוב נוסחת LET אחת ב-'+app+' לעמודה חדשה בשם "'+E.name+'": '+E.what+'. תן שמות למשתנים, והסבר כל חלק בשורה. אם חסר ערך בשורה, התא נשאר ריק. העמודות: '+names+'.']];
      r.hints=['LET נותן שם לחישוב ביניים, כדי לכתוב אותו פעם אחת ולהשתמש בו כמה פעמים.','בדיקת מדגם: מסננים כל ערך של העמודה החדשה, ובודקים ביד שלוש שורות מכל ערך.'];
      r.sol=colSol(E);
    }
    r.check='<p>שלוש שורות שבדקתם ביד, ויצא בהן מה שציפיתם. אחת מהן שורה עם ערך חסר.</p>';
    r.bonus='<p>בקשו מה-AI עמודה נוספת שתעזור לענות על השאלה: "'+esc(f.ask)+'"</p>';
    r.stuck=['הקובץ הנקי, מסוף סבב 2',num(f.st.clean)+' שורות'+(f.hood?', עם העמודה "שכונה מתוקנת"':''),UP+'files/'+f.n+'-clean.xlsx',f.n+' '+f.short+' - נקי.xlsx'];
  }
  if(n===4){
    var p0=f.piv[0], p1=f.piv[1], p2=f.piv[2];
    if(L==='b'){
      r.task='<p>בנו טבלת ציר אחת: <b>'+esc(p0.title)+'</b>. הוסיפו גרף אחד, וכתבו משפט תובנה אחד, עם מספר.</p>';
      r.prompts=[['טבלת ציר ראשונה','איך בונים ב-'+app+' טבלת ציר שמראה '+p0.title+'? צעד אחר צעד. העמודות: '+names+'.'],
                 ['תובנה','מתוך הטבלה הזו: מה הממצא האחד החשוב? משפט אחד עם מספר, ומה עלול להפריך אותו.']];
      r.hints=[pivHow(f,p0),'ממיינים מהגדול לקטן. השורה הראשונה היא המועמדת לתובנה. את טבלת הציר מדביקים ל-AI לפני פרומפט התובנה.'];
      r.sol='<p>בקובץ המועשר אמור לצאת:</p>'+pivTable(p0);
    }else if(L==='m'){
      r.task='<p>בנו שתי טבלאות ציר: <b>'+esc(p0.title)+'</b>, ו<b>'+esc(p1.title)+'</b>. הוסיפו מבט שלישי: <b>'+esc(p2.title)+'</b>. נסחו תובנה אחת, והעבירו אותה חמש בדיקות.</p>';
      r.prompts=[['הצלבות','אילו 3 הצלבות יעניינו את '+f.role+'? לכל אחת, איך בונים את טבלת הציר ב-'+app+', צעד אחר צעד. העמודות: '+names+'.'],
                 ['תובנה','מתוך הטבלה הזו: מה הממצא האחד החשוב? משפט אחד עם מספר, ומה עלול להפריך אותו.']];
      r.hints=[pivHow(f,p0)+' בטבלה השנייה: '+pivHow(f,p1),'תובנה אומרת מה גבוה או נמוך, בכמה, וביחס למה. "יש הבדלים בין השכונות" היא לא תובנה.'];
      r.sol='<p>בקובץ המועשר אמור לצאת:</p>'+pivTable(p0)+pivTable(p1)+pivTable(p2);
    }else{
      r.task='<p>ענו על שאלה שמשלבת שתי עמודות או יותר: <b>'+esc(f.adv)+'</b> בדקו חריגות, ונסחו תובנה עם המלצה.</p>';
      r.prompts=[['שאלה משולבת','השאלה: '+f.adv+' אילו עמודות צריך להצליב ב-'+app+', ואיך בודקים שהתוצאה לא נובעת מכמה שורות חריגות? צעד אחר צעד. העמודות: '+names+'.'],
                 ['תובנה','מתוך הטבלה הזו: מה הממצא האחד החשוב? משפט אחד עם מספר, ומה עלול להפריך אותו.']];
      r.hints=['מתחילים מטבלת ציר עם שדה אחד בשורות ושדה שני בעמודות.','בודקים כמה שורות עומדות מאחורי כל תא. תא שנשען על שלוש שורות הוא לא ממצא.'];
      r.sol='<p>אין כאן תשובה אחת. נקודת פתיחה, מהקובץ המועשר:</p>'+pivTable(p0)+pivTable(p1);
    }
    r.check='<p>תובנה אחת במשפט, עם מספר, שעברה את חמש הבדיקות בכלי <a class="ln" href="../../insight/">חילוץ תובנה</a>.</p>';
    r.bonus='<p>שאלו את ה-AI מה עלול להפריך את התובנה שלכם, ובדקו טענה אחת מהרשימה בגיליון.</p>';
  }
  if(n===5){
    var kn=f.kpis.map(function(k){ return k.name; }).join(', ');
    if(L==='b'){
      r.task='<p>כתבו דו"ח של עמוד אחד ל'+esc(f.role)+': שלושה ממצאים, שלושה מספרים, המלצה אחת.</p>';
      r.prompts=[['דו"ח למנהל','כתוב דו"ח של עמוד ל'+f.role+' מתוך הנתונים האלה: 3 ממצאים עם מספר, המלצה אחת. שפה פשוטה, בלי סופרלטיבים.']];
      r.hints=['מדביקים ל-AI את טבלאות הציר מסבב 4, לא את כל הקובץ. כך הוא כותב על מה שמצאתם.','מוחקים כל משפט שאין בו מספר או המלצה.'];
    }else if(L==='m'){
      r.task='<p>בנו דף מצב בגיליון חדש: שלושה מדדים בראש ('+esc(kn)+'), שני גרפים ('+esc(f.charts.join(', '))+'), ומסנן (Slicer) על "'+esc(f.filters[0])+'".</p>';
      r.prompts=[['דף מצב בגיליון','איך בונים ב-'+app+', בגיליון חדש, דף מצב מהטבלה: בראש שלושה מדדים ('+kn+'), מתחת שני גרפים ('+f.charts.join(', ')+'), ומסנן על '+f.filters.join(', ')+'? צעד אחר צעד. העמודות: '+names+'.']];
      r.hints=['מדד הוא תא אחד עם נוסחה אחת, בגופן גדול.','הגרפים נבנים מטבלאות ציר, והמסנן מתחבר לטבלאות הציר. כך לחיצה אחת מעדכנת את כל הדף.'];
    }else{
      r.task='<p>בנו דשבורד HTML אינטראקטיבי מהקובץ: מדדים, מסננים וגרפים בקובץ אחד שנפתח בדפדפן. אפשרות אחרת: כלל שמסמן חריגות בגיליון אוטומטית, לפי העמודה "'+esc(C.name)+'".</p>';
      r.prompts=[['דשבורד','בנה דשבורד HTML בעברית (RTL) מהקובץ: מדדים '+kn+', מסננים '+f.filters.join(', ')+', גרפים '+f.charts.join(', ')+'. קובץ אחד, בלי ספריות חיצוניות, שעובד גם בלי אינטרנט.']];
      r.hints=['את הדשבורד בונים בצ\'אט, עם הקובץ המועשר, ולא בחלונית שבתוך הגיליון. שומרים את התשובה כקובץ ‎.html ופותחים בדפדפן.',
               'משהו לא עובד? מדביקים ל-AI את השגיאה, או מתארים מה רואים, ומבקשים את כל הקובץ מחדש.'];
      r.tip='<p><b>סקריפט או אוטומציה נחסמים במחשב של הרשות?</b> עוברים לדשבורד ה-HTML: הוא קובץ רגיל, ולא צריך בשבילו הרשאות.</p>';
    }
    r.sol='<p>שלושת המספרים, מהקובץ המועשר:</p>'+kpiList(f);
    r.check='<p>שלושה מספרים במוצר שלכם, שאימתם מול הגיליון. אלה המדדים של הקובץ הזה:</p>'+kpiList(f);
    r.bonus='<p>יש לכם ממצא והמלצה? הכלי <a class="ln" href="../../pitch/">מהנתון למצגת</a> הופך אותם לפרומפט שבונה מצגת.</p>';
  }
  if(n>=4) r.stuck=['הקובץ המועשר, מסוף סבב 3',num(f.st.clean)+' שורות, עם '+f.cols.length+' העמודות החדשות',UP+'files/'+f.n+'-rich.xlsx',f.n+' '+f.short+' - מועשר.xlsx'];
  return r;
}

function steps(n,T,f){
  var open={
    copilot:['פותחים את הקובץ ב-Excel, בחשבון העבודה.','פותחים את חלונית Copilot בתוך הגיליון.'],
    gemini:['מעלים את הקובץ ל-Google Drive ופותחים אותו ב-Google Sheets.','פותחים את חלונית Gemini בתוך הגיליון.'],
    chat:['בקובץ התרגול אין נתונים אמיתיים. בקובץ מהעבודה, קודם <a class="ln" href="../../data/anonymize/">מתממים</a>.','מעלים את הקובץ לצ\'אט. את הגיליון עצמו משאירים פתוח ב-Excel או ב-Sheets.']
  }[T];
  var inSheet=T!=='chat';
  var rest={
    1:['מדביקים את הפרומפט, ומבקשים נוסחה.', inSheet?'מוסיפים את הנוסחה לתא ריק ליד הטבלה, ומשווים לתשובה.':'מעתיקים את הנוסחה מהתשובה, מדביקים בתא ריק בגיליון, ומשווים לתשובה.'],
    2:['מדביקים את פרומפט האבחון, וקוראים את רשימת הבעיות.','מסירים כפילויות בגיליון עצמו: בתפריט "נתונים" יש פעולה להסרת כפילויות.','סופרים שורות לפני ואחרי.'],
    3:['מדביקים את הפרומפט, ומקבלים נוסחה לשורה 2.','מדביקים את הנוסחה בעמודה חדשה בסוף הטבלה, וגוררים למטה.','בודקים שלוש שורות ביד.'],
    4:['מדביקים את הפרומפט, ומקבלים הוראות לטבלת ציר.','בונים את טבלת הציר בגיליון: מסמנים את הטבלה, ובתפריט "הוספה" בוחרים טבלת ציר.', inSheet?'מבקשים מה-AI משפט תובנה על טבלת הציר.':'מדביקים את טבלת הציר לצ\'אט, ומבקשים משפט תובנה.'],
    5:[inSheet?'מסמנים את טבלאות הציר מסבב 4, ומדביקים את הפרומפט.':'מדביקים לצ\'אט את טבלאות הציר מסבב 4, ואת הפרומפט.','בונים את המוצר: דו"ח, דף מצב או דשבורד.','מאמתים שלושה מספרים מול הגיליון.']
  }[n];
  var note={copilot:'Copilot לא מגיב לקובץ? בדרך כלל הקובץ צריך להיות שמור ב-OneDrive או ב-SharePoint, עם שמירה אוטומטית. אין רישיון? עוברים למסלול הצ\'אט.',
            gemini:'אין חלונית Gemini בגיליון? היא תלויה במנוי של הארגון. עוברים למסלול הצ\'אט.',
            chat:'הצ\'אט לא רואה את הגיליון שלכם, רק את הקובץ שהעליתם. כל נוסחה בודקים בגיליון.'}[T];
  return '<ol>'+open.concat(rest).map(function(s){ return '<li>'+s+'</li>'; }).join('')+'</ol><p class="out">'+note+'</p>';
}

function round(){
  var c=choice(), f=F[c.file], n=ROUND, R=ROUNDS[n], r=build(n,c,f), li=LEVELS.map(function(l){return l[0];}).indexOf(c.level);
  function sel(id,lab,list,cur,cls){ return '<div><label for="'+id+'">'+lab+'</label><select id="'+id+'" class="'+(cls||'')+'">'+list.map(function(o){ return '<option value="'+o[0]+'"'+(o[0]===cur?' selected':'')+'>'+esc(o[1])+'</option>'; }).join('')+'</select></div>'; }
  var files=Object.keys(F).sort().map(function(k){ return [k,k+' · '+F[k].short]; });
  var h='<div class="bar3 lv-'+c.level+'">'+sel('sFile','הקובץ שלי',files,c.file)+sel('sLevel','הרמה שלי',LEVELS,c.level,'lvsel')+sel('sTool','הכלי שלי',TOOLS,c.tool)+'</div>';
  h+='<nav class="strip" aria-label="חמשת הסבבים">'+[1,2,3,4,5].map(function(i){ return '<a href="../r'+i+'/"'+(i===n?' aria-current="step"':'')+'><b>'+i+'</b><span>'+ROUNDS[i].t+'</span></a>'; }).join('')+'</nav>';
  h+='<div class="rhead"><h1>סבב '+n+': '+R.t+'</h1><p>'+R.s+'</p></div>';
  h+='<section class="box task lv-'+c.level+'"><h2>המשימה <span class="lvtag">רמה '+label(LEVELS,c.level)+'</span> <span class="out">קובץ '+f.n+': '+esc(f.short)+'</span></h2>'+r.task+(r.ex?'<div class="ex">'+r.ex+'</div>':'')+
     '<div class="shift"><button type="button" id="easier"'+(li===0?' disabled':'')+'>קשה לי: רמה אחת למטה</button><button type="button" id="harder"'+(li===2?' disabled':'')+'>קל לי: רמה אחת למעלה</button></div></section>';
  h+='<section class="box"><h2>השלבים: '+esc(label(TOOLS,c.tool))+'</h2>'+steps(n,c.tool,f)+'</section>';
  h+='<section class="box"><h2>פרומפט להעתקה</h2>'+r.prompts.map(function(p){ return prompt(p[0],p[1]); }).join('')+'</section>';
  if(r.tip) h+='<section class="box tip">'+r.tip+'</section>';
  h+='<section class="box"><h2>נתקעתם באמצע?</h2>'+
     '<details class="hint"><summary>רמז 1</summary><div class="in"><p>'+r.hints[0]+'</p></div></details>'+
     '<details class="hint locked"><summary>רמז 2</summary><div class="in"><p>'+r.hints[1]+'</p></div></details>'+
     '<details class="hint sol locked"><summary>פתרון</summary><div class="in">'+r.sol+'</div></details></section>';
  h+='<section class="box check"><h2>נקודת ביקורת: מה צריך להיות על המסך כדי להמשיך</h2>'+r.check+'</section>';
  h+='<section class="box"><h2>סיימתם מוקדם? בונוס</h2>'+r.bonus+'</section>';
  h+='<section class="box stuck"><span class="t">'+(n<=2?'הקובץ לסבב הזה':'לא הגעתם לכאן עם קובץ מוכן?')+'</span><span class="d">'+r.stuck[0]+' · '+r.stuck[1]+'</span>'+dl(r.stuck[2],r.stuck[3],'הורדה ←')+'</section>';
  h+='<details class="hint" style="margin-top:14px"><summary>שני פרומפטים שעוזרים בכל סבב</summary><div class="in">'+
     prompt('תקלה','קיבלתי את השגיאה הזאת: [מדביקים כאן את השגיאה]. תקן ותן לי את הכל מחדש.')+
     prompt('בדיקה עצמית','אילו מספרים בתשובה חישבת בעצמך, ואילו לקחת מנוסחה? מה כדאי לאמת ידנית?')+'</div></details>';
  h+='<div class="pn">'+(n>1?'<a href="../r'+(n-1)+'/">→ סבב '+(n-1)+': '+ROUNDS[n-1].t+'</a>':'<a href="../">→ לדף הסבבים</a>')+
     (n<5?'<a class="nx" href="../r'+(n+1)+'/">סבב '+(n+1)+': '+ROUNDS[n+1].t+' ←</a>':'<a class="nx" href="../my-file/">הקובץ שלי בשבוע הבא ←</a>')+'</div>';
  var app=document.getElementById('app'); app.innerHTML=h;
  function set(k,v){ var x=choice(); x[k]=v; save(x); var y=window.scrollY; round(); window.scrollTo(0,y); }
  document.getElementById('sFile').onchange=function(){ set('file',this.value); };
  document.getElementById('sLevel').onchange=function(){ set('level',this.value); };
  document.getElementById('sTool').onchange=function(){ set('tool',this.value); };
  document.getElementById('easier').onclick=function(){ set('level',LEVELS[li-1][0]); };
  document.getElementById('harder').onclick=function(){ set('level',LEVELS[li+1][0]); };
  var hs=[].slice.call(app.querySelectorAll('section .hint'));
  hs.forEach(function(d,i){
    d.querySelector('summary').addEventListener('click',function(e){ if(d.classList.contains('locked')) e.preventDefault(); });
    d.addEventListener('toggle',function(){ if(d.open&&hs[i+1]) hs[i+1].classList.remove('locked'); });
  });
  if(!c.set) save(c);
}

if(ROUND) round();
else if(document.getElementById('pick')){ identity(); home(); }
window.DataAI={choice:choice,files:F,levels:LEVELS,tools:TOOLS,rounds:ROUNDS,rules:RULES,esc:esc,build:build,steps:steps};
})();
