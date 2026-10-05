/* RomanAI site shell, loaded by every page:
   1. Top strip: RomanAI wordmark + a clear way back. Configure on the script tag:
        data-crumbs='כל ההכשרות|../../;קבלת החלטות מבוססת נתונים|../'   (label|href pairs, outermost first)
      The last pair is where the "חזרה" button goes. href "auto" resolves from ?from=maazim|ai.
   2. The seven floating data symbols (brand signature): in every hero (.cover / [data-sym]) and in the footer.
   3. Footer: symbols strip and credit.
   4. Small "?" help: <button class="help" data-help="הסבר קצר">?</button> (or SiteHelp.html('...') in templates). */
(function(){
  'use strict';
  var me=document.currentScript, base=new URL('.', me.src).href;   // .../assets/

  /* ---------- data symbols ---------- */
  function symbols(dark){
    var a=dark?'#FFFDF8':'#1F5F5B', b=dark?'#C9E2DE':'#B4502F', c=dark?'#F0B95A':'#D99A2B';
    var S={
      bars:'<path d="M8 40V26M18 40V14M28 40V22M38 40V8" stroke="'+a+'" stroke-width="5" stroke-linecap="round"/><path d="M4 42h40" stroke="'+b+'" stroke-width="2.5" stroke-linecap="round"/>',
      table:'<rect x="6" y="8" width="36" height="32" rx="5" stroke="'+b+'" stroke-width="3" fill="none"/><path d="M6 18h36M6 29h36M19 8v32" stroke="'+b+'" stroke-width="2.5"/><rect x="6" y="8" width="36" height="10" rx="5" fill="'+a+'" opacity=".55"/>',
      line:'<path d="M6 36l10-10 8 6 16-18" stroke="'+a+'" stroke-width="4" fill="none" stroke-linecap="round" stroke-linejoin="round"/><circle cx="40" cy="14" r="4" fill="'+c+'"/><path d="M4 42h40" stroke="'+b+'" stroke-width="2.5" stroke-linecap="round"/>',
      donut:'<circle cx="24" cy="24" r="15" stroke="'+b+'" stroke-width="7" fill="none" opacity=".45"/><path d="M24 9a15 15 0 0 1 14.3 19.6" stroke="'+c+'" stroke-width="7" fill="none"/>',
      scatter:'<path d="M6 6v36h36" stroke="'+b+'" stroke-width="2.5" fill="none" stroke-linecap="round"/><circle cx="14" cy="32" r="3.5" fill="'+a+'"/><circle cx="22" cy="26" r="3.5" fill="'+a+'"/><circle cx="30" cy="20" r="3.5" fill="'+c+'"/><circle cx="37" cy="12" r="3.5" fill="'+a+'"/>',
      pin:'<path d="M24 43s13-12.5 13-23a13 13 0 0 0-26 0c0 10.5 13 23 13 23z" stroke="'+a+'" stroke-width="3.5" fill="none"/><circle cx="24" cy="20" r="5" fill="'+c+'"/>',
      db:'<ellipse cx="24" cy="12" rx="14" ry="5" stroke="'+b+'" stroke-width="3" fill="none"/><path d="M10 12v24c0 2.8 6.3 5 14 5s14-2.2 14-5V12M10 24c0 2.8 6.3 5 14 5s14-2.2 14-5" stroke="'+b+'" stroke-width="3" fill="none"/>'
    };
    return S;
  }
  var ORDER=['bars','table','line','donut','scatter','pin','db'];
  function one(name,dark){ return '<svg viewBox="0 0 48 48" fill="none" aria-hidden="true">'+(symbols(dark)[name]||symbols(dark).bars)+'</svg>'; }
  function strip(dark){ return '<div class="sym '+(dark?'dark':'light')+'" aria-hidden="true">'+ORDER.map(function(n){ return one(n,dark); }).join('')+'</div>'; }
  function paintIcons(root){
    (root||document).querySelectorAll('[data-ico]').forEach(function(el){ if(!el.firstChild) el.innerHTML=one(el.dataset.ico,false); });
    (root||document).querySelectorAll('[data-sym]').forEach(function(el){ if(!el.firstChild) el.innerHTML=strip(el.dataset.sym==='dark'); });
  }

  /* ---------- top strip ---------- */
  function topbar(){
    var conf=me.dataset.crumbs, home=me.dataset.home||'./';
    var pairs=conf?conf.split(';').map(function(p){ var a=p.split('|'); return {t:a[0],h:a[1]}; }):[];
    pairs.forEach(function(p){
      if(p.h!=='auto') return;
      var from=new URLSearchParams(location.search).get('from');
      if(from==='maazim'){ p.t='מאיצים דיגיטליים'; p.h='../maazim/'; } else { p.t='הכשרת בינה מלאכותית'; p.h='../ai/'; }
    });
    if(pairs.length) home=pairs[0].h;
    var nav=document.createElement('nav'); nav.className='topbar'; nav.setAttribute('aria-label','ניווט');
    var html='<a class="mark" href="'+home+'" title="לדף הבית">Roman<span>AI</span></a>';
    if(pairs.length){
      var up=pairs[pairs.length-1], here=me.dataset.here||document.title.split(/[—·]/)[0].trim();
      html+='<a class="back" href="'+up.h+'">→ חזרה ל'+up.t+'</a>'+
        '<span class="trail">'+pairs.map(function(p){ return '<a href="'+p.h+'">'+p.t+'</a><span class="sep">‹</span>'; }).join('')+'<span class="here">'+here+'</span></span>';
    }
    nav.innerHTML=html;
    document.body.insertBefore(nav,document.body.firstChild);
    document.body.classList.add('has-topbar');
    /* a tool header that only repeats the page title is noise; keep it when it holds controls */
    document.querySelectorAll('header.bar').forEach(function(h){
      if(!h.querySelector('.bar-actions,.steps,button')) h.classList.add('plain');
    });
  }

  /* ---------- hero + footer ---------- */
  function hero(){
    document.querySelectorAll('.cover').forEach(function(c){ c.insertAdjacentHTML('beforeend',strip(true)); });
  }
  function footer(){
    var f=document.createElement('footer'); f.className='brandfoot';
    f.innerHTML='<div class="in">'+
      strip(false)+
      '<div class="by">נבנה ע"י רומן גרינשטיין · <a href="mailto:romangash@gmail.com">romangash@gmail.com</a></div></div>';
    var old=document.querySelector('footer.credit');
    if(old) old.parentNode.insertBefore(f,old); else document.body.appendChild(f);
  }

  /* ---------- "?" help ---------- */
  var pop=null, openBtn=null;
  function closeHelp(){ if(pop){ pop.remove(); pop=null; } if(openBtn){ openBtn.setAttribute('aria-expanded','false'); openBtn=null; } }
  document.addEventListener('click',function(e){
    var b=e.target.closest&&e.target.closest('.help');
    if(!b){ closeHelp(); return; }
    e.preventDefault(); e.stopPropagation();
    if(openBtn===b){ closeHelp(); return; }
    closeHelp();
    pop=document.createElement('div'); pop.className='help-pop'; pop.setAttribute('role','tooltip'); pop.innerHTML=b.dataset.help;
    document.body.appendChild(pop);
    var r=b.getBoundingClientRect(), w=pop.offsetWidth;
    var left=Math.min(Math.max(8, r.left+r.width/2-w/2), document.documentElement.clientWidth-w-8);
    pop.style.left=(left+window.scrollX)+'px'; pop.style.top=(r.bottom+8+window.scrollY)+'px';
    openBtn=b; b.setAttribute('aria-expanded','true');
  });
  document.addEventListener('keydown',function(e){ if(e.key==='Escape') closeHelp(); });
  function helpHtml(text){ return '<button type="button" class="help" aria-label="הסבר" aria-expanded="false" data-help="'+String(text).replace(/&/g,'&amp;').replace(/"/g,'&quot;')+'">?</button>'; }

  /* help marks for the older tools, whose markup is rendered by their own scripts */
  var AUTO=[
    ['[data-page="role-worksheet"] .titles h1','דף עבודה אישי. ממלאים שלושה חלקים, ובסוף לוחצים <b>תמונה</b> כדי להוריד. בצד יש טיימר לתרגיל: בוחרים 3, 5 או 7 דקות ולוחצים <b>התחל</b>.'],
    ['[data-page="role-worksheet"] .s1 .sh h3','שם התפקיד כמו שהייתם מציגים אותו למישהו מחוץ לרשות. משפט אחד מספיק.'],
    ['[data-page="role-worksheet"] .s2 .sh h3','דברים שלוקחים ימים או שבועות ושיש להם תוצר: דוח, תוכנית, מכרז, אירוע.'],
    ['[data-page="role-worksheet"] .s3 .sh h3','דברים קטנים שחוזרים כל יום או כל שבוע: מיילים, פניות, עדכון טבלאות.'],
    ['[data-page="role-worksheet"] .aihint','הכפתור <b>AI</b> ליד כל משימה מסמן אותה כמשימה שבה AI יכול לעזור. אפשר לסמן ולבטל.'],
    ['[data-page="insight"] .checks-head h3','חמש שאלות ששואלים על כל תובנה לפני שסומכים עליה. מסמנים כל בדיקה שעשיתם, וכותבים מה מצאתם.']
  ];
  var autoTimer;
  function autoHelp(){
    AUTO.forEach(function(a){
      document.querySelectorAll(a[0]).forEach(function(el){ if(!el.querySelector('.help')) el.insertAdjacentHTML('beforeend',helpHtml(a[1])); });
    });
  }
  function init(){
    var seg=location.pathname.replace(/index\.html$/,'').split('/').filter(Boolean);
    document.body.dataset.page=seg[seg.length-1]||'home';
    topbar(); hero(); footer(); paintIcons(); autoHelp();
    new MutationObserver(function(){ clearTimeout(autoTimer); autoTimer=setTimeout(autoHelp,120); }).observe(document.body,{childList:true,subtree:true});
  }
  if(document.body) init(); else document.addEventListener('DOMContentLoaded',init);
  window.SiteHelp={html:helpHtml};
  window.SiteSym={paint:paintIcons, strip:strip};
})();
