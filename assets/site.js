/* Site-wide helpers, loaded by every page:
   1. Back strip / breadcrumbs. Configure on the script tag:
        data-crumbs='כל ההכשרות|../../;קבלת החלטות מבוססת נתונים|../'   (label|href pairs, outermost first)
        data-here='שם העמוד'
      The last pair is where the big "חזרה" button goes. Use href "auto" to resolve
      from ?from=maazim|ai (the shared role-worksheet page).
   2. Small animated data drawings: any element with data-viz="bars|hbars|line|donut|table|dash|scatter"
      gets an animated SVG. data-side="1" on the script tag adds floating ornaments in the page margins. */
(function(){
  'use strict';
  var me=document.currentScript;

  /* ---------- drawings ---------- */
  var d=function(i,step){ return 'style="animation-delay:'+(-(i*step)).toFixed(2)+'s"'; };
  var VIZ={
    bars:function(){
      var h=[34,48,62,40,26];
      return h.map(function(v,i){ return '<rect class="bar '+(i===2?'h':'g')+'" '+d(i,.7)+' x="'+(12+i*20)+'" y="'+(70-v)+'" width="14" height="'+v+'" rx="2.5"/>'; }).join('')+
        '<line class="ax" x1="6" y1="71" x2="114" y2="71"/>';
    },
    hbars:function(){
      var w=[92,70,54,38,24];
      return w.map(function(v,i){ return '<rect class="hbar '+(i===0?'h':'g')+'" '+d(i,.8)+' x="'+(110-v)+'" y="'+(8+i*14)+'" width="'+v+'" height="9" rx="2.5"/>'; }).join('')+
        '<line class="ax" x1="111" y1="4" x2="111" y2="78"/>';
    },
    line:function(){
      var p='M8 60 L26 50 L44 56 L62 30 L80 38 L98 16 L112 22';
      return '<path class="ln2" d="M8 66 L26 62 L44 64 L62 54 L80 58 L98 48 L112 50"/>'+
        '<path class="ln draw" pathLength="1" d="'+p+'"/>'+
        '<circle class="a" r="4.5"><animateMotion dur="5s" repeatCount="indefinite" path="'+p+'"/></circle>'+
        '<line class="ax" x1="6" y1="72" x2="114" y2="72"/>';
    },
    donut:function(){
      return '<circle class="ring g" cx="60" cy="40" r="26"/>'+
        '<circle class="ring h" cx="60" cy="40" r="26" stroke-dasharray="164" stroke-dashoffset="150" stroke-linecap="round"/>'+
        '<text x="60" y="46" text-anchor="middle" font-size="17" data-tick="63" data-suffix="%">63%</text>';
    },
    table:function(){
      var s='<rect class="h" x="8" y="8" width="104" height="10" rx="2"/>';
      for(var r=0;r<5;r++) for(var c=0;c<3;c++)
        s+='<rect class="row g" style="animation-delay:'+(r)+'s" x="'+(8+c*36)+'" y="'+(22+r*11)+'" width="'+(c===0?32:32)+'" height="7" rx="1.5"/>';
      return s;
    },
    dash:function(){
      return '<rect class="g" opacity=".35" x="4" y="4" width="52" height="34" rx="5"/>'+
        '<text x="30" y="27" text-anchor="middle" font-size="16" data-tick="136">136</text>'+
        '<rect class="g" opacity=".35" x="62" y="4" width="54" height="34" rx="5"/>'+
        [14,22,10,18].map(function(v,i){ return '<rect class="bar '+(i===1?'h':'g')+'" '+d(i,.6)+' x="'+(69+i*11)+'" y="'+(33-v)+'" width="7" height="'+v+'" rx="1.5"/>'; }).join('')+
        '<rect class="g" opacity=".35" x="4" y="44" width="112" height="32" rx="5"/>'+
        '<path class="ln draw" pathLength="1" d="M10 68 L28 60 L46 64 L64 52 L82 57 L100 48 L110 51"/>';
    },
    scatter:function(){
      var pts=[[16,60],[28,54],[38,58],[46,44],[58,46],[66,34],[78,38],[88,24],[100,26],[108,14]];
      return pts.map(function(p,i){ return '<circle class="bob '+(i===7?'a':'h')+'" '+d(i,.35)+' cx="'+p[0]+'" cy="'+p[1]+'" r="'+(i===7?5.5:4)+'" opacity=".85"/>'; }).join('')+
        '<line class="ax" x1="6" y1="72" x2="114" y2="72"/><line class="ax" x1="6" y1="6" x2="6" y2="72"/>';
    }
  };
  function svg(type){ return '<svg class="vz" viewBox="0 0 120 80" aria-hidden="true">'+(VIZ[type]||VIZ.bars)()+'</svg>'; }
  function paint(root){
    (root||document).querySelectorAll('[data-viz]').forEach(function(el){
      if(el.dataset.vizDone) return; el.dataset.vizDone='1'; el.innerHTML=svg(el.dataset.viz);
    });
  }
  /* numbers that keep drifting a little */
  setInterval(function(){
    document.querySelectorAll('[data-tick]').forEach(function(t){
      var base=+t.dataset.tick, v=Math.round(base+(Math.random()-.5)*base*.08);
      t.textContent=v+(t.dataset.suffix||'');
    });
  },1400);

  /* ---------- back strip ---------- */
  function crumbs(){
    var conf=me&&me.dataset.crumbs; if(!conf) return;
    var pairs=conf.split(';').map(function(p){ var a=p.split('|'); return {t:a[0],h:a[1]}; });
    pairs.forEach(function(p){
      if(p.h!=='auto') return;
      var from=new URLSearchParams(location.search).get('from');
      if(from==='maazim'){ p.t='מאיצים דיגיטליים'; p.h='../maazim/'; } else { p.t='הכשרת בינה מלאכותית'; p.h='../ai/'; }
    });
    var up=pairs[pairs.length-1], here=me.dataset.here||document.title.split(/[—·]/)[0].trim();
    var nav=document.createElement('nav'); nav.className='crumbs'; nav.setAttribute('aria-label','ניווט');
    nav.innerHTML='<a class="back" href="'+up.h+'">→ חזרה ל'+up.t+'</a>'+
      '<span class="trail">'+pairs.map(function(p){ return '<a href="'+p.h+'">'+p.t+'</a> <span class="sep">‹</span> '; }).join('')+'<span class="here">'+here+'</span></span>'+
      '<span class="spark vz-dark" data-viz="line"></span>';
    document.body.insertBefore(nav,document.body.firstChild);
    document.body.classList.add('has-crumbs');
  }
  function sides(){
    if(!me||!me.dataset.side) return;
    [['r',['bars','donut','table']],['l',['line','dash','scatter']]].forEach(function(s){
      var el=document.createElement('div'); el.className='vz-side '+s[0];
      el.innerHTML=s[1].map(function(t){ return '<div class="vz-card" data-viz="'+t+'"></div>'; }).join('');
      document.body.appendChild(el);
    });
  }
  function init(){ crumbs(); sides(); paint(); }
  if(document.body) init(); else document.addEventListener('DOMContentLoaded',init);
  window.SiteViz={paint:paint};
})();
