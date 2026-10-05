/* Shared helpers for the new tools of "קבלת החלטות מבוססת נתונים".
   Everything participants fill in stays in their browser (localStorage). */
window.DD = (function(){
  'use strict';
  function esc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function load(key, fallback){
    try{ var s=JSON.parse(localStorage.getItem(key)); if(s) return Object.assign(fallback, s); }catch(e){}
    return fallback;
  }
  function save(key, data){ try{ localStorage.setItem(key, JSON.stringify(data)); }catch(e){} }
  function identity(){
    try{ var s=JSON.parse(localStorage.getItem('toolkit_identity')||'{}'); return [s.name,s.org,s.dept,s.role].filter(Boolean).join(' · '); }
    catch(e){ return ''; }
  }
  var tt;
  function toast(m){
    var el=document.getElementById('toast'); if(!el) return;
    el.textContent=m; el.classList.add('show'); clearTimeout(tt);
    tt=setTimeout(function(){ el.classList.remove('show'); },2400);
  }
  function copy(text, msg){
    (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject())
      .then(function(){ toast(msg||'הועתק ✓'); })
      .catch(function(){ window.prompt('העתיקו:', text); });
  }
  function download(name, content, type){
    var b=new Blob([content],{type:type||'application/json'}), u=URL.createObjectURL(b), a=document.createElement('a');
    a.href=u; a.download=name; document.body.appendChild(a); a.click();
    setTimeout(function(){ a.remove(); URL.revokeObjectURL(u); },1500);
  }
  function safeName(s, fallback){ return (s||fallback).replace(/[\\/:*?"<>|]/g,'').slice(0,50); }
  /* Word-openable document from simple HTML */
  function downloadDoc(name, title, bodyHtml){
    var html='<html dir="rtl"><head><meta charset="utf-8"><title>'+esc(title)+'</title>'+
      '<style>body{font-family:Arial,sans-serif;direction:rtl;line-height:1.6}h1{color:#8F3D22}h2{color:#1F5F5B;border-bottom:1px solid #ccc;padding-bottom:3px;margin-top:22px}table{border-collapse:collapse;width:100%}td,th{border:1px solid #bbb;padding:5px 7px;text-align:right}</style>'+
      '</head><body>'+bodyHtml+'<p style="color:#888;font-size:11px;margin-top:30px">נבנה באתר המלווה · קבלת החלטות מבוססת נתונים · רומן גרינשטיין</p></body></html>';
    download(safeName(name,'מסמך')+'.doc', '﻿'+html, 'application/msword');
  }
  function loadFile(cb){
    var inp=document.createElement('input'); inp.type='file'; inp.accept='application/json,.json';
    inp.onchange=function(){ var f=inp.files[0]; if(!f) return; var r=new FileReader();
      r.onload=function(){ try{ cb(JSON.parse(r.result)); toast('נטען ✓'); }catch(e){ toast('קובץ לא תקין'); } }; r.readAsText(f); };
    inp.click();
  }
  function tabs(root, key){
    var btns=root.querySelectorAll('.tabs [data-tab]'), panes=root.querySelectorAll('.tab-pane');
    function show(id){
      btns.forEach(function(b){ b.classList.toggle('on', b.dataset.tab===id); });
      panes.forEach(function(p){ p.style.display = p.id===id ? '' : 'none'; });
      try{ localStorage.setItem(key, id); history.replaceState(null,'','#'+id); }catch(e){}
    }
    btns.forEach(function(b){ b.addEventListener('click', function(){ show(b.dataset.tab); window.scrollTo({top:0}); }); });
    var start=(location.hash||'').slice(1);
    if(!document.getElementById(start)){ try{ start=localStorage.getItem(key)||''; }catch(e){ start=''; } }
    show(document.getElementById(start) ? start : btns[0].dataset.tab);
  }
  var QKEY='dd_question_v3';
  /* "כדי לעזור ל[מי מחליט] להחליט [איזו החלטה] — [מה מודדים] [מתי ואיפה]?"  f(value, placeholder) formats each part */
  function questionSentence(q, f){
    return 'כדי לעזור ל'+f(q.who,'מי מחליט')+' להחליט '+f(q.decision,'איזו החלטה')+' — '+f(q.measure,'מה מודדים')+' '+f(q.scope,'מתי ואיפה')+'?';
  }
  /* the business question written in the question builder, as plain text ('' when not written yet) */
  function businessQuestion(){
    try{
      var q=(JSON.parse(localStorage.getItem(QKEY)||'{}').q)||{};
      if((q.own||'').trim()) return q.own.trim();
      if(!(q.who||q.decision||q.measure||q.scope)) return '';
      return questionSentence(q, function(v,ph){ return (v||'').trim()||'['+ph+']'; });
    }catch(e){ return ''; }
  }
  /* the directions written in part 3 of the question builder, one item per line: {internal:[], open:[], field:[]} */
  function builderIdeas(){
    var out={internal:[],open:[],field:[]};
    try{ var d=(JSON.parse(localStorage.getItem(QKEY)||'{}').ideas)||{}; Object.keys(out).forEach(function(k){ out[k]=String(d[k]||'').split(/\n/).map(function(x){ return x.trim(); }).filter(Boolean); }); }catch(e){}
    return out;
  }
  function help(text){ return '<button type="button" class="help" aria-label="הסבר" aria-expanded="false" data-help="'+esc(text)+'">?</button>'; }
  return {questionSentence:questionSentence, businessQuestion:businessQuestion, builderIdeas:builderIdeas, help:help, esc:esc, load:load, save:save, identity:identity, toast:toast, copy:copy, download:download, downloadDoc:downloadDoc, loadFile:loadFile, safeName:safeName, tabs:tabs};
})();
