#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从引擎输出生成交互式确认页（含 localStorage 自动保存）"""
import csv
import json
import os
import sys
import html as H

TPL = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITLE__</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Helvetica Neue",sans-serif;background:#fafaf9;color:#1c1917;line-height:1.6;padding:24px 16px 200px}
.wrap{max-width:1040px;margin:0 auto}
h1{font-size:22px;font-weight:500;margin-bottom:4px}
.sub{font-size:13px;color:#78716c;margin-bottom:6px}
.sub2{font-size:12px;color:#a8a29e;margin-bottom:18px}
.bar{position:sticky;top:0;background:#fafaf9;padding:12px 0;border-bottom:1px solid #e7e5e4;margin-bottom:18px;z-index:20;display:flex;align-items:flex-end;gap:16px;flex-wrap:wrap}
.stats{display:flex;gap:18px;flex-wrap:wrap;flex:1}
.stat b{font-weight:500;font-size:19px;display:block;line-height:1.2}
.stat span{color:#78716c;font-size:11.5px}
.btn{background:#1c1917;color:#fff;border:none;padding:10px 15px;border-radius:8px;font-size:13.5px;cursor:pointer;font-family:inherit;white-space:nowrap}
.btn:hover{background:#44403c}
.btn.ghost{background:#fff;color:#1c1917;border:1px solid #d6d3d1}
.saved{font-size:11.5px;color:#15803d;padding:5px 9px;background:#f0fdf4;border-radius:6px;display:none}
.saved.on{display:inline-block}
.card{background:#fff;border:1px solid #e7e5e4;border-radius:12px;padding:16px;margin-bottom:12px}
.card.done{background:#f0fdf4;border-color:#86efac}
.card.custom{background:#f8fafc;border-color:#93c5fd}
.hd{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:10px}
.amt{font-size:20px;font-weight:500;white-space:nowrap}
.meta{font-size:13px;color:#57534e;margin-top:2px}
.tag{display:inline-block;font-size:11px;padding:2px 7px;border-radius:4px;background:#fef3c7;color:#92400e;margin-right:5px;margin-bottom:3px}
.tag.p0{background:#fee2e2;color:#991b1b}
.tag.p1{background:#fef3c7;color:#92400e}
.tag.p2{background:#f5f5f4;color:#57534e}
.tag.err{background:#dbeafe;color:#1e40af}
.tag.bk{background:#ede9fe;color:#5b21b6}
.tag.in{background:#dcfce7;color:#166534}
.tag.ok{background:#dcfce7;color:#166534}
.opts{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}
.opt{border:1px solid #d6d3d1;background:#fff;border-radius:8px;padding:9px 13px;font-size:13.5px;cursor:pointer;font-family:inherit;text-align:left;transition:.12s;line-height:1.35}
.opt:hover{border-color:#78716c;background:#fafaf9}
.opt.sel{background:#1c1917;color:#fff;border-color:#1c1917}
.opt.custom{background:#eff6ff;border-color:#60a5fa;color:#1e40af}
.opt small{display:block;font-size:11px;opacity:.6;margin-top:2px}
.hint{font-size:12px;color:#78716c;margin-top:9px}
.done-tag{color:#15803d;font-weight:500}
.link{background:none;border:none;color:#2563eb;font-size:12.5px;cursor:pointer;padding:6px 0;font-family:inherit;text-decoration:underline}
.notearea{margin-top:12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px}
.notearea .nt{font-size:12px;color:#57534e;font-weight:500;margin-bottom:7px}
.notearea .nt .lock{font-size:11px;color:#94a3b8;font-weight:400;margin-left:6px}
.quick{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}
.qtag{border:1px solid #cbd5e1;background:#fff;color:#475569;border-radius:999px;padding:4px 10px;font-size:12px;cursor:pointer;font-family:inherit}
.qtag:hover{background:#1c1917;color:#fff;border-color:#1c1917}
.noteinput{width:100%;padding:9px 11px;border:1px solid #cbd5e1;border-radius:7px;font-size:13px;font-family:inherit;background:#fff;resize:vertical;min-height:54px;line-height:1.6}
.noteinput:focus{outline:2px solid #93c5fd;border-color:#3b82f6}
.noteinput.filled{border-color:#86efac;background:#f0fdf4}
.editor{background:#eef2f7;border:1px solid #cbd5e1;border-radius:8px;padding:13px;margin-top:10px}
.frow{display:flex;gap:8px;align-items:center;margin-bottom:9px;flex-wrap:wrap}
.frow:last-child{margin-bottom:0}
.frow label{font-size:12.5px;color:#475569;min-width:54px;font-weight:500}
.frow select,.frow input{flex:1;min-width:140px;padding:8px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;font-family:inherit;background:#fff}
.path{font-size:11.5px;color:#64748b;background:#f1f5f9;padding:7px 10px;border-radius:6px;margin:-4px 0 9px}
.sect{font-size:15px;font-weight:500;margin:26px 0 10px;padding-bottom:8px;border-bottom:1px solid #e7e5e4}
table{width:100%;border-collapse:collapse;font-size:12.5px;background:#fff;border:1px solid #e7e5e4;border-radius:10px;overflow:hidden}
th{background:#f5f5f4;text-align:left;padding:8px 10px;font-weight:500;font-size:12px;color:#57534e}
td{padding:6px 10px;border-top:1px solid #f5f5f4}
tr.fixed td{background:#f8fafc}
#toast{position:fixed;bottom:26px;left:50%;transform:translateX(-50%);background:#1c1917;color:#fff;padding:12px 20px;border-radius:8px;font-size:13.5px;display:none;z-index:99;max-width:90vw}
#dock{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid #d6d3d1;padding:14px 20px;display:none;z-index:30;box-shadow:0 -2px 12px rgba(0,0,0,.08)}
#dock.on{display:block}
#dock .dt{font-size:13px;font-weight:500;margin-bottom:8px;display:flex;align-items:center;gap:8px}
#dock .db{font-size:12px;white-space:pre-wrap;max-height:130px;overflow:auto;background:#f8fafc;border:1px solid #e2e8f0;border-radius:7px;padding:10px;font-family:ui-monospace,Menlo,monospace;line-height:1.55}
#dock .da{display:flex;gap:8px;margin-top:9px;flex-wrap:wrap}
#dock .da .btn{padding:8px 14px;font-size:13px}
</style>
</head>
<body>
<div class="wrap">
<h1>__TITLE__</h1>
<div class="sub">__SUB__</div>
<div class="sub2">分类：__BK__ 账本 × 收支方向 × __C1__ 一级 / __C2__ 二级 · 账户 __ACC__ 个 · <span style="color:#15803d">内容自动保存，刷新不丢</span></div>
<div class="bar">
  <div class="stats">
    <div class="stat"><b id="s1">__N1__</b><span>待确认</span></div>
    <div class="stat"><b>__TOTAL__</b><span>总笔数</span></div>
    <div class="stat"><b>__AUTO__</b><span>自动归类</span></div>
    <div class="stat"><b>__EXP__</b><span>支出合计</span></div>
    <div class="stat"><b id="s5">0</b><span>手工修正</span></div>
  </div>
  <span class="saved" id="saved">已自动保存</span>
  <button class="btn" onclick="showOut()">生成结果</button>
  <button class="btn ghost" onclick="window.print()">打印</button>
  <button class="btn ghost" onclick="resetAll()">重置</button>
</div>
<div id="list"></div>
<div class="sect">已自动归类 __AUTO__ 笔 — 发现错误点「修正」</div>
<div id="tbl"></div>
</div>
<div id="dock">
  <div class="dt">确认结果 <button class="btn ghost" onclick="document.getElementById('dock').classList.remove('on')" style="padding:2px 8px;font-size:11px;margin-left:auto">收起</button></div>
  <div class="db" id="dockBody"></div>
  <div class="da">
    <button class="btn" onclick="copyOut()">复制全部</button>
    <button class="btn ghost" onclick="downloadTxt()">下载 txt</button>
  </div>
</div>
<div id="toast"></div>
<script>
var ROWS=__ROWS__, TREE=__TREE__, ACC=__ACC__;
var QUICK=['退款','报销','垫付','AA分摊','借出','还款','礼金','代买','退货','赔付','公积金','医保','公司','家人','朋友','出差','旅游','医疗','学习','其他'];
var NL=String.fromCharCode(10), KEY='__KEY__';
var state={};
ROWS.forEach(function(r){ state[r.id]={choice:null,custom:false,bk:null,sd:null,c1:null,c2:null,acc:null,note:''}; });
var saveTimer=null;
function save(){ try{ localStorage.setItem(KEY, JSON.stringify(state)); }catch(e){}
  var el=document.getElementById('saved'); if(el){ el.classList.add('on'); clearTimeout(saveTimer);
  saveTimer=setTimeout(function(){ el.classList.remove('on'); },1400); } }
function load(){ try{ var raw=localStorage.getItem(KEY); if(!raw) return null; var o=JSON.parse(raw), n=0;
  for(var k in o){ if(state.hasOwnProperty(k)&&o[k]){ state[k]=o[k]; n++; } } return n; }catch(e){ return null; } }
function clearSave(){ try{ localStorage.removeItem(KEY); }catch(e){} }
function toast(m){ var t=document.getElementById('toast'); t.textContent=m; t.style.display='block';
  clearTimeout(t._t); t._t=setTimeout(function(){t.style.display='none';},2400); }
function esc(s){ return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function ohtml(a,cur){ var h='<option value="">(不填)</option>';
  a.forEach(function(v){ h+='<option'+(v===cur?' selected':'')+'>'+esc(v)+'</option>'; }); return h; }
function sides(bk){ return bk?Object.keys(TREE[bk]||{}):[]; }
function c1list(bk,sd){ if(!bk||!sd) return []; return Object.keys((TREE[bk]||{})[sd]||{}); }
function c2list(bk,sd,c1){ if(!bk||!sd||!c1) return []; return ((TREE[bk]||{})[sd]||{})[c1]||[]; }
function eff(r){ var s=state[r.id];
  return {bk:s.bk||r.bk, sd:s.sd||r.sd, c1:s.c1||r.c1,
          c2:(s.c2!==null&&s.c2!==undefined)?s.c2:r.c2, acc:s.acc||r.acc}; }
function editor(r){ var s=state[r.id], e=eff(r), subs=c2list(e.bk,e.sd,e.c1);
  var accs=ACC.map(function(a){return a.name;});
  return '<div class="editor">'
  +'<div class="frow"><label>账本</label><select onchange="setBk('+r.id+',this.value)">'+ohtml(Object.keys(TREE),e.bk)+'</select></div>'
  +'<div class="frow"><label>方向</label><select onchange="setSd('+r.id+',this.value)">'+ohtml(sides(e.bk),e.sd)+'</select></div>'
  +'<div class="frow"><label>一级分类</label><select onchange="setC1('+r.id+',this.value)">'+ohtml(c1list(e.bk,e.sd),e.c1)+'</select></div>'
  +'<div class="frow"><label>二级分类</label><select onchange="setF('+r.id+',\'c2\',this.value)">'+ohtml(subs,e.c2)+'</select></div>'
  +'<div class="path">'+esc(e.bk)+' › '+esc(e.sd)+' › '+esc(e.c1)+(e.c2?' › '+esc(e.c2):'  （无二级）')+'</div>'
  +'<div class="frow"><label>账户</label><select onchange="setF('+r.id+',\'acc\',this.value)">'+ohtml(accs,e.acc)+'</select></div>'
  +'<div class="frow"><label>备注</label><input id="rm_'+r.id+'" type="text" value="'+esc(s.note||r.rem||'')+'" oninput="onIN('+r.id+',this.value)"></div>'
  +'</div>'; }
function noteBox(r){ var s=state[r.id];
  return '<div class="notearea"><div class="nt">补充说明（用于 AI 判断这笔的真实性质）<span class="lock">不影响已填分类</span></div>'
  +'<div class="quick">'+QUICK.map(function(q){ return '<button class="qtag" onclick="addTag('+r.id+',\''+q+'\')">+ '+q+'</button>'; }).join('')+'</div>'
  +'<textarea class="noteinput'+(s.note?' filled':'')+'" id="nt_'+r.id+'" placeholder="例：这是XX账单的退款 / 垫付的钱，已找对方报销 / AA分摊，5人平摊" oninput="onNT('+r.id+',this.value,this)">'+esc(s.note||'')+'</textarea></div>'; }
function render(){
  var need=ROWS.filter(function(r){return r.conf==='需确认';}).sort(function(a,b){return b.amt-a.amt;});
  var html='';
  need.forEach(function(r){ var s=state[r.id];
    var cls=r.amt>=300?'p0':(r.amt>=100?'p1':'p2');
    var pt=r.amt>=300?'P0 大额':(r.amt>=100?'P1 中额':'P2 小额');
    var o='';
    (r.options||[]).forEach(function(op,i){
      o+='<button class="opt'+(s.choice===i?' sel':'')+'" onclick="pick('+r.id+','+i+')">'+esc(op.label)+'<small>'+esc(op.desc)+'</small></button>'; });
    html+='<div class="card '+(s.custom?'custom':(s.choice!==null?'done':''))+'">'
    +'<div class="hd"><div>'
    +'<span class="tag '+cls+'">'+pt+'</span><span class="tag '+(r.t==='收入'?'in':'')+'">'+esc(r.t)+'</span><span class="tag bk">'+esc(r.bk)+'</span>'
    +(s.custom?'<span class="tag err">已修正</span>':'')
    +'<div class="meta">'+esc(r.dt)+' · '+esc((r.party||r.rem||'-'))+'</div>'
    +(r.goods&&r.goods!=='/'?'<div class="meta">'+esc(r.goods)+'</div>':'')
    +'<div class="meta">现归类 '+esc(r.sd)+' › '+esc(r.c1)+(r.c2?' › '+esc(r.c2):'')+' · 账户 '+esc(r.acc||'-')+'</div></div>'
    +'<div class="amt">'+r.amt.toFixed(2)+'</div></div>'
    +'<div class="opts">'+o+'<button class="opt custom" onclick="tog('+r.id+')">'+(s.custom?'收起':'✎ 手工修正')+'</button></div>'
    +'<div class="hint">'+hintOf(r)+'</div>'+noteBox(r)+(s.custom?editor(r):'')+'</div>'; });
  document.getElementById('list').innerHTML=html;
  var l=Object.keys(state).filter(function(k){return state[k].choice===null&&!state[k].custom;}).length;
  var fx=Object.keys(state).filter(function(k){return state[k].custom;}).length;
  document.getElementById('s1').textContent=l;
  document.getElementById('s5').textContent=fx;
  var auto=ROWS.filter(function(r){return r.conf!=='需确认';});
  var t='<table><tr><th>日期</th><th>类型</th><th>金额</th><th>账本</th><th>方向</th><th>一级</th><th>二级</th><th>账户</th><th>说明</th><th>依据</th><th></th></tr>';
  auto.forEach(function(r){ var s=state[r.id], e=eff(r);
    t+='<tr class="'+(s.custom?'fixed':'')+'"><td>'+esc(r.dt)+'</td><td>'+esc(r.t)+'</td><td style="text-align:right">'+r.amt.toFixed(2)
    +'</td><td>'+esc(e.bk)+'</td><td>'+esc(e.sd)+'</td><td>'+esc(e.c1)+'</td><td>'+esc(e.c2)+'</td><td>'+esc((e.acc||'').replace(/[「」]/g,''))
    +'</td><td>'+esc((s.note||'').slice(0,20))+'</td><td style="color:#a8a29e">'+esc(r.why)+'</td>'
    +'<td><button class="link" onclick="tog('+r.id+')">'+(s.custom?'收起':'修正')+'</button></td></tr>';
    if(s.custom) t+='<tr class="fixed"><td colspan="11">'+editor(r)+'</td></tr>'; });
  document.getElementById('tbl').innerHTML=t+'</table>'; }
function hintOf(r){ var s=state[r.id], e=eff(r);
  if(s.custom) return '<span class="done-tag">已改为：</span>'+esc(e.bk)+' › '+esc(e.sd)+' › '+esc(e.c1)+(e.c2?' › '+esc(e.c2):'');
  if(s.choice!==null){ var o=r.options[s.choice]; return '<span class="done-tag">已选：</span>'+esc(o.label); }
  return '点选项快速判定性质，或「✎ 手工修正」按账本逐层选'; }
function pick(id,i){ state[id].choice=i; state[id].custom=false; save(); render();
  var l=Object.keys(state).filter(function(k){return state[k].choice===null&&!state[k].custom;}).length;
  if(l===0) toast('全部处理完成，点「生成结果」'); }
function tog(id){ state[id].custom=!state[id].custom; save(); render();
  if(state[id].custom) setTimeout(function(){ var e=document.getElementById('rm_'+id); if(e) e.focus(); },30); }
function setBk(id,v){ state[id].bk=v; state[id].sd=null; state[id].c1=null; state[id].c2=null; state[id].custom=true; save(); render(); }
function setSd(id,v){ state[id].sd=v; state[id].c1=null; state[id].c2=null; state[id].custom=true; save(); render(); }
function setC1(id,v){ state[id].c1=v; state[id].c2=null; state[id].custom=true; save(); render(); }
function setF(id,k,v){ state[id][k]=v; state[id].custom=true; save(); render(); }
function onIN(id,v){ state[id].note=v; state[id].custom=true; save(); }
function onNT(id,v,el){ state[id].note=v; if(el) el.classList.toggle('filled',!!v.trim());
  var n=Object.keys(state).filter(function(k){return state[k].custom;}).length;
  var e=document.getElementById('s5'); if(e) e.textContent=n; save(); }
function addTag(id,tag){ var el=document.getElementById('nt_'+id);
  var cur=((el&&typeof el.value==='string')?el.value.trim():'')||state[id].note||'';
  if(cur.indexOf(tag)>=0){ toast('已包含「'+tag+'」'); return; }
  state[id].note=cur?(cur+'，'+tag):tag; state[id].custom=true; save(); render();
  var e2=document.getElementById('nt_'+id);
  if(e2&&typeof e2.value==='string'){ try{ e2.focus(); e2.setSelectionRange(e2.value.length,e2.value.length); }catch(err){} } }
function resetAll(){ if(!confirm('确定清空所有已填写的内容？此操作不可恢复。')) return;
  clearSave(); ROWS.forEach(function(r){ state[r.id]={choice:null,custom:false,bk:null,sd:null,c1:null,c2:null,acc:null,note:''}; });
  document.getElementById('dock').classList.remove('on'); render(); toast('已清空'); }
function buildText(){
  var need=ROWS.filter(function(r){return r.conf==='需确认';}).sort(function(a,b){return b.amt-a.amt;});
  var picks=[],miss=[];
  need.forEach(function(r){ var s=state[r.id];
    if(s.custom){}
    else if(s.choice!==null){ var o=r.options[s.choice];
      picks.push(r.dt+'  '+r.amt.toFixed(2)+'  '+(r.party||r.rem)+'  → '+o.label+(s.note?'  [说明: '+s.note+']':'')); }
    else miss.push(r.amt.toFixed(2)+' '+r.party); });
  var fx=[];
  ROWS.forEach(function(r){ var s=state[r.id]; if(!s.custom) return; var e=eff(r);
    fx.push((r.conf==='需确认'?'':'[自动归类修正] ')+r.dt+'  '+r.amt.toFixed(2)+'  '+(r.party||r.rem)+'  → '+e.bk+' › '+e.sd+' › '+e.c1+(e.c2?' › '+e.c2:'')+' / '+(e.acc||'-')+(s.note?'  [说明: '+s.note+']':'')); });
  var out=[];
  if(miss.length) out.push('⚠️ 未处理 '+miss.length+' 笔: '+miss.join('、'));
  else out.push('✅ 待确认 '+need.length+' 笔已全部处理');
  var af=fx.filter(function(x){return x.indexOf('[自动归类修正]')===0;});
  var nf=fx.filter(function(x){return x.indexOf('[自动归类修正]')!==0;});
  if(af.length) out.push('','【自动归类笔的修正 '+af.length+' 条】',af.join(NL));
  if(nf.length) out.push('','【手工填写的待确认笔】',nf.join(NL));
  if(picks.length) out.push('','【点选确认的待确认笔】',picks.join(NL));
  var notes=[];
  ROWS.forEach(function(r){ var s=state[r.id];
    if(s.note&&s.note!==r.rem) notes.push(r.dt+'  '+r.amt.toFixed(2)+'  '+(r.party||r.rem)+'  说明: '+s.note); });
  if(notes.length) out.push('','【补充说明 '+notes.length+' 条】',notes.join(NL));
  return out.join(NL); }
function showOut(){ document.getElementById('dockBody').textContent=buildText();
  document.getElementById('dock').classList.add('on');
  try{ window.scrollTo({top:document.body.scrollHeight,behavior:'smooth'}); }catch(e){} }
function copyOut(){ var txt=buildText();
  if(navigator.clipboard&&navigator.clipboard.writeText){
    navigator.clipboard.writeText(txt).then(function(){toast('已复制到剪贴板');},function(){fb(txt);}); } else fb(txt); }
function fb(txt){ var ta=document.createElement('textarea'); ta.value=txt;
  ta.style.position='fixed'; ta.style.left='-9999px'; document.body.appendChild(ta); ta.select();
  try{ document.execCommand('copy'); toast('已复制'); }catch(e){ toast('复制失败，请手动选择'); }
  document.body.removeChild(ta); }
function downloadTxt(){ var txt=buildText();
  var blob=new Blob([txt],{type:'text/plain;charset=utf-8'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='__FILENAME__.txt';
  document.body.appendChild(a); a.click(); document.body.removeChild(a); toast('已下载'); }
var restored=load(); render();
if(restored){ var f=Object.keys(state).filter(function(k){var s=state[k]; return s.choice!==null||s.custom;}).length;
  setTimeout(function(){ toast('已恢复上次填写的 '+f+' 笔内容'); },500); }
</script>
</body>
</html>'''

QUICK_TAGS = ['退款', '报销', '垫付', 'AA分摊', '借出', '还款', '礼金', '代买',
              '退货', '赔付', '公司', '家人', '朋友', '出差', '旅游', '医疗', '学习']


def build(rows, cfg, outpath, title='iCost 记账确认', sub='', key='icost_v1'):
    """rows: 引擎输出的行列表（需含 id/dt/t/amt/bk/sd/c1/c2/acc/rem/why/conf/party/goods）"""
    import collections
    TREE = cfg['cats']
    accs = [a['name'] for a in cfg['accts'] and
            [x for g, v in cfg['accts'].items() if not g.startswith('历史')
             for x in v['accounts'] if x.get('main')]]
    accs.sort(key=lambda n: -next((a.get('hist', 0)
                 for g, v in cfg['accts'].items() if not g.startswith('历史')
                 for a in v['accounts'] if a['name'] == n), 0))
    c1n = sum(len(v) for bk in TREE.values() for v in bk.values())
    c2n = sum(len(c2) for bk in TREE.values() for v in bk.values() for c2 in v.values())
    need = [r for r in rows if r.get('conf') == '需确认']
    exp = sum(r['amt'] for r in rows if r['t'] == '支出')

    # 为需确认项生成选项
    KIND = {'转账待定': ['借出', '还款', 'AA分摊', '礼金'],
            '朋友往来': ['群收款', '其他收入', '转账红包'],
            'AA收款': ['群收款', '其他收入'],
            '待确认': []}
    for r in rows:
        r.setdefault('options', [])
        if r.get('conf') == '需确认' and not r['options']:
            labels = KIND.get(r.get('why'), [])
            if not labels:
                labels = ['日常餐饮', '生活日用', '购物', '其他支出']
            r['options'] = [{'label': x, 'desc': '按实际情况选择'} for x in labels]

    page = (TPL
            .replace('__TITLE__', title)
            .replace('__SUB__', sub)
            .replace('__BK__', str(len(TREE)))
            .replace('__C1__', str(c1n)).replace('__C2__', str(c2n))
            .replace('__ACC__', str(len(accs)))
            .replace('__TOTAL__', str(len(rows)))
            .replace('__N1__', str(len(need)))
            .replace('__AUTO__', str(len(rows) - len(need)))
            .replace('__EXP__', '{:,.0f}'.format(exp))
            .replace('__KEY__', key)
            .replace('__FILENAME__', 'iCost确认结果')
            .replace('__ROWS__', json.dumps(rows, ensure_ascii=False))
            .replace('__TREE__', json.dumps(TREE, ensure_ascii=False))
            .replace('__ACC2__', json.dumps(accs, ensure_ascii=False))
            .replace('var ACC=__ACC__;', 'var ACC=__ACC2__;'))
    with open(outpath, 'w', encoding='utf-8') as f:
        f.write(page)
    return len(need)


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('--rows', required=True, help='记账明细.csv')
    ap.add_argument('--out', required=True, help='输出 html 路径')
    a = ap.parse_args()
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import engine
    cfg = engine.load_config()
    rows = []
    with open(a.rows, encoding='utf-8-sig') as f:
        rd = csv.DictReader(f)
        for i, r in enumerate(rd):
            rows.append({
                'id': i, 'dt': r['日期'].replace('-', '/')[:16], 't': r['类型'],
                'amt': float(r['金额']), 'bk': r['账本'], 'sd': r['方向'],
                'c1': r['一级分类'], 'c2': r['二级分类'], 'acc': r['账户'],
                'rem': r['备注'], 'why': r['归类依据'], 'conf': r['置信度'],
                'party': r.get('交易对方', ''), 'goods': '',
            })
    n = build(rows, cfg, a.out)
    print('已生成确认页: %s（待确认 %d 笔）' % (a.out, n))
