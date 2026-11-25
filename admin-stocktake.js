/* كيتشن فاست - الجرد (مرحلة 4): جلسة جرد، مقارنة بالرصيد المتوقع، اعتماد، طباعة، تصدير Excel. ملف مستقل. */
(function(){try{
if(window.__kfStk)return;window.__kfStk=1;
const UN={kg:'كيلو',l:'لتر',pcs:'قطعة'};
const f2=x=>(Math.round((+x||0)*100)/100).toLocaleString('en-US');
const f3=x=>(Math.round((+x||0)*1000)/1000).toLocaleString('en-US');
const TD=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Africa/Cairo'});
const dt=s=>new Date(s).toLocaleDateString('ar-EG',{timeZone:'Africa/Cairo'});
let ST=null,LN=[];
const cn=v=>v===''||v==null?null:+v;
const dif=l=>l.counted==null?null:+l.counted-+l.expected;
const un=l=>UN[l.materials.unit]||'';

async function list(){
 const r=ok(await sb.from('stocktakes').select('id,name,status,created_at').order('created_at',{ascending:false}).limit(50));
 V.innerHTML=`<div class="c"><b class="g">الجرد</b><p class="m" style="margin:4px 0 8px">ابدأ جلسة جرد: النظام بياخد صورة من الرصيد المتوقع لكل خامة، والشيف يعدّ الموجود فعلاً، وإنت تسجل العدّ. الفرق بيظهر بالكمية والقيمة، ولما تعتمد الجرد الرصيد بيتصحح.</p><label>اسم الجلسة (اختياري)<input id="stn" placeholder="جرد ${TD()}"></label><button class="b" data-x="stn">➕ بدء جلسة جرد جديدة</button></div>`
 +(r.map(s=>`<div class="c"><div class="r" data-sid="${s.id}"><span><b>${esc(s.name)}</b> <span class="m">${dt(s.created_at)}</span></span><span>${s.status=='approved'?'✅ معتمد':'📝 مسودة'}</span><button class="b" data-x="sto">فتح</button></div></div>`).join('')||'<p class="m">لسه مفيش جلسات جرد.</p>')}

async function start(){const id=ok(await sb.rpc('start_stocktake',{p_name:($('stn').value||'').trim()}));await open(id)}

async function open(id){
 const[a,b]=await Promise.all([sb.from('stocktakes').select('*').eq('id',id).single(),sb.from('stocktake_lines').select('id,material_id,expected,avg_cost,counted,materials(name,unit,kind)').eq('stocktake_id',id)]);
 ST=ok(a);LN=ok(b).sort((x,y)=>((x.materials.kind=='pack')-(y.materials.kind=='pack'))||x.materials.name.localeCompare(y.materials.name,'ar'));draw()}

function draw(){
 const dr=ST.status=='draft';
 const rows=LN.map(l=>{const m=l.materials;return`<div class="r" data-lid="${l.id}"><span style="min-width:140px"><b>${esc(m.name)}</b> <span class="m">${un(l)}${m.kind=='pack'?' · تغليف':''}</span></span><span>المتوقع: ${f3(l.expected)}</span>${dr?`<label style="display:flex;gap:6px;align-items:center">الفعلي<input data-f="cnt" type="number" min="0" step="any" value="${l.counted==null?'':+l.counted}" style="width:110px"></label>`:`<span>الفعلي: ${l.counted==null?'—':f3(l.counted)}</span>`}<span class="df"></span><span class="vl m"></span></div>`}).join('');
 V.innerHTML=`<div class="c"><div class="r"><b class="g">${esc(ST.name)}</b><span class="m">${dt(ST.created_at)} · ${dr?'📝 مسودة':'✅ معتمد '+dt(ST.approved_at)}</span></div><div class="r" style="margin-top:8px"><button class="b" data-x="stb">⬅️ الجلسات</button>${dr?'<button class="b" data-x="sts">💾 حفظ العدّ</button><button class="b" data-x="sta">✅ اعتماد الجرد</button>':''}<button class="b" data-x="stp1">🖨️ ورقة عدّ للشيف</button><button class="b" data-x="stp2">🖨️ طباعة التقرير</button><button class="b" data-x="stx">📊 تصدير Excel</button>${dr?'<button class="b d" data-x="std">🗑️ حذف</button>':''}</div></div><div class="c" id="stt"></div>${rows}`;upd()}

function cur(){return LN.map(l=>{const i=document.querySelector(`[data-lid="${l.id}"] [data-f=cnt]`);return i?{...l,counted:cn(i.value)}:l})}
function upd(){const L=cur();let sh=0,su=0,n=0;
 L.forEach(l=>{const row=document.querySelector(`[data-lid="${l.id}"]`);if(!row)return;const d=dif(l),df=row.querySelector('.df'),vl=row.querySelector('.vl');
  if(d==null){df.textContent='';vl.textContent='';return}
  n++;const v=d*(+l.avg_cost);if(v<0)sh+=v;else su+=v;
  df.innerHTML=`<b style="color:${d<0?'#ff6a6a':d>0?'var(--gold)':'inherit'}">${d>0?'+':''}${f3(d)} ${un(l)}</b>`;vl.textContent=(v>0?'+':'')+f2(v)+' ج'});
 $('stt').innerHTML=`<b>تم عدّ ${n} من ${LN.length} خامة</b> · <b style="color:#ff6a6a">عجز: ${f2(sh)} ج</b> · <b style="color:var(--gold)">زيادة: +${f2(su)} ج</b> · <b>الصافي: ${f2(sh+su)} ج</b>`}

async function save(){const L=cur();ok(await sb.from('stocktake_lines').upsert(L.map(l=>({id:l.id,stocktake_id:ST.id,material_id:l.material_id,expected:l.expected,avg_cost:l.avg_cost,counted:l.counted}))));LN=LN.map((l,i)=>({...l,counted:L[i].counted}))}

function pr(title,head,rows,foot){
 const h=`<!doctype html><html dir="rtl" lang="ar"><head><meta charset="utf-8"><title>${esc(title)}</title><style>@page{size:A4;margin:12mm}body{font-family:Tahoma,Arial,sans-serif;direction:rtl;color:#000}h2{margin:0 0 4px}.s{color:#555;font-size:12px;margin-bottom:10px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #444;padding:6px 8px;text-align:right;font-size:13px}th{background:#eee}td.b{height:28px}tfoot td{font-weight:bold;background:#f6f6f6}</style></head><body><h2>كيتشن فاست — ${esc(title)}</h2><div class="s">${esc(ST.name)} · ${dt(ST.created_at)}</div><table><thead><tr>${head.map(x=>`<th>${x}</th>`).join('')}</tr></thead><tbody>${rows}</tbody>${foot||''}</table></body></html>`;
 const f=document.createElement('iframe');f.style.cssText='position:fixed;right:0;bottom:0;width:0;height:0;border:0';document.body.appendChild(f);
 const d=f.contentWindow.document;d.open();d.write(h);d.close();
 setTimeout(()=>{f.contentWindow.focus();f.contentWindow.print();setTimeout(()=>f.remove(),3000)},400)}
function pChef(){pr('ورقة عدّ الجرد',['#','الخامة','الوحدة','الكمية الفعلية','ملاحظات'],LN.map((l,i)=>`<tr><td>${i+1}</td><td>${esc(l.materials.name)}</td><td>${un(l)}</td><td class="b"></td><td class="b"></td></tr>`).join(''))}
function pRep(){const L=cur();let sh=0,su=0;
 const rows=L.map(l=>{const d=dif(l),v=d==null?null:d*(+l.avg_cost);if(v!=null){if(v<0)sh+=v;else su+=v}
  return`<tr><td>${esc(l.materials.name)}</td><td>${un(l)}</td><td>${f3(l.expected)}</td><td>${l.counted==null?'—':f3(l.counted)}</td><td>${d==null?'—':(d>0?'+':'')+f3(d)}</td><td>${v==null?'—':(v>0?'+':'')+f2(v)}</td></tr>`}).join('');
 pr('تقرير الجرد',['الخامة','الوحدة','المتوقع','الفعلي','الفرق','قيمة الفرق (ج)'],rows,`<tfoot><tr><td colspan="5">إجمالي العجز</td><td>${f2(sh)}</td></tr><tr><td colspan="5">إجمالي الزيادة</td><td>+${f2(su)}</td></tr><tr><td colspan="5">الصافي</td><td>${f2(sh+su)}</td></tr></tfoot>`)}

const loadX=()=>window.XLSX?Promise.resolve():new Promise((res,rej)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';s.onload=res;s.onerror=rej;document.head.appendChild(s)});
async function xl(){const L=cur(),fn='stocktake-'+(ST.created_at||'').slice(0,10);
 const aoa=[['الخامة','الوحدة','المتوقع','الفعلي','الفرق','متوسط السعر','قيمة الفرق']];
 L.forEach((l,i)=>{const r=i+2,d=dif(l);aoa.push([l.materials.name,un(l),+l.expected,l.counted==null?'':+l.counted,d==null?'':{t:'n',v:d,f:`D${r}-C${r}`},+l.avg_cost,d==null?'':{t:'n',v:d*(+l.avg_cost),f:`E${r}*F${r}`}])});
 const n=L.length+1,vs=L.map(l=>{const d=dif(l);return d==null?0:d*(+l.avg_cost)}),sh=vs.filter(v=>v<0).reduce((a,b)=>a+b,0),su=vs.filter(v=>v>0).reduce((a,b)=>a+b,0);
 aoa.push([],['إجمالي العجز','','','','','',{t:'n',v:sh,f:`SUMIF(G2:G${n},"<0")`}],['إجمالي الزيادة','','','','','',{t:'n',v:su,f:`SUMIF(G2:G${n},">0")`}],['الصافي','','','','','',{t:'n',v:sh+su,f:`SUM(G2:G${n})`}]);
 try{await loadX();const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=[{wch:26},{wch:10},{wch:12},{wch:12},{wch:12},{wch:14},{wch:16}];const wb=XLSX.utils.book_new();wb.Workbook={Views:[{RTL:true}]};XLSX.utils.book_append_sheet(wb,ws,'الجرد');XLSX.writeFile(wb,fn+'.xlsx');toast('تم تصدير ملف Excel')}
 catch(e){const csv='\ufeff'+aoa.map(r=>r.map(c=>{const v=c&&c.v!==undefined?c.v:c;return'"'+String(v).replace(/"/g,'""')+'"'}).join(',')).join('\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download=fn+'.csv';a.click();toast('تعذر تحميل مكتبة Excel، اتصدّر ملف CSV بدلاً منه',1)}}

V.addEventListener('click',async e=>{const b=e.target.closest('[data-x^="st"]');if(!b)return;const x=b.dataset.x;
 try{
  if(x=='stn')await start();
  else if(x=='sto')await open(b.closest('[data-sid]').dataset.sid);
  else if(x=='stb')await list();
  else if(x=='sts'){await save();toast('تم حفظ العدّ')}
  else if(x=='sta'){const L=cur(),n=L.filter(l=>l.counted==null).length;if(!L.some(l=>l.counted!=null))return toast('سجّل العدّ الأول',1);
   if(!confirm('اعتماد الجرد؟ الرصيد هيتصحح للكميات اللي سجلتها.'+(n?'\n'+n+' خامة لسه ما اتعدّتش ومش هتتغير.':'')+'\nبعد الاعتماد مفيش تعديل.'))return;
   await save();ok(await sb.rpc('approve_stocktake',{p_id:ST.id}));toast('تم اعتماد الجرد');await open(ST.id)}
  else if(x=='std'){if(!confirm('حذف جلسة الجرد دي؟'))return;ok(await sb.from('stocktakes').delete().eq('id',ST.id));toast('تم الحذف');await list()}
  else if(x=='stp1')pChef();
  else if(x=='stp2')pRep();
  else if(x=='stx')await xl();
 }catch(err){fail(err)}});
V.addEventListener('input',e=>{if(e.target.dataset.f=='cnt')upd()});

TABS.stocktake=['الجرد',list];
$('nav').innerHTML=Object.entries(TABS).map(([k,v])=>`<button data-k="${k}">${v[0]}</button>`).join('');
}catch(e){console.error('admin-stocktake:',e)}})();
