(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const text = v => String(v ?? '').trim();
  const norm = v => text(v).toUpperCase().replace(/\s+/g, '');
  const key = (a,f) => `${norm(a)}|${norm(f)}`;
  const aliases = {
    account:['账号','店铺','店铺名称','店铺名','账号名称','account'], fnsku:['FNSKU','fnsku'],
    price:['采购单价','采购价','采购价格','采购成本','单价','purchase price'], currency:['币种','货币','currency'],
    title:['商品名称','产品名称','Listing标题','品名','商品名','listing标题'], hs:['海关编码','HTS code','HS编码','HS code','海关代码'],
    id:['Shipment ID','shipment id','货件ID','货件编号'], ref:['Reference ID','reference id','参考编号'],
    date:['货件建立','货件创建日期','货件创建时间'], total:['货件总件数'], address:['仓库地址'],
    qty:['数量','发货数量'], received:['箱数','接收数量','亚马逊接收数量'], company:['公司英文名'], companyCn:['公司中文名'],
  };
  function get(row, name) {
    const candidates = aliases[name].map(x => norm(x));
    const k = Object.keys(row).find(x => candidates.includes(norm(x)));
    return k === undefined ? '' : row[k];
  }
  function num(v) { if (v === '' || v == null) return null; let s = String(v).trim().replace(/,/g,''); return s !== '' && Number.isFinite(Number(s)) ? Number(s) : null; }
  function date(v) {
    if (!v && v !== 0) return '';
    if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}/${v.getMonth()+1}/${v.getDate()}`;
    if (typeof v === 'number' && window.XLSX) { const d = XLSX.SSF.parse_date_code(v); return d ? `${d.y}/${d.m}/${d.d}` : ''; }
    return text(v);
  }
  const state = { shipments:[], selected:null, products:[], demo:false, blocked:new Set() };
  const exchangeRate=6.66824; // USD/CNY, November 2026 month-to-date through Nov 6
  let catalog=window.initialCatalog;
  try { const saved=JSON.parse(localStorage.getItem('claim_catalog_v2')||'null'); if(saved&&saved.listing&&saved.products&&saved.addresses) { const originals=new Map(window.initialCatalog.products.map(p=>[norm(p.sku),p])); saved.products.forEach(p=>{const original=originals.get(norm(p.sku));if(!p.image&&original?.image){p.image=original.image;p.imageWidth=original.imageWidth;p.imageHeight=original.imageHeight}}); catalog=saved; } } catch (_) {}
  function warehouse(address) {
    const code=(text(address).match(/^([A-Z0-9]{3,6})(?:\s|$|-)/i)||[])[1];
    const found=(catalog?.addresses||[]).filter(a=>norm(a.code)===norm(code));
    const unique=[...new Map(found.map(a=>[`${a.address}|${a.zip}`,a])).values()];
    return unique.length===1?unique[0]:null;
  }
  function freight(s){const ps=s.items.map(i=>match(i,s.account));return s.items.every((i,n)=>i.qty>0&&ps[n]?.weight>0)?s.items.reduce((sum,i,n)=>sum+i.qty*ps[n].weight,0):null;}
  function countryName(address){const m=text(address).match(/(?:,|-)\s*(United States|United Kingdom|Canada|Germany|France|Japan|US|CA|GB)\s*$/i);return m?({US:'United States',CA:'Canada',GB:'United Kingdom'}[m[1].toUpperCase()]||m[1]):'';}
  try { state.products = JSON.parse(localStorage.getItem('claim_products_v1') || '[]'); if (!Array.isArray(state.products)) state.products=[]; } catch (_) { state.products=[]; }
  try { state.blocked = new Set(JSON.parse(localStorage.getItem('claim_conflicts_v1') || '[]')); } catch (_) { state.blocked=new Set(); }
  function tell(message, bad=false) { try { localStorage.setItem('claim_conflicts_v1',JSON.stringify([...state.blocked])); } catch (_) {} const e=$('status'); if(e){ e.textContent=message; e.dataset.kind=bad?'error':'info'; } }
  function read(file) { return file.arrayBuffer().then(buf => { if (!window.XLSX) throw Error('表格解析组件加载失败，请检查网络并刷新页面。'); const wb=XLSX.read(buf,{type:'array',cellDates:true}); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:'',raw:true}); }); }
  function match(item, account) { const skus=[...new Set((catalog?.listing||[]).filter(x=>norm(x.fnsku)===norm(item.fnsku)).map(x=>norm(x.sku)).filter(Boolean))]; if(skus.length!==1)return null;const candidates=(catalog?.products||[]).filter(p=>norm(p.sku)===skus[0]);if(candidates.length!==1)return null;const p=candidates[0],cost=num(p.cost),weight=num(p.weight);if(norm(p.unit)!=='KG')return {...p,price:cost===null?null:Number((cost/exchangeRate).toFixed(2)),currency:'USD',weight:null};return {...p,price:cost===null?null:Number((cost/exchangeRate).toFixed(2)),currency:'USD',weight}; }
  function grouped(rows) {
    const groups = new Map();
    for (const [index,r] of rows.entries()) {
      const id=text(get(r,'id')); if(!id) continue;
      const originalAddress=text(get(r,'address')), wh=warehouse(originalAddress),zips=originalAddress.match(/\b\d{5}(?:-\d{4})?\b/g)||[],sourceZip=zips.length>1?zips[zips.length-1]:(zips[0]&&zips[0].includes('-')?zips[0]:''); const h={id,account:text(get(r,'account')),ref:text(get(r,'ref')),date:date(get(r,'date')),address:wh&&/^[A-Z0-9]{3,6}$/i.test(originalAddress)?`${wh.code} - ${wh.address}`:originalAddress,zip:sourceZip||wh?.zip||'',zipConflict:!!(sourceZip&&wh?.zip&&sourceZip!==wh.zip),company:text(get(r,'company')),companyCn:text(get(r,'companyCn')),total:num(get(r,'total'))};
      if (!groups.has(id)) groups.set(id,{...h,items:[],conflicts:[],sourceRows:[]});
      const s=groups.get(id);
      for(const field of ['account','ref','date','address','company','companyCn','total']) if(String(s[field] ?? '') !== String(h[field] ?? '')) s.conflicts.push(`${{account:'账号',ref:'Reference ID',date:'货件建立日期',address:'仓库地址',company:'公司英文名',companyCn:'公司中文名',total:'货件总件数'}[field]}不一致（第 ${index+2} 行）`);
      s.items.push({fnsku:text(get(r,'fnsku')),title:text(get(r,'title')),qty:num(get(r,'qty')),received:num(get(r,'received')),hs:text(get(r,'hs'))}); s.sourceRows.push(index+2);
    }
    return [...groups.values()];
  }
  function issues(s) {
    const list=[...new Set(s.conflicts)];
    for(const [v,n] of [[s.id,'Shipment ID'],[s.account,'账号'],[s.ref,'Reference ID'],[s.date,'货件建立日期'],[s.address,'仓库地址'],[s.company,'公司英文名']]) if(!v) list.push(`缺少${n}`);
    if (/^[A-Z0-9]{3,6}$/i.test(s.address||'')) list.push('仓库代码未匹配到唯一地址，请核对地址库');
    if(s.zipConflict) list.push('索赔表邮编与地址库不一致，请核对');
    if(s.total == null) list.push('缺少货件总件数');
    s.items.forEach((i,n)=>{
      const p=match(i,s.account);
      if(!i.fnsku) list.push(`第 ${n+1} 项缺少 FNSKU`);
      if(!i.title) list.push(`第 ${n+1} 项缺少商品名称`);
      if(i.qty == null || i.qty <= 0 || !Number.isInteger(i.qty)) list.push(`第 ${n+1} 项发货数量无效`);
      if(!p) list.push(`第 ${n+1} 项的 FNSKU → SKU → 产品资料缺失或存在冲突`);
      else if(p.price == null || p.price < 0) list.push(`第 ${n+1} 项缺少采购单价`);
      if(!i.hs && !p?.hs) list.push(`第 ${n+1} 项缺少海关编码`);
      if(p && !p.currency) list.push(`第 ${n+1} 项缺少币种`);
      if(p && !(p.weight>0)) list.push(`第 ${n+1} 项缺少有效重量（kg）`);
    });
    const currencies=new Set(s.items.map(i=>match(i,s.account)?.currency).filter(Boolean).map(norm)); if(currencies.size>1) list.push('同一货件存在不同币种');
    return [...new Set(list)];
  }
  function money(n) { return Number(n).toFixed(2); }
  function renderProducts() {
    const el=$('productsTable')?.querySelector('tbody'); if(!el) return;
    el.innerHTML=state.products.length ? state.products.map(p=>`<tr><td>${esc(p.account)}</td><td class="mono">${esc(p.fnsku)}</td><td>${esc(p.title || '—')}</td><td>${p.price==null?'—':esc(money(p.price))} ${esc(p.currency)}</td><td>${esc(p.hs||'—')}</td></tr>`).join('') : '<tr><td colspan="5" class="empty">暂无产品数据，导入领星表格后自动匹配</td></tr>';
    const count=$('productCount'); if(count) count.textContent=state.products.length;
  }
  function render() {
    const target=$('shipments')?.querySelector('tbody'); if(!target) return;
    target.innerHTML=state.shipments.length ? state.shipments.map((s,index)=>{const problems=issues(s); const units=s.items.reduce((a,i)=>a+(i.qty||0),0);return `<tr data-index="${index}" class="shipment-row ${state.selected===index?'selected':''}" tabindex="0"><td><span class="mono strong">${esc(s.id)}</span><br><span class="subtle">${esc(s.ref||'缺少 Reference ID')}</span></td><td>${s.items.length} 项</td><td>${units} 件</td><td><span class="pill ${problems.length?'warn':'success'}">${problems.length?`待处理 ${problems.length}`:'可生成'}</span></td><td>查看 →</td></tr>`}).join('') : '<tr><td colspan="5" class="empty">上传索赔资料后，货件将显示在这里</td></tr>';
    target.querySelectorAll('[data-index]').forEach(el=>{const select=()=>{state.selected=Number(el.dataset.index);render()};el.addEventListener('click',select);el.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();select()}})});
    const total=$('shipmentCount'); if(total) total.textContent=state.shipments.length;
    const ready=$('readyCount'); if(ready) ready.textContent=state.shipments.filter(s=>!issues(s).length).length;
    const attention=$('attentionCount'); if(attention) attention.textContent=state.shipments.filter(s=>issues(s).length).length;
    renderPreview(); renderProducts();
  }
  function renderPreview() {
    const el=$('preview'); if(!el) return; const s=state.shipments[state.selected];
    if(!s){el.innerHTML='<div class="preview-empty"><div class="preview-icon">▤</div><strong>选择一条货件</strong><p>在货件列表中选择货件，查看装箱单和待处理事项。</p></div>'; return;}
    const ps=issues(s);
    const rows=s.items.map(i=>{const p=match(i,s.account);const amount=i.qty!=null && p?.price!=null ? money(i.qty*p.price):'';const img=p?.image?.startsWith('data:image/jpeg;base64,')?`<img class="packing-photo" src="${esc(p.image)}" alt="商品图片">`:'';return `<tr class="packing-item"><td>${esc(i.title)}</td><td>${esc(i.hs||p?.hs||'')}</td><td>${i.qty??''}</td><td>${i.qty??''}</td><td>${esc(i.fnsku)}</td><td>${img}</td><td>${p?.price==null?'':money(p.price)}</td><td>${amount}</td></tr>`}).join('');
    const currency=match(s.items[0],s.account)?.currency||'';
    const country = countryName(s.address);
    el.innerHTML=`<div class="preview-top"><div><span class="eyebrow">PACKING LIST</span><h3>${esc(s.id)}</h3><p>${esc(s.ref)} · ${esc(s.date)}</p></div><span class="pill ${ps.length?'warn':'success'}">${ps.length?'待补全':'可生成'}</span></div>${ps.length?`<div class="notice"><strong>生成前需要处理</strong><ul>${ps.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>`:''}<div class="packing-scroll"><div class="packing-document"><div class="packing-title">PACKING LIST INFORMATION</div><table class="packing-top"><colgroup><col style="width:13.8%"><col style="width:33.3%"><col style="width:13%"><col style="width:10.8%"><col style="width:10.4%"><col style="width:18.7%"></colgroup><tbody><tr><td>TO 收件方</td><td>${esc(s.company||'')}</td><td>Date 时间:</td><td class="center">${esc(s.date||'')}</td><td>PSC 件数:</td><td class="center">${s.total??''}</td></tr><tr><td>Name 收件人:</td><td class="center">Amazon.com Services LLC</td><td>VAT（选填）:</td><td colspan="3"></td></tr><tr><td rowspan="2">Address 地址:</td><td class="center" rowspan="2">${esc(s.address||'')}</td><td>EORI（选填）:</td><td colspan="3"></td></tr><tr><td>Zip code邮编:</td><td colspan="3"></td></tr><tr class="packing-tel"><td>Tel 电话:</td><td></td><td colspan="2" class="center">Shipment ID:</td><td colspan="2" class="center">${esc(s.id)}</td></tr><tr><td>Country 国家:</td><td class="center">${esc(country)}</td><td colspan="2" class="center">Reference ID:</td><td colspan="2" class="center">${esc(s.ref||'')}</td></tr></tbody></table><table class="packing-items"><colgroup><col style="width:26.2%"><col style="width:9.1%"><col style="width:4.7%"><col style="width:7.2%"><col style="width:12.9%"><col style="width:10.8%"><col style="width:10.4%"><col style="width:18.7%"></colgroup><thead><tr><th>Product Name &amp; Description<br>产品名称</th><th>HTS code（海关编码）</th><th>CTN<br>箱数</th><th>Quantity<br>数量</th><th>FNSKU:</th><th>产品图片</th><th>Unit Price单价<br>(${esc(currency)})</th><th>Total 总价值（${esc(currency)}）</th></tr></thead><tbody>${rows}<tr class="packing-summary"><td>Freight 运费:</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr class="packing-summary"><td>SubTotal（合计）:</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr class="packing-origin"><td colspan="8">Made in China</td></tr></tbody></table><div class="packing-signatures"><span>出货人（签字）：</span><span>验收人（签字）：${esc(s.companyCn||'')}</span></div></div></div><p class="preview-note">CTN 与发货数量一致；签名、印章和未提供的图片留空。采购单价以导入的领星资料为准。</p>`;
    const products=s.items.map(i=>match(i,s.account));
    const canSum=s.items.length>0 && s.items.every((i,n)=>Number.isInteger(i.qty)&&i.qty>0&&products[n]&&products[n].price!=null&&products[n].price>=0&&products[n].currency) && new Set(products.map(p=>norm(p.currency))).size===1;
    const fr=freight(s), summary=el.querySelectorAll('.packing-summary');
    if(summary[0] && fr!==null) summary[0].firstElementChild.textContent='Freight 运费：$'+money(fr);
    if(summary[1] && canSum && fr!==null) summary[1].firstElementChild.textContent='SubTotal（合计）：$'+money(s.items.reduce((sum,i,n)=>sum+i.qty*products[n].price,0)+fr);
    const zipLabel=[...el.querySelectorAll('td')].find(x=>x.textContent.trim()==='Zip code邮编:'); if(zipLabel?.nextElementSibling)zipLabel.nextElementSibling.textContent=s.zip||'';
    const note=el.querySelector('.preview-note'); if(note)note.textContent='运费按当前明细的商品毛重 × 发货数量 × 1 USD/kg 估算；合计=商品金额+运费。汇率：1 USD = 6.66824 CNY（2026年11月截至6日均值）。';
  }
  async function makeDoc(s) {
    if (!window.docx) throw Error('Word 组件加载失败，请检查网络并刷新页面。');
    const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,AlignmentType,VerticalAlign,PageOrientation,HeightRule,TableLayoutType} = window.docx;
    const usable = 14738;
    const topCols = [2034, 4911, 1916, 1591, 1533, 2753];
    const itemCols = [3854,1341,695,1061,1901,1592,1533,2761];
    const borders = Object.fromEntries(['top','bottom','left','right','insideHorizontal','insideVertical'].map(k=>[k,{style:'single',size:6,color:'000000'}]));
    const para = (value, opts={}) => new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: {before:0,after:0,line:220},
      children: String(value ?? '').split('\n').map((line,index)=>new TextRun({text:line,break:index?1:0,bold:!!opts.bold,size:opts.size||18,font:'Arial'}))
    });
    function c(value, width, opts={}) {
      return new TableCell({
        width:{size:width,type:WidthType.DXA}, columnSpan:opts.span||1, rowSpan:opts.rows||1,
        verticalAlign:VerticalAlign.CENTER, margins:{top:25,bottom:25,left:40,right:40},
        children:[para(value,opts)]
      });
    }
    const row = (cells,height,header=false) => new TableRow({children:cells, cantSplit:true,
      ...(height ? {height:{value:height,rule:HeightRule.ATLEAST}} : {}), ...(header ? {tableHeader:true} : {})});
    const sumCols=(cols,start,count)=>cols.slice(start,start+count).reduce((a,b)=>a+b,0);
    const top = new Table({width:{size:usable,type:WidthType.DXA},columnWidths:topCols,layout:TableLayoutType.FIXED,borders,rows:[
      row([c('PACKING LIST INFORMATION',usable,{span:6,center:true,bold:true,size:20})],470),
      row([c('TO 收件方',topCols[0]),c(s.company,topCols[1]),c('Date 时间：',topCols[2]),c(s.date,topCols[3],{center:true}),c('PSC 件数：',topCols[4]),c(s.total,topCols[5],{center:true})],240),
      row([c('Name 收件人：',topCols[0]),c('Amazon.com Services LLC',topCols[1],{center:true}),c('VAT（选填）：',topCols[2]),c('',sumCols(topCols,3,3),{span:3})],240),
      row([c('Address 地址：',topCols[0],{rows:2}),c(s.address,topCols[1],{rows:2,center:true}),c('EORI（选填）：',topCols[2]),c('',sumCols(topCols,3,3),{span:3})],240),
      row([c('Zip code邮编：',topCols[2]),c(s.zip||'',sumCols(topCols,3,3),{span:3})],430),
      row([c('Tel 电话：',topCols[0]),c('',topCols[1]),c('Shipment ID：',sumCols(topCols,2,2),{span:2,center:true}),c(s.id,sumCols(topCols,4,2),{span:2,center:true})],600),
      row([c('Country 国家：',topCols[0]),c(countryName(s.address),topCols[1],{center:true}),c('Reference ID：',sumCols(topCols,2,2),{span:2,center:true}),c(s.ref,sumCols(topCols,4,2),{span:2,center:true})],240)
    ]});
    let sum=0;
    const currency=match(s.items[0],s.account).currency;
    const headings=['Product Name & Description\n产品名称','HTS code（海关编码）','CTN\n箱数','Quantity\n数量','FNSKU:','产品图片',`Unit Price单价\n(${currency})`,`Total 总价值（${currency}）`];
    const detail=s.items.map(i=>{const p=match(i,s.account); const amount=i.qty*p.price;sum+=amount;
      const values=[i.title,i.hs||p.hs,i.qty,i.qty,i.fnsku,'',money(p.price),money(amount)];
      const cells=values.map((v,j)=>c(v,itemCols[j],{center:j===0||j===2||j===3||j>=6}));
      if(p.image?.startsWith('data:image/jpeg;base64,')){const bytes=Uint8Array.from(atob(p.image.split(',')[1]),ch=>ch.charCodeAt(0)),w=p.imageWidth||200,h=p.imageHeight||200,scale=75/Math.max(w,h);cells[5]=new TableCell({width:{size:itemCols[5],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({type:'jpg',data:bytes,transformation:{width:Math.round(w*scale),height:Math.round(h*scale)}})]})]});}
      return row(cells,1100);
    });
    const blankRow=(label,last='')=>row([c(label,itemCols[0]),...itemCols.slice(1,7).map(w=>c('',w)),c(last,itemCols[7],{center:true})],220);
    const items = new Table({width:{size:usable,type:WidthType.DXA},columnWidths:itemCols,layout:TableLayoutType.FIXED,borders,rows:[
      row(headings.map((v,j)=>c(v,itemCols[j],{center:j===0})),460,true),...detail,
      blankRow('Freight 运费：$'+money(freight(s))),blankRow('SubTotal（合计）：$'+money(sum+freight(s))),
      row([c('Made in China',usable,{span:8,center:true})],230)
    ]});
    const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838,orientation:PageOrientation.LANDSCAPE},margin:{top:1500,bottom:900,left:1050,right:1050}}},children:[
      top,items,
      new Paragraph({spacing:{before:300,after:0},alignment:AlignmentType.RIGHT,children:[new TextRun({text:'出货人（签字）：                         验收人（签字）：'+(s.companyCn||''),size:18,font:'Arial'})]})
    ]}]});
    return Packer.toBlob(doc);
  }
  function download(blob,name){const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.style.display='none';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),60000)}
  function downloadMessage(id, message, bad=false) { const e=$(id); if(e){e.textContent=message;e.dataset.kind=bad?'error':'info';} }
  function filename(s){return `PackingList_${s.id.replace(/[^a-z0-9_-]/gi,'_')}_${s.ref.replace(/[^a-z0-9_-]/gi,'_')}.docx`}
  async function importClaim(file){try{if(!file)return;if(!/\.(xlsx|xls|csv)$/i.test(file.name))throw Error('请选择 XLSX、XLS 或 CSV 格式的索赔表。');const rows=await read(file);const groups=grouped(rows);if(!groups.length)throw Error('没有识别到 Shipment ID。请确认第一行为表头。');state.shipments=groups;state.selected=0;state.demo=false;render();tell(`已导入 ${groups.length} 个货件，共 ${rows.length} 行。请检查待处理事项。`)}catch(err){tell(err.message,true)}}
  $('claimFile')?.addEventListener('change',async e=>{try{await importClaim(e.target.files[0])}finally{e.target.value=''}});
  const dropzone=$('claimDropzone');
  if(dropzone){
    let depth=0;
    dropzone.addEventListener('dragenter',e=>{e.preventDefault();depth++;dropzone.classList.add('drag-active')});
    dropzone.addEventListener('dragover',e=>{e.preventDefault();if(e.dataTransfer)e.dataTransfer.dropEffect='copy'});
    dropzone.addEventListener('dragleave',e=>{e.preventDefault();depth=Math.max(0,depth-1);if(!depth)dropzone.classList.remove('drag-active')});
    dropzone.addEventListener('drop',async e=>{e.preventDefault();depth=0;dropzone.classList.remove('drag-active');const files=e.dataTransfer?.files;if(!files?.length)return;if(files.length!==1){tell('请一次拖入一份索赔表。',true);return}await importClaim(files[0])});
  }
  $('productFile')?.addEventListener('change',async e=>{try{const rows=await read(e.target.files[0]);let added=0,updated=0,skipped=0;const seen=new Map();let conflicts=0;for(const r of rows){const account=text(get(r,'account')),fnsku=text(get(r,'fnsku')),price=num(get(r,'price')),currency=text(get(r,'currency')),title=text(get(r,'title')),hs=text(get(r,'hs'));if(!account||!fnsku||price==null||price<0){skipped++;continue}const k=key(account,fnsku),p={account,fnsku,price,currency,title,hs};if(seen.has(k)&&JSON.stringify(seen.get(k))!==JSON.stringify(p)){if(seen.get(k)!==null) conflicts++;seen.set(k,null)}else if(!seen.has(k))seen.set(k,p)} for(const [k,p] of seen.entries()){if(!p){state.blocked.add(k);continue}state.blocked.delete(k);const idx=state.products.findIndex(x=>key(x.account,x.fnsku)===key(p.account,p.fnsku));if(idx<0){state.products.push(p);added++}else{state.products[idx]=p;updated++}}localStorage.setItem('claim_products_v1',JSON.stringify(state.products));render();document.querySelector('[data-view="workspace"]')?.click();tell(`产品导入完成：新增 ${added}，更新 ${updated}，跳过 ${skipped}，冲突 ${conflicts}。缺账号、FNSKU 或采购价的记录不会导入；冲突记录需修正后重新导入。`,conflicts>0)}catch(err){document.querySelector('[data-view="workspace"]')?.click();tell(err.message,true)}finally{e.target.value=''}});
  $('templateBtn')?.addEventListener('click',()=>{if(!window.XLSX){tell('模板组件加载失败，请刷新后重试。',true);return}const fields=['账号','货件建立','货件总件数','仓库地址','Shipment ID','Reference ID','商品名称','数量','FNSKU','公司英文名','公司中文名','海关编码'];const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([fields]);ws['!cols']=fields.map((_,i)=>({wch:[20,18,16,40,23,20,40,12,22,32,30,20][i]}));XLSX.utils.book_append_sheet(wb,ws,'索赔录入表');const info=XLSX.utils.aoa_to_sheet([['填写说明'],['第一行表头请保留；每行填写一项商品。'],['同一货件多项商品请逐行填写，货件、账号、地址、日期等共同信息保持一致。'],['账号、货件建立、货件总件数、仓库地址、Shipment ID、Reference ID、商品名称、数量、FNSKU、公司英文名为核对字段。'],['公司中文名、海关编码按实际情况填写；产品及费用信息仍需系统匹配并人工核对。']]);info['!cols']=[{wch:90}];XLSX.utils.book_append_sheet(wb,info,'填写说明');XLSX.writeFile(wb,'索赔装箱单录入模板.xlsx')});
  $('exportBtn')?.addEventListener('click',async()=>{try{const s=state.shipments[state.selected];if(!s)throw Error('请先导入索赔资料并选择一个货件。');const problems=issues(s);if(problems.length)throw Error('暂不能生成：'+problems.join('；')+'。请补全资料后重新导入。');if(!window.docx)throw Error('Word 组件未加载，请确认完整解压体验版后重新打开页面。');downloadMessage('downloadStatus','正在生成 Word…');download(await makeDoc(s),filename(s));downloadMessage('downloadStatus','Word 已生成。请查看浏览器下载列表并核对内容。');}catch(err){downloadMessage('downloadStatus',err.message||'生成失败，请重试。',true)}});
  $('exportAllBtn')?.addEventListener('click',async()=>{try{if(!state.shipments.length)throw Error('请先导入索赔资料。');if(!window.JSZip||!window.docx)throw Error('下载组件未加载，请确认完整解压体验版后重新打开页面。');const good=state.shipments.filter(s=>!issues(s).length);if(!good.length)throw Error('没有可生成的货件。请选择货件查看缺失资料，补全后重新导入。');downloadMessage('batchStatus','正在生成并打包…');const zip=new JSZip();for(const s of good){const word=await makeDoc(s);zip.file(filename(s),await word.arrayBuffer())}download(await zip.generateAsync({type:'blob'}),'Packing_Lists.zip');downloadMessage('batchStatus',`已打包 ${good.length} 份，跳过 ${state.shipments.length-good.length} 份待处理货件。请查看浏览器下载列表。`)}catch(err){downloadMessage('batchStatus',err.message||'打包失败，请重试。',true)}});
  async function embeddedProductImages(buffer, rows){
    const archive=await JSZip.loadAsync(buffer),relFile=archive.file('xl/drawings/_rels/drawing1.xml.rels'),drawingFile=archive.file('xl/drawings/drawing1.xml');if(!relFile||!drawingFile)return new Map();
    const parser=new DOMParser(),rels=parser.parseFromString(await relFile.async('string'),'application/xml'),drawing=parser.parseFromString(await drawingFile.async('string'),'application/xml');
    const paths=new Map([...rels.getElementsByTagNameNS('*','Relationship')].map(r=>[r.getAttribute('Id'),r.getAttribute('Target')])),result=new Map();
    for(const anchor of [...drawing.documentElement.children]){const from=anchor.getElementsByTagNameNS('*','from')[0],blip=anchor.getElementsByTagNameNS('*','blip')[0];if(!from||!blip)continue;const col=Number(from.getElementsByTagNameNS('*','col')[0]?.textContent),row=Number(from.getElementsByTagNameNS('*','row')[0]?.textContent);if(col!==2||row<1||!rows[row]?.[0])continue;
      const target=paths.get(blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','embed'));if(!target)continue;const path=new URL(target,'https://local.invalid/xl/drawings/drawing1.xml').pathname.slice(1),entry=archive.file(path);if(!entry)continue;
      const blob=await entry.async('blob'),bitmap=await createImageBitmap(blob),scale=Math.min(1,200/Math.max(bitmap.width,bitmap.height)),canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(bitmap.width*scale));canvas.height=Math.max(1,Math.round(bitmap.height*scale));const context=canvas.getContext('2d');context.fillStyle='#fff';context.fillRect(0,0,canvas.width,canvas.height);context.drawImage(bitmap,0,0,canvas.width,canvas.height);bitmap.close();result.set(norm(rows[row][0]),{image:canvas.toDataURL('image/jpeg',0.82),imageWidth:canvas.width,imageHeight:canvas.height});
    }return result;
  }
  for(const [id,kind] of [['listingUpdate','listing'],['catalogUpdate','products'],['addressUpdate','addresses']]) $(id)?.addEventListener('change',async e=>{try{
    const buffer=await e.target.files[0].arrayBuffer(),wb=XLSX.read(buffer,{type:'array'});const sheet=kind==='addresses'?wb.Sheets['地址库']:wb.Sheets[wb.SheetNames[0]];if(!sheet)throw Error('未找到地址库工作表');
    const a=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true}),h=a[0]||[];let next;
    if(kind==='listing'){const fi=h.findIndex(x=>norm(x)==='FNSKU'),si=h.findIndex(x=>norm(x)==='SKU');if(fi<0||si<0)throw Error('Listing 表需要 FNSKU 与 SKU 列');next=a.slice(1).map(r=>({fnsku:r[fi],sku:r[si]})).filter(r=>r.fnsku&&r.sku)}
    if(kind==='products'){if(norm(h[0])!=='*SKU'||h.length<18||!String(h[3]).includes('CNY')||!String(h[17]).toUpperCase().includes('HSCODE'))throw Error('产品表格式不符，请使用含采购成本、重量和 HSCODE 的更新版');const images=await embeddedProductImages(buffer,a),old=new Map(catalog.products.map(p=>[norm(p.sku),p]));next=a.slice(1).map(r=>{const previous=old.get(norm(r[0]));return {sku:r[0],cost:r[3],weight:r[4],unit:r[5],hs:r[17],...(images.get(norm(r[0]))||(previous?.image?{image:previous.image,imageWidth:previous.imageWidth,imageHeight:previous.imageHeight}:{}))}}).filter(r=>r.sku)}
    if(kind==='addresses'){if(!String(h[1]).includes('地址编码')||!String(h[9]).includes('邮编'))throw Error('地址库列格式不符');next=a.slice(1).map(r=>({code:r[1],address:[r[3],r[4],r[5],r[6],r[7],r[8]].filter(Boolean).join(', '),zip:String(r[9]||'')})).filter(r=>r.code)}
    if(!next.length)throw Error('未识别到有效数据，原资料未改变');const updated={...catalog,[kind]:next};localStorage.setItem('claim_catalog_v2',JSON.stringify(updated));catalog=updated;render();$('maintenanceStatus').textContent=`更新成功：${next.length} 条。请重新导入索赔表以刷新地址。`;
  }catch(err){$('maintenanceStatus').textContent='更新失败：'+err.message}finally{e.target.value=''}});
  document.querySelectorAll('.nav-link').forEach(button=>button.addEventListener('click',()=>{document.querySelectorAll('.nav-link').forEach(x=>x.classList.toggle('active',x===button));$('workspaceSection').classList.toggle('hidden',button.dataset.view!=='workspace');$('productsSection').classList.toggle('hidden',button.dataset.view!=='products');}));
  render();
})();
