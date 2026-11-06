const $=s=>document.querySelector(s);
const input=$('#files'),drop=$('#drop'),notice=$('#notice');
let rawRows=[],fileMeta=[],duplicates=0;
const rmGroups={
  'RM Fon':['801973','801997'],'RM Mint':['801974','801975','801976','801977'],
  'RM May':['803075','803099','803085'],'RM TOON':['803071','803117'],
  'RM CHI':['803079','803083','803092','803127','803119','803076','803088'],
  'RM TA':['803122'],'RM Pim':['803105','803108','803110','803066','803080'],
  'RM M':['603935','603936','603937','803077','803090','803097','803102','803120','803113'],
  'RM BOM':['603927','603930','803068','803078','803098','803123','803104'],
  'RM PondNBI':['803067','803084','803087','803100','803124','803126'],
  'RM MAX':['803081','803086','803072','803125','603941'],
  'RM Nopadon':['803073','803103'],'RM Ark':['803101','803131','803107']
};
const rmByCode=Object.fromEntries(Object.entries(rmGroups).flatMap(([rm,codes])=>codes.map(code=>[code,rm])));
const rmCodes=new Set(rmGroups['RM M']);
const branchNames={
  '801973':'01Don Mueang12','801997':'01Don Mueang15','801974':'01Lak Si06','801975':'01Lak Si07','801976':'01Lak Si08','801977':'01Lak Si09',
  '803075':'03Mueang Nonthaburi21','803099':'03Mueang Nonthaburi29','803085':'03Mueang Nonthaburi25','803071':'03Pak Kret12','803117':'03Pak Kret18',
  '803079':'03Mueang Nonthaburi23','803083':'03Mueang Nonthaburi24','803092':'03Mueang Nonthaburi26','803127':'03Mueang Nonthaburi34','803119':'03Mueang Nonthaburi32','803076':'03Mueang Nonthaburi22','803088':'03Pak Kret14','803122':'03Pak Kret19',
  '803105':'03Pak Kret15','803108':'03Pak Kret16','803110':'03Pak Kret17','803066':'03Pak Kret11','803080':'03Pak Kret13',
  '603935':'03Bang Krui040','603936':'03Bang Krui039','603937':'03Bang Krui041','803077':'03Bang Krui11','803090':'03Bang Krui12','803097':'03Bang Krui13','803102':'03Bang Krui14','803120':'03Bang Krui19','803113':'03Bang Krui16',
  '603927':'03Bang Yai056','603930':'03Bang Yai057','803068':'03Bang Yai09','803078':'03Bang Yai10','803098':'03Bang Yai13','803123':'03Bang Yai16','803104':'03Bang Yai13',
  '803067':'03Bang Bua Thong14','803084':'03Bang Bua Thong18','803087':'03Bang Bua Thong19','803100':'03Bang Bua Thong21','803124':'03Bang Bua Thong22','803126':'03Bang Bua Thong27',
  '803081':'03Sai Noi08','803086':'03Sai Noi09','803072':'03Sai Noi07','803125':'03Sai Noi10','603941':'03Sai Noi06','803073':'03Bang Bua Thong16','803103':'03Bang Bua Thong23','803101':'03Sai Noi10','803131':'03Bang Bua Thong28','803107':'03Sai Noi006'
};
const FIELDS={
  awb:['หมายเลขพัสดุ','เลขพัสดุ','AWB'],source:['แหล่งที่มาคำสั่งซื้อ','แหล่งที่มา','Order source'],platform:['แพลตฟอร์ม','Platform'],
  recordDate:['วันที่','วันที่ได้รับ','วันที่บันทึก'],deadlineNormal:['กำหนดเวลาการเซ็นรับที่เลื่อนออกไป','เวลาที่คาดว่าจะเกินกำหนด'],
  deadlineReverse:['ตีกลับ-ระยะเวลาที่เกินระยะเวลา Long-tail'],targetNormal:['ส่งปกติ-เวลา Long-tail (H)'],
  remainingNormal:['ส่งปกติ-ระยะเวลา Long-tail ที่้เหลือ','ส่งปกติ-ระยะเวลา Long-tail ที่เหลือ'],targetReverse:['逆向-长尾单时效(H)'],
  remainingReverse:['ตีกลับ-ระยะเวลา Long-tail ที่เหลือ','ตีกลับ-ระยะเวลา Long-tail ที่้เหลือ'],submittedAt:['เวลายื่นเรื่องพัสดุตีกลับ'],
  checkedAt:['ตรวจสอบเวลาตีกลับ'],lastScanType:['ประเภทการสแกนล่าสุด'],lastScanAt:['เวลาที่สแกนล่าสุด'],
  lastBranch:['DP ที่สแกนล่าสุด'],lastBranchCode:['最后扫描DP编码'],destBranch:['สาขาปลายทาง'],destBranchCode:['目的DP编码'],
  overdueType:['ประเภทเกินเวลา'],reverseFlag:['พัสดุตีกลับหรือไม่'],closedFlag:['แน่ใจหรือไม่ว่าจะปิดงาน']
};
const norm=v=>String(v??'').trim().toLowerCase().replace(/\s+/g,'');
const plain=v=>v==null?'':String(v).trim();
const num=v=>{const n=Number(String(v??'').replace(/,/g,'').trim());return String(v??'').trim()!==''&&Number.isFinite(n)?n:null};
const fmt=n=>Number(n||0).toLocaleString('th-TH');
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function showNotice(msg){notice.textContent=msg||'';notice.classList.toggle('show',!!msg)}

// Minimal XLSX reader: it reads the OOXML ZIP in-browser, without sending files to a server.
async function unzipXlsx(file){
  const buf=await file.arrayBuffer(),view=new DataView(buf),bytes=new Uint8Array(buf),decoder=new TextDecoder();
  const u16=i=>view.getUint16(i,true),u32=i=>view.getUint32(i,true);
  let eocd=-1;for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65558);i--){if(u32(i)===0x06054b50){eocd=i;break}}
  if(eocd<0)throw new Error('ไม่พบโครงสร้าง XLSX');
  const count=u16(eocd+10),central=u32(eocd+16),entries=new Map();let pos=central;
  for(let n=0;n<count;n++){
    if(u32(pos)!==0x02014b50)throw new Error('ตารางไฟล์ XLSX ไม่สมบูรณ์');
    const method=u16(pos+10),compressed=u32(pos+20),nameLen=u16(pos+28),extraLen=u16(pos+30),commentLen=u16(pos+32),local=u32(pos+42);
    const name=decoder.decode(bytes.slice(pos+46,pos+46+nameLen));const localName=u16(local+26),localExtra=u16(local+28),start=local+30+localName+localExtra;
    const packed=bytes.slice(start,start+compressed);let content;
    if(method===0)content=packed;else if(method===8){if(typeof DecompressionStream==='undefined')throw new Error('เบราว์เซอร์นี้ยังอ่านไฟล์ XLSX แบบบีบอัดไม่ได้');const ds=new DecompressionStream('deflate-raw');content=new Uint8Array(await new Response(new Blob([packed]).stream().pipeThrough(ds)).arrayBuffer())}else{pos+=46+nameLen+extraLen+commentLen;continue}
    entries.set(name,decoder.decode(content));pos+=46+nameLen+extraLen+commentLen;
  }
  return entries;
}
function xml(text){return new DOMParser().parseFromString(text,'application/xml')}
function excelDate(serial){const d=new Date(Date.UTC(1899,11,30)+Number(serial)*86400000);return d.toISOString().replace('T',' ').slice(0,19)}
function colNum(ref){let n=0;for(const c of ref.match(/^[A-Z]+/i)?.[0]||''){n=n*26+c.toUpperCase().charCodeAt(0)-64}return n-1}
function textOf(el,tag){return el.getElementsByTagNameNS('*',tag)[0]?.textContent||''}
async function parseXlsx(file){
  if(!file.name.toLowerCase().endsWith('.xlsx'))throw new Error('รองรับเฉพาะไฟล์ .xlsx');
  if(file.size>30*1024*1024)throw new Error('ไฟล์เกิน 30 MB');
  const entries=await unzipXlsx(file),shared=[];
  if(entries.has('xl/sharedStrings.xml')){const doc=xml(entries.get('xl/sharedStrings.xml'));for(const si of doc.getElementsByTagNameNS('*','si'))shared.push([...si.getElementsByTagNameNS('*','t')].map(t=>t.textContent).join(''))}
  const styles=entries.has('xl/styles.xml')?xml(entries.get('xl/styles.xml')):null,custom={};
  if(styles)for(const f of styles.getElementsByTagNameNS('*','numFmt'))custom[f.getAttribute('numFmtId')]=f.getAttribute('formatCode');
  const xfs=styles?[...styles.getElementsByTagNameNS('*','cellXfs')[0]?.children||[]]:[];
  const isDateStyle=style=>{const id=Number(xfs[style]?.getAttribute('numFmtId')||0),fmt=custom[id]||'';return [14,15,16,17,18,19,20,21,22,45,46,47].includes(id)||(/[ymdhis]/i.test(fmt)&&!/["].*[ymdhis].*[\"]/.test(fmt))};
  const sheets=[...entries.keys()].filter(n=>/^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();const output=[];
  for(const sheetPath of sheets){const doc=xml(entries.get(sheetPath)),matrix=[];
    for(const rowEl of doc.getElementsByTagNameNS('*','row')){const row=[];for(const cell of rowEl.getElementsByTagNameNS('*','c')){const idx=colNum(cell.getAttribute('r'));if(idx<0)continue;const type=cell.getAttribute('t'),v=textOf(cell,'v');let value='';
      if(type==='s')value=shared[Number(v)]??'';else if(type==='inlineStr')value=[...cell.getElementsByTagNameNS('*','t')].map(t=>t.textContent).join('');else if(type==='b')value=v==='1'?'TRUE':'FALSE';else if(v!==''&&isDateStyle(Number(cell.getAttribute('s')||0))&&Number.isFinite(Number(v)))value=excelDate(v);else value=v;
      row[idx]=value;
    }matrix[Number(rowEl.getAttribute('r'))-1]=row}
    let headerIndex=-1,headers=[];for(let i=0;i<Math.min(25,matrix.length);i++){const row=matrix[i]||[];if(findCol(row,FIELDS.awb)!=null&&findCol(row,FIELDS.overdueType)!=null){headerIndex=i;headers=row;break}}
    if(headerIndex<0)continue;const indexes=Object.fromEntries(Object.entries(FIELDS).map(([k,candidates])=>[k,findCol(headers,candidates)]));
    for(let i=headerIndex+1;i<matrix.length;i++){const row=matrix[i]||[],get=k=>{const j=indexes[k];return j!=null?plain(row[j]):''};const awb=get('awb');if(!awb)continue;const item=Object.fromEntries(Object.keys(FIELDS).map(k=>[k,get(k)]));item.awb=awb;item.sourceFile=file.name;item.platform=item.platform||item.source;output.push(classify(item))}
  }
  if(!output.length)throw new Error('ไม่พบหัวตาราง “หมายเลขพัสดุ” และ “ประเภทเกินเวลา”');return output;
}
function findCol(headers,names){const map=new Map(headers.map((h,i)=>[norm(h),i]).filter(([h])=>h));for(const name of names){if(map.has(norm(name)))return map.get(norm(name))}return null}
function classify(r){const flag=plain(r.reverseFlag).toUpperCase(),overdue=r.overdueType||'',reverse=flag==='Y';const remaining=reverse?r.remainingReverse:r.remainingNormal,n=num(remaining),token=norm(remaining);let status;
  if(token.includes('已超时')||token.includes('เกินกำหนด')||token.includes('เกินsla')||token.includes('overdue'))status='เกิน SLA';
  else if(n!==null)status=n>0?'ภายใน SLA':'เกิน SLA';
  else if(overdue.includes('วันนี้'))status='ครบกำหนดวันนี้';
  else if(remaining)status='มีค่า SLA ในไฟล์';else status='ไม่มีค่าคงเหลือในไฟล์';
  const code=branchCode(r.lastBranchCode);return {...r,direction:reverse?'ตีกลับ':'ส่งปกติ',returnFlag:flag,remaining,targetHours:reverse?r.targetReverse:r.targetNormal,deadline:reverse?(r.deadlineReverse||r.deadlineNormal):r.deadlineNormal,status,branchCode:code,branch:branchNames[code]||r.lastBranch||'ไม่ระบุสาขา',branchRM:rmByCode[code]||''};
}
function branchCode(v){const n=num(v);return n!==null&&Number.isInteger(n)?String(n):plain(v)}
function parseDate(s){if(!s)return null;const m=String(s).match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;const d=new Date(s);return Number.isNaN(+d)?null:d.toISOString().slice(0,10)}
function selectedCodes(){return [...document.querySelectorAll('#branchOptions input:checked')].map(x=>x.value)}
function currentRows(){const from=$('#fromDate').value,to=$('#toDate').value,dir=$('#direction').value,st=$('#status').value,flag=$('#returnFlag').value,codes=selectedCodes(),q=$('#search').value.trim().toLowerCase();return rawRows.filter(r=>{const d=parseDate(r.recordDate);if(from&&d&&d<from)return false;if(to&&d&&d>to)return false;if(dir!=='all'&&r.direction!==dir)return false;if(flag!=='all'&&r.returnFlag!==flag)return false;if(st!=='all'&&r.status!==st)return false;if(codes.length&&!codes.includes(String(r.branchCode)))return false;const rm=$('#rmFilter')?.value||'RM M';if(rm!=='all'&&rm!=='custom'&&rmByCode[String(r.branchCode)]!==rm)return false;if(String(r.branchCode)==='803120')return false;if(q&&!r.awb.toLowerCase().includes(q))return false;return true})}
function sortedBranchAwbRows(){return currentRows().filter(r=>r.branchCode&&['N','Y'].includes(r.returnFlag)).sort((a,b)=>{const d=Number(a.branchCode)-Number(b.branchCode);return (Number.isFinite(d)&&d!==0?d:a.branchCode.localeCompare(b.branchCode))||a.awb.localeCompare(b.awb)})}
function tagClass(s){return s==='เกิน SLA'?'over':s==='ครบกำหนดวันนี้'?'today':s==='ภายใน SLA'?'ok':'neutral'}
function pct(n,d){return d?`${(n/d*100).toFixed(1)}%`:'0.0%'}
function bars(target,entries,total){target.innerHTML=entries.length?entries.map(e=>`<div class="bar-row"><span class="bar-name" title="${esc(e.label)}">${esc(e.label)}</span><div class="bar-track"><div class="bar-fill ${e.color||''}" style="width:${total?Math.max(2,e.value/total*100):0}%"></div></div><span class="bar-count">${fmt(e.value)}</span></div>`).join(''):'<div class="empty">ไม่มีรายการในเงื่อนไขนี้</div>'}
function renderFiles(){$('#fileList').innerHTML=fileMeta.map(f=>`<div class="file-chip"><b>XLSX</b><span>${esc(f.name)}</span></div>`).join('');$('#fileEmpty').classList.toggle('hidden',fileMeta.length>0);$('#heroFiles').textContent=fmt(fileMeta.length)}
function render(){const rows=currentRows(),n=rows.length,over=rows.filter(r=>r.status==='เกิน SLA').length,today=rows.filter(r=>r.status==='ครบกำหนดวันนี้').length,ok=rows.filter(r=>r.status==='ภายใน SLA').length;$('#heroRows').textContent=fmt(n);$('#cTotal').textContent=fmt(n);$('#cFiles').textContent=`${fmt(fileMeta.length)} ไฟล์ · ${fmt(duplicates)} แถวซ้ำตัดออก`;$('#cOver').textContent=fmt(over);$('#cOverPct').textContent=`${pct(over,n)} ของมุมมองปัจจุบัน`;$('#cToday').textContent=fmt(today);$('#cOk').textContent=fmt(ok);$('#cOkPct').textContent=`${pct(ok,n)} ของมุมมองปัจจุบัน`;
  const labels=['เกิน SLA','ครบกำหนดวันนี้','ภายใน SLA','ไม่มีค่าคงเหลือในไฟล์','มีค่า SLA ในไฟล์'],colors=['red','amber','','',''];bars($('#statusBars'),labels.map((x,i)=>({label:x,value:rows.filter(r=>r.status===x).length,color:colors[i]})).filter(x=>x.value),n);$('#statusTotal').textContent=`${fmt(n)} รายการ`;
  const by=new Map();rows.forEach(r=>{const label=`${r.branch}${r.branchCode?` · ${r.branchCode}`:''}`;by.set(label,(by.get(label)||0)+1)});const top=[...by].sort((a,b)=>b[1]-a[1]).slice(0,7).map(([label,value])=>({label,value}));bars($('#branchBars'),top,Math.max(1,...top.map(x=>x.value)));$('#branchTotal').textContent=`${fmt(by.size)} สาขา`;
  const grouped=new Map();rows.forEach(r=>{const k=String(r.branchCode||'ไม่ระบุ');if(!grouped.has(k))grouped.set(k,{N:[],Y:[]});if(r.returnFlag==='Y')grouped.get(k).Y.push(r.awb);else if(r.returnFlag==='N')grouped.get(k).N.push(r.awb)});$('#pivotBody').innerHTML=[...grouped].sort((a,b)=>Number(a[0])-Number(b[0])).map(([code,v])=>`<tr><td>${esc(branchNames[code]||'')} <span class="subtle">${esc(code)}</span></td><td>${fmt(v.N.length)}</td><td class="mono awb-list">${esc(v.N.join(', ')||'—')}</td><td>${fmt(v.Y.length)}</td><td class="mono awb-list">${esc(v.Y.join(', ')||'—')}</td><td>${fmt(v.N.length+v.Y.length)}</td><td><button class="button row-copy" data-code="${esc(code)}">คัดลอก</button></td></tr>`).join('')||'<tr><td colspan="7" class="empty">ไม่มีสาขาในตัวกรอง</td></tr>';
  const branchAwb=sortedBranchAwbRows();$('#copyListBody').innerHTML=branchAwb.length?branchAwb.slice(0,2000).map(r=>`<tr><td>${esc(parseDate(r.recordDate)||r.recordDate||'—')}</td><td class="mono">${esc(r.awb)}</td><td>${esc(r.branchCode)} · ${esc(r.branch)}</td><td>${r.returnFlag==='Y'?'Y · ตีกลับ':'N · นำจ่าย'}</td></tr>`).join(''):'<tr><td colspan="4" class="empty">ไม่มีรายการในตัวกรอง</td></tr>';$('#copyListMeta').textContent=`เรียงตามรหัสสาขา · ${fmt(branchAwb.length)} รายการ · คัดลอกพร้อมหัวตารางได้`;
  $('#tbody').innerHTML=rows.length?rows.slice(0,500).map(r=>`<tr><td>${esc(parseDate(r.recordDate)||r.recordDate||'—')}</td><td>${esc(r.branch)}<div class="subtle">${esc(r.branchCode||'')}</div></td><td class="mono">${esc(r.awb)}</td><td>${r.returnFlag==='Y'?'ตีกลับ':r.returnFlag==='N'?'นำจ่าย':'—'}</td><td>${esc(r.source||r.platform||'—')}</td><td>${esc(r.direction)}</td><td><span class="tag ${tagClass(r.status)}">${esc(r.status)}</span></td><td>${esc(r.deadline||'—')}</td><td>${esc(r.targetHours||'—')}</td><td>${esc(r.remaining||'—')}</td><td>${esc(r.lastScanType||'—')}<div class="subtle">${esc(r.lastScanAt||'')}</div></td><td><button class="button row-copy" data-code="${esc(r.branchCode)}" data-awb="${esc(r.awb)}">คัดลอก</button></td></tr>`).join(''):'<tr><td colspan="12"><div class="empty-state"><div class="big">⌕</div><b>ไม่พบรายการตามตัวกรอง</b><span>ลองเปลี่ยนสาขา ช่วงวันที่ หรือสถานะ SLA</span></div></td></tr>';
  $('#tableMeta').textContent=`แสดง ${fmt(Math.min(n,500))} จาก ${fmt(n)} รายการ${n>500?' · จำกัดตารางที่ 500 แถวแรก':''}`;$('#export').disabled=n===0;$('#copyFiltered').disabled=n===0;$('#copyPivot').disabled=n===0;$('#copyBranchAwb').disabled=branchAwb.length===0;$('#dedupeInfo').textContent=`อ่าน ${fmt(rawRows.length+duplicates)} แถว · ใช้ ${fmt(rawRows.length)} AWB ไม่ซ้ำ`}
async function loadFiles(files){const selected=[...files];if(!selected.length)return;showNotice('กำลังอ่านไฟล์ในเบราว์เซอร์…');$('#heroUpload').disabled=true;try{let parsed=[];const warnings=[];for(const file of selected){try{parsed.push(...await parseXlsx(file))}catch(e){warnings.push(`${file.name}: ${e.message}`)}}if(!parsed.length)throw new Error(warnings.join(' · ')||'ไม่มีแถวข้อมูลที่อ่านได้');const map=new Map();duplicates=0;for(let i=0;i<parsed.length;i++){const r=parsed[i],old=map.get(r.awb);if(old){duplicates++;if((r.lastScanAt||'')>=(old.lastScanAt||''))map.set(r.awb,r)}else map.set(r.awb,r)}rawRows=[...map.values()];fileMeta=selected.filter(f=>parsed.some(r=>r.sourceFile===f.name)).map(f=>({name:f.name}));renderFiles();buildBranchOptions();showNotice(warnings.length?warnings.join(' · '):'อ่านไฟล์สำเร็จ · '+selected.length+' ไฟล์ · '+fmt(parsed.length)+' แถวก่อนตัดรายการซ้ำ');render()}catch(e){showNotice(e.message||'อ่านไฟล์ไม่สำเร็จ กรุณาตรวจว่าเป็น .xlsx จากรายงาน SLA')}finally{$('#heroUpload').disabled=false}}
function branchButtonText(){const rm=$('#rmFilter').value,n=selectedCodes().length;return rm==='all'?'ทุกสาขา ▾':rm==='custom'?`เลือก ${n} สาขา ▾`:`${rm} · ${n} สาขา ▾`}
function setRM(rm){$('#rmFilter').value=rm;document.querySelectorAll('#branchOptions input').forEach(x=>x.checked=rm==='all'||(rm!=='custom'&&rmByCode[x.value]===rm));$('#branchButton').textContent=branchButtonText();render()}
function buildBranchOptions(){const all=[...new Set([...Object.keys(branchNames),...rawRows.map(r=>String(r.branchCode||''))].filter(c=>c&&c!=='803120'))].sort();$('#branchOptions').innerHTML=all.map(c=>`<label><input type="checkbox" value="${esc(c)}" ${rmCodes.has(c)?'checked':''}><span>${esc(c)} · ${esc(branchNames[c]||'สาขาในไฟล์')} · ${esc(rmByCode[c]||'ไม่ระบุ RM')}</span></label>`).join('');$('#rmFilter').value='RM M';$('#branchButton').textContent=branchButtonText();$('#rmFilter').addEventListener('change',()=>setRM($('#rmFilter').value));$('#branchOptions').querySelectorAll('input').forEach(x=>x.addEventListener('change',()=>{$('#rmFilter').value='custom';$('#branchButton').textContent=branchButtonText();render()}))}
drop.addEventListener('click',()=>input.click());$('#heroUpload').addEventListener('click',()=>input.click());input.addEventListener('change',e=>loadFiles(e.target.files));drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('drag')});drop.addEventListener('dragleave',()=>drop.classList.remove('drag'));drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('drag');loadFiles(e.dataTransfer.files)});
['fromDate','toDate','direction','returnFlag','status','search'].forEach(id=>$('#'+id).addEventListener(id==='search'?'input':'change',render));$('#branchButton').addEventListener('click',()=>$('#branchPop').classList.toggle('hidden'));$('#branchSearch').addEventListener('input',e=>{const q=e.target.value.toLowerCase();$('#branchOptions').querySelectorAll('label').forEach(el=>el.classList.toggle('hidden',!el.textContent.toLowerCase().includes(q)))});$('#selectRm').addEventListener('click',()=>setRM('RM M'));$('#selectAllBranches').addEventListener('click',()=>setRM('all'));$('#clearBranches').addEventListener('click',()=>setRM('custom'));$('#reset').addEventListener('click',()=>{$('#fromDate').value='';$('#toDate').value='';$('#direction').value='all';$('#returnFlag').value='all';$('#status').value='all';$('#search').value='';setRM('RM M')});
$('#copyFiltered').addEventListener('click',async()=>{const text=currentRows().map(r=>`${r.awb}\t${r.branchCode}`).join('\n');try{await navigator.clipboard.writeText(text);showNotice(`คัดลอก ${fmt(currentRows().length)} รายการแล้ว`)}catch{showNotice('คัดลอกไม่ได้: กรุณาเปิดหน้าเว็บผ่านเบราว์เซอร์ที่อนุญาต Clipboard')}});$('#copyPivot').addEventListener('click',async()=>{const rows=currentRows(),groups=new Map();rows.forEach(r=>{if(!groups.has(r.branchCode))groups.set(r.branchCode,{N:[],Y:[]});if(r.returnFlag==='N'||r.returnFlag==='Y')groups.get(r.branchCode)[r.returnFlag].push(r.awb)});const text=[...groups].sort((a,b)=>Number(a[0])-Number(b[0])).map(([c,v])=>`${c}\tนำจ่าย (${v.N.length}): ${v.N.join(', ')||'—'}\tตีกลับ (${v.Y.length}): ${v.Y.join(', ')||'—'}\tรวม ${v.N.length+v.Y.length}`).join('\n');try{await navigator.clipboard.writeText(text);showNotice('คัดลอกสรุปพร้อมเลข AWB แล้ว')}catch{showNotice('คัดลอกไม่ได้: เบราว์เซอร์ไม่อนุญาต Clipboard')}});document.addEventListener('click',async e=>{const b=e.target.closest('.row-copy');if(!b)return;if(b.dataset.awb){await navigator.clipboard.writeText(`${b.dataset.awb}\t${b.dataset.code}`)}else{const c=b.dataset.code;await navigator.clipboard.writeText(currentRows().filter(r=>r.branchCode===c).sort((a,b)=>a.awb.localeCompare(b.awb)).map(r=>`${r.branchCode}\t${r.awb}\t${r.returnFlag==='Y'?'ตีกลับ':'นำจ่าย'}`).join('\n'))}showNotice('คัดลอกข้อมูลพร้อม AWB แล้ว')});
$('#copyBranchAwb').addEventListener('click',async()=>{const rows=sortedBranchAwbRows(),text=['วันที่\tAWB\tสาขา รหัส / ชื่อ\tประเภท SLA',...rows.map(r=>`${parseDate(r.recordDate)||r.recordDate||'—'}\t${r.awb}\t${r.branchCode} · ${r.branch}\t${r.returnFlag==='Y'?'Y · ตีกลับ':'N · นำจ่าย'}`)].join('\n');try{await navigator.clipboard.writeText(text);showNotice(`คัดลอกรายงานแบบแชท ${fmt(rows.length)} รายการแล้ว`)}catch{showNotice('คัดลอกไม่ได้: เบราว์เซอร์ไม่อนุญาต Clipboard')}});
$('#export').addEventListener('click',()=>{const rows=currentRows(),cols=['recordDate','branchCode','branch','awb','source','direction','status','deadline','targetHours','remaining','lastScanType','lastScanAt','sourceFile'];const q=v=>'"'+String(v??'').replace(/"/g,'""')+'"';const csv='\ufeff'+[cols.join(','),...rows.map(r=>cols.map(c=>q(r[c])).join(','))].join('\r\n');const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download='sla-dashboard-filtered.csv';a.click();URL.revokeObjectURL(a.href)});
buildBranchOptions();render();

