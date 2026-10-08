import { api, configured, preparePages, uploadPages, combinePages, supabase } from './backend.js';
import { makeExcel, makePdf, parseFinishedWorkbook, validateList } from './exports.js';

const $ = (id) => document.getElementById(id);
const state = { files: [], imageUrls: [], uploads: [], data: null, currentPage: 0, listId: null, finalized: false, origin: 'recognition', editCount: 0, role: 'sales' };

function escapeHtml(value) { return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function showMessage(message, error=false) { const el=$('review-section').classList.contains('hidden')?$('global-message'):$('message'); el.textContent=message; el.classList.remove('hidden'); el.classList.toggle('error',error); el.scrollIntoView({behavior:'smooth',block:'nearest'}); }
function clearMessage() { $('message').classList.add('hidden');$('global-message').classList.add('hidden'); }

async function loadStatus() {
  const data=await api('status');
  state.role=data.role;
  const connected=data.key_mode==='personal'?data.has_personal_key:data.has_shared_key;
  $('key-status').textContent=connected ? `Gemini connected · ${data.key_mode}` : `${data.key_mode} API key needed`;
  $('key-status').classList.toggle('off',!connected);
  $('key-mode').value=data.key_mode;
  $('settings-key-state').textContent=`Shared: ${data.has_shared_key?'configured':'missing'} · Personal: ${data.has_personal_key?'configured':'missing'}`;
  $('save-shared-key').classList.toggle('hidden',data.role!=='admin');
  $('admin-section').classList.toggle('hidden',data.role!=='admin');
  $('training-count').textContent=`${data.approved_examples||0} approved examples`;
  if(data.role==='admin') loadTrainingQueue().catch(error=>showMessage(error.message,true));
  return data;
}

function setFinalized(value) {
  state.finalized=value;
  $('download-excel').disabled=!value;
  $('download-pdf').disabled=!value;
  $('final-button').textContent=value?'✓ Final saved — update if edited':'✓ Final — save verified list';
}

function setFiles(files) {
  const known=new Set(state.files.map(f=>`${f.name}:${f.size}:${f.lastModified}`));
  for(const file of files) { const id=`${file.name}:${file.size}:${file.lastModified}`; if(!known.has(id)){state.files.push(file);known.add(id);} }
  renderFiles();
}
function renderFiles() {
  $('file-list').innerHTML=state.files.map((file,i)=>`<div class="file-row"><span class="file-index">${i+1}</span><img class="file-thumb" src="${file.type.startsWith('image/')?escapeHtml(URL.createObjectURL(file)):''}" alt=""><span class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span><div class="file-controls"><button class="mini-button" data-file-action="up" data-index="${i}" title="Move up">↑</button><button class="mini-button" data-file-action="down" data-index="${i}" title="Move down">↓</button><button class="mini-button" data-file-action="remove" data-index="${i}" title="Remove">×</button></div></div>`).join('');
  $('upload-count').textContent=state.files.length?`${state.files.length} file${state.files.length===1?'':'s'} selected`:'No files selected';
  $('process-button').disabled=!state.files.length;
}

function setPage(index) {
  state.currentPage=index;
  document.querySelectorAll('.page-tab').forEach((el,i)=>el.classList.toggle('active',i===index));
  $('page-indicator').textContent=`${index+1} / ${state.imageUrls.length}`;
  const url=state.imageUrls[index];
  $('source-view').classList.remove('zoomed');
  $('source-view').innerHTML=url?`<img src="${url}" alt="Original packing list page ${index+1}">`:'<span>PDF page preview unavailable</span>';
}

function prepareImages() {
  state.imageUrls.filter(url=>url.startsWith('blob:')).forEach(url=>URL.revokeObjectURL(url));
  const pageCount=state.uploads.length;
  state.imageUrls=state.uploads.map(x=>x.url);
  $('page-tabs').innerHTML=Array.from({length:pageCount},(_,i)=>`<button type="button" class="page-tab" data-page="${i}">Page ${i+1}</button>`).join('');
  if(state.imageUrls.length) setPage(0);
  else {$('source-view').innerHTML='<span>Source photos were not retained for this saved list.</span>';$('page-indicator').textContent='';}
}

function updateSummary() {
  if(!state.data)return;
  $('box-count').textContent=`${state.data.boxes.length} boxes`;
  const count=state.data.boxes.reduce((sum,box)=>sum+box.items.filter(item=>item.code||item.size||item.quantity||item.note).length,0);
  $('item-count').textContent=`${count} items`;
}

function rowHtml(boxIndex,item,index) {
  const uncertain=item.needs_review?' uncertain':'';
  const f=(name,value,placeholder='')=>`<input class="item-input${uncertain}" data-box="${boxIndex}" data-row="${index}" data-field="${name}" value="${escapeHtml(value)}" placeholder="${placeholder}" title="${escapeHtml(item.source_text||'')}">`;
  const source=item.source_page?`<button class="source-link" data-page="${item.source_page-1}" type="button">Page ${item.source_page}</button>`:'';
  return `<tr><td>${index+1}</td><td>${f('code',item.code,'PC-404')}</td><td>${f('size',item.size,'S / M')}</td><td>${f('quantity',item.quantity??'','Qty')}</td><td>${f('note',item.note,'Only base / cup')}</td><td>${source}</td><td><button class="mini-button" type="button" data-row-action="remove" data-box="${boxIndex}" data-row="${index}" title="Remove row">×</button></td></tr>`;
}

function renderBoxes() {
  const boxes=state.data.boxes;
  $('boxes').innerHTML=boxes.map((box,bi)=>{
    const rows=Math.max(5,box.items.length);
    const body=Array.from({length:rows},(_,i)=>rowHtml(bi,box.items[i]||{code:'',size:'',quantity:null,note:''},i)).join('');
    const needs=box.needs_review||box.items.some(x=>x.needs_review);
    return `<article class="box-card${needs?' review':''}"><div class="box-header"><div class="box-title">Box <input class="box-number" type="number" min="1" value="${box.number}" data-box-number="${bi}">${needs?'<span class="review-chip">Check handwriting</span>':''}</div><div class="box-actions">${needs?`<button class="link-button" type="button" data-box-action="checked" data-box="${bi}">Mark checked</button>`:''}<button class="link-button" type="button" data-box-action="add" data-box="${bi}">+ Add row</button><button class="link-button danger" type="button" data-box-action="remove" data-box="${bi}">Remove box</button></div></div><table class="entry-table"><thead><tr><th>No.</th><th>Item code</th><th>Size</th><th>Qty.</th><th>Note</th><th>Source</th><th></th></tr></thead><tbody>${body}</tbody></table><div class="box-footer"><span>${box.items.length} detected item${box.items.length===1?'':'s'}</span><span>${Math.max(0,5-box.items.length)} blank print row${Math.max(0,5-box.items.length)===1?'':'s'} minimum</span></div></article>`;
  }).join('');
  updateSummary();
}

function renderWarnings() {
  const warnings=state.data.warnings||[];
  $('warnings').classList.toggle('hidden',!warnings.length);
  $('warnings').innerHTML=warnings.length?`<strong>Items to check</strong><ul>${warnings.map(w=>`<li>${escapeHtml(w)}</li>`).join('')}</ul>`:'';
}

function showReview(data) {
  state.data=data;
  state.editCount=0;
  $('keep-for-training').checked=state.origin==='manual_import'||Math.random()<.1;
  setFinalized(false);
  $('customer').value=data.customer||'';
  $('packing-date').value=data.packing_date||new Date().toLocaleDateString('en-GB');
  $('private-mark').value=data.private_mark||'';
  $('transport').value=data.transport||'';
  $('review-section').classList.remove('hidden');
  document.querySelectorAll('.step').forEach((el,i)=>el.classList.toggle('active',i===1));
  prepareImages();renderWarnings();renderBoxes();
  $('review-section').scrollIntoView({behavior:'smooth'});
}

async function processFiles() {
  clearMessage();
  $('process-button').disabled=true;$('progress').classList.remove('hidden');$('progress-text').textContent='Preparing pages…';
  try {
    const pages=await preparePages(state.files);
    state.uploads=await uploadPages(pages);
    state.origin='recognition'; state.listId=null;
    const extracted=[];
    for(let i=0;i<state.uploads.length;i++){
      $('progress-text').textContent=`Reading page ${i+1} of ${state.uploads.length}…`;
      const last=extracted.at(-1)?.boxes?.at(-1)?.number||null;
      try{const result=await api('recognize',{path:state.uploads[i].path,continuation:last});extracted.push(result.page);}
      catch(error){if(extracted.length){const partial=combinePages(extracted);partial.warnings.push(`Page ${i+1} could not be read: ${error.message}`);showReview(partial);}throw error;}
      $('progress-fill').style.width=`${Math.round(100*(i+1)/state.uploads.length)}%`;
    }
    showReview(combinePages(extracted));
  }catch(error){showMessage(error.message,true);}finally{$('process-button').disabled=false;$('progress').classList.add('hidden');}
}

async function importTraining() {
  clearMessage();
  const sheet=$('training-sheet').files[0], photos=Array.from($('training-photos').files);
  if(!sheet||!photos.length){showMessage('Choose a finished Excel and its matching photos.',true);return;}
  $('import-training').disabled=true;
  try{
    const data=await parseFinishedWorkbook(sheet);
    state.uploads=await uploadPages(await preparePages(photos));state.origin='manual_import';state.listId=null;
    showReview(data);
    showMessage(`Imported ${data.boxes.length} filled boxes. Compare with photos, then press Final. Admin will review it for shared training hints.`);
  }catch(error){showMessage(error.message,true);}finally{$('import-training').disabled=false;}
}

async function finalizeList() {
  clearMessage();
  if(!state.uploads.length&&!state.listId){showMessage('Process or import source photos first.',true);return;}
  const button=$('final-button');button.disabled=true;
  try{
    const data=collectData();validateList(data);
    const info=await api('finalize',{id:state.listId,paths:state.listId?[]:state.uploads.map(x=>x.path),hashes:state.uploads.map(x=>x.hash),data,origin:state.origin,keep_for_training:$('keep-for-training').checked});
    state.listId=info.id;
    setFinalized(true);
    loadHistory().catch(()=>{});
    showMessage(info.training_status==='training_candidate'?'Final saved. Source photos are queued for admin review. Excel and PDF are ready.':'Final saved. Excel and PDF are ready; source photos were deleted from cloud storage.');
    loadStatus().catch(()=>{});
  }catch(error){showMessage(error.message,true);}finally{button.disabled=false;}
}

function collectData() {
  const data=structuredClone(state.data);
  data.customer=$('customer').value.trim(); data.packing_date=$('packing-date').value.trim(); data.private_mark=$('private-mark').value.trim(); data.transport=$('transport').value.trim();
  data.boxes=data.boxes.map(box=>({...box,items:box.items.filter(item=>item.code||item.size||item.quantity!==null&&item.quantity!==''||item.note).map(item=>({...item,quantity:item.quantity===''?null:item.quantity}))}));
  return data;
}

async function exportFile(kind,preview=false) {
  clearMessage();
  try{
    if(!preview&&!state.finalized)throw new Error('Check the list and press Final before downloading.');
    const data=collectData();
    const blob=kind==='xlsx'?await makeExcel(data):makePdf(data);const url=URL.createObjectURL(blob);
    if(preview){const opened=window.open(url,'_blank');if(!opened)showMessage('Popup blocked. Allow popups to preview PDF.',true);}
    else{const filename=`${(data.customer||'PACKING LIST').replace(/[^A-Za-z0-9 _-]/g,'').slice(0,50)} ${data.packing_date||''}.${kind}`;const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();showMessage(`${kind.toUpperCase()} downloaded.`);}
    setTimeout(()=>URL.revokeObjectURL(url),60000);
  }catch(error){showMessage(error.message,true);}
}

$('file-input').addEventListener('change',e=>{setFiles(e.target.files);e.target.value='';});
const drop=$('drop-zone');
drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('dragging');});
drop.addEventListener('dragleave',()=>drop.classList.remove('dragging'));
drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('dragging');setFiles(e.dataTransfer.files);});
$('file-list').addEventListener('click',e=>{const button=e.target.closest('[data-file-action]');if(!button)return;const i=Number(button.dataset.index),action=button.dataset.fileAction;if(action==='remove')state.files.splice(i,1);else{const j=i+(action==='up'?-1:1);if(j>=0&&j<state.files.length)[state.files[i],state.files[j]]=[state.files[j],state.files[i]];}renderFiles();});
$('process-button').addEventListener('click',processFiles);
$('import-training').addEventListener('click',importTraining);
$('final-button').addEventListener('click',finalizeList);
$('settings-button').addEventListener('click',()=>$('settings-dialog').showModal());
$('save-key').addEventListener('click',()=>saveKey('personal'));
$('save-shared-key').addEventListener('click',()=>saveKey('shared'));
$('key-mode').addEventListener('change',async e=>{try{await api('set-key-mode',{mode:e.target.value});await loadStatus();}catch(error){$('settings-error').textContent=error.message;}});
$('page-tabs').addEventListener('click',e=>{const button=e.target.closest('[data-page]');if(button)setPage(Number(button.dataset.page));});
$('source-view').addEventListener('click',()=>{$('source-view').classList.toggle('zoomed');});
function markEdited(){setFinalized(false);state.editCount++;if(state.editCount>=2)$('keep-for-training').checked=true;}
document.querySelectorAll('#customer,#packing-date,#private-mark,#transport').forEach(input=>input.addEventListener('input',markEdited));
$('boxes').addEventListener('input',e=>{const target=e.target;markEdited();if(target.dataset.boxNumber!==undefined){state.data.boxes[Number(target.dataset.boxNumber)].number=Number(target.value);updateSummary();return;}if(target.dataset.field!==undefined){const bi=Number(target.dataset.box),ri=Number(target.dataset.row);const box=state.data.boxes[bi];while(box.items.length<=ri)box.items.push({code:'',size:'',quantity:null,note:'',source_page:null,source_text:'',needs_review:false});box.items[ri][target.dataset.field]=target.dataset.field==='quantity'?(target.value===''?null:Number(target.value)):target.value;updateSummary();}});
$('boxes').addEventListener('click',e=>{const page=e.target.closest('[data-page]');if(page){setPage(Number(page.dataset.page));return;}const row=e.target.closest('[data-row-action]');if(row){markEdited();const box=state.data.boxes[Number(row.dataset.box)];box.items.splice(Number(row.dataset.row),1);renderBoxes();return;}const button=e.target.closest('[data-box-action]');if(!button)return;markEdited();const bi=Number(button.dataset.box);if(button.dataset.boxAction==='remove'){state.data.boxes.splice(bi,1);}else if(button.dataset.boxAction==='checked'){state.data.boxes[bi].needs_review=false;state.data.boxes[bi].items.forEach(item=>item.needs_review=false);}else{state.data.boxes[bi].items.push({code:'',size:'',quantity:null,note:'',source_page:null,source_text:'',needs_review:false});}renderBoxes();});
$('add-box').addEventListener('click',()=>{markEdited();const number=state.data.boxes.reduce((m,b)=>Math.max(m,b.number),0)+1;state.data.boxes.push({number,items:[],needs_review:false});renderBoxes();});
$('preview-pdf').addEventListener('click',()=>exportFile('pdf',true));
$('download-excel').addEventListener('click',()=>exportFile('xlsx'));
$('download-pdf').addEventListener('click',()=>exportFile('pdf'));

async function saveKey(scope){
  const key=$('api-key').value.trim();$('settings-error').textContent='';
  try{await api('set-key',{scope,key});$('api-key').value='';await loadStatus();$('settings-dialog').close();showMessage(`${scope==='shared'?'Company shared':'Personal'} Gemini key updated.`);}
  catch(error){$('settings-error').textContent=error.message;}
}
async function loadTrainingQueue(){
  const {candidates}=await api('training-queue');
  $('training-queue').innerHTML=candidates.length?candidates.map(c=>`<div class="queue-card" data-case="${c.id}"><h4>${escapeHtml(c.data.customer||'Untitled list')}</h4><p>${c.data.boxes.length} boxes · ${escapeHtml(c.origin)} · ${new Date(c.created_at).toLocaleString()}</p><p>${c.source_paths.length} source page(s) retained for review</p><details><summary>View verified rows</summary>${c.data.boxes.map(box=>`<div class="queue-box"><strong>Box ${escapeHtml(box.number)}</strong><ul>${box.items.map(item=>`<li>${escapeHtml(item.code)} ${escapeHtml(item.size)} — ${escapeHtml(item.quantity)} ${escapeHtml(item.note||'')}</li>`).join('')}</ul></div>`).join('')}</details><div class="queue-actions"><button class="secondary-button" data-review="open">Open photos</button><button class="primary-button" data-review="approve">Approve example</button><button class="link-button danger" data-review="reject">Reject & delete photos</button></div></div>`).join(''):'No cases waiting for review.';
  state.candidates=candidates;
}
async function loadHistory(){
  const {data,error}=await supabase.from('packing_lists').select('id,data,status,source_paths,source_hashes,origin,created_at').order('created_at',{ascending:false}).limit(40);
  if(error)throw error;
  state.history=data||[];
  $('list-history').innerHTML=state.history.length?state.history.map(list=>`<button class="history-row" type="button" data-list="${list.id}"><strong>${escapeHtml(list.data.customer||'Untitled')}</strong><span>${escapeHtml(list.data.boxes.length)} boxes · ${escapeHtml(new Date(list.created_at).toLocaleDateString())} · ${escapeHtml(list.status.replaceAll('_',' '))}</span><span>Open →</span></button>`).join(''):'No saved lists yet.';
}
$('refresh-history').addEventListener('click',()=>loadHistory().catch(error=>showMessage(error.message,true)));
$('list-history').addEventListener('click',async e=>{
  const button=e.target.closest('[data-list]');if(!button)return;
  const list=state.history?.find(x=>x.id===button.dataset.list);if(!list)return;
  try{
    state.files=[];renderFiles();state.origin=list.origin;state.listId=list.id;
    state.uploads=[];
    for(const [index,path] of list.source_paths.entries()){
      const {data,error}=await supabase.storage.from('packing-sources').createSignedUrl(path,600);
      if(error)throw error;
      state.uploads.push({path,hash:list.source_hashes[index]||'',url:data.signedUrl});
    }
    showReview(structuredClone(list.data));
    $('keep-for-training').checked=list.source_paths.length>0;
    setFinalized(true);
  }catch(error){showMessage(error.message,true);}
});
$('training-queue').addEventListener('click',async e=>{
  const button=e.target.closest('[data-review]'), card=e.target.closest('[data-case]');if(!button||!card)return;
  const candidate=state.candidates?.find(c=>c.id===card.dataset.case);if(!candidate)return;
  try{
    if(button.dataset.review==='open'){
      for(const path of candidate.source_paths){const {data,error}=await supabase.storage.from('packing-sources').createSignedUrl(path,300);if(error)throw error;window.open(data.signedUrl,'_blank');}
    }else{
      const approve=button.dataset.review==='approve';
      if(!approve&&!confirm('Reject this training example and delete its source photos?'))return;
      await api('review-training',{id:candidate.id,approve});await loadTrainingQueue();await loadStatus();
    }
  }catch(error){showMessage(error.message,true);}
});
$('refresh-training').addEventListener('click',()=>loadTrainingQueue().catch(error=>showMessage(error.message,true)));
$('invite-button').addEventListener('click',async()=>{
  const email=$('invite-email').value.trim();
  try{await api('invite',{email});$('invite-email').value='';showMessage(`Invite sent to ${email}.`);}
  catch(error){showMessage(error.message,true);}
});
$('login-button').addEventListener('click',async()=>{
  if(!supabase){$('login-message').textContent='Supabase project is not configured yet.';return;}
  const email=$('login-email').value.trim();
  try{const {error}=await supabase.auth.signInWithOtp({email,options:{shouldCreateUser:false,emailRedirectTo:location.origin}});if(error)throw error;$('login-message').textContent='Check your email for the sign-in link.';}
  catch(error){$('login-message').textContent=error.message;}
});
$('sign-out').addEventListener('click',async()=>{await supabase.auth.signOut();location.reload();});
async function boot(){
  if(!configured){$('login-message').textContent='Supabase project is not configured yet.';return;}
  const {data:{session}}=await supabase.auth.getSession();
  if(!session)return;
  $('auth-screen').classList.add('hidden');$('app-main').classList.remove('hidden');
  document.querySelector('.top-actions').classList.remove('hidden');
  $('user-email').textContent=session.user.email||'';
  await loadStatus();
  await loadHistory();
}
boot().catch(error=>{$('login-message').textContent=error.message;});
supabase?.auth.onAuthStateChange((_event,session)=>{if(session&&$('app-main').classList.contains('hidden'))boot().catch(console.error);});
