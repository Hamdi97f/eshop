const SHEET_ID = '1ztGpWzH-TJzkdCfSU4S-XirGlNIpIeAjDQoYDlP5chk';
const ADMIN_KEY = 'CHANGE_THIS_ADMIN_KEY';

function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(name, headers){
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0 && headers) sh.appendRow(headers);
  return sh;
}
function setup_(){
  sheet_('Orders',['Date','Name','Phone','Governorate','Delegation','Address','Qty','Product Price','Delivery','Total','Status']);
  const c=sheet_('Config',['Key','Value']);
  if(c.getLastRow()===1){
    [['name','6 Dessous de Plat en Bois'],['price','29'],['delivery','6'],['color','#d8a94d'],['headline','حماية لسطحك بأناقة طبيعية'],['description','Pack de 6 pièces — عملي وأنيق ويحمي السطح من حرارة الأواني.'],['images','[]'],['mainImage','0']].forEach(r=>c.appendRow(r));
  }
}
function doGet(e){
  setup_();
  const a=(e.parameter.action||'config').toLowerCase();
  if(a==='config') return json_({ok:true,config:readConfig_()});
  if(a==='orders'){
    if(e.parameter.key!==ADMIN_KEY) return json_({ok:false,error:'unauthorized'});
    return json_({ok:true,orders:readOrders_()});
  }
  return json_({ok:true});
}
function doPost(e){
  setup_();
  let p={}; try{p=JSON.parse(e.postData.contents||'{}')}catch(err){return json_({ok:false,error:'invalid_json'})}
  const a=(p.action||'order').toLowerCase();
  if(a==='order'){
    const sh=sheet_('Orders');
    sh.appendRow([new Date(),p.name||'',p.phone||'',p.governorate||'',p.delegation||'',p.address||'',p.qty||1,p.product_price||29,p.delivery||6,p.total||35,'new']);
    return json_({ok:true,message:'order_saved'});
  }
  if(p.key!==ADMIN_KEY) return json_({ok:false,error:'unauthorized'});
  if(a==='saveConfig'){
    const c=sheet_('Config');
    const cfg=p.config||{}; const rows=c.getDataRange().getValues();
    Object.keys(cfg).forEach(k=>{let found=-1;for(let i=1;i<rows.length;i++)if(rows[i][0]===k){found=i+1;break}if(found===-1)c.appendRow([k,typeof cfg[k]==='object'?JSON.stringify(cfg[k]):String(cfg[k])]);else c.getRange(found,2).setValue(typeof cfg[k]==='object'?JSON.stringify(cfg[k]):String(cfg[k]));});
    return json_({ok:true});
  }
  if(a==='status'){
    const sh=sheet_('Orders'); const row=Number(p.row); if(row<2)return json_({ok:false,error:'bad_row'}); sh.getRange(row,11).setValue(p.status||'new'); return json_({ok:true});
  }
  return json_({ok:false,error:'unknown_action'});
}
function readConfig_(){
  const sh=sheet_('Config'); const out={}; sh.getDataRange().getValues().slice(1).forEach(r=>{if(r[0]){let v=r[1];try{v=JSON.parse(v)}catch(e){}out[r[0]]=v}}); return out;
}
function readOrders_(){
  const sh=sheet_('Orders'); const v=sh.getDataRange().getValues(); if(v.length<2)return [];
  return v.slice(1).map((r,i)=>({row:i+2,date:r[0],name:r[1],phone:r[2],governorate:r[3],delegation:r[4],address:r[5],qty:r[6],product_price:r[7],delivery:r[8],total:r[9],status:r[10]}));
}
