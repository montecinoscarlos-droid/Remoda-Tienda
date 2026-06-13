const SHEET_PRODUCTOS = 'Productos';
const SHEET_PEDIDOS = 'Pedidos';
const ADMIN_PASSWORD = '1234'; // Cambia esta clave de administrador
const ADMIN_TOKEN = 'REMODA_ADMIN_TOKEN_2026';
const DRIVE_FOLDER_NAME = 'ReModa_Fotos';
const WHATSAPP_TIENDA = '56928117329';

const HEADERS_PRODUCTOS = ['id','nombre','marca','talla','precio','stock','estado','foto1','foto2','whatsapp','actualizado'];
const HEADERS_PEDIDOS = ['fecha','cliente','detalle','total'];

function doGet(e){
  setup();
  const action = e && e.parameter ? e.parameter.action : '';
  if(action === 'listProducts') return json({ok:true, productos:getProducts()});
  return json({ok:true,mensaje:'Apps Script GET funcionando',whatsapp:WHATSAPP_TIENDA,productos:getProducts().length});
}

function doPost(e){
  try{
    setup();
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if(data.action === 'login') return json({ok:data.password===ADMIN_PASSWORD, token:data.password===ADMIN_PASSWORD?ADMIN_TOKEN:''});
    if(data.action === 'listProducts') return json({ok:true, productos:getProducts()});
    if(data.action === 'saveProduct'){
      if(data.token !== ADMIN_TOKEN) return json({ok:false,error:'No autorizado'});
      return json(saveProduct(data));
    }
    if(data.action === 'discountStock'){
      if(data.token !== ADMIN_TOKEN) return json({ok:false,error:'No autorizado'});
      return json(discountStock(data.id, Number(data.cantidad||1)));
    }
    if(data.action === 'deleteProduct'){
      if(data.token !== ADMIN_TOKEN) return json({ok:false,error:'No autorizado'});
      return json(deleteProduct(data.id));
    }
    if(data.action === 'saveOrder') return json(saveOrder(data));
    return json({ok:false,error:'Acción no válida'});
  }catch(err){return json({ok:false,error:String(err)})}
}

function json(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function ss(){return SpreadsheetApp.getActiveSpreadsheet()}

function setup(){
  const book=ss();
  let sh=book.getSheetByName(SHEET_PRODUCTOS)||book.insertSheet(SHEET_PRODUCTOS);
  if(sh.getLastRow()===0){
    sh.appendRow(HEADERS_PRODUCTOS);
  }else{
    asegurarColumnas(sh, HEADERS_PRODUCTOS);
  }
  let ped=book.getSheetByName(SHEET_PEDIDOS)||book.insertSheet(SHEET_PEDIDOS);
  if(ped.getLastRow()===0) ped.appendRow(HEADERS_PEDIDOS);
}

function normalizarClave(h){
  return String(h || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,'');
}

function canonHeader(h){
  const k=normalizarClave(h);
  if(k==='id'||k==='codigo'||k==='codigoproducto') return 'id';
  if(k==='nombre'||k==='producto'||k==='prenda') return 'nombre';
  if(k==='marca') return 'marca';
  if(k==='talla') return 'talla';
  if(k==='precio'||k==='valor') return 'precio';
  if(k==='stock'||k==='cantidad') return 'stock';
  if(k==='estado'||k==='disponibilidad') return 'estado';
  if(k==='foto1'||k==='imagen1'||k==='urlfoto1'||k==='foto1url') return 'foto1';
  if(k==='foto2'||k==='imagen2'||k==='urlfoto2'||k==='foto2url') return 'foto2';
  if(k==='whatsapp'||k==='wsp'||k==='telefono') return 'whatsapp';
  if(k==='actualizado'||k==='fechaactualizacion') return 'actualizado';
  return k;
}

function asegurarColumnas(sh, requeridos){
  const lastCol = Math.max(1, sh.getLastColumn());
  const head = sh.getRange(1,1,1,lastCol).getValues()[0];
  const canon = head.map(canonHeader);
  requeridos.forEach(h=>{
    if(canon.indexOf(h) === -1){
      sh.getRange(1, sh.getLastColumn()+1).setValue(h);
      canon.push(h);
    }
  });
}

function headerMap(sh){
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const map={};
  head.forEach((h,i)=>{ map[canonHeader(h)] = i+1; });
  return map;
}

function getProducts(){
  setup();
  const sh=ss().getSheetByName(SHEET_PRODUCTOS);
  const values=sh.getDataRange().getValues();
  if(values.length < 2) return [];
  const head=values[0].map(canonHeader);
  return values.slice(1).map(r=>{
    const o={};
    head.forEach((h,i)=>o[h]=r[i]);
    o.id = String(o.id || '').trim();
    if(!o.id) return null;
    if(!o.whatsapp) o.whatsapp=WHATSAPP_TIENDA;
    if(!o.estado) o.estado=Number(o.stock||0)>0?'Disponible':'Vendido';
    // IMPORTANTE: convierte enlaces antiguos de Drive a enlace de miniatura visible en GitHub Pages.
    o.foto1 = normalizarFotoDrive(o.foto1);
    o.foto2 = normalizarFotoDrive(o.foto2);
    return o;
  }).filter(Boolean);
}

function normalizarFotoDrive(url){
  url = String(url || '').trim();
  if(!url) return '';
  // Si la celda tiene solo el ID del archivo, también sirve.
  if(/^[a-zA-Z0-9_-]{20,}$/.test(url)){
    return 'https://drive.google.com/thumbnail?id=' + url + '&sz=w1200';
  }
  // Soporta: open?id=, uc?id=, uc?export=view&id=, file/d/ID/view, thumbnail?id=
  let m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/) || url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if(m && m[1]){
    return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w1200';
  }
  return url;
}

function generarId(){
  return 'RM' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Santiago', 'yyyyMMddHHmmss');
}

function saveProduct(data){
  setup();
  const p=data.producto||{};
  p.id = String(p.id || '').trim();
  if(!p.id) p.id = generarId();
  p.nombre = String(p.nombre || '').trim();
  p.marca = String(p.marca || '').trim();
  p.talla = String(p.talla || '').trim();
  p.precio = Number(p.precio || 0);
  p.stock = Number(p.stock || 0);
  p.estado = p.estado || (p.stock > 0 ? 'Disponible' : 'Vendido');
  p.whatsapp = p.whatsapp || WHATSAPP_TIENDA;

  if(data.foto1Base64) p.foto1=uploadBase64(data.foto1Base64, p.id+'_foto1');
  if(data.foto2Base64) p.foto2=uploadBase64(data.foto2Base64, p.id+'_foto2');

  const sh=ss().getSheetByName(SHEET_PRODUCTOS);
  asegurarColumnas(sh, HEADERS_PRODUCTOS);
  const map=headerMap(sh);
  const values=sh.getDataRange().getValues();
  const idCol=map.id;
  let row=-1;
  for(let i=1;i<values.length;i++){
    if(String(values[i][idCol-1]).trim()===String(p.id)) { row=i+1; break; }
  }
  if(row<0){
    row=sh.getLastRow()+1;
    sh.getRange(row, idCol).setValue(p.id);
  }

  HEADERS_PRODUCTOS.forEach(h=>{
    if(!map[h]) return;
    const value = h==='actualizado' ? new Date() : (p[h] !== undefined ? p[h] : '');
    sh.getRange(row, map[h]).setValue(value);
  });
  return {ok:true, producto:p};
}

function deleteProduct(id){
  setup(); if(!id) return {ok:false,error:'Falta ID'};
  const sh=ss().getSheetByName(SHEET_PRODUCTOS); const map=headerMap(sh); const values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(String(values[i][map.id-1]).trim()===String(id).trim()){
      sh.deleteRow(i+1);
      return {ok:true,id:id};
    }
  }
  return {ok:false,error:'ID no encontrado'};
}

function uploadBase64(base64, name){
  const folder=getFolder();
  const m=base64.match(/^data:(.*?);base64,(.*)$/); if(!m) return '';
  const mime=m[1]; const bytes=Utilities.base64Decode(m[2]);
  const ext=mime.includes('png')?'.png':'.jpg';
  const blob=Utilities.newBlob(bytes,mime,name+'_'+Date.now()+ext);
  const file=folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w1200';
}
function getFolder(){
  const it=DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return it.hasNext()?it.next():DriveApp.createFolder(DRIVE_FOLDER_NAME);
}
function saveOrder(data){
  setup(); const sh=ss().getSheetByName(SHEET_PEDIDOS);
  const items = data.items || [];
  const total = data.total || items.reduce((a,b)=>a+(Number(b.precio||0)*Number(b.cantidad||1)),0);
  sh.appendRow([new Date(), data.cliente||'', JSON.stringify(items), total]);
  if(data.descontarStock === true) items.forEach(item=>discountStock(item.id, Number(item.cantidad||1)));
  return {ok:true};
}
function discountStock(id, cantidad){
  setup(); if(!id) return {ok:false,error:'Falta ID'};
  const sh=ss().getSheetByName(SHEET_PRODUCTOS); const map=headerMap(sh); const values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(String(values[i][map.id-1]).trim()===String(id).trim()){
      const actual=Number(values[i][map.stock-1]||0);
      const nuevo=Math.max(0, actual-Number(cantidad||1));
      sh.getRange(i+1, map.stock).setValue(nuevo);
      if(nuevo<=0 && map.estado) sh.getRange(i+1, map.estado).setValue('Vendido');
      if(map.actualizado) sh.getRange(i+1, map.actualizado).setValue(new Date());
      return {ok:true,id:id,stock:nuevo};
    }
  }
  return {ok:false,error:'ID no encontrado'};
}

// NOTE: ordenar por actualizado en frontend o backend recomendado
