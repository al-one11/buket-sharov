const PRODUCTS = [
  {id:'bouquet-1', category:'Букеты', name:'Букет цветов', description:'Готовый букет для особенного момента', price:'', oldPrice:'', visible:true},
  {id:'flowers-1', category:'Цветы', name:'Свежие цветы', description:'Цветы для букета или композиции', price:'', oldPrice:'', visible:true},
  {id:'balloons-1', category:'Шары', name:'Воздушные шары', description:'Яркое дополнение к подарку', price:'', oldPrice:'', visible:true},
  {id:'gift-1', category:'Подарки', name:'Подарок', description:'Красивое дополнение к цветам', price:'', oldPrice:'', visible:true}
];

// 1) Создайте Google Таблицу.
// 2) Откройте Расширения → Apps Script и вставьте этот код.
// 3) Разверните как веб-приложение: выполнять от вашего аккаунта, доступ — Все.
// 4) Скопируйте URL веб-приложения в CONFIG.API_URL в admin.html.
// 5) Смените пароль ниже на свой.
const OWNER_PASSWORD = 'CHANGE_ME_1234';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Товары') || ss.insertSheet('Товары');
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,7).setValues([['id','category','name','description','price','oldPrice','visible']]);
    sh.getRange(2,1,PRODUCTS.length,7).setValues(PRODUCTS.map(p => [p.id,p.category,p.name,p.description,p.price,p.oldPrice,p.visible]));
  }
  return sh;
}

function doGet() {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  const products = values.slice(1).filter(r => r[0]).map(r => ({id:String(r[0]),category:String(r[1]||''),name:String(r[2]||''),description:String(r[3]||''),price:String(r[4]||''),oldPrice:String(r[5]||''),visible:r[6] !== false && String(r[6]).toLowerCase() !== 'false'}));
  return json_({ok:true, products:products});
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.password !== OWNER_PASSWORD) return json_({ok:false,error:'Неверный пароль'});
    if (!Array.isArray(body.products)) return json_({ok:false,error:'Некорректные данные'});
    const sh = sheet_();
    sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),7).clearContent();
    const rows = body.products.map(p => [p.id||'',p.category||'',p.name||'',p.description||'',p.price||'',p.oldPrice||'',p.visible !== false]);
    if (rows.length) sh.getRange(2,1,rows.length,7).setValues(rows);
    return json_({ok:true,updatedAt:new Date().toISOString()});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
