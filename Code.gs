/**
 * PARTE DE TRABAJO — Google Apps Script v5
 * Листы: PARTE DE TRABAJO + servicio
 */
var SHEET_NAME    = "PARTE DE TRABAJO";
var SERVICE_SHEET = "servicio";
var PRICE_PER_HOUR = 28;

function out(data, callback) {
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p  = (e && e.parameter) ? e.parameter : {};
  var cb = p.callback || null;
  try {
    if (p.action === "getLists")    return out(getListsData(), cb);
    if (p.action === "getRowCount") return out(getRowCountData(), cb);
    return out({ status:"ok", message:"Parte de Trabajo API v5" }, cb);
  } catch(err) {
    return out({ status:"error", message:err.message }, cb);
  }
}

function doPost(e) {
  try {
    var raw = "";
    if (e.parameter && e.parameter.data) {
      raw = e.parameter.data;
    } else if (e.postData) {
      var body = e.postData.contents || "";
      var pairs = body.split("&");
      for (var i=0;i<pairs.length;i++) {
        var kv = pairs[i].split("=");
        if (kv.length >= 2 && decodeURIComponent(kv[0]) === "data") {
          raw = decodeURIComponent(kv.slice(1).join("=").replace(/\+/g," "));
          break;
        }
      }
      if (!raw) raw = body;
    }
    if (!raw) throw new Error("No data received");

    var d = JSON.parse(raw);
    Logger.log("POST data: " + JSON.stringify(d).substring(0,200));

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet not found: " + SHEET_NAME);

    var horas  = parseInt(d.horas)    || 0;
    var precio = parseFloat(d.precio) || 0;
    var total  = horas * PRICE_PER_HOUR + precio;

    sheet.appendRow([
      d.fecha||"", d.cliente||"", d.obra||"", d.operario||"",
      d.ayudante||"", d.vehiculo||"", d.trabajos||"", d.photos_obra||"",
      d.tipo||"", horas, d.compania||"", d.materiales||"",
      d.photos_mat||"", precio, total
    ]);
    var row = sheet.getLastRow();
    formatRow(sheet, row);
    Logger.log("✅ Row added: " + row);
    return out({ status:"ok", row:row, total:total });
  } catch(err) {
    Logger.log("❌ " + err.message);
    return out({ status:"error", message:err.message });
  }
}

function getListsData() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SERVICE_SHEET);
  if (!sheet) throw new Error("Sheet 'servicio' not found");
  var rows = sheet.getDataRange().getValues();
  var cl=[],ob=[],op=[],ti=[],co=[];
  for (var i=1;i<rows.length;i++) {
    var r=rows[i];
    pu(cl,r[0]); pu(ob,r[1]); pu(op,r[2]); pu(op,r[3]); pu(ti,r[4]); pu(co,r[5]);
  }
  return {status:"ok",data:{clientes:cl,obras:ob,operarios:op,tipos:ti,companias:co}};
}

function getRowCountData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return {status:"error",count:0};
  return {status:"ok",count:sheet.getLastRow()};
}

function pu(arr,val) {
  if (!val) return;
  var v=String(val).trim();
  if (v && arr.indexOf(v)===-1) arr.push(v);
}

function formatRow(sheet,r) {
  try {
    sheet.getRange(r,1).setNumberFormat("dd/mm/yyyy");
    sheet.getRange(r,10).setNumberFormat("0");
    sheet.getRange(r,14).setNumberFormat("#,##0.00");
    sheet.getRange(r,15).setNumberFormat("#,##0.00");
    if (r%2===0) sheet.getRange(r,1,1,15).setBackground("#F8F9FA");
  } catch(e) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("🌐 Traducir / Перевод")
    .addItem("🇪🇸 Español","translateToSpanish")
    .addItem("🇷🇺 Русский","translateToRussian")
    .addItem("↩ Original","restoreOriginal")
    .addToUi();
}

var H = {
  original:["FECHA","CLIENTE","OBRA ","OPERARIO","AYUDANTE","VEHÍCULO","TRABAJOS REALIZADOS","FOTOS DE OBRA","TIPO DE TRABAJO","HORAS (28 €/h)","COMPANIA ","MATERIALES UTILIZADOS","Foto de los materiales comprados o del recibo","Precio (EUR)","TOTAL"],
  es:["FECHA","CLIENTE","OBRA","OPERARIO","AYUDANTE","VEHÍCULO","TRABAJOS REALIZADOS","FOTOS DE OBRA","TIPO DE TRABAJO","HORAS","COMPAÑÍA","MATERIALES UTILIZADOS","FOTO MATERIALES / RECIBO","PRECIO (EUR)","TOTAL"],
  ru:["ДАТА","КЛИЕНТ","ОБЪЕКТ","ИСПОЛНИТЕЛЬ","ПОМОЩНИК","ТРАНСПОРТ","ВЫПОЛНЕННЫЕ РАБОТЫ","ФОТО ОБЪЕКТА","ТИП РАБОТЫ","ЧАСЫ","КОМПАНИЯ","ИСПОЛЬЗОВАННЫЕ МАТЕРИАЛЫ","ФОТО МАТЕРИАЛОВ / ЧЕК","ЦЕНА (EUR)","ИТОГО"]
};
function sh(a){var s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);if(s)s.getRange(1,1,1,a.length).setValues([a]);}
function translateToSpanish(){sh(H.es);SpreadsheetApp.getUi().alert("✅ Español");}
function translateToRussian(){sh(H.ru);SpreadsheetApp.getUi().alert("✅ Русский");}
function restoreOriginal(){sh(H.original);SpreadsheetApp.getUi().alert("✅ Original");}

function testSetup() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet: "+ss.getName());
  var m=ss.getSheetByName(SHEET_NAME);
  Logger.log(m?"✅ PARTE DE TRABAJO rows:"+m.getLastRow():"❌ Not found");
  var s=ss.getSheetByName(SERVICE_SHEET);
  Logger.log(s?"✅ servicio rows:"+s.getLastRow():"❌ Not found");
  if(s) Logger.log("getLists: "+JSON.stringify(getListsData()).substring(0,300));
}

function testPost() {
  var e={parameter:{data:JSON.stringify({fecha:"2026-06-26",cliente:"CASA FUSTER",obra:"ORIGINAL",
    operario:"MOHA",ayudante:"",vehiculo:"TEST",trabajos:"Test v5",
    tipo:"Construcción / Строительные работы",horas:8,compania:"OBRAMAT",
    materiales:"Test",precio:0,photos_obra:"",photos_mat:""})},postData:null};
  Logger.log("testPost: "+doPost(e).getContent());
}
