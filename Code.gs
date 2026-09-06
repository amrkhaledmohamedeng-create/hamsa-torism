/**
 * Hamsa - Leads backend for Google Sheets
 * Attach this script to the Google Sheet that will store customer requests.
 */
const SHEET_NAME = 'Leads';
const HEADERS = [
  'Timestamp','Name','Phone','Governorate','Service','Program',
  'Persons','TravelDate','PaymentMethod','DownPayment','InstallmentMonths',
  'Notes','Status'
];

function setup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) sh = ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow() === 0) sh.appendRow(HEADERS);
  else if(sh.getRange(1,1,1,HEADERS.length).getValues()[0].every(v=>v==='')) sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
}

function doGet(){
  return ContentService.createTextOutput(JSON.stringify({ok:true,service:'hamsa-leads'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    setup();
    const p = e && e.parameter ? e.parameter : {};
    // Honeypot: bots that fill this hidden field are ignored.
    if((p.website || '').trim() !== '') return json({ok:true});
    const name = String(p.name || '').trim();
    const phone = String(p.phone || '').trim();
    const service = String(p.service || '').trim();
    if(!name || !phone || !service) return json({ok:false,error:'البيانات الأساسية ناقصة'});
    if(!/^01\d{9}$/.test(phone.replace(/[\s-]/g,''))) return json({ok:false,error:'رقم الموبايل غير صحيح'});
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    sh.appendRow([
      new Date(),
      name,
      phone,
      String(p.governorate || '').trim(),
      service,
      String(p.program || '').trim(),
      String(p.persons || '1').trim(),
      String(p.travelDate || '').trim(),
      String(p.paymentMethod || '').trim(),
      String(p.downPayment || '').trim(),
      String(p.installmentMonths || '').trim(),
      String(p.notes || '').trim(),
      'جديد'
    ]);
    return json({ok:true,message:'تم تسجيل الطلب'});
  }catch(err){
    return json({ok:false,error:String(err)});
  }
}

function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
