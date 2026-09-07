/**
 * Hamsa - Leads + Offers backend for Google Sheets
 * ================================================
 * هذا الكود فيه جزءان:
 * 1) نظام الطلبات (Leads) - زي القديم بالظبط.
 * 2) نظام العروض (Offers) - جديد، بيخليك تتحكم في عروض الموقع
 *    عن طريق شيت اسمه "Offers" بدل ما تعدل كود الموقع.
 *
 * طريقة الاستخدام:
 * - افتح Extensions > Apps Script في نفس الشيت بتاعك.
 * - امسح الكود القديم كله والصقه هذا بدلاً منه.
 * - شغّل الدالة setup() مرة واحدة يدويًا (من القائمة أعلى المحرر)
 *   عشان تتنشئ شيتات Leads و Offers ويتحط فيها أعمدة تلقائيًا.
 * - اعمل Deploy > New deployment > Web app (Execute as: Me, Access: Anyone)
 *   وخد الرابط ده، هو نفسه اللي هيتحط في نموذج الحجز وفي كود جلب العروض.
 */

const SHEET_NAME = 'Leads';
const HEADERS = [
  'Timestamp','Name','Phone','Governorate','Service','Program',
  'Persons','TravelDate','PaymentMethod','DownPayment','InstallmentMonths',
  'Notes','Status'
];

// ================== إعداد شيت العروض (Offers) ==================
const OFFERS_SHEET_NAME = 'Offers';
const OFFERS_HEADERS = [
  'ID','Section','Title','Subtitle','Items','Price','PriceLabel',
  'Badge','Image','Order','Active'
];
/*
  شرح أعمدة شيت Offers:
  - ID          : رقم أو كود مميز للعرض (اختياري، للترتيب الداخلي فقط)
  - Section     : اسم القسم اللي يظهر فيه العرض. القيم المسموحة:
                  umrah   = قسم برامج العمرة
                  hajj    = قسم برامج الحج
                  cash    = قسم عمرة الكاش
  - Title       : عنوان العرض (مثال: عمرة التيسير)
  - Subtitle    : وصف قصير تحت العنوان (اختياري)
  - Items       : تفاصيل العرض، كل نقطة تتفصل عن التانية بعلامة | (شرطة رأسية)
                  مثال: 1000 جنيه × 24 شهر|2000 جنيه × 12 شهر
  - Price       : السعر أو النص اللي يظهر في الشريط الملون تحت الكارت
  - PriceLabel  : يسيبه فاضي عادةً (احتياطي لأي تسمية إضافية)
  - Badge       : كلمة صغيرة زي "الأكثر طلبًا" (اختياري)
  - Image       : رابط صورة للكارت (اختياري - لو فاضي هياخد صورة افتراضية)
  - Order       : رقم يحدد ترتيب ظهور العرض (الأصغر يظهر أولاً)
  - Active      : اكتب "نعم" عشان يظهر العرض، أو "لا" عشان يختفي بدون ما تمسحه
*/

function setup(){
  setupLeads();
  setupOffers();
}

function setupLeads(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh) sh = ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow() === 0) sh.appendRow(HEADERS);
  else if(sh.getRange(1,1,1,HEADERS.length).getValues()[0].every(v=>v==='')) sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
}

function setupOffers(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  if(!sh) sh = ss.insertSheet(OFFERS_SHEET_NAME);
  if(sh.getLastRow() === 0){
    sh.appendRow(OFFERS_HEADERS);
    // صفوف نموذجية عشان تفهم الشكل المطلوب وتقدر تعدل عليها
    sh.appendRow([1,'umrah','عمرة التيسير','ابدأ من 1,000 جنيه',
      '1000 جنيه × 24 شهر|1250 جنيه × 24 شهر|1500 جنيه × 24 شهر|2000 جنيه × 12 شهر',
      'ابدأ من 1,000 جنيه','', '', '', 1, 'نعم']);
    sh.appendRow([2,'cash','ريع بخش - اقتصادي بالمواصلات','11 ليلة مكة + 3 ليالي المدينة',
      'رباعي 40,500|ثلاثي 43,500|ثنائي 48,900|سينجل 64,500',
      '40,500 جنيه','', '', '', 1, 'نعم']);
  } else if(sh.getRange(1,1,1,OFFERS_HEADERS.length).getValues()[0].every(v=>v==='')){
    sh.getRange(1,1,1,OFFERS_HEADERS.length).setValues([OFFERS_HEADERS]);
  }
  sh.setFrozenRows(1);
}

// ================== الردود على الطلبات (GET/POST) ==================

function doGet(e){
  setup();
  const p = e && e.parameter ? e.parameter : {};
  if(p.action === 'offers'){
    return getOffersJson(p.section || '');
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true,service:'hamsa-leads'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOffersJson(sectionFilter){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(OFFERS_SHEET_NAME);
  if(!sh) return json({ok:false,error:'لا يوجد شيت عروض بعد',offers:[]});

  const lastRow = sh.getLastRow();
  if(lastRow < 2) return json({ok:true,offers:[]});

  const values = sh.getRange(2,1,lastRow-1,OFFERS_HEADERS.length).getValues();
  let offers = values.map(function(row){
    return {
      id: row[0],
      section: String(row[1]||'').trim(),
      title: String(row[2]||'').trim(),
      subtitle: String(row[3]||'').trim(),
      items: String(row[4]||'').split('|').map(function(s){return s.trim();}).filter(Boolean),
      price: String(row[5]||'').trim(),
      priceLabel: String(row[6]||'').trim(),
      badge: String(row[7]||'').trim(),
      image: String(row[8]||'').trim(),
      order: Number(row[9]) || 999,
      active: String(row[10]||'').trim()
    };
  });

  // فلترة العروض المفعلة فقط
  offers = offers.filter(function(o){
    return o.title && (o.active === 'نعم' || o.active.toLowerCase() === 'yes' || o.active === 'true' || o.active === '1');
  });

  if(sectionFilter){
    offers = offers.filter(function(o){ return o.section === sectionFilter; });
  }

  offers.sort(function(a,b){ return a.order - b.order; });

  return json({ok:true,offers:offers});
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
