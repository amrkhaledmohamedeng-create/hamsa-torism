# همسه — نموذج الاستفسار + Google Sheets

## 1) إنشاء الشيت
أنشئ Google Sheet جديد وافتح Extensions → Apps Script. الصق محتوى Code.gs، ثم شغّل setup مرة واحدة. سيتم إنشاء ورقة Leads تلقائيًا.

## 2) نشر الباك إند
Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone. انسخ رابط /exec.

## 3) ربط الموقع
في index.html ابحث عن:
YOUR_APPS_SCRIPT_WEB_APP_URL
واستبدله برابط /exec الذي نسخته.

## 4) التشغيل
ارفع index.html على GitHub Pages. النموذج أصبح داخل نفس الصفحة، والطلبات تُضاف مباشرة إلى ورقة Leads.

الأعمدة: Timestamp, Name, Phone, Governorate, Service, Program, Persons, TravelDate, PaymentMethod, DownPayment, InstallmentMonths, Notes, Status.
