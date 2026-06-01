# 📱 متابعة المخدومات (Sunday School Attendance App)

تطبيق ويب لتتبع حضور المخدومات في مدارس الأحد.

## 🚀 التشغيل

التطبيق يعمل مباشرة كـ Static Site. يمكنك:

1. **فتح الملف محلياً:**
   ```bash
   # باستخدام أي خادم محلي
   npx serve .
   # أو
   python -m http.server 8080
   ```

2. **استضافة على GitHub Pages:**
   - ارفع الملفات على مستودع GitHub
   - فعّل GitHub Pages من الإعدادات
   - اختر الفرع `main` والمجلد `root`

3. **استضافة على Netlify/Vercel:**
   - اسحب المجلد مباشرة على الموقع

## 📁 هيكل الملفات

```
.
├── index.html              # الصفحة الرئيسية
├── assets/
│   ├── index-DhHX_UUi.css  # ملفات الأنماط
│   └── index-BaoUA8bS.js   # ملف JavaScript المُجمّع
└── README.md               # هذا الملف
```

## 🛠️ التقنيات المستخدمة

- React (مُجمّع)
- CSS مخصص
- Firebase Auth (لتسجيل الدخول)
- PWA Ready (يدعم التثبيت كتطبيق)

## 📝 ملاحظات

- التطبيق يستخدم Firebase للمصادقة (تسجيل الدخول بـ Google)
- البيانات تُخزن محلياً (LocalStorage / IndexedDB)
- يدعم العمل بدون إنترنت (Offline Mode)

## 📄 الترخيص

جميع الحقوق محفوظة.
