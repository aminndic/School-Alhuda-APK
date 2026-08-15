# نظام مدرسة الهدى الثانوية بنين - Android APK

هذا المشروع يحوّل تطبيق Flask الأصلي إلى تطبيق Android يعمل محلياً باستخدام Chaquopy + WebView + SQLite.

## البناء عبر GitHub Actions

1. ارفع كامل محتويات هذا المشروع إلى مستودع GitHub.
2. افتح Actions.
3. اختر `Build School Alhuda APK`.
4. اضغط `Run workflow`.
5. بعد النجاح افتح الـ workflow ثم Artifact باسم `school-alhuda-debug-apk`.

## ملاحظات

- Python 3.11 وChaquopy 17.0.
- Android minSdk 24.
- التطبيق يشغل Flask على `127.0.0.1:5000` داخل الهاتف.
- قاعدة البيانات تحفظ في مساحة التطبيق الداخلية، لذلك لا تحتاج إنترنت.
- ملف Amiri-Regular.ttf يمكن وضعه لاحقاً داخل المشروع/التطبيق لتفعيل إخراج PDF العربي بشكل أفضل.
