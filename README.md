# Tebyan
A program to modify Arabic texts to display them correctly in applications and games that do not support the Arabic language.

[Download Tebyan](https://github.com/MrGamesKingPro/Tebyan/releases/tag/Tebyan)

Or use a version Python 

Install Libraries

pip install tk ttkbootstrap arabic_reshaper python-bidi

برنامج "تبيان" مصمم لمعالجة النصوص العربية لجعلها قابلة للعرض بشكل صحيح في التطبيقات والألعاب التي قد لا تدعم اللغة العربية بشكل مباشر (مثل عدم عرض الحروف المتصلة أو عكس اتجاه النص). يقوم البرنامج بتحويل النص العربي الأصلي إلى صيغة "مرئية" تستخدم أشكال الحروف المنفصلة والمرتبة بشكل صحيح للعرض في البيئات غير الداعم

يحتوي البرنامج على واجهتين رئيسيتين يمكن التنقل بينهما باستخدام علامات التبويب:

1.  **معالجة نصوص:** لمعالجة نص تقوم بكتابته مباشرة في البرنامج.

![01](https://github.com/user-attachments/assets/93ae4e61-769a-4958-9a01-8dc1208fad63)


2.  **معالجة ملفات:** لمعالجة النصوص العربية الموجودة داخل ملفات نصية (txt, .csv, .json. , .xml).

![02](https://github.com/user-attachments/assets/ec794cf7-0f30-4105-8538-aa33e468bf7b)

كيفية الاستخدام:

1.  **معالجة نصوص**
    *   .اكتب أو الصق النص العربي في المربع العلوي.
    *   .اضغط على زر 'معالجة النص'
    *   .سيظهر النص المعدل في المربع السفلي، جاهزاً للنسخ.
    *   .يمكنك نسخ النص المعدل بالكامل باستخدام زر 'نسخ الكل' الموجود أسفله.
    *   .يمكنك مسح مربعات النص الإدخال والإخراج باستخدام زر 'مسح الكل (النصوص)'
    *   .يمكنك حفظ النص المعدل من قائمة 'ملف'

2.  **معالجة ملفات**
    *   .(.txt)، XML (.xml)، JSON (.json)، أو CSV (.csv) اضغط على 'فتح ملفًا أو أكثر' لتحديد ملفات نصية ذات إمتداد
    *   .اضغط على 'مسار المجلد الحفظ' لتحديد المجلد الذي سيتم حفظ الملفات المعالجة فيه
    *   .سيتم إنشاء مجلد فرعي باسم 'processed_files' داخل المجلد المختار (إذا لم يكن موجوداً)
    *   .اضغط على 'معالجة الملفات المحددة'.
    *   .سيتم معالجة النصوص العربية فقط داخل الملفات وحفظ نسخ جديدة بنفس إسم والتنسيق في المجلد المخصص
    *   .تابع شريط التقدم وسجل الحالة لمعرفة حالة العملية.
    *   .يمكنك مسح قائمة الملفات ومجلد الحفظ والسجل باستخدام زر 'مسح الكل (الملفات)'

ملاحظات:
*    .UTF-8 تأكد من أن الملفات تستخدم ترميز
*   .في ملفات CSV و xml و JSON، سيتم محاولة معالجة كل القيم النصية
*   .يمكنك استخدام زر الفأرة الأيمن للقص/النسخ/اللصق في مربعات النص (حسب ما إذا كان المربع قابلاً للكتابة أو للقراءة فقط)
