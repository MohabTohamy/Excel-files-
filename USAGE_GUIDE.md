# دليل الاستخدام - إضافة الإرشادات العربية
# Usage Guide - Adding Arabic Guidance Rows

## نظرة عامة / Overview

هذا الدليل يشرح كيفية استخدام البرنامج النصي `add_arabic_guidance.py` لإضافة صفوف إرشادية باللغة العربية تلقائيًا إلى ملفات Excel.

This guide explains how to use the `add_arabic_guidance.py` script to automatically add Arabic guidance rows to Excel files.

## المتطلبات / Requirements

```bash
# تثبيت المكتبات المطلوبة / Install required libraries
pip install openpyxl
```

## الاستخدام الأساسي / Basic Usage

```bash
# انتقل إلى مجلد المستودع / Navigate to repository folder
cd /path/to/repository

# تشغيل البرنامج النصي / Run the script
python3 add_arabic_guidance.py
```

البرنامج النصي سيقوم تلقائيًا بـ:
1. البحث عن جميع ملفات .xlsx في المجلد الحالي
2. فحص كل ملف للتحقق من وجود إرشادات عربية في الصف 2
3. إضافة صف إرشادي للملفات التي لا تحتوي على إرشادات
4. طباعة تقرير تفصيلي بالنتائج

The script will automatically:
1. Find all .xlsx files in the current directory
2. Check each file for existing Arabic guidance in row 2
3. Add a guidance row to files without guidance
4. Print a detailed summary report

## مثال على المخرجات / Example Output

```
============================================================
Adding Arabic Guidance Rows to Excel Files
إضافة صفوف الإرشادات العربية لملفات Excel
============================================================

Found 10 Excel file(s)
Files: template-bridges.xlsx, template-safety.xlsx, ...


============================================================
Processing: template-bridges.xlsx
============================================================
Headers found: 8 columns
Inserted new row at position 2
Added guidance for 8 columns
✓ Successfully saved template-bridges.xlsx

============================================================
Processing: قالب-fwd.xlsx
============================================================
⊘ Skipped: قالب-fwd.xlsx already has Arabic guidance in row 2

============================================================
SUMMARY REPORT / تقرير ملخص
============================================================

✓ Successfully processed: 1 file(s)
  - SUCCESS: template-bridges.xlsx (8 columns)

⊘ Skipped (already have guidance): 1 file(s)
  - SKIPPED: قالب-fwd.xlsx (already has guidance)

============================================================
Processing complete! / اكتملت المعالجة!
============================================================
```

## أمثلة على الإرشادات لأنواع مختلفة من الأعمدة / Examples of Guidance for Different Column Types

### ملفات الجسور / Bridge Files
- **Bridge ID** → "أدخل رمز تعريف الجسر"
- **Bridge Type** → "أدخل نوع الجسر (خرساني، معدني، إلخ)"
- **Span Length** → "أدخل طول البحر/الفتحة بالمتر"
- **Condition** → "أدخل حالة الجسر (ممتاز، جيد، متوسط، سيء)"
- **Load Capacity** → "أدخل الحمولة القصوى بالطن"

### ملفات السلامة المرورية / Traffic Safety Files
- **Date** → "أدخل التاريخ بصيغة DD/MM/YYYY"
- **Time** → "أدخل الوقت بصيغة HH:MM"
- **Accident Type** → "أدخل نوع الحادث (اصطدام، انقلاب، دهس، إلخ)"
- **Severity** → "أدخل درجة الخطورة (Low=منخفض، Medium=متوسط، High=عالي)"
- **Casualties** → "أدخل عدد الإصابات أو الوفيات"
- **Weather** → "أدخل حالة الطقس (صحو، ممطر، ضبابي، إلخ)"

### ملفات الإنارة / Lighting Files
- **Light Type** → "أدخل نوع الإنارة (LED، صوديوم، هاليد معدني، إلخ)"
- **Power** → "أدخل القدرة بالواط (W)"
- **Height** → "أدخل الارتفاع بالمتر"
- **Pole Material** → "أدخل مادة العمود (حديد، ألمنيوم، خرسانة)"
- **Working Status** → "أدخل حالة التشغيل (يعمل، معطل، يحتاج صيانة)"

### ملفات السرعة / Speed Files
- **Speed** → "أدخل السرعة بالكيلومتر/ساعة (km/h)"
- **Vehicle Type** → "أدخل نوع المركبة (سيارة خاصة، شاحنة، حافلة، دراجة)"
- **Lane** → "أدخل رقم المسار (L1, L2, إلخ)"
- **Direction** → "أدخل اتجاه الحركة (شمال، جنوب، شرق، غرب)"

### الأعمدة العامة / General Columns
- **Code/ID** → "أدخل رمز/معرف فريد (مثال: 001، ABC-123)"
- **Date** → "أدخل التاريخ بصيغة DD/MM/YYYY"
- **Location** → "أدخل الموقع أو الإحداثيات"
- **GPS Coordinates** → "أدخل خط الطول والعرض (Lat, Long)"
- **Description** → "أدخل وصف تفصيلي"
- **Notes** → "أدخل أي ملاحظات إضافية"
- **Status** → "أدخل الحالة (جيد، متوسط، سيء، إلخ)"

## التنسيق المطبق / Applied Formatting

جميع صفوف الإرشادات تحصل على التنسيق التالي:
- **حجم الخط / Font Size:** 9
- **النمط / Style:** مائل / Italic
- **لون النص / Text Color:** رمادي داكن / Dark Gray (#404040)
- **لون الخلفية / Background Color:** رمادي فاتح / Light Gray (#D3D3D3)
- **المحاذاة / Alignment:** توسيط أفقي وعمودي / Center horizontal and vertical
- **التفاف النص / Text Wrap:** مفعّل / Enabled
- **ارتفاع الصف / Row Height:** 30

## الكشف الذكي عن الأعمدة / Intelligent Column Detection

البرنامج النصي يستخدم مطابقة الأنماط الذكية لتحديد نوع كل عمود:

The script uses intelligent pattern matching to determine column types:

1. **البحث عن الكلمات المفتاحية / Keyword Search**
   - يبحث عن كلمات مثل "code", "date", "location", "bridge", "speed", إلخ
   - Searches for keywords like "code", "date", "location", "bridge", "speed", etc.

2. **دعم اللغتين / Bilingual Support**
   - يدعم أسماء الأعمدة بالإنجليزية والعربية
   - Supports both English and Arabic column names

3. **إرشادات افتراضية / Default Guidance**
   - إذا لم يتم التعرف على نوع العمود، يستخدم "أدخل البيانات المطلوبة"
   - If column type is not recognized, uses "أدخل البيانات المطلوبة"

## كيف يعمل الكشف التلقائي / How Auto-Detection Works

```python
# مثال 1: الكشف عن أعمدة التاريخ
# Example 1: Date column detection
"Date" → "أدخل التاريخ بصيغة DD/MM/YYYY"
"SurveyDate" → "أدخل تاريخ المسح بصيغة DD/MM/YYYY"
"تاريخ" → "أدخل التاريخ بصيغة DD/MM/YYYY"

# مثال 2: الكشف عن أعمدة الرموز
# Example 2: Code column detection
"BridgeID" → "أدخل رمز تعريف الجسر"
"SectionCode" → "أدخل رمز القطاع/المقطع"
"Code" → "أدخل رمز/معرف فريد (مثال: 001، ABC-123)"

# مثال 3: الكشف عن أعمدة القياسات
# Example 3: Measurement column detection
"Length" → "أدخل الطول بالمتر أو السنتيمتر"
"Width" → "أدخل العرض بالمتر أو السنتيمتر"
"Area" → "أدخل المساحة بالمتر المربع"
```

## الأسئلة الشائعة / FAQ

### 1. ماذا يحدث إذا كان الصف الثاني يحتوي بالفعل على بيانات؟
**What happens if row 2 already contains data?**

البرنامج النصي يتحقق من وجود نص عربي في الصف 2. إذا وُجد، يتم تخطي الملف.

The script checks for Arabic text in row 2. If found, the file is skipped.

### 2. هل يمكنني تشغيل البرنامج النصي عدة مرات؟
**Can I run the script multiple times?**

نعم! البرنامج النصي آمن للتشغيل المتكرر. سيتخطى الملفات التي تحتوي بالفعل على إرشادات.

Yes! The script is safe to run multiple times. It will skip files that already have guidance.

### 3. كيف يمكنني إضافة إرشادات مخصصة لنوع عمود جديد؟
**How can I add custom guidance for a new column type?**

قم بتعديل قاموس `guidance_patterns` في دالة `get_guidance_for_column()` في ملف `add_arabic_guidance.py`.

Edit the `guidance_patterns` dictionary in the `get_guidance_for_column()` function in `add_arabic_guidance.py`.

### 4. هل يحفظ البرنامج النصي نسخة احتياطية من الملفات الأصلية؟
**Does the script save a backup of the original files?**

لا، البرنامج النصي يستبدل الملفات مباشرة. يُنصح بإنشاء نسخة احتياطية يدويًا قبل التشغيل أو استخدام نظام التحكم في الإصدارات (Git).

No, the script replaces files directly. It's recommended to create a manual backup or use version control (Git) before running.

## استكشاف الأخطاء / Troubleshooting

### خطأ: ModuleNotFoundError: No module named 'openpyxl'
```bash
pip install openpyxl
```

### خطأ: PermissionError
تأكد من أن ملفات Excel غير مفتوحة في برنامج آخر.

Make sure Excel files are not open in another program.

### البرنامج النصي لا يجد أي ملفات
تأكد من تشغيل البرنامج النصي في نفس المجلد الذي يحتوي على ملفات .xlsx.

Make sure to run the script in the same directory that contains .xlsx files.

## الدعم / Support

إذا واجهت أي مشاكل أو لديك اقتراحات، يرجى فتح issue في المستودع.

If you encounter any issues or have suggestions, please open an issue in the repository.
