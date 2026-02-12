# قوالب Excel لإدارة بيانات الطرق
# Excel Templates for Road Data Management

هذا المستودع يحتوي على قوالب Excel لإدارة وتسجيل بيانات الطرق والعيوب.

This repository contains Excel templates for managing and recording road data and defects.

## الملفات / Files

### قوالب البيانات الفنية / Technical Data Templates

1. **قالب-fwd.xlsx** - قالب لتسجيل قراءات جهاز FWD (Falling Weight Deflectometer)
2. **قالب-gpr.xlsx** - قالب لتسجيل بيانات GPR (Ground Penetrating Radar)
3. **قالب-iri.xlsx** - قالب لتسجيل قيم IRI (International Roughness Index)
4. **قالب-skid.xlsx** - قالب لتسجيل بيانات مقاومة الانزلاق (Skid Resistance)

### قوالب العيوب / Defect Templates

5. **قالب-عيوب-التقاطعات.xlsx** - قالب لتسجيل عيوب التقاطعات
6. **قالب-عيوب-الطرق-الرئيسية.xlsx** - قالب لتسجيل عيوب الطرق الرئيسية
7. **قالب-عيوب-الطرق-الفرعية.xlsx** - قالب لتسجيل عيوب الطرق الفرعية

## الميزات / Features

✓ جميع القوالب تحتوي على صف إرشادي باللغة العربية تحت الرؤوس مباشرة
✓ الصف الإرشادي يوضح نوع البيانات المطلوب إدخالها في كل عمود
✓ تنسيق واضح ومميز للصف الإرشادي (خط صغير، مائل، خلفية رمادية فاتحة)

✓ All templates contain Arabic guidance rows directly under the headers
✓ Guidance rows explain what type of data should be entered in each column
✓ Clear and distinctive formatting for guidance rows (small font, italic, light gray background)

## كيفية الاستخدام / How to Use

1. اختر القالب المناسب لنوع البيانات التي تريد تسجيلها
2. افتح الملف في Microsoft Excel أو LibreOffice Calc
3. اقرأ الصف الإرشادي (الصف الثاني) لفهم نوع البيانات المطلوبة
4. ابدأ بإدخال بياناتك من الصف الثالث فما بعد

1. Choose the appropriate template for the type of data you want to record
2. Open the file in Microsoft Excel or LibreOffice Calc
3. Read the guidance row (row 2) to understand the required data types
4. Start entering your data from row 3 onwards

## البنية / Structure

كل ملف Excel يحتوي على:
- **الصف 1**: رؤوس الأعمدة (Headers)
- **الصف 2**: إرشادات باللغة العربية (Arabic Guidance)
- **الصف 3 وما بعد**: بيانات العينات والإدخالات (Sample Data and Entries)

Each Excel file contains:
- **Row 1**: Column Headers
- **Row 2**: Arabic Guidance
- **Row 3 onwards**: Sample Data and Entries

## صيانة القوالب / Template Maintenance

### إضافة صف إرشادي لقوالب جديدة / Adding Guidance Rows to New Templates

إذا كنت بحاجة لإضافة صف إرشادي لقوالب جديدة، استخدم البرنامج النصي المرفق:

If you need to add guidance rows to new templates, use the included script:

```bash
# تثبيت المكتبات المطلوبة / Install required libraries
pip install openpyxl

# تشغيل البرنامج النصي / Run the script
python3 add_guidance_rows.py
```

**ملاحظة**: يمكنك تعديل ملف `add_guidance_rows.py` لإضافة إرشادات مخصصة لقوالب جديدة.

**Note**: You can modify `add_guidance_rows.py` to add custom guidance for new templates.

## المتطلبات التقنية / Technical Requirements

- Microsoft Excel 2010 أو أحدث / or newer
- LibreOffice Calc 5.0 أو أحدث / or newer
- Python 3.6+ و openpyxl (للصيانة فقط / for maintenance only)

## الترخيص / License

هذا المستودع متاح للاستخدام الحر.

This repository is available for free use.
