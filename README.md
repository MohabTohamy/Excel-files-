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

يتوفر برنامجان نصيان لإضافة صفوف الإرشادات:

Two scripts are available for adding guidance rows:

#### 1. البرنامج النصي الذكي (موصى به) / Intelligent Script (Recommended)

**`add_arabic_guidance.py`** - برنامج نصي محسّن يستخدم الكشف الذكي عن أنواع الأعمدة

**`add_arabic_guidance.py`** - Enhanced script with intelligent column type detection

```bash
# تثبيت المكتبات المطلوبة / Install required libraries
pip install openpyxl

# تشغيل البرنامج النصي / Run the script
python3 add_arabic_guidance.py
```

**المميزات / Features:**
- ✅ كشف تلقائي للأعمدة التي تحتوي بالفعل على إرشادات عربية / Auto-detects files with existing Arabic guidance
- ✅ مطابقة ذكية للأنماط لأنواع الأعمدة المختلفة / Intelligent pattern matching for different column types
- ✅ دعم شامل للعديد من أنواع الأعمدة (الجسور، السلامة المرورية، الإنارة، السرعة، إلخ) / Comprehensive support for many column types
- ✅ تنسيق صحيح (خط 9، مائل، رمادي) / Proper formatting (size 9, italic, gray)
- ✅ تقرير تفصيلي بالملفات المعالجة / Detailed report of processed files
- ✅ معالجة جميع ملفات .xlsx في المجلد / Processes all .xlsx files in the directory

**أنواع الأعمدة المدعومة / Supported Column Types:**
- المعرفات والرموز / IDs and Codes (Code, ID, SectionCode, BridgeID, etc.)
- التواريخ والأوقات / Dates and Times (Date, Time, SurveyDate, etc.)
- المواقع والإحداثيات / Locations and GPS (Location, GPS, Coordinates, Lat, Long)
- الشوارع والطرق / Streets and Roads (Street, Lane, Direction)
- القياسات والأبعاد / Measurements and Dimensions (Length, Width, Height, Area)
- الحالة والخطورة / Status and Severity (Status, Condition, Severity)
- الجسور / Bridges (Bridge Type, Span Length, Load Capacity)
- السلامة المرورية / Traffic Safety (Accident Type, Casualties, Weather)
- الإنارة / Lighting (Light Type, Power, Pole Material)
- السرعة والمركبات / Speed and Vehicles (Speed, Vehicle Type)
- التحسينات / Improvements (Improvement Type, Cost, Priority)
- القياسات الفنية / Technical Measurements (FWD, GPR, IRI, SKID)

#### 2. البرنامج النصي الأصلي / Original Script

**`add_guidance_rows.py`** - للملفات السبعة الأصلية فقط / For the original seven files only

```bash
python3 add_guidance_rows.py
```

**ملاحظة**: البرنامج النصي الجديد `add_arabic_guidance.py` يوفر كشف تلقائي أفضل ودعم أوسع لأنواع الأعمدة المختلفة.

**Note**: The new `add_arabic_guidance.py` script provides better auto-detection and broader support for different column types.

## المتطلبات التقنية / Technical Requirements

- Microsoft Excel 2010 أو أحدث / or newer
- LibreOffice Calc 5.0 أو أحدث / or newer
- Python 3.6+ و openpyxl (للصيانة فقط / for maintenance only)

## الترخيص / License

هذا المستودع متاح للاستخدام الحر.

This repository is available for free use.
