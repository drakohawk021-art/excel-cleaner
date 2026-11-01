import pandas as pd
import re
from datetime import datetime
# 1) قراءة الملف (كل حاجة كنص عشان نتحكم في التنضيف)
input_file = input("اكتب اسم الملف (مثال: dirty_clients.xlsx): ").strip()
df = pd.read_excel(input_file, dtype=str)

# توحيد أسماء الأعمدة المختلفة
aliases = {
    "الموبايل": "التليفون",
    "الهاتف": "التليفون",
    "رقم التليفون": "التليفون",
    "رقم الموبايل": "التليفون",
    "السعر": "المبلغ",
    "القيمة": "المبلغ",
    "المبلغ المدفوع": "المبلغ",
    "تاريخ": "التاريخ",
    "الإسم": "الاسم",
    "اسم العميل": "الاسم",
}
df.columns = [aliases.get(col.strip(), col.strip()) for col in df.columns]
original_count = len(df)
print(f"عدد الصفوف الأصلي: {original_count}")

# 2) حذف الصفوف اللي مفيهاش اسم (ونحفظها في ملف منفصل)
missing_name = df["الاسم"].isna()
df[missing_name].to_excel("removed_rows.xlsx", index=False)
df = df[~missing_name].copy()
missing_count = int(missing_name.sum())
print(f"اتحذف {missing_count} صف ناقص الاسم")

# 3) تنضيف الأسماء: شيل المسافات الزيادة
df["الاسم"] = df["الاسم"].str.strip().str.replace(r"\s+", " ", regex=True)

# 4) توحيد التليفون: 010-1234-5678
def clean_phone(phone):
    digits = re.sub(r"\D", "", str(phone))
    if len(digits) == 11:
        return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    return phone  # لو الرقم غريب نسيبه زي ما هو

df["التليفون"] = df["التليفون"].apply(clean_phone)

# 5) توحيد التاريخ: 2024-01-15
def clean_date(value):
    if pd.isna(value):
        return None
    value = str(value).strip()
    formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d %B %Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return f"غير معروف: {value}"

df["التاريخ"] = df["التاريخ"].apply(clean_date)

# 6) تنضيف المبلغ: شيل أي حاجة مش رقم
df["المبلغ"] = df["المبلغ"].str.replace(r"[^\d.]", "", regex=True).astype(float)

# 7) شيل المكرر
before = len(df)
df = df.drop_duplicates()
duplicates_count = before - len(df)
print(f"اتشال {duplicates_count} صف مكرر")

# 8) حفظ النتيجة
df.to_excel("clean_clients.xlsx", index=False)
print(f"تم! عدد الصفوف النهائي: {len(df)}")

# 9) كتابة التقرير
report = f"""تقرير تنضيف الملف: {input_file}
-----------------------------
عدد الصفوف الأصلي: {original_count}
صفوف ناقصة الاسم (اتحذفت): {missing_count}
صفوف مكررة (اتشالت): {duplicates_count}
عدد الصفوف النهائي: {len(df)}

تم توحيد: المسافات، صيغة التليفون، صيغة التاريخ، المبالغ.
ملاحظة: التواريخ بصيغة يوم/شهر/سنة، والصفوف المحذوفة محفوظة في removed_rows.xlsx
"""
with open("report.txt", "w", encoding="utf-8") as file:
    file.write(report)
print("تم إنشاء report.txt")