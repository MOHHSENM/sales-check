import glob
import os
import re
import pandas as pd


# ============================================================
# خواندن مشترک فایل‌های ورودی
# ============================================================
print("=" * 60)
print("خواندن فایل‌های ورودی...")
print("=" * 60)

forosh_file = "فروش منشور.xlsx"
df_forosh = pd.read_excel(forosh_file)
print(f"✓ '{forosh_file}' خوانده شد ({len(df_forosh)} ردیف)")

grid_files = sorted(glob.glob("GridData*.xlsx"))
if not grid_files:
    raise FileNotFoundError("هیچ فایل GridData ای در پوشه پیدا نشد!")
print(f"✓ {len(grid_files)} فایل GridData پیدا شد:")
for f in grid_files:
    print(f"   - {os.path.basename(f)}")

df_grid_raw = pd.concat([pd.read_excel(f) for f in grid_files], ignore_index=True)
print(f"   مجموع ردیف بعد از ادغام: {len(df_grid_raw)}")


# ============================================================
# گزارش ۱: مقایسه بر اساس کد پرسنلی → نتیجه.xlsx
# ============================================================
def report_personnel(df_forosh, df_grid_raw):
    print("\n" + "=" * 60)
    print("گزارش ۱: مقایسه بر اساس کد پرسنلی")
    print("=" * 60)

    # کپی مستقل تا تغییرات این گزارش به گزارش بعدی آسیب نزنه
    df_grid_final = df_grid_raw.copy()

    # استخراج ستون‌های مورد نیاز
    df_forosh_sub = df_forosh[["کدپرسنلی↕", "مدت", "مانده قسطی↕"]].copy()
    df_forosh_sub.columns = ["کد پرسنلی", "تعداد اقساط", "مانده_منشور"]

    df_forosh_sub["کد پرسنلی"] = pd.to_numeric(df_forosh_sub["کد پرسنلی"], errors="coerce")
    df_forosh_sub["تعداد اقساط"] = pd.to_numeric(df_forosh_sub["تعداد اقساط"], errors="coerce")
    df_forosh_sub["مانده_منشور"] = pd.to_numeric(
        df_forosh_sub["مانده_منشور"], errors="coerce"
    ).fillna(0)

    df_forosh_agg = (
        df_forosh_sub.groupby(["کد پرسنلی", "تعداد اقساط"], as_index=False)["مانده_منشور"].sum()
    )

    df_grid_final["مانده قابل پرداخت_clean"] = (
        df_grid_final["مانده قابل پرداخت"].astype(str).str.replace(",", "").str.replace("٬", "")
    )
    df_grid_final["مانده قابل پرداخت_clean"] = pd.to_numeric(
        df_grid_final["مانده قابل پرداخت_clean"], errors="coerce"
    ).fillna(0)

    df_grid_sub = df_grid_final[["کد پرسنلی", "تعداد اقساط", "مانده قابل پرداخت_clean"]].copy()
    df_grid_sub.columns = ["کد پرسنلی", "تعداد اقساط", "مانده_گرید"]
    df_grid_sub["کد پرسنلی"] = pd.to_numeric(df_grid_sub["کد پرسنلی"], errors="coerce")
    df_grid_sub["تعداد اقساط"] = pd.to_numeric(df_grid_sub["تعداد اقساط"], errors="coerce")

    df_grid_agg = (
        df_grid_sub.groupby(["کد پرسنلی", "تعداد اقساط"], as_index=False)["مانده_گرید"].sum()
    )

    df_merged = pd.merge(
        df_forosh_agg, df_grid_agg,
        on=["کد پرسنلی", "تعداد اقساط"], how="outer",
    ).fillna(0)

    df_merged["اختلاف"] = df_merged["مانده_منشور"] - df_merged["مانده_گرید"]

    def check_status(row):
        if row["مانده_منشور"] == 0:
            return "فقط در گرید دیتا"
        elif row["مانده_گرید"] == 0:
            return "فقط در منشور"
        elif row["اختلاف"] != 0:
            return "دارای مغایرت مبلغ"
        else:
            return "مطابق"

    df_merged["وضعیت"] = df_merged.apply(check_status, axis=1)

    df_result = df_merged[[
        "کد پرسنلی", "تعداد اقساط", "مانده_منشور", "مانده_گرید", "اختلاف", "وضعیت"
    ]]

    output_filename = "نتیجه.xlsx"
    df_result.to_excel(output_filename, index=False)
    print(f"✅ '{output_filename}' ساخته شد ({len(df_result)} ردیف)")


# ============================================================
# گزارش ۲: مقایسه بر اساس کد غرفه → نتیجه غرفه.xlsx
# ============================================================
def report_booth(df_forosh, df_grid_raw):
    print("\n" + "=" * 60)
    print("گزارش ۲: مقایسه بر اساس کد غرفه")
    print("=" * 60)

    df_grid = df_grid_raw.copy()

    def extract_code(val):
        if pd.isna(val):
            return None
        s = str(val).strip()
        m = re.search(r"_(\d+)\s*$", s)
        if m:
            return int(m.group(1))
        m2 = re.search(r"(\d+)\s*$", s)
        return int(m2.group(1)) if m2 else None

    df_grid["کد گرید"] = df_grid["کد فروشگاه"].apply(extract_code)
    df_grid["قیمت_clean"] = pd.to_numeric(
        df_grid["قیمت"].astype(str).str.replace(",", "").str.replace("٬", ""),
        errors="coerce",
    ).fillna(0)
    df_grid["تعداد اقساط"] = pd.to_numeric(df_grid["تعداد اقساط"], errors="coerce")
    df_grid = df_grid.dropna(subset=["کد گرید", "تعداد اقساط"])
    df_grid["کد گرید"] = df_grid["کد گرید"].astype(int)
    df_grid["تعداد اقساط"] = df_grid["تعداد اقساط"].astype(int)

    df_grid_step1 = (
        df_grid.groupby(["کد گرید", "تعداد اقساط"], as_index=False)["قیمت_clean"].sum()
    )

    print("   خواندن 'کد غرفه منشور.xlsx'...")
    df_map = pd.read_excel("کد غرفه منشور.xlsx")
    df_map = df_map.rename(columns={
        df_map.columns[1]: "کد گرید",
        df_map.columns[2]: "کد غرفه",
        df_map.columns[3]: "تعداد اقساط",
    })
    df_map = df_map[["کد گرید", "کد غرفه", "تعداد اقساط"]].copy()
    for c in ["کد گرید", "کد غرفه", "تعداد اقساط"]:
        df_map[c] = pd.to_numeric(df_map[c], errors="coerce")
    df_map = df_map.dropna()
    for c in ["کد گرید", "کد غرفه", "تعداد اقساط"]:
        df_map[c] = df_map[c].astype(int)
    df_map_unique = df_map.drop_duplicates(subset=["کد گرید", "تعداد اقساط"])

    df_mapped = pd.merge(df_grid_step1, df_map_unique, on=["کد گرید", "تعداد اقساط"], how="left")
    missing = df_mapped["کد غرفه"].isna().sum()
    print(f"   ردیف‌های بدون نگاشت: {missing}")

    df_with_code = df_mapped[df_mapped["کد غرفه"].notna()].copy()
    df_with_code["کد غرفه"] = df_with_code["کد غرفه"].astype(int)

    df_grid_booth_agg = (
        df_with_code.groupby("کد غرفه", as_index=False)
        .agg({
            "قیمت_clean": "sum",
            "کد گرید": lambda x: ", ".join(str(v) for v in sorted(set(x))),
        })
        .rename(columns={"قیمت_clean": "جمع قیمت گرید"})
    )
    print(f"   تعداد کد غرفه یکتا از گرید: {len(df_grid_booth_agg)}")

    df_no_code = df_mapped[df_mapped["کد غرفه"].isna()].copy()
    df_no_code_grouped = (
        df_no_code.groupby("کد گرید", as_index=False)["قیمت_clean"]
        .sum()
        .rename(columns={"قیمت_clean": "جمع قیمت گرید"})
    )
    df_no_code_grouped["کد غرفه"] = pd.NA
    df_no_code_grouped["کد گرید"] = df_no_code_grouped["کد گرید"].astype(str)

    forosh_code_col = df_forosh.columns[5]   # ستون F
    forosh_mande_col = df_forosh.columns[1]  # ستون B

    df_forosh_sub_b = df_forosh[[forosh_code_col, forosh_mande_col]].copy()
    df_forosh_sub_b.columns = ["کد غرفه", "مانده"]
    df_forosh_sub_b["کد غرفه"] = pd.to_numeric(df_forosh_sub_b["کد غرفه"], errors="coerce")
    df_forosh_sub_b["مانده"] = pd.to_numeric(
        df_forosh_sub_b["مانده"].astype(str).str.replace(",", ""), errors="coerce"
    ).fillna(0)
    df_forosh_sub_b = df_forosh_sub_b.dropna(subset=["کد غرفه"])
    df_forosh_sub_b["کد غرفه"] = df_forosh_sub_b["کد غرفه"].astype(int)

    df_forosh_agg_b = (
        df_forosh_sub_b.groupby("کد غرفه", as_index=False)["مانده"].sum()
        .rename(columns={"مانده": "جمع مانده منشور"})
    )

    df_final1 = pd.merge(df_grid_booth_agg, df_forosh_agg_b, on="کد غرفه", how="outer")
    df_final2 = df_no_code_grouped.copy()
    df_final2["جمع مانده منشور"] = 0

    df_final = pd.concat([df_final1, df_final2], ignore_index=True)
    df_final["جمع قیمت گرید"] = df_final["جمع قیمت گرید"].fillna(0)
    df_final["جمع مانده منشور"] = df_final["جمع مانده منشور"].fillna(0)
    df_final["کد گرید"] = df_final["کد گرید"].fillna("")
    df_final["اختلاف"] = df_final["جمع قیمت گرید"] - df_final["جمع مانده منشور"]

    def status(row):
        if pd.isna(row["کد غرفه"]):
            return "بدون نگاشت کد غرفه"
        if row["جمع قیمت گرید"] == 0:
            return "فقط در منشور"
        if row["جمع مانده منشور"] == 0:
            return "فقط در گرید"
        if row["اختلاف"] != 0:
            return "دارای مغایرت"
        return "مطابق"

    df_final["وضعیت"] = df_final.apply(status, axis=1)
    df_final["کد غرفه"] = df_final["کد غرفه"].astype("Int64")

    df_out = df_final[[
        "کد غرفه", "کد گرید",
        "جمع قیمت گرید", "جمع مانده منشور", "اختلاف", "وضعیت",
    ]].sort_values("کد غرفه", na_position="last")

    out_name = "نتیجه غرفه.xlsx"
    df_out.to_excel(out_name, index=False)
    print(f"✅ '{out_name}' ساخته شد ({len(df_out)} ردیف)")


# ============================================================
# اجرا
# ============================================================
report_personnel(df_forosh, df_grid_raw)
report_booth(df_forosh, df_grid_raw)

print("\n" + "=" * 60)
print("✅ هر دو گزارش با موفقیت ساخته شدند.")
print("=" * 60)
