import os
import shutil
import pandas as pd
import numpy as np
import subprocess
import csv
import re
from collections import defaultdict

# ---------- Step 1: 圖片計數 ----------

def fill_color_count_only(image_folder, excel_path, count_col_name='顏色總數'):
    if not os.path.exists(image_folder):
        print(f"⚠️ 圖片資料夾不存在：{image_folder}，跳過圖片統計")
        return pd.read_excel(excel_path)

    df = pd.read_excel(excel_path)
    files = os.listdir(image_folder)

    if count_col_name not in df.columns:
        df[count_col_name] = 0

    valid_exts = ('.jpg', '.jpeg', '.png', '.webp')
    for idx, row in df.iterrows():
        product_name = str(row.iloc[0])
        prefix = product_name[:14]
        matched_photos = [f for f in files if f.startswith(prefix) and f.lower().endswith(valid_exts)]
        df.at[idx, count_col_name] = len(matched_photos)

    df.to_excel(excel_path, index=False)
    print(f'✅ 已更新 {excel_path} 的「{count_col_name}」欄位')
    return df

def sum_color_counts(excel_path, col1='顏色總數B', col2='顏色總數合計'):
    df = pd.read_excel(excel_path)
    for col in [col1, col2]:
        if col not in df.columns:
            df[col] = 0
    df['顏色總數合計'] = df[col1].fillna(0).astype(int)
    df.to_excel(excel_path, index=False)
    print(f'✅ 已更新 {col1} 到 顏色總數合計')

# ---------- 複製並重新命名（同款分組流水號） ----------

def copy_and_rename_folder(src_folder, dst_parent, new_folder_name):
    dst_folder = os.path.join(dst_parent, new_folder_name)
    try:
        if os.path.exists(dst_folder):
            shutil.rmtree(dst_folder)
        os.makedirs(dst_folder, exist_ok=True)
        valid_exts = ('.jpg', '.jpeg', '.png', '.webp')
        files = [f for f in os.listdir(src_folder) if f.lower().endswith(valid_exts)]
        files.sort()  # 保持順序

        # 依前14碼分組
        group_dict = defaultdict(list)
        for file in files:
            prefix = file[:14]
            group_dict[prefix].append(file)

        # 各組內依檔名排序，流水號從A1開始
        for group_files in group_dict.values():
            group_files.sort()
            for idx, file in enumerate(group_files, 1):
                src_file = os.path.join(src_folder, file)
                name, ext = os.path.splitext(file)
                new_name = f"{name}-A{idx}{ext}"
                dst_file = os.path.join(dst_folder, new_name)
                shutil.copy2(src_file, dst_file)

        print(f'✅ 已依同款分組並重新命名複製到：{dst_folder}')
    except Exception as e:
        print(f'❌ 複製資料夾時發生錯誤：{e}')
    return dst_folder

# ---------- 複製圖片到目標資料夾 ----------

def copy_images_to_target(source_dirs, target_dir):
    os.makedirs(target_dir, exist_ok=True)
    img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif')
    for src in source_dirs:
        if os.path.exists(src):
            print(f"✅ 找到資料夾：{src}")
            for file in os.listdir(src):
                if file.lower().endswith(img_exts):
                    src_file = os.path.join(src, file)
                    dst_file = os.path.join(target_dir, file)
                    try:
                        shutil.copy2(src_file, dst_file)
                    except Exception as e:
                        print(f"❌ 複製檔案 {src_file} 發生錯誤：{e}")
        else:
            print(f"⚠️ 找不到資料夾：{src}")
    print("✅ 圖片已全部複製完畢！")

# ---------- Git 提交與推送 ----------

def git_commit_push(repo_path, commit_message):
    try:
        os.chdir(repo_path)
        result = subprocess.run(['git', 'status', '--porcelain'], capture_output=True, text=True, encoding='utf-8')
        if (result.stdout or '').strip() == '':
            print('ℹ️ 沒有檔案變更，無需提交。')
        else:
            subprocess.run(['git', 'add', '.'], check=True, encoding='utf-8')
            subprocess.run(['git', 'commit', '-m', commit_message], check=True, encoding='utf-8')
            subprocess.run(['git', 'push', 'origin', 'main'], check=True, encoding='utf-8')
            print('✅ 已成功推送到 GitHub main 分支。')
    except Exception as e:
        print(f'❌ Git 操作發生錯誤：{e}')

# ---------- 產生圖片 CDN 連結清單 CSV ----------

def generate_cdn_csv(github_user, repo_name, image_root, csv_filename):
    cdn_base = f'https://cdn.jsdelivr.net/gh/{github_user}/{repo_name}@main'
    # 絕對路徑直接用，不再 join
    csv_path = csv_filename if os.path.isabs(csv_filename) else os.path.join(image_root, csv_filename)
    valid_exts = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'}
    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['檔名', 'CDN 連結'])
        for root, dirs, files in os.walk(image_root):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext in valid_exts:
                    rel_path = os.path.relpath(os.path.join(root, file), image_root).replace('\\', '/')
                    url = f'{cdn_base}/{rel_path}'
                    writer.writerow([rel_path, url])
    print(f"✅ 已建立圖片 CDN 清單：{csv_path}")

# ---------- 合併多個 Excel 並補齊8欄位 ----------

def merge_excels(excel_paths, output_path):
    # 強制補齊8個標準欄位
    standard_cols = [
        "商品名称", "尺码", "分类", "单价",
        "顏色總數A", "顏色總數B", "顏色總數C", "顏色總數合計"
    ]
    df_list = []
    for path in excel_paths:
        if os.path.exists(path):
            df = pd.read_excel(path)
            # 補齊缺少欄位
            for col in standard_cols:
                if col not in df.columns:
                    df[col] = ""
            df = df[standard_cols]
            df_list.append(df)
        else:
            print(f"⚠️ 找不到 Excel 檔案：{path}")
    if df_list:
        merged_df = pd.concat(df_list, ignore_index=True)
        merged_df.to_excel(output_path, index=False)
        print(f"✅ 已合併所有 Excel 並輸出至：{output_path}")
    else:
        print("⚠️ 無可合併的 Excel 檔案")

# ---------- 處理子資料夾 ----------

def process_subfolder(date_folder, subfolder_name):
    full_subfolder = os.path.join(date_folder, subfolder_name)
    print(f"開始處理子資料夾：{full_subfolder}")

    if not os.path.exists(full_subfolder):
        print(f"⚠️ 找不到 {full_subfolder}，跳過此資料夾。")
        return None

    prefix = subfolder_name[0]
    excel_files = [f for f in os.listdir(full_subfolder)
                   if f.startswith(prefix) and f.endswith('价格表.xlsx')]
    if not excel_files:
        print(f"⚠️ {full_subfolder} 找不到價格表 Excel 檔案，跳過。")
        return None

    excel_path = os.path.join(full_subfolder, excel_files[0])
    dst_parent = full_subfolder

    image_folder2 = os.path.join(full_subfolder, '3.穿搭規格')
    dst_folder2 = None
    if os.path.exists(image_folder2):
        dst_folder2 = copy_and_rename_folder(image_folder2, dst_parent, '3-1浮水印')
        fill_color_count_only(image_folder2, excel_path, count_col_name='顏色總數B')

    sum_color_counts(excel_path, col1='顏色總數B', col2='顏色總數合計')

    source_dirs = [
        dst_folder2 if dst_folder2 else '',
        os.path.join(full_subfolder, '4.穿搭其他')
    ]
    source_dirs = [d for d in source_dirs if d and os.path.exists(d)]

    target_dir = r'E:\github\target'
    copy_images_to_target(source_dirs, target_dir)

    return excel_path

# ---------- 主流程 ----------

if __name__ == '__main__':
    parent_folder = r'C:\Users\a0987\photo\images'
    # 自動尋找所有 WEEK+數字 資料夾
    week_folders = [
        f for f in os.listdir(parent_folder)
        if os.path.isdir(os.path.join(parent_folder, f)) and re.match(r'^WEEK\d+$', f, re.IGNORECASE)
    ]

    subfolders_to_process = ['A-当日新款', 'Z-当日新款']
    processed_excels = []

    for week_folder in week_folders:
        base_folder = os.path.join(parent_folder, week_folder)
        date_folders = [f for f in os.listdir(base_folder)
                        if os.path.isdir(os.path.join(base_folder, f))]
        for date_folder in date_folders:
            for subfolder_name in subfolders_to_process:
                full_path = os.path.join(base_folder, date_folder, subfolder_name)
                if os.path.exists(full_path):
                    print(f"處理：{full_path}")
                    excel_path = process_subfolder(os.path.join(base_folder, date_folder), subfolder_name)
                    if excel_path:
                        processed_excels.append(excel_path)
                else:
                    print(f"⚠️ 不存在資料夾：{full_path}")

    target_dir = r'E:\github\target'
    github_user = 'buyerbumblebee'
    repo_name = 'photo'

    git_commit_push(target_dir, '新增圖片分類 - 全部日期完成')

    # 產生圖片連結 CSV，輸出到 C:\Users\a0987\photo\images\1.圖片網址.csv
    generate_cdn_csv(
        github_user,
        repo_name,
        target_dir,
        csv_filename=r'C:\Users\a0987\photo\images\1.圖片網址.csv'
    )

    # 合併 Excel，輸出到 C:\Users\a0987\photo\images\2.價格總表.xlsx
    merged_excel_output = r'C:\Users\a0987\photo\images\2.價格總表.xlsx'
    merge_excels(processed_excels, merged_excel_output)

    print("✅ 所有檔案已依需求輸出完畢！")

    # ---------- 整合圖片網址與價格總表 ----------
    image_path = r'C:\Users\a0987\photo\images\1.圖片網址.csv'
    price_path = r'C:\Users\a0987\photo\images\2.價格總表.xlsx'
    output_path = r'C:\Users\a0987\photo\images\整合.xlsx'

    # 先確認檔案是否存在
    if not os.path.exists(image_path):
        print(f"找不到圖片網址檔案：{image_path}")
        exit()
    if not os.path.exists(price_path):
        print(f"找不到價格總表檔案：{price_path}")
        exit()

    # 讀取資料
    image_df = pd.read_csv(image_path, encoding='utf-8')
    price_df = pd.read_excel(price_path, engine='openpyxl')

    # 建立圖片對照表（key=檔名前兩段）
    image_map = {}
    for idx, row in image_df.iterrows():
        filename = str(row[0])  # A欄（第1欄，index=0）
        link = row[1]           # B欄（第2欄，index=1）
        if pd.isna(filename):
            continue
        key = '-'.join(filename.split('-')[:2])
        image_map.setdefault(key, []).append(link)

    # 寫入圖片連結，補齊20欄
    img_cols = [f'圖{i}' for i in range(1, 21)]
    img_data = []
    for _, row in price_df.iterrows():
        product_name = str(row[0])  # 商品名稱通常在A欄（第1欄，index=0）
        if pd.isna(product_name):
            img_data.append(['']*20)
            continue
        key = '-'.join(product_name.split('-')[:2])
        links = image_map.get(key, [])
        links = links[:20] + [''] * (20 - len(links))
        img_data.append(links)

    img_df = pd.DataFrame(img_data, columns=img_cols)

    # 合併主檔與圖片欄
    price_df = price_df.reset_index(drop=True)
    df_final = pd.concat([price_df, img_df], axis=1)

    # 刪除整行沒有任何圖片連結的列
    df_final = df_final[~(df_final[img_cols].eq('').all(axis=1))]

    # 刪除H欄(第8欄, index=7)為0的列
    if df_final.shape[1] > 7:
        df_final = df_final[~((df_final.iloc[:, 7] == 0) | (df_final.iloc[:, 7] == '0'))]

    # NaN補空字串
    df_final = df_final.replace(np.nan, '', regex=True)

    # 輸出結果
    df_final.to_excel(output_path, index=False, engine='openpyxl')

    print(f"資料整合完成，請查看：{output_path}")
