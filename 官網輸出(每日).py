import os
import re
import shutil
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import subprocess
import csv
import numpy as np

# ---------- 最前面：自動更名價格表 ----------
def rename_price_excel(folder_path):
    files = os.listdir(folder_path)
    pattern = re.compile(r'.*价格表\.xlsx$')
    old_file = None
    for f in files:
        if pattern.match(f):
            old_file = f
            break
    if old_file:
        old_full_path = os.path.join(folder_path, old_file)
        new_file = '2.價格總表.xlsx'
        new_full_path = os.path.join(folder_path, new_file)
        if old_full_path != new_full_path:
            if os.path.exists(new_full_path):
                os.remove(new_full_path)
            os.rename(old_full_path, new_full_path)
            print(f'已將 "{old_file}" 重新命名為 "{new_file}"')
        return new_full_path
    else:
        raise FileNotFoundError('找不到任何「價格表.xlsx」檔案')

# ---------- Step 1: 圖片計數 + 浮水印 ----------
def fill_color_count_only(image_folder, excel_path, count_col_name='顏色總數'):
    df = pd.read_excel(excel_path)
    files = os.listdir(image_folder)
    if count_col_name not in df.columns:
        df[count_col_name] = 0
    for idx, row in df.iterrows():
        product_name = str(row.iloc[0])
        prefix = product_name[:14]
        matched_photos = [f for f in files if f.startswith(prefix)]
        df.at[idx, count_col_name] = len(matched_photos)
    df.to_excel(excel_path, index=False)
    print(f'已更新 {excel_path} 的「{count_col_name}」欄位')
    return df

def sum_color_counts(excel_path, col1='顏色總數A', col2='顏色總數B', col3='顏色總數C', sum_col='顏色總數合計'):
    df = pd.read_excel(excel_path)
    for col in [col1, col2, col3]:
        if col not in df.columns:
            df[col] = 0
    df[sum_col] = df[col1].fillna(0).astype(int) + df[col2].fillna(0).astype(int) + df[col3].fillna(0).astype(int)
    df.to_excel(excel_path, index=False)
    print(f'已計算 {col1} + {col2} + {col3} 到 {sum_col}')

def copy_and_rename_folder(src_folder, dst_parent, new_folder_name):
    dst_folder = os.path.join(dst_parent, new_folder_name)
    try:
        if os.path.exists(dst_folder):
            shutil.rmtree(dst_folder)
        shutil.copytree(src_folder, dst_folder)
        print(f'已複製資料夾到：{dst_folder}')
    except Exception as e:
        print(f'複製資料夾時發生錯誤：{e}')
    return dst_folder

def add_watermark(image_path, watermark_text, output_path, font_size=40):
    try:
        image = Image.open(image_path).convert("RGBA")
    except Exception as e:
        print(f"無法開啟圖片 {image_path}，錯誤：{e}")
        return
    watermark_layer = Image.new("RGBA", image.size, (0,0,0,0))
    font_path = "msjh.ttc"
    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        print(f"找不到字型檔 {font_path}，將使用預設字型")
        font = ImageFont.load_default()
    draw = ImageDraw.Draw(watermark_layer)
    bbox = draw.textbbox((0, 0), watermark_text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    padding_x, padding_y, margin = 8, 5, 20
    x = image.width - text_width - padding_x*2 - margin
    y = image.height - text_height - padding_y*2 - margin
    rect_xy = [x, y + 10, x + text_width + padding_x*2, y + text_height + padding_y*2 + 10]
    draw.rectangle(rect_xy, fill=(255,255,255,255))
    draw.text((x + padding_x, y + padding_y), watermark_text, font=font, fill=(0,0,0,255))
    watermarked_image = Image.alpha_composite(image, watermark_layer)
    try:
        watermarked_image.convert("RGB").save(output_path, "PNG")
    except Exception as e:
        print(f"無法儲存圖片 {output_path}，錯誤：{e}")

def batch_watermark_by_excel(excel_path, folder, prefix_letter, count_col_name='顏色總數'):
    df = pd.read_excel(excel_path)
    for idx, row in df.iterrows():
        style = str(row.get('商品名称', '') or '')
        if not style:
            print(f"第 {idx} 行商品名稱為空，略過")
            continue
        color_count = row.get(count_col_name, 0)
        try:
            color_count = int(color_count)
        except Exception:
            color_count = 0
        prefix = style[:14]
        matched_files = [f for f in os.listdir(folder)
                         if f.startswith(prefix) and f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        matched_files.sort()
        if len(matched_files) < color_count:
            print(f"{style} 圖片數不足，略過")
            continue
        for i in range(color_count):
            img_path = os.path.join(folder, matched_files[i])
            watermark_text = f"圖{i+1}"
            base, ext = os.path.splitext(matched_files[i])
            suffix = f"{prefix_letter}{i+1}"
            new_filename = f"{base}-{suffix}{ext}"
            output_path = os.path.join(folder, new_filename)
            add_watermark(img_path, watermark_text, output_path)
            try:
                os.remove(img_path)
            except Exception as e:
                print(f"刪除原檔 {img_path} 時發生錯誤：{e}")
            print(f"已為 {matched_files[i]} 加上浮水印 {watermark_text}，另存為 {new_filename}，原檔已刪除")
    print(f"✅ [{prefix_letter}] 已依Excel與前14碼對照批次加上浮水印，且只保留新檔案")

# ---------- Step 2: 複製圖片到目標資料夾 ----------
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
                        print(f"複製檔案 {src_file} 時發生錯誤：{e}")
        else:
            print(f"⚠️ 找不到資料夾：{src}")
    print("✅ 圖片已全部複製完畢！")

# ---------- Step 3: Git 提交與推送 ----------
def git_commit_push(repo_path, commit_message):
    try:
        os.chdir(repo_path)
        result = subprocess.run(['git', 'status', '--porcelain'], capture_output=True, text=True, encoding='utf-8')
        if (result.stdout or '').strip() == '':
            print('沒有檔案變更，無需提交。')
        else:
            subprocess.run(['git', 'add', '.'], check=True, encoding='utf-8')
            subprocess.run(['git', 'commit', '-m', commit_message], check=True, encoding='utf-8')
            subprocess.run(['git', 'push', 'origin', 'main'], check=True, encoding='utf-8')
            print('已成功推送到 GitHub main 分支。')
    except Exception as e:
        print(f'發生錯誤：{e}')

# ---------- Step 4: 產生圖片 CDN 連結清單 CSV ----------
def generate_cdn_csv(github_user, repo_name, image_root, csv_filename):
    cdn_base = f'https://cdn.jsdelivr.net/gh/{github_user}/{repo_name}@main'
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
    print(f"✅ 已建立圖片 CDN 連結清單：{csv_path}")

# ---------- 主流程 ----------
if __name__ == '__main__':

    # 1. 自動尋找並更名價格表
    folder_path = r'C:\Users\a0987\photo\images'
    excel_path = rename_price_excel(folder_path)  # 回傳新檔案完整路徑
    dst_parent = folder_path

    # 2. 批市規格
    image_folder1 = os.path.join(folder_path, '1.批市規格')
    dst_folder1 = copy_and_rename_folder(image_folder1, dst_parent, '1-1批市浮水印')
    fill_color_count_only(image_folder1, excel_path, count_col_name='顏色總數A')
    batch_watermark_by_excel(excel_path, dst_folder1, prefix_letter="A", count_col_name='顏色總數A')

    # 3. 穿搭規格
    image_folder2 = os.path.join(folder_path, '3.穿搭規格')
    dst_folder2 = copy_and_rename_folder(image_folder2, dst_parent, '3-1浮水印')
    fill_color_count_only(image_folder2, excel_path, count_col_name='顏色總數B')
    batch_watermark_by_excel(excel_path, dst_folder2, prefix_letter="B", count_col_name='顏色總數B')

    # 4. 男裝規格
    image_folder3 = os.path.join(folder_path, '5.男裝規格')
    dst_folder3 = copy_and_rename_folder(image_folder3, dst_parent, '5-1男裝浮水印')
    fill_color_count_only(image_folder3, excel_path, count_col_name='顏色總數C')
    batch_watermark_by_excel(excel_path, dst_folder3, prefix_letter="C", count_col_name='顏色總數C')

    # 5. 顏色數量加總
    sum_color_counts(
        excel_path,
        col1='顏色總數A',
        col2='顏色總數B',
        col3='顏色總數C',
        sum_col='顏色總數合計'
    )

    # 6. 複製所有圖片（含男裝）
    source_dirs = [
        image_folder1,
        dst_folder1,
        os.path.join(folder_path, '2.批市其他'),
        image_folder2,
        dst_folder2,
        os.path.join(folder_path, '4.穿搭其他'),
        image_folder3, 
        dst_folder3, 
        os.path.join(folder_path, '6.男裝其他')
    ]
    target_dir = r'E:\github\target'
    copy_images_to_target(source_dirs, target_dir)

    # 7. Git 操作
    git_commit_push(target_dir, '新增圖片分類')

    # 8. 產生 CDN 連結 CSV（指定新檔名與路徑）
    github_user = 'buyerbumblebee'
    repo_name = 'photo'
    image_csv_path = r'C:\Users\a0987\photo\images\1.圖片網址.csv'
    generate_cdn_csv(
        github_user,
        repo_name,
        target_dir,
        csv_filename=image_csv_path
    )

    # 9. 圖片網址與價格總表整合
    print("開始整合圖片網址與價格總表...")
    import pandas as pd
    import numpy as np

    # 指定你的檔案路徑
    price_path = os.path.join(folder_path, '2.價格總表.xlsx')
    output_path = os.path.join(folder_path, '整合.xlsx')

    # 先確認檔案是否存在
    if not os.path.exists(image_csv_path):
        print(f"找不到圖片網址檔案：{image_csv_path}")
        exit()
    if not os.path.exists(price_path):
        print(f"找不到價格總表檔案：{price_path}")
        exit()

    # 讀取資料
    image_df = pd.read_csv(image_csv_path, encoding='utf-8')
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

    # ✅ 新增：刪除整行沒有任何圖片連結的列
    df_final = df_final[~(df_final[img_cols].eq('').all(axis=1))]

    # 刪除H欄(第8欄, index=7)為0的列
    if df_final.shape[1] > 7:
        df_final = df_final[~((df_final.iloc[:, 7] == 0) | (df_final.iloc[:, 7] == '0'))]

    # NaN補空字串
    df_final = df_final.replace(np.nan, '', regex=True)

    # 輸出結果
    df_final.to_excel(output_path, index=False, engine='openpyxl')

    print(f"資料整合完成，請查看：{output_path}")
