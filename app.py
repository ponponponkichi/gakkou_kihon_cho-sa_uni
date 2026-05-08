import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import time
import zipfile
import io
import datetime
import shutil
import pandas as pd
import openpyxl
import streamlit as st
import tempfile

# --- Streamlitの基本設定 ---
st.set_page_config(page_title="学校基本調査 一括DL＆整理ツール", layout="wide")

# 列名の表記ゆれを統一するための辞書
rename_dict = {}

# 過去のログから判明している様式のマスターリスト
known_forms = [
    "07go_1_教員数（本務者）（再掲）", "07go_A_学生数", "07go_B_教員数（本務者）", "07go_C_職員数", "07go_Z_教員数（兼務者）",
    "08go_2_学科別学生数のうち休学者数", "08go_3_学科別学生数のうち最低在学年限超過学生数（編入学者は除く。）", "08go_7_専攻科，別科及び科目等履修生等の学生数", "08go_D_学科別学生数|入学志願者数|入学者数", "08go_E", "08go_G_出身高校の所在地県別入学者数", "08go_O_年齢別入学者数（再掲）", "08go_R_短期大学・高等専門学校・専修学校（専門課程）・高等学校等専攻科からの編入学者数",
    "09go_4", "09go_5", "09go_8", "09go_H", "09go_I", "09go_S",
    "10go_6", "10go_9", "10go_J", "10go_K", "10go_Q", "10go_T",
    "11go_bekkei", "11go_gkssu",
    "20go",
    "30go_2_1", "30go_2_1_bekkei", "30go_2_2"
]

def get_hidden_header_index(filepath):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
        for row_idx in range(1, sheet.max_row + 1):
            if sheet.row_dimensions[row_idx].hidden:
                return row_idx - 1
    except Exception:
        pass
    return None

# --- UIの構築 ---
st.title("📊 学校基本調査 一括ダウンロード＆整理ツール")
st.write("大学改革支援・学位授与機構の「大学基本情報」から、「学校基本調査」指定様式のデータを一括取得し・複数年データを1本化します。")

# ラジオボタン
mode_str = st.radio(
    "処理モードを選択してください",
    ("① 全様式を1本化する（5分ほどかかる）", "② 指定様式のみ1本化する（おすすめ）"),
    horizontal=True
)
current_mode = "all" if "①" in mode_str else "select"

# チェックボックス（指定様式のみの場合）
selected_forms = []
if current_mode == "select":
    st.write("対象とする様式にチェックを入れてください：")
    cols = st.columns(4)
    for i, form in enumerate(known_forms):
        with cols[i % 4]:
            if st.checkbox(form, key=form):
                selected_forms.append(form)

st.divider()

# 実行ボタン
if st.button("🚀 全自動処理を開始する", type="primary"):
    
    if current_mode == "select" and not selected_forms:
        st.error("エラー: 「指定様式のみ1本化する」が選ばれましたが、様式が1つもチェックされていません。")
        st.stop()

    # --- UI更新用のプレースホルダー ---
    progress_bar = st.progress(0)
    status_text = st.empty()
    log_area = st.empty()
    log_messages = []

    def log(msg):
        log_messages.append(msg)
        log_area.text_area("実行ログ", value="\n".join(log_messages), height=300, disabled=True)

    log("="*50)
    log("ダウンロードを開始します...")
    log(f"処理モード: {'すべての様式' if current_mode == 'all' else f'指定された様式 ({len(selected_forms)}個)'}")

    with tempfile.TemporaryDirectory() as temp_dir:
        base_save_dir = os.path.join(temp_dir, "processed_data")
        os.makedirs(base_save_dir)
        temp_raw_dir = os.path.join(base_save_dir, "_raw_data_temp")
        os.makedirs(temp_raw_dir)

        try:
            # ==========================================
            # フェーズ1: ダウンロード＆一時解凍
            # ==========================================
            log("\n【フェーズ1】データのダウンロードを開始します...")
            top_url = "https://portal.niad.ac.jp/ptrt/table.html"
            response = requests.get(top_url)
            response.encoding = response.apparent_encoding
            soup = BeautifulSoup(response.text, 'html.parser')
            
            year_links = [a for a in soup.find_all('a', href=True) 
                          if a['href'].endswith('.html') and any(char.isdigit() for char in a['href'])]

            seen_urls = set()
            unique_links = []
            for link in year_links:
                full_url = urljoin(top_url, link.get('href'))
                if full_url not in seen_urls and "table.html" not in full_url:
                    seen_urls.add(full_url)
                    unique_links.append(link)

            total_links = len(unique_links)
            if total_links == 0:
                log("ダウンロード可能な年度データが見つかりませんでした。")
                st.stop()

            log(f"{total_links}年度分のデータを取得します。")

            for i, link in enumerate(unique_links):
                progress_val = (i / total_links) * 0.5
                progress_bar.progress(progress_val)
                status_text.text(f"フェーズ1進行中... ({int(progress_val * 100)}%)")

                page_url = urljoin(top_url, link.get('href'))
                year_text = link.get_text(strip=True).replace('\n', '').replace(' ', '')
                safe_year_text = year_text.replace('/', '_').replace(':', '_').replace('（', '(').replace('）', ')')
                year_folder_path = os.path.join(temp_raw_dir, safe_year_text)

                try:
                    page_res = requests.get(page_url)
                    page_res.encoding = page_res.apparent_encoding
                    page_soup = BeautifulSoup(page_res.text, 'html.parser')
                    
                    download_tag = None
                    for a_tag in page_soup.find_all('a', href=True):
                        if "DOWNLOAD" in a_tag.get_text().upper() or a_tag['href'].endswith('.zip'):
                            download_tag = a_tag
                            break
                    
                    if download_tag:
                        zip_url = urljoin(page_url, download_tag.get('href'))
                        log(f" -> ダウンロード＆解凍: {safe_year_text}")
                        zip_res = requests.get(zip_url)
                        with zipfile.ZipFile(io.BytesIO(zip_res.content)) as z:
                            z.extractall(year_folder_path)
                        time.sleep(1)
                    else:
                        log(f"  × スキップ: {safe_year_text}内にデータが見つかりません")
                except Exception as e:
                    log(f"  ! エラー ({safe_year_text}): {e}")

            # ==========================================
            # フェーズ2: 様式別のフォルダ分け
            # ==========================================
            progress_bar.progress(0.5)
            status_text.text("フェーズ2進行中... (50%)")
            log("\n【フェーズ2】ダウンロードしたファイルを様式別に整理します...")
            
            copied_count = 0
            for root, dirs, files in os.walk(temp_raw_dir):
                for file in files:
                    if file.endswith(('.xls', '.xlsx')):
                        file_path = os.path.join(root, file)
                        parts = file.split('_', 1)
                        if len(parts) < 2:
                            continue
                        
                        form_type, _ = os.path.splitext(parts[1])
                        form_type = form_type.replace('-', '_')
                        
                        if current_mode == "select" and form_type not in selected_forms:
                            continue
                        
                        form_dir = os.path.join(base_save_dir, form_type)
                        if not os.path.exists(form_dir):
                            os.makedirs(form_dir)
                            
                        dest_file_path = os.path.join(form_dir, file)
                        shutil.copy2(file_path, dest_file_path)
                        copied_count += 1
            
            log(f" -> {copied_count} 個の対象ファイルを各様式フォルダへ移動しました。")
            shutil.rmtree(temp_raw_dir, ignore_errors=True)

            # ==========================================
            # フェーズ3: データ結合とCSV一本化
            # ==========================================
            progress_bar.progress(0.6)
            status_text.text("フェーズ3進行中... (60%)")
            log("\n【フェーズ3】各様式フォルダ内でのデータ結合を開始します...")

            form_folders = [d for d in os.listdir(base_save_dir) if os.path.isdir(os.path.join(base_save_dir, d)) and d != "_raw_data_temp"]
            if current_mode == "select":
                form_folders = [d for d in form_folders if d in selected_forms]

            total_forms = len(form_folders)

            if total_forms == 0:
                log("結合対象のフォルダがありませんでした。")
            else:
                for i, form_folder in enumerate(form_folders):
                    progress_val = 0.6 + ((i / total_forms) * 0.4)
                    progress_bar.progress(progress_val)
                    status_text.text(f"フェーズ3進行中... ({int(progress_val * 100)}%)")

                    form_path = os.path.join(base_save_dir, form_folder)
                    log(f"\n■ 処理中: 様式 [{form_folder}]")
                    
                    target_header_idx = None
                    for file in os.listdir(form_path):
                        if file.endswith('.xlsx'):
                            idx = get_hidden_header_index(os.path.join(form_path, file))
                            if idx is not None:
                                target_header_idx = idx
                                log(f"  -> 非表示ヘッダー行を特定: {idx + 1}行目")
                                break
                    
                    if target_header_idx is None:
                        log(f"  ! 非表示行が見つからないため、この様式の結合をスキップします。")
                        continue

                    merged_data = []
                    for file in os.listdir(form_path):
                        if file.endswith(('.xls', '.xlsx')):
                            file_path = os.path.join(form_path, file)
                            try:
                                df = pd.read_excel(file_path, header=target_header_idx)
                                df.dropna(how='all', axis=0, inplace=True)
                                df.dropna(how='all', axis=1, inplace=True)

                                year = file.split('_')[0]
                                df.insert(0, 'ファイル年度', year)
                                
                                if rename_dict:
                                    df.rename(columns=rename_dict, inplace=True)
                                
                                record_count = len(df)
                                merged_data.append(df)
                                
                                # ★修正ポイント1: 【】をやめて _() に変更
                                name_part, ext_part = os.path.splitext(file)
                                new_filename = f"{name_part}_({record_count}_records){ext_part}"
                                os.rename(file_path, os.path.join(form_path, new_filename))
                                log(f"    - 読込成功: {new_filename}")
                            except Exception as e:
                                log(f"    × 読込エラー: {file} - {e}")
                                
                    if merged_data:
                        final_df = pd.concat(merged_data, ignore_index=True)
                        final_record_count = len(final_df)
                        # ★修正ポイント2: 【】をやめて _() に変更
                        merged_filename = f"{form_folder}_merged_({final_record_count}_records).csv"
                        save_path = os.path.join(form_path, merged_filename)
                        final_df.to_csv(save_path, index=False, encoding='utf-8-sig')
                        log(f"  ★ 結合完了: {merged_filename} を生成しました")

            with open(os.path.join(base_save_dir, "run_log.txt"), "w", encoding="utf-8") as f:
                f.write("\n".join(log_messages))

            # ==========================================
            # フェーズ4: ZIP化とダウンロードボタンの生成
            # ==========================================
            progress_bar.progress(1.0)
            status_text.text("処理完了 (100%)")
            log("\n==================================================")
            log("すべての処理が完了しました！ZIPファイルを作成しています...")

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for root, dirs, files in os.walk(base_save_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        arcname = os.path.relpath(file_path, base_save_dir)
                        zip_file.write(file_path, arcname)
            
            st.success("✅ 全ての処理が完了しました！下のボタンから結果をダウンロードしてください。")
            
            st.download_button(
                label="📁 整理済みのデータをダウンロード (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"school_basic_survey_data_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                type="primary"
            )

        except Exception as e:
            st.error(f"\n予期せぬエラーが発生しました: {e}")
            log(f"エラー詳細: {e}")
