import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import time
import re
import io
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# 0. 頁面全域設定
# ==========================================
st.set_page_config(
    page_title="Montbell 自動化中心 v3.6 (防斷線版)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 優化
st.markdown("""
    <style>
    div.stButton > button {
        height: 3.5em;
        font-size: 1.2em !important;
        font-weight: bold;
        border-radius: 10px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    }
    .main-content {
        padding: 20px;
        background-color: #f9f9f9;
        border-radius: 15px;
        margin-top: 20px;
        border: 1px solid #eee;
    }
    </style>
""", unsafe_allow_html=True)

# 初始化 Session State
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'all_in_one'
# [v3.6 新增] 用於控制停止的標記
if 'stop_flag' not in st.session_state:
    st.session_state.stop_flag = False

def set_page(page_name):
    st.session_state.current_page = page_name

# ==========================================
# 1. 核心邏輯與工具函式
# ==========================================
def get_gemini_response(prompt, api_key, model_name):
    """呼叫 Gemini API (已解除安全限制)"""
    if not api_key: return "Error: 請輸入 Key"
    try:
        genai.configure(api_key=api_key)
        safety_settings = {
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
        }
        generation_config = {"temperature": 0.2, "top_p": 0.8, "top_k": 40, "max_output_tokens": 2048}
        model = genai.GenerativeModel(model_name.strip(), generation_config=generation_config)
        response = model.generate_content(prompt, safety_settings=safety_settings)
        return response.text.strip()
    except Exception as e:
        if "SAFETY" in str(e): return "Error: 內容被安全性攔截"
        return f"Error: {str(e)}"

def get_available_models(api_key):
    try:
        genai.configure(api_key=api_key)
        return [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    except: return []

def scrape_montbell_single(model):
    """爬蟲邏輯 (加入更強的錯誤捕捉)"""
    headers = {'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja-JP'}
    base_url = "https://webshop.montbell.jp/"
    search_url = "https://webshop.montbell.jp/goods/list_search.php?top_sk="
    info = {'型號': model, '商品名': '', '價格': '', '商品描述': '', '規格': '', '機能': '', '商品URL': ''}
    try:
        target_url = f"{base_url}goods/disp.php?product_id={model}"
        resp = requests.get(target_url, headers=headers, timeout=10)
        if resp.status_code != 200:
            search_resp = requests.get(f"{search_url}{model}", headers=headers, timeout=10)
            if search_resp.status_code == 200:
                soup_s = BeautifulSoup(search_resp.text, 'html.parser')
                link = soup_s.select_one('div.product a, div.goods-container a')
                if link:
                    target_url = base_url + link['href'].lstrip('/')
                    resp = requests.get(target_url, headers=headers, timeout=10)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            info['商品URL'] = target_url
            name = soup.select_one('h1.goods-detail__ttl-main, h1')
            if name: info['商品名'] = name.text.strip()
            price = soup.select_one('.goods-detail__price, span.selling_price')
            if price: info['價格'] = price.text.strip()
            desc = soup.select('.column1.type01 .innerCont p')
            if desc: info['商品描述'] = desc[0].text.strip()
            spec = soup.select('.column1.type01, div.explanationBox')
            for s in spec:
                if '仕様' in s.text: info['規格'] = s.text.strip()
                if '機能' in s.text: info['機能'] = s.text.strip()
            if not info['規格']:
                sf = soup.select_one('div.explanationBox')
                if sf: info['規格'] = sf.text.strip()
    except Exception: pass # 爬蟲失敗就回傳空值，不中斷
    return info

# [v3.6 新增] 自動備份函式
def auto_save_to_local(data_list, filename="backup_temp.xlsx"):
    """將目前進度寫入本地 Excel (避免瀏覽器崩潰資料全失)"""
    try:
        df_backup = pd.DataFrame(data_list)
        df_backup.to_excel(filename, index=False)
        return True
    except:
        return False

# Prompt Generators
def create_trans_prompt(text): return f"角色：專業戶外用品譯者(台灣)。任務：將日文翻譯為繁體中文。原則：1.專有名詞台式化。2.語氣自然。3.不解釋。原文：{text}"
def create_refine_prompt(text, limit): return f"任務：提取賣點並精簡。限制：{limit}字內。原文：{text}"
def create_spec_prompt(text): return f"任務：優化規格表。規則：保留【】標題，去除贅字，縮寫，保持換行。原文：{text}"

# ==========================================
# 2. 側邊欄與導航
# ==========================================
with st.sidebar:
    st.title("🛠️ 設定中心")
    api_key = st.text_input("API Key", type="password")
    
    # 模型選擇
    model_options = ["gemini-pro"]
    if api_key:
        detected = get_available_models(api_key)
        if detected: model_options = detected
    selected_model = st.selectbox("AI 模型", model_options, index=0)
    
    if st.button("測試連線"):
        try:
            genai.configure(api_key=api_key)
            m = genai.GenerativeModel(selected_model)
            m.generate_content("Hi")
            st.success("✅ 連線成功")
        except Exception as e: st.error(f"❌ 失敗: {e}")
        
    st.markdown("---")
    st.info("ℹ️ **v3.6 安全機制**：\n每處理 20 筆資料，系統會自動在您的資料夾產生一份 `backup_temp.xlsx`。")

st.title("🏔️ Montbell 自動化中心 v3.6")

# 四大導航鍵
nav1, nav2, nav3, nav4 = st.columns(4)
with nav1:
    if st.button("⚡ 一鍵全自動", use_container_width=True): set_page('all_in_one')
with nav2:
    if st.button("📥 獨立爬蟲", use_container_width=True): set_page('scraper')
with nav3:
    if st.button("🈺 獨立翻譯", use_container_width=True): set_page('translator')
with nav4:
    if st.button("✨ 獨立優化", use_container_width=True): set_page('refiner')
st.markdown("---")

# ==========================================
# 3. 功能頁面實作
# ==========================================

if st.session_state.current_page == 'all_in_one':
    st.markdown("### ⚡ 一鍵全自動處理 (含斷線保護)")
    
    c_in, c_set = st.columns([1, 1])
    with c_in: uploaded_file = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_all")
    with c_set:
        with st.expander("⚙️ 設定", expanded=True):
            sheet_name = st.text_input("工作表", "工作表1", key="sn_all")
            col_idx = st.number_input("型號欄位索引", 0, key="mi_all")
            limit = st.number_input("字數限制", 50, 10, key="cl_all")
            # [v3.6] 讓使用者設定多少筆存一次
            autosave_interval = st.number_input("自動存檔頻率 (筆數)", 10, 100, 20, help="每處理幾筆就備份一次到本地硬碟")

    # [v3.6] 停止按鈕的 UI 邏輯比較特殊，我們用一個 Checkbox 來模擬「請求停止」
    stop_requested = st.checkbox("🛑 緊急停止 (勾選後，程式將在處理完當前這一筆後停止並結算)", key="stop_chk")

    if st.button("🚀 開始執行", type="primary", use_container_width=True, key="btn_all"):
        if not uploaded_file or not api_key:
            st.error("❌ 資料不全：請檢查 API Key 或 檔案")
        else:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                models = []
                for i, r in df.iterrows():
                    if i >= 1 and col_idx < len(r):
                        m = str(r.iloc[col_idx]).strip()
                        if re.match(r'^\d{7}$', m): models.append(m)
                
                if not models:
                    st.error("找不到有效型號")
                else:
                    results = []
                    total = len(models)
                    
                    # [v3.6] 使用 empty 容器來顯示即時狀態，避免畫面太亂
                    status_box = st.status("🚀 任務初始化...", expanded=True)
                    prog_bar = st.progress(0)
                    
                    for i, m in enumerate(models):
                        # [v3.6] 檢查是否按下停止
                        if stop_requested:
                            status_box.update(label="🛑 使用者請求停止！正在結算...", state="error")
                            st.warning(f"已在第 {i} 筆停止。目前資料已保存。")
                            break

                        pct = int((i+1)/total*100)
                        status_box.update(label=f"⏳ [{i+1}/{total}] 正在處理: {m} ({pct}%)")
                        
                        try:
                            # 1.爬蟲
                            raw = scrape_montbell_single(m)
                            # 2.翻譯
                            trans = raw.copy()
                            if raw['商品名'] and raw['商品名'] != '未找到':
                                trans['商品名_TW'] = get_gemini_response(create_trans_prompt(raw['商品名']), api_key, selected_model)
                                trans['商品描述_TW'] = get_gemini_response(create_trans_prompt(raw['商品描述']), api_key, selected_model)
                                trans['規格_TW'] = get_gemini_response(create_trans_prompt(raw['規格']), api_key, selected_model)
                                trans['機能_TW'] = get_gemini_response(create_trans_prompt(raw['機能']), api_key, selected_model)
                            else: trans['商品名_TW'] = "查無資料"
                            # 3.優化
                            if raw['商品名'] and raw['商品名'] != '未找到':
                                trans['精簡描述_AI'] = get_gemini_response(create_refine_prompt(trans['商品描述_TW'], limit), api_key, selected_model)
                                trans['規格_結構化_AI'] = get_gemini_response(create_spec_prompt(trans['規格_TW']), api_key, selected_model)
                            else:
                                trans['精簡描述_AI'] = ""
                                trans['規格_結構化_AI'] = ""
                            
                            results.append(trans)
                            
                            # [v3.6] 自動存檔機制
                            if (i + 1) % autosave_interval == 0:
                                save_success = auto_save_to_local(results, "backup_all_in_one.xlsx")
                                if save_success:
                                    st.toast(f"💾 已自動備份 {i+1} 筆資料到 backup_all_in_one.xlsx", icon="✅")

                        except Exception as e:
                            # [v3.6] 錯誤捕捉：不要崩潰，記錄錯誤並繼續
                            st.error(f"處理 {m} 時發生錯誤: {e}")
                            # 為了安全，發生錯誤時也存一次檔
                            auto_save_to_local(results, "backup_error_save.xlsx")
                            continue

                        prog_bar.progress((i+1)/total)
                        time.sleep(0.5)
                    
                    status_box.update(label="✅ 任務結束！", state="complete", expanded=False)
                    
                    df_final = pd.DataFrame(results)
                    st.success(f"共完成 {len(df_final)} 筆資料。")
                    
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as w: df_final.to_excel(w, index=False)
                    st.download_button("📥 下載最終結果", out.getvalue(), "montbell_final.xlsx", "primary")

# --- 其他頁面 (爬蟲/翻譯/優化) 結構類似，皆加入自動存檔邏輯 ---
elif st.session_state.current_page == 'scraper':
    st.markdown("### 📥 獨立爬蟲 (含備份)")
    up_1 = st.file_uploader("上傳 Excel", key="up_1")
    c1, c2 = st.columns(2)
    with c1: sheet_1 = st.text_input("工作表", "工作表1", key="sn_1")
    with c2: idx_1, row_1 = st.number_input("索引", 0, key="mi_1"), st.number_input("開始列", 2, key="sr_1")
    stop_1 = st.checkbox("🛑 停止爬蟲", key="stop_1")

    if st.button("開始", key="btn_1") and up_1:
        df = pd.read_excel(up_1, sheet_name=sheet_1)
        models = [str(r.iloc[idx_1]).strip() for i, r in df.iterrows() if i>=row_1-1 and idx_1<len(r) and re.match(r'^\d{7}$', str(r.iloc[idx_1]).strip())]
        
        res = []
        prog = st.progress(0)
        for i, m in enumerate(models):
            if stop_1: 
                st.warning("已停止"); break
            res.append(scrape_montbell_single(m))
            
            if (i+1)%20 == 0: 
                auto_save_to_local(res, "backup_scrape.xlsx")
                st.toast(f"已備份 {i+1} 筆")
                
            prog.progress((i+1)/len(models), text=f"進度 {int((i+1)/len(models)*100)}%")
            time.sleep(0.5)
            
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w: pd.DataFrame(res).to_excel(w, index=False)
        st.download_button("下載", out.getvalue(), "scraped.xlsx")

elif st.session_state.current_page == 'translator':
    st.markdown("### 🈺 獨立翻譯 (含備份)")
    up_2 = st.file_uploader("上傳 Excel", key="up_2")
    if up_2 and api_key:
        df_t = pd.read_excel(up_2)
        cols = st.multiselect("翻譯欄位", df_t.columns)
        stop_2 = st.checkbox("🛑 停止翻譯", key="stop_2")
        
        if st.button("開始", key="btn_2") and cols:
            new_df = df_t.copy()
            prog = st.progress(0)
            total = len(df_t) * len(cols)
            curr = 0
            for col in cols:
                new_df[f"{col}_TW"] = ""
                for i, r in new_df.iterrows():
                    if stop_2: break
                    if pd.notna(r[col]):
                        new_df.at[i, f"{col}_TW"] = get_gemini_response(create_trans_prompt(str(r[col])), api_key, selected_model)
                    curr += 1
                    if curr % 20 == 0:
                        auto_save_to_local(new_df.to_dict('records'), "backup_trans.xlsx")
                        st.toast("已自動備份")
                    prog.progress(curr/total, text=f"{int(curr/total*100)}%")
                    time.sleep(0.5)
                if stop_2: break
            
            if stop_2: st.warning("已停止")
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: new_df.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "translated.xlsx")

elif st.session_state.current_page == 'refiner':
    st.markdown("### ✨ 獨立優化 (含備份)")
    up_3 = st.file_uploader("上傳 Excel", key="up_3")
    if up_3 and api_key:
        df_r = pd.read_excel(up_3)
        c_d = st.selectbox("描述", df_r.columns)
        c_s = st.selectbox("規格", ["(不處理)"] + list(df_r.columns))
        lim = st.slider("字數", 10, 200, 50)
        stop_3 = st.checkbox("🛑 停止優化", key="stop_3")
        
        if st.button("開始", key="btn_3"):
            res_d, res_s = [], []
            prog = st.progress(0)
            total = len(df_r)
            for i, r in df_r.iterrows():
                if stop_3: 
                    st.warning("已停止"); break
                
                if pd.notna(r[c_d]): res_d.append(get_gemini_response(create_refine_prompt(str(r[c_d]), lim), api_key, selected_model))
                else: res_d.append("")
                
                if c_s != "(不處理)" and pd.notna(r[c_s]): res_s.append(get_gemini_response(create_spec_prompt(str(r[c_s])), api_key, selected_model))
                else: res_s.append("")
                
                if (i+1)%20 == 0:
                    temp_df = df_r.iloc[:len(res_d)].copy()
                    temp_df['精簡_AI'] = res_d
                    if c_s != "(不處理)": temp_df['規格_AI'] = res_s
                    auto_save_to_local(temp_df.to_dict('records'), "backup_refine.xlsx")
                    st.toast("已自動備份")
                    
                prog.progress((i+1)/total, text=f"{int((i+1)/total*100)}%")
                time.sleep(0.5)
            
            df_r = df_r.iloc[:len(res_d)] # 裁切到停止點
            df_r['精簡_AI'] = res_d
            if c_s != "(不處理)": df_r['規格_AI'] = res_s
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: df_r.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "refined.xlsx")