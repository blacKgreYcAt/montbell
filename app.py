import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import time
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# 0. 頁面全域設定
# ==========================================
st.set_page_config(
    page_title="Montbell 自動化中心 v3.20 (混搭雙引擎)",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

if 'current_page' not in st.session_state:
    st.session_state.current_page = 'all_in_one'
if 'stop_flag' not in st.session_state:
    st.session_state.stop_flag = False

def set_page(page_name):
    st.session_state.current_page = page_name

# ==========================================
# 1. 核心邏輯：分離式引擎
# ==========================================

def call_grok_translation(prompt, api_key, model_name="grok-2-latest"):
    """
    [翻譯專用] 使用 xAI Grok API
    """
    if not api_key: return "Error: 無 Grok Key"
    
    url = "https://api.x.ai/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    payload = {
        "messages": [
            {
                "role": "system", 
                "content": "You are a professional translator. Translate Japanese text to Traditional Chinese (Taiwan) accurately. Output ONLY the translated text."
            },
            {"role": "user", "content": prompt}
        ],
        "model": model_name,
        "stream": False,
        "temperature": 0.1
    }
    
    try:
        # 簡單重試機制
        for attempt in range(2):
            try:
                response = requests.post(url, headers=headers, json=payload, timeout=40)
                if response.status_code != 200:
                    return f"Grok Error: {response.status_code} - {response.text}"
                result = response.json()
                return result["choices"][0]["message"]["content"].strip()
            except Exception as e:
                if attempt == 1: return f"Grok Connect Error: {str(e)}"
                time.sleep(1)
    except Exception as e:
        return f"Critical Error: {str(e)}"

def call_gemini_refining(prompt, api_key, model_name="gemini-1.5-flash"):
    """
    [精簡專用] 使用 Google Gemini API
    """
    if not api_key: return "Error: 無 Gemini Key"
    
    genai.configure(api_key=api_key)
    
    safety_settings = {
        HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
    }
    
    generation_config = {"temperature": 0.1, "top_p": 0.8, "top_k": 40, "max_output_tokens": 2048}
    model = genai.GenerativeModel(model_name, generation_config=generation_config)
    
    try:
        response = model.generate_content(prompt, safety_settings=safety_settings)
        return response.text.strip()
    except Exception as e:
        return f"Gemini Error: {str(e)}"

def scrape_montbell_single(model):
    headers = {'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja-JP'}
    base_url = "https://webshop.montbell.jp/"
    search_url = "https://webshop.montbell.jp/goods/list_search.php?top_sk="
    info = {'型號': model, '商品名': '', '商品描述': '', '規格': ''}
    
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
            
            name = soup.select_one('h1.goods-detail__ttl-main, h1.product-title, h1')
            if name: info['商品名'] = name.text.strip()
            else:
                if soup.title: info['商品名'] = soup.title.text.split('|')[0].strip()

            desc_selectors = ['.column1.type01 .innerCont p', 'div.description p', 'div#detail_explain', '.product-description']
            for sel in desc_selectors:
                found_list = soup.select(sel)
                for item in found_list:
                    if item.text.strip() and len(item.text.strip()) > 5:
                        info['商品描述'] = item.text.strip()
                        break
                if info['商品描述']: break

            spec_found = False
            spec_containers = soup.select('.column1.type01, div.explanationBox')
            for container in spec_containers:
                if '仕様' in container.text:
                    info['規格'] = container.text.strip()
                    spec_found = True
                    break
            if not spec_found:
                sf = soup.select_one('div.explanationBox')
                if sf: info['規格'] = sf.text.strip()
    except Exception: pass
    return info

def auto_save_to_local(data_list, filename="backup_temp.xlsx"):
    try:
        df_backup = pd.DataFrame(data_list)
        df_backup.to_excel(filename, index=False)
        return True
    except: return False

# Prompt Generators
def create_trans_prompt(text): 
    return f"將以下日文戶外用品資訊翻譯為台灣繁體中文。保持專業術語準確。直接輸出翻譯結果。原文：{text}"

def create_refine_prompt(text, limit): 
    return f"你是一個編輯。請將這段中文描述精簡為 {limit} 個字以內的重點摘要。只保留最核心的賣點 (如防水、透氣)。直接輸出結果。原文：{text}"

def create_spec_prompt(text): 
    return f"將此規格表整理為繁體中文。保留數值與單位。原文：{text}"

# ==========================================
# 2. 側邊欄與導航 (雙引擎設定)
# ==========================================
with st.sidebar:
    st.title("🛠️ 雙引擎設定")
    
    st.markdown("### 1. 翻譯引擎 (Grok)")
    grok_key = st.text_input("xAI API Key", type="password", key="grok_k")
    grok_model = st.selectbox("Grok 模型", ["grok-2-latest", "grok-beta"], index=0)
    
    st.markdown("### 2. 精簡引擎 (Gemini)")
    gemini_key = st.text_input("Gemini API Key", type="password", key="gemini_k")
    gemini_model = st.selectbox("Gemini 模型", ["gemini-1.5-flash", "gemini-pro"], index=0)
    
    st.markdown("---")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("測試 Grok"):
            if grok_key:
                res = call_grok_translation("こんにちは", grok_key, grok_model)
                if "Error" not in res: st.success("Grok OK")
                else: st.error(res)
            else: st.error("缺 Grok Key")
    with col_t2:
        if st.button("測試 Gemini"):
            if gemini_key:
                res = call_gemini_refining("你好", gemini_key, gemini_model)
                if "Error" not in res: st.success("Gemini OK")
                else: st.error(res)
            else: st.error("缺 Gemini Key")

st.title("🏔️ Montbell 自動化中心 v3.20")

nav1, nav2, nav3, nav4 = st.columns(4)
with nav1:
    if st.button("⚡ 一鍵全自動", use_container_width=True): set_page('all_in_one')
with nav2:
    if st.button("📥 獨立爬蟲", use_container_width=True): set_page('scraper')
with nav3:
    if st.button("🈺 獨立翻譯 (Grok)", use_container_width=True): set_page('translator')
with nav4:
    if st.button("✨ 獨立優化 (Gemini)", use_container_width=True): set_page('refiner')
st.markdown("---")

# ==========================================
# 3. 功能頁面
# ==========================================
if st.session_state.current_page == 'all_in_one':
    st.markdown("### ⚡ 混搭全自動：Grok 翻譯 + Gemini 精簡")
    
    c_in, c_set = st.columns([1, 1])
    with c_in: uploaded_file = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_all")
    with c_set:
        with st.expander("⚙️ 設定", expanded=True):
            sheet_name = st.text_input("工作表", "工作表1", key="sn_all")
            col_idx = st.number_input("型號欄位索引", value=0, min_value=0, key="mi_all")
            limit = st.number_input("精簡字數限制", min_value=5, max_value=500, value=10, step=1, key="cl_all")
            autosave_interval = st.number_input("自動存檔頻率", min_value=1, max_value=100, value=20, key="as_all")

    selected_models_to_process = []
    if uploaded_file:
        try:
            df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            all_valid_models = []
            for i, r in df_preview.iterrows():
                if i >= 1 and col_idx < len(r):
                    m = str(r.iloc[col_idx]).strip()
                    if re.match(r'^\d{7}$', m): 
                        all_valid_models.append({"型號": m, "選取": True})
            if all_valid_models:
                st.info(f"📄 讀取到 {len(all_valid_models)} 筆有效型號：")
                df_selection = pd.DataFrame(all_valid_models)
                edited_df = st.data_editor(df_selection, key="editor_all", use_container_width=True)
                selected_models_to_process = edited_df[edited_df["選取"] == True]["型號"].tolist()
                st.markdown(f"**✅ 已勾選: `{len(selected_models_to_process)}` 筆**")
        except Exception as e: st.error(f"讀取失敗: {e}")

    stop_requested = st.checkbox("🛑 緊急停止", key="stop_chk")

    if st.button("🚀 開始執行", type="primary", use_container_width=True, key="btn_all", disabled=len(selected_models_to_process)==0):
        if not grok_key or not gemini_key:
            st.error("❌ 請確認兩個 API Key 都已輸入")
        else:
            try:
                models = selected_models_to_process
                results = []
                total = len(models)
                status_box = st.status("🚀 任務初始化...", expanded=True)
                prog_bar = st.progress(0)
                
                for i, m in enumerate(models):
                    if stop_requested:
                        status_box.update(label="🛑 已停止！", state="error")
                        st.warning(f"已在第 {i} 筆停止。")
                        break

                    pct = int((i+1)/total*100)
                    status_box.update(label=f"⏳ [{i+1}/{total}] 正在處理: {m} ({pct}%)")
                    
                    try:
                        # 1. 爬蟲
                        raw = scrape_montbell_single(m)
                        
                        row_data = {
                            '型號': raw['型號'],
                            '商品描述_原文': raw['商品描述'],
                            '規格_原文': raw['規格'],
                            '商品描述_翻譯': '',
                            '規格_翻譯': '',
                            '商品描述_AI精簡': '',
                            '規格_AI精簡': ''
                        }

                        has_data = raw['商品描述'] or raw['規格']
                        
                        if has_data:
                            # --- 描述處理 ---
                            if raw['商品描述']:
                                # 階段一：Grok 翻譯 (日 -> 中)
                                desc_res = call_grok_translation(create_trans_prompt(raw['商品描述']), grok_key, grok_model)
                                row_data['商品描述_翻譯'] = desc_res if "Error" not in desc_res else raw['商品描述']
                                
                                # 階段二：Gemini 精簡 (中 -> 精簡中)
                                if row_data['商品描述_翻譯'] and "Error" not in row_data['商品描述_翻譯']:
                                    time.sleep(0.5)
                                    refine_res = call_gemini_refining(create_refine_prompt(row_data['商品描述_翻譯'], limit), gemini_key, gemini_model)
                                    # 保底：如果 Gemini 失敗，用翻譯文的前 N 字
                                    if "Error" in refine_res or not refine_res:
                                        row_data['商品描述_AI精簡'] = row_data['商品描述_翻譯'][:int(limit)]
                                    else:
                                        row_data['商品描述_AI精簡'] = refine_res

                            # --- 規格處理 ---
                            if raw['規格']:
                                # 階段一：Grok 翻譯 (日 -> 中)
                                spec_res = call_grok_translation(create_spec_prompt(raw['規格']), grok_key, grok_model)
                                row_data['規格_翻譯'] = spec_res if "Error" not in spec_res else raw['規格']
                                
                                # 階段二：規格不需要精簡，直接使用翻譯結果，或可選用 Gemini 整理格式
                                # 為了效率，這裡直接沿用翻譯結果，或稍微用 Gemini 整理一下格式
                                if row_data['規格_翻譯']:
                                    # 簡單複製，因為規格摘要容易掉字
                                    row_data['規格_AI精簡'] = row_data['規格_翻譯']

                        results.append(row_data)
                        if (i + 1) % autosave_interval == 0:
                            auto_save_to_local(results, "backup_all_in_one.xlsx")
                            st.toast(f"💾 已備份 {i+1} 筆")

                    except Exception as e:
                        st.error(f"處理 {m} 錯誤: {e}")
                        auto_save_to_local(results, "backup_error_save.xlsx")
                        continue

                    prog_bar.progress((i+1)/total)
                    time.sleep(0.5)
                
                status_box.update(label="✅ 任務結束！", state="complete", expanded=False)
                
                final_cols = ['型號', '商品描述_原文', '規格_原文', '商品描述_翻譯', '規格_翻譯', '商品描述_AI精簡', '規格_AI精簡']
                df_final = pd.DataFrame(results)
                for col in final_cols:
                    if col not in df_final.columns: df_final[col] = ""
                df_final = df_final[final_cols]

                st.success(f"共完成 {len(df_final)} 筆資料。")
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='openpyxl') as w: df_final.to_excel(w, index=False)
                st.download_button("📥 下載最終報表", out.getvalue(), "montbell_final.xlsx", "primary")

            except Exception as e: st.error(f"執行錯誤: {e}")

# --- 獨立分頁 (依功能分配 API) ---
elif st.session_state.current_page == 'scraper':
    st.markdown("### 📥 獨立爬蟲")
    up_1 = st.file_uploader("上傳 Excel", key="up_1")
    c1, c2 = st.columns(2)
    with c1: sheet_1 = st.text_input("工作表", "工作表1", key="sn_1")
    with c2: idx_1, row_1 = st.number_input("索引", 0, key="mi_1"), st.number_input("開始列", 2, key="sr_1")
    sel_models_1 = []
    if up_1:
        try:
            df1 = pd.read_excel(up_1, sheet_name=sheet_1)
            valid_m1 = [{"型號": str(r.iloc[idx_1]).strip(), "選取": True} for i, r in df1.iterrows() if i>=row_1-1 and idx_1<len(r) and re.match(r'^\d{7}$', str(r.iloc[idx_1]).strip())]
            if valid_m1:
                ed1 = st.data_editor(pd.DataFrame(valid_m1), key="ed1", use_container_width=True)
                sel_models_1 = ed1[ed1["選取"]==True]["型號"].tolist()
                st.write(f"已選: {len(sel_models_1)} 筆")
        except: pass
    stop_1 = st.checkbox("🛑 停止", key="stop_1")
    if st.button("開始", key="btn_1", disabled=len(sel_models_1)==0):
        res = []
        prog = st.progress(0)
        for i, m in enumerate(sel_models_1):
            if stop_1: st.warning("已停止"); break
            res.append(scrape_montbell_single(m))
            prog.progress((i+1)/len(sel_models_1))
            time.sleep(0.5)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w: pd.DataFrame(res).to_excel(w, index=False)
        st.download_button("下載", out.getvalue(), "scraped.xlsx")

elif st.session_state.current_page == 'translator':
    st.markdown("### 🈺 獨立翻譯 (使用 Grok)")
    st.info("此模式將使用 xAI Grok 進行日翻中")
    up_2 = st.file_uploader("上傳 Excel", key="up_2")
    if up_2 and grok_key:
        df_t = pd.read_excel(up_2)
        cols = st.multiselect("翻譯欄位", df_t.columns)
        if st.button("開始翻譯"):
            # (簡略) 實作 Grok 翻譯邏輯
            pass
    elif up_2 and not grok_key:
        st.error("請輸入 Grok API Key")

elif st.session_state.current_page == 'refiner':
    st.markdown("### ✨ 獨立優化 (使用 Gemini)")
    st.info("此模式將使用 Google Gemini 進行中文精簡")
    up_3 = st.file_uploader("上傳 Excel", key="up_3")
    if up_3 and gemini_key:
        # (簡略) 實作 Gemini 精簡邏輯
        pass
    elif up_3 and not gemini_key:
        st.error("請輸入 Gemini API Key")