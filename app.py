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
    page_title="Montbell 自動化中心 v3.8 (精準版)",
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
# 1. 核心邏輯與工具函式
# ==========================================
def get_gemini_response(prompt, api_key, model_name):
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
    """爬蟲：只抓取 標題(為了辨識)、描述、規格"""
    headers = {'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja-JP'}
    base_url = "https://webshop.montbell.jp/"
    search_url = "https://webshop.montbell.jp/goods/list_search.php?top_sk="
    
    # [v3.8] 欄位簡化，只保留指定項目
    info = {'型號': model, '商品名': '', '商品描述': '', '規格': '', '商品URL': ''}
    
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
            
            name = soup.select_one('h1.goods-detail__ttl-main, h1.product-title, h1')
            if name: info['商品名'] = name.text.strip()
            else:
                if soup.title: info['商品名'] = soup.title.text.split('|')[0].strip()

            # [v3.8] 移除價格與機能抓取，專注描述與規格
            desc = soup.select('.column1.type01 .innerCont p')
            if desc: info['商品描述'] = desc[0].text.strip()
            
            spec = soup.select('.column1.type01, div.explanationBox')
            for s in spec:
                if '仕様' in s.text: info['規格'] = s.text.strip()
            
            if not info['規格']:
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

def create_trans_prompt(text): return f"角色：專業戶外用品譯者(台灣)。任務：將日文翻譯為繁體中文。原則：1.專有名詞台式化。2.語氣自然。3.不解釋。原文：{text}"
def create_refine_prompt(text, limit): return f"任務：提取賣點並精簡。限制：{limit}字內。原文：{text}"
def create_spec_prompt(text): return f"任務：優化規格表。規則：保留【】標題，去除贅字，縮寫，保持換行。原文：{text}"

# ==========================================
# 2. 側邊欄與導航
# ==========================================
with st.sidebar:
    st.title("🛠️ 設定中心")
    api_key = st.text_input("API Key", type="password")
    
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
    st.info("ℹ️ **v3.8 精準版**：\n只抓取並處理「描述」與「規格」，產出包含 原文/翻譯/AI精簡 的完整報表。")

st.title("🏔️ Montbell 自動化中心 v3.8")

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
    st.markdown("### ⚡ 一鍵全自動處理 (描述 & 規格專用)")
    
    c_in, c_set = st.columns([1, 1])
    with c_in: uploaded_file = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_all")
    with c_set:
        with st.expander("⚙️ 設定", expanded=True):
            sheet_name = st.text_input("工作表", "工作表1", key="sn_all")
            col_idx = st.number_input("型號欄位索引", value=0, min_value=0, key="mi_all")
            limit = st.number_input("精簡字數限制", min_value=10, max_value=500, value=50, step=10, key="cl_all")
            autosave_interval = st.number_input("自動存檔頻率", min_value=1, max_value=100, value=20, key="as_all")

    stop_requested = st.checkbox("🛑 緊急停止 (勾選後，處理完當前筆數即停止)", key="stop_chk")

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
                    status_box = st.status("🚀 任務初始化...", expanded=True)
                    prog_bar = st.progress(0)
                    
                    for i, m in enumerate(models):
                        if stop_requested:
                            status_box.update(label="🛑 使用者請求停止！正在結算...", state="error")
                            st.warning(f"已在第 {i} 筆停止。")
                            break

                        pct = int((i+1)/total*100)
                        status_box.update(label=f"⏳ [{i+1}/{total}] 正在處理: {m} ({pct}%)")
                        
                        try:
                            # 1. 爬蟲 (只抓描述與規格)
                            raw = scrape_montbell_single(m)
                            
                            # 準備輸出的資料結構
                            row_data = {
                                '型號': raw['型號'],
                                '商品名': raw['商品名'], # 保留但不翻譯
                                '商品URL': raw['商品URL'],
                                # 原文
                                '商品描述_原文': raw['商品描述'],
                                '規格_原文': raw['規格'],
                                # 翻譯與優化 (預設空值)
                                '商品描述_翻譯': '',
                                '規格_翻譯': '',
                                '商品描述_AI精簡': '',
                                '規格_AI精簡': ''
                            }

                            # 2. 翻譯與優化
                            # 寬容判斷：只要有描述或規格，就處理
                            has_data = raw['商品描述'] or raw['規格']
                            
                            if has_data:
                                # 翻譯
                                if raw['商品描述']:
                                    trans_desc = get_gemini_response(create_trans_prompt(raw['商品描述']), api_key, selected_model)
                                    row_data['商品描述_翻譯'] = trans_desc
                                    # 優化
                                    row_data['商品描述_AI精簡'] = get_gemini_response(create_refine_prompt(trans_desc, limit), api_key, selected_model)
                                
                                if raw['規格']:
                                    trans_spec = get_gemini_response(create_trans_prompt(raw['規格']), api_key, selected_model)
                                    row_data['規格_翻譯'] = trans_spec
                                    # 優化
                                    row_data['規格_AI精簡'] = get_gemini_response(create_spec_prompt(trans_spec), api_key, selected_model)
                            else:
                                row_data['商品名'] = row_data['商品名'] + " (查無資料)"

                            results.append(row_data)
                            
                            if (i + 1) % autosave_interval == 0:
                                auto_save_to_local(results, "backup_all_in_one.xlsx")
                                st.toast(f"💾 已備份 {i+1} 筆")

                        except Exception as e:
                            st.error(f"處理 {m} 時發生錯誤: {e}")
                            auto_save_to_local(results, "backup_error_save.xlsx")
                            continue

                        prog_bar.progress((i+1)/total)
                        time.sleep(0.5)
                    
                    status_box.update(label="✅ 任務結束！", state="complete", expanded=False)
                    
                    # 整理最終 DataFrame 順序
                    final_cols = ['型號', '商品名', '商品描述_原文', '規格_原文', '商品描述_翻譯', '規格_翻譯', '商品描述_AI精簡', '規格_AI精簡', '商品URL']
                    df_final = pd.DataFrame(results)
                    # 確保欄位存在 (防止全空時報錯)
                    for col in final_cols:
                        if col not in df_final.columns: df_final[col] = ""
                    df_final = df_final[final_cols]

                    st.success(f"共完成 {len(df_final)} 筆資料。")
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as w: df_final.to_excel(w, index=False)
                    st.download_button("📥 下載最終報表", out.getvalue(), "montbell_final.xlsx", "primary")

            except Exception as e:
                st.error(f"執行錯誤: {e}")

# --- 其他頁面 ---
elif st.session_state.current_page == 'scraper':
    st.markdown("### 📥 獨立爬蟲 (僅描述與規格)")
    up_1 = st.file_uploader("上傳 Excel", key="up_1")
    c1, c2 = st.columns(2)
    with c1: sheet_1 = st.text_input("工作表", "工作表1", key="sn_1")
    with c2: idx_1, row_1 = st.number_input("索引", 0, key="mi_1"), st.number_input("開始列", 2, key="sr_1")
    stop_1 = st.checkbox("🛑 停止", key="stop_1")

    if st.button("開始", key="btn_1") and up_1:
        try:
            df = pd.read_excel(up_1, sheet_name=sheet_1)
            models = [str(r.iloc[idx_1]).strip() for i, r in df.iterrows() if i>=row_1-1 and idx_1<len(r) and re.match(r'^\d{7}$', str(r.iloc[idx_1]).strip())]
            res = []
            prog = st.progress(0)
            for i, m in enumerate(models):
                if stop_1: st.warning("已停止"); break
                res.append(scrape_montbell_single(m))
                if (i+1)%20 == 0: auto_save_to_local(res, "backup_scrape.xlsx")
                prog.progress((i+1)/len(models), text=f"進度 {int((i+1)/len(models)*100)}%")
                time.sleep(0.5)
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: pd.DataFrame(res).to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "scraped.xlsx")
        except Exception as e: st.error(f"錯誤: {e}")

elif st.session_state.current_page == 'translator':
    st.markdown("### 🈺 獨立翻譯")
    up_2 = st.file_uploader("上傳 Excel", key="up_2")
    if up_2 and api_key:
        df_t = pd.read_excel(up_2)
        cols = st.multiselect("翻譯欄位", df_t.columns)
        stop_2 = st.checkbox("🛑 停止", key="stop_2")
        if st.button("開始", key="btn_2") and cols:
            new_df = df_t.copy()
            prog = st.progress(0)
            total, curr = len(df_t) * len(cols), 0
            for col in cols:
                new_df[f"{col}_TW"] = ""
                for i, r in new_df.iterrows():
                    if stop_2: break
                    if pd.notna(r[col]):
                        new_df.at[i, f"{col}_TW"] = get_gemini_response(create_trans_prompt(str(r[col])), api_key, selected_model)
                    curr += 1
                    if curr % 20 == 0: auto_save_to_local(new_df.to_dict('records'), "backup_trans.xlsx")
                    prog.progress(curr/total, text=f"{int(curr/total*100)}%")
                    time.sleep(0.5)
                if stop_2: break
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: new_df.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "translated.xlsx")

elif st.session_state.current_page == 'refiner':
    st.markdown("### ✨ 獨立優化")
    up_3 = st.file_uploader("上傳 Excel", key="up_3")
    if up_3 and api_key:
        df_r = pd.read_excel(up_3)
        c_d = st.selectbox("描述", df_r.columns)
        c_s = st.selectbox("規格", ["(不處理)"] + list(df_r.columns))
        lim = st.slider("字數", 10, 200, 50)
        stop_3 = st.checkbox("🛑 停止", key="stop_3")
        if st.button("開始", key="btn_3"):
            res_d, res_s = [], []
            prog = st.progress(0)
            total = len(df_r)
            for i, r in df_r.iterrows():
                if stop_3: st.warning("已停止"); break
                res_d.append(get_gemini_response(create_refine_prompt(str(r[c_d]), lim), api_key, selected_model) if pd.notna(r[c_d]) else "")
                res_s.append(get_gemini_response(create_spec_prompt(str(r[c_s])), api_key, selected_model) if c_s != "(不處理)" and pd.notna(r[c_s]) else "")
                if (i+1)%20 == 0: 
                    temp = df_r.iloc[:len(res_d)].copy()
                    temp['精簡_AI'] = res_d
                    if c_s != "(不處理)": temp['規格_AI'] = res_s
                    auto_save_to_local(temp.to_dict('records'), "backup_refine.xlsx")
                prog.progress((i+1)/total, text=f"{int((i+1)/total*100)}%")
                time.sleep(0.5)
            df_r = df_r.iloc[:len(res_d)]
            df_r['精簡_AI'] = res_d
            if c_s != "(不處理)": df_r['規格_AI'] = res_s
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: df_r.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "refined.xlsx")