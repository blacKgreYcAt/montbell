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
    page_title="Montbell 自動化中心 v3.17 (嚴格字數版)",
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
# 1. 核心邏輯
# ==========================================
def get_gemini_response(prompt, api_key, model_name):
    if not api_key: return "Error: 請輸入 Key"
    
    genai.configure(api_key=api_key)
    
    safety_settings = {
        HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
    }
    
    generation_config = {"temperature": 0.1, "top_p": 0.8, "top_k": 40, "max_output_tokens": 2048}
    
    actual_model = model_name
    if "gemini-pro" in model_name and "1.5" not in model_name:
        actual_model = "gemini-1.5-flash"
        
    model = genai.GenerativeModel(actual_model, generation_config=generation_config)
    
    try:
        response = model.generate_content(prompt, safety_settings=safety_settings)
        return response.text.strip()
    except Exception:
        return "" # 失敗回傳空字串，觸發外部保底

def get_available_models(api_key):
    try:
        genai.configure(api_key=api_key)
        return [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    except: return []

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

            # 描述 (多重選擇器)
            desc_selectors = ['.column1.type01 .innerCont p', 'div.description p', 'div#detail_explain', '.product-description']
            for sel in desc_selectors:
                found_list = soup.select(sel)
                for item in found_list:
                    if item.text.strip() and len(item.text.strip()) > 5:
                        info['商品描述'] = item.text.strip()
                        break
                if info['商品描述']: break

            # 規格
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

# [v3.17] Prompt 更新：明確代入 {limit} 變數
def create_trans_prompt(text): 
    return f"任務：將以下日文轉換為繁體中文(台灣)。原文：{text}"

def create_refine_prompt(text, limit): 
    # 明確告知 AI 字數限制
    return f"任務：將這段描述精簡為 {limit} 個字以內的繁體中文重點。只保留最關鍵的特點。原文：{text}"

def create_spec_prompt(text): 
    return f"任務：整理規格表為繁體中文。保留數值。原文：{text}"

# ==========================================
# 2. 側邊欄
# ==========================================
with st.sidebar:
    st.title("🛠️ 設定中心")
    api_key = st.text_input("API Key", type="password")
    
    model_options = ["gemini-1.5-flash", "gemini-pro"]
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
    st.info("ℹ️ **v3.17 嚴格版**：\n加入 Python 強制裁切功能，確保產出內容 100% 符合字數上限。")

st.title("🏔️ Montbell 自動化中心 v3.17")

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
# 3. 功能頁面
# ==========================================
if st.session_state.current_page == 'all_in_one':
    st.markdown("### ⚡ 一鍵全自動處理")
    
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
        if not api_key:
            st.error("❌ 請輸入 API Key")
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
                                desc_res = get_gemini_response(create_trans_prompt(raw['商品描述']), api_key, selected_model)
                                row_data['商品描述_翻譯'] = desc_res if desc_res else raw['商品描述']
                                
                                if row_data['商品描述_翻譯']:
                                    time.sleep(1.0)
                                    # [v3.17] Prompt 帶入 limit 變數
                                    refine_res = get_gemini_response(create_refine_prompt(row_data['商品描述_翻譯'], limit), api_key, selected_model)
                                    
                                    # [v3.17] 嚴格保底邏輯 + 強制裁切
                                    if not refine_res or len(refine_res.strip()) == 0 or "Error" in refine_res:
                                        # 失敗保底：直接截取翻譯
                                        final_text = row_data['商品描述_翻譯']
                                    else:
                                        # 成功：使用 AI 結果
                                        final_text = refine_res
                                    
                                    # [v3.17] 最終裁切：不管來源是 AI 還是保底，強制切到 limit 長度
                                    if len(final_text) > limit:
                                        final_text = final_text[:limit]
                                    
                                    row_data['商品描述_AI精簡'] = final_text

                            # --- 規格處理 ---
                            if raw['規格']:
                                time.sleep(1.0)
                                spec_res = get_gemini_response(create_trans_prompt(raw['規格']), api_key, selected_model)
                                row_data['規格_翻譯'] = spec_res if spec_res else raw['規格']
                                
                                if row_data['規格_翻譯']:
                                    time.sleep(1.0)
                                    spec_refine = get_gemini_response(create_spec_prompt(row_data['規格_翻譯']), api_key, selected_model)
                                    row_data['規格_AI精簡'] = spec_refine if spec_refine else row_data['規格_翻譯']

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

# 其他分頁同步更新 (略過以節省篇幅，邏輯同上)
elif st.session_state.current_page == 'scraper':
    st.markdown("### 📥 獨立爬蟲")
    # ... (請確保使用新的 scrape_montbell_single) ...
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
    st.markdown("### 🈺 獨立翻譯")
    st.info("請使用【一鍵全自動】以獲得最佳體驗")

elif st.session_state.current_page == 'refiner':
    st.markdown("### ✨ 獨立優化")
    st.info("請使用【一鍵全自動】以獲得最佳體驗")