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
    page_title="Montbell 自動化中心 v3.10 (篩選版)",
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
def get_gemini_response(prompt, api_key, model_name, retry=True):
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
        try:
            return response.text.strip()
        except ValueError:
            if retry:
                simple_prompt = f"Translate the following Japanese text to Traditional Chinese (Taiwan): {prompt.split('原文：')[-1]}"
                return get_gemini_response(simple_prompt, api_key, model_name, retry=False)
            return "Error: 內容被安全性攔截"
    except Exception as e:
        return f"Error: {str(e)}"

def get_available_models(api_key):
    try:
        genai.configure(api_key=api_key)
        return [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
    except: return []

def scrape_montbell_single(model):
    """爬蟲：只抓取 標題、描述、規格"""
    headers = {'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja-JP'}
    base_url = "https://webshop.montbell.jp/"
    search_url = "https://webshop.montbell.jp/goods/list_search.php?top_sk="
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
    st.info("ℹ️ **v3.10 篩選版**：\n上傳檔案後，您現在可以先勾選想要處理的項目，再開始執行。")

st.title("🏔️ Montbell 自動化中心 v3.10")

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

    # [v3.10] 新增：資料預覽與篩選區
    selected_models_to_process = []
    if uploaded_file:
        try:
            df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            # 提取所有符合格式的型號
            all_valid_models = []
            for i, r in df_preview.iterrows():
                if i >= 1 and col_idx < len(r):
                    m = str(r.iloc[col_idx]).strip()
                    if re.match(r'^\d{7}$', m): 
                        all_valid_models.append({"型號": m, "原始行數": i+2, "選取": True}) # 預設全選
            
            if all_valid_models:
                st.info(f"📄 讀取到 {len(all_valid_models)} 筆有效型號，請勾選您要處理的項目：")
                
                # 顯示可編輯的表格 (Data Editor)
                df_selection = pd.DataFrame(all_valid_models)
                edited_df = st.data_editor(
                    df_selection,
                    column_config={
                        "選取": st.column_config.CheckboxColumn(
                            "處理?",
                            help="勾選以進行處理",
                            default=True,
                        )
                    },
                    disabled=["型號", "原始行數"],
                    hide_index=True,
                    use_container_width=True,
                    key="editor_all"
                )
                
                # 獲取使用者最終勾選的型號
                selected_models_to_process = edited_df[edited_df["選取"] == True]["型號"].tolist()
                st.markdown(f"**✅ 目前已選擇: `{len(selected_models_to_process)}` 筆**")
                
            else:
                st.warning("⚠️ 在指定的欄位中找不到符合格式 (7碼數字) 的型號，請檢查設定。")
        except Exception as e:
            st.error(f"讀取 Excel 失敗: {e}")

    stop_requested = st.checkbox("🛑 緊急停止 (勾選後，處理完當前筆數即停止)", key="stop_chk")

    # 按鈕邏輯：只有當有選擇型號時才啟用
    if st.button("🚀 開始執行", type="primary", use_container_width=True, key="btn_all", disabled=len(selected_models_to_process)==0):
        if not api_key:
            st.error("❌ 請輸入 API Key")
        else:
            try:
                models = selected_models_to_process # 使用篩選後的清單
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
                        # 1. 爬蟲
                        raw = scrape_montbell_single(m)
                        
                        row_data = {
                            '型號': raw['型號'],
                            '商品名': raw['商品名'],
                            '商品URL': raw['商品URL'],
                            '商品描述_原文': raw['商品描述'],
                            '規格_原文': raw['規格'],
                            '商品描述_翻譯': '',
                            '規格_翻譯': '',
                            '商品描述_AI精簡': '',
                            '規格_AI精簡': ''
                        }

                        # 2. 翻譯與優化
                        has_data = raw['商品描述'] or raw['規格']
                        
                        if has_data:
                            if raw['商品描述']:
                                desc_res = get_gemini_response(create_trans_prompt(raw['商品描述']), api_key, selected_model)
                                row_data['商品描述_翻譯'] = desc_res
                                if "Error" not in desc_res:
                                    row_data['商品描述_AI精簡'] = get_gemini_response(create_refine_prompt(desc_res, limit), api_key, selected_model)

                            if raw['規格']:
                                spec_res = get_gemini_response(create_trans_prompt(raw['規格']), api_key, selected_model)
                                row_data['規格_翻譯'] = spec_res
                                if "Error" not in spec_res:
                                    row_data['規格_AI精簡'] = get_gemini_response(create_spec_prompt(spec_res), api_key, selected_model)
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
                
                final_cols = ['型號', '商品名', '商品描述_原文', '規格_原文', '商品描述_翻譯', '規格_翻譯', '商品描述_AI精簡', '規格_AI精簡', '商品URL']
                df_final = pd.DataFrame(results)
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
    
    # [v3.10] 爬蟲頁面的篩選
    sel_models_1 = []
    if up_1:
        try:
            df1 = pd.read_excel(up_1, sheet_name=sheet_1)
            valid_m1 = [{"型號": str(r.iloc[idx_1]).strip(), "選取": True} for i, r in df1.iterrows() if i>=row_1-1 and idx_1<len(r) and re.match(r'^\d{7}$', str(r.iloc[idx_1]).strip())]
            if valid_m1:
                st.info("勾選要爬取的項目：")
                ed1 = st.data_editor(pd.DataFrame(valid_m1), key="ed1", hide_index=True, use_container_width=True)
                sel_models_1 = ed1[ed1["選取"]==True]["型號"].tolist()
                st.write(f"已選: {len(sel_models_1)} 筆")
        except: pass

    stop_1 = st.checkbox("🛑 停止", key="stop_1")

    if st.button("開始", key="btn_1", disabled=len(sel_models_1)==0):
        try:
            res = []
            prog = st.progress(0)
            for i, m in enumerate(sel_models_1):
                if stop_1: st.warning("已停止"); break
                res.append(scrape_montbell_single(m))
                if (i+1)%20 == 0: auto_save_to_local(res, "backup_scrape.xlsx")
                prog.progress((i+1)/len(sel_models_1), text=f"進度 {int((i+1)/len(sel_models_1)*100)}%")
                time.sleep(0.5)
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: pd.DataFrame(res).to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "scraped.xlsx")
        except Exception as e: st.error(f"錯誤: {e}")

elif st.session_state.current_page == 'translator':
    st.markdown("### 🈺 獨立翻譯")
    up_2 = st.file_uploader("上傳 Excel", key="up_2")
    
    # [v3.10] 翻譯頁面的篩選
    df_t = pd.DataFrame()
    sel_indices_2 = [] # 翻譯需要記錄 index
    if up_2:
        try:
            df_t = pd.read_excel(up_2)
            df_t['選取'] = True # 預設全選
            st.info("勾選要翻譯的資料列：")
            # 顯示前幾欄供辨識
            disp_cols = list(df_t.columns[:5]) + ['選取']
            ed2 = st.data_editor(df_t[disp_cols], key="ed2", hide_index=False, use_container_width=True)
            # 找出被勾選的 index
            sel_indices_2 = ed2[ed2['選取']==True].index.tolist()
            st.write(f"已選: {len(sel_indices_2)} 筆")
        except: pass

    cols = st.multiselect("翻譯欄位", df_t.columns if not df_t.empty else [])
    stop_2 = st.checkbox("🛑 停止", key="stop_2")
    
    if st.button("開始", key="btn_2", disabled=len(sel_indices_2)==0 or not cols):
        if api_key:
            new_df = df_t.copy()
            prog = st.progress(0)
            # 計算總工作量
            total_ops = len(sel_indices_2) * len(cols)
            curr_op = 0
            
            for col in cols:
                new_df[f"{col}_TW"] = "" if f"{col}_TW" not in new_df.columns else new_df[f"{col}_TW"]
                
                for i in sel_indices_2: # 只遍歷選取的 index
                    if stop_2: break
                    val = new_df.at[i, col]
                    if pd.notna(val):
                        res = get_gemini_response(create_trans_prompt(str(val)), api_key, selected_model)
                        new_df.at[i, f"{col}_TW"] = res
                    curr_op += 1
                    if curr_op % 20 == 0: auto_save_to_local(new_df.to_dict('records'), "backup_trans.xlsx")
                    prog.progress(curr_op/total_ops, text=f"{int(curr_op/total_ops*100)}%")
                    time.sleep(0.5)
                if stop_2: break
            
            out = io.BytesIO()
            # 只輸出選取的部分，或是全部？通常希望保留原始檔結構，所以輸出全部，但只有選取的有翻譯
            with pd.ExcelWriter(out, engine='openpyxl') as w: new_df.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "translated.xlsx")

elif st.session_state.current_page == 'refiner':
    st.markdown("### ✨ 獨立優化")
    up_3 = st.file_uploader("上傳 Excel", key="up_3")
    
    # [v3.10] 優化頁面的篩選
    df_r = pd.DataFrame()
    sel_indices_3 = []
    if up_3:
        try:
            df_r = pd.read_excel(up_3)
            df_r['選取'] = True
            st.info("勾選要優化的資料列：")
            disp_cols_r = list(df_r.columns[:5]) + ['選取']
            ed3 = st.data_editor(df_r[disp_cols_r], key="ed3", hide_index=False, use_container_width=True)
            sel_indices_3 = ed3[ed3['選取']==True].index.tolist()
            st.write(f"已選: {len(sel_indices_3)} 筆")
        except: pass

    if not df_r.empty:
        c_d = st.selectbox("描述", df_r.columns)
        c_s = st.selectbox("規格", ["(不處理)"] + list(df_r.columns))
    
    lim = st.slider("字數", 10, 200, 50)
    stop_3 = st.checkbox("🛑 停止", key="stop_3")
    
    if st.button("開始", key="btn_3", disabled=len(sel_indices_3)==0):
        if api_key:
            # 初始化欄位
            df_r['精簡_AI'] = "" if '精簡_AI' not in df_r.columns else df_r['精簡_AI']
            if c_s != "(不處理)": df_r['規格_AI'] = "" if '規格_AI' not in df_r.columns else df_r['規格_AI']
            
            prog = st.progress(0)
            total = len(sel_indices_3)
            
            for idx, i in enumerate(sel_indices_3):
                if stop_3: st.warning("已停止"); break
                
                r = df_r.iloc[i]
                
                # 描述
                d_val = get_gemini_response(create_refine_prompt(str(r[c_d]), lim), api_key, selected_model) if pd.notna(r[c_d]) else ""
                if "Error" in d_val: d_val = ""
                df_r.at[i, '精簡_AI'] = d_val
                
                # 規格
                if c_s != "(不處理)" and pd.notna(r[c_s]):
                    s_val = get_gemini_response(create_spec_prompt(str(r[c_s])), api_key, selected_model)
                    if "Error" in s_val: s_val = ""
                    df_r.at[i, '規格_AI'] = s_val
                
                if (idx+1)%20 == 0: auto_save_to_local(df_r.to_dict('records'), "backup_refine.xlsx")
                prog.progress((idx+1)/total, text=f"{int((idx+1)/total*100)}%")
                time.sleep(0.5)
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: df_r.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "refined.xlsx")