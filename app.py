import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import google.generativeai as genai
import time
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ==========================================
# 0. 頁面全域設定
# ==========================================
st.set_page_config(
    page_title="Montbell 自動化中心 v3.2",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS 優化
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        font-weight: bold;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 1. 核心邏輯函式庫
# ==========================================

def get_gemini_response(prompt, api_key, model_name):
    """呼叫 Gemini API 的通用函式"""
    if not api_key:
        return "Error: 請輸入 API Key"
    try:
        genai.configure(api_key=api_key)
        generation_config = {
            "temperature": 0.2,
            "top_p": 0.8,
            "top_k": 40,
            "max_output_tokens": 2048,
        }
        # 確保模型名稱沒有多餘空白
        clean_model_name = model_name.strip()
        model = genai.GenerativeModel(clean_model_name, generation_config=generation_config)
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"Error: {str(e)}"

def get_available_models(api_key):
    """[v3.2 新增] 自動偵測目前環境可用的模型列表"""
    try:
        genai.configure(api_key=api_key)
        models = []
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                # 只取 'models/' 後面的名稱，例如 'gemini-pro'
                name = m.name.replace('models/', '')
                models.append(name)
        return models
    except Exception as e:
        return []

def scrape_montbell_single(model):
    """爬取單一商品邏輯"""
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
                spec_fallback = soup.select_one('div.explanationBox')
                if spec_fallback: info['規格'] = spec_fallback.text.strip()
                
    except Exception as e:
        print(f"Scrape Error {model}: {e}")
    
    return info

def create_trans_prompt(text):
    return f"""
    角色：專業戶外用品譯者 (台灣市場)。
    任務：將日文翻譯為繁體中文 (台灣)。
    原則：1.專有名詞使用台灣習慣用語(如:透湿->透氣)。2.語氣通順自然。3.不解釋，直接輸出翻譯。
    原文：{text}
    """

def create_refine_prompt(text, limit):
    return f"任務：提取商品核心賣點並精簡。限制：{limit}個中文字內。原文：{text}"

def create_spec_prompt(text):
    return f"任務：優化並精簡產品規格表。規則：保留【】內標題，去除贅字，使用縮寫，保持換行。原文：{text}"

# ==========================================
# 2. 側邊欄：全域設定 (v3.2 智能偵測版)
# ==========================================
with st.sidebar:
    st.title("🛠️ 設定中心")
    st.info("👋 Hi Benjamin, v3.2 Auto-Detect")
    
    st.markdown("### 1. API 金鑰")
    api_key = st.text_input("Google Gemini API Key", type="password", placeholder="貼上 Key...")
    
    st.markdown("### 2. 模型選擇")
    
    # [v3.2] 預設的 fallback 選項，以防偵測失敗
    default_options = ["gemini-pro"] 
    model_options = default_options
    
    if api_key:
        # [v3.2] 嘗試自動獲取可用模型列表
        detected_models = get_available_models(api_key)
        if detected_models:
            model_options = detected_models
            st.success(f"已偵測到 {len(detected_models)} 個可用模型")
        else:
            st.warning("無法自動偵測模型，將使用預設列表。")
    
    selected_model = st.selectbox(
        "AI 模型", 
        model_options, 
        index=0,
        help="此列表由系統自動偵測您的 API Key 可用權限。"
    )
    
    # 測試按鈕
    if st.button("測試目前選擇的模型"):
        if not api_key:
            st.error("請先輸入 API Key")
        else:
            try:
                genai.configure(api_key=api_key)
                m = genai.GenerativeModel(selected_model)
                m.generate_content("Hello")
                st.success(f"✅ {selected_model} 連線成功！")
            except Exception as e:
                st.error(f"❌ 測試失敗: {e}")

    st.markdown("---")
    st.caption("Design for Montbell Workflow")

# ==========================================
# 3. 主畫面
# ==========================================
st.title("🏔️ Montbell 自動化中心 v3.2")

tabs = st.tabs(["⚡ 一鍵全自動", "📥 分步：爬蟲", "🈺 分步：翻譯", "✨ 分步：優化"])

# ==========================================
# TAB 1: 一鍵全自動
# ==========================================
with tabs[0]:
    st.header("⚡ 一鍵全自動處理流程")
    
    col_in, col_set = st.columns([1, 1])
    with col_in:
        uploaded_file_all = st.file_uploader("上傳型號 Excel", type=["xlsx", "xls"], key="up_all")
    with col_set:
        with st.expander("參數設定", expanded=True):
            sheet_name_all = st.text_input("工作表名稱", value="工作表1", key="sn_all")
            model_col_idx_all = st.number_input("型號欄位索引", value=0, min_value=0, key="mi_all")
            char_limit_all = st.number_input("精簡字數限制", value=50, min_value=10, key="cl_all")
            
    if st.button("🚀 啟動全自動排程", type="primary", key="btn_all"):
        if not uploaded_file_all or not api_key:
            st.error("請檢查：1.是否已上傳檔案 2.是否已輸入 API Key")
        else:
            try:
                df = pd.read_excel(uploaded_file_all, sheet_name=sheet_name_all)
                models = []
                for idx, row in df.iterrows():
                    if idx >= 1:
                        if model_col_idx_all < len(row):
                            m = str(row.iloc[model_col_idx_all]).strip()
                            if re.match(r'^\d{7}$', m): models.append(m)
                
                if not models:
                    st.error("找不到有效型號 (7碼數字)。")
                else:
                    results_final = []
                    with st.status(f"正在處理 {len(models)} 筆商品...", expanded=True) as status:
                        prog_bar = st.progress(0)
                        for i, model in enumerate(models):
                            status.update(label=f"[{i+1}/{len(models)}] 處理型號：{model}")
                            
                            # 1. 爬蟲
                            raw = scrape_montbell_single(model)
                            
                            # 2. 翻譯
                            trans = raw.copy()
                            if raw['商品名'] != '未找到':
                                trans['商品名_TW'] = get_gemini_response(create_trans_prompt(raw['商品名']), api_key, selected_model)
                                trans['商品描述_TW'] = get_gemini_response(create_trans_prompt(raw['商品描述']), api_key, selected_model)
                                trans['規格_TW'] = get_gemini_response(create_trans_prompt(raw['規格']), api_key, selected_model)
                                trans['機能_TW'] = get_gemini_response(create_trans_prompt(raw['機能']), api_key, selected_model)
                            else:
                                trans['商品名_TW'] = "查無資料"
                            
                            # 3. 優化
                            if raw['商品名'] != '未找到':
                                trans['精簡描述_AI'] = get_gemini_response(create_refine_prompt(trans['商品描述_TW'], char_limit_all), api_key, selected_model)
                                trans['規格_結構化_AI'] = get_gemini_response(create_spec_prompt(trans['規格_TW']), api_key, selected_model)
                            else:
                                trans['精簡描述_AI'] = ""
                                trans['規格_結構化_AI'] = ""

                            results_final.append(trans)
                            prog_bar.progress((i+1)/len(models))
                            time.sleep(1) 
                        
                        status.update(label="✅ 完成！", state="complete", expanded=False)

                    df_final = pd.DataFrame(results_final)
                    st.success(f"完成！共 {len(df_final)} 筆。")
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as w: df_final.to_excel(w, index=False)
                    st.download_button("📥 下載最終報表", out.getvalue(), "montbell_final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

            except Exception as e:
                st.error(f"執行錯誤: {e}")

# ==========================================
# TAB 2: 爬蟲 (僅下載)
# ==========================================
with tabs[1]:
    st.header("📥 爬蟲下載")
    uploaded_file = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_1")
    col1, col2 = st.columns(2)
    with col1:
        sheet_name = st.text_input("工作表", value="工作表1", key="sn_1")
    with col2:
        model_col_idx = st.number_input("型號欄位索引", value=0, key="mi_1")
        start_row = st.number_input("開始列", value=2, key="sr_1")
    
    if st.button("開始爬取", key="btn_1") and uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        models = []
        for idx, row in df.iterrows():
            if idx >= start_row - 1:
                if model_col_idx < len(row):
                    m = str(row.iloc[model_col_idx]).strip()
                    if re.match(r'^\d{7}$', m): models.append(m)
        
        res = []
        progress = st.progress(0)
        for i, m in enumerate(models):
            res.append(scrape_montbell_single(m))
            progress.progress((i+1)/len(models))
            time.sleep(0.5)
        
        df_res = pd.DataFrame(res)
        st.success("完成")
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w: df_res.to_excel(w, index=False)
        st.download_button("下載", out.getvalue(), "scraped.xlsx")

# ==========================================
# TAB 3: 翻譯 (僅翻譯)
# ==========================================
with tabs[2]:
    st.header("🈺 AI 翻譯")
    up_trans = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_2")
    if up_trans and api_key:
        df_t = pd.read_excel(up_trans)
        cols = st.multiselect("選擇翻譯欄位", df_t.columns)
        if st.button("開始翻譯", key="btn_2"):
            new_df = df_t.copy()
            prog = st.progress(0)
            total = len(df_t) * len(cols)
            curr = 0
            for c in cols:
                new_df[f"{c}_TW"] = ""
                for i, r in new_df.iterrows():
                    if pd.notna(r[c]):
                        new_df.at[i, f"{c}_TW"] = get_gemini_response(create_trans_prompt(str(r[c])), api_key, selected_model)
                    curr += 1
                    prog.progress(curr/total)
                    time.sleep(0.5)
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: new_df.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "translated.xlsx")

# ==========================================
# TAB 4: 優化 (僅優化)
# ==========================================
with tabs[3]:
    st.header("✨ 優化精簡")
    up_ref = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_3")
    if up_ref and api_key:
        df_r = pd.read_excel(up_ref)
        c_desc = st.selectbox("描述欄位", df_r.columns)
        c_spec = st.selectbox("規格欄位", ["(不處理)"] + list(df_r.columns))
        limit = st.slider("字數限制", 10, 200, 50)
        
        if st.button("開始優化", key="btn_3"):
            res_d, res_s = [], []
            prog = st.progress(0)
            for i, r in df_r.iterrows():
                if pd.notna(r[c_desc]):
                    res_d.append(get_gemini_response(create_refine_prompt(str(r[c_desc]), limit), api_key, selected_model))
                else: res_d.append("")
                
                if c_spec != "(不處理)" and pd.notna(r[c_spec]):
                    res_s.append(get_gemini_response(create_spec_prompt(str(r[c_spec])), api_key, selected_model))
                else: res_s.append("")
                prog.progress((i+1)/len(df_r))
                time.sleep(0.5)
            
            df_r['精簡_AI'] = res_d
            if c_spec != "(不處理)": df_r['規格_AI'] = res_s
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as w: df_r.to_excel(w, index=False)
            st.download_button("下載", out.getvalue(), "refined.xlsx")