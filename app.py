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
    page_title="Montbell 自動化中心 v3.1",
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
            "temperature": 0.2, # 低溫度確保翻譯準確
            "top_p": 0.8,
            "top_k": 40,
            "max_output_tokens": 2048,
        }
        model = genai.GenerativeModel(model_name, generation_config=generation_config)
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        # 如果遇到 404 錯誤，嘗試給出更友善的提示
        if "404" in str(e):
            return f"Error: 模型名稱錯誤或不支援 ({model_name})。建議切換至 gemini-1.5-flash。"
        return f"Error: {str(e)}"

def scrape_montbell_single(model):
    """爬取單一商品邏輯 (回傳 dict)"""
    headers = {'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja-JP'}
    base_url = "https://webshop.montbell.jp/"
    search_url = "https://webshop.montbell.jp/goods/list_search.php?top_sk="
    
    info = {'型號': model, '商品名': '', '價格': '', '商品描述': '', '規格': '', '機能': '', '商品URL': ''}
    
    try:
        # 1. 直接訪問
        target_url = f"{base_url}goods/disp.php?product_id={model}"
        resp = requests.get(target_url, headers=headers, timeout=10)
        
        # 2. 搜尋備案
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
    原則：
    1. 專有名詞使用台灣戶外圈習慣用語 (如：透湿->透氣)。
    2. 語氣通順自然。
    3. 不要有任何解釋，直接輸出翻譯結果。
    原文：{text}
    """

def create_refine_prompt(text, limit):
    return f"""
    任務：提取商品核心賣點並精簡。
    限制：{limit}個中文字內。
    原文：{text}
    """

def create_spec_prompt(text):
    return f"""
    任務：優化並精簡產品規格表。
    規則：保留【】內標題，去除贅字，使用縮寫，保持換行格式。
    原文：{text}
    """

# ==========================================
# 2. 側邊欄：全域設定
# ==========================================
with st.sidebar:
    st.title("🛠️ 設定中心")
    st.info("👋 Hi Benjamin, v3.1 Fix")
    
    st.markdown("### 1. API 金鑰")
    api_key = st.text_input("Google Gemini API Key", type="password", placeholder="貼上 Key...")
    
    # 新增：API 檢測按鈕
    col_test, col_status = st.columns([1, 2])
    with col_test:
        test_btn = st.button("測試連線")
    
    if test_btn and api_key:
        try:
            genai.configure(api_key=api_key)
            # [FIX] 這裡強制使用最穩定的 Flash 模型進行測試，避免 gemini-pro 404 錯誤
            m = genai.GenerativeModel("gemini-1.5-flash")
            response = m.generate_content("Test connection")
            st.sidebar.success("✅ API 連線成功！")
        except Exception as e:
            st.sidebar.error(f"❌ 連線失敗: {e}")

    st.markdown("### 2. 模型選擇")
    # [FIX] 移除了舊版 gemini-pro，改用明確版本號
    model_options = ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-1.0-pro"]
    selected_model = st.selectbox("AI 模型", model_options, index=0, help="Flash最快(推薦)，Pro品質較好")
    
    st.markdown("---")
    st.caption("Design for Montbell Workflow")

# ==========================================
# 3. 主畫面
# ==========================================
st.title("🏔️ Montbell 自動化中心 v3.1")

tabs = st.tabs(["⚡ 一鍵全自動 (All-in-One)", "📥 分步：爬蟲", "🈺 分步：翻譯", "✨ 分步：優化"])

# ==========================================
# TAB 1: 一鍵全自動 (All-in-One)
# ==========================================
with tabs[0]:
    st.header("⚡ 一鍵全自動處理流程")
    st.caption("上傳型號表 -> 系統自動：1.爬取官網 -> 2.翻譯成中文 -> 3.精簡優化 -> 輸出最終檔。")
    
    col_in, col_set = st.columns([1, 1])
    with col_in:
        uploaded_file_all = st.file_uploader("上傳型號 Excel", type=["xlsx", "xls"], key="up_all")
    with col_set:
        with st.expander("參數設定 (點擊展開)", expanded=True):
            sheet_name_all = st.text_input("工作表名稱", value="工作表1", key="sn_all")
            model_col_idx_all = st.number_input("型號欄位索引 (A=0, B=1...)", value=0, min_value=0, key="mi_all")
            char_limit_all = st.number_input("描述精簡字數限制", value=50, min_value=10, key="cl_all")
            
    if st.button("🚀 啟動全自動排程", type="primary", key="btn_all"):
        if not uploaded_file_all or not api_key:
            st.error("請檢查：1.是否已上傳檔案 2.是否已輸入 API Key")
        else:
            try:
                # 讀取 Excel
                df = pd.read_excel(uploaded_file_all, sheet_name=sheet_name_all)
                models = []
                for idx, row in df.iterrows():
                    if idx >= 1: # 假設 Header 後一行開始
                        if model_col_idx_all < len(row):
                            m = str(row.iloc[model_col_idx_all]).strip()
                            if re.match(r'^\d{7}$', m): models.append(m)
                
                if not models:
                    st.error("找不到有效型號 (7碼數字)。")
                else:
                    results_final = []
                    
                    # 使用 st.status 顯示複合進度
                    with st.status(f"正在處理 {len(models)} 筆商品 (爬蟲+翻譯+優化)...", expanded=True) as status:
                        prog_bar = st.progress(0)
                        
                        for i, model in enumerate(models):
                            status.update(label=f"[{i+1}/{len(models)}] 處理型號：{model} ...")
                            
                            # 1. 爬蟲
                            raw_data = scrape_montbell_single(model)
                            
                            # 2. 翻譯 (針對主要欄位)
                            trans_data = raw_data.copy()
                            if raw_data['商品名'] != '未找到':
                                trans_data['商品名_TW'] = get_gemini_response(create_trans_prompt(raw_data['商品名']), api_key, selected_model)
                                trans_data['商品描述_TW'] = get_gemini_response(create_trans_prompt(raw_data['商品描述']), api_key, selected_model)
                                trans_data['規格_TW'] = get_gemini_response(create_trans_prompt(raw_data['規格']), api_key, selected_model)
                                trans_data['機能_TW'] = get_gemini_response(create_trans_prompt(raw_data['機能']), api_key, selected_model)
                            else:
                                trans_data['商品名_TW'] = "查無資料"
                            
                            # 3. 優化 (精簡)
                            if raw_data['商品名'] != '未找到':
                                trans_data['精簡描述_AI'] = get_gemini_response(create_refine_prompt(trans_data['商品描述_TW'], char_limit_all), api_key, selected_model)
                                trans_data['規格_結構化_AI'] = get_gemini_response(create_spec_prompt(trans_data['規格_TW']), api_key, selected_model)
                            else:
                                trans_data['精簡描述_AI'] = ""
                                trans_data['規格_結構化_AI'] = ""

                            results_final.append(trans_data)
                            prog_bar.progress((i+1)/len(models))
                            time.sleep(1) # 避免 API 過熱
                        
                        status.update(label="✅ 全自動流程執行完畢！", state="complete", expanded=False)

                    # 輸出
                    df_final = pd.DataFrame(results_final)
                    st.success(f"完成！共產出 {len(df_final)} 筆資料。")
                    
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False)
                    st.download_button("📥 下載最終完整報表", out.getvalue(), "montbell_full_auto.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

            except Exception as e:
                st.error(f"執行錯誤: {e}")

# ==========================================
# TAB 2: 爬蟲 (Scraper)
# ==========================================
with tabs[1]:
    st.header("📥 步驟一：官網爬蟲 (僅下載)")
    uploaded_file = st.file_uploader("上傳 Excel", type=["xlsx", "xls"], key="up_1")
    col1, col2 = st.columns(2)
    with col1:
        sheet_name = st.text_input("工作表", value="工作表1", key="sn_1")
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
        st.success("爬取完成")
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as w: df_res.to_excel(w, index=False)
        st.download_button("下載 Excel", out.getvalue(), "scraped.xlsx")

# ==========================================
# TAB 3: 翻譯 (Translator)
# ==========================================
with tabs[2]:
    st.header("🈺 步驟二：AI 翻譯 (僅翻譯)")
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
            st.download_button("下載翻譯檔", out.getvalue(), "translated.xlsx")

# ==========================================
# TAB 4: 優化 (Refiner)
# ==========================================
with tabs[3]:
    st.header("✨ 步驟三：優化精簡 (僅優化)")
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
            st.download_button("下載優化檔", out.getvalue(), "refined.xlsx")