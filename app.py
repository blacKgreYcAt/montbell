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
    page_title="Montbell 商品資料自動化中心",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定義 CSS 優化視覺 (隱藏預設 Footer，優化按鈕樣式)
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
# 1. 核心邏輯函式庫 (Backend Logic)
# ==========================================

def get_gemini_response(prompt, api_key, model_name="gemini-1.5-flash"):
    """呼叫 Gemini API 的通用函式"""
    if not api_key:
        return "Error: 請先輸入 API Key"
    try:
        genai.configure(api_key=api_key)
        generation_config = {
            "temperature": 0.2,
            "top_p": 0.8,
            "top_k": 40,
            "max_output_tokens": 2048,
        }
        model = genai.GenerativeModel(model_name, generation_config=generation_config)
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        return f"Error: {str(e)}"

# ==========================================
# 2. 側邊欄：全域設定
# ==========================================
with st.sidebar:
    st.title("🛠️ 設定中心")
    st.info("👋 嗨 Benjamin，歡迎回來！")
    
    st.markdown("### 🔑 API 金鑰設定")
    api_key = st.text_input("Google Gemini API Key", type="password", placeholder="貼上您的 Key...")
    
    if api_key:
        st.success("API Key 已載入")
    else:
        st.warning("請輸入 Key 以啟用 AI 功能")
        
    st.markdown("---")
    st.markdown("### ℹ️ 關於工具")
    st.caption("此工具由 Python 驅動，整合了爬蟲與 Gemini AI，專為 Montbell 資料處理設計。")
    st.caption("v2.0 - UI Optimized")

# ==========================================
# 3. 主畫面：分頁導航
# ==========================================
st.title("🏔️ Montbell 商品資料自動化中心")
st.markdown("請依序執行以下步驟，完成資料的 **獲取**、**在地化** 與 **優化**。")

# 使用 Tabs 取代 Radio Button，視覺更現代
tab1, tab2, tab3 = st.tabs(["📥 步驟一：官網爬蟲", "🈺 步驟二：AI 翻譯 (TW)", "✨ 步驟三：資料優化"])

# ==========================================
# TAB 1: 爬蟲 (Scraper)
# ==========================================
with tab1:
    st.header("Montbell 日本官網資料下載")
    st.caption("上傳包含「商品型號」的 Excel，系統將自動從官網抓取圖片、價格與規格。")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.markdown("#### 1. 檔案設定")
        uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx", "xls"], key="uploader_1")
        
        with st.expander("進階參數設定", expanded=False):
            sheet_name = st.text_input("工作表名稱", value="工作表1")
            start_row = st.number_input("資料開始列 (Header後一行)", value=2, min_value=1)
            model_col_idx = st.number_input("型號欄位索引 (A=0, B=1...)", value=0, min_value=0)
            
    with col2:
        st.markdown("#### 2. 執行面板")
        if uploaded_file:
            # 預覽檔案
            try:
                df_preview = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                st.dataframe(df_preview.head(3), use_container_width=True)
                st.caption(f"預覽前 3 筆資料。將從第 {start_row} 列開始讀取，型號位於第 {model_col_idx} 欄。")
                
                if st.button("🚀 開始爬取資料", type="primary", key="btn_scrape"):
                    # 讀取並過濾型號
                    real_start_row = start_row - 1
                    models = []
                    for index, row in df_preview.iterrows():
                        if index < real_start_row: continue
                        if model_col_idx < len(row):
                            model = str(row.iloc[model_col_idx]).strip()
                            if re.match(r'^\d{7}$', model): models.append(model)
                    
                    if not models:
                        st.error("未找到符合格式 (7碼數字) 的型號，請檢查設定。")
                    else:
                        # 使用 st.status 顯示進度，介面更乾淨
                        results = []
                        with st.status(f"正在爬取 {len(models)} 筆商品...", expanded=True) as status:
                            progress_bar = st.progress(0)
                            
                            # 爬蟲設定
                            headers = {'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'ja-JP'}
                            base_url = "https://webshop.montbell.jp/"
                            
                            for i, model in enumerate(models):
                                status.update(label=f"正在處理 ({i+1}/{len(models)}): {model}")
                                progress_bar.progress((i + 1) / len(models))
                                
                                product_info = {'型號': model, '商品名': '未找到', '價格': '', '商品描述': '', '規格': '', '機能': ''}
                                try:
                                    # 簡化的爬蟲邏輯 (為節省篇幅，核心邏輯與前版相同)
                                    target_url = f"{base_url}goods/disp.php?product_id={model}"
                                    resp = requests.get(target_url, headers=headers, timeout=10)
                                    if resp.status_code == 200:
                                        soup = BeautifulSoup(resp.text, 'html.parser')
                                        product_info['商品URL'] = target_url
                                        
                                        name = soup.select_one('h1.goods-detail__ttl-main, h1')
                                        if name: product_info['商品名'] = name.text.strip()
                                        
                                        price = soup.select_one('.goods-detail__price, span.selling_price')
                                        if price: product_info['價格'] = price.text.strip()
                                        
                                        desc = soup.select('.column1.type01 .innerCont p')
                                        if desc: product_info['商品描述'] = desc[0].text.strip()
                                        
                                        spec = soup.select_one('div.explanationBox')
                                        if spec: product_info['規格'] = spec.text.strip()

                                except Exception as e:
                                    st.write(f"Error: {model} - {e}")
                                
                                results.append(product_info)
                                time.sleep(1) # 禮貌性延遲
                                
                            status.update(label="✅ 爬取完成！", state="complete", expanded=False)
                        
                        # 結果處理
                        result_df = pd.DataFrame(results)
                        st.success(f"成功獲取 {len(result_df)} 筆資料！")
                        
                        # 下載按鈕
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            result_df.to_excel(writer, index=False)
                        
                        st.download_button(
                            label="📥 下載爬取結果 (Excel)",
                            data=output.getvalue(),
                            file_name="montbell_data_scraped.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"讀取 Excel 失敗: {e}")
        else:
            st.info("請先上傳 Excel 檔案以開始操作。")

# ==========================================
# TAB 2: 翻譯 (Translator)
# ==========================================
with tab2:
    st.header("AI 智能翻譯 (日 -> 繁中)")
    st.caption("透過 Gemini AI，將日文資料轉換為符合台灣戶外市場用語的在地化內容。")
    
    if not api_key:
        st.error("⚠️ 請先在左側邊欄輸入 API Key 才能使用此功能。")
    else:
        uploaded_file_trans = st.file_uploader("上傳檔案 (通常是步驟一的結果)", type=["xlsx", "xls"], key="uploader_2")
        
        if uploaded_file_trans:
            df_trans = pd.read_excel(uploaded_file_trans)
            
            col_config, col_action = st.columns([1, 2])
            
            with col_config:
                st.markdown("#### 1. 欄位選擇")
                cols_to_translate = st.multiselect(
                    "選擇需要翻譯的欄位", 
                    df_trans.columns,
                    default=[c for c in df_trans.columns if c in ['商品名', '商品描述', '規格', '機能']]
                )
                st.info("💡 提示：AI 將會扮演「專業戶外譯者」的角色進行翻譯。")

            with col_action:
                st.markdown("#### 2. 預覽與執行")
                st.dataframe(df_trans.head(3), use_container_width=True)
                
                if st.button("🌏 開始 AI 翻譯", type="primary", key="btn_trans"):
                    if not cols_to_translate:
                        st.warning("請至少選擇一個欄位。")
                    else:
                        new_df = df_trans.copy()
                        total_steps = len(df_trans) * len(cols_to_translate)
                        current_step = 0
                        
                        with st.status("正在進行 AI 翻譯...", expanded=True) as status:
                            progress_bar = st.progress(0)
                            
                            for col in cols_to_translate:
                                new_col_name = f"{col}_TW"
                                new_df[new_col_name] = ""
                                
                                for idx, row in new_df.iterrows():
                                    original_text = str(row[col])
                                    if pd.notna(row[col]) and original_text.strip() != "":
                                        status.update(label=f"翻譯中: [{col}] 第 {idx+1} 筆...")
                                        
                                        # 專業 Persona Prompt
                                        prompt = f"""
                                        角色設定：你是一位翻譯經驗豐富的專業譯者，對於戶外商品的機能名詞十分熟悉，同時對於社群行銷的用字也很了解，能夠將日文資料翻譯為符合台灣市場需求的內容。
                                        任務：請將以下的日文商品資料翻譯成繁體中文 (台灣)。
                                        翻譯原則：
                                        1. 專有名詞請使用台灣戶外圈習慣的用語 (例如：透湿 -> 透氣)。
                                        2. 語氣要通順自然，適合閱讀，避免生硬的直譯。
                                        3. 嚴格禁止自我指涉，直接輸出翻譯內容。
                                        原文：{original_text}
                                        """
                                        
                                        trans_text = get_gemini_response(prompt, api_key)
                                        new_df.at[idx, new_col_name] = trans_text
                                        time.sleep(0.5)
                                    
                                    current_step += 1
                                    progress_bar.progress(current_step / total_steps)
                                    
                            status.update(label="✅ 翻譯作業完成！", state="complete", expanded=False)
                        
                        st.success("翻譯成功！")
                        output_trans = io.BytesIO()
                        with pd.ExcelWriter(output_trans, engine='openpyxl') as writer:
                            new_df.to_excel(writer, index=False)
                            
                        st.download_button(
                            label="📥 下載翻譯結果 (Excel)",
                            data=output_trans.getvalue(),
                            file_name="montbell_data_translated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

# ==========================================
# TAB 3: 優化 (Refiner)
# ==========================================
with tab3:
    st.header("資料精簡與結構化")
    st.caption("將翻譯後的長篇大論，轉化為適合電商上架的精簡賣點與規格表。")

    if not api_key:
        st.error("⚠️ 請先在左側邊欄輸入 API Key。")
    else:
        uploaded_file_refine = st.file_uploader("上傳檔案 (通常是步驟二的結果)", type=["xlsx", "xls"], key="uploader_3")
        
        if uploaded_file_refine:
            df_refine = pd.read_excel(uploaded_file_refine)
            
            # 版面配置：左側設定，右側說明
            c1, c2 = st.columns([1, 1])
            
            with c1:
                st.subheader("參數設定")
                col_desc = st.selectbox("選擇【商品描述】來源欄位", df_refine.columns, index=len(df_refine.columns)-1 if '商品描述_TW' in df_refine.columns else 0)
                col_spec = st.selectbox("選擇【規格】來源欄位 (選填)", ["(不處理)"] + list(df_refine.columns))
                
                st.markdown("---")
                char_limit = st.slider("商品描述字數限制", min_value=30, max_value=200, value=50, step=10)
                refine_specs_opt = st.toggle("啟用規格 AI 結構化 (整理為 Key-Value 格式)", value=True)
                
            with c2:
                st.subheader("操作說明")
                st.markdown("""
                此步驟將執行以下優化：
                * **描述精簡**：提取核心賣點，去除贅字，符合字數限制。
                * **規格結構化**：將雜亂的規格文字整理成易讀的列表 (如啟用)。
                """)
                st.warning("注意：此步驟會消耗較多 Token，請耐心等待。")

            st.markdown("---")
            if st.button("✨ 開始資料優化", type="primary", key="btn_refine"):
                with st.status("AI 正在施展魔法...", expanded=True) as status:
                    progress = st.progress(0)
                    results_desc = []
                    results_spec = []
                    total = len(df_refine)
                    
                    for idx, row in df_refine.iterrows():
                        status.update(label=f"正在優化第 {idx+1}/{total} 筆...")
                        progress.progress((idx+1)/total)
                        
                        # 1. 描述
                        if pd.notna(row[col_desc]):
                            p_desc = f"提取商品核心賣點並精簡至{char_limit}字內。原文：{str(row[col_desc])}"
                            results_desc.append(get_gemini_response(p_desc, api_key))
                        else:
                            results_desc.append("")
                            
                        # 2. 規格
                        if col_spec != "(不處理)" and refine_specs_opt and pd.notna(row[col_spec]):
                            p_spec = f"優化產品規格表，保留【】標題，去除贅字，使用縮寫。原文：{str(row[col_spec])}"
                            results_spec.append(get_gemini_response(p_spec, api_key))
                        elif col_spec != "(不處理)":
                            results_spec.append(row[col_spec])
                        else:
                            results_spec.append("")
                            
                        time.sleep(0.5)
                    
                    status.update(label="✨ 優化完成！", state="complete", expanded=False)

                # 寫入與下載
                df_refine['精簡描述_AI'] = results_desc
                if col_spec != "(不處理)":
                    df_refine['規格_結構化_AI'] = results_spec
                
                st.success("所有資料處理完畢！")
                
                # 顯示 Before / After 比較 (取第一筆範例)
                with st.expander("👀 查看優化前後對比 (範例)", expanded=True):
                    c_a, c_b = st.columns(2)
                    with c_a:
                        st.markdown("**處理前 (描述)**")
                        st.text(str(df_refine.iloc[0][col_desc])[:100] + "...")
                    with c_b:
                        st.markdown(f"**處理後 (精簡 {char_limit} 字)**")
                        st.success(results_desc[0])

                output_final = io.BytesIO()
                with pd.ExcelWriter(output_final, engine='openpyxl') as writer:
                    df_refine.to_excel(writer, index=False)
                    
                st.download_button(
                    label="📥 下載最終成品 (Excel)",
                    data=output_final.getvalue(),
                    file_name="montbell_final_optimized.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )