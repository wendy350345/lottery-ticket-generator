import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import io
import os

st.set_page_config(page_title="萬用抽獎券生成器 V6", layout="wide")
st.title("🎟️ 萬用抽獎券生成器 V6 (解析度同步修正版)")

# --- 側邊欄設定 ---
with st.sidebar:
    st.header("⚙️ 樣式設定")
    font_mode = st.radio("字體來源", ["思源黑體", "上傳字體檔 (.ttf/.otf)"])
    uploaded_font = st.file_uploader("上傳字體檔案", type=["ttf", "ttc", "otf"]) if font_mode == "上傳字體檔 (.ttf/.otf)" else None
    
    fixed_text = st.text_input("固定標題", "2026 年度尾牙")
    fixed_size = st.slider("標題大小", 10, 300, 60) # 增大滑桿範圍適應高解析度
    fixed_y = st.slider("標題垂直位置 (%)", 0, 100, 15)
    
    data_size = st.slider("資料文字大小", 20, 300, 120)
    data_y = st.slider("資料垂直位置 (%)", 0, 100, 65)
    text_color = st.color_picker("文字顏色", "#000000")
    line_spacing = st.slider("行間距", 0, 100, 20)

# 工具函數：動態載入字體
def load_my_font(size):
    # 如果使用者有上傳字體，優先使用上傳的
    if font_mode == "上傳字體檔 (.ttf/.otf)" and uploaded_font is not None:
        return ImageFont.truetype(io.BytesIO(uploaded_font.getvalue()), size)
    
    # 否則使用專案資料夾內的思源黑體
    # 這裡請確認檔案名稱與你下載的一致
    local_font_path = "SOURCEHANSANSTC-REGULAR.otf" 
    
    if os.path.exists(local_font_path):
        return ImageFont.truetype(local_font_path, size)
    else:
        # 如果本機也沒有，才回退到微軟正黑體或預設字體
        try:
            return ImageFont.truetype("C:\\Windows\\Fonts\\msjh.ttc", size)
        except:
            st.warning("找不到思源黑體或系統字體，使用預設字體（中文可能亂碼）")
            return ImageFont.load_default()

# --- 檔案上傳 ---
col1, col2 = st.columns(2)
with col1:
    bg_file = st.file_uploader("1. 上傳背景底圖", type=["png", "jpg", "jpeg"])
with col2:
    data_file = st.file_uploader("2. 上傳 Excel 名單", type=["xlsx"])

if bg_file and data_file:
    df = pd.read_excel(data_file)
    cols = st.multiselect("請選擇要印出的 Excel 欄位", df.columns)
    
    # 預先計算 A4 格子尺寸，讓預覽與輸出基準一致
    orig_bg = Image.open(bg_file).convert("RGB")
    is_landscape = orig_bg.width > orig_bg.height
    A4_W, A4_H = (3508, 2480) if is_landscape else (2480, 3508)
    margin = 60
    t_w = (A4_W - 2 * margin) // 3
    t_h = (A4_H - 2 * margin) // 3

    # --- 預覽區域 ---
    st.subheader("👁️ 效果預覽 (以列印解析度為基準)")
    
    # 【關鍵】預覽時先將圖片 resize 到輸出的實際大小
    preview_ticket = orig_bg.resize((t_w, t_h), Image.LANCZOS)
    draw_preview = ImageDraw.Draw(preview_ticket)
    
    f_font = load_my_font(fixed_size)
    d_font = load_my_font(data_size)

    # 畫標題
    draw_preview.text((t_w/2, t_h * fixed_y / 100), fixed_text, font=f_font, fill=text_color, anchor="mm")
    # 畫第一筆資料
    if cols:
        sample_text = "\n".join([str(df.iloc[0][c]) for c in cols])
        draw_preview.multiline_text((t_w/2, t_h * data_y / 100), sample_text, font=d_font, fill=text_color, anchor="mm", align="center", spacing=line_spacing)
    
    st.image(preview_ticket, caption="預覽圖與輸出 PDF 的文字比例現已同步", use_container_width=True)

    # --- 批次生成 PDF ---
    if st.button("🚀 生成 A4 PDF", type="primary"):
        pages = []
        curr_page = Image.new('RGB', (A4_W, A4_H), 'white')
        draw_page = ImageDraw.Draw(curr_page)
        
        prog = st.progress(0)
        total_count = len(df)

        for i, (idx, row) in enumerate(df.iterrows()):
            # 每次製作一張小券，確保從乾淨的底圖 resize
            ticket = orig_bg.resize((t_w, t_h), Image.LANCZOS)
            t_draw = ImageDraw.Draw(ticket)
            
            # 畫文字
            t_draw.text((t_w/2, t_h * fixed_y / 100), fixed_text, font=f_font, fill=text_color, anchor="mm")
            row_txt = "\n".join([str(row[c]) for c in cols])
            t_draw.multiline_text((t_w/2, t_h * data_y / 100), row_txt, font=d_font, fill=text_color, anchor="mm", align="center", spacing=line_spacing)
            
            # 拼貼
            x = margin + (i % 3) * t_w
            y = margin + ((i // 3) % 3) * t_h
            curr_page.paste(ticket, (x, y))
            draw_page.rectangle([x, y, x + t_w, y + t_h], outline="#D3D3D3", width=1)
            
            if (i + 1) % 9 == 0 or (i + 1) == total_count:
                pages.append(curr_page)
                curr_page = Image.new('RGB', (A4_W, A4_H), 'white')
                draw_page = ImageDraw.Draw(curr_page)
            
            prog.progress((i + 1) / total_count)
            
        pdf_out = io.BytesIO()
        pages[0].save(pdf_out, format="PDF", save_all=True, append_images=pages[1:])
        st.success("✅ 完成！PDF 字體大小現在應該與預覽完全一致。")
        st.download_button("📥 下載 PDF", data=pdf_out.getvalue(), file_name="tickets_final.pdf")

