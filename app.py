import streamlit as st
import pandas as pd
import google.generativeai as genai
from google.generativeai import types
import os
import re
from PIL import Image

# Load environment variables
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env_path):
        try:
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" in line:
                        k, v = line.split("=", 1)
                        v = v.strip().strip('"').strip("'")
                        os.environ.setdefault(k.strip(), v)
        except Exception:
            pass

def generate_content_with_retries(model, contents, config, max_retries=1):
    try:
        return model.generate_content(contents=contents, generation_config=config)
    except Exception as e:
        msg = str(e)
        lower = msg.lower()
        if 'resource_exhausted' in lower or 'quota' in lower or '429' in lower or 'rate limit' in lower:
            retry_message = ''
            m = re.search(r'please retry in\s*([0-9]+(?:\.[0-9]+)?)s', msg, re.IGNORECASE)
            if m:
                retry_message = f" Vui lòng thử lại sau khoảng {m.group(1)} giây."
            raise RuntimeError(f"API Gemini đang hết quota. {retry_message} Nếu vẫn gặp lỗi, kiểm tra lại tài khoản/billing của bạn.\n{msg}") from e
        raise

st.set_page_config(
    page_title="Dược Hậu Giang - DHG Pharma Advisor",
    layout="wide",
    page_icon="🌿"
)

# ========== CSS: Ẩn sidebar, hiển thị khi hover ==========
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Quicksand:wght@500;600;700&display=swap');

    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif;
        background: linear-gradient(135deg, #f0fdfa 0%, #e6fffa 100%);
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Quicksand', sans-serif;
        font-weight: 600;
        color: #0f766e;
    }

    /* Ẩn sidebar mặc định */
    [data-testid="stSidebar"] {
        width: 0px;
        min-width: 0px;
        transition: all 0.3s ease;
        overflow: hidden;
    }

    [data-testid="stSidebar"]:hover {
        width: 280px;
        min-width: 280px;
    }

    [data-testid="stSidebar"] > div {
        width: 280px;
    }

    .stButton > button {
        background: linear-gradient(135deg, #14b8a6 0%, #0d9488 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 10px 20px;
        font-weight: 600;
        font-family: 'Inter', sans-serif;
        box-shadow: 0 4px 12px rgba(20, 184, 166, 0.3);
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        box-shadow: 0 6px 20px rgba(20, 184, 166, 0.5);
        transform: translateY(-1px);
    }

    [data-testid="stChatMessage"] {
        border-radius: 18px;
        padding: 15px;
        margin-bottom: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        font-family: 'Inter', sans-serif;
    }
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarAssistant"]) {
        background-color: #f0fdfa;
        border: 1px solid #99f6e4;
    }
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {
        background: linear-gradient(135deg, #0f766e 0%, #14b8a6 100%) !important;
        border: 1px solid #0d9488;
    }
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) p,
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) li,
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) span,
    [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) div {
        color: #ffffff !important;
        font-weight: 500 !important;
    }

    .product-card {
        background: white;
        border-radius: 20px;
        padding: 16px;
        border: 1px solid #e6fffa;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s;
        margin-bottom: 12px;
    }
    .product-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 25px rgba(20, 184, 166, 0.15);
    }

    [data-testid="stChatInput"] {
        border: 2px solid #14b8a6 !important;
        border-radius: 16px !important;
    }

    .streamlit-expanderHeader {
        background: #f0fdfa;
        border-radius: 12px;
        color: #0f766e;
        font-weight: 600;
    }

    .promo-card {
        background: linear-gradient(135deg, #ccfbf1, #99f6e4);
        border-radius: 20px;
        padding: 30px 20px;
        text-align: center;
        color: #0f766e;
        box-shadow: 0 6px 20px rgba(20, 184, 166, 0.15);
        transition: transform 0.2s;
    }
    .promo-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 25px rgba(20, 184, 166, 0.3);
    }
</style>
""", unsafe_allow_html=True)

# Hàm tính tiền
def parse_price(price_str):
    try:
        if pd.isna(price_str) or not price_str:
            return 0
        cleaned = re.sub(r'[^\d]', '', str(price_str))
        return int(cleaned) if cleaned else 0
    except:
        return 0

# Load dữ liệu
def load_data():
    for fname in ["data_san_pham.xlsx", "data san pham.xlsx", "data_san_pham.xlsx"]:
        if os.path.exists(fname):
            try:
                return pd.read_excel(fname)
            except Exception as e:
                st.error(f"⚠️ Lỗi đọc file '{fname}': {e}")
                return pd.DataFrame()
    st.error("⚠️ Không tìm thấy file dữ liệu sản phẩm.")
    return pd.DataFrame(columns=['Tên sản phẩm', 'Phân loại', 'Chuyên mục ', 'Xuất xứ ', 'Giá tiền', 'Công dụng'])

df_products = load_data()

# Session state
if "messages" not in st.session_state:
    st.session_state.messages = []
if "cart" not in st.session_state:
    st.session_state.cart = []
if "recommended_products" not in st.session_state:
    st.session_state.recommended_products = []

def add_product_to_cart(product_names: list[str]):
    added_items = []
    if df_products.empty or 'Tên sản phẩm' not in df_products.columns:
        return "Kho hàng trống."
    for p_name_input in product_names:
        match = df_products[df_products['Tên sản phẩm'].str.lower().str.contains(p_name_input.lower(), na=False)]
        if not match.empty:
            actual_name = match.iloc[0]['Tên sản phẩm']
            p_price = match.iloc[0].get('Giá tiền', 'Liên hệ')
            st.session_state.cart.append({"name": actual_name, "price": p_price})
            added_items.append(actual_name)
    if added_items:
        return f"Đã tự động thêm thành công: {', '.join(added_items)} vào giỏ hàng."
    return "Không tìm thấy sản phẩm phù hợp trong kho để tự động thêm."

# ========== SIDEBAR (ẩn, hover hiện) ==========
with st.sidebar:
    st.markdown("""
        <div style="text-align: center; padding-bottom: 10px;">
            <h2 style="color: #0f766e; margin-bottom: 0; font-family: 'Quicksand', sans-serif;">
                Dược Hậu Giang
            </h2>
            <p style="color: #0d9488; font-size: 0.85rem; margin-top: 0; font-weight: 500;">
                DHG Pharma Advisor
            </p>
            <hr style="border: 1px solid #ccfbf1; margin: 15px 0;">
        </div>
    """, unsafe_allow_html=True)

    page = st.radio(
        "Danh Mục Chính",
        ["🏠 Trang Chủ", "🛒 Giỏ Hàng", "🧴 Sản Phẩm", "💬 Tư Vấn AI"],
        label_visibility="collapsed"
    )

    st.divider()
    st.markdown("### ⚙️ Quản Lý")
    if st.button("🗑️ Xóa Lịch Sử Chat", use_container_width=True):
        st.session_state.messages = []
        st.session_state.recommended_products = []
        st.toast("🧹 Đã làm sạch hội thoại!")
        st.rerun()

# ========== TAB 1: TRANG CHỦ ==========
if page == "🏠 Trang Chủ":
    st.image("https://via.placeholder.com/1200x400/f0fdfa/0f766e?text=DƯỢC+HẬU+GIANG+-+DHG+PHARMA+ADVISOR",
             use_container_width=True)

    st.markdown("""
        <div style="text-align: center; padding: 20px 0 10px 0;">
            <h1 style="color: #0f766e; font-size: 3rem; font-weight: 700; margin-bottom: 5px;">
                Dược Hậu Giang
            </h1>
            <p style="color: #0d9488; font-size: 1.2rem; font-weight: 500;">
                DHG Pharma Advisor – Giải pháp chăm sóc da khoa học
            </p>
            <hr style="width: 50%; border: 2px solid #ccfbf1; margin: 20px auto;">
        </div>
    """, unsafe_allow_html=True)

    st.subheader("🔥 Khuyến mãi hôm nay")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="promo-card">
            <div style="font-size: 2.5rem; margin-bottom: 10px;">🌞</div>
            <h3 style="margin: 0 0 10px 0; color: #0f766e;">GIẢM 30%</h3>
            <p style="font-weight: 500; margin: 0;">Kem Chống Nắng</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="promo-card">
            <div style="font-size: 2.5rem; margin-bottom: 10px;">🧼</div>
            <h3 style="margin: 0 0 10px 0; color: #0f766e;">MUA 2 TẶNG 1</h3>
            <p style="font-weight: 500; margin: 0;">Sữa Rửa Mặt</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="promo-card">
            <div style="font-size: 2.5rem; margin-bottom: 10px;">🚚</div>
            <h3 style="margin: 0 0 10px 0; color: #0f766e;">FREESHIP</h3>
            <p style="font-weight: 500; margin: 0;">Đơn từ 500k</p>
        </div>
        """, unsafe_allow_html=True)

    st.divider()
    st.info("💡 Mẹo làm đẹp: Uống đủ 2 lít nước mỗi ngày và luôn dùng kem chống nắng để bảo vệ làn da tối ưu.")

# ========== TAB 2: GIỎ HÀNG ==========
elif page == "🛒 Giỏ Hàng":
    st.markdown("<h2 style='color:#0f766e;'>🛒 Giỏ Hàng Của Bạn</h2>", unsafe_allow_html=True)

    if not st.session_state.cart:
        st.info("🛒 Giỏ hàng hiện đang trống. Hãy vào tab Sản Phẩm hoặc chat với AI để thêm sản phẩm nhé!")
    else:
        total_price = 0
        total_items = len(st.session_state.cart)
        for idx, item in enumerate(st.session_state.cart):
            col_item, col_del = st.columns([4, 1])
            col_item.write(f"• **{item['name']}**\n_{item['price']}_")
            total_price += parse_price(item['price'])
            if col_del.button("❌", key=f"del_cart_{idx}"):
                st.session_state.cart.pop(idx)
                st.rerun()

        st.divider()
        st.markdown(f"📦 **Tổng số lượng:** {total_items} sản phẩm")
        st.markdown(f"💰 **Tổng tiền tạm tính:** <span style='color:#0f766e; font-weight:bold; font-size:1.3rem;'>{total_price:,}đ</span>", unsafe_allow_html=True)
        st.divider()
        if st.button("🔥 Gửi Đơn Đặt Hàng", use_container_width=True):
            st.balloons()
            st.success("✅ Đơn hàng đã được chuyển tới hệ thống xử lý!")
            st.session_state.cart = []
            st.rerun()

# ========== TAB 3: SẢN PHẨM ==========
elif page == "🧴 Sản Phẩm":
    st.markdown("<h2 style='color:#0f766e;'>🧴 Danh Mục Sản Phẩm</h2>", unsafe_allow_html=True)

    if df_products.empty:
        st.warning("📭 Chưa có dữ liệu sản phẩm.")
    else:
        CHUYEN_MUC_LIST = ["Tất cả", "Mỹ Phẩm", "Thuốc da liễu", "Thuốc Trị Mụn", "Chăm Sóc Cơ Thể", "Thuốc Trị Sẹo", "Thuốc Da Liễu"]
        selected_cat = st.selectbox("🔎 Lọc theo chuyên mục:", CHUYEN_MUC_LIST)

        chuyen_muc_col = 'Chuyên mục ' if 'Chuyên mục ' in df_products.columns else ('Chuyên mục' if 'Chuyên mục' in df_products.columns else None)
        if chuyen_muc_col and selected_cat != "Tất cả":
            df_filtered = df_products[df_products[chuyen_muc_col].str.strip().str.lower() == selected_cat.strip().lower()]
        else:
            df_filtered = df_products
        if df_filtered.empty: df_filtered = df_products

        cols = st.columns(3)
        col_index = 0
        for index, row in df_filtered.iterrows():
            with cols[col_index % 3]:
                p_name = row.get('Tên sản phẩm', f'Mỹ phẩm #{index+1}')
                p_brand = row.get('Xuất xứ ', row.get('Xuất xứ', 'Chính hãng'))
                p_price = row.get('Giá tiền', 'Liên hệ')
                p_effect = row.get('Công dụng', 'Sản phẩm chăm sóc da')
                p_image = row.get('Hình ảnh', None)
                if pd.notna(p_image) and str(p_image).strip() != "" and str(p_image).lower() != "nan":
                    p_image = str(p_image).replace('\\', '/').strip()
                else:
                    p_image = None

                st.markdown(f"""
                    <div class="product-card">
                        <span style="background:#ccfbf1; color:#0f766e; font-size:0.75rem; font-weight:600; padding:3px 10px; border-radius:20px; display:inline-block; margin-bottom:8px;">
                            {p_brand}
                        </span>
                        <h4 style="margin:0; color:#0f766e; font-size:1.1rem; font-weight:600;">{p_name}</h4>
                    </div>
                """, unsafe_allow_html=True)

                if p_image and os.path.exists(p_image):
                    try:
                        st.image(Image.open(p_image), use_container_width=True)
                    except:
                        st.caption(f"⚠️ Lỗi ảnh: {p_image}")
                else:
                    st.caption("📸 Chưa có ảnh minh họa")

                st.markdown(f"""
                    <div class="product-card" style="margin-top:0; border-top:none; border-top-left-radius:0; border-top-right-radius:0; padding-top:5px;">
                        <p style="font-size:0.85rem; color:#115e59; margin:5px 0 8px 0; height:3em; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;">
                            {p_effect}
                        </p>
                        <h5 style="color:#0f766e; margin:0 0 10px 0; font-weight:700; font-size:1.2rem;">💰 {p_price}</h5>
                    </div>
                """, unsafe_allow_html=True)

                if st.button(f"🛒 Thêm vào giỏ hàng", key=f"prod_btn_{index}"):
                    st.session_state.cart.append({"name": p_name, "price": p_price})
                    st.toast(f"✅ Đã thêm {p_name} vào giỏ hàng!")
                    st.rerun()
            col_index += 1

# ========== TAB 4: TƯ VẤN AI ==========
elif page == "💬 Tư Vấn AI":
    st.markdown("<h2 style='color:#0f766e;'>💬 Tư Vấn Cùng Chuyên Gia AI</h2>", unsafe_allow_html=True)

    with st.expander("⚙️ Thông Tin Cá Nhân (Tùy Chỉnh)", expanded=False):
        user_name = st.text_input("Họ và tên:", value="Khách hàng")
        user_age = st.number_input("Tuổi:", min_value=1, max_value=100, value=20)
        user_skin_type = st.selectbox("Loại da hiện tại:", ["Chưa xác định loại da", "Da dầu mụn", "Da khô nhạy cảm", "Da hỗn hợp", "Da thường"])

    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if prompt := st.chat_input("Hãy mô tả rõ hơn về tình trạng da hiện tại của bạn..."):
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})

    GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY")

    if not GEMINI_API_KEY:
     st.error("❌ Chưa cấu hình GEMINI_API_KEY trong Streamlit Secrets.")
    st.stop()

    knowledge_base = df_products.to_string(index=False) if not df_products.empty else "Không có dữ liệu."

    AGREE_KEYWORDS = [
            "thêm vào giỏ", "cho vào giỏ", "bỏ vào giỏ", "thêm vào",
            "thêm đi", "lấy đi", "lấy hết", "mua đi", "mua hết",
            "chốt đi", "chốt luôn", "lấy cho chị", "lấy cho mình",
            "lấy cho t", "thêm cho t", "add đi", "lấy luôn", "mua luôn",
            "cho chị", "cho mình", "thêm luôn", "ok thêm", "ừ thêm",
            "thêm cái đó", "thêm mấy cái", "thêm sp", "thêm sản phẩm",
            "lấy sp", "muốn mua"
        ]

    is_agreeing = False
    if len(prompt.strip()) <= 30:
            is_agreeing = any(word in prompt.lower() for word in AGREE_KEYWORDS)

    if not is_agreeing:
            st.session_state.recommended_products = []

    if is_agreeing and st.session_state.recommended_products:
            with st.chat_message("assistant"):
                with st.spinner("Đang thêm vào giỏ hàng..."):
                    result_msg = add_product_to_cart(st.session_state.recommended_products)
                    answer = f"{result_msg} Bạn kiểm tra giỏ hàng nhé! 🛒"
                    st.markdown(answer)
                    st.session_state.messages.append({"role": "assistant", "content": answer})
                    st.session_state.recommended_products = []
            st.rerun()

    elif is_agreeing and not st.session_state.recommended_products:
            with st.chat_message("assistant"):
                answer = "Mình chưa xác định được sản phẩm nào. Bạn hãy mô tả tình trạng da để mình tư vấn nhé!"
                st.markdown(answer)
                st.session_state.messages.append({"role": "assistant", "content": answer})
            st.rerun()

    else:
            prompt_he_thong = f"""
            Bạn là chuyên gia tư vấn Skincare chuyên nghiệp. Khách tên {user_name}, {user_age} tuổi, loại da: {user_skin_type}.
            {"Khách chưa biết loại da của mình - hãy hỏi thêm về tình trạng da (bóng dầu, khô, mụn...) để tư vấn phù hợp, hoặc gợi ý sản phẩm phù hợp cho nhiều loại da." if user_skin_type == "Chưa xác định loại da" else ""}
            DANH SÁCH SẢN PHẨM TRONG KHO:
            {knowledge_base}
            
            NHIỆM VỤ:
            1. Tư vấn routine khoa học ngắn gọn bằng sản phẩm có sẵn ở trên. 
            2. Cuối câu hỏi KHÁCH CÓ HÀI LÒNG VÀ MUỐN THÊM CÁC SẢN PHẨM NÀY VÀO GIỎ HÀNG KHÔNG.
            3. Xuất ra một dòng cuối cùng ở định dạng chính xác như sau: [RECOMMEND: tên_sản_phẩm_1, tên_sản_phẩm_2] để hệ thống ghi nhớ.
            """

            with st.chat_message("assistant"):
                with st.spinner("AI đang phân tích da..."):
                    try:
                        genai.configure(api_key=GEMINI_API_KEY)
                        # Sử dụng model ổn định, tương thích mọi API key
                        model = genai.GenerativeModel(
                            model_name='gemini-2.5-flash',
                            system_instruction=prompt_he_thong
                        )
                        response = generate_content_with_retries(
                            model=model,
                            contents=prompt,
                            config=types.GenerationConfig(temperature=0.4),
                            max_retries=3
                        )
                        raw_answer = getattr(response, 'text', str(response))
                        match_rec = re.search(r'\[RECOMMEND:\s*(.*?)\]', raw_answer)
                        if match_rec:
                            products_extracted = [p.strip() for p in match_rec.group(1).split(',')]
                            st.session_state.recommended_products = products_extracted
                            clean_answer = re.sub(r'\[RECOMMEND:\s*(.*?)\]', '', raw_answer).strip()
                        else:
                            clean_answer = raw_answer
                        st.markdown(clean_answer)
                        st.session_state.messages.append({"role": "assistant", "content": clean_answer})
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi kết nối AI: {e}. Thử lại hoặc xóa lịch sử chat nhé!")
