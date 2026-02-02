import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Analytics Dashboard",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }
    
    .stApp {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    }
    
    .main-header {
        font-size: 3rem;
        font-weight: 800;
        background: linear-gradient(135deg, #00d4ff 0%, #7b2cbf 50%, #ff6b6b 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 1rem;
    }
    
    .sub-header {
        font-size: 1.2rem;
        color: #a0aec0;
        text-align: center;
        margin-bottom: 3rem;
    }
    
    .menu-card {
        background: linear-gradient(145deg, #1e2a4a 0%, #2d3a5a 100%);
        border-radius: 20px;
        padding: 2rem;
        border: 1px solid rgba(255,255,255,0.1);
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        text-align: center;
        transition: all 0.3s ease;
        height: 250px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .menu-card:hover {
        transform: translateY(-10px);
        box-shadow: 0 15px 40px rgba(0,212,255,0.3);
        border-color: rgba(0,212,255,0.5);
    }
    
    .menu-icon {
        font-size: 4rem;
        margin-bottom: 1rem;
    }
    
    .menu-title {
        font-size: 1.5rem;
        font-weight: 600;
        color: #ffffff;
        margin-bottom: 0.5rem;
    }
    
    .menu-desc {
        font-size: 0.9rem;
        color: #a0aec0;
    }
    
    .stSidebar {
        background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%);
    }
    
    .stSidebar [data-testid="stMarkdownContainer"] p {
        color: #ffffff;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<h1 class="main-header">🏠 Analytics Dashboard</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Pilih dashboard yang ingin Anda lihat dari menu di sidebar</p>', unsafe_allow_html=True)

# Sidebar info
with st.sidebar:
    st.markdown("## 📌 Navigasi")
    st.markdown("---")
    st.info("Gunakan menu di atas untuk berpindah antar dashboard")
    st.markdown("---")
    st.markdown("### 📊 Dashboard Tersedia")
    st.markdown("""
    - **Promo Dashboard** - Analisis performa promosi
    - **Ended Promo** - Summary promo yang berakhir Jan 2026
    """)

# Menu Cards
st.markdown("### 📂 Pilih Dashboard")
st.markdown("<br>", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class="menu-card">
        <div class="menu-icon">📊</div>
        <div class="menu-title">Promo Dashboard</div>
        <div class="menu-desc">Analisis performa promosi dan kontribusi terhadap Net Sales</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Buka Promo Dashboard", use_container_width=True):
        st.switch_page("pages/1_📊_Promo_Dashboard.py")

with col2:
    st.markdown("""
    <div class="menu-card">
        <div class="menu-icon">📈</div>
        <div class="menu-title">Ended Promo</div>
        <div class="menu-desc">Summary Promo yang Berakhir - January 2026</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Buka Ended Promo", use_container_width=True):
        st.switch_page("pages/2_📈_Ended_Promo.py")

with col3:
    st.markdown("""
    <div class="menu-card">
        <div class="menu-icon">⚙️</div>
        <div class="menu-title">Coming Soon</div>
        <div class="menu-desc">Dashboard tambahan akan hadir segera</div>
    </div>
    """, unsafe_allow_html=True)
    st.button("Coming Soon", use_container_width=True, disabled=True)

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #a0aec0; padding: 2rem;'>
        <p>🏠 Analytics Dashboard | Built with ❤️ using Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)
