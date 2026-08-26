"""
Lacoste Rapor - Streamlit Cloud
"""
import streamlit as st
import sys, tempfile
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

st.set_page_config(page_title="🐊 Lacoste Rapor", page_icon="🐊", layout="centered")

st.markdown("""
<style>
.stApp{background:#1a1008}
h1,h2,h3{color:#f0d8a0}
.stButton>button{background:#C8A84B!important;color:#2c2416!important;font-weight:700!important;border:none!important;border-radius:10px!important;padding:.7rem 2rem!important;font-size:15px!important;width:100%}
.stButton>button:hover{background:#d4b050!important}
.stDownloadButton>button{background:#1e7a45!important;color:#fff!important;font-weight:700!important;border:none!important;border-radius:10px!important;padding:.7rem 2rem!important;font-size:15px!important;width:100%}
footer{display:none}
</style>
""", unsafe_allow_html=True)

st.title("🐊 Lacoste Rapor Sistemi")
st.markdown("Haftalık takip Excel dosyalarını yükleyin → Rapor otomatik oluşur → İndirin")

with st.expander("ℹ️ Nasıl kullanılır?"):
    st.markdown("""
    1. **26AW TAKİP.xlsx** ve/veya **26SS TAKİP.xlsx** dosyalarını yükleyin
    2. Bu dosyalar **ANA TABLO** sayfası içermeli (haftalık satış takip dosyası)
    3. **Rapor Oluştur** butonuna tıklayın
    4. **Raporu İndir** ile HTML dosyasını indirin
    5. İndirilen dosyayı Edge veya Chrome'da açın
    """)

st.divider()

col1, col2 = st.columns(2)
with col1:
    st.markdown("**◆ 26AW Takip Dosyası**")
    aw_file = st.file_uploader("26AW", type=["xlsx","xls"], key="aw", label_visibility="collapsed")
    if aw_file:
        st.success(f"✅ {aw_file.name}")

with col2:
    st.markdown("**◆ 26SS Takip Dosyası**")
    ss_file = st.file_uploader("26SS", type=["xlsx","xls"], key="ss", label_visibility="collapsed")
    if ss_file:
        st.success(f"✅ {ss_file.name}")

st.caption("Not: Sadece biri yüklemek yeterlidir")
st.divider()

if st.button("⚡ Rapor Oluştur", use_container_width=True):
    if not aw_file and not ss_file:
        st.error("Lütfen en az bir Excel dosyası yükleyin!")
        st.stop()

    try:
        import lacoste_rapor_v2 as lr
        import importlib
        importlib.reload(lr)

        bar = st.progress(0, text="Başlatılıyor...")

        with tempfile.TemporaryDirectory() as tmp:
            aw_path = ss_path = None

            if aw_file:
                aw_path = Path(tmp) / aw_file.name
                aw_path.write_bytes(aw_file.getvalue())

            if ss_file:
                ss_path = Path(tmp) / ss_file.name
                ss_path.write_bytes(ss_file.getvalue())

            aw_path = aw_path or ss_path
            ss_path = ss_path or aw_path

            bar.progress(20, text="26AW okunuyor...")
            df_aw = lr.oku_excel(aw_path)

            bar.progress(50, text="26SS okunuyor...")
            df_ss = lr.oku_excel(ss_path)

            bar.progress(80, text="Rapor oluşturuluyor...")
            hafta = datetime.now().strftime("%d.%m.%Y - Hafta %V")
            html  = lr.uret_html_cift(df_aw, df_ss, hafta,
                                       (aw_file or ss_file).name,
                                       (ss_file or aw_file).name)
            bar.progress(100, text="Tamamlandı!")

        dosya_adi = f"lacoste_{datetime.now().strftime('%Y%m%d_%H%M')}.html"
        st.success(f"✅ Rapor hazır! ({len(html)//1024} KB)")
        st.download_button(
            label="📥 Raporu İndir (.html)",
            data=html.encode("utf-8"),
            file_name=dosya_adi,
            mime="text/html",
            use_container_width=True
        )
        st.info("İndirilen dosyayı Edge veya Chrome'da açın.")

    except KeyError as e:
        st.error(f"Excel formatı hatalı: {e} kolonu bulunamadı.")
        st.warning("Lütfen **ANA TABLO** sayfası içeren haftalık takip dosyasını yükleyin. Şablon dosyası değil.")
    except Exception as e:
        st.error(f"Hata: {e}")
        with st.expander("Hata detayı"):
            st.exception(e)
