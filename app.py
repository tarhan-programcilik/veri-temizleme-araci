import io
from datetime import datetime
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Sayfa Yapılandırması
st.set_page_config(
    page_title="DataCleaner Studio 📊 (Portfolio Demo)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Kısıtlama Sabitleri (Portfolio Demo Limits)
MAX_DEMO_ROWS = 200
MAX_FILE_SIZE_MB = 2.0
UPWORK_PROFILE_URL = "https://www.upwork.com"  # Profil linkinizi buraya güncelleyebilirsiniz

# Özel CSS ile Modern Portföy Görünümü
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #64748B;
        margin-bottom: 1.2rem;
    }
    .hire-card {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        color: white;
        padding: 16px;
        border-radius: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .hire-card h4 {
        color: #FFFFFF !important;
        margin-top: 0;
        margin-bottom: 8px;
    }
    .hire-card p {
        color: #E0E7FF;
        font-size: 0.85rem;
        margin-bottom: 12px;
    }
    .demo-badge {
        background-color: #FEF3C7;
        color: #92400E;
        border: 1px solid #FCD34D;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 10px;
    }
    .locked-card {
        background-color: #FFF1F2;
        border: 1px solid #FECDD3;
        border-radius: 10px;
        padding: 24px;
        text-align: center;
        margin: 20px 0;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Oturum değişkenlerini tanımlar."""
    if "df" not in st.session_state:
        st.session_state.df = None
    if "raw_df" not in st.session_state:
        st.session_state.raw_df = None
    if "history_log" not in st.session_state:
        st.session_state.history_log = []
    if "active_file_id" not in st.session_state:
        st.session_state.active_file_id = None
    if "file_trial_used" not in st.session_state:
        st.session_state.file_trial_used = False
    if "is_row_capped" not in st.session_state:
        st.session_state.is_row_capped = False
    if "lang" not in st.session_state:
        st.session_state.lang = "en"


init_session_state()


def log_action(message: str):
    """Yapılan işlemi zaman damgasıyla kaydeder."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.history_log.append(f"[{timestamp}] {message}")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """openpyxl ile stillendirilmiş ve Upwork portföy filigranlı Excel oluşturur."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Clean_Data')
        worksheet = writer.sheets['Clean_Data']

        # Başlık biçimlendirmesi
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        # Otomatik sütun genişliği
        for col_idx, col in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(
                len(str(col)),
                df[col].astype(str).str.len().max() if not df.empty else 0
            )
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

        # Demo Filigranı / Attribution Satırı
        last_row = len(df) + 3
        watermark_cell = worksheet.cell(row=last_row, column=1)
        watermark_cell.value = "⚠️ Portfolio Demo Sample - Capped at 200 rows | Created by Tarhan Programcılık (Available for hire on Upwork)"
        watermark_cell.font = Font(name="Calibri", size=10, italic=True, color="64748B", bold=True)

    return output.getvalue()


# -------------------------------------------------------------
# KENAR ÇUBUĞU (SIDEBAR) - DİL, UPWORK CTA & KISITLI DOSYA YÜKLEME
# -------------------------------------------------------------
with st.sidebar:
    # Dil Seçimi
    lang_choice = st.selectbox(
        "Language / Dil",
        ["English (EN)", "Türkçe (TR)"],
        index=0 if st.session_state.lang == "en" else 1,
        label_visibility="collapsed"
    )
    is_tr = "TR" in lang_choice
    st.session_state.lang = "tr" if is_tr else "en"

    # Upwork "Hire Me" Kartı
    st.markdown(f"""
    <div class="hire-card">
        <h4>💼 {"Upwork Profilim" if is_tr else "Hire Me on Upwork"}</h4>
        <p>{"İşletmeniz için özel veri işleme hatları (ETL), Python otomasyonları, Streamlit panelleri veya yapay zeka entegrasyonu mu arıyorsunuz?" if is_tr else "Looking for custom data pipelines, automated ETL workflows, custom Streamlit dashboards, or AI solutions?"}</p>
        <a href="{UPWORK_PROFILE_URL}" target="_blank" style="text-decoration:none;">
            <button style="width:100%; background-color:#10B981; color:white; border:none; padding:8px 12px; border-radius:6px; font-weight:600; cursor:pointer;">
                {"👉 Upwork'ten İletişime Geçin" if is_tr else "👉 Contact Me on Upwork"}
            </button>
        </a>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f'<span class="demo-badge">🔒 {"PORTFÖY VİTRİN SÜRÜMÜ" if is_tr else "PORTFOLIO DEMO EDITION"}</span>', unsafe_allow_html=True)

    st.header("📂 " + ("Dosya Yükleme" if is_tr else "Upload Dataset"))
    st.caption(("⚡ Demo: 1 dosya deneme hakkı, maks 2MB, 200 satır." if is_tr else "⚡ Demo limits: 1 dataset per session, max 2MB, up to 200 rows."))

    uploaded_file = st.file_uploader(
        "Excel (.xlsx, .xls) / CSV (.csv)" if not is_tr else "Excel veya CSV dosyası seçin",
        type=["csv", "xlsx", "xls"],
        help="Portfolio showcase limit: 2MB" if not is_tr else "Portföy demo sınırı: Maksimum 2MB"
    )

    if uploaded_file is not None:
        file_size_mb = uploaded_file.size / (1024 * 1024)
        current_file_id = f"{uploaded_file.name}_{uploaded_file.size}"

        # 1. Dosya Boyutu Kısıtı (Max 2MB)
        if file_size_mb > MAX_FILE_SIZE_MB:
            msg = f"❌ Dosya boyutu ({file_size_mb:.2f} MB) demo sınırını ({MAX_FILE_SIZE_MB} MB) aşıyor. Lütfen daha küçük bir demo dosyası yükleyin." if is_tr else f"❌ File size ({file_size_mb:.2f} MB) exceeds the 2MB demo threshold. Please test with a smaller dataset."
            st.error(msg)
            st.stop()

        # 2. Tek Seferlik Deneme Kısıtı (Tek Dosya Hakkı)
        if st.session_state.file_trial_used and st.session_state.active_file_id != current_file_id:
            st.session_state.quota_blocked = True
        else:
            st.session_state.quota_blocked = False

            # Yeni ve ilk dosya ise oturumu başlat
            if st.session_state.active_file_id != current_file_id:
                st.session_state.active_file_id = current_file_id
                st.session_state.file_trial_used = True
                st.session_state.history_log = []

                try:
                    if uploaded_file.name.endswith('.csv'):
                        encoding_opt = st.selectbox(
                            "Encoding" if not is_tr else "Karakter Kodlaması",
                            ["utf-8", "utf-8-sig", "latin1", "iso-8859-9", "cp1254"]
                        )
                        sep_opt = st.selectbox(
                            "Delimiter" if not is_tr else "Ayraç (Delimiter)",
                            [",", ";", "\\t", "|"]
                        )
                        df_loaded = pd.read_csv(uploaded_file, sep=sep_opt, encoding=encoding_opt)
                    else:
                        excel_file = pd.ExcelFile(uploaded_file)
                        sheet_names = excel_file.sheet_names
                        selected_sheet = st.selectbox("Worksheet" if not is_tr else "Çalışma Sayfası (Sheet)", sheet_names)
                        df_loaded = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

                    # 3. Satır Sınırı Kısıtı (Max 200 Satır)
                    if len(df_loaded) > MAX_DEMO_ROWS:
                        df_loaded = df_loaded.iloc[:MAX_DEMO_ROWS].copy()
                        st.session_state.is_row_capped = True
                        log_action(f"Demo Mode: Dataset capped at first {MAX_DEMO_ROWS} rows for showcase.")
                    else:
                        st.session_state.is_row_capped = False

                    st.session_state.raw_df = df_loaded.copy()
                    st.session_state.df = df_loaded.copy()
                    log_action(f"Loaded '{uploaded_file.name}' ({len(df_loaded)} rows, {len(df_loaded.columns)} columns).")
                    st.success("File loaded successfully!" if not is_tr else "Dosya başarıyla yüklendi!")
                except Exception as e:
                    st.error(f"Error reading file: {str(e)}" if not is_tr else f"Dosya okunurken bir hata oluştu: {str(e)}")

        if st.session_state.df is not None and not st.session_state.get("quota_blocked", False):
            st.divider()
            st.subheader("🔄 " + ("Hızlı İşlemler" if is_tr else "Quick Actions"))
            if st.button("⏪ " + ("Orijinal Veriye Sıfırla" if is_tr else "Reset to Original"), use_container_width=True):
                st.session_state.df = st.session_state.raw_df.copy()
                log_action("Reset dataset to original imported state." if not is_tr else "Veri seti ilk haline sıfırlandı.")
                st.rerun()

            st.caption(f"Active File: **{uploaded_file.name}**" if not is_tr else f"Aktif Dosya: **{uploaded_file.name}**")


# -------------------------------------------------------------
# ANA EKRAN
# -------------------------------------------------------------
header_title = "📊 DataCleaner Studio (Portfolio Demo)" if not is_tr else "📊 DataCleaner Studio (Portföy Demosu)"
header_sub = "Interactive data cleaning & profiling showcase built with Python & Streamlit." if not is_tr else "Python, Streamlit ve openpyxl ile geliştirilmiş veri temizleme ve özetleme vitrini."

st.markdown(f'<div class="main-header">{header_title}</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sub-header">{header_sub}</div>', unsafe_allow_html=True)

# KOTA DOLDU / TEK DOSYA KİLİDİ
if st.session_state.get("quota_blocked", False):
    st.markdown(f"""
    <div class="locked-card">
        <h2 style="color:#BE123C; margin-top:0;">🔒 {"Demo Deneme Sınırına Ulaşıldı" if is_tr else "Demo Trial Quota Reached"}</h2>
        <p style="font-size:1.05rem; color:#475569; max-width:700px; margin:0 auto 20px auto;">
            {"Bu portföy demosunda oturum başına <b>sadece 1 adet dosya deneme hakkı</b> sunulmaktadır. Sınırsız dosya boyutu, toplu veri işleme (batch processing), otomatik ETL hatları veya özel şirket araçları geliştirmek için benimle Upwork üzerinden iletişime geçebilirsiniz!" if is_tr else "This portfolio showcase is limited to <b>1 dataset trial per session</b>. To unlock unlimited processing, automated ETL pipelines, cloud database integrations, or bespoke internal tools, feel free to hire me on Upwork!"}
        </p>
        <a href="{UPWORK_PROFILE_URL}" target="_blank" style="text-decoration:none;">
            <button style="background-color:#1E3A8A; color:white; border:none; padding:12px 28px; border-radius:8px; font-size:1rem; font-weight:700; cursor:pointer;">
                💼 {"Upwork Profilime Git & Projeyi Konuşalım" if is_tr else "Hire Me on Upwork / Discuss Your Project"}
            </button>
        </a>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Dosya Yüklenmemişse Hoş Geldiniz Ekranı
if st.session_state.df is None:
    st.info(
        "👈 Please upload a sample **Excel (.xlsx)** or **CSV** dataset (max 2MB) from the left sidebar to test the interface." 
        if not is_tr else 
        "👈 Arayüzü test etmek için lütfen sol panelden örnek bir **Excel (.xlsx)** veya **CSV** dosyası yükleyin (Maksimum 2MB)."
    )
    
    # Upwork Portföy Bilgilendirme Kartı
    st.markdown("---")
    c_info1, c_info2 = st.columns(2)
    with c_info1:
        st.markdown("### 🎯 " + ("What This Demo Demonstrates" if not is_tr else "Bu Portföy Demosu Neleri Gösterir?"))
        if not is_tr:
            st.markdown("""
            - **End-to-End Data Cleansing**: Duplicate detection, missing data imputation, outlier filtering (IQR).
            - **Automated Profiling & EDA**: Column health metrics, data type inference, statistical summaries.
            - **Production-Ready Python**: Built with pandas, Streamlit, and openpyxl with clean state management.
            - **Custom Spreadsheet Styling**: Programmatic formatting with branded header styles and auto-fit columns.
            """)
        else:
            st.markdown("""
            - **Uçtan Uca Veri Temizleme**: Çift kayıt temizliği, eksik veri doldurma, IQR aykırı değer analizi.
            - **Otomatik Profilleme**: Sütun sağlık metrikleri, veri tipi kontrolü, istatistiksel özetler.
            - **Üretime Hazır Mimari**: pandas, Streamlit ve openpyxl ile modüler ve oturum güvenli yapı.
            - **Özel Excel Biçimlendirme**: Otomatik başlık stillendirme ve dinamik sütun genişliği optimizasyonu.
            """)
    with c_info2:
        st.markdown("### 💼 " + ("Hire Me on Upwork" if not is_tr else "Upwork Üzerinden Çalışalım"))
        if not is_tr:
            st.markdown("""
            Need a tailored solution for your company?
            - Custom Web Dashboards & Internal SaaS Tools
            - Automated Web Scraping & Data Extraction
            - ETL & Database Pipelines (PostgreSQL, BigQuery, Snowflake)
            - AI/LLM App Development (LangChain, OpenAI API)
            """)
        else:
            st.markdown("""
            Şirketiniz veya projeniz için özel bir geliştirme mi istiyorsunuz?
            - Özel Web Panelleri ve Dahili Yönetim Araçları
            - Otomatik Web Kazıma (Scraping) ve Veri Çıkarma
            - ETL ve Veri Tabanı Entegrasyonları (PostgreSQL, BigQuery vb.)
            - Yapay Zeka / LLM Uygulama Geliştirme
            """)
    st.stop()

# 200 Satır Kısıtı Bildirimi
if st.session_state.is_row_capped:
    cap_msg = f"ℹ️ **Portfolio Showcase Notice**: Dataset has been capped at the first **{MAX_DEMO_ROWS} rows** for demo purposes. Full enterprise versions process unlimited rows without restriction." if not is_tr else f"ℹ️ **Portföy Vitrin Bildirimi**: Dosyanız demo amacıyla ilk **{MAX_DEMO_ROWS} satır** ile sınırlandırılmıştır. Tam kurumsal sürümlerde satır sınırı bulunmamaktadır."
    st.warning(cap_msg)

df = st.session_state.df

# -------------------------------------------------------------
# ANA SEKME YAPISI
# -------------------------------------------------------------
tab1_title = "📈 1. Overview & Profiling" if not is_tr else "📈 1. Veri Özeti & Analiz"
tab2_title = "🧹 2. Data Cleaning" if not is_tr else "🧹 2. Veri Temizleme"
tab3_title = "💾 3. Export & Download" if not is_tr else "💾 3. Dışa Aktar & İndir"

tab_summary, tab_cleaning, tab_export = st.tabs([tab1_title, tab2_title, tab3_title])

# =============================================================
# SEKME 1: VERİ ÖZETİ VE PROFİLLEME
# =============================================================
with tab_summary:
    st.subheader("📌 " + ("Key Metrics & Dataset Overview" if not is_tr else "Genel Bakış & Temel Metrikler"))
    
    total_rows = len(df)
    total_cols = len(df.columns)
    total_cells = total_rows * total_cols
    total_missing = df.isnull().sum().sum()
    missing_pct = (total_missing / total_cells * 100) if total_cells > 0 else 0
    duplicate_rows = df.duplicated().sum()
    memory_usage_mb = df.memory_usage(deep=True).sum() / (1024 * 1024)

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Rows (Capped)" if (st.session_state.is_row_capped and not is_tr) else ("Toplam Satır (Kısıtlı)" if st.session_state.is_row_capped else ("Rows" if not is_tr else "Toplam Satır")), f"{total_rows:,}")
    col2.metric("Columns" if not is_tr else "Toplam Sütun", total_cols)
    col3.metric("Missing Cells" if not is_tr else "Eksik Değerler", f"{total_missing:,}", f"{missing_pct:.1f}%")
    col4.metric("Duplicates" if not is_tr else "Yinelenen Satırlar", f"{duplicate_rows:,}")
    col5.metric("Memory Usage" if not is_tr else "Bellek Kullanımı", f"{memory_usage_mb:.2f} MB")

    st.divider()

    st.subheader("👀 " + ("Data Preview" if not is_tr else "Veri Önizlemesi"))
    view_options = ["First 10 Rows", "Last 10 Rows", "Random 10 Rows", "Full Dataset"] if not is_tr else ["İlk 10 Satır", "Son 10 Satır", "Rastgele 10 Satır", "Tüm Veri"]
    view_option = st.radio("View Mode:" if not is_tr else "Görünüm Modu:", view_options, horizontal=True)
    if "First" in view_option or "İlk" in view_option:
        st.dataframe(df.head(10), use_container_width=True)
    elif "Last" in view_option or "Son" in view_option:
        st.dataframe(df.tail(10), use_container_width=True)
    elif "Random" in view_option or "Rastgele" in view_option:
        st.dataframe(df.sample(min(10, total_rows)), use_container_width=True)
    else:
        st.dataframe(df, use_container_width=True)

    st.divider()

    st.subheader("📋 " + ("Column Health & Data Types" if not is_tr else "Sütun Sağlığı ve Veri Tipleri"))
    col_info = []
    for col in df.columns:
        null_count = df[col].isnull().sum()
        col_info.append({
            "Column Name" if not is_tr else "Sütun Adı": col,
            "Type" if not is_tr else "Veri Tipi": str(df[col].dtype),
            "Non-Null Count" if not is_tr else "Dolu Satır": total_rows - null_count,
            "Null Count" if not is_tr else "Eksik Satır": null_count,
            "Null %" if not is_tr else "Eksiklik Oranı (%)": round((null_count / total_rows) * 100, 2) if total_rows > 0 else 0,
            "Unique Values" if not is_tr else "Benzersiz Değer": df[col].nunique(),
            "Sample Value" if not is_tr else "Örnek Değer": str(df[col].dropna().iloc[0]) if not df[col].dropna().empty else "-"
        })
    info_df = pd.DataFrame(col_info)
    st.dataframe(info_df, use_container_width=True)

    null_key = "Null Count" if not is_tr else "Eksik Satır"
    null_pct_key = "Null %" if not is_tr else "Eksiklik Oranı (%)"
    name_key = "Column Name" if not is_tr else "Sütun Adı"
    missing_cols = info_df[info_df[null_key] > 0]
    
    if not missing_cols.empty:
        st.subheader("📉 " + ("Missing Value Distribution by Column" if not is_tr else "Sütunlardaki Eksik Değer Dağılımı"))
        chart_data = missing_cols.set_index(name_key)[null_pct_key]
        st.bar_chart(chart_data)
    else:
        st.success("🎉 " + ("No missing values found in this dataset!" if not is_tr else "Tebrikler! Veri setinde hiç eksik değer (null/NaN) bulunmuyor."))

    st.divider()
    st.subheader("📊 " + ("Statistical Summary (Describe)" if not is_tr else "İstatistiksel Dağılım (Describe)"))
    desc_opts = ["Numeric Columns", "All Columns"] if not is_tr else ["Sayısal Sütunlar", "Tüm Sütunlar"]
    describe_type = st.radio("Summary Scope:" if not is_tr else "İstatistik Kapsamı:", desc_opts, horizontal=True)
    if "Numeric" in describe_type or "Sayısal" in describe_type:
        num_df = df.select_dtypes(include='number')
        if not num_df.empty:
            st.dataframe(num_df.describe().T, use_container_width=True)
        else:
            st.info("No numerical columns found." if not is_tr else "Veri setinde sayısal sütun bulunamadı.")
    else:
        st.dataframe(df.describe(include='all').astype(str).T, use_container_width=True)


# =============================================================
# SEKME 2: VERİ TEMİZLEME ARAÇLARI
# =============================================================
with tab_cleaning:
    st.subheader("🛠️ " + ("Data Cleaning & Transformation Tools" if not is_tr else "Temizleme & Veri Düzenleme Araçları"))

    # 1. DUPLICATES
    with st.expander("👯 1. " + ("Remove Duplicate Rows" if not is_tr else "Yinelenen (Duplicate) Satırları Temizleme"), expanded=False):
        dup_count = df.duplicated().sum()
        st.write(("Total Exact Duplicate Rows: " if not is_tr else "Mevcut Tam Yinelenen Satır Sayısı: ") + f"**{dup_count}**")

        dup_opts = ["All Columns Identical", "Subset of Specific Columns"] if not is_tr else ["Tüm Sütunlar Aynı Olanlar", "Belirli Sütunları Aynı Olanlar"]
        dup_mode = st.radio("Duplicate Criterion:" if not is_tr else "Yinelenme Kriteri:", dup_opts, key="dup_mode")
        subset_cols = None
        if "Subset" in dup_mode or "Belirli" in dup_mode:
            subset_cols = st.multiselect("Select Subset Columns:" if not is_tr else "Kriter Sütunları Seçin:", df.columns.tolist())

        keep_opts = ["Keep First Occurrence", "Keep Last Occurrence"] if not is_tr else ["İlk Kayıt (first)", "Son Kayıt (last)"]
        keep_option = st.selectbox("Record to Retain:" if not is_tr else "Hangi Kayıt Saklansın?", keep_opts, index=0)
        keep_val = "first" if ("First" in keep_option or "İlk" in keep_option) else "last"

        if st.button("🗑️ " + ("Drop Duplicates" if not is_tr else "Yinelenen Satırları Temizle"), type="primary"):
            cols_to_check = subset_cols if (subset_cols and len(subset_cols) > 0) else None
            before_len = len(df)
            st.session_state.df = df.drop_duplicates(subset=cols_to_check, keep=keep_val).reset_index(drop=True)
            removed = before_len - len(st.session_state.df)
            log_action(f"Dropped {removed} duplicate rows (retained: {keep_val}).")
            st.success(f"Successfully removed {removed} duplicate row(s)!" if not is_tr else f"{removed} adet satır başarıyla silindi!")
            st.rerun()

    # 2. MISSING VALUES
    with st.expander("❓ 2. " + ("Missing Value Management (Drop / Impute)" if not is_tr else "Eksik Değer (NaN / Null) Yönetimi"), expanded=False):
        st.markdown("**A) " + ("Drop Rows with Missing Values" if not is_tr else "Eksik Değer İçeren Satırları Silme") + "**")
        drop_na_opts = ["Drop if ALL cells are empty", "Drop if ANY cell is empty", "Drop if specific columns are empty"] if not is_tr else ["Tüm satır boşsa sil (all)", "Herhangi bir hücre boşsa sil (any)", "Seçili sütunlar boşsa sil (subset)"]
        drop_na_mode = st.selectbox("Condition:" if not is_tr else "Silme Koşulu:", drop_na_opts)
        drop_subset = None
        if "specific" in drop_na_mode or "Seçili" in drop_na_mode:
            drop_subset = st.multiselect("Select target columns:" if not is_tr else "Kontrol edilecek sütunlar:", df.columns.tolist(), key="dropna_cols")

        if st.button("❌ " + ("Drop Missing Rows" if not is_tr else "Eksik Satırları Sil")):
            before_len = len(df)
            how_val = "all" if ("ALL" in drop_na_mode or "all" in drop_na_mode) else "any"
            st.session_state.df = df.dropna(how=how_val, subset=drop_subset if drop_subset else None).reset_index(drop=True)
            removed = before_len - len(st.session_state.df)
            log_action(f"Removed {removed} rows with missing values.")
            st.success(f"Dropped {removed} rows." if not is_tr else f"{removed} adet satır silindi.")
            st.rerun()

        st.divider()
        st.markdown("**B) " + ("Impute / Fill Missing Values" if not is_tr else "Eksik Değerleri Doldurma (Imputation)") + "**")
        fill_col = st.selectbox("Target Column:" if not is_tr else "Doldurulacak Sütun:", df.columns.tolist(), key="fill_col")
        
        col_type = df[fill_col].dtype
        null_in_col = df[fill_col].isnull().sum()
        caption_txt = f"'{fill_col}' has **{null_in_col}** missing entries. (Type: {col_type})" if not is_tr else f"'{fill_col}' sütununda **{null_in_col}** adet eksik değer var. (Tip: {col_type})"
        st.caption(caption_txt)

        base_methods = ["Constant Value", "Mode (Most Frequent)", "Forward Fill (ffill)", "Backward Fill (bfill)"] if not is_tr else ["Sabit Değer", "En Çok Tekrar Eden Değer (Mod)", "Önceki Değerle Doldur (ffill)", "Sonraki Değerle Doldur (bfill)"]
        if pd.api.types.is_numeric_dtype(df[fill_col]):
            num_methods = ["Mean (Average)", "Median"] if not is_tr else ["Ortalama (Mean)", "Medyan (Median)"]
            fill_methods = num_methods + base_methods
        else:
            fill_methods = base_methods

        chosen_method = st.selectbox("Imputation Strategy:" if not is_tr else "Doldurma Yöntemi:", fill_methods)
        custom_fill_value = None
        if "Constant" in chosen_method or "Sabit" in chosen_method:
            default_val = "0" if pd.api.types.is_numeric_dtype(df[fill_col]) else "Unknown"
            custom_fill_value = st.text_input("Fill Value:" if not is_tr else "Doldurulacak Değer:", value=default_val)

        if st.button("✏️ " + ("Apply Imputation" if not is_tr else "Eksik Değerleri Doldur")):
            if null_in_col == 0:
                st.info("No missing values in this column." if not is_tr else "Bu sütunda zaten eksik değer yok.")
            else:
                updated_col = df[fill_col].copy()
                if "Mean" in chosen_method or "Ortalama" in chosen_method:
                    val = updated_col.mean()
                    updated_col.fillna(val, inplace=True)
                elif "Median" in chosen_method or "Medyan" in chosen_method:
                    val = updated_col.median()
                    updated_col.fillna(val, inplace=True)
                elif "Mode" in chosen_method or "Mod" in chosen_method:
                    mode_val = updated_col.mode()
                    val = mode_val[0] if not mode_val.empty else ""
                    updated_col.fillna(val, inplace=True)
                elif "Forward" in chosen_method or "Önceki" in chosen_method:
                    updated_col.ffill(inplace=True)
                elif "Backward" in chosen_method or "Sonraki" in chosen_method:
                    updated_col.bfill(inplace=True)
                elif "Constant" in chosen_method or "Sabit" in chosen_method:
                    if pd.api.types.is_numeric_dtype(df[fill_col]):
                        try:
                            val = float(custom_fill_value) if "." in custom_fill_value else int(custom_fill_value)
                        except ValueError:
                            val = custom_fill_value
                    else:
                        val = custom_fill_value
                    updated_col.fillna(val, inplace=True)

                st.session_state.df[fill_col] = updated_col
                log_action(f"Imputed missing entries in '{fill_col}' via '{chosen_method}'.")
                st.success(f"Column '{fill_col}' successfully imputed!" if not is_tr else f"'{fill_col}' sütunu başarıyla dolduruldu.")
                st.rerun()

    # 3. COLUMN OPERATIONS
    with st.expander("📑 3. " + ("Column Operations (Drop, Normalize, Rename)" if not is_tr else "Sütun Yönetimi (Silme, Yeniden Adlandırma & Normalizasyon)"), expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**" + ("Drop Unwanted Columns" if not is_tr else "Gereksiz Sütunları Kaldırma") + "**")
            cols_to_drop = st.multiselect("Select columns to drop:" if not is_tr else "Silinecek Sütunları Seçin:", df.columns.tolist())
            if st.button("🗑️ " + ("Drop Selected Columns" if not is_tr else "Seçili Sütunları Sil"), disabled=len(cols_to_drop) == 0):
                st.session_state.df.drop(columns=cols_to_drop, inplace=True)
                log_action(f"Dropped columns: {', '.join(cols_to_drop)}")
                st.success(f"Removed {len(cols_to_drop)} column(s)." if not is_tr else f"{len(cols_to_drop)} sütun silindi.")
                st.rerun()

        with c2:
            st.markdown("**" + ("Auto-Normalize Column Headers" if not is_tr else "Sütun İsimlerini Normalize Etme") + "**")
            st.caption("Converts to snake_case, removes special characters and trims spaces." if not is_tr else "Boşlukları alt çizgiye dönüştürür, küçük harf yapar, Türkçe karakterleri standartlaştırır.")
            if st.button("✨ " + ("Auto-Clean Headers" if not is_tr else "Sütun İsimlerini Otomatik Düzelt")):
                tr_map = str.maketrans("çğıöşüÇĞİÖŞÜ ", "cgiosuCGIOSU_")
                new_cols = [c.strip().translate(tr_map).lower() for c in df.columns]
                st.session_state.df.columns = new_cols
                log_action("Normalized all column headers to snake_case.")
                st.success("Column headers cleaned!" if not is_tr else "Sütun adları başarıyla düzenlendi.")
                st.rerun()

        st.divider()
        st.markdown("**" + ("Rename a Specific Column" if not is_tr else "Tek Bir Sütunu Yeniden Adlandırma") + "**")
        col_to_rename = st.selectbox("Column to rename:" if not is_tr else "Adı değiştirilecek sütun:", df.columns.tolist(), key="rename_col")
        new_name = st.text_input("New Name:" if not is_tr else "Yeni İsim:", value=col_to_rename)
        if st.button("Update Name" if not is_tr else "Adı Güncelle", disabled=(new_name == col_to_rename or not new_name.strip())):
            st.session_state.df.rename(columns={col_to_rename: new_name.strip()}, inplace=True)
            log_action(f"Renamed '{col_to_rename}' -> '{new_name.strip()}'.")
            st.success(f"Renamed to '{new_name.strip()}'." if not is_tr else f"Sütun adı güncellendi: {new_name.strip()}")
            st.rerun()

    # 4. TYPE CASTING
    with st.expander("🔄 4. " + ("Type Casting (Data Type Conversion)" if not is_tr else "Veri Tipi Dönüştürme (Type Casting)"), expanded=False):
        c_col, c_type = st.columns(2)
        with c_col:
            target_col = st.selectbox("Target Column:" if not is_tr else "Dönüştürülecek Sütun:", df.columns.tolist(), key="cast_col")
        with c_type:
            types_list = ["Numeric (Float)", "Integer", "String (Text)", "Datetime", "Categorical"] if not is_tr else ["Sayısal (Numeric / Float)", "Tam Sayı (Integer)", "Metin (String / Text)", "Tarih/Saat (Datetime)", "Kategori (Category)"]
            target_type = st.selectbox("Destination Type:" if not is_tr else "Hedef Veri Tipi:", types_list)

        if st.button("Convert Type" if not is_tr else "Dönüşümü Uygula"):
            try:
                if "Numeric" in target_type or "Sayısal" in target_type:
                    st.session_state.df[target_col] = pd.to_numeric(st.session_state.df[target_col], errors='coerce')
                elif "Integer" in target_type or "Tam Sayı" in target_type:
                    st.session_state.df[target_col] = pd.to_numeric(st.session_state.df[target_col], errors='coerce').fillna(0).astype('int64')
                elif "String" in target_type or "Metin" in target_type:
                    st.session_state.df[target_col] = st.session_state.df[target_col].astype(str)
                elif "Datetime" in target_type or "Tarih" in target_type:
                    st.session_state.df[target_col] = pd.to_datetime(st.session_state.df[target_col], errors='coerce')
                elif "Categorical" in target_type or "Kategori" in target_type:
                    st.session_state.df[target_col] = st.session_state.df[target_col].astype('category')

                log_action(f"Cast '{target_col}' to '{target_type}'.")
                st.success(f"Column '{target_col}' converted to {target_type}!" if not is_tr else f"'{target_col}' tipi başarıyla güncellendi!")
                st.rerun()
            except Exception as e:
                st.error(f"Conversion error: {str(e)}" if not is_tr else f"Dönüştürme hatası: {str(e)}")

    # 5. STRING CLEANING
    with st.expander("🔤 5. " + ("String Formatting (Strip & Case)" if not is_tr else "Metin Temizleme (Boşluk Temizleme & Büyük/Küçük Harf)"), expanded=False):
        text_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
        if not text_cols:
            st.info("No string/text columns available." if not is_tr else "Veri setinde metin (string/object) sütunu bulunmuyor.")
        else:
            selected_text_col = st.selectbox("Text Column:" if not is_tr else "Metin Sütunu:", text_cols, key="text_col_select")
            str_actions = ["Strip Whitespace", "lowercase", "UPPERCASE", "Title Case"] if not is_tr else ["Baş/Son Boşlukları Kırp (Strip Whitespace)", "Küçük Harf Yap (lowercase)", "BÜYÜK HARF YAP (UPPERCASE)", "Baş Harfleri Büyüt (Title Case)"]
            text_action = st.radio("Formatting Action:" if not is_tr else "Uygulanacak Metin İşlemi:", str_actions, horizontal=True)
            if st.button("Apply Text Formatting" if not is_tr else "Metin İşlemini Uygula"):
                series = df[selected_text_col].astype(str)
                if "Strip" in text_action or "Kırp" in text_action:
                    st.session_state.df[selected_text_col] = series.str.strip()
                elif "lower" in text_action or "Küçük" in text_action:
                    st.session_state.df[selected_text_col] = series.str.lower()
                elif "UPPER" in text_action or "BÜYÜK" in text_action:
                    st.session_state.df[selected_text_col] = series.str.upper()
                elif "Title" in text_action or "Baş Harfleri" in text_action:
                    st.session_state.df[selected_text_col] = series.str.title()

                log_action(f"Formatted text column '{selected_text_col}' ({text_action}).")
                st.success("Text formatting applied successfully!" if not is_tr else "Metin düzenleme tamamlandı.")
                st.rerun()

    # 6. OUTLIER FILTER
    with st.expander("📐 6. " + ("Outlier Detection & Removal (Numeric)" if not is_tr else "Aykırı Değer Filtreleme (Sayısal Sütunlar)"), expanded=False):
        num_cols = df.select_dtypes(include='number').columns.tolist()
        if not num_cols:
            st.info("No numerical columns for outlier inspection." if not is_tr else "Aykırı değer tespiti için sayısal sütun bulunmuyor.")
        else:
            selected_num_col = st.selectbox("Select Numeric Column:" if not is_tr else "Sayısal Sütun Seçin:", num_cols, key="outlier_col")
            method_opts = ["IQR (Interquartile Range - Recommended)", "Custom Min / Max Range"] if not is_tr else ["IQR (Çeyrekler Açıklığı - Önerilen)", "Özel Min / Max Eşiği"]
            outlier_mode = st.radio("Method:" if not is_tr else "Yöntem:", method_opts, horizontal=True)

            if "IQR" in outlier_mode:
                q1 = df[selected_num_col].quantile(0.25)
                q3 = df[selected_num_col].quantile(0.75)
                iqr = q3 - q1
                factor = st.slider("IQR Multiplier (Default: 1.5):" if not is_tr else "IQR Çarpanı (Varsayılan: 1.5):", min_value=1.0, max_value=3.0, value=1.5, step=0.1)
                lower_bound = q1 - factor * iqr
                upper_bound = q3 + factor * iqr
                outliers_count = ((df[selected_num_col] < lower_bound) | (df[selected_num_col] > upper_bound)).sum()
                
                bnd_txt = f"Lower Bound: **{lower_bound:.2f}**, Upper Bound: **{upper_bound:.2f}** | Detected Outliers: **{outliers_count}**" if not is_tr else f"Alt Sınır: **{lower_bound:.2f}**, Üst Sınır: **{upper_bound:.2f}** | Tespit Edilen Aykırı Satır: **{outliers_count}**"
                st.write(bnd_txt)

                btn_drop = "Drop Outlier Rows" if not is_tr else "Aykırı Değerleri Filtrele (Satırları Sil)"
                if st.button(btn_drop, disabled=outliers_count == 0):
                    before_len = len(df)
                    st.session_state.df = df[(df[selected_num_col] >= lower_bound) & (df[selected_num_col] <= upper_bound)].reset_index(drop=True)
                    removed = before_len - len(st.session_state.df)
                    log_action(f"Filtered {removed} outliers from '{selected_num_col}' (IQR={factor}).")
                    st.success(f"Removed {removed} outlier row(s)." if not is_tr else f"{removed} adet aykırı satır filtrelendi.")
                    st.rerun()
            else:
                col_min = float(df[selected_num_col].min()) if not df[selected_num_col].empty else 0.0
                col_max = float(df[selected_num_col].max()) if not df[selected_num_col].empty else 100.0
                c_min, c_max = st.columns(2)
                with c_min:
                    min_val = st.number_input("Minimum Value:" if not is_tr else "Minimum Değer:", value=col_min)
                with c_max:
                    max_val = st.number_input("Maximum Value:" if not is_tr else "Maksimum Değer:", value=col_max)

                btn_range = "Filter by Range" if not is_tr else "Eşik Değerlere Göre Filtrele"
                if st.button(btn_range):
                    before_len = len(df)
                    st.session_state.df = df[(df[selected_num_col] >= min_val) & (df[selected_num_col] <= max_val)].reset_index(drop=True)
                    removed = before_len - len(st.session_state.df)
                    log_action(f"Filtered '{selected_num_col}' with range [{min_val}, {max_val}].")
                    st.success(f"Filtered {removed} row(s)." if not is_tr else f"{removed} satır filtrelendi.")
                    st.rerun()


# =============================================================
# SEKME 3: DIŞA AKTARMA VE İNDİRME
# =============================================================
with tab_export:
    st.subheader("💾 " + ("Download Cleaned Dataset" if not is_tr else "Temizlenmiş Veriyi İndirme"))
    
    dim_txt = f"Current Dimensions: **{len(df):,} Rows** × **{len(df.columns)} Columns**" if not is_tr else f"Mevcut Veri Boyutu: **{len(df):,} Satır** × **{len(df.columns)} Sütun**"
    st.write(dim_txt)
    st.dataframe(df.head(10), use_container_width=True)

    col_dl1, col_dl2 = st.columns(2)

    # EXCEL EXPORT
    with col_dl1:
        st.markdown("### 📗 Excel (.xlsx)")
        st.caption("Styled navy headers, bold white text, auto-fitted columns & demo watermark." if not is_tr else "openpyxl ile stillendirilmiş, otomatik sütun genişlikli ve demo filigranlı dosya.")
        with st.spinner("Preparing Excel file..." if not is_tr else "Excel dosyası hazırlanıyor..."):
            excel_data = to_excel_bytes(df)
        
        default_excel_name = f"demo_cleaned_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="📥 Download Excel (.xlsx)" if not is_tr else "📥 Excel (.xlsx) Olarak İndir",
            data=excel_data,
            file_name=default_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

    # CSV EXPORT
    with col_dl2:
        st.markdown("### 📄 CSV (.csv)")
        st.caption("UTF-8 BOM encoded with customizable delimiter (Excel compatible)." if not is_tr else "Excel ile uyumlu (UTF-8 BOM) ve ayraç seçilebilir CSV.")
        csv_sep = st.selectbox("CSV Delimiter:" if not is_tr else "CSV Ayırıcı Karakter:", [",", ";", "\\t"], index=0)
        
        csv_data = df.to_csv(index=False, sep=csv_sep).encode('utf-8-sig')
        default_csv_name = f"demo_cleaned_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        st.download_button(
            label="📥 Download CSV (.csv)" if not is_tr else "📥 CSV (.csv) Olarak İndir",
            data=csv_data,
            file_name=default_csv_name,
            mime="text/csv",
            use_container_width=True
        )

    st.divider()

    # AUDIT LOG
    st.subheader("📜 " + ("Action History & Audit Log" if not is_tr else "Yapılan İşlemlerin Geçmişi (Audit Log)"))
    if st.session_state.history_log:
        for log_entry in reversed(st.session_state.history_log):
            st.markdown(f"- {log_entry}")
    else:
        st.info("No cleaning actions performed yet." if not is_tr else "Henüz herhangi bir temizleme işlemi uygulanmadı.")
