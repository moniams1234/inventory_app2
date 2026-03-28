"""
app.py – Główna aplikacja Streamlit: Wiekowanie zapasów i kalkulacja rezerw.
"""
from __future__ import annotations

import io
from datetime import date, datetime

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ---------------- LOGIN ----------------
def login():
    if "logged" not in st.session_state:
        st.session_state.logged = False

    if not st.session_state.logged:
        st.markdown("## 🔐 Dostęp do aplikacji")

        password = st.text_input("Podaj hasło:", type="password")

        if password:
            if password == st.secrets["APP_PASSWORD"]:
                st.session_state.logged = True
                st.success("Zalogowano")
                st.rerun()
            else:
                st.error("Nieprawidłowe hasło")

        st.stop()

login()
# ---------------------------------------

from export import df_to_csv_bytes, export_to_excel, summary_to_csv_bytes
from processing import DEFAULT_MAPPING_PATH, process_data
from utils import (
    display_financial_metrics,
    display_metrics_row,
    style_detail_df,
    style_summary_df,
)

# ---------------------------------------------------------------------------
# Konfiguracja strony
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Wiekowanie zapasów i kalkulacja rezerw",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# CSS – kolorystyka szaro-pomarańczowa
# ---------------------------------------------------------------------------
st.markdown(
    """
<style>
/* Import Google Fonts */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* Główny styl */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Nagłówek główny */
.main-header {
    background: linear-gradient(135deg, #404040 0%, #1a1a1a 100%);
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    border-left: 6px solid #E8650A;
}
.main-header h1 {
    color: #FFFFFF;
    font-size: 1.8rem;
    font-weight: 700;
    margin: 0;
}
.main-header p {
    color: #AAAAAA;
    font-size: 0.9rem;
    margin: 0.3rem 0 0 0;
}

/* Sekcje */
.section-header {
    background: linear-gradient(90deg, #E8650A 0%, #FAD7B8 100%);
    padding: 0.5rem 1rem;
    border-radius: 6px;
    margin: 1rem 0 0.8rem 0;
    font-weight: 600;
    color: #1a1a1a;
    font-size: 0.95rem;
}

/* Badge aktywnego mappingu */
.badge-default {
    background: #E8650A;
    color: white;
    padding: 0.25rem 0.8rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 600;
    display: inline-block;
}
.badge-user {
    background: #404040;
    color: white;
    padding: 0.25rem 0.8rem;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 600;
    display: inline-block;
}

/* Instrukcja */
.instruction-box {
    background: #FFF8F3;
    border: 1px solid #FAD7B8;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin-bottom: 1rem;
    font-size: 0.88rem;
    color: #404040;
}
.instruction-box ol {
    margin: 0.5rem 0 0 1rem;
    padding: 0;
}
.instruction-box li {
    margin-bottom: 0.3rem;
}

/* Metryki */
div[data-testid="stMetric"] {
    background: #F8F8F8;
    border: 1px solid #E0E0E0;
    border-radius: 8px;
    padding: 0.6rem 1rem;
    border-top: 3px solid #E8650A;
}

/* Przyciski */
.stButton > button {
    background: linear-gradient(135deg, #E8650A, #C45508);
    color: white;
    font-weight: 600;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 2rem;
    font-size: 1rem;
    transition: all 0.2s;
    box-shadow: 0 2px 8px rgba(232,101,10,0.3);
}
.stButton > button:hover {
    background: linear-gradient(135deg, #C45508, #A34406);
    box-shadow: 0 4px 12px rgba(232,101,10,0.4);
    transform: translateY(-1px);
}

/* Pobieranie */
.stDownloadButton > button {
    background: #404040;
    color: white;
    border: none;
    border-radius: 6px;
    font-weight: 500;
}
.stDownloadButton > button:hover {
    background: #606060;
}

/* Separator */
hr {
    border-top: 2px solid #FAD7B8;
    margin: 1.5rem 0;
}

/* Ukryj zbędne elementy Streamlit */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Zakładki */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 6px 6px 0 0;
    font-weight: 500;
}
</style>
""",
    unsafe_allow_html=True,
)
def render_charts(df: pd.DataFrame) -> None:
    """Renderuje dashboard wykresów na podstawie przetworzonych danych."""
    st.markdown('<div class="section-header">📈 Dashboard i wizualizacje</div>', unsafe_allow_html=True)

    chart_df = df.copy()

    # Bezpieczeństwo typów
    chart_df["Wartość mag."] = pd.to_numeric(chart_df["Wartość mag."], errors="coerce").fillna(0.0)
    chart_df["Kwota rezerwy"] = pd.to_numeric(chart_df["Kwota rezerwy"], errors="coerce").fillna(0.0)
    chart_df["Rodzaj indeksu"] = chart_df["Rodzaj indeksu"].fillna("BRAK")
    chart_df["Type of materials"] = chart_df["Type of materials"].fillna("UNMAPPED")
    chart_df["Przedział wiekowania"] = chart_df["Przedział wiekowania"].fillna("BRAK")
    chart_df["Magazyn"] = chart_df["Magazyn"].fillna("BRAK")

    # -------------------------
    # 1. Udział PROWAX / NON PROWAX
    # -------------------------
    share_df = (
        chart_df.groupby("Rodzaj indeksu", as_index=False)["Wartość mag."]
        .sum()
        .sort_values("Wartość mag.", ascending=False)
    )

    fig_share = px.pie(
        share_df,
        names="Rodzaj indeksu",
        values="Wartość mag.",
        hole=0.55,
        title="Udział procentowy stanu magazynowego: PROWAX / NON PROWAX",
    )
    fig_share.update_traces(textposition="inside", textinfo="percent+label")
    fig_share.update_layout(
        margin=dict(l=20, r=20, t=60, b=20),
        legend_title_text="Rodzaj indeksu",
        height=420,
    )

    # -------------------------
    # 2. Stan magazynu i rezerwa: PROWAX / NON PROWAX
    # -------------------------
    compare_df = (
        chart_df.groupby("Rodzaj indeksu", as_index=False)[["Wartość mag.", "Kwota rezerwy"]]
        .sum()
    )

    compare_long = compare_df.melt(
        id_vars="Rodzaj indeksu",
        value_vars=["Wartość mag.", "Kwota rezerwy"],
        var_name="Miara",
        value_name="Wartość",
    )

    fig_compare = px.bar(
        compare_long,
        x="Wartość",
        y="Rodzaj indeksu",
        color="Miara",
        barmode="group",
        orientation="h",
        title="PROWAX / NON PROWAX – stan magazynu i rezerwa",
        text_auto=".2s",
    )
    fig_compare.update_layout(
        margin=dict(l=20, r=20, t=60, b=20),
        height=420,
        xaxis_title="Wartość [PLN]",
        yaxis_title="",
        legend_title_text="Miara",
    )

    # -------------------------
    # 3. Rezerwa wg Type of materials
    # -------------------------
    reserve_by_type = (
        chart_df.groupby("Type of materials", as_index=False)["Kwota rezerwy"]
        .sum()
        .sort_values("Kwota rezerwy", ascending=False)
    )

    fig_type = px.bar(
        reserve_by_type,
        x="Type of materials",
        y="Kwota rezerwy",
        title="Kwota rezerwy wg Type of materials",
        text_auto=".2s",
    )
    fig_type.update_layout(
        margin=dict(l=20, r=20, t=60, b=20),
        height=420,
        xaxis_title="Type of materials",
        yaxis_title="Kwota rezerwy [PLN]",
    )

    # -------------------------
    # 4. Wiekowanie zapasu wg przedziałów i rodzaju indeksu
    # -------------------------
    age_order = ["0-3 mcy", "3-6 mcy", "6-9 mcy", "9-12 mcy", "pow 12 mcy", "data > dzień analizy", "BRAK"]

    aging_df = (
        chart_df.groupby(["Przedział wiekowania", "Rodzaj indeksu"], as_index=False)["Wartość mag."]
        .sum()
    )
    aging_df["Przedział wiekowania"] = pd.Categorical(
        aging_df["Przedział wiekowania"],
        categories=age_order,
        ordered=True,
    )
    aging_df = aging_df.sort_values("Przedział wiekowania")

    fig_aging = px.bar(
        aging_df,
        x="Przedział wiekowania",
        y="Wartość mag.",
        color="Rodzaj indeksu",
        barmode="stack",
        title="Struktura wieku zapasu wg przedziałów",
        text_auto=".2s",
    )
    fig_aging.update_layout(
        margin=dict(l=20, r=20, t=60, b=20),
        height=420,
        xaxis_title="Przedział wiekowania",
        yaxis_title="Wartość magazynowa [PLN]",
        legend_title_text="Rodzaj indeksu",
    )

    # -------------------------
    # 5. TOP 10 magazynów wg rezerwy
    # -------------------------
    top_mag = (
        chart_df.groupby("Magazyn", as_index=False)["Kwota rezerwy"]
        .sum()
        .sort_values("Kwota rezerwy", ascending=False)
        .head(10)
    )

    fig_top = px.bar(
        top_mag.sort_values("Kwota rezerwy", ascending=True),
        x="Kwota rezerwy",
        y="Magazyn",
        orientation="h",
        title="TOP 10 magazynów wg kwoty rezerwy",
        text_auto=".2s",
    )
    fig_top.update_layout(
        margin=dict(l=20, r=20, t=60, b=20),
        height=460,
        xaxis_title="Kwota rezerwy [PLN]",
        yaxis_title="",
    )

    # -------------------------
    # Layout 2x2 + 1 pełna szerokość
    # -------------------------
    row1_col1, row1_col2 = st.columns(2, gap="large")
    with row1_col1:
        st.plotly_chart(fig_share, use_container_width=True)
    with row1_col2:
        st.plotly_chart(fig_compare, use_container_width=True)

    row2_col1, row2_col2 = st.columns(2, gap="large")
    with row2_col1:
        st.plotly_chart(fig_type, use_container_width=True)
    with row2_col2:
        st.plotly_chart(fig_aging, use_container_width=True)

    st.plotly_chart(fig_top, use_container_width=True)
    
# ---------------------------------------------------------------------------
# Nagłówek
# ---------------------------------------------------------------------------
st.markdown(
    """
<div class="main-header">
    <h1>📦 Wiekowanie zapasów i kalkulacja rezerw</h1>
    <p>Automatyczne wiekowanie, mapowanie typów materiałów i kalkulacja rezerw bilansowych</p>
</div>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Instrukcja użytkowania
# ---------------------------------------------------------------------------
with st.expander("📖 Jak korzystać z aplikacji?", expanded=False):
    st.markdown(
        """
<div class="instruction-box">
<strong>Kroki użytkowania:</strong>
<ol>
    <li><strong>Wybierz datę analizy</strong> – na jaką datę ma być wykonane wiekowanie zapasów.</li>
    <li><strong>Wgraj plik zapasów</strong> – plik Excel z arkuszem <em>MyPrint</em>, nagłówki w wierszu 4.</li>
    <li><strong>Wybierz źródło mappingu</strong> – domyślny (wbudowany) lub własny plik Excel z arkuszami <em>Mapp1</em> i <em>Mapp2</em>.</li>
    <li><strong>Kliknij „Przelicz"</strong> – aplikacja wykona wiekowanie i wyliczy rezerwy.</li>
    <li><strong>Pobierz wyniki</strong> – plik Excel z danymi szczegółowymi, podsumowaniem i logiem walidacji.</li>
</ol>
<strong>Wymagane kolumny w pliku zapasów:</strong>
<code>Index materiałowy, Magazyn, Typ surowca, Data przyjęcia, Wartość mag.</code>
</div>
""",
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# Sekcja 1 – Data analizy
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">📅 1. Data analizy</div>', unsafe_allow_html=True)

col_date, col_info = st.columns([1, 2])
with col_date:
    analysis_date = st.date_input(
        "Wybierz datę, na którą wykonać wiekowanie:",
        value=date.today(),
        format="DD.MM.YYYY",
        help="Wiek zapasu liczymy od 'Data przyjęcia' do tej daty.",
    )

# ---------------------------------------------------------------------------
# Sekcja 2 – Upload pliku zapasów
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">📂 2. Plik z zapasami</div>', unsafe_allow_html=True)

stock_file = st.file_uploader(
    "Wgraj plik Excel z zapasami (arkusz: MyPrint, nagłówki w wierszu 4):",
    type=["xlsx", "xls"],
    key="stock_uploader",
    help="Plik musi zawierać arkusz 'MyPrint' z nagłówkami w wierszu 4.",
)

if stock_file:
    st.success(f"✅ Wgrany plik: **{stock_file.name}** ({stock_file.size / 1024:.1f} KB)")

# ---------------------------------------------------------------------------
# Sekcja 3 – Źródło mappingu
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">🗂️ 3. Źródło mappingu</div>', unsafe_allow_html=True)

mapping_source = st.radio(
    "Wybierz źródło mappingu:",
    options=["Dane domyślne", "Chcę załadować nowe"],
    horizontal=True,
    help=(
        "**Dane domyślne** – wbudowany plik mappingu (Mapp1 + Mapp2). "
        "**Chcę załadować nowe** – własny plik Excel z arkuszami Mapp1 i Mapp2."
    ),
)

mapping_file = None
mapping_ok = True

if mapping_source == "Dane domyślne":
    if DEFAULT_MAPPING_PATH.exists():
        st.markdown(
            '<span class="badge-default">🟠 Aktywny mapping: domyślny</span>',
            unsafe_allow_html=True,
        )
        st.caption(f"Plik: `data/default_mapping.xlsx` (Mapp1: lista PROWAX | Mapp2: typy materiałów)")
    else:
        st.error(
            "❌ Domyślny plik mappingu nie istnieje! "
            "Umieść plik `default_mapping.xlsx` w katalogu `data/`."
        )
        mapping_ok = False

else:  # Chcę załadować nowe
    mapping_file = st.file_uploader(
        "Wgraj plik Excel z mappingiem (arkusze: Mapp1, Mapp2):",
        type=["xlsx", "xls"],
        key="mapping_uploader",
        help="Plik musi zawierać arkusz 'Mapp1' (indeksy PROWAX w kolumnie B) i 'Mapp2' (Magazyn + Typ surowca → Type of materials).",
    )
    if mapping_file:
        st.markdown(
            '<span class="badge-user">⚫ Aktywny mapping: plik użytkownika</span>',
            unsafe_allow_html=True,
        )
        st.caption(f"Plik: **{mapping_file.name}** ({mapping_file.size / 1024:.1f} KB)")
    else:
        st.warning(
            "⚠️ Nie wgrałeś pliku mappingu. "
            "Wgraj plik Excel z arkuszami Mapp1 i Mapp2, aby kontynuować."
        )
        mapping_ok = False

st.markdown("---")

# ---------------------------------------------------------------------------
# Sekcja 4 – Przycisk przeliczenia
# ---------------------------------------------------------------------------
st.markdown('<div class="section-header">⚙️ 4. Przelicz</div>', unsafe_allow_html=True)

can_run = stock_file is not None and mapping_ok

if not stock_file:
    st.info("ℹ️ Wgraj plik z zapasami, aby aktywować przycisk przeliczenia.")
elif not mapping_ok and mapping_source == "Chcę załadować nowe":
    st.info("ℹ️ Wgraj plik mappingu, aby aktywować przycisk przeliczenia.")

run_btn = st.button(
    "🚀 Przelicz wiekowanie i rezerwy",
    disabled=not can_run,
    use_container_width=True,
)

# ---------------------------------------------------------------------------
# Przetwarzanie
# ---------------------------------------------------------------------------
if run_btn and can_run:
    with st.spinner("⏳ Trwa przetwarzanie danych..."):
        # Resetuj pozycję w plikach
        stock_file.seek(0)
        if mapping_file:
            mapping_file.seek(0)

        result = process_data(
            stock_file=stock_file,
            analysis_date=analysis_date,
            mapping_source="default" if mapping_source == "Dane domyślne" else "user",
            mapping_file=mapping_file,
        )

    # Błędy krytyczne
    if result["errors"]:
        for err in result["errors"]:
            st.error(f"❌ {err}")

    if not result["success"]:
        st.error("❌ Przetwarzanie nie powiodło się. Sprawdź powyższe błędy.")
        st.stop()

    # Ostrzeżenia
    for warn in result["warnings"]:
        st.warning(warn)

    # Sukces
    st.success(
        f"✅ Przetwarzanie zakończone pomyślnie! "
        f"Mapping: **{result['mapping_source_label']}** | "
        f"Data analizy: **{analysis_date.strftime('%d.%m.%Y')}**"
    )

    df: pd.DataFrame = result["df"]
    summary: pd.DataFrame = result["summary"]
    stats: dict = result["stats"]

  # Metryki
st.markdown('<div class="section-header">📊 Statystyki przetwarzania</div>', unsafe_allow_html=True)
display_metrics_row(stats)
st.markdown("<br>", unsafe_allow_html=True)
display_financial_metrics(stats)

st.markdown("---")

# Wykresy
render_charts(df)

st.markdown("---")

# Podgląd wyników
tab1, tab2 = st.tabs(["📋 Dane szczegółowe (pierwsze 100 wierszy)", "📊 Tabela podsumowująca"])

with tab1:
        st.markdown(f"**Łącznie rekordów:** {len(df):,}".replace(",", " "))
        try:
            st.dataframe(style_detail_df(df), use_container_width=True, height=450)
        except Exception:
            # Fallback bez stylowania
            display_cols = [c for c in df.columns if c in [
                "Index materiałowy", "Magazyn", "Typ surowca", "Data przyjęcia",
                "Wartość mag.", "Rodzaj indeksu", "Type of materials",
                "Przedział wiekowania", "% rezerwy", "Status pozycji", "Kwota rezerwy",
            ]]
            st.dataframe(df[display_cols].head(100), use_container_width=True, height=450)

with tab2:
        try:
            st.dataframe(style_summary_df(summary), use_container_width=True)
        except Exception:
            flat = summary.copy()
            if isinstance(flat.columns, pd.MultiIndex):
                flat.columns = [" | ".join(str(c) for c in col) for col in flat.columns]
            st.dataframe(flat.reset_index(), use_container_width=True)

st.markdown("---")

 # Eksport
st.markdown('<div class="section-header">💾 5. Pobierz wyniki</div>', unsafe_allow_html=True)

with st.spinner("⏳ Generowanie pliku Excel..."):
        stock_file.seek(0)
        excel_bytes = export_to_excel(
            df=df,
            summary=summary,
            analysis_date=analysis_date,
            stats=stats,
            warnings_list=result["warnings"],
            errors_list=result["errors"],
            mapping_source_label=result["mapping_source_label"],
        )

filename_date = analysis_date.strftime("%Y%m%d")
excel_filename = f"wiekowanie_zapasow_{filename_date}.xlsx"

col1, col2, col3 = st.columns(3)

with col1:
       st.download_button(
            label="📥 Pobierz Excel (pełny)",
            data=excel_bytes,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Plik Excel z arkuszami: Dane szczegółowe, Podsumowanie, Log walidacji.",
        )

with col2:
        csv_detail = df_to_csv_bytes(df)
        st.download_button(
            label="📄 CSV – dane szczegółowe",
            data=csv_detail,
            file_name=f"zapasy_szczegolowe_{filename_date}.csv",
            mime="text/csv",
            use_container_width=True,
        )

with col3:
        csv_summary = summary_to_csv_bytes(summary)
        st.download_button(
            label="📄 CSV – podsumowanie",
            data=csv_summary,
            file_name=f"zapasy_podsumowanie_{filename_date}.csv",
            mime="text/csv",
            use_container_width=True,
        )

st.markdown(
        """
        <div style="text-align:center; color:#808080; font-size:0.8rem; margin-top:2rem;">
        Wiekowanie zapasów i kalkulacja rezerw &nbsp;|&nbsp;
        Dane przetworzone lokalnie, nie są przesyłane na zewnętrzne serwery.
        </div>
        """,
        unsafe_allow_html=True,
    )

elif not run_btn:
    # Stan początkowy – podpowiedź
    if stock_file and mapping_ok:
        st.info("👆 Wszystko gotowe! Kliknij przycisk **Przelicz wiekowanie i rezerwy**.")
