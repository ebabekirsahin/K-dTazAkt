"""
╔══════════════════════════════════════════════════════════════════════╗
║   KIDEM TAZMİNATI AKTÜERYAL HESAPLAMA ARACI  —  IAS 19 / TMS 19    ║
║   Yöntem: Projected Unit Credit (PUC)                                ║
║   Çıktı : Excel raporu + PDF aktüeryal rapor                         ║
╚══════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import math
from datetime import date, datetime

# ─────────────────────────────────────────────────────────────────────
# SAYFA AYARLARI
# ─────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Kıdem Tazminatı Aktüeryal Hesaplama",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Kıdem Tazminatı Aktüeryal Hesaplama")
st.caption("IAS 19 / TMS 19 — Projected Unit Credit Yöntemi")

# ─────────────────────────────────────────────────────────────────────
# BÖLÜM 1 — AKTÜERYAL VARSAYIMLAR (Sidebar)
# ─────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Aktüeryal Varsayımlar")

    st.subheader("Finansal Varsayımlar")
    discount_rate = st.number_input(
        "İskonto Oranı (%)",
        min_value=0.0, max_value=50.0, value=22.0, step=0.5,
        help="Devlet iç borçlanma senetleri (DİBS) faiz oranı baz alınır."
    ) / 100

    salary_increase = st.number_input(
        "Maaş Artış Oranı (%)",
        min_value=0.0, max_value=100.0, value=30.0, step=0.5,
        help="Enflasyon beklentisi + reel maaş artışı."
    ) / 100

    ceiling_2025 = st.number_input(
        "Kıdem Tazminatı Tavanı (TL)",
        min_value=1000.0, value=35058.58, step=100.0,
        help="2025 yılı kıdem tazminatı tavanı. Her yıl güncellenir."
    )

    st.subheader("Demografik Varsayımlar")
    turnover_rate = st.number_input(
        "İşten Ayrılma Olasılığı (%)",
        min_value=0.0, max_value=50.0, value=8.0, step=0.5,
        help="Yıllık gönüllü işten ayrılma / devir hızı."
    ) / 100

    retirement_age = st.number_input(
        "Emeklilik Yaşı",
        min_value=50, max_value=70, value=58, step=1,
        help="Kıdem tazminatına hak kazanılan ortalama emeklilik yaşı."
    )

    mortality_rate = st.number_input(
        "Ölüm Olasılığı (%)",
        min_value=0.0, max_value=5.0, value=0.3, step=0.05,
        help="TRH 2010 Türkiye mortalite tablosu — basitleştirilmiş sabit oran."
    ) / 100

    st.subheader("Duyarlılık Analizi")
    run_sensitivity = st.checkbox("±1% Senaryo Analizi Yap", value=True)

# ─────────────────────────────────────────────────────────────────────
# BÖLÜM 2 — VERİ GİRİŞİ
# ─────────────────────────────────────────────────────────────────────
st.header("1️⃣ Çalışan Verisi Girişi")

input_method = st.radio(
    "Veri giriş yöntemi:",
    ["Manuel Giriş (Form)", "Excel / CSV Yükle"],
    horizontal=True
)

SAMPLE_DATA = pd.DataFrame({
    "Ad Soyad":        ["Ayşe Kaya", "Mehmet Demir", "Fatma Yılmaz", "Ali Çelik", "Zeynep Arslan"],
    "Doğum Yılı":      [1985, 1978, 1990, 1982, 1995],
    "İşe Giriş Yılı":  [2010, 2005, 2015, 2008, 2020],
    "Aylık Brüt Maaş": [45000, 80000, 35000, 60000, 28000],
    "Cinsiyet":        ["K", "E", "K", "E", "K"],
})

if input_method == "Manuel Giriş (Form)":
    st.info("Aşağıdaki tabloya tıklayarak veri girebilir veya satır ekleyebilirsiniz.")
    df_input = st.data_editor(
        SAMPLE_DATA,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Ad Soyad":        st.column_config.TextColumn("Ad Soyad"),
            "Doğum Yılı":      st.column_config.NumberColumn("Doğum Yılı", min_value=1950, max_value=2005, format="%d"),
            "İşe Giriş Yılı":  st.column_config.NumberColumn("İşe Giriş Yılı", min_value=1980, max_value=2025, format="%d"),
            "Aylık Brüt Maaş": st.column_config.NumberColumn("Aylık Brüt Maaş (TL)", min_value=0, format="%.0f"),
            "Cinsiyet":        st.column_config.SelectboxColumn("Cinsiyet", options=["E", "K"]),
        }
    )
else:
    uploaded = st.file_uploader(
        "Excel veya CSV dosyası yükleyin",
        type=["xlsx", "csv"],
        help="Kolonlar: Ad Soyad, Doğum Yılı, İşe Giriş Yılı, Aylık Brüt Maaş, Cinsiyet"
    )
    if uploaded:
        if uploaded.name.endswith(".csv"):
            df_input = pd.read_csv(uploaded)
        else:
            df_input = pd.read_excel(uploaded)
        st.success(f"✅ {len(df_input)} çalışan yüklendi.")
        st.dataframe(df_input, use_container_width=True)
    else:
        st.warning("Henüz dosya yüklenmedi — örnek veriyle devam ediliyor.")
        df_input = SAMPLE_DATA.copy()

# ─────────────────────────────────────────────────────────────────────
# BÖLÜM 3 — PUC HESAPLAMA MOTORU
# ─────────────────────────────────────────────────────────────────────

def puc_hesapla(df: pd.DataFrame, r: float, s: float, tavan: float,
                turnover: float, ret_age: int, mort: float) -> pd.DataFrame:
    """
    ══════════════════════════════════════════════════════════════════
    PROJECTED UNIT CREDIT (PUC) MOTORU — IAS 19 Para. 67-98
    ══════════════════════════════════════════════════════════════════

    Her çalışan için:
      1. Bugüne kadar kazanılmış hizmet yılı (past service)
      2. Emekliliğe kalan süre (t)
      3. Emeklilikteki tahmini maaş (gelecekteki değer)
      4. Emeklilikteki kıdem tazminatı ödemesi (tavan sınırlı)
      5. Hayatta kalma & işte kalma olasılıkları
      6. Bugünkü değer (present value = PV)
      7. Kazanılmış birim (DBO = PV × past/total)
    """

    bugun = date.today()
    yil = bugun.year

    sonuclar = []
    for _, row in df.iterrows():
        yas          = yil - int(row["Doğum Yılı"])
        kidem        = yil - int(row["İşe Giriş Yılı"])
        kidem        = max(kidem, 1)                        # min 1 yıl
        maaş         = float(row["Aylık Brüt Maaş"])

        # Emekliliğe kalan süre
        t = max(ret_age - yas, 1)

        # Toplam beklenen hizmet yılı (bugüne + kalan)
        toplam_kidem = kidem + t

        # ── 1. Emeklilikteki tahmini maaş ──────────────────────────
        # Maaş her yıl (s) oranında büyür
        maas_emekli = maaş * (1 + s) ** t

        # ── 2. Kıdem tazminatı ödemesi ─────────────────────────────
        # Her hizmet yılı için 1 aylık brüt maaş (tavan sınırlı)
        # Toplam = toplam_kidem × min(maas_emekli, tavan)
        maas_efektif = min(maas_emekli, tavan)
        odeme_emekli = maas_efektif * toplam_kidem

        # ── 3. Hayatta kalma & işte kalma olasılığı ────────────────
        # Basit model: her yıl (mort + turnover) kaybediliyor
        # p = (1 - mort - turnover) ^ t   →  t yıl sonra işte olma
        p_hayatta = (1 - mort - turnover) ** t
        p_hayatta = max(p_hayatta, 0.0)

        # ── 4. Present Value (iskonto) ──────────────────────────────
        # PV = ödeme × p × e^(-r×t)   [sürekli bileşik yaklaşım]
        # Alternatif: PV = ödeme × p / (1+r)^t  (dönemsel)
        pv_toplam = odeme_emekli * p_hayatta / ((1 + r) ** t)

        # ── 5. Kazanılmış Birim (DBO — Defined Benefit Obligation) ─
        # Bugüne kadar kazanılan kısım = past / toplam
        dbo = pv_toplam * (kidem / toplam_kidem)

        # ── 6. Cari Hizmet Maliyeti ─────────────────────────────────
        # CSC = 1 yıllık kazanım = PV × (1/toplam_kidem)
        csc = pv_toplam * (1 / toplam_kidem)

        # ── 7. Faiz Maliyeti ────────────────────────────────────────
        # Faiz = DBO(başlangıç) × r
        faiz = dbo * r

        sonuclar.append({
            "Ad Soyad":             row["Ad Soyad"],
            "Yaş":                  yas,
            "Kıdem (yıl)":          kidem,
            "Mevcut Maaş":          round(maaş, 2),
            "Tahmini Maaş (emekli)": round(maas_emekli, 2),
            "Emekliliğe Kalan (t)": t,
            "Hayatta Kalma %":      round(p_hayatta * 100, 1),
            "PV Toplam":            round(pv_toplam, 2),
            "DBO (Yükümlülük)":     round(dbo, 2),
            "Cari Hizmet Maliyeti": round(csc, 2),
            "Faiz Maliyeti":        round(faiz, 2),
            "P&L Etkisi":           round(csc + faiz, 2),
        })

    return pd.DataFrame(sonuclar)


def sensitivity_hesapla(df: pd.DataFrame, r: float, s: float, tavan: float,
                         turnover: float, ret_age: int, mort: float) -> pd.DataFrame:
    """±1% iskonto ve maaş artışı duyarlılık analizi"""
    senaryolar = {
        "Baz Senaryo":          (r,       s),
        "İskonto +1%":          (r+0.01,  s),
        "İskonto -1%":          (r-0.01,  s),
        "Maaş Artışı +1%":      (r,       s+0.01),
        "Maaş Artışı -1%":      (r,       s-0.01),
    }
    rows = []
    baz_dbo = None
    for label, (ri, si) in senaryolar.items():
        res = puc_hesapla(df, ri, si, tavan, turnover, ret_age, mort)
        toplam_dbo = res["DBO (Yükümlülük)"].sum()
        toplam_csc = res["Cari Hizmet Maliyeti"].sum()
        if baz_dbo is None:
            baz_dbo = toplam_dbo
        fark = toplam_dbo - baz_dbo
        fark_pct = (fark / baz_dbo * 100) if baz_dbo else 0
        rows.append({
            "Senaryo":         label,
            "İskonto (%)":     f"{ri*100:.1f}%",
            "Maaş Artışı (%)": f"{si*100:.1f}%",
            "DBO (TL)":        round(toplam_dbo, 2),
            "Fark (TL)":       round(fark, 2),
            "Fark (%)":        round(fark_pct, 2),
        })
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────
# BÖLÜM 4 — HESAPLA butonu
# ─────────────────────────────────────────────────────────────────────
st.header("2️⃣ Hesaplama")

if st.button("🔢 Hesapla", type="primary", use_container_width=True):

    with st.spinner("Hesaplanıyor…"):
        df_sonuc = puc_hesapla(
            df_input, discount_rate, salary_increase, ceiling_2025,
            turnover_rate, retirement_age, mortality_rate
        )
        df_sens = sensitivity_hesapla(
            df_input, discount_rate, salary_increase, ceiling_2025,
            turnover_rate, retirement_age, mortality_rate
        ) if run_sensitivity else None

    st.session_state["df_sonuc"] = df_sonuc
    st.session_state["df_sens"]  = df_sens
    st.session_state["params"]   = {
        "discount_rate":   discount_rate,
        "salary_increase": salary_increase,
        "ceiling_2025":    ceiling_2025,
        "turnover_rate":   turnover_rate,
        "retirement_age":  retirement_age,
        "mortality_rate":  mortality_rate,
    }

# ─────────────────────────────────────────────────────────────────────
# BÖLÜM 5 — SONUÇLAR
# ─────────────────────────────────────────────────────────────────────
if "df_sonuc" in st.session_state:
    df_sonuc = st.session_state["df_sonuc"]
    df_sens  = st.session_state["df_sens"]
    params   = st.session_state["params"]

    st.header("3️⃣ Sonuçlar")

    # KPI kartları
    toplam_dbo  = df_sonuc["DBO (Yükümlülük)"].sum()
    toplam_csc  = df_sonuc["Cari Hizmet Maliyeti"].sum()
    toplam_faiz = df_sonuc["Faiz Maliyeti"].sum()
    toplam_pl   = df_sonuc["P&L Etkisi"].sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Toplam DBO", f"₺{toplam_dbo:,.0f}", help="Tanımlanmış Fayda Yükümlülüğü (bilanço)")
    c2.metric("Cari Hizmet Maliyeti", f"₺{toplam_csc:,.0f}", help="Bu yıl kazanılan yükümlülük artışı")
    c3.metric("Faiz Maliyeti", f"₺{toplam_faiz:,.0f}", help="DBO × iskonto oranı")
    c4.metric("P&L Etkisi", f"₺{toplam_pl:,.0f}", help="Gelir tablosu toplam etkisi (CSC + Faiz)")

    st.divider()

    # ── Sekme yapısı ──────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Çalışan Detayı",
        "📈 DBO Hareketi",
        "🔍 Duyarlılık Analizi",
        "📖 Yöntem Açıklaması"
    ])

    # TAB 1 — Çalışan detayı
    with tab1:
        st.subheader("Çalışan Bazlı PUC Sonuçları")
        fmt = {
            "Mevcut Maaş":             "₺{:,.0f}",
            "Tahmini Maaş (emekli)":   "₺{:,.0f}",
            "PV Toplam":               "₺{:,.0f}",
            "DBO (Yükümlülük)":        "₺{:,.0f}",
            "Cari Hizmet Maliyeti":    "₺{:,.0f}",
            "Faiz Maliyeti":           "₺{:,.0f}",
            "P&L Etkisi":              "₺{:,.0f}",
        }
        st.dataframe(df_sonuc.style.format(fmt), use_container_width=True)

        # DBO dağılımı grafiği
        st.subheader("DBO Dağılımı")
        chart_data = df_sonuc.set_index("Ad Soyad")[["DBO (Yükümlülük)", "Cari Hizmet Maliyeti", "Faiz Maliyeti"]]
        st.bar_chart(chart_data)

    # TAB 2 — DBO Hareketi (Reconciliation)
    with tab2:
        st.subheader("DBO Hareket Tablosu (IAS 19)")
        st.caption("Dönem başı yükümlülükten dönem sonuna geçiş — bilanço dipnotu formatı")

        # Basit hareket tablosu (tek dönem — önceki dönem sıfır kabul)
        onceki_dbo   = 0.0   # ilk yıl / eğitim için sıfır
        cari_hizmet  = toplam_csc
        faiz_maliyet = toplam_faiz
        odemeler     = 0.0   # dönem içi ödenen (eğitimde sıfır)
        akt_kazanc   = 0.0   # aktüeryal kazanç/kayıp (eğitimde sıfır)
        donem_sonu   = onceki_dbo + cari_hizmet + faiz_maliyet - odemeler + akt_kazanc

        hareket = pd.DataFrame({
            "Kalem": [
                "Dönem Başı DBO",
                "  + Cari Hizmet Maliyeti (P&L)",
                "  + Faiz Maliyeti (P&L)",
                "  - Dönem İçi Ödemeler",
                "  ± Aktüeryal Kazanç / Kayıp (OCI)",
                "Dönem Sonu DBO",
            ],
            "Tutar (TL)": [
                onceki_dbo,
                cari_hizmet,
                faiz_maliyet,
                -odemeler,
                akt_kazanc,
                donem_sonu,
            ]
        })
        hareket_fmt = hareket.copy()
        hareket_fmt["Tutar (TL)"] = hareket_fmt["Tutar (TL)"].apply(lambda x: f"₺{x:,.2f}")
        st.table(hareket_fmt)

        st.info(
            "💡 **OCI (Diğer Kapsamlı Gelir):** Gerçekleşen ile beklenen arasındaki fark "
            "aktüeryal kazanç/kayıp olarak OCI'ye yazılır — gelir tablosunu etkilemez."
        )

    # TAB 3 — Duyarlılık
    with tab3:
        st.subheader("Duyarlılık Analizi — ±1% Senaryo")
        if df_sens is not None:
            st.dataframe(
                df_sens.style.format({
                    "DBO (TL)":   "₺{:,.2f}",
                    "Fark (TL)":  "₺{:,.2f}",
                    "Fark (%)":   "{:.2f}%",
                }),
                use_container_width=True
            )
            st.bar_chart(df_sens.set_index("Senaryo")["DBO (TL)"])
            st.caption(
                "IAS 19 Para. 145: Her önemli aktüeryal varsayım için makul ölçüde mümkün "
                "değişikliklerin etkisi dipnotlarda açıklanmalıdır."
            )
        else:
            st.info("Duyarlılık analizi kapalı. Sidebar'dan aktifleştirin.")

    # TAB 4 — Yöntem açıklaması
    with tab4:
        st.subheader("Projected Unit Credit (PUC) Yöntemi — IAS 19")
        st.markdown("""
**PUC neden kullanılır?**
IAS 19 paragraf 67, tanımlanmış fayda planlarında *yalnızca* Projected Unit Credit yöntemini
zorunlu kılar. Bu yöntemde her çalışanın birikmiş hizmetine karşılık gelen yükümlülük ayrı
ayrı ölçülür.

---

**Formüller:**

| Büyüklük | Formül |
|---|---|
| Tahmini maaş (emeklilik) | `maaş × (1 + s)^t` |
| Beklenen ödeme (emeklilik) | `min(maaş_emekli, tavan) × toplam_kıdem_yılı` |
| Hayatta kalma olasılığı | `(1 − ölüm − ayrılma)^t` |
| Bugünkü değer (PV) | `beklenen_ödeme × p / (1+r)^t` |
| **DBO** (bilanço) | `PV × (geçmiş_kıdem / toplam_beklenen_kıdem)` |
| Cari hizmet maliyeti | `PV × (1 / toplam_beklenen_kıdem)` |
| Faiz maliyeti | `DBO × r` |

---

**Gelir tablosu vs OCI ayrımı:**
- **P&L:** Cari hizmet maliyeti + Faiz maliyeti
- **OCI:** Aktüeryal kazanç/kayıplar (gerçekleşme ile beklenti farkı)
- **Bilanço:** Net DBO (varlık veya yükümlülük olarak)

---

**Türkiye'ye özgü notlar:**
- Kıdem tazminatı tavanı her Ocak ve Temmuz güncellenir
- Ertelenmiş vergi: Geçici fark × kurumlar vergisi oranı (%25)
- 1 yılı doldurmayan çalışanlar genellikle hesaba dahil edilmez
        """)

    # ─────────────────────────────────────────────────────────────
    # BÖLÜM 6 — EXCEL RAPORU
    # ─────────────────────────────────────────────────────────────
    st.header("4️⃣ Rapor İndir")

    def excel_olustur(df_sonuc, df_sens, params):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:

            # Sayfa 1: Özet
            ozet = pd.DataFrame({
                "Parametre": [
                    "Rapor Tarihi", "İskonto Oranı", "Maaş Artış Oranı",
                    "Kıdem Tazminatı Tavanı", "İşten Ayrılma Olasılığı",
                    "Emeklilik Yaşı", "Ölüm Olasılığı",
                    "---", "Toplam DBO", "Toplam CSC", "Toplam Faiz", "P&L Etkisi"
                ],
                "Değer": [
                    str(date.today()),
                    f"{params['discount_rate']*100:.1f}%",
                    f"{params['salary_increase']*100:.1f}%",
                    f"₺{params['ceiling_2025']:,.2f}",
                    f"{params['turnover_rate']*100:.1f}%",
                    str(params['retirement_age']),
                    f"{params['mortality_rate']*100:.2f}%",
                    "",
                    f"₺{df_sonuc['DBO (Yükümlülük)'].sum():,.2f}",
                    f"₺{df_sonuc['Cari Hizmet Maliyeti'].sum():,.2f}",
                    f"₺{df_sonuc['Faiz Maliyeti'].sum():,.2f}",
                    f"₺{df_sonuc['P&L Etkisi'].sum():,.2f}",
                ]
            })
            ozet.to_excel(writer, sheet_name="Özet", index=False)

            # Sayfa 2: Çalışan detayı
            df_sonuc.to_excel(writer, sheet_name="Çalışan Detayı", index=False)

            # Sayfa 3: DBO Hareketi
            hareket = pd.DataFrame({
                "Kalem": ["Dönem Başı DBO", "Cari Hizmet Maliyeti", "Faiz Maliyeti",
                           "Dönem İçi Ödemeler", "Aktüeryal Kazanç/Kayıp", "Dönem Sonu DBO"],
                "Tutar": [0, df_sonuc['Cari Hizmet Maliyeti'].sum(),
                           df_sonuc['Faiz Maliyeti'].sum(), 0, 0,
                           df_sonuc['DBO (Yükümlülük)'].sum()]
            })
            hareket.to_excel(writer, sheet_name="DBO Hareketi", index=False)

            # Sayfa 4: Duyarlılık
            if df_sens is not None:
                df_sens.to_excel(writer, sheet_name="Duyarlılık Analizi", index=False)

        buf.seek(0)
        return buf

    # ─────────────────────────────────────────────────────────────
    # BÖLÜM 7 — PDF RAPORU
    # ─────────────────────────────────────────────────────────────
    def pdf_olustur(df_sonuc, df_sens, params):
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                             Table, TableStyle, HRFlowable)
            from reportlab.lib.enums import TA_CENTER, TA_LEFT

            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=A4,
                                    leftMargin=2*cm, rightMargin=2*cm,
                                    topMargin=2*cm, bottomMargin=2*cm)
            styles = getSampleStyleSheet()
            story  = []

            # Başlık stili
            title_style = ParagraphStyle("title", parent=styles["Title"],
                                         fontSize=16, spaceAfter=6, alignment=TA_CENTER)
            h2_style    = ParagraphStyle("h2", parent=styles["Heading2"],
                                         fontSize=12, spaceAfter=4, spaceBefore=12)
            body_style  = styles["Normal"]
            body_style.fontSize = 9

            # Başlık
            story.append(Paragraph("KIDEM TAZMİNATI AKTÜERYAL RAPORU", title_style))
            story.append(Paragraph("IAS 19 / TMS 19 — Projected Unit Credit Yöntemi", styles["Normal"]))
            story.append(Paragraph(f"Rapor Tarihi: {date.today().strftime('%d.%m.%Y')}", styles["Normal"]))
            story.append(HRFlowable(width="100%", thickness=1, color=colors.darkblue))
            story.append(Spacer(1, 0.3*cm))

            # Varsayımlar tablosu
            story.append(Paragraph("1. Aktüeryal Varsayımlar", h2_style))
            veri = [
                ["Parametre", "Değer"],
                ["İskonto Oranı",           f"{params['discount_rate']*100:.1f}%"],
                ["Maaş Artış Oranı",        f"{params['salary_increase']*100:.1f}%"],
                ["Kıdem Tavanı",            f"₺{params['ceiling_2025']:,.2f}"],
                ["Gönüllü Ayrılma Oranı",   f"{params['turnover_rate']*100:.1f}%"],
                ["Emeklilik Yaşı",          str(params['retirement_age'])],
                ["Ölüm Olasılığı",          f"{params['mortality_rate']*100:.2f}%"],
            ]
            t = Table(veri, colWidths=[9*cm, 7*cm])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.darkblue),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                ("FONTSIZE",   (0,0), (-1,-1), 9),
                ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4ff")]),
                ("ALIGN",      (1,0), (1,-1), "RIGHT"),
            ]))
            story.append(t)
            story.append(Spacer(1, 0.5*cm))

            # Özet sonuçlar
            story.append(Paragraph("2. Özet Sonuçlar", h2_style))
            ozet_veri = [
                ["Büyüklük", "Tutar (TL)"],
                ["Toplam DBO (Bilanço Yükümlülüğü)",  f"₺{df_sonuc['DBO (Yükümlülük)'].sum():>15,.2f}"],
                ["Cari Hizmet Maliyeti (P&L)",         f"₺{df_sonuc['Cari Hizmet Maliyeti'].sum():>15,.2f}"],
                ["Faiz Maliyeti (P&L)",                f"₺{df_sonuc['Faiz Maliyeti'].sum():>15,.2f}"],
                ["Toplam P&L Etkisi",                  f"₺{df_sonuc['P&L Etkisi'].sum():>15,.2f}"],
            ]
            t2 = Table(ozet_veri, colWidths=[10*cm, 6*cm])
            t2.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.darkblue),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                ("FONTSIZE",   (0,0), (-1,-1), 9),
                ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4ff")]),
                ("ALIGN",      (1,1), (1,-1), "RIGHT"),
                ("FONTNAME",   (0,-1), (-1,-1), "Helvetica-Bold"),
            ]))
            story.append(t2)
            story.append(Spacer(1, 0.5*cm))

            # Çalışan detayı
            story.append(Paragraph("3. Çalışan Bazlı DBO", h2_style))
            col_header = ["Ad Soyad", "Yaş", "Kıdem", "DBO (TL)", "CSC (TL)", "Faiz (TL)"]
            rows = [col_header]
            for _, r in df_sonuc.iterrows():
                rows.append([
                    str(r["Ad Soyad"]),
                    str(r["Yaş"]),
                    str(r["Kıdem (yıl)"]),
                    f"₺{r['DBO (Yükümlülük)']:,.0f}",
                    f"₺{r['Cari Hizmet Maliyeti']:,.0f}",
                    f"₺{r['Faiz Maliyeti']:,.0f}",
                ])
            t3 = Table(rows, colWidths=[5*cm, 1.5*cm, 1.5*cm, 3.5*cm, 3*cm, 2.5*cm])
            t3.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.darkblue),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                ("FONTSIZE",   (0,0), (-1,-1), 8),
                ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
                ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4ff")]),
                ("ALIGN",      (1,0), (-1,-1), "RIGHT"),
            ]))
            story.append(t3)
            story.append(Spacer(1, 0.5*cm))

            # Duyarlılık
            if df_sens is not None:
                story.append(Paragraph("4. Duyarlılık Analizi (±1%)", h2_style))
                sens_header = ["Senaryo", "İskonto", "Maaş Artışı", "DBO (TL)", "Fark (TL)", "Fark (%)"]
                sens_rows   = [sens_header]
                for _, r in df_sens.iterrows():
                    sens_rows.append([
                        str(r["Senaryo"]),
                        str(r["İskonto (%)"]),
                        str(r["Maaş Artışı (%)"]),
                        f"₺{r['DBO (TL)']:,.0f}",
                        f"₺{r['Fark (TL)']:,.0f}",
                        f"{r['Fark (%)']:.2f}%",
                    ])
                t4 = Table(sens_rows, colWidths=[4.5*cm, 2*cm, 2.5*cm, 3.5*cm, 3*cm, 1.5*cm])
                t4.setStyle(TableStyle([
                    ("BACKGROUND", (0,0), (-1,0), colors.darkblue),
                    ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
                    ("FONTSIZE",   (0,0), (-1,-1), 8),
                    ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
                    ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f0f4ff")]),
                    ("ALIGN",      (1,0), (-1,-1), "RIGHT"),
                    ("FONTNAME",   (0,1), (0,1), "Helvetica-Bold"),  # baz senaryo bold
                ]))
                story.append(t4)

            # Dipnot
            story.append(Spacer(1, 1*cm))
            story.append(HRFlowable(width="100%", thickness=0.5, color=colors.grey))
            story.append(Paragraph(
                "Bu rapor IAS 19 / TMS 19 kapsamında Projected Unit Credit yöntemiyle hazırlanmıştır. "
                "Aktüeryal değerleme, lisanslı bir aktüer tarafından onaylanmalıdır.",
                ParagraphStyle("footer", parent=styles["Normal"], fontSize=7, textColor=colors.grey)
            ))

            doc.build(story)
            buf.seek(0)
            return buf

        except ImportError:
            return None

    col_xl, col_pdf = st.columns(2)

    with col_xl:
        xl_buf = excel_olustur(df_sonuc, df_sens, params)
        st.download_button(
            label="⬇️ Excel Raporu İndir (.xlsx)",
            data=xl_buf,
            file_name=f"kidem_tazminati_raporu_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_pdf:
        pdf_buf = pdf_olustur(df_sonuc, df_sens, params)
        if pdf_buf:
            st.download_button(
                label="⬇️ PDF Aktüeryal Rapor İndir (.pdf)",
                data=pdf_buf,
                file_name=f"kidem_tazminati_aktueryal_{date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        else:
            st.warning("PDF için: `pip install reportlab`")
