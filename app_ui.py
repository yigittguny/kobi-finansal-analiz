import io
import os
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


st.set_page_config(page_title="KOBİ Finansal Analiz", layout="centered")
st.title("KOBİ Finansal Analiz Sistemi")
st.caption("Excel yükle → analiz al → grafikleri gör → PDF raporu indir")

uploaded_file = st.file_uploader("Excel Dosyası Yükle (.xlsx)", type=["xlsx"])


def register_turkish_font():
    """
    Windows'ta Arial kullanır (Türkçe karakterleri düzgün basar).
    Bulamazsa Helvetica'ya düşer (Türkçe yine bozulabilir ama uygulama çökmez).
    """
    font_path = r"C:\Windows\Fonts\arial.ttf"
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont("TRFont", font_path))
            return "TRFont"
        except Exception:
            return "Helvetica"
    return "Helvetica"


if uploaded_file is None:
    st.info("Başlamak için bir Excel dosyası yükle. Kolonlar: Satis, Maliyet")
    st.stop()

# Excel oku
try:
    df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Excel okunamadı: {e}")
    st.stop()

# Kolon kontrolü
required_cols = {"Satis", "Maliyet"}
if not required_cols.issubset(set(df.columns)):
    st.error("Excel dosyasında şu kolonlar olmalı: Satis, Maliyet")
    st.stop()

# Temizlik
df = df.copy()
df["Satis"] = pd.to_numeric(df["Satis"], errors="coerce")
df["Maliyet"] = pd.to_numeric(df["Maliyet"], errors="coerce")
df = df.dropna(subset=["Satis", "Maliyet"])

if df.empty:
    st.error("Geçerli veri yok. Satis ve Maliyet sayısal olmalı.")
    st.stop()

# Hesaplamalar
df["Kar"] = df["Satis"] - df["Maliyet"]
df["Kar_Marji"] = 0.0
df.loc[df["Satis"] != 0, "Kar_Marji"] = (df.loc[df["Satis"] != 0, "Kar"] / df.loc[df["Satis"] != 0, "Satis"]) * 100

toplam_satis = float(df["Satis"].sum())
toplam_kar = float(df["Kar"].sum())
ortalama_marj = float(df["Kar_Marji"].mean())

st.subheader("📊 Finansal Özet")
c1, c2, c3 = st.columns(3)
c1.metric("Toplam Satış", round(toplam_satis, 2))
c2.metric("Toplam Kâr", round(toplam_kar, 2))
c3.metric("Ortalama Kâr Marjı (%)", round(ortalama_marj, 2))

# Risk
st.subheader("⚠ Risk Analizi")
if ortalama_marj < 20:
    risk = "YÜKSEK"
    st.error("Risk Seviyesi: YÜKSEK")
    tavsiye = "Maliyetleri kontrol et. Fiyatlandırmayı ve tedarik maliyetlerini gözden geçir."
elif ortalama_marj < 35:
    risk = "ORTA"
    st.warning("Risk Seviyesi: ORTA")
    tavsiye = "Marjı artırmak için operasyon ve satın alma süreçlerinde optimizasyon yap."
else:
    risk = "DÜŞÜK"
    st.success("Risk Seviyesi: DÜŞÜK")
    tavsiye = "Genel tablo sağlıklı. Ölçekleme ve büyüme stratejileri planlanabilir."

st.write("Tavsiye:", tavsiye)

# İçgörüler
st.subheader("🔎 Kritik İçgörüler")
en_karli = df.loc[df["Kar"].idxmax()]
en_dusuk_marj = df.loc[df["Kar_Marji"].idxmin()]

st.success("En Karlı Satır")
st.write(en_karli)

st.error("En Düşük Marjlı Satır")
st.write(en_dusuk_marj)

# Grafikler
st.subheader("📈 Grafikler")

fig1, ax1 = plt.subplots()
ax1.bar(df.index.astype(str), df["Kar"])
ax1.set_xlabel("Satır")
ax1.set_ylabel("Kâr")
ax1.set_title("Satır Bazlı Kâr Analizi")
st.pyplot(fig1)

fig2, ax2 = plt.subplots()
ax2.plot(df.index.astype(str), df["Kar_Marji"], marker="o")
ax2.set_xlabel("Satır")
ax2.set_ylabel("Kâr Marjı (%)")
ax2.set_title("Satır Bazlı Kâr Marjı")
st.pyplot(fig2)

# PDF
st.subheader("📄 PDF Rapor")

if st.button("PDF Raporu Oluştur"):
    base_font = register_turkish_font()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer)
    styles = getSampleStyleSheet()

    # Fontu tüm stillere uygula
    for key in ["Normal", "Heading1", "Heading2", "Heading3"]:
        if key in styles:
            styles[key].fontName = base_font

    elements = []
    elements.append(Paragraph("KOBİ Finansal Analiz Raporu", styles["Heading1"]))
    elements.append(Spacer(1, 0.25 * inch))

    # Özet tablo
    summary_data = [
        ["Toplam Satış", f"{round(toplam_satis, 2)}"],
        ["Toplam Kâr", f"{round(toplam_kar, 2)}"],
        ["Ortalama Kâr Marjı", f"%{round(ortalama_marj, 2)}"],
        ["Risk Seviyesi", risk],
    ]
    summary_table = Table(summary_data, colWidths=[180, 300])
    summary_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), base_font),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 0.2 * inch))

    # Tavsiye metni
    elements.append(Paragraph("Tavsiye", styles["Heading3"]))
    elements.append(Paragraph(tavsiye, styles["Normal"]))
    elements.append(Spacer(1, 0.2 * inch))

    # İçgörüler tablosu
    elements.append(Paragraph("Kritik İçgörüler", styles["Heading3"]))

    insights_data = [
        ["", "Satış", "Maliyet", "Kâr", "Kâr Marjı (%)"],
        [
            "En Karlı Satır",
            f"{round(float(en_karli['Satis']), 2)}",
            f"{round(float(en_karli['Maliyet']), 2)}",
            f"{round(float(en_karli['Kar']), 2)}",
            f"{round(float(en_karli['Kar_Marji']), 2)}",
        ],
        [
            "En Düşük Marjlı Satır",
            f"{round(float(en_dusuk_marj['Satis']), 2)}",
            f"{round(float(en_dusuk_marj['Maliyet']), 2)}",
            f"{round(float(en_dusuk_marj['Kar']), 2)}",
            f"{round(float(en_dusuk_marj['Kar_Marji']), 2)}",
        ],
    ]
    insights_table = Table(insights_data, colWidths=[170, 90, 90, 90, 120])
    insights_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), base_font),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(insights_table)

    doc.build(elements)
    pdf_bytes = buffer.getvalue()
    buffer.close()

    st.download_button(
        label="PDF Raporu İndir",
        data=pdf_bytes,
        file_name="kobi_finansal_analiz_raporu.pdf",
        mime="application/pdf",
    )