from io import BytesIO
import json
import os
import tempfile
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
from pypdf import PdfReader
import streamlit as st

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim", page_icon="✏️", layout="wide"
)

# --- GEMINI CLIENT BAĞLANTISI ---
try:
  api_key = st.secrets["API_KEY"]
  client = genai.Client(api_key=api_key)
except Exception as e:
  st.error(
      "API Anahtarı bulunamadı! Lütfen Streamlit Cloud Secrets ayarlarından"
      " 'API_KEY' tanımladığınızdan emin olun."
  )
  st.stop()


# --- YARDIMCI FONKSİYONLAR ---
def pdf_oku(dosya):
  reader = PdfReader(dosya)
  metin = ""
  for sayfa in reader.pages:
    metin += sayfa.extract_text() or ""
  return metin


def docx_oku(dosya):
  doc = Document(dosya)
  metin = "\n".join([p.text for p.text in doc.paragraphs])
  return metin


@st.cache_data
def verileri_yukle():
  kazanimlar = []
  referanslar = []
  try:
    if os.path.exists("kazanimlar.json"):
      with open("kazanimlar.json", "r", encoding="utf-8") as f:
        kazanimlar = json.load(f)
  except Exception:
    pass
  try:
    if os.path.exists("referans_sorular.json"):
      with open("referans_sorular.json", "r", encoding="utf-8") as f:
        referanslar = json.load(f)
  except Exception:
    pass
  return kazanimlar, referanslar


kazanimlar_data, referanslar_data = verileri_yukle()


def soru_word_yap(soru_metni):
  doc = Document()
  doc.add_heading("Kurşun Kalem Dijital Eğitim - Soru Çıktısı", level=1)
  doc.add_paragraph(soru_metni)
  f = BytesIO()
  doc.save(f)
  f.seek(0)
  return f


# --- SINIF VE KONU YAPISI ---
SINIF_KONTENJLARI = {
    "5. Sınıf Türkçe": [
        "Sözcükte Anlam (Gerçek, Mecaz, Terim)",
        "Cümlede Anlam (Öznellik, Nesnellik, Neden-Sonuç)",
        "Parçada Anlam (Ana Düşünce, Konu, Başlık)",
        "Metin Türleri (Hikaye, Masal, Fabl)",
        "Yazım Kuralları ve Noktalama İşaretleri",
    ],
    "6. Sınıf Türkçe": [
        "Sözcükte Anlam (Eş, Zıt, Yakın Anlamlılar, Deyimler)",
        "Cümlede Anlam (Amaç-Sonuç, Koşul-Sonuç)",
        "Parçada Anlam (Ana Duygu, Düşünceyi Geliştirme Yolları)",
        "İsimler (Adlar) ve Tamlamalar",
        "Zamirler (Adıllar) ve Ek Fiil",
    ],
    "7. Sınıf Türkçe": [
        "Söz Sanatları (Kişileştirme, Benzetme, Abartma, Konuşturma)",
        "Fiiller (Kip, Kişi, Anlam Kayması)",
        "Fiilimsiler (İsim-Fiil, Sıfat-Fiil, Zarf-Fiil)",
        "Cümlenin Ögeleri (Özne, Nesne, Dolaylı Tümleç, Yüklem)",
        "Anlatım Bozuklukları (Anlamsal Bozukluklar)",
    ],
    "8. Sınıf Türkçe (LGS Hazırlık)": [
        "Fiilimsiler ve Cümlenin Ögeleri (İleri Düzey)",
        "Cümle Vurgusu ve Çatı Özellikleri",
        "Metin Türleri ve Düşünceyi Geliştirme Yolları",
        "Sözel Mantık ve Muhakeme Becerileri",
        "Grafik, Tablo ve Görsel Okuma Yorumlama",
        "Paragraf Stratejileri (Ana Düşünce, Yardımcı Düşünce, Akışı Bozan)",
    ],
}

# --- ARAYÜZ BAŞLIĞI ---
st.title(
    "✏️ Kurşun Kalem Dijital Eğitim - Sınıf Bazlı Soru ve İçerik Yönetim Sistemi"
)
st.markdown(
    "5, 6, 7 ve 8. Sınıf müfredatına uygun olarak konularınızı yönetin,"
    " sorular üretin ve ders videolarınızı analiz edin."
)

# --- SEKMELER ---
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📝 Sınıf & LGS Soru Üretici",
        "📂 Sınıf & Konu Dosya Yönetimi",
        "📊 Soru Havuzu Analizi",
        "🎥 Sınıf Bazlı Video Analiz Paneli",
    ]
)

with tab1:
  st.header("Beceri Temelli Sınıf & LGS Soru Üretici")

  col_s1, col_s2 = st.columns(2)
  with col_s1:
    secilen_sinif_s = st.selectbox(
        "Hedef Sınıf Seviyesi:", list(SINIF_KONTENJLARI.keys()), key="soru_sinif"
    )
  with col_s2:
    secilen_konu_s = st.selectbox(
        "İlgili Konu / Kazanım:",
        SINIF_KONTENJLARI[secilen_sinif_s],
        key="soru_konu",
    )

  ek_detay = st.text_area(
      "Özel İstem / Ek Açıklama (İsteğe Bağlı):",
      placeholder=(
          "Örn: Sadece uzun paragraflı yeni nesil soru olsun veya şu kuralı"
          " vurgula."
      ),
      height=70,
  )

  col1, col2, col3 = st.columns(3)
  with col1:
    zorluk_derecesi = st.selectbox(
        "Zorluk Seviyesi", ["Beceri Temelli (Yeni Nesil)", "Zor", "Orta"]
    )
  with col2:
    soru_sayisi = st.number_input(
        "Üretilecek Soru Sayısı", min_value=1, max_value=5, value=1
    )
  with col3:
    soru_turu = st.selectbox(
        "Soru Tipi",
        ["Çoktan Seçmeli (4 Şıklı)", "Açık Uçlu", "Metne Dayalı Yorum"],
    )

  if st.button("Soruları Üret", type="primary"):
    with st.spinner("Yapay zekâ soruları hazırlıyor, lütfen bekleyin..."):
      try:
        prompt = f"""
                  Sen uzman bir Türkçe eğitimcisi ve soru hazırlama komisyonu başkanısın.
                  Aşağıdaki kriterlere tam uygun olacak şekilde soru/sorular hazarla:
                  - Sınıf Düzeyi: {secilen_sinif_s}
                  - Konu / Kazanım: {secilen_konu_s}
                  - Ek Açıklama: {ek_detay}
                  - Zorluk: {zorluk_derecesi}
                  - Soru Sayısı: {soru_sayisi}
                  - Soru Tipi: {soru_turu}
                  - MEB kazanımlarına ve LGS/okul sınavı mantığına tam uygun, öğrencinin okuduğunu anlama ve yorumlama becerilerini ölçen
