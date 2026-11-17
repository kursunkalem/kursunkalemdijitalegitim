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
  metin = "\n".join([p.text for p in doc.paragraphs])
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


# --- ARAYÜZ BAŞLIĞI ---
st.title(
    "✏️ Kurşun Kalem Dijital Eğitim - Dijital Soru ve Video Ders Platformu"
)
st.markdown(
    "LGS standartlarına uygun sorular üretin, referans belgelerinizi yönetin"
    " ve ders videolarını canlı olarak izleyip yapay zekâyla analiz edin."
)

# --- SEKMELER ---
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📝 Soru Üretim Modülü",
        "📂 Kazanım & Referans Yönetimi",
        "📊 Soru Havuzu Analizi",
        "🎥 Video Kütüphanesi ve Oynatıcı",
    ]
)

with tab1:
  st.header("Beceri Temelli Soru Üretici (LGS Standartları)")
  kazanim_input = st.text_area(
      "Kazanım veya Konu İçeriğini Girin:",
      placeholder="Örn: Metnin ana düşüncesini ve yardımcı düşüncelerini ayırt eder.",
      height=100,
  )

  col1, col2, col3 = st.columns(3)
  with col1:
    zorluk_derecesi = st.selectbox(
        "Zorluk Seviyesi", ["Beceri Temelli (LGS)", "Zor", "Orta"]
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
    if not kazanim_input or len(kazanim_input.strip()) < 3:
      st.warning(
          "Lütfen soru üretimi için daha detaylı bir kazanım veya konu"
          " içeriği girin."
      )
    else:
      with st.spinner("Yapay zekâ soruları hazırlıyor, lütfen bekleyin..."):
        try:
          prompt = f"""
                    Sen uzman bir Türkçe eğitimcisi ve LGS soru hazırlama komisyonu başkanısın.
                    Aşağıdaki kriterlere tam uygun olacak şekilde soru/sorular hazarla:
                    - Kazanım/İçerik: {kazanim_input}
                    - Zorluk: {zorluk_derecesi}
                    - Soru Sayısı: {soru_sayisi}
                    - Soru Tipi: {soru_turu}
                    - LGS formatına uygun, öğrencinin okuduğunu anlama ve yorumlama becerilerini ölçen nitelikli öncüller kullan.
                    - Cevap anahtarını ve detaylı açıklamasını mutlaka ekle.
                    """
          response = client.models.generate_content(
              model="gemini-2.5-flash",
              contents=prompt,
              config=types.GenerateContentConfig(
                  temperature=0.4, max_output_tokens=2048
              ),
          )
          st.success("Sorular başarıyla üretildi!")
          st.markdown(response.text)

          word_dosyasi = soru_word_yap(response.text)
          st.download_button(
              label="📥 Soruları Word (DOCX) Olarak İndir",
              data=word_dosyasi,
              file_name="kursunkalem_lgs_sorulari.docx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              ),
          )
        except Exception as e:
          st.error(f"Soru üretilirken hata oluştu: {e}")

with tab2:
  st.header("📂 Kazanım & Referans Dosya Yönetimi")
  yuklenen_dosya = st.file_uploader(
      "Harici Soru veya Ders Notu Yükle (PDF, DOCX, TXT)",
      type=["pdf", "docx", "txt"],
  )
  if yuklenen_dosya is not None:
    uzanti = yuklenen_dosya.name.split(".")[-1].lower()
    metin_icerik = (
        pdf_oku(yuklenen_dosya)
        if uzanti == "pdf"
        else (
            docx_oku(yuklenen_dosya)
            if uzanti == "docx"
            else yuklenen_dosya.read().decode("utf-8")
        )
    )
    st.text_area(
        "Belge İçeriği Önizlemesi:", value=metin_icerik[:3000], height=200
    )

    if st.button("Belgeyi Analiz Et"):
      with st.spinner("Analiz ediliyor..."):
        p = "Bu dokümanı incele, LGS Türkçe müfredatı bağlamında ana kazanımları çıkar."
        resp = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[p, metin_icerik[:3000]],
            config=types.GenerateContentConfig(
                temperature=0.3, max_output_tokens=1500
            ),
        )
        st.markdown(resp.text)

with tab3:
  st.header("📊 Soru Havuzu Analiz Paneli")
  if referanslar_data:
    st.write(f"Toplam {len(referanslar_data)} referans soru aktif.")
    st.json(referanslar_data[:2])
  else:
    st.warning("Referans soru dosyası bulunamadı.")

with tab4:
  st.header("🎥 Video Kütüphanesi ve Canlı Oynatıcı")
  st.markdown(
      "İster sitede sabit yer alan YouTube ders videolarını seçerek izleyin, ister"
      " bilgisayarınızdan kendi videonuzu yükleyip yapay zekâya analiz"
      " ettirin."
  )

  video_secimi = st.radio(
      "Video Modu:",
      [
          "📺 Sitedeki Sabit Ders Videoları (YouTube)",
          "💻 Bilgisayarımdan Video Yükle ve İzle",
      ],
  )

  if video_secimi == "📺 Sitedeki Sabit Ders Videoları (YouTube)":
    secilen_youtube_ders = st.selectbox(
        "Sitedeki Ders Listesinden Seçin:",
        [
            "Örnek Ders 1: Paragraf Stratejileri",
            "Örnek Ders 2: Fiilimsiler Konu Anlatımı",
        ],
    )

    col_y1, col_y2 = st.columns([1.2, 1])
    with col_y1:
      st.subheader(f"▶️ Oynatılıyor: {secilen_youtube_ders}")
      if "Paragraf" in secilen_youtube_ders:
        st.video("https://www.youtube.com/watch?v=dQw4w9WgXcQ")
      else:
        st.video("https://www.youtube.com/watch?v=dQw4w9WgXcQ")

    with col_y2:
      st.subheader("⚙️ Bilgi")
      st.info(
          "Bu bölüm sitede sabit yer alan dersler içindir. Kendi yükleyeceğiniz"
          " videoları analiz ettirmek için diğer modu seçebilirsiniz."
      )

  else:
    uploaded_video = st.file_uploader(
        "Bilgisayardan Ders Videosu Seç (MP4, MOV, AVI)",
        type=["mp4", "mov", "avi"],
    )

    if uploaded_video is not None:
      col_video, col_panel = st.columns([1.2, 1])

      with col_video:
        st.subheader("▶️ Sitedeki Canlı Oynatıcı")
        st.video(uploaded_video)

      with col_panel:
        st.subheader("⚙️ Yapay Zekâ Analiz Paneli")
        islem_turu = st.selectbox(
            "İşlem Seçin:",
            [
                "Ders Özeti ve Ana Başlıklar",
                "Videodan Beceri Temelli Soru Üret",
                "Video Konuşma Dökümü (Transkript)",
            ],
        )

        if st.button("Analizi Başlat", type="primary"):
          with st.spinner("Video işleniyor ve analiz ediliyor..."):
            try:
              with tempfile.NamedTemporaryFile(
                  delete=False, suffix=".mp4"
              ) as tmp_file:
                uploaded_video.seek(0)
                tmp_file.write(uploaded_video.read())
                tmp_file_path = tmp_file.name

              video_file = client.files.upload(file=tmp_file_path)

              if islem_turu == "Ders Özeti ve Ana Başlıklar":
                video_prompt = (
                    "Bu ders videosunu analiz et. Öğrenciler için maddeler halinde"
                    " detaylı bir ders özeti çıkar."
                )
              elif islem_turu == "Videodan Beceri Temelli Soru Üret":
                video_prompt = (
                    "Bu ders videosundaki konuya dayanarak LGS formatında"
                    " beceri temelli sorular ve cevap anahtarı hazarla."
                )
              else:
                video_prompt = (
                    "Bu videodaki konuşmaları kronolojik sırayla metin"
                    " dökümü (transkript) haline getir."
                )

              response = client.models.generate_content(
                  model="gemini-2.5-flash",
                  contents=[video_file, video_prompt],
                  config=types.GenerateContentConfig(
                      temperature=0.3, max_output_tokens=2048
                  ),
              )

              st.success("Analiz tamamlandı!")
              st.markdown(response.text)

              os.unlink(tmp_file_path)

            except Exception as e:
              st.error(f"Hata oluştu: {e}")
