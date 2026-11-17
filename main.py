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

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim", page_icon="✏️", layout="wide"
)

# --- GEMINI CLIENT BAĞLANTISI ---
# Streamlit Secrets üzerinden API anahtarını güvenli bir şekilde alıyoruz
try:
  api_key = st.secrets["API_KEY"]
  client = genai.Client(api_key=api_key)
except Exception as e:
  st.error(
      "API Anahtarı bulunamadı! Lütfen Streamlit Cloud Secrets ayarlarından"
      " 'API_KEY' tanımladığınızdan emin olun."
  )
  st.stop()

# --- BAŞLIK VE ARAYÜZ ---
st.title("✏️ Kurşun Kalem Dijital Eğitim - Soru Üretim ve Video Analiz Paneli")
st.markdown(
    "LGS standartlarına uygun beceri temelli sorular üretin ve ders videolarınızı"
    " yapay zekâ ile analiz edin."
)

# Sekmeler (Tab Yapısı)
tab1, tab2 = st.tabs(["📝 Soru Üretim Modülü", "🎥 Yapay Zekâ Video Analizi"])

with tab1:
  st.header("Beceri Temelli Soru Üretici")

  # Örnek Alanlar (Mevcut kodun yapısına uygun)
  kazanim_input = st.text_area(
      "Kazanım veya Konu İçeriğini Girin:",
      placeholder="Örn: Metnin ana düşüncesini ve yardımcı düşüncelerini ayırt eder.",
  )

  col1, col2 = st.columns(2)
  with col1:
    zorluk_derecesi = st.selectbox(
        "Zorluk Seviyesi", ["Orta", "Zor", "Beceri Temelli (LGS)"]
    )
  with col2:
    soru_sayisi = st.number_input(
        "Üretilecek Soru Sayısı", min_value=1, max_value=5, value=1
    )

  if st.button("Soruları Üret", type="primary"):
    if not kazanim_input:
      st.warning("Lütfen bir kazanım veya konu içeriği girin.")
    else:
      with st.spinner(
          "Yapay zekâ LGS formatında soruları hazırlıyor, lütfen bekleyin..."
      ):
        try:
          prompt = f"""
                    Aşağıdaki kazanım doğrultusunda {zorluk_derecesi} seviyesinde {soru_sayisi} adet LGS formatında beceri temelli Türkçe sorusu hazırla.
                    Kazanım/İçerik: {kazanim_input}
                    """
          response = client.models.generate_content(
              model="gemini-2.5-flash",
              contents=prompt,
              config=types.GenerateContentConfig(
                  temperature=0.4,
                  max_output_tokens=2048,
              ),
          )
          st.success("Sorular başarıyla üretildi!")
          st.markdown(response.text)
        except Exception as e:
          st.error(f"Soru üretilirken bir hata oluştu: {e}")

with tab2:
  st.header("🎥 Yapay Zekâ Destekli Ders Videosu Analizi")
  st.markdown(
      "Bir ders videosu yükleyerek videonun dökümünü çıkarabilir, özetini"
      " alabilir veya doğrudan videoya dayalı sorular üretebilirsiniz."
  )

  # Video Dosyası Yükleme
  uploaded_video = st.file_uploader(
      "Ders Videosu Yükleyin (MP4, MOV, AVI)", type=["mp4", "mov", "avi"]
  )

  if uploaded_video is not None:
    # Videoyu ekranda oynat
    st.video(uploaded_video)

    islem_turu = st.selectbox(
        "Yapılacak İşlem",
        [
            "Ders Özeti ve Ana Başlıklar",
            "Videodan Beceri Temelli Soru Üret",
            "Video Dökümü (Transkript) Çıkar",
        ],
    )

    if st.button("Video İşlemini Başlat", type="primary"):
      with st.spinner(
          "Video geçici olarak işleniyor ve yapay zekâya gönderiliyor..."
      ):
        try:
          # Yüklenen videoyu geçici bir dosyaya kaydediyoruz (Gemini File API için gerekli)
          with tempfile.NamedTemporaryFile(
              delete=False, suffix=".mp4"
          ) as tmp_file:
            tmp_file.write(uploaded_video.read())
            tmp_file_path = tmp_file.name

          # Videoyu Gemini File API'ye yükle
          st.info(
              "Video Gemini bulutuna yükleniyor, dosya boyutuna göre birkaç"
              " saniye sürebilir..."
          )
          video_file = client.files.upload(file=tmp_file_path)

          # İşlem türüne göre prompt belirleme
          if islem_turu == "Ders Özeti ve Ana Başlıklar":
            video_prompt = (
                "Bu ders videosunu dikkatlice izle/analiz et. Öğrenciler için"
                " anlaşılır, maddeler halinde detaylı bir ders özeti ve ana"
                " başlıklar çıkar."
            )
          elif islem_turu == "Videodan Beceri Temelli Soru Üret":
            video_prompt = (
                "Bu ders videosunda anlatılan konuya dayanarak LGS formatında"
                " 2 adet nitelikli, beceri temelli çoktan seçmeli soru ve"
                " cevap anahtarı hazarla."
            )
          else:
            video_prompt = (
                "Bu videodaki konuşmaları ve anlatılanları kronolojik sırayla"
                " metin dökümü (transkript) haline getir."
            )

          # Model çağrısı (Multimodal video desteği)
          response = client.models.generate_content(
              model="gemini-2.5-flash",
              contents=[video_file, video_prompt],
              config=types.GenerateContentConfig(
                  temperature=0.3,
                  max_output_tokens=2048,
              ),
          )

          st.success("Video analizi tamamlandı!")
          st.markdown(response.text)

          # Geçici dosyayı temizle
          os.unlink(tmp_file_path)

        except Exception as e:
          st.error(f"Video işlenirken bir hata oluştu: {e}")
