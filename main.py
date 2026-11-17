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
try:
  api_key = st.secrets["API_KEY"]
  client = genai.Client(api_key=api_key)
except Exception as e:
  st.error(
      "API Anahtarı bulunamadı! Lütfen Streamlit Cloud Secrets ayarlarından"
      " 'API_KEY' tanımladığınızdan emin olun."
  )
  st.stop()

# --- ARAYÜZ ---
st.title("✏️ Kurşun Kalem Dijital Eğitim")
st.markdown("LGS standartlarına uygun beceri temelli sorular ve video analiz paneli.")

# Sekmeler (Sade ve Net Yapı)
tab1, tab2 = st.tabs(["📝 Soru Üretim Modülü", "🎥 Ders Videosu İzleme ve Analiz"])

with tab1:
  st.header("Beceri Temelli Soru Üretici")

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
      with st.spinner("Sorular hazırlanıyor, lütfen bekleyin..."):
        try:
          prompt = f"""
                    Aşağıdaki kazanım doğrultusunda {zorluk_derecesi} seviyesinde {soru_sayisi} adet LGS formatında beceri temelli Türkçe sorusu hazarla. Cevap anahtarını da ekle.
                    Kazanım/İçerik: {kazanim_input}
                    """
          response = client.models.generate_content(
              model="gemini-3.8-flash",
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
  st.header("🎥 Ders Videosu İzleme ve Yapay Zekâ Analizi")
  st.markdown(
      "Bilgisayarından bir ders videosu yükleyerek uygulamada canlı olarak"
      " izleyebilir ve yapay zekâya analiz ettirebilirsin."
  )

  uploaded_video = st.file_uploader(
      "Ders Videosu Seç (MP4, MOV, AVI)", type=["mp4", "mov", "avi"]
  )

  if uploaded_video is not None:
    # Yan yana düzen: Sol tarafta video oynatıcı, sağ tarafta analiz paneli
    col_video, col_panel = st.columns([1.2, 1])

    with col_video:
      st.subheader("▶️ Video Oynatıcı")
      st.video(uploaded_video)

    with col_panel:
      st.subheader("⚙️ Yapay Zekâ İşlemleri")
      islem_turu = st.selectbox(
          "Yapılacak İşlem",
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
                model="gemini-3.8-flash",
                contents=[video_file, video_prompt],
                config=types.GenerateContentConfig(
                    temperature=0.3,
                    max_output_tokens=2048,
                ),
            )

            st.success("Analiz tamamlandı!")
            st.markdown(response.text)

            os.unlink(tmp_file_path)

          except Exception as e:
            st.error(f"İşlem sırasında hata oluştu: {e}")
