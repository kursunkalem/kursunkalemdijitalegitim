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
    " sorular üretin ve ders video/ses dosyalarınızı analiz edin."
)

# --- SEKMELER ---
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📝 Sınıf & LGS Soru Üretici",
        "📂 Sınıf & Konu Dosya Yönetimi",
        "📊 Soru Havuzu Analizi",
        "🎥 Sınıf Bazlı Medya (Video/Ses) Analiz Paneli",
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
                  - MEB kazanımlarına ve LGS/okul sınavı mantığına tam uygun, öğrencinin okuduğunu anlama ve yorumlama becerilerini ölçen nitelikli öncüller kullan.
                  - Cevap anahtarını ve detaylı açıklamasını mutlaka ekle.
                  """
        response = client.models.generate_content(
            model="gemini-3.8-flash",
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
            file_name=f"kursunkalem_{secilen_sinif_s.split('.')[0]}_sinif_sorulari.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
        )
      except Exception as e:
        st.error(f"Soru üretilirken hata oluştu: {e}")

with tab2:
  st.header("📂 Sınıf ve Konu Bazlı Dosya Yönetimi")
  st.markdown(
      "Ders notlarınızı veya test dokümanlarınızı sınıf ve ünite konularına göre"
      " sisteme yükleyin."
  )

  col_d1, col_d2 = st.columns(2)
  with col_d1:
    yonetim_sinif = st.selectbox(
        "Sınıf Seçin:", list(SINIF_KONTENJLARI.keys()), key="yonetim_sinif"
    )
  with col_d2:
    yonetim_konu = st.selectbox(
        "Konu Seçin:", SINIF_KONTENJLARI[yonetim_sinif], key="yonetim_konu"
    )

  yuklenen_dosya = st.file_uploader(
      f"'{yonetim_konu}' için Ders Notu / Doküman Yükle (PDF, DOCX, TXT)",
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

    st.success(
        f"'{yuklenen_dosya.name}' dosyası başarıyla okunarak {yonetim_sinif} -"
        f" '{yonetim_konu}' başlığına bağlandı!"
    )
    st.text_area(
        "Belge İçeriği Önizlemesi:", value=metin_icerik[:3000], height=200
    )

    if st.button("Belgeyi Yapay Zekâ ile Analiz Et"):
      with st.spinner("Analiz ediliyor..."):
        p = f"Bu dokümanı {yonetim_sinif} seviyesindeki '{yonetim_konu}' başlığı altında incele, temel kazanımları ve önemli bilgileri özetle."
        resp = client.models.generate_content(
            model="gemini-3.8-flash",
            contents=[p, metin_icerik[:3000]],
            config=types.GenerateContentConfig(
                temperature=0.3, max_output_tokens=1500
            ),
        )
        st.markdown(resp.text)

with tab3:
  st.header("📊 Soru Havuzu Analiz Paneli")
  if referanslar_data:
    st.success(f"Sistemde toplam **{len(referanslar_data)}** referans soru aktif.")
    for i, soru in enumerate(referanslar_data):
      with st.expander(
          f"📌 Soru {i+1} | Branş: {soru.get('brans', 'Belirtilmemiş')} | Sınıf:"
          f" {soru.get('sinif', 'Belirtilmemiş')}"
      ):
        st.markdown(f"**Kazanım:** {soru.get('kazanim', 'Belirtilmemiş')}")
        st.markdown("---")
        st.markdown("**Soru Metni / Bağlam:**")
        st.info(soru.get("baglam_metni", "Metin bulunamadı."))
  else:
    st.warning("Henüz havuzda referans soru bulunmuyor.")

with tab4:
  st.header("🎥 Sınıf ve Konu Bazlı Medya (Video / Ses) Analiz Paneli")
  st.markdown(
      "Bilgisayarınızdaki ders videosunu veya ses kaydını (MP4, MOV, AVI, MP3,"
      " WAV, M4A) yükleyin; sınıf ve konuyu seçerek yapay zekâya analiz"
      " ettirin."
  )

  col_v1, col_v2 = st.columns(2)
  with col_v1:
    video_sinif = st.selectbox(
        "Medya Dosyasının Ait Olduğu Sınıf:",
        list(SINIF_KONTENJLARI.keys()),
        key="video_sinif",
    )
  with col_v2:
    video_konu = st.selectbox(
        "Medya Dosyasının Ait Olduğu Konu:",
        SINIF_KONTENJLARI[video_sinif],
        key="video_konu",
    )

  uploaded_media = st.file_uploader(
      "Medya Dosyası Seçin (Video veya MP3/Ses Dosyası)",
      type=["mp4", "mov", "avi", "mp3", "wav", "m4a"],
  )

  if uploaded_media is not None:
    dosya_uzanti = uploaded_media.name.split(".")[-1].lower()
    col_media, col_panel = st.columns([1.2, 1])

    with col_media:
      st.subheader(f"▶️ Oynatıcı ({video_sinif})")
      if dosya_uzanti in ["mp3", "wav", "m4a"]:
        st.audio(uploaded_media)
      else:
        st.video(uploaded_media)

    with col_panel:
      st.subheader("⚙️ Yapay Zekâ Analiz İşlemleri")
      islem_turu = st.selectbox(
          "İşlem Seçin:",
          [
              "Ders Özeti ve Ana Başlıklar",
              "İçerikten Beceri Temelli Soru Üret",
              "Konuşma Dökümü (Transkript)",
          ],
      )

      if st.button("Medya Analizini Başlat", type="primary"):
        with st.spinner("Dosya işleniyor ve analiz ediliyor..."):
          try:
            with tempfile.NamedTemporaryFile(
                delete=False, suffix=f".{dosya_uzanti}"
            ) as tmp_file:
              uploaded_media.seek(0)
              tmp_file.write(uploaded_media.read())
              tmp_file_path = tmp_file.name

            media_file = client.files.upload(file=tmp_file_path)

            if islem_turu == "Ders Özeti ve Ana Başlıklar":
              media_prompt = f"Bu ders kaydını/videosunu {video_sinif} seviyesi '{video_konu}' konusu bağlamında analiz et. Öğrenciler için maddeler halinde detaylı bir ders özeti çıkar."
            elif islem_turu == "İçerikten Beceri Temelli Soru Üret":
              media_prompt = f"Bu medya içeriğindeki konuya dayanarak {video_sinif} - '{video_konu}' için uygun formatta beceri temelli sorular ve cevap anahtarı hazarla."
            else:
              media_prompt = "Bu ses/video kaydındaki konuşmaları kronolojik sırayla metin dökümü (transkript) haline getir."

            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=[media_file, media_prompt],
                config=types.GenerateContentConfig(
                    temperature=0.3, max_output_tokens=2048
                ),
            )

            st.success("Analiz tamamlandı!")
            st.markdown(response.text)

            os.unlink(tmp_file_path)

          except Exception as e:
            st.error(f"Hata oluştu: {e}")
