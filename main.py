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


# --- KALICI ARŞİV (JSON) YÖNETİMİ ---
ARSIV_DOSYASI = "ders_arsivi.json"


def arsivi_yukle():
  if os.path.exists(ARSIV_DOSYASI):
    try:
      with open(ARSIV_DOSYASI, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}


def arsivi_kaydet(arsiv_verisi):
  with open(ARSIV_DOSYASI, "w", encoding="utf-8") as f:
    json.dump(arsiv_verisi, f, ensure_ascii=False, indent=4)


arsiv = arsivi_yukle()


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
    "✏️ Kurşun Kalem Dijital Eğitim - Sınıf ve Konu Arşiv Yönetim Sistemi"
)
st.markdown(
    "Sınıf ve konu seçerek ders notlarınızı, ses/video kayıtlarınızı kalıcı"
    " olarak saklayın ve ilgili konuyu açtığınızda görüntüleyin."
)

# --- SEKMELER ---
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📝 Sınıf & LGS Soru Üretici",
        "📂 Konu Bazlı Dosya Arşivi",
        "📊 Soru Havuzu Analizi",
        "🎥 Konu Bazlı Medya (Ses/Video) Arşivi ve Analiz",
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
  st.header("📂 Konu Bazlı Dosya Arşivi")
  st.markdown(
      "Önce sınıf ve konuyu seçin. Yüklediğiniz dokümanlar yalnızca o seçtiğiniz"
      " konunun altında görünecektir."
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

  with st.expander(
      "➕ Bu Konuya Yeni Doküman / Ders Notu Yükle", expanded=False
  ):
    yuklenen_dosya = st.file_uploader(
        "Dosya Seç (PDF, DOCX, TXT)",
        type=["pdf", "docx", "txt"],
        key="file_uploader_tab2",
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

      if yonetim_sinif not in arsiv:
        arsiv[yonetim_sinif] = {}
      if yonetim_konu not in arsiv[yonetim_sinif]:
        arsiv[yonetim_sinif][yonetim_konu] = {"dokumanlar": [], "medyalar": []}

      mevcut_dosyalar = [
          d["ad"] for d in arsiv[yonetim_sinif][yonetim_konu]["dokumanlar"]
      ]
      if yuklenen_dosya.name not in mevcut_dosyalar:
        arsiv[yonetim_sinif][yonetim_konu]["dokumanlar"].append(
            {"ad": yuklenen_dosya.name, "icerik": metin_icerik}
        )
        arsivi_kaydet(arsiv)
        st.success(
            f"'{yuklenen_dosya.name}' başarıyla {yonetim_sinif} ->"
            f" '{yonetim_konu}' arşivine eklendi!"
        )

  st.divider()
  st.subheader(
      f"📚 {yonetim_sinif} / {yonetim_konu} Arşivindeki Dokümanlar"
  )

  konu_dokumanlari = (
      arsiv.get(yonetim_sinif, {}).get(yonetim_konu, {}).get("dokumanlar", [])
  )
  if konu_dokumanlari:
    for idx, doc in enumerate(konu_dokumanlari):
      with st.expander(f"📄 {doc['ad']}", expanded=False):
        st.text_area(
            "İçerik Önizlemesi:",
            value=doc["icerik"][:3000],
            height=150,
            key=f"doc_preview_{idx}",
        )
        if st.button(
            "Bu Dokümanı Yapay Zekâ ile Analiz Et", key=f"analiz_doc_{idx}"
        ):
          with st.spinner("Analiz ediliyor..."):
            p = f"Bu dokümanı {yonetim_sinif} seviyesindeki '{yonetim_konu}' başlığı altında incele, temel kazanımları özetle."
            resp = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=[p, doc["icerik"][:3000]],
                config=types.GenerateContentConfig(
                    temperature=0.3, max_output_tokens=1500
                ),
            )
            st.markdown(resp.text)
  else:
    st.info(
        f"'{yonetim_konu}' başlığı altında henüz kayıtlı bir doküman"
        " bulunmuyor."
    )

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
  st.header("🎥 Konu Bazlı Medya (Ses / Video) Arşivi ve Analiz")
  st.markdown(
      "Önce sınıf ve konuyu seçin. Yüklediğiniz medya dosyaları yalnızca o"
      " seçtiğiniz konunun altında görünecektir."
  )

  col_v1, col_v2 = st.columns(2)
  with col_v1:
    video_sinif = st.selectbox(
        "Sınıf Seçin:", list(SINIF_KONTENJLARI.keys()), key="video_sinif"
    )
  with col_v2:
    video_konu = st.selectbox(
        "Konu Seçin:", SINIF_KONTENJLARI[video_sinif], key="video_konu"
    )

  with st.expander(
      "➕ Bu Konuya Yeni Medya Dosyası (Video / Ses) Yükle", expanded=False
  ):
    uploaded_media = st.file_uploader(
        "Medya Dosyası Seç (MP4, MOV, AVI, MP3, WAV, M4A)",
        type=["mp4", "mov", "avi", "mp3", "wav", "m4a"],
        key="media_uploader_tab4",
    )

    if uploaded_media is not None:
      dosya_uzanti = uploaded_media.name.split(".")[-1].lower()

      with tempfile.NamedTemporaryFile(
          delete=False, suffix=f".{dosya_uzanti}"
      ) as tmp_file:
        uploaded_media.seek(0)
        tmp_file.write(uploaded_media.read())
        tmp_file_path = tmp_file.name

      try:
        media_file = client.files.upload(file=tmp_file_path)

        if video_sinif not in arsiv:
          arsiv[video_sinif] = {}
        if video_konu not in arsiv[video_sinif]:
          arsiv[video_sinif][video_konu] = {"dokumanlar": [], "medyalar": []}

        uploaded_media.seek(0)
        media_bytes = uploaded_media.read()

        arsiv[video_sinif][video_konu]["medyalar"].append(
            {
                "ad": uploaded_media.name,
                "uzanti": dosya_uzanti,
                "veri": list(media_bytes),
                "gemini_file_name": media_file.name,
            }
        )
        arsivi_kaydet(arsiv)
        st.success(
            f"'{uploaded_media.name}' başarıyla {video_sinif} -> '{video_konu}'"
            " arşivine eklendi!"
        )
      except Exception as e:
        st.error(f"Medya yüklenirken hata oluştu: {e}")
      finally:
        if os.path.exists(tmp_file_path):
          os.unlink(tmp_file_path)

  st.divider()
  st.subheader(f"🎬 {video_sinif} / {video_konu} Arşivindeki Medyalar")

  konu_medyalari = (
      arsiv.get(video_sinif, {}).get(video_konu, {}).get("medyalar", [])
  )
  if konu_medyalari:
    for idx, medya in enumerate(konu_medyalari):
      with st.expander(f"🎵 / 🎥 {medya['ad']}", expanded=True):
        col_m1, col_m2 = st.columns([1.2, 1])

        with col_m1:
          st.subheader("▶️ Oynatıcı")
          media_data_bytes = bytes(medya["veri"])
          if medya["uzanti"] in ["mp3", "wav", "m4a"]:
            st.audio(media_data_bytes)
          else:
            st.video(media_data_bytes)

        with col_m2:
          st.subheader("⚙️ Yapay Zekâ Analiz İşlemleri")
          islem_turu = st.selectbox(
              "İşlem Seçin:",
              [
                  "Ders Özeti ve Ana Başlıklar",
                  "İçerikten Beceri Temelli Soru Üret",
                  "Konuşma Dökümü (Transkript)",
              ],
              key=f"islem_{idx}",
          )

          if st.button(
              "Arşivdeki Medyayı Analiz Et", key=f"btn_analiz_{idx}"
          ):
            with st.spinner("Yapay zekâ arşivdeki dosyayı inceliyor..."):
              try:
                gemini_f = client.files.get(name=medya["gemini_file_name"])

                if islem_turu == "Ders Özeti ve Ana Başlıklar":
                  media_prompt = f"Bu ders kaydını/videosunu {video_sinif} seviyesi '{video_konu}' konusu bağlamında analiz et. Öğrenciler için maddeler halinde detaylı bir ders özeti çıkar."
                elif islem_turu == "İçerikten Beceri Temelli Soru Üret":
                  media_prompt = f"Bu medya içeriğindeki konuya dayanarak {video_sinif} - '{video_konu}' için uygun formatta beceri temelli sorular ve cevap anahtarı hazarla."
                else:
                  media_prompt = "Bu ses/video kaydındaki konuşmaları kronolojik sırayla metin dökümü (transkript) haline getir."

                response = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=[gemini_f, media_prompt],
                    config=types.GenerateContentConfig(
                        temperature=0.3, max_output_tokens=2048
                    ),
                )

                st.success("Analiz tamamlandı!")
                st.markdown(response.text)
              except Exception as e:
                st.error(
                    f"Analiz sırasında hata oluştu (Dosya süresi dolmuş olabilir,"
                    f" yeniden yükleyebilirsiniz): {e}"
                )
  else:
    st.info(
        f"'{video_konu}' başlığı altında henüz kayıtlı bir ses/video medya"
        " dosyası bulunmuyor."
    )
