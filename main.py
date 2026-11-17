import json
import os
import streamlit as st
from google import genai

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim", page_icon="✏️", layout="wide"
)

# --- KLASÖR VE DOSYA YOLLARI ---
DATA_DIR = "veri"
MEDIA_DIR = "media_arsivi"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(MEDIA_DIR, exist_ok=True)

SINIF_KONULARI_DOSYA = os.path.join(DATA_DIR, "sinif_konulari.json")
DERS_ARSIVI_DOSYA = os.path.join(DATA_DIR, "ders_arsivi.json")


# --- VARSAYILAN VERİLER ---
def varsayilan_konulari_yukle():
  return {
      "5. Sınıf Türkçe": [
          "Sözcükte Anlam (Gerçek, Mecaz, Terim)",
          "Sözcükler Arası Anlam İlişkileri",
          "Deyimler ve Atasözleri",
          "Cümlede Anlam",
          "Metnin Anlamı ve Yorumu",
      ],
      "6. Sınıf Türkçe": [
          "Sözcükte ve Söz Öbeklerinde Anlam",
          "Cümlede Anlam İlişkileri",
          "Parçada Anlam ve Yapı",
          "Metin Türleri",
          "Yazım Kuralları ve Noktalama",
      ],
      "7. Sınıf Türkçe": [
          "Sözcükte Anlam ve Çok Boyutlu Yorum",
          "Cümlede Düşünceyi Geliştirme Yolları",
          "Fiilimsiler",
          "Anlatım Bozuklukları",
          "Grafik, Tablo ve Görsel Okuma",
      ],
      "8. Sınıf Türkçe (LGS)": [
          "Sözel Mantık ve Muhakeme",
          "Görsel, Grafik ve Tablo Yorumlama",
          "Fiil Çatısı ve Cümle Ögeleri",
          "Metinler Arası Karşılaştırma",
          "Yazım Kuralları ve Noktalama İşaretleri (LGS Odaklı)",
      ],
  }


# --- JSON YÖNETİM FONKSİYONLARI ---
def veri_yukle(dosya_yolu, varsayilan_fonk):
  if os.path.exists(dosya_yolu):
    try:
      with open(dosya_yolu, "r", encoding="utf-8") as f:
        return json.load(f)
    except:
      return varsayilan_fonk()
  else:
    veri = varsayilan_fonk()
    veri_kaydet(dosya_yolu, veri)
    return veri


def veri_kaydet(dosya_yolu, veri):
  with open(dosya_yolu, "w", encoding="utf-8") as f:
    json.dump(veri, f, ensure_ascii=False, indent=4)


# Verileri başlat
sinif_konulari = veri_yukle(SINIF_KONULARI_DOSYA, varsayilan_konulari_yukle)
ders_arsivi = veri_yukle(
    DERS_ARSIVI_DOSYA, lambda: {"dokumanlar": [], "medyalar": []}
)

# --- BAŞLIK VE AÇIKLAMA ---
st.markdown(
    "# ✏️ Kurşun Kalem Dijital Eğitim - Sınıf ve Konu Arşiv Yönetim Sistemi"
)
st.write(
    "Sınıf/konu yönetimi yapın, soru üretin ve doküman/medyalarınızı konuya"
    " özel saklayın."
)

# --- SEKME YAPISI (Fiyatlandırma Sekmesi Eklendi) ---
sekme1, sekme2, sekme3, sekme4, sekme5, sekme6 = st.tabs([
    "⚙️ Sınıf & Konu Yönetimi",
    "📝 Sınıf & LGS Soru Üretici",
    "📁 Konu Bazlı Dosya Arşivi",
    "📊 Soru Havuzu Analizi",
    "🎬 Konu Bazlı Medya Arşivi",
    "💎 Abonelik & Fiyatlandırma",
])

# ==========================================
# 1. SEKME: SINIF & KONU YÖNETİMİ
# ==========================================
with sekme1:
  st.subheader("📚 Sınıf Seviyeleri ve Konu / Kazanım Yönetimi")
  st.write(
      "Mevcut sınıflara yeni konular ekleyebilir veya yeni sınıf seviyesi"
      " tanımlayabilirsiniz."
  )

  col1, col2 = st.columns(2)

  with col1:
    st.markdown("### Yeni Sınıf Ekle")
    yeni_sinif_adi = st.text_input(
        "Sınıf Adı (Örn: hazırlık veya 4. Sınıf Türkçe)"
    )
    if st.button("Sınıfı Ekle"):
      if yeni_sinif_adi and yeni_sinif_adi not in sinif_konulari:
        sinif_konulari[yeni_sinif_adi] = []
        veri_kaydet(SINIF_KONULARI_DOSYA, sinif_konulari)
        st.success(f"'{yeni_sinif_adi}' başarıyla eklendi!")
        st.rerun()
      elif yeni_sinif_adi in sinif_konulari:
        st.warning("Bu sınıf zaten mevcut.")

  with col2:
    st.markdown("### Mevcut Sınıfa Konu Ekle")
    secilen_sinif_ekle = st.selectbox(
        "Sınıf Seçin", list(sinif_konulari.keys()), key="sinif_sec_ekle"
    )
    yeni_konu_adi = st.text_input("Konu / Kazanım Adı")
    if st.button("Konuyu Ekle"):
      if yeni_konu_adi:
        if yeni_konu_adi not in sinif_konulari[secilen_sinif_ekle]:
          sinif_konulari[secilen_sinif_ekle].append(yeni_konu_adi)
          veri_kaydet(SINIF_KONULARI_DOSYA, sinif_konulari)
          st.success(f"'{yeni_konu_adi}' konuya eklendi!")
          st.rerun()
        else:
          st.warning("Bu konu bu sınıfta zaten var.")

  st.divider()
  st.markdown("### 📋 Mevcut Sınıf ve Konu Listesi")
  for sinif, konular in sinif_konulari.items():
    with st.expander(f"📌 {sinif} ({len(konular)} Konu)"):
      if konular:
        for idx, konu in enumerate(konular):
          col_a, col_b = st.columns([0.85, 0.15])
          col_a.write(f"- {konu}")
          if col_b.button(
              "Sil", key=f"sil_{sinif}_{idx}", type="secondary"
          ):
            sinif_konulari[sinif].pop(idx)
            veri_kaydet(SINIF_KONULARI_DOSYA, sinif_konulari)
            st.rerun()
      else:
        st.info("Bu sınıfa ait henüz konu eklenmemiş.")

# ==========================================
# 2. SEKME: SORU ÜRETİCİ (GEMINI API)
# ==========================================
with sekme2:
  st.subheader("Beceri Temelli Sınıf & LGS Soru Üretici")

  api_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

  col_s1, col_s2 = st.columns(2)
  with col_s1:
    hedef_sinif = st.selectbox(
        "Hedef Sınıf Seviyesi:",
        list(sinif_konulari.keys()),
        key="soru_hedef_sinif",
    )
  with col_s2:
    konu_listesi = (
        sinif_konulari.get(hedef_sinif, ["Genel"])
        if sinif_konulari.get(hedef_sinif)
        else ["Genel"]
    )
    ilgili_konu = st.selectbox(
        "İlgili Konu / Kazanım:", konu_listesi, key="soru_ilgili_konu"
    )

  ozel_istem = st.text_area(
      "Özel İstem / Ek Açıklama (İsteğe Bağlı):",
      placeholder="Örn: Öncüllü, grafik yorumlamaya dayalı, çeldiricisi güçlü...",
  )

  col_z1, col_z2, col_z3 = st.columns(3)
  with col_z1:
    zorluk = st.selectbox(
        "Zorluk Seviyesi",
        ["Beceri Temelli (Yeni Nesil)", "Orta / MEB Tarzı", "Güçlük / Seçici"],
    )
  with col_z2:
    soru_sayisi = st.number_input(
        "Üretilecek Soru Sayısı", min_value=1, max_value=5, value=1
    )
  with col_z3:
    soru_tipi = st.selectbox(
        "Soru Tipi", ["Çoktan Seçmeli (4 Şıkkı)", "Açık Uçlu / Cevap Anahtarlı"]
    )

  if st.button("Soruları Üret", type="primary"):
    if not api_key:
      st.error(
          "⚠️ GEMINI_API_KEY bulunamadı! Lütfen Streamlit secrets veya çevre"
          " değişkenlerine ekleyin."
      )
    else:
      try:
        client = genai.Client(api_key=api_key)
        prompt = (
            f"Sen uzman bir Türkçe eğitimi ve LGS soru yazarısın. Aşağıdaki"
            f" kriterlere uygun {soru_sayisi} adet {zorluk} seviyesinde"
            f" {soru_tipi} üret.\n"
            f"- Sınıf Seviyesi: {hedef_sinif}\n"
            f"- Konu / Kazanım: {ilgili_konu}\n"
            f"- Ek Açıklama: {ozel_istem}\n\n"
            "Lütfen soruları MEB beceri temelli LGS formatına uygun olarak"
            " öncüllü, metinli ve net çeldiricilerle hazırla. Cevap anahtarını"
            " ve çözüm açıklamasını da ekle."
        )

        with st.spinner("Yapay zekâ soruları hazırlıyor..."):
          response = client.models.generate_content(
              model="gemini-2.5-flash",
              contents=prompt,
          )
          st.markdown("### 📄 Üretilen Soru Çıktısı")
          st.markdown(response.text)
      except Exception as e:
        st.error(f"Soru üretilirken hata oluştu: {e}")

# ==========================================
# 3. SEKME: KONU BAZLI DOSYA ARŞİVİ
# ==========================================
with sekme3:
  st.subheader("📁 Konu Bazlı Doküman Arşivi (PDF, DOCX, TXT)")

  with st.form("dosya_yukleme_formu", clear_on_submit=True):
    f_sinif = st.selectbox(
        "Sınıf Seçin", list(sinif_konulari.keys()), key="dosya_sinif"
    )
    f_konu_list = sinif_konulari.get(f_sinif, ["Genel"])
    f_konu = st.selectbox("Konu Seçin", f_konu_list, key="dosya_konu")
    yuklenen_dosya = st.file_uploader(
        "Doküman Yükle (PDF, DOCX, TXT)", type=["pdf", "docx", "txt"]
    )
    dosya_kaydet_btn = st.form_submit_button("Arşive Kaydet")

    if dosya_kaydet_btn and yuklenen_dosya:
      dosya_yolu = os.path.join(MEDIA_DIR, yuklenen_dosya.name)
      with open(dosya_yolu, "wb") as f:
        f.write(yuklenen_dosya.getbuffer())

      ders_arsivi["dokumanlar"].append({
          "sinif": f_sinif,
          "konu": f_konu,
          "dosya_adi": yuklenen_dosya.name,
          "yol": dosya_yolu,
      })
      veri_kaydet(DERS_ARSIVI_DOSYA, ders_arsivi)
      st.success(f"'{yuklenen_dosya.name}' başarıyla arşive eklendi!")

  st.divider()
  st.markdown("### 🗂️ Arşivdeki Dokümanlar")
  if ders_arsivi["dokumanlar"]:
    for idx, doc in enumerate(ders_arsivi["dokumanlar"]):
      col1, col2, col3 = st.columns([2, 2, 1])
      col1.write(f"**{doc['sinif']}** / {doc['konu']}")
      col2.write(f"📄 {doc['dosya_adi']}")
      if col3.button("Sil", key=f"doc_sil_{idx}"):
        ders_arsivi["dokumanlar"].pop(idx)
        veri_kaydet(DERS_ARSIVI_DOSYA, ders_arsivi)
        st.rerun()
  else:
    st.info("Henüz yüklenmiş bir doküman bulunmuyor.")

# ==========================================
# 4. SEKME: SORU HAVUZU ANALİZİ
# ==========================================
with sekme4:
  st.subheader("📊 Soru Havuzu Analizi ve İstatistikler")
  toplam_sinif = len(sinif_konulari)
  toplam_konu = sum(len(k) for k in sinif_konulari.values())
  toplam_dokuman = len(ders_arsivi["dokumanlar"])
  toplam_medya = len(ders_arsivi["medyalar"])

  col1, col2, col3, col4 = st.columns(4)
  col1.metric("Toplam Sınıf Seviyesi", toplam_sinif)
  col2.metric("Toplam Konu / Kazanım", toplam_konu)
  col3.metric("Arşivlenen Doküman", toplam_dokuman)
  col4.metric("Arşivlenen Medya", toplam_medya)

  st.info("Sistemdeki toplam veri ve içerik dağılımını buradan takip edebilirsiniz.")

# ==========================================
# 5. SEKME: KONU BAZLI MEDYA ARŞİVİ
# ==========================================
with sekme5:
  st.subheader("🎬 Konu Bazlı Medya Arşivi (Video / Ses ve Analiz)")

  with st.form("medya_yukleme_formu", clear_on_submit=True):
    m_sinif = st.selectbox(
        "Sınıf Seçin", list(sinif_konulari.keys()), key="medya_sinif"
    )
    m_konu_list = sinif_konulari.get(m_sinif, ["Genel"])
    m_konu = st.selectbox("Konu Seçin", m_konu_list, key="medya_konu")
    yuklenen_medya = st.file_uploader(
        "Medya Yükle (MP4, MOV, AVI, MP3, WAV, M4A)",
        type=["mp4", "mov", "avi", "mp3", "wav", "m4a"],
    )
    medya_kaydet_btn = st.form_submit_button("Medya Arşivine Kaydet")

    if medya_kaydet_btn and yuklenen_medya:
      medya_yolu = os.path.join(MEDIA_DIR, yuklenen_medya.name)
      with open(medya_yolu, "wb") as f:
        f.write(yuklenen_medya.getbuffer())

      ders_arsivi["medyalar"].append({
          "sinif": m_sinif,
          "konu": m_konu,
          "dosya_adi": yuklenen_medya.name,
          "yol": medya_yolu,
      })
      veri_kaydet(DERS_ARSIVI_DOSYA, ders_arsivi)
      st.success(f"'{yuklenen_medya.name}' medya arşive eklendi!")

  st.divider()
  st.markdown("### 🎦 Arşivdeki Medya Dosyaları")
  if ders_arsivi["medyalar"]:
    for idx, med in enumerate(ders_arsivi["medyalar"]):
      col1, col2, col3 = st.columns([2, 2, 1])
      col1.write(f"**{med['sinif']}** / {med['konu']}")
      col2.write(f"🎥 {med['dosya_adi']}")
      if col3.button("Sil", key=f"med_sil_{idx}"):
        ders_arsivi["medyalar"].pop(idx)
        veri_kaydet(DERS_ARSIVI_DOSYA, ders_arsivi)
        st.rerun()
  else:
    st.info("Henüz yüklenmiş bir medya dosyası bulunmuyor.")

# ==========================================
# 6. SEKME: ABONELİK & FİYATLANDIRMA (YENİ)
# ==========================================
with sekme6:
  st.subheader("💎 Kurşun Kalem Dijital - Pro Paketler ve Abonelik")
  st.write(
      "Öğretmenler, özel ders verenler ve eğitim kurumları için hazırlanan"
      " avantajlı paketlerimizle tanışın."
  )

  col_p1, col_p2 = st.columns(2)

  with col_p1:
    st.markdown("### 📌 Bireysel Standart Paket")
    st.markdown("### **199 TL / Ay**")
    st.markdown(
        "- Ayda **300 Adet** Beceri Temelli Soru Üretimi\n- 5., 6., 7. ve 8."
        " Sınıf Müfredat Tam Erişim\n- PDF Çıktı Desteği\n- Standart AI Yanıt"
        " Hızı"
    )
    # Stripe ödeme linkini buraya ekleyebilirsin
    st.link_button(
        "Standart Paketi Seç",
        "https://buy.stripe.com/buraya_standart_linkin_gelecek",
    )

  with col_p2:
    st.markdown("### 🚀 Pro / Zümre Paketi")
    st.markdown("### **599 TL / Ay**")
    st.markdown(
        "- **Sınırsız Soru Üretimi**\n- Sınırsız Konu ve Kazanım Ekleme\n-"
        " Gelişmiş Doküman ve Medya Arşivi\n- Öncelikli ve Hızlı AI Yanıt Süresi"
    )
    # Stripe pro ödeme linkini buraya ekleyebilirsin
    st.link_button(
        "Pro Paketi Seç", "https://buy.stripe.com/buraya_pro_linkin_gelecek"
    )

  st.divider()
  st.info(
      "💡 **Not:** Satın alma işlemi sonrasında üyeliğinizin anında aktif"
      " edilmesi için ödeme dekontunuzu veya kayıt e-postanızı sistem"
      " yöneticisine iletebilirsiniz."
  )
