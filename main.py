import sqlite3
import streamlit as st

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# Veritabanı Kurulumu
def init_db():
  conn = sqlite3.connect("okul.db")
  c = conn.cursor()
  c.execute("""
        CREATE TABLE IF NOT EXISTS sorular (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            soru_metni TEXT,
            secenek_a TEXT,
            secenek_b TEXT,
            secenek_c TEXT,
            secenek_d TEXT,
            dogru_cevap TEXT,
            kazanim TEXT
        )
    """)
  c.execute("""
        CREATE TABLE IF NOT EXISTS ogrenciler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            isim TEXT,
            puan INTEGER,
            cozulen_soru INTEGER
        )
    """)
  conn.commit()
  conn.close()


init_db()

# Yan Menü (Navigasyon)
st.sidebar.title("✏️ Kurşun Kalem Dijital")
st.sidebar.markdown("---")
secim = st.sidebar.selectbox(
    "Modül Seçin",
    [
        "👨‍🏫 Yazar / Soru Üretim Paneli",
        "🎯 Öğrenci Soru Çözüm Ekranı",
        "🏆 Liderlik & Karnem",
        "💡 Soru Stratejileri",
    ],
)

# 1. YAZAR / SORU ÜRETİM PANELİ (Senin asıl panelin)
if secim == "👨‍🏫 Yazar / Soru Üretim Paneli":
  st.title("👨‍🏫 Kurşun Kalem - Soru Yazar Paneli")
  st.write(
      "Kendi hazırladığın LGS ve Türkçe sorularını sisteme ekleyebilir, soru"
      " havuzunu yönetebilirsin."
  )

  with st.form("soru_formu"):
    s_metni = st.text_area("Soru Kökü ve Metni")
    col1, col2 = st.columns(2)
    with col1:
      s_a = st.text_input("A Seçeneği")
      s_b = st.text_input("B Seçeneği")
    with col2:
      s_c = st.text_input("C Seçeneği")
      s_d = st.text_input("D Seçeneği")

    dogru = st.selectbox("Doğru Cevap", ["A", "B", "C", "D"])
    kazanim = st.text_input("İlgili Kazanım (Örn: Fiilimsiler, Sözel Mantık)")

    kaydet = st.form_submit_button("Soruyu Sisteme Kaydet")
    if kaydet:
      if s_metni and kazanim:
        conn = sqlite3.connect("okul.db")
        c = conn.cursor()
        c.execute(
            "INSERT INTO sorular (soru_metni, secenek_a, secenek_b, secenek_c,"
            " secenek_d, dogru_cevap, kazanım) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (s_metni, s_a, s_b, s_c, s_d, dogru, kazanim),
        )
        conn.commit()
        conn.close()
        st.success("🎉 Soru başarıyla havuza eklendi!")
      else:
        st.warning("Lütfen soru metnini ve kazanımı boş bırakmayın.")

# 2. ÖĞRENCİ SORU ÇÖZÜM EKRANI
elif secim == "🎯 Öğrenci Soru Çözüm Ekranı":
  st.title("🎯 Öğrenci Test Çözüm Paneli")
  st.write("Sistemdeki havuzdan rastgele gelen soruları çözmeye başla!")

  conn = sqlite3.connect("okul.db")
  c = conn.cursor()
  c.execute("SELECT id, soru_metni, secenek_a, secenek_b, secenek_c, secenek_d, dogru_cevap FROM sorular")
  sorular = c.fetchall()
  conn.close()

  if not sorular:
    st.info(
        "Henüz havuzda soru bulunmuyor. 'Yazar / Soru Üretim Paneli' üzerinden"
        " soru ekleyebilirsin!"
    )
  else:
    for idx, s in enumerate(sorular, 1):
      st.markdown(f"### Soru {idx}")
      st.info(s[1])
      secim_ogrenci = st.radio(
          f"Seçenekler (Soru {idx}):",
          [f"A) {s[2]}", f"B) {s[3]}", f"C) {s[4]}", f"D) {s[5]}"],
          key=f"soru_{s[0]}",
          index=None,
      )

      if st.button(f"Cevabı Kontrol Et (Soru {idx})", key=f"btn_{s[0]}"):
        if secim_ogrenci is None:
          st.warning("Lütfen bir seçenek seç.")
        else:
          verilen_harf = secim_ogrenci[0]
          if verilen_harf == s[6]:
            st.success("🎉 Doğru cevap!")
          else:
            st.error(f"❌ Yanlış. Doğru cevap: **{s[6]}**")

# 3. LİDERLİK & KARNEM
elif secim == "🏆 Liderlik & Karnem":
  st.title("🏆 Liderlik Tablosu ve Gelişim")
  st.write("Öğrencilerin puan durumları ve başarı sıralaması.")

  conn = sqlite3.connect("okul.db")
  c = conn.cursor()
  c.execute(
      "SELECT isim, puan, cozulen_soru FROM ogrenciler ORDER BY puan DESC"
  )
  ogrenciler = c.fetchall()
  conn.close()

  if ogrenciler:
    for sira, (isim, puan, soru) in enumerate(ogrenciler, 1):
      st.write(f"**{sira}.** {isim} - Puan: {puan} (Çözülen: {soru})")
  else:
    st.info("Henüz kayıtlı öğrenci verisi bulunmuyor.")

# 4. SORU STRATEJİLERİ
elif secim == "💡 Soru Stratejileri":
  st.title("💡 LGS Türkçe Çözüm Stratejileri")
  st.write("Sınavda hız kazandıracak taktikler.")
  st.markdown(
      "- **Sözel Mantık:** Önce kesin bilgileri ele alıp tablo oluşturun.\n-"
      " **Paragraf:** Ana düşünce sorularında giriş ve sonuç cümlelerine"
      " odaklanın."
  )
