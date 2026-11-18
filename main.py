import sqlite3
import streamlit as st

# 1. Sayfa Yapılandırması (En başta olmalı)
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# 2. Veritabanı Bağlantısı ve Tablo Oluşturma
def init_db():
  conn = sqlite3.connect("okul.db")
  c = conn.cursor()
  # Öğrenci tablosu
  c.execute("""
        CREATE TABLE IF NOT EXISTS ogrenciler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            isim TEXT,
            puan INTEGER,
            cozulen_soru INTEGER,
            dogru_sayisi INTEGER
        )
    """)
  # Örnek başlangıç verisi yoksa ekle
  c.execute("SELECT COUNT(*) FROM ogrenciler")
  if c.fetchone()[0] == 0:
    ornekler = [
        ("Zeynep Yılmaz", 450, 45, 41),
        ("Ahmet Demir", 380, 40, 34),
        ("Elif Çelik", 520, 55, 50),
        ("Kerem Kaya", 310, 35, 28),
    ]
    c.executemany(
        "INSERT INTO ogrenciler (isim, puan, cozulen_soru, dogru_sayisi) VALUES"
        " (?, ?, ?, ?)",
        ornekler,
    )
  conn.commit()
  conn.close()


init_db()

# 3. Kenar Çubuğu (Sidebar) Navigasyon
st.sidebar.title("📚 Kurşun Kalem")
st.sidebar.markdown("---")
menu = st.sidebar.selectbox(
    "Menü Seçin",
    [
        "🎯 Günlük Test Çöz",
        "📊 Gelişim Karnem",
        "🏆 Liderlik Tablosu",
        "💡 Soru Stratejileri",
    ],
)

# --- SAYFA 1: GÜNLÜK TEST ÇÖZ ---
if menu == "🎯 Günlük Test Çöz":
  st.title("🎯 LGS Türkçe Beceri Testi")
  st.write(
      "Günün sorusunu dikkatle oku, doğru seçeneği işaretle ve anında"
      " değerlendir!"
  )

  st.markdown("---")
  st.markdown("### Soru 1 (Sözel Mantık / Fiilimsiler)")
  st.info(
      "Aşağıdaki cümlelerin hangisinde altı çizili sözcük **sıfat-fiil**"
      " görevinde **kullanılmamıştır**?"
  )

  secim = st.radio(
      "Seçenekler:",
      [
          "A) Akacak kan damarda durmaz.",
          "B) Sararmış yapraklar bahçeyi kaplamıştı.",
          "C) Görünmez kaza hepimizi sarstı.",
          "D) Çocuklar bahçede neşeyle koşuyordu.",
      ],
      index=None,
  )

  if st.button("Cevabı Gönder"):
    if secim is None:
      st.warning("Lütfen bir seçenek işaretleyin!")
    elif secim.startswith("D"):
      st.success(
          "🎉 **Tebrikler, Doğru!** 'Koşuyordu' sözcüğü cümlenin yüklemidir ve"
          " çekimli bir fiildir; sıfat-fiil değildir."
      )
      # Veritabanında skoru artır
      conn = sqlite3.connect("okul.db")
      c = conn.cursor()
      c.execute(
          "UPDATE ogrenciler SET puan = puan + 10, cozulen_soru = cozulen_soru +"
          " 1, dogru_sayisi = dogru_sayisi + 1 WHERE id = 1"
      )
      conn.commit()
      conn.close()
    else:
      st.error(
          "❌ **Yanlış Cevap.** Doğru seçenek **D** olmalıydı. A, B ve C"
          " seçeneklerindeki fiilimsiler isimleri nitelerken, D seçeneğindeki"
          " sözcük çekimli fiildir."
      )

# --- SAYFA 2: GELİŞİM KARNEM ---
elif menu == "📊 Gelişim Karnem":
  st.title("📊 Kişisel Gelişim ve Başarı Karnesi")
  st.write("Genel performansını ve doğru/yanlış analizini buradan incele.")

  conn = sqlite3.connect("okul.db")
  c = conn.cursor()
  c.execute(
      "SELECT puan, cozulen_soru, dogru_sayisi FROM ogrenciler WHERE id = 1"
  )
  veri = c.fetchone()
  conn.close()

  puan, soru, dogru = veri if veri else (0, 0, 0)
  yanlis = soru - dogru if soru >= dogru else 0

  col1, col2, col3 = st.columns(3)
  col1.metric("Toplam Puan", f"{puan} P")
  col2.metric("Çözülen Soru", f"{soru} Soru")
  col3.metric("Doğru / Yanlış", f"{dogru} D / {yanlis} Y")

  st.markdown("---")
  st.subheader("💡 Yapay Zeka Tavsiyesi")
  st.info(
      "Son çözdüğün testlere göre **'Fiilimsiler'** ve **'Sözel Mantık'**"
      " konularında ekstra soru çözmelisin. Tavsiye edilen testler yakında"
      " yüklenecektir."
  )

# --- SAYFA 3: LİDERLİK TABLOSU ---
elif menu == "🏆 Liderlik Tablosu":
  st.title("🏆 Şampiyonlar Ligi - Liderlik Tablosu")
  st.write(
      "En çok puan toplayan ve düzenli soru çözen başarılı öğrencilerimiz!"
  )

  conn = sqlite3.connect("okul.db")
  c = conn.cursor()
  c.execute(
      "SELECT isim, puan, cozulen_soru FROM ogrenciler ORDER BY puan DESC"
  )
  ogrenciler = c.fetchall()
  conn.close()

  # Tablo gösterimi
  tablo_verisi = []
  for sira, (isim, p, s) in enumerate(ogrenciler, 1):
    madalya = "🥇" if sira == 1 else ("🥈" if sira == 2 else ("🥉" if sira == 3 else "🏅"))
    tablo_verisi.append(
        {
            "Sıra": f"{madalya} {sira}",
            "Öğrenci": isim,
            "Puan": p,
            "Çözülen Soru": s,
        }
    )

  st.table(tablo_verisi)

# --- SAYFA 4: SORU STRATEJİLERİ ---
elif menu == "💡 Soru Stratejileri":
  st.title("💡 LGS Türkçe Çözüm Stratejileri")
  st.write(
      "Sınavda hayat kurtaracak pratik ipuçları ve paragraf taktikleri."
  )

  with st.expander("📌 Sözel Mantık Sorularını Çözme Taktikleri"):
    st.write(
        "1. Öncülleri kesin yargılardan başlayarak eleyin.\n"
        "2. Tablo veya şema çizmek her zaman işinizi kolaylaştırır.\n"
        "3. Olumsuz köklere ('değildir', 'olamaz') özellikle dikkat edin."
    )

  with st.expander("📌 Paragraf Sorularında 'Ana Düşünce' Nasıl Bulunur?"):
    st.write(
        "1. Parçanın ilk ve son cümlelerine odaklanın.\n"
        "2. Yazarın okuyucuya vermek istediği temel mesajı tek cümleyle"
        " özetlemeye çalışın.\n"
        "3. Kendi düşüncelerinizi değil, metnin içindeki anlamı esas alın."
    )
