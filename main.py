import streamlit as st
import sqlite3
import random
import os
import datetime
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- SAYFA YAPILANDIRMASI ---
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim",
    page_icon="✏️",
    layout="wide"
)

# --- VERİTABANI BAĞLANTISI VE KURULUMU ---
DB_NAME = "kursun_kalem.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Kullanıcılar Tablosu
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            role TEXT,
            points INTEGER DEFAULT 0,
            package TEXT DEFAULT 'Ücretsiz (Freemium)'
        )
    """)
    # Soru Havuzu Tablosu
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject TEXT,
            grade TEXT,
            question_text TEXT,
            option_a TEXT,
            option_b TEXT,
            option_c TEXT,
            option_d TEXT,
            correct_answer TEXT,
            explanation TEXT
        )
    """)
    conn.commit()
    conn.close()

init_db()

# --- ÖRNEK VERİ EKLEME (İLK ÇALIŞTIRMA İÇİN) ---
def seed_initial_data():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM questions")
    if cursor.fetchone()[0] == 0:
        sample_questions = [
            (
                "LGS Türkçe", "8. Sınıf",
                "Aşağıdaki cümlelerin hangisinde bir yazım yanlışı vardır?",
                "A) Herkes bu konuda aynı fikirdeydi.",
                "B) Yanlış anlamalara mahal vermemek için açık konuştu.",
                "C) Türkçeyi korumak hepimizin görevidir.",
                "D) Herhangi birşey olduğunda beni ararsın.",
                "D",
                "Çeldiriciye dikkat! 'Herhangi bir' ayrı yazılır ancak 'şey' kelimesi daima ayrı yazılır. Doğru kullanım 'herhangi bir şey' şeklinde olmalıdır."
            ),
            (
                "LGS Türkçe", "8. Sınıf",
                "Bir sinema salonunda A, B, C, D koltuklarına oturan 4 arkadaşın oturma düzeniyle ilgili bilinenler şunlardır: \n1. Cansu en kenardaki koltukta oturmaktadır.\n2. Ahmet, Burak'ın hemen solundadır.\nBuna göre kim kesinlikle C koltuğundadır?",
                "A) Ahmet", "B) Burak", "C) Cansu", "D) Deniz",
                "C",
                "Sözel mantık kuralı gereği kesin olan öncüller yerleştirildiğinde Cansu'nun yeri netleşir."
            )
        ]
        cursor.executemany("""
            INSERT INTO questions (subject, grade, question_text, option_a, option_b, option_c, option_d, correct_answer, explanation)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, sample_questions)
        conn.commit()
    conn.close()

seed_initial_data()

# --- YARDIMCI FONKSİYONLAR: 2 SÜTUNLU WORD (.DOCX) OLUŞTURUCU ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def create_double_column_docx(questions_list, test_title="LGS TÜRKÇE DENEME SINAVI"):
    doc = docx.Document()
    
    # Kenar Boşlukları (Daraltılmış - 1.27 cm)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
    # Başlık
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"--- {test_title} ---")
    run_title.bold = True
    run_title.font.size = Pt(14)
    run_title.font.color.rgb = RGBColor(26, 82, 118)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Kurşun Kalem Dijital Eğitim - Soru Fabrikası Çıktısı\nAd Soyad: ...................................  No: ...........")
    run_sub.font.size = Pt(9)
    run_sub.font.italic = True
    
    doc.add_paragraph() # Boşluk
    
    # 2 Sütunlu Yapı için Tablo Oluşturma
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Inches(3.5)
    table.columns[1].width = Inches(3.5)
    
    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)
    
    set_cell_background(cell_left, "F9F9F9")
    set_cell_background(cell_right, "F9F9F9")
    
    p_left = cell_left.paragraphs[0]
    p_right = cell_right.paragraphs[0]
    
    current_target_p = p_left
    
    for idx, q in enumerate(questions_list, 1):
        if idx > 1:
            current_target_p = current_target_p.insert_paragraph_after()
            
        run_q = current_target_p.add_run(f"Soru {idx} [{q[1]} - {q[2]}]:\n{q[3]}\n")
        run_q.bold = True
        run_q.font.size = Pt(9.5)
        
        for opt in [q[4], q[5], q[6], q[7]]:
            run_opt = current_target_p.add_run(f"{opt}\n")
            run_opt.font.size = Pt(9)
            
        current_target_p.add_run("\n")
        
        if idx == len(questions_list) // 2:
            current_target_p = p_right

    file_path = "kursun_kalem_2sutun_test.docx"
    doc.save(file_path)
    return file_path

# --- KENAR ÇUBUĞU (NAVİGASYON VE SAYAÇ) ---
st.sidebar.title("✏️ Kurşun Kalem Eğitim")
st.sidebar.caption("SaaS Ekosistem & Soru Fabrikası")

menu = st.sidebar.selectbox(
    "Modül Seçin",
    [
        "🏠 Ana Sayfa & Geri Sayım",
        "👨‍🎓 Öğrenci Portalı & Soru Çözümü",
        "👪 Veli Takip & Performans",
        "🏢 B2B Lisans & Kurum Paneli",
        "🤖 Yapay Zeka Soru Fabrikası",
        "📄 2 Sütunlu Word Test Üretici",
        "💎 Paketler & Abonelik (Freemium)"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("🚀 **Sistem Durumu:** Aktif\n\n📌 **Aktif Paket:** Kurumsal & Bireysel Hibrit\n\n🐱 **Maskotlar:** Gri Cafer & Beyaz Cevdet onayladı!")

# --- 1. ANA SAYFA & GERİ SAYIM ---
if menu == "🏠 Ana Sayfa & Geri Sayım":
    st.title("📚 Kurşun Kalem Dijital Eğitim Platformuna Hoş Geldiniz")
    st.markdown("İlkokuldan YKS'ye uzanan, yapay zeka destekli Türkiye'nin en akıllı soru ve koçluk ekosistemi.")
    
    # --- LGS 2027 GERİ SAYIM SAYACI ---
    st.markdown("### ⏳ 2027 LGS Sınavına Geri Sayım (13 Haziran 2027)")
    exam_date = datetime.datetime(2027, 6, 13, 9, 30, 0)
    now = datetime.datetime.now()
    remaining = exam_date - now
    
    if remaining.total_seconds() > 0:
        days = remaining.days
        hours, remainder = divmod(remaining.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(label="📅 Gün", value=f"{days} Gün")
        c2.metric(label="⏰ Saat", value=f"{hours} Saat")
        c3.metric(label="⏱️ Dakika", value=f"{minutes} Dk")
        c4.metric(label="⚡ Saniye", value=f"{seconds} Sn")
    else:
        st.success("🎉 LGS 2027 Sınavı Gerçekleşti!")
    
    st.markdown("---")
    
    # Canlı Sayaçlar
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric(label="Üretilen Soru Havuzu", value="142,500+", delta="+1,200 bugün")
    with col2:
        st.metric(label="Aktif Öğrenci/Veli", value="18,420", delta="+340 bu hafta")
    with col3:
        st.metric(label="B2B Anlaşmalı Kurum", value="48", delta="4 Yeni")
    with col4:
        st.metric(label="Çözülen Soru Sayısı", value="2.4 Milyon", delta="%98 Başarı")
        
    st.markdown("---")
    st.subheader("💡 Platform Özellikleri")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown("### 🎯 Soru Fabrikası")
        st.write("MEB formatına uygun saniyeler içinde yeni nesil soru üretin ve veritabanına kaydedin.")
    with col_b:
        st.markdown("### 📄 Profesyonel Word Çıktısı")
        st.write("Soru havuzundaki soruları tek tıkla 2 sütunlu, yayıncı kalitesinde .docx test kağıtlarına dönüştürün.")
    with col_c:
        st.markdown("### 📊 Veli & Kurum Paneli")
        st.write("Öğrenci gelişimini takip edin, kurumsal lisanslarla okulları sisteme dahil edin.")

# --- 2. ÖĞRENCİ PORTALI & SORU ÇÖZÜMÜ ---
elif menu == "👨‍🎓 Öğrenci Portalı & Soru Çözümü":
    st.title("👨‍🎓 Öğrenci Portalı & Soru Çözüm Merkezi")
    st.markdown("Veritabanındaki güncel soruları çözerek KKP puanları toplayın ve eksiklerinizi kapatın.")
    
    tab1, tab2 = st.tabs(["📝 Soru Bankası & Test Çöz", "🏆 KKP Oyunlaştırma & Rozetler"])
    
    with tab1:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, subject, grade, question_text, option_a, option_b, option_c, option_d, correct_answer, explanation FROM questions")
        questions = cursor.fetchall()
        conn.close()
        
        if not questions:
            st.warning("Soru havuzunda henüz soru yok. 'Yapay Zeka Soru Fabrikası' modülünden soru ekleyebilirsiniz.")
        else:
            selected_q = st.selectbox("Çözmek istediğiniz soruyu seçin:", questions, format_func=lambda x: f"[{x[1]}] {x[3][:50]}...")
            if selected_q:
                st.write(f"**Soru:** {selected_q[3]}")
                user_choice = st.radio("Seçenekler:", [selected_q[4], selected_q[5], selected_q[6], selected_q[7]], key=f"q_{selected_q[0]}")
                
                if st.button("Cevabı Kontrol Et"):
                    chosen_letter = user_choice.split(")")[0].strip()
                    if chosen_letter == selected_q[8]:
                        st.success("🎉 Tebrikler, doğru cevap!")
                        st.info(f"**Çözüm Açıklaması:** {selected_q[9]}")
                    else:
                        st.error(f"❌ Yanlış! Doğru cevap: {selected_q[8]}")
                        st.info(f"**Çözüm Açıklaması:** {selected_q[9]}")

    with tab2:
        st.subheader("🏆 Kurşun Kalem Puanı (KKP) ve Durumunuz")
        st.metric("Toplam Puanınız", "1,450 KKP", "Seviye: Altın Öğrenci 🥇")
        st.progress(75)
        st.write("Sonraki seviyeye kalan soru hedefi: **50 soru**")

# --- 3. VELİ TAKİP & PERFORMANS ---
elif menu == "👪 Veli Takip & Performans":
    st.title("👪 Veli Takip ve Performans Paneli")
    st.markdown("Çocuğunuzun günlük çalışma saatlerini, test sonuçlarını ve kazanım eksiklerini anlık takip edin.")
    
    student_name = st.selectbox("Öğrenci Seçin", ["Ahmet Furkan Kurşun (8. Sınıf - LGS)"])
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Bu Hafta Çözülen Soru", "340 Soru", "+45 dünden")
    with col2:
        st.metric("Genel Başarı Oranı", "%82", "+4%")
    with col3:
        st.metric("Günlük Çalışma Süresi", "2 Saat 15 Dk", "Hedef: 2 Saat")
        
    st.markdown("---")
    st.subheader("📊 Haftalık Kazanım Analizi")
    st.write("✅ **Türkçe - Fiilimsiler:** %95 (Çok İyi)")
    st.write("⚠️ **Matematik - Kareköklü İfadeler:** %60 (Tekrar Edilmeli)")
    st.write("🚨 **Fen - Basit Makineler:** %45 (Acil Soru Çözümü Önerilir)")
    
    if st.button("Veli Raporunu İndir"):
        st.success("Veli bilgilendirme raporu başarıyla hazırlandı.")

# --- 4. B2B LİSANS & KURUM PANELİ ---
elif menu == "🏢 B2B Lisans & Kurum Paneli":
    st.title("🏢 B2B Kurumsal Lisans & Okul Paneli")
    st.markdown("Dershaneler, özel okullar ve kurs merkezleri için toplu öğrenci ve öğretmen lisans yönetimi.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Kurum Bilgileri")
        st.text_input("Kurum Adı", "Örnek Özel LGS Kursu")
        st.text_input("Yetkili Adı Soyadı", "Kürşat Han Kurşun")
        st.number_input("Aktif Öğrenci Lisans Adedi", min_value=10, max_value=1000, value=150)
        st.button("Lisans Paketini Güncelle")
        
    with col2:
        st.subheader("Kurum Öğretmenleri")
        st.write("Sisteme kayıtlı öğretmenler kendi adına soru üretebilir ve kurumsal test indirebilir.")
        if st.button("Öğretmenleri Listele"):
            st.json(["Uğur Hoca (Matematik)", "Adil Hoca (Türkçe)", "Zeynep Hoca (Fen Bilimleri)"])

# --- 5. YAPAY ZEKA SORU FABRİKASI ---
elif menu == "🤖 Yapay Zeka Soru Fabrikası":
    st.title("🤖 Yapay Zeka Soru Fabrikası")
    st.markdown("MEB beceri temelli sınav formatına uygun, saniyeler içinde yeni nesil soru üretin ve veritabanına ekleyin.")
    
    with st.form("question_factory_form"):
        subject = st.selectbox("Ders Seçimi", ["LGS Türkçe", "LGS Matematik", "LGS Fen Bilimleri", "YKS Türkçe", "YKS Matematik"])
        grade = st.selectbox("Sınıf Seviyesi", ["8. Sınıf (LGS)", "11. Sınıf", "12. Sınıf (YKS)"])
        topic = st.text_input("Konu Başlığı (Örn: Sözel Mantık / Paragraf Yapısı)")
        count = st.slider("Üretilecek Soru Adedi", min_value=1, max_value=10, value=3)
        
        submitted = st.form_submit_button("Soruları Üret ve Havuza Ekle")
        if submitted:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            for i in range(count):
                q_text_sample = f"[{topic}] ile ilgili sistemde üretilen beceri temelli yeni nesil soru #{i+1}"
                cursor.execute("""
                    INSERT INTO questions (subject, grade, question_text, option_a, option_b, option_c, option_d, correct_answer, explanation)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (subject, grade, q_text_sample, "A) Seçenek 1", "B) Seçenek 2", "C) Seçenek 3", "D) Doğru Seçenek", "D", "Detaylı çözüm açıklaması ve ipucu."))
            conn.commit()
            conn.close()
            st.success(f"Başarıyla {count} adet soru üretildi ve veritabanına kaydedildi!")

# --- 6. 2 SÜTUNLU WORD TEST ÜRETİCİ ---
elif menu == "📄 2 Sütunlu Word Test Üretici":
    st.title("📄 2 Sütunlu Profesyonel Word (.docx) Test Üretici")
    st.markdown("Veritabanındaki soruları yayıncı standartlarında iki sütunlu şık test kağıtları olarak indirin.")
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT id, subject, grade, question_text, option_a, option_b, option_c, option_d FROM questions")
    all_qs = cursor.fetchall()
    conn.close()
    
    if not all_qs:
        st.warning("Veritabanında henüz soru bulunmuyor. Önce 'Yapay Zeka Soru Fabrikası' modülünden soru üretin.")
    else:
        st.write(f"Mevcut Soru Havuzundaki Toplam Soru: **{len(all_qs)}**")
        test_title_input = st.text_input("Test Başlığı", "KURŞUN KALEM - LGS TÜRKÇE DENEMESİ")
        
        if st.button("2 Sütunlu Word Dosyasını (.docx) Oluştur"):
            file_path = create_double_column_docx(all_qs, test_title=test_title_input)
            with open(file_path, "rb") as f:
                st.download_button(
                    label="📥 Word Test Dosyasını İndir",
                    data=f,
                    file_name="kursun_kalem_test.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            st.success("Test dosyası başarıyla hazırlandı!")

# --- 7. PAKETLER & ABONELİK (FREEMIUM) ---
elif menu == "💎 Paketler & Abonelik (Freemium)":
    st.title("💎 Abonelik Paketleri ve Fiyatlandırma")
    st.markdown("İster bireysel öğrenci, ister kurum olun; size en uygun paketi seçin.")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🟢 Freemium (Ücretsiz)")
        st.write("- Sınırlı Soru Havuzu")
        st.write("- Günlük 5 Soru Çözümü")
        st.write("- Temel Veli Paneli")
        st.markdown("**Ücretsiz**")
        st.button("Mevcut Plan", key="free_btn")
        
    with col2:
        st.markdown("### 🚀 Bireysel Pro (Öğrenci)")
        st.write("- Sınırsız Soru Fabrikası")
        st.write("- Sınırsız Test Çözümü")
        st.write("- Detaylı Veli Takip Raporu")
        st.write("- KKP Puan ve Rozetler")
        st.markdown("**299 TL / Ay**")
        st.button("Hemen Abone Ol (İyzico)", key="pro_btn")
        
    with col3:
        st.markdown("### 🏢 Kurumsal (B2B Lisans)")
        st.write("- Okul / Kurs Paneli")
        st.write("- Sınırsız Öğretmen & Öğrenci")
        st.write("- Kuruma Özel Logo ile Word Çıktısı")
        st.write("- Öncelikli Teknik Destek")
        st.markdown("**4.999 TL / Yıl**")
        st.button("Kurumsal Başvuru Yap", key="b2b_btn")

# --- ALT BİLGİ ---
st.markdown("---")
st.caption("© 2026 Kurşun Kalem Dijital Eğitim Teknolojileri | Tüm hakları saklıdır. Gri Cafer & Beyaz Cevdet onaylıdır 🐾")
