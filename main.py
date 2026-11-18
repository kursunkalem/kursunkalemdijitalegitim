import os
import json
import sqlite3
import datetime
import time
from io import BytesIO
import streamlit as st
from google import genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader

API_KEY = "AQ.Ab8RN6Ln16aqw1zeiFjrpBNsDE9rLkAwtjet6QfCfMam4Uzu6w"

st.set_page_config(page_title="Kurşun Kalem Dijital Eğitim - Geleceği Yazan Dijital Eğitim", page_icon="✏️", layout="wide")

def kurumsal_logo_goster():
    st.markdown("""
        <div style="display: flex; align-items: center; gap: 12px; background: linear-gradient(135deg, #2c3e50, #34495e); padding: 15px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <div style="background: #e74c3c; width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 24px; color: white; font-weight: bold; border: 2px solid #f39c12;">✏️</div>
            <div>
                <h2 style="color: white; margin: 0; font-size: 20px; font-family: sans-serif;">KURŞUN KALEM DİJİTAL EĞİTİM</h2>
                <p style="color: #f1c40f; margin: 0; font-size: 13px; font-weight: 600;">Geleceği yazan dijital eğitim • Profesyonel Soru Fabrikası & Akıllı Öğrenci Portalı</p>
            </div>
        </div>
    """, unsafe_allow_html=True)

MAARIF_EGITIM_CERCEVESI = """
TÜRKİYE YÜZYILI MAARİF MODELİ ORTAK EĞİTİM VE YETKİNLİK ÇERÇEVESİ:
1. Erdem-Değer-Eylem Bütünlüğü: Öğrencilerin bilgelik, erdem ve ahlaki olgunluğa ulaşması; edinilen bilgilerin eyleme ve karaktere dönüştürülmesi.
2. Beceriler Çerçevesi: Zihinsel beceriler (analiz, sentez, eleştirel düşünme, problem çözme), sosyal-duygusal beceriler ve eylem becerileri.
3. Okuryazarlık Türleri: Bilgi, dijital, finansal, sanat, vatandaşlık ve sürdürülebilirlik okuryazarlığı.
4. Soru ve Ölçme Standartları: Ezbere dayalı sorulardan uzak; bağlam kurabilen, günlük hayatla ilişkilendirilmiş, mantık-muhakeme, veri okuryazarlığı ve metinler arası sentez yeteneğini ölçen beceri temelli yapılar.
"""

def init_db():
    conn = sqlite3.connect("okul.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS ogrenciler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            isim TEXT,
            puan INTEGER,
            cozulen_soru INTEGER,
            dogru_sayisi INTEGER
        )
    """)
    c.execute("SELECT COUNT(*) FROM ogrenciler")
    if c.fetchone()[0] == 0:
        ornekler = [
            ("Kürşat Kurşun", 450, 45, 41),
            ("Burak Aygören", 380, 40, 34),
            ("Elif Öztürk", 520, 55, 50),
            ("Kullanıcı Öğrenci", 120, 12, 10)
        ]
        c.executemany("INSERT INTO ogrenciler (isim, puan, cozulen_soru, dogru_sayisi) VALUES (?, ?, ?, ?)", ornekler)
    conn.commit()
    conn.close()

init_db()

@st.cache_data
def load_kazanimlar():
    dosya_adi = "kazanimlar.json"
    varsayilan_veri = {
        "Türkçe": {
            "5. Sınıf": {
                "Okuma ve Anlama": [
                    "T.5.1.1. Bağlamdan yararlanarak bilmediği kelime gruplarının anlamını tahmin eder.",
                    "T.5.1.7. Metinle ilgili soruları cevaplar.",
                    "T.5.1.12. Grafik, tablo ve çizelgeyle sunulan bilgileri yorumlar."
                ]
            },
            "6. Sınıf": {
                "Okuma ve Anlama": [
                    "T.6.1.4. Metindeki ana düşünceyi / ana duyguyu belirler.",
                    "T.6.1.8. Metindeki yardımcı düşünceleri tespit eder."
                ]
            },
            "7. Sınıf": {
                "Okuma ve Anlama": [
                    "T.7.1.3. Deyim ve atasözlerinin metne katkısını kavrar.",
                    "T.7.1.9. Metnin anlatım biçimlerini belirler."
                ]
            },
            "8. Sınıf (LGS / Beceri Temelli)": {
                "Okuma, Mantık ve Muhakeme": [
                    "T.8.1.5. Metindeki söz sanatlarını tespit ederek işlevlerini kavrar.",
                    "T.8.1.11. Grafik, tablo ve çizelgeyle sunulan verileri yorumlayarak mantık-muhakeme yapar.",
                    "T.8.1.18. Farklı kaynaklardan alınan metinler arasında sentez ve karşılaştırma yapar."
                ]
            }
        },
        "Matematik": {
            "8. Sınıf (LGS / Beceri Temelli)": {
                "Sayılar ve İşlemler": [
                    "M.8.1.1. Verilen pozitif tam sayıların pozitif tam sayı çarpanlarını bulur.",
                    "M.8.1.2. İki doğal sayının en büyük ortak bölenini (EBOB) ve en küçük ortak katını (EKOK) hesaplar."
                ]
            }
        }
    }
    if not os.path.exists(dosya_adi):
        try:
            with open(dosya_adi, "w", encoding="utf-8") as f:
                json.dump(varsayilan_veri, f, ensure_ascii=False, indent=2)
        except: pass
    try:
        with open(dosya_adi, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return varsayilan_veri

kazanim_db = load_kazanimlar()

def load_referanslar():
    dosya_adi = "referans_sorular.json"
    if not os.path.exists(dosya_adi):
        try:
            with open(dosya_adi, "w", encoding="utf-8") as f:
                json.dump([], f, ensure_ascii=False, indent=2)
        except: pass
    try:
        with open(dosya_adi, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []

referans_havuzu = load_referanslar()

def load_arsiv():
    dosya_adi = "soru_arsivi.json"
    if not os.path.exists(dosya_adi):
        return []
    try:
        with open(dosya_adi, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []

def save_arsiv(yeni_test):
    arsiv = load_arsiv()
    arsiv.insert(0, yeni_test)
    try:
        with open("soru_arsivi.json", "w", encoding="utf-8") as f:
            json.dump(arsiv[:20], f, ensure_ascii=False, indent=2)
    except: pass

st.sidebar.title("🧭 Sistem Navigasyonu")
ana_modul = st.sidebar.radio(
    "Çalışma Alanı Seçin:",
    [
        "🏠 Ana Sayfa & 2027 LGS Geri Sayım",
        "🛠️ Soru & Redaksiyon Fabrikası (Yazar Paneli)",
        "🎓 Öğrenci Portalı & Soru Çözüm",
        "👪 Veli Takip Paneli",
        "🏢 B2B Lisans & Kurum Yönetimi",
        "💎 Paketler & Abonelik Vitrini"
    ]
)
st.sidebar.markdown("---")

# --- ANA SAYFA & GERİ SAYIM ---
if ana_modul == "🏠 Ana Sayfa & 2027 LGS Geri Sayım":
    kurumsal_logo_goster()
    st.title("📚 Kurşun Kalem Dijital Eğitim Platformuna Hoş Geldiniz")
    st.markdown("İlkokuldan YKS'ye uzanan, yapay zeka destekli Türkiye'nin en akıllı soru ve koçluk ekosistemi.")
    
    # --- CANLI AKAN LGS 2027 GERİ SAYIM SAYACI ---
    st.markdown("### ⏳ 2027 LGS Sınavına Geri Sayım (13 Haziran 2027)")
    
    # Canlı sayaç konteyneri ve otomatik yenileme döngüsü
    sayac_yeri = st.empty()
    
    exam_date = datetime.datetime(2027, 6, 13, 9, 30, 0)
    now = datetime.datetime.now()
    remaining = exam_date - now
    
    if remaining.total_seconds() > 0:
        days = remaining.days
        hours, remainder = divmod(remaining.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        
        with sayac_yeri.container():
            c1, c2, c3, c4 = st.columns(4)
            c1.metric(label="📅 Gün", value=f"{days} Gün")
            c2.metric(label="⏰ Saat", value=f"{hours} Saat")
            c3.metric(label="⏱️ Dakika", value=f"{minutes} Dk")
            c4.metric(label="⚡ Saniye", value=f"{seconds} Sn")
    else:
        sayac_yeri.success("🎉 LGS 2027 Sınavı Gerçekleşti!")

    st.markdown("---")
    
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

# --- YAZAR / SORU FABRİKASI PANELİ ---
elif ana_modul == "🛠️ Soru & Redaksiyon Fabrikası (Yazar Paneli)":
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "🎯 Soru Üretici",
        "✍️ Profesyonel Redaktör & Editör",
        "🤖 Akıllı Komut (Asistan)",
        "📚 Geçmiş Arşivi",
        "➕ Soru Ekle",
        "📄 Çoklu Format Dosya Yükle",
        "⚙️ Havuz Yönetimi"
    ])

    with tab7:
        kurumsal_logo_goster()
        st.header("Mevcut Referans Soru & Dosya Havuzu")
        if not referans_havuzu:
            st.info("Henüz havuza eklenmiş referans soru veya dosya bulunmuyor.")
        else:
            for idx, ref in enumerate(referans_havuzu):
                with st.expander(f"Kayıt {idx + 1} - {ref.get('brans', 'Türkçe')} / {ref.get('sinif', '')} ({ref.get('kazanim', '')[:40]}...)"):
                    st.write(f"**İçerik / Öncül Metin:** {ref.get('baglam_metni')}")
                    st.write(f"**Soru / Başlık:** {ref.get('soru_koku')}")
                    st.write(f"**Doğru Cevap:** {ref.get('dogru_cevap')}")

            if st.button("🗑️ Tüm Referans ve Dosya Havuzunu Temizle", type="secondary"):
                try:
                    with open("referans_sorular.json", "w", encoding="utf-8") as f:
                        json.dump([], f, ensure_ascii=False, indent=2)
                    st.success("Referans havuzu sıfırlandı.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Sıfırlama hatası: {e}")

    with tab6:
        kurumsal_logo_goster()
        st.header("📄 / 🖼️ PDF, Word (.docx) veya Görsel (JPEG/PNG) ile Soru Kaynağı Yükle")
        st.markdown("Elindeki deneme sınavı PDF, Word belgesi ya da sınav fotoğrafını buraya yükle. Program içeriği otomatik olarak okuyup yapay zekanın referans havuzuna işleyecektir.")
        
        yuklenen_dosya = st.file_uploader("Dosya Seçin", type=["pdf", "docx", "jpg", "jpeg", "png"])
        
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            p_brans = st.selectbox("Kaynak Branşı", list(kazanim_db.keys()), key="p_brans")
        with col_p2:
            p_sinif = st.selectbox("Sınıf Seviyesi", ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS / Beceri Temelli)"], key="p_sinif")
        
        p_kazanim = st.text_input("Hedef Kazanım veya Konu Açıklaması", value="Genel Deneme Analizi ve Beceri Temelli Soru Standardı")

        if yuklenen_dosya is not None:
            dosya_uzantisi = yuklenen_dosya.name.split('.')[-1].lower()
            cikti_metni = ""

            try:
                if dosya_uzantisi == "pdf":
                    reader = PdfReader(yuklenen_dosya)
                    for page in reader.pages:
                        metin = page.extract_text()
                        if metin: cikti_metni += metin + "\n"
                    st.success(f"PDF başarıyla okundu! Toplam {len(reader.pages)} sayfa metin çıkarıldı.")
                
                elif dosya_uzantisi == "docx":
                    doc_file = Document(yuklenen_dosya)
                    for p in doc_file.paragraphs:
                        if p.text: cikti_metni += p.text + "\n"
                    st.success("Word belgesi başarıyla okundu!")
                
                elif dosya_uzantisi in ["jpg", "jpeg", "png"]:
                    with st.spinner("Görsel yapay zeka ile taranıyor ve metne dönüştürülüyor..."):
                        client = genai.Client(api_key=API_KEY)
                        dosya_bytes = yuklenen_dosya.getvalue()
                        
                        image_resp = client.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=[
                                {
                                    "inline_data": {
                                        "mime_type": f"image/{'jpeg' if dosya_uzantisi in ['jpg', 'jpeg'] else 'png'}",
                                        "data": dosya_bytes
                                    }
                                },
                                "Bu görseldeki tüm test sorularını, bağlam metinlerini, seçenekleri ve soru köklerini eksiksiz olarak metne dök."
                            ]
                        )
                        cikti_metni = image_resp.text
                        st.success("Görsel başarıyla tarandı ve metne dönüştürüldü!")

                if cikti_metni:
                    with st.expander("📖 Okunan Dosya / Görsel İçeriğini Önizle"):
                        st.text_area("Çıkarılan İçerik Metni", cikti_metni[:2500] + ("..." if len(cikti_metni) > 2500 else ""), height=200)

                    if st.button("📥 İçeriği Referans Havuzuna İşle", type="primary"):
                        yeni_referans = {
                            "brans": p_brans, "sinif": p_sinif, "kazanim": p_kazanim,
                            "baglam_metni": f"[{dosya_uzantisi.upper()} DOSYA/GÖRSEL KAYNAĞI]\n{cikti_metni[:4000]}",
                            "soru_koku": f"Yüklenen {dosya_uzantisi.upper()} dosyasındaki kurumsal soru tarzı ve standartları",
                            "sik_a": "Referans", "sik_b": "Referans", "sik_c": "Referans", "sik_d": "Referans", "dogru_cevap": "A"
                        }
                        mevcut_veriler = load_referanslar()
                        mevcut_veriler.append(yeni_referans)

                        with open("referans_sorular.json", "w", encoding="utf-8") as f:
                            json.dump(mevcut_veriler, f, ensure_ascii=False, indent=2)
                        st.success("Dosya içeriği başarıyla referans havuzuna eklendi! Artık soru üretirken bu kaynak baz alınacak.")
            except Exception as e:
                st.error(f"Dosya işlenirken bir hata oluştu: {e}")

    with tab5:
        kurumsal_logo_goster()
        st.header("Havuza Yeni Referans Soru Ekle")
        with st.form("referans_form"):
            ref_brans = st.selectbox("Branş", list(kazanim_db.keys()))
            ref_sinif = st.selectbox("Sınıf Seviyesi", ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS / Beceri Temelli)"])
            ref_kazanim = st.text_input("Kazanım Açıklaması")
            ref_baglam = st.text_area("Bağlam Metni / Senaryo / Veri")
            ref_kok = st.text_input("Soru Kökü")
            
            col1, col2 = st.columns(2)
            with col1:
                ref_a = st.text_input("A Şıkkı")
                ref_c = st.text_input("C Şıkkı")
            with col2:
                ref_b = st.text_input("B Şıkkı")
                ref_d = st.text_input("D Şıkkı")
                
            ref_dogru = st.selectbox("Doğru Cevap", ["A", "B", "C", "D"])
            submitted = st.form_submit_button("💾 Referans Havuzuna Kaydet")
            
            if submitted:
                yeni_soru = {
                    "brans": ref_brans, "sinif": ref_sinif, "kazanim": ref_kazanim,
                    "baglam_metni": ref_baglam, "soru_koku": ref_kok,
                    "sik_a": ref_a, "sik_b": ref_b, "sik_c": ref_c, "sik_d": ref_d, "dogru_cevap": ref_dogru
                }
                mevcut_veriler = load_referanslar()
                mevcut_veriler.append(yeni_soru)
                try:
                    with open("referans_sorular.json", "w", encoding="utf-8") as f:
                        json.dump(mevcut_veriler, f, ensure_ascii=False, indent=2)
                    st.success("Referans soru başarıyla kaydedildi!")
                except Exception as e:
                    st.error(f"Kayıt hatası: {e}")

    with tab4:
        kurumsal_logo_goster()
        st.header("📚 Daha Önce Üretilen Test Arşivi")
        arsiv_listesi = load_arsiv()
        if not arsiv_listesi:
            st.info("Henüz arşivde kayıtlı test bulunmuyor.")
        else:
            for i, test in enumerate(arsiv_listesi):
                with st.expander(f"📁 {test.get('baslik', 'Test')} ({len(test.get('sorular', []))} Soru)"):
                    st.write(f"**Kurum / Başlık:** {test.get('kurum')}")
                    for s_idx, soru in enumerate(test.get('sorular', []), 1):
                        st.markdown(f"**Soru {s_idx}:** {soru.get('soru_koku')}")
                        st.write(f"Doğru Cevap: **{soru.get('dogru_cevap')}**")

    with tab3:
        kurumsal_logo_goster()
        st.title("🤖 Kurşun Kalem Akıllı Asistanı")
        st.markdown("Yapay zekaya doğal dilde komut vererek dilediğiniz kriterlerde anında soru ve test ürettirin.")
        
        serbest_komut = st.text_area("Ne tür bir test üretmek istiyorsunuz?", placeholder="Örn: 8. sınıf fen bilimleri DNA ve genetik kod ile ilgili orta-zor seviyede 2 soru hazırla.")
        
        if st.button("🚀 Asistanla Üret", type="primary"):
            if not serbest_komut.strip():
                st.warning("Lütfen bir komut yazın.")
            else:
                with st.spinner("Kurşun Kalem Asistanı soruları hazırlıyor..."):
                    try:
                        client = genai.Client(api_key=API_KEY)
                        asistan_prompt = (
                            "Sen Kurşun Kalem Dijital Eğitim'in uzman yapay zeka eğitim asistanısın. "
                            "Kullanıcının doğal dilde verdiği komutu analiz et ve Türkiye Yüzyılı Maarif Modeli'ne uygun, "
                            "özgün çoktan seçmeli sorular üret. Çözüm analizi istemiyorum. "
                            "Çıktıyı YALNIZCA geçerli bir JSON dizisi (liste) formatında (hiçbir ek metin, açıklama veya backtick eklemeden) şu yapıda ver:\n"
                            "[\n  {\n    \"baglam_metni\": \"...\",\n    \"soru_koku\": \"...\",\n    \"sik_a\": \"...\",\n    \"sik_b\": \"...\",\n    \"sik_c\": \"...\",\n    \"sik_d\": \"...\",\n    \"dogru_cevap\": \"A/B/C/D\"\n  }\n]"
                        )
                        resp = client.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=serbest_komut,
                            config={
                                "system_instruction": asistan_prompt,
                                "temperature": 0.4,
                                "max_output_tokens": 2048
                            }
                        )
                        text = resp.text.strip()
                        if text.startswith("```json"): text = text[7:]
                        if text.endswith("```"): text = text[:-3]
                        
                        asistan_sorular = json.loads(text.strip())
                        st.success(f"{len(asistan_sorular)} adet soru asistan tarafından başarıyla üretildi!")
                        
                        for idx, s in enumerate(asistan_sorular, 1):
                            with st.expander(f"Asistan Soru {idx}"):
                                if s.get("baglam_metni"): st.write(s.get("baglam_metni"))
                                st.markdown(f"**{s.get('soru_koku')}**")
                                st.write(f"A) {s.get('sik_a')}")
                                st.write(f"B) {s.get('sik_b')}")
                                st.write(f"C) {s.get('sik_c')}")
                                st.write(f"D) {s.get('sik_d')}")
                                st.success(f"Doğru Cevap: {s.get('dogru_cevap')}")
                    except Exception as e:
                        st.error(f"Asistan üretim hatası: {e}")

    with tab2:
        kurumsal_logo_goster()
        st.title("✍️ Profesyonel Metin Redaktörü & Editör")
        st.markdown("Yazdığın, dışarıdan aldığın veya düzenlenmesini istediğin soru metinlerini, paragrafları veya test taslaklarını **Türkiye Yüzyılı Maarif Modeli** standartlarına, dilbilgisi kurallarına ve akademik akıcılığa göre kusursuz hale getir.")
        
        redakte_edilecek_metin = st.text_area(
            "Redakte ve Editörlük Yapılacak Metni / Soruyu Buraya Yapıştırın:", 
            height=250,
            placeholder="Örn: Buraya hatalı veya geliştirilmesi gereken soru metnini, açıklamayı veya paragrafı yapıştırın..."
        )
        
        col_ed1, col_ed2 = st.columns(2)
        with col_ed1:
            editor_modu = st.selectbox(
                "Düzenleme ve Redaksiyon Odağı",
                [
                    "Akademik & Maarif Modeli Dil Standartlaştırma",
                    "Soru Kökü ve Çeldirici Kalitesini Artırma",
                    "Yazım, Noktalama ve İmlâ Kusursuzlaştırma",
                    "Metinler Arası Sentez ve Mantık Denetimi"
                ]
            )
        with col_ed2:
            hedef_kitle = st.selectbox("Hedef Sınıf / Seviye", ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS Standardı)"])

        if st.button("✨ Metni Profesyonelce Redakte Et", type="primary"):
            if not redakte_edilecek_metin.strip():
                st.warning("Lütfen redakte edilecek bir metin girin.")
            else:
                with st.spinner("Baş Editör ve Redaktör metni inceliyor ve Maarif Modeli standartlarına göre yeniden yapılandırıyor..."):
                    try:
                        client = genai.Client(api_key=API_KEY)
                        editor_system_prompt = (
                            "Sen kıdemli bir Eğitim Yayıncılığı Baş Editörü ve Baş Redaktörüsün. "
                            "Aşağıdaki eğitim çerçevesine ve Maarif Modeli esaslarına tam bağlı kalarak metni profesyonelce redakte et:\n"
                            + MAARIF_EGITIM_CERCEVESI + "\n"
                            "Görevin: Kullanıcının verdiği ham ya da kusurlu metni / soruyu; Türkçe dil kurallarına, "
                            "akademik terminolojiye, LGS ve kazanım ölçme standartlarına göre kusursuzlaştır. "
                            "Çıktıda önce 'Düzenlenmiş Metin / Soru' sun, ardından alt bölümde yapılan düzeltmeleri ve pedagojik gerekçeleri kısaca maddeler halinde açıkla."
                        )

                        editor_resp = client.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=f"Odak Modu: {editor_modu}\nHedef Seviye: {hedef_kitle}\n\nHam Metin:\n{redakte_edilecek_metin}",
                            config={
                                "system_instruction": editor_system_prompt,
                                "temperature": 0.4,
                                "max_output_tokens": 2048
                            }
                        )

                        st.success("Redaksiyon işlemi tamamlandı!")
                        st.markdown("### 📝 Editör Onaylı Sonuç:")
                        st.markdown(editor_resp.text)
                    except Exception as e:
                        st.error(f"Redaksiyon sırasında hata oluştu: {e}")

    with tab1:
        kurumsal_logo_goster()
        st.title("🎯 Kurşun Kalem Dijital Eğitim - Soru Üretim Stüdyosu")
        st.markdown("Türkiye Yüzyılı Maarif Modeli standartlarında, **Otomatik Redaktör ve Kalite Kontrollü** ve **2 Sütunlu LGS Kitapçık Formatında** profesyonel değerlendirme testleri oluşturun.")

        with st.sidebar:
            kurumsal_logo_goster()
            st.header("⚙️ Test ve Kurum Parametreleri")
            
            kurum_adi = st.text_input("Kurum / Okul Adı", value="Kurşun Kalem Dijital Eğitim")
            secilen_brans = st.selectbox("Branş Seçimi", list(kazanim_db.keys()), key="yazar_brans")
            
            sinif_listesi = list(kazanim_db[secilen_brans].keys())
            secilen_sinif = st.selectbox("Sınıf Seviyesi", sinif_listesi, key="yazar_sinif")
            
            alan_listesi = list(kazanim_db[secilen_brans][secilen_sinif].keys())
            secilen_alan = st.selectbox("Öğrenme Alanı", alan_listesi, key="yazar_alan")
            
            secilen_kazanim = st.selectbox("Hedef Kazanım", kazanim_db[secilen_brans][secilen_sinif][secilen_alan], key="yazar_kazanim")
            
            soru_adedi = st.slider("Üretilecek Soru Adedi (Test Modu)", min_value=1, max_value=5, value=2)
            zorluk_seviyesi = st.select_slider("Zorluk Derecesi", options=["Orta", "Orta-Zor", "İleri Düzey"], value="Orta-Zor")
            tema_girdisi = st.text_input("Bağlam / Tema / Senaryo", placeholder="Örn: Sürdürülebilirlik, deney analizi...")
            
            kalite_kontrol = st.checkbox("🔍 Otomatik Soru Kalite & Redaksiyon Denetimi", value=True, help="Üretilen soruların çeldirici kalitesini ve kazanım uyumunu arka planda denetler.")
            
            uret_butonu = st.button("✨ Kurumsal Soru Paketi Üret", type="primary", use_container_width=True)

        referanslar_str = json.dumps(referans_havuzu, ensure_ascii=False, indent=2)
        if len(referanslar_str) > 3000:
            referanslar_str = referanslar_str[:3000] + "\n...(Kalan referanslar hız için kırpıldı)"

        SYSTEM_PROMPT = (
            f"Sen uzman bir {secilen_brans} eğitimcisi, müfredat geliştirme uzmanı ve kıdemli soru yazarısın.\n"
            f"Görevin; {kurum_adi} akademik çizgisine ve aşağıdaki Türkiye Yüzyılı Maarif Modeli ortak çerçevesine tam uygun, "
            "öğrencinin zihinsel becerilerini, analiz ve sentez yeteneğini geliştiren özgün çoktan seçmeli sorular üretmektir. Çözüm analizi İSTEMİYORUM.\n\n"
            + MAARIF_EGITIM_CERCEVESI + "\n\n"
            "Aşağıda verilen YÜKLENEN DOSYA/GÖRSEL KAYNAK verilerindeki soru kalıplarını, çeldirici niteliklerini ve metin uzunluklarını incele; benzer nitelikte özgün sorular üret.\n"
            "Çıktıyı YALNIZCA geçerli bir JSON dizisi (liste) formatında (hiçbir ek metin, açıklama veya markdown backtick eklemeden) şu yapıda ver:\n"
            "[\n  {\n    \"baglam_metni\": \"...\",\n    \"soru_koku\": \"...\",\n    \"sik_a\": \"...\",\n    \"sik_b\": \"...\",\n    \"sik_c\": \"...\",\n    \"sik_d\": \"...\",\n    \"dogru_cevap\": \"A/B/C/D\"\n  }\n]\n\n"
            "REFERANS DOSYA İÇERİKLERİ:\n" + referanslar_str
        )

        def toplu_soru_uret(brans, sinif, kazanim, tema, zorluk, adet, validator_aktif):
            client = genai.Client(api_key=API_KEY)
            kullanici_istegi = (
                f"Kriterler:\n- Branş: {brans}\n- Sınıf: {sinif}\n- Kazanım: {kazanim}\n"
                f"- Tema/Bağlam: {tema}\n- Zorluk: {zorluk}\n- Üretilecek Soru Adedi: {adet}\n"
                f"Bu kriterlere uygun tam {adet} adet özgün beceri temelli soru üret ve JSON listesi olarak dön."
            )

            response = client.models.generate_content(
                model="gemini-3.8-flash",
                contents=kullanici_istegi,
                config={
                    "system_instruction": SYSTEM_PROMPT,
                    "temperature": 0.4,
                    "max_output_tokens": 2048
                }
            )

            text = response.text.strip()
            if text.startswith("```json"): text = text[7:]
            if text.endswith("```"): text = text[:-3]
            sorular = json.loads(text.strip())

            if validator_aktif and sorular:
                validator_prompt = (
                    "Sen bir Baş Redaktör ve Kalite Kontrol Uzmanısın. Aşağıdaki soru listesini incele; "
                    "eğer şıklarda veya soru kökünde mantık hatası, zayıf çeldirici veya kazanım uyumsuzluğu varsa düzelt. "
                    "Çıktıyı YALNIZCA aynı JSON listesi formatında geri ver:"
                )
                val_resp = client.models.generate_content(
                    model="gemini-3.8-flash",
                    contents=json.dumps(sorular, ensure_ascii=False),
                    config={
                        "system_instruction": validator_prompt,
                        "temperature": 0.3,
                        "max_output_tokens": 2048
                    }
                )
                val_text = val_resp.text.strip()
                if val_text.startswith("```json"): val_text = val_text[7:]
                if val_text.endswith("```"): val_text = val_text[:-3]
                try:
                    sorular = json.loads(val_text.strip())
                except: pass

            return sorular

        def word_iki_sutunlu_olustur(kurum, brans, sinif, kazanim, soru_listesi):
            doc = Document()
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(0.8)
                section.bottom_margin = Inches(0.8)
                section.left_margin = Inches(0.8)
                section.right_margin = Inches(0.8)

            p_logo = doc.add_paragraph()
            run_logo = p_logo.add_run("✏️ [ KURŞUN KALEM DİJİTAL EĞİTİM ]")
            run_logo.bold = True
            run_logo.font.color.rgb = RGBColor(192, 57, 43)
            run_logo.font.size = Pt(12)

            doc.add_heading(f"{kurum} - {brans} ({sinif}) Değerlendirme Testi", level=1)
            p_info = doc.add_paragraph("Adı Soyadı: ..........................................      Sınıf/No: ............ / ............ ")
            p_info.runs[0].italic = True
            p_kazanim = doc.add_paragraph(f"Hedef Kazanım: {kazanim}")
            p_kazanim.runs[0].italic = True
            doc.add_paragraph("-" * 65)

            table = doc.add_table(rows=1, cols=2)
            table.autofit = False

            table.columns[0].width = Inches(3.3)
            table.columns[1].width = Inches(3.3)

            cell_left = table.cell(0, 0)
            cell_right = table.cell(0, 1)

            for i, soru in enumerate(soru_listesi, 1):
                target_cell = cell_left if (i % 2 != 0) else cell_right

                p_sno = target_cell.add_paragraph()
                p_sno.add_run(f"Soru {i}").bold = True

                if soru.get("baglam_metni"):
                    target_cell.add_paragraph(soru.get("baglam_metni"))

                p_kok = target_cell.add_paragraph()
                p_kok.add_run(soru.get("soru_koku", "")).bold = True

                target_cell.add_paragraph(f"A) {soru.get('sik_a', '')}")
                target_cell.add_paragraph(f"B) {soru.get('sik_b', '')}")
                target_cell.add_paragraph(f"C) {soru.get('sik_c', '')}")
                target_cell.add_paragraph(f"D) {soru.get('sik_d', '')}")
                target_cell.add_paragraph("-" * 30)

            doc.add_page_break()
            doc.add_heading("Cevap Anahtarı", level=1)
            for i, soru in enumerate(soru_listesi, 1):
                p_cevap = doc.add_paragraph()
                p_cevap.add_run(f"Soru {i} Doğru Cevap: {soru.get('dogru_cevap', '')}").bold = True

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            return buffer

        if uret_butonu:
            if not API_KEY:
                st.error("API anahtarı tanımlı değil.")
            elif not tema_girdisi.strip():
                st.warning("Lütfen bir bağlam veya tema girin.")
            else:
                with st.spinner(f"{kurum_adi} standartlarında redaksiyon denetimiyle {soru_adedi} soru hazırlanıyor..."):
                    try:
                        uretilen_sorular = toplu_soru_uret(
                            secilen_brans, secilen_sinif, secilen_kazanim,
                            tema_girdisi, zorluk_seviyesi, soru_adedi, kalite_kontrol
                        )
                        st.session_state["toplu_sorular"] = uretilen_sorular
                        st.session_state["test_meta"] = (kurum_adi, secilen_brans, secilen_sinif, secilen_kazanim)

                        save_arsiv({
                            "baslik": f"{secilen_brans} - {secilen_sinif} Testi",
                            "kurum": kurum_adi,
                            "sorular": uretilen_sorular
                        })
                    except Exception as e:
                        st.error(f"Soru üretim hatası oluştu: {e}")

        if "toplu_sorular" in st.session_state:
            sorular = st.session_state["toplu_sorular"]
            k_kurum, k_brans, k_sinif, k_kazanim = st.session_state["test_meta"]

            st.success(f"✨ {len(sorular)} adet özgün soru redaksiyon ve kalite süzgecinden geçerek başarıyla üretildi!")

            for idx, soru in enumerate(sorular, 1):
                with st.expander(f"📄 Soru {idx} Önizlemesi"):
                    if soru.get("baglam_metni"):
                        st.write(soru.get("baglam_metni"))
                    st.markdown(f"**{soru.get('soru_koku', '')}**")
                    st.write(f"**A)** {soru.get('sik_a', '')}")
                    st.write(f"**B)** {soru.get('sik_b', '')}")
                    st.write(f"**C)** {soru.get('sik_c', '')}")
                    st.write(f"**D)** {soru.get('sik_d', '')}")
                    st.success(f"**Doğru Cevap:** {soru.get('dogru_cevap', '')}")

            docx_buffer = word_iki_sutunlu_olustur(k_kurum, k_brans, k_sinif, k_kazanim, sorular)

            safe_filename = f"{k_brans}_{k_sinif}".replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')
            st.download_button(
                label="📥 2 Sütunlu LGS Formatında Word (.docx) İndir",
                data=docx_buffer,
                file_name=f"{safe_filename}_kurumsal_test.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# --- ÖĞRENCİ PORTALI & SORU ÇÖZÜM ---
elif ana_modul == "🎓 Öğrenci Portalı & Soru Çözüm":
    kurumsal_logo_goster()
    st.title("🎓 Öğrenci Portalı & Soru Çözüm Merkezi")
    st.markdown("Veritabanındaki güncel soruları çözerek KKP puanları toplayın ve eksiklerinizi kapatın.")

    ogrenci_tab1, ogrenci_tab2, ogrenci_tab3, ogrenci_tab4 = st.tabs([
        "🎯 Soru Bankası & Test Çöz",
        "🧠 Yapay Zeka Akıllı Koç",
        "📊 Kişisel Gelişim Karnem",
        "🏆 Şampiyonlar Liderlik Tablosu"
    ])

    with ogrenci_tab1:
        st.subheader("İnteraktif Soru Çözümü")
        arsiv_listesi = load_arsiv()
        if not arsiv_listesi:
            st.info("Henüz arşivde test bulunmuyor. Yazar panelinden test üreterek buraya aktarabilirsin.")
        else:
            secilen_test = st.selectbox("Test Seçin", [t.get('baslik') for t in arsiv_listesi])
            secilen_test_obj = next((t for t in arsiv_listesi if t.get('baslik') == secilen_test), None)

            if secilen_test_obj:
                for idx, soru in enumerate(secilen_test_obj.get('sorular', []), 1):
                    st.markdown(f"---")
                    st.markdown(f"### Soru {idx}")
                    if soru.get("baglam_metni"): st.write(soru.get("baglam_metni"))
                    st.markdown(f"**{soru.get('soru_koku')}**")
                    
                    cevap_secimi = st.radio(
                        f"Seçim (Soru {idx}):",
                        [f"A) {soru.get('sik_a')}", f"B) {soru.get('sik_b')}", f"C) {soru.get('sik_c')}", f"D) {soru.get('sik_d')}"],
                        key=f"ogrenci_soru_{idx}",
                        index=None
                    )

                    if st.button(f"Cevabı Kontrol Et (Soru {idx})", key=f"btn_ogrenci_{idx}"):
                        if cevap_secimi is None:
                            st.warning("Şık seçiniz.")
                        else:
                            secilen_harf = cevap_secimi[0]
                            dogru_harf = soru.get('dogru_cevap').strip().upper()
                            if secilen_harf == dogru_harf:
                                st.success(f"🎉 Tebrikler! Doğru cevap ({dogru_harf}).")
                                conn = sqlite3.connect("okul.db")
                                c = conn.cursor()
                                c.execute("UPDATE ogrenciler SET puan = puan + 15, cozulen_soru = cozulen_soru + 1, dogru_sayisi = dogru_sayisi + 1 WHERE id = 4")
                                conn.commit()
                                conn.close()
                            else:
                                st.error(f"❌ Yanlış! Doğru seçenek: **{dogru_harf}**")

    with ogrenci_tab2:
        st.subheader("🧠 Kurşun Kalem Yapay Zeka Akıllı Koç")
        st.markdown("Çözdüğün testlerin sonuçlarına göre eksiklerini analiz eden ve sana özel çalışma programı çıkaran akıllı koç asistanı.")

        ogrenci_mesaji = st.text_input("Koça sorunu yaz:", placeholder="Örn: Sözel mantık sorularında çok zaman kaybediyorum, ne yapmalıyım?")
        if st.button("🤖 Akıllı Koçtan Tavsiye Al", type="primary"):
            if not ogrenci_mesaji.strip():
                st.warning("Lütfen koça bir soru veya sorun yaz.")
            else:
                with st.spinner("Akıllı Koç analiz yapıyor ve sana özel strateji hazırlıyor..."):
                    try:
                        client = genai.Client(api_key=API_KEY)
                        koc_prompt = (
                            "Sen Kurşun Kalem Dijital Eğitim'in 'Geleceği yazan dijital eğitim' felsefesine bağlı; "
                            "ultra akıllı, motive edici ve nokta atışı taktikler veren LGS Eğitim Koçusun. "
                            "Öğrencinin yazdığı soruya veya sorununa karşılık; ona net, pratik ve taktiksel bir rehberlik sun."
                        )
                        resp = client.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=ogrenci_mesaji,
                            config={"system_instruction": koc_prompt, "temperature": 0.5}
                        )
                        st.success("Akıllı Koç Tavsiyesi:")
                        st.markdown(resp.text)
                    except Exception as e:
                        st.error(f"Koç yanıt üretirken hata oluştu: {e}")

    with ogrenci_tab3:
        st.subheader("📊 Kişisel Gelişim ve Başarı Karnem")
        conn = sqlite3.connect("okul.db")
        c = conn.cursor()
        c.execute("SELECT puan, cozulen_soru, dogru_sayisi FROM ogrenciler WHERE id = 4")
        veri = c.fetchone()
        conn.close()

        puan, soru_sayisi, dogru_sayisi = veri if veri else (120, 12, 10)
        yanlis_sayisi = soru_sayisi - dogru_sayisi if soru_sayisi >= dogru_sayisi else 0
        basari = int((dogru_sayisi / soru_sayisi) * 100) if soru_sayisi > 0 else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Toplam Puan", f"{puan} P")
        col2.metric("Çözülen Soru", f"{soru_sayisi} Soru")
        col3.metric("Doğru / Yanlış", f"{dogru_sayisi} D / {yanlis_sayisi} Y")
        col4.metric("Başarı Oranı", f"%{basari}")

        st.markdown("---")
        st.info("💡 **Analiz Notu:** Sözel mantıkta hızlandın fakat paragraf hızını artırmak için günde 15 paragraf sorusu çözmelisin.")

    with ogrenci_tab4:
        st.subheader("🏆 Şampiyonlar Ligi Liderlik Tablosu")
        conn = sqlite3.connect("okul.db")
        c = conn.cursor()
        c.execute("SELECT isim, puan, cozulen_soru FROM ogrenciler ORDER BY puan DESC")
        ogrenciler = c.fetchall()
        conn.close()

        tablo = []
        for sira, (isim, p, s) in enumerate(ogrenciler, 1):
            madalya = "🥇" if sira == 1 else ("🥈" if sira == 2 else ("🥉" if sira == 3 else "🏅"))
            tablo.append({"Sıra": f"{madalya} {sira}", "Öğrenci": isim, "Puan": p, "Çözülen Soru": s})
        st.table(tablo)

# --- VELİ TAKİP PANELİ ---
elif ana_modul == "👪 Veli Takip Paneli":
    kurumsal_logo_goster()
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
    st.subheader("📊 Haftalık Kazanım Analizi (Zayıf / Güçlü Alanlar)")
    st.write("✅ **Türkçe - Fiilimsiler:** %95 (Çok İyi)")
    st.write("⚠️ **Matematik - Kareköklü İfadeler:** %60 (Tekrar Edilmeli)")
    st.write("🚨 **Fen - Basit Makineler:** %45 (Acil Soru Çözümü Önerilir)")
    
    if st.button("Veli Raporunu İndir"):
        st.success("Veli bilgilendirme raporu başarıyla hazırlandı.")

# --- B2B LİSANS & KURUM YÖNETİMİ ---
elif ana_modul == "🏢 B2B Lisans & Kurum Yönetimi":
    kurumsal_logo_goster()
    st.title("🏢 B2B Kurumsal Lisans & Okul Paneli")
    st.markdown("Dershaneler, özel okullar ve kurs merkezleri için toplu öğrenci ve öğretmen lisans yönetimi.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Kurum Bilgileri")
        st.text_input("Kurum Adı", "Örnek Özel LGS Kursu")
        st.text_input("Yetkili Adı Soyadı", "Elif Öztürk")
        st.number_input("Aktif Öğrenci Lisans Adedi", min_value=10, max_value=1000, value=150)
        st.button("Lisans Paketini Güncelle")
        
    with col2:
        st.subheader("Kurum Öğretmenleri")
        st.write("Sisteme kayıtlı öğretmenler kendi adına soru üretebilir ve kurumsal test indirebilir.")
        if st.button("Öğretmenleri Listele"):
            st.json(["Uğur Hoca (Matematik)", "Adil Hoca (Türkçe)", "Zeynep Hoca (Fen Bilimleri)"])

# --- PAKETLER & ABONELİK VİTRİNİ ---
elif ana_modul == "💎 Paketler & Abonelik Vitrini":
    kurumsal_logo_goster()
    
    st.markdown("""
        <style>
        .paket-karti {
            background-color: #ffffff;
            border: 1px solid #e0e0e0;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.05);
            margin-bottom: 20px;
            transition: transform 0.3s ease;
        }
        .paket-karti:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 20px rgba(0,0,0,0.1);
        }
        .fiyat-etiketi {
            font-size: 24px;
            font-weight: bold;
            color: #2c3e50;
            margin-top: 10px;
            margin-bottom: 5px;
        }
        .taksit-bilgisi {
            font-size: 12px;
            color: #7f8c8d;
            margin-bottom: 15px;
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("💎 Kurşun Kalem - Eğitim Paketleri ve Fiyatlandırma Vitrini")
    st.markdown("🚀 *Geleceği yazan dijital eğitim* felsefesiyle; öğrencilerimiz ve kurumlarımız için en uygun yapay zeka destekli paketlerimiz.")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
            <div class="paket-karti">
                <h4>🎯 LGS 2027 Türkçe & Sözel Beceri Paketi</h4>
                <p style="font-size: 13px; color: #555;">LGS'de yeni nesil sözel mantık ve paragraf atlamak isteyenler için özel koçluk destekli paket.</p>
                <hr style="margin: 10px 0;">
                <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                    <li>✓ Bireysel Koçluk & Akıllı Takip</li>
                    <li>✓ Sınırsız Yapay Zeka Soru Üretimi</li>
                    <li>✓ Kişiselleştirilmiş Eksik Kapatma Testleri</li>
                    <li>✓ 2 Sütunlu LGS PDF & Word Çıktıları</li>
                    <li>✓ Şampiyonlar Liderlik Tablosu Erişimi</li>
                </ul>
                <div class="fiyat-etiketi">34.900 TL</div>
                <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 2.908 TL)</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Hemen Al (Paket 1)", key="btn_p1", use_container_width=True):
            st.success("Tebrikler! LGS Sözel Paket sepetinize eklendi. Ödeme sayfasına yönlendiriliyorsunuz...")

    with col2:
        st.markdown("""
            <div class="paket-karti">
                <h4>🚀 LGS 2027 Tüm Dersler Tam Set</h4>
                <p style="font-size: 13px; color: #555;">Türkçe, Matematik, Fen ve Tüm Sözel-Sayısal derslerde en tepeyi hedefleyenler için tam kapsamlı set.</p>
                <hr style="margin: 10px 0;">
                <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                    <li>✓ Bütün Derslerde Yapay Zeka Desteği</li>
                    <li>✓ 5.000+ Beceri Temelli Soru Bankası</li>
                    <li>✓ Akıllı Kazanım Analizi ve Karne</li>
                    <li>✓ 7/24 AI Akıllı Koçluk Asistanı</li>
                    <li>✓ Sınırsız Video / Strateji Rehberliği</li>
                </ul>
                <div class="fiyat-etiketi">44.900 TL</div>
                <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 3.741 TL)</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Hemen Al (Paket 2)", key="btn_p2", use_container_width=True):
            st.success("Tebrikler! LGS Tam Set sepetinize eklendi. Ödeme sayfasına yönlendiriliyorsunuz...")

    with col3:
        st.markdown("""
            <div class="paket-karti">
                <h4>🏢 Kurumsal Okul & Kurs Lisansı (B2B)</h4>
                <p style="font-size: 13px; color: #555;">Dershaneler, özel okullar ve kurs merkezleri için toplu öğrenci ve öğretmen lisans yönetimi.</p>
                <hr style="margin: 10px 0;">
                <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                    <li>✓ Okul / Kurs Yönetim Paneli</li>
                    <li>✓ Sınırsız Öğretmen & Öğrenci Lisansı</li>
                    <li>✓ Kuruma Özel Logo ile Word Çıktısı</li>
                    <li>✓ Öncelikli Teknik Destek</li>
                </ul>
                <div class="fiyat-etiketi">4.999 TL / Yıl</div>
                <div class="taksit-bilgisi">Kurumsal Faturalı ve Esnek Lisanslama</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("Kurumsal Başvuru Yap", key="btn_b2b", use_container_width=True):
            st.success("Kurumsal başvuru formunuz alındı. Eğitim danışmanlarımız sizinle iletişime geçecektir.")

# --- ALT BİLGİ ---
st.markdown("---")
st.caption("© 2026 Kurşun Kalem Dijital Eğitim Teknolojileri | Tüm hakları saklıdır.")
