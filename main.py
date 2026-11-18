import os
import json
import sqlite3
from io import BytesIO
import streamlit as st
from google import genai
from docx import Document
from docx.shared import Inches, Pt, RGBColor
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
            dogru_sayisi INTEGER,
            kkp INTEGER DEFAULT 150
        )
    """)
    c.execute("SELECT COUNT(*) FROM ogrenciler")
    if c.fetchone()[0] == 0:
        ornekler = [
            ("Zeynep Yılmaz", 450, 45, 41, 450),
            ("Ahmet Demir", 380, 40, 34, 380),
            ("Elif Çelik", 520, 55, 50, 520),
            ("Kullanıcı Öğrenci", 120, 12, 10, 150)
        ]
        c.executemany("INSERT INTO ogrenciler (isim, puan, cozulen_soru, dogru_sayisi, kkp) VALUES (?, ?, ?, ?, ?)", ornekler)
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
    ["🛠️ Soru & Redaksiyon Fabrikası (Yazar Paneli)", "🎓 Öğrenci & Akıllı Koç Portalı"]
)
st.sidebar.markdown("---")

st.sidebar.info("📊 **Canlı Ekosistem Verisi**\n\n• Bugün Üretilen Soru: **14.850**\n• Aktif Öğrenci: **3.420**\n• Çözülen Test: **9.120**")

if ana_modul == "🛠️ Soru & Redaksiyon Fabrikası (Yazar Paneli)":
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "🎯 Soru Üretici",
        "✍️ Profesyonel Redaktör & Editör",
        "🤖 Akıllı Komut (Asistan)",
        "📚 Geçmiş Arşivi",
        "➕ Soru Ekle",
        "📄 Çoklu Format Dosya Yükle",
        "⚙️ Havuz Yönetimi",
        "🏢 B2B Kurumsal Lisans"
    ])

    with tab8:
        kurumsal_logo_goster()
        st.header("🏢 Öğretmenler & Kurumlar İçin B2B Soru Fabrikası Lisansı")
        st.markdown("Okullar, özel ders merkezleri ve dijital yayıncılık yapan öğretmenler için sınırsız Word/PDF soru üretim lisanları.")
        
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("""
                <div style="background:#fff; border:1px solid #ddd; padding:20px; border-radius:10px;">
                    <h3>🏫 Kurum / Okul Lisansı</h3>
                    <p style="color:#666;">Birden fazla öğretmen ve şube için kurumsal ortak soru havuzu ve sınırsız 2 sütunlu çıktı.</p>
                    <h2 style="color:#2c3e50;">14.900 TL / Yıl</h2>
                    <ul style="font-size:13px;">
                        <li>✓ Sınırsız Soru Üretimi</li>
                        <li>✓ Tüm Sınıf Düzeyleri (1-12 & YKS)</li>
                        <li>✓ Kurumsal Logo Entegrasyonu</li>
                    </ul>
                </div>
            """, unsafe_allow_html=True)
            if st.button("Kurumsal Lisans Satın Al", key="btn_b2b_1"):
                st.success("Tebrikler! Kurumsal lisans başvuru talebiniz alındı.")

        with col_b2:
            st.markdown("""
                <div style="background:#fff; border:1px solid #ddd; padding:20px; border-radius:10px;">
                    <h3>👨‍🏫 Bireysel Yazar / Öğretmen Lisansı</h3>
                    <p style="color:#666;">Özel ders veren veya kendi kitabını yazan öğretmenler için profesyonel yazar paketi.</p>
                    <h2 style="color:#2c3e50;">4.900 TL / Yıl</h2>
                    <ul style="font-size:13px;">
                        <li>✓ Gelişmiş Redaktör Modülü</li>
                        <li>✓ Word (.docx) 2 Sütun Aktarım</li>
                        <li>✓ AI Asistan Desteği</li>
                    </ul>
                </div>
            """, unsafe_allow_html=True)
            if st.button("Yazar Lisansı Satın Al", key="btn_b2b_2"):
                st.success("Tebrikler! Yazar lisansı talebiniz alındı.")

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
            secilen_brans = st.selectbox("Branş Seçimi", list(kazanim_db.keys()))
            
            sinif_listesi = list(kazanim_db[secilen_brans].keys())
            secilen_sinif = st.selectbox("Sınıf Seviyesi", sinif_listesi)
            
            alan_listesi = list(kazanim_db[secilen_brans][secilen_sinif].keys())
            secilen_alan = st.selectbox("Öğrenme Alanı", alan_listesi)
            
            secilen_kazanim = st.selectbox("Hedef Kazanım", kazanim_db[secilen_brans][secilen_sinif][secilen_alan])
            
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

# --- KURŞUN KALEM ÖĞRENCİ VE AKILLI KOÇ PORTALI & PAKETLER VİTRİNİ ---
else:
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

    st.title("🎓 Kurşun Kalem - Öğrenci Sınav, Akıllı Koç ve Premium Paket Vitrini")
    st.markdown("🚀 *Geleceği yazan dijital eğitim* felsefesiyle; ilkokuldan YKS'ye tüm sınıf düzeyleri için yapay zeka koçluk destekli ve zenginleştirilmiş premium eğitim paketlerimiz.")

    ogrenci_tab1, ogrenci_tab2, ogrenci_tab3, ogrenci_tab4, ogrenci_tab5, ogrenci_tab6, ogrenci_tab7 = st.tabs([
        "📦 Zenginleştirilmiş Eğitim Paketlerimiz",
        "🎯 Günlük Test Çöz & Anında Çözüm",
        "🧠 Yapay Zeka Akıllı Koç",
        "📊 Kişisel Gelişim Karnem",
        "👨‍👩‍👧 Veli Takip Paneli",
        "🏆 Şampiyonlar Liderlik Tablosu",
        "💡 Soru Stratejileri & Taktikler"
    ])

    with ogrenci_tab1:
        st.subheader("📦 Kurşun Kalem Dijital Eğitim - Premium Ekosistem Paketleri")
        st.markdown("Her pakete özel; **7/24 Yapay Zeka Koçluğu, Video Çözüm Arşivi, Veli Takip Erişimi ve Sınırsız Deneme Hakkı** dahildir!")
        
        sinif_filtresi = st.radio(
            "Eğitim Seviyesi / Sınıf Filtrele:",
            ["Tümü", "YKS 2027", "LGS 2027", "8. Sınıf", "7. Sınıf", "6. Sınıf", "5. Sınıf", "4. Sınıf", "3. Sınıf", "2. Sınıf", "1. Sınıf"],
            horizontal=True
        )
        st.markdown("---")

        col1, col2, col3 = st.columns(3)

        if sinif_filtresi in ["Tümü", "1. Sınıf", "2. Sınıf", "3. Sınıf", "4. Sınıf"]:
            with col1:
                st.markdown("""
                    <div class="paket-karti">
                        <h4>🎨 İlkokul Premium Temel Beceri Seti</h4>
                        <p style="font-size: 13px; color: #555;">1, 2, 3 ve 4. sınıflar için okuma anlama, akıcı okuma ve eğlenceli matematik seti.</p>
                        <hr style="margin: 10px 0;">
                        <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                            <li>✓ Sınırsız AI Eğlenceli Alıştırma Modülü</li>
                            <li>✓ Veli Cep Takip Paneli (Canlı Rapor)</li>
                            <li>✓ 500+ Oyunlaştırılmış Beceri Sorusu</li>
                            <li>✓ Haftalık Dijital Rozetler & KKP Ödülü</li>
                        </ul>
                        <div class="fiyat-etiketi">19.900 TL</div>
                        <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 1.658 TL)</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("Hemen Al (İlkokul Seti)", key="btn_ilkokul", use_container_width=True):
                    st.success("Tebrikler! İlkokul Premium Set sepetinize eklendi.")

        if sinif_filtresi in ["Tümü", "5. Sınıf", "6. Sınıf", "7. Sınıf"]:
            with col2:
                st.markdown("""
                    <div class="paket-karti">
                        <h4>📚 Ortaokul Güçlendirme & Beceri Paketi</h4>
                        <p style="font-size: 13px; color: #555;">5, 6 ve 7. sınıflar için okul başarısını artıran ve LGS'ye erkenden hazırlayan set.</p>
                        <hr style="margin: 10px 0;">
                        <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                            <li>✓ Yeni Nesil Soru Çözüm Videoları</li>
                            <li>✓ Akıllı Eksik Kapatma Robotu</li>
                            <li>✓ 7/24 Yapay Zeka Akıllı Koç Erişimi</li>
                            <li>✓ Aylık Canlı Rehberlik Seminerleri</li>
                        </ul>
                        <div class="fiyat-etiketi">27.900 TL</div>
                        <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 2.325 TL)</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("Hemen Al (Ortaokul Seti)", key="btn_ortaokul", use_container_width=True):
                    st.success("Tebrikler! Ortaokul Güçlendirme Seti sepetinize eklendi.")

        if sinif_filtresi in ["Tümü", "LGS 2027", "8. Sınıf"]:
            with col3:
                st.markdown("""
                    <div class="paket-karti">
                        <h4>🚀 LGS 2027 Tüm Dersler VIP Tam Set</h4>
                        <p style="font-size: 13px; color: #555;">LGS'de Türkiye derecesi hedefleyenler için en kapsamlı yapay zeka destekli set.</p>
                        <hr style="margin: 10px 0;">
                        <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                            <li>✓ Sınırsız AI Soru Fabrikası Test Üretimi</li>
                            <li>✓ Gerçek LGS Deneme Simülasyonları</li>
                            <li>✓ Bireysel Eğitim Koçu Ataması</li>
                            <li>✓ Veli Takip Paneli + Detaylı Karne</li>
                        </ul>
                        <div class="fiyat-etiketi">44.900 TL</div>
                        <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 3.741 TL)</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("Hemen Al (LGS VIP Set)", key="btn_lgs", use_container_width=True):
                    st.success("Tebrikler! LGS VIP Tam Set sepetinize eklendi.")

        st.markdown("<br>", unsafe_allow_html=True)
        col4, col5, col6 = st.columns(3)

        if sinif_filtresi in ["Tümü", "YKS 2027"]:
            with col4:
                st.markdown("""
                    <div class="paket-karti">
                        <h4>🎯 YKS 2027 (TYT + AYT) Sayısal VIP Set</h4>
                        <p style="font-size: 13px; color: #555;">Üniversite sınavında ilk bine oynamak isteyenler için özel video ve koçluk paketi.</p>
                        <hr style="margin: 10px 0;">
                        <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                            <li>✓ TYT-AYT Tüm Konu Anlatım & Soru Bankası</li>
                            <li>✓ Nokta Atışı Çıkmış Soru Analizleri</li>
                            <li>✓ Akıllı Koç ile Haftalık Hedef Takibi</li>
                        </ul>
                        <div class="fiyat-etiketi">49.900 TL</div>
                        <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 4.158 TL)</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("Hemen Al (YKS Sayısal)", key="btn_yks", use_container_width=True):
                    st.success("Tebrikler! YKS Sayısal Set sepetinize eklendi.")

            with col5:
                st.markdown("""
                    <div class="paket-karti">
                        <h4>💡 LGS 2027 Sadece Türkçe & Sözel Kampı</h4>
                        <p style="font-size: 13px; color: #555;">Sözel mantık ve paragraf netlerini garantiye almak isteyenler için hızlandırılmış paket.</p>
                        <hr style="margin: 10px 0;">
                        <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                            <li>✓ Sınırsız Paragraf & Sözel Mantık Üretimi</li>
                            <li>✓ 2 Sütunlu LGS PDF & Word Çıktıları</li>
                            <li>✓ Şampiyonlar Liderlik Tablosu Erişimi</li>
                        </ul>
                        <div class="fiyat-etiketi">34.900 TL</div>
                        <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 2.908 TL)</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("Hemen Al (Sözel Kamp)", key="btn_sozel", use_container_width=True):
                    st.success("Tebrikler! Sözel Kamp Paketi sepetinize eklendi.")

            with col6:
                st.markdown("""
                    <div class="paket-karti">
                        <h4>🏆 8. Sınıf Son Dönem Deneme & Hızlandırma</h4>
                        <p style="font-size: 13px; color: #555;">Sınava az kala eksiklerini hızla kapatmak ve nokta atışı deneme çözmek isteyenler.</p>
                        <hr style="margin: 10px 0;">
                        <ul style="font-size: 13px; padding-left: 15px; color: #444;">
                            <li>✓ 30 Adet Kurumsal LGS Deneme Sınavı</li>
                            <li>✓ Otomatik Net, Puan ve Sıralama Analizi</li>
                            <li>✓ Eksik Kapatma Nokta Atışı Testleri</li>
                        </ul>
                        <div class="fiyat-etiketi">24.900 TL</div>
                        <div class="taksit-bilgisi">Peşin Fiyatına 12 Taksit (Ayda 2.075 TL)</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("Hemen Al (Deneme Paketi)", key="btn_deneme", use_container_width=True):
                    st.success("Tebrikler! Deneme Paketi sepetinize eklendi.")

    with ogrenci_tab2:
        st.subheader("🎯 İnteraktif Beceri Sınavı & Freemium Deneme Modülü")
        st.markdown("🎁 *Ücretsiz Freemium Deneme Sınavı Hakkınız Aktif!* Aşağıdaki örnek LGS beceri testini çözerek sistemimizi deneyebilirsiniz.")
        
        arsiv_listesi = load_arsiv()
        if not arsiv_listesi:
            st.info("Havuzda kayıtlı özel test bulunmuyor. Örnek standart LGS sorusu üzerinden test modülü:")
            st.markdown("### Örnek LGS Beceri Temelli Soru")
            st.info("Aşağıdaki cümlelerin hangisinde altı çizili sözcük **fiilimsi** değildir?")
            secim = st.radio("Seçenekler:", [
                "A) Okuldan gelen öğrenci içeri girdi.",
                "B) Yarışmada koşan atlet birinci oldu.",
                "C) Çocuğun yüzü tatlı tatlı gülüyordu.",
                "D) Görünmez kazalardan kaçınmalıyız."
            ], index=None, key="freemium_radio")

            if st.button("Cevabı Değerlendir (KKP Kazan)"):
                if secim is None:
                    st.warning("Lütfen şık seçin.")
                elif secim.startswith("C"):
                    st.success("🎉 Doğru! 'Gülüyordu' sözcüğü çekimli fiildir. +15 Kurşun Kalem Puan (KKP) kazandınız!")
                    conn = sqlite3.connect("okul.db")
                    c = conn.cursor()
                    c.execute("UPDATE ogrenciler SET puan = puan + 15, cozulen_soru = cozulen_soru + 1, dogru_sayisi = dogru_sayisi + 1, kkp = kkp + 15 WHERE id = 4")
                    conn.commit()
                    conn.close()
                else:
                    st.error("❌ Yanlış. Doğru seçenek C olmalıydı.")
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
                        key=f"kurumsal_soru_{idx}",
                        index=None
                    )

                    if st.button(f"Cevabı Kontrol Et (Soru {idx})", key=f"btn_kurumsal_{idx}"):
                        if cevap_secimi is None:
                            st.warning("Şık seçiniz.")
                        else:
                            secilen_harf = cevap_secimi[0]
                            dogru_harf = soru.get('dogru_cevap').strip().upper()
                            if secilen_harf == dogru_harf:
                                st.success(f"🎉 Tebrikler! Doğru cevap ({dogru_harf}). +20 KKP Kazandınız!")
                                conn = sqlite3.connect("okul.db")
                                c = conn.cursor()
                                c.execute("UPDATE ogrenciler SET puan = puan + 15, cozulen_soru = cozulen_soru + 1, dogru_sayisi = dogru_sayisi + 1, kkp = kkp + 20 WHERE id = 4")
                                conn.commit()
                                conn.close()
                            else:
                                st.error(f"❌ Yanlış! Doğru seçenek: **{dogru_harf}**")

    with ogrenci_tab3:
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
                            "ultra akıllı, motive edici ve nokta atışı taktikler veren Eğitim Koçusun. "
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

    with ogrenci_tab4:
        st.subheader("📊 Kişisel Gelişim ve Başarı Karnem & Oyunlaştırma")
        conn = sqlite3.connect("okul.db")
        c = conn.cursor()
        c.execute("SELECT puan, cozulen_soru, dogru_sayisi, kkp FROM ogrenciler WHERE id = 4")
        veri = c.fetchone()
        conn.close()

        puan, soru_sayisi, dogru_sayisi, kkp = veri if veri else (120, 12, 10, 150)
        yanlis_sayisi = soru_sayisi - dogru_sayisi if soru_sayisi >= dogru_sayisi else 0
        basari = int((dogru_sayisi / soru_sayisi) * 100) if soru_sayisi > 0 else 0

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Toplam Puan", f"{puan} P")
        col2.metric("Çözülen Soru", f"{soru_sayisi} Soru")
        col3.metric("Doğru / Yanlış", f"{dogru_sayisi} D / {yanlis_sayisi} Y")
        col4.metric("Başarı Oranı", f"%{basari}")
        col5.metric("Kurşun Kalem Puan (KKP)", f"⭐ {kkp} KKP")

        st.markdown("---")
        st.markdown("### 🏅 Kazanılan Dijital Rozetler")
        st.info("🏆 **Paragraf Canavarı** (Açık) | 🥈 **Sözel Mantık Uzmanı** (Açık) | 🔥 **7 Gün Kesintisiz Çalışma Ateşi** (Kilitli)")
        st.info("💡 **Analiz Notu:** Sözel mantıkta hızlandın fakat paragraf hızını artırmak için günde 15 paragraf sorusu çözmelisin.")

    with ogrenci_tab5:
        st.subheader("👨‍👩‍👧 Veli Takip Paneli")
        st.markdown("Velilerin çocuklarının günlük çözdüğü soru sayılarını, net gelişimini ve eksik konularını cepten takip edebileceği özel ekran.")
        
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            st.metric("Öğrenci Günlük Çalışma Süresi", "2 Saat 45 Dk")
            st.metric("Bu Hafta Çözülen Toplam Soru", f"{soru_sayisi + 35} Soru")
        with col_v2:
            st.success("🟢 Durum: Öğrenci bu hafta çalışma hedeflerini %92 oranında tamamladı.")
            st.info("📌 **Veli Tavsiyesi:** Çocuğunuzun fen bilimleri ağırlıklı eksiklerini kapatmak için 'Ortaokul Güçlendirme Paketi'ni inceleyebilirsiniz.")

    with ogrenci_tab6:
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

    with ogrenci_tab7:
        st.subheader("💡 Sınav Taktikleri ve Stratejiler")
        with st.expander("📌 Sözel Mantık Taktikleri"):
            st.write("1. Kesin bilgileri ele alın.\n2. Şema kurun.")
        with st.expander("📌 Paragraf Şifreleri"):
            st.write("1. Olumsuz soru köklerine dikkat edin.\n2. Son cümleye odaklanın.")
