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

st.set_page_config(page_title="Kurşun Kalem Dijital Eğitim - Profesyonel Soru & Öğrenci Portalı", page_icon="🎯", layout="wide")

def kurumsal_logo_goster():
    st.markdown("""
        <div style="display: flex; align-items: center; gap: 12px; background: linear-gradient(135deg, #2c3e50, #34495e); padding: 15px; border-radius: 12px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
            <div style="background: #e74c3c; width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 24px; color: white; font-weight: bold; border: 2px solid #f39c12;">✏️</div>
            <div>
                <h2 style="color: white; margin: 0; font-size: 20px; font-family: sans-serif;">KURŞUN KALEM DİJİTAL EĞİTİM</h2>
                <p style="color: #bdc3c7; margin: 0; font-size: 12px;">Türkiye Yüzyılı Maarif Modeli • Profesyonel Soru Fabrikası & Öğrenci Sınav Portalı</p>
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
            ("Zeynep Yılmaz", 450, 45, 41),
            ("Ahmet Demir", 380, 40, 34),
            ("Elif Çelik", 520, 55, 50),
            ("Kullanıcı Öğrenci", 100, 10, 8)
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
    ["🛠️ Soru & Redaksiyon Fabrikası (Yazar Paneli)", "🎓 Öğrenci Eğitim & Sınav Portalı"]
)
st.sidebar.markdown("---")

if ana_modul == "🛠️ Soru & Redaksiyon Fabrikası (Yazar Paneli)":
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
            secilen_brans = st.selectbox("Branş Seçimi", list(kazanim_db.keys()))
            
            sinif_listesi = list(kazanim_db[secilen_brans].keys())
            secilen_sinif = st.selectbox("Sınıf Seviyesi", sinif_listesi)
            
            alan_listesi = list(kazanim_db[secilen_brans][secilen_sinif].keys())
            secilen_alan = st.selectbox("Öğrenme Alanı", alan_listesi)
            
            secilen_kazanim = st.selectbox("Hedef Kazanım", kazanim_db[secilen_brans][secilen_sinif][secilen_alan])
            
            soru_adedi = st.slider("Üretilecek Soru Adedi (Test Modu)", min_value=1, max_value=5, value=2)
            zorluk_seviyesi = st.select_slider("Zorluk Derecesi", options=["Orta", "Orta-Zor", "İleri Düzey"], value="Orta-Zor")
            tema_girdisi = st.text_input("Bağlam / Tema / Senaryo", placeholder="Örn: Sürdürülebilirlik, deney
