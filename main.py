import io
import os
import re
import time
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types

# ==============================================================================
# SAYFA YAPILANDIRMASI VE KURUMSAL SABİTLER
# ==============================================================================
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "Kurşun Kalem Dijital Eğitim"
BRAND_TAGLINE = (
    "Geleceği yazan dijital eğitim • Profesyonel Soru Fabrikası, Dijikoç & Akıllı Öğrenci Portalı"
)
DEDICATION_TEXT = (
    "Bu program Sevgili Annem Ayşe Kurşun'a ve Sevgili Babam Sinan Gürbüz Kurşun'a "
    "ithaf edilmiştir."
)

# 2027 Sınav Hedef Tarihleri
LGS_2027_TARGET = datetime(2027, 6, 13, 9, 30, 0)
YKS_2027_TARGET = datetime(2027, 6, 19, 10, 15, 0)

PAGES = [
    "Ana Sayfa & Geri Sayım",
    "Soru & Redaksiyon Fabrikası",
    "Öğrenci Portalı & Çözüm",
    "Dijikoç",
    "Veli Takip Paneli",
    "B2B Lisans & Kurum",
    "Paketler & Abonelik Vitrini",
]

SUBSCRIPTION_TIERS: List[Dict[str, Any]] = [
    {
        "id": "ogrenci_temel",
        "segment": "Öğrenci",
        "name": "LGS Başlangıç Dijital",
        "monthly_price": 490,
        "badge": "Temel",
        "desc": "5–8. sınıf öğrencileri için günlük kazanım tekrarları ve akıllı soru çözümleri.",
        "features": [
            "Aylık 250 yapay zekâ destekli soru üretimi",
            "Öğrenci Portalı & adımlı çözüm desteği",
            "2027 LGS canlı geri sayım ve konu takibi",
            "Standart DOCX test çıktısı",
        ],
    },
    {
        "id": "ogrenci_pro",
        "segment": "Öğrenci",
        "name": "2027 LGS Hedef Pro",
        "monthly_price": 890,
        "badge": "Popüler",
        "desc": "Yeni nesil muhakeme soruları, Dijikoç birebir çalışma planı ve veli raporu.",
        "features": [
            "Aylık 1.000 yapay zekâ soru ve redaksiyon hakkı",
            "Dijikoç haftalık akıllı çalışma programı",
            "Veli Takip Paneli canlı bildirimleri",
            "2 sütunlu LGS kitapçık DOCX çıktısı",
        ],
    },
    {
        "id": "ogrenci_elit",
        "segment": "Öğrenci",
        "name": "LGS Tam Donanım Elit",
        "monthly_price": 1390,
        "badge": "Kapsamlı",
        "desc": "Yüksek soru üretim kotası, yoğun redaksiyon denetimi ve özel koçluk senaryoları.",
        "features": [
            "Aylık 2.500 soru üretimi + otomatik redaksiyon",
            "Dijikoç günlük deneme analizi & çeldirici raporu",
            "Veli Takip Paneli + haftalık özet bülteni",
            "Özel kurum/öğrenci başlıklı DOCX testler",
        ],
    },
    {
        "id": "yks_sayisal",
        "segment": "YKS",
        "name": "YKS MF (Sayısal) Akademi",
        "monthly_price": 990,
        "badge": "TYT + AYT",
        "desc": "TYT ve AYT Matematik, Fizik, Kimya, Biyoloji ve Türkçe odaklı ileri hazırlık paketi.",
        "features": [
            "TYT & AYT kazanım ve ÖSYM tarzı soru üretimi",
            "Dijikoç net artırma ve süre yönetimi koçu",
            "Konu bazlı hata defteri ve çözüm desteği",
            "2 sütunlu deneme/test DOCX çıktısı",
        ],
    },
    {
        "id": "yks_ea_sozel",
        "segment": "YKS",
        "name": "YKS EA & Sözel Atölye",
        "monthly_price": 940,
        "badge": "TYT + AYT",
        "desc": "Türk Dili ve Edebiyatı, Tarih, Coğrafya, Matematik ve paragraf analizi paketi.",
        "features": [
            "Paragraf, edebiyat ve sosyal bilimler soru fabrikası",
            "Dijikoç deneme stratejisi ve okuma hızı koçu",
            "Detaylı çeldirici ve kazanım eşleştirme",
            "Düzenlenebilir DOCX çalışma kağıtları",
        ],
    },
    {
        "id": "yks_vip",
        "segment": "YKS",
        "name": "YKS Derece VIP",
        "monthly_price": 1590,
        "badge": "VIP",
        "desc": "Yoğun deneme analizi, kişiselleştirilmiş Dijikoç ve yüksek soru üretim kotası.",
        "features": [
            "Aylık 3.000 soru üretimi ve kalite denetimi",
            "Dijikoç birebir hedef üniversite & net simülasyonu",
            "Veli/danışman görünümü ve gelişim grafikleri",
            "Kurumsal standartta DOCX kitapçık çıktısı",
        ],
    },
    {
        "id": "veli_aile",
        "segment": "Veli",
        "name": "Aile Takip & Rehberlik",
        "monthly_price": 590,
        "badge": "Veli Odaklı",
        "desc": "Öğrencinin çalışma düzenini, deneme ivmesini ve eksik kazanımlarını şeffaf izleyin.",
        "features": [
            "Veli Takip Paneli haftalık performans özeti",
            "Dijikoç çalışma disiplini ve motivasyon planı",
            "Eksik kazanım uyarıları ve öneri listesi",
            "1 öğrenci profili dahil",
        ],
    },
    {
        "id": "veli_coklu",
        "segment": "Veli",
        "name": "Çoklu Kardeş Aile Paketi",
        "monthly_price": 990,
        "badge": "2 Öğrenci",
        "desc": "İki kardeş için ayrı sınıf/sınav seviyelerinde takip, soru üretimi ve Dijikoç desteği.",
        "features": [
            "2 ayrı öğrenci profili (5–12. sınıf / LGS & YKS)",
            "Aylık 1.500 ortak soru üretim kotası",
            "Veli Takip Paneli karşılaştırmalı görünüm",
            "DOCX test indirme ve yazdırma desteği",
        ],
    },
    {
        "id": "b2b_ogretmen",
        "segment": "Kurum / B2B",
        "name": "Öğretmen & Zümre Stüdyosu",
        "monthly_price": 1790,
        "badge": "Zümre",
        "desc": "Branş öğretmenleri ve zümre başkanları için hızlı kazanım testi ve redaksiyon aracı.",
        "features": [
            "Aylık 4.000 soru üretimi + otomatik redaksiyon",
            "Türkiye Yüzyılı Maarif Modeli kazanım uyumu",
            "2 sütunlu LGS kitapçık DOCX çıktısı",
            "5 öğretmen koltuğu",
        ],
    },
    {
        "id": "b2b_kurs",
        "segment": "Kurum / B2B",
        "name": "Etüt Merkezi & Kurs Paketi",
        "monthly_price": 4490,
        "badge": "Kurumsal",
        "desc": "Kurs merkezleri ve butik okullar için ölçeklenebilir soru fabrikası ve öğrenci portalı.",
        "features": [
            "Aylık 12.000 soru üretimi ve toplu test hazırlama",
            "Öğrenci Portalı, Dijikoç ve Veli Paneli erişimi",
            "Kurum logolu/başlıklı DOCX sınav çıktıları",
            "25 öğretmen + 250 öğrenci hesabı",
        ],
    },
    {
        "id": "b2b_kampus",
        "segment": "Kurum / B2B",
        "name": "Okul & Kampüs Lisansı",
        "monthly_price": 9900,
        "badge": "Kampüs",
        "desc": "Tüm kademelerde standart ölçme-değerlendirme ve akıllı koçluk altyapısı.",
        "features": [
            "Aylık 35.000 soru üretimi + öncelikli redaksiyon",
            "Tüm modüllerde kurumsal yetkilendirme",
            "Zümre ve idareci raporlama ekranları",
            "Sınırsız test çıktısı & özel lisans desteği",
        ],
    },
    {
        "id": "b2b_ozel",
        "segment": "Kurum / B2B",
        "name": "Kurumsal Zincir & Yayınevi",
        "monthly_price": 19900,
        "badge": "Enterprise",
        "desc": "Çok şubeli eğitim kurumları ve içerik ekipleri için özel kapasite ve entegrasyon.",
        "features": [
            "Özel yüksek hacimli soru üretim kotası",
            "Özel editörlük ve çeldirici kalite standartları",
            "Kurumsal onboarding ve teknik destek",
            "Çoklu kampüs yönetimi",
        ],
    },
]

CURRICULUM_TOPICS: Dict[str, Dict[str, List[str]]] = {
    "Türkçe": {
        "Okuma ve Anlama": [
            "T.5.1.1. Bağlamdan yararlanarak bilmediği kelime gruplarının anlamını tahmin etme",
            "T.6.1.4. Metindeki ana düşünce ve yardımcı düşünceleri belirleme",
            "T.7.1.6. Metinler arası karşılaştırma yapma ve çıkarımda bulunma",
            "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
        ],
        "Dil Bilgisi ve Yazım": [
            "Fiilimsiler ve cümle ögeleri ilişkisi",
            "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            "Cümlede anlam ve anlatım bozukluklarının tespiti",
        ],
    },
    "Matematik": {
        "Sayılar ve İşlemler": [
            "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
            "Üslü ve kareköklü ifadelerde akıl yürütme ve tahmin",
            "Cebirsel ifadeler, özdeşlikler ve doğrusal denklemler",
        ],
        "Geometri ve Veri Analizi": [
            "Üçgenler, eşlik-benzerlik ve dönüşüm geometrisi",
            "Olasılık ve istatistiksel veri yorumlama",
        ],
    },
    "Fen Bilimleri": {
        "Fiziksel ve Kimyasal Süreçler": [
            "Mevsimlerin oluşumu ve iklim hareketleri",
            "DNA, genetik kod ve kalıtım problemleri",
            "Basınç, madde ve endüstri uygulamaları",
        ],
        "Canlılar ve Enerji": [
            "Basit makineler ve enerji dönüşümleri",
            "Elektrik yükleri ve elektrik enerjisi",
        ],
    },
    "Sosyal Bilgiler / İnkılap Tarihi": {
        "Tarih ve Kültür": [
            "Bir Kahraman Doğuyor ve Millî Uyanış süreci",
            "Millî Birlik ve Çağdaş Türkiye",
            "Birey, toplum ve aktif vatandaşlık",
        ],
    },
    "İngilizce": {
        "Okuma ve İletişim": [
            "Friendship, Teen Life ve In the Kitchen bağlamsal okuma",
            "On the Phone, The Internet ve Tourism akıl yürütme",
        ],
    },
}

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2.5rem;
        }
        .kp-hero {
            background: linear-gradient(135deg, #1b2a4a 0%, #2c3e50 60%, #7f1d1d 100%);
            color: #ffffff;
            padding: 1.4rem 1.7rem;
            border-radius: 14px;
            margin-bottom: 1.2rem;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.14);
        }
        .kp-hero h1 {
            margin: 0;
            font-size: 1.65rem;
            font-weight: 700;
            letter-spacing: 0.2px;
            color: #ffffff;
        }
        .kp-hero p {
            margin: 0.35rem 0 0 0;
            font-size: 0.95rem;
            color: #f1f5f9;
            opacity: 0.95;
        }
        .kp-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1.15rem 1.25rem;
            margin-bottom: 0.9rem;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        }
        .kp-metric-box {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1rem;
            text-align: center;
        }
        .kp-metric-num {
            font-size: 1.55rem;
            font-weight: 700;
            color: #1e293b;
            line-height: 1.2;
        }
        .kp-metric-lbl {
            font-size: 0.8rem;
            color: #64748b;
            margin-top: 0.25rem;
            text-transform: uppercase;
            letter-spacing: 0.4px;
        }
        .kp-dedication {
            background: #fffbeb;
            border-left: 4px solid #b45309;
            padding: 0.75rem 1rem;
            border-radius: 8px;
            color: #78350f;
            font-size: 0.9rem;
            margin-top: 0.75rem;
        }
        div[data-testid="stSidebar"] .stButton > button {
            width: 100%;
            text-align: left;
            justify-content: flex-start;
            border-radius: 8px;
            padding: 0.5rem 0.75rem;
            font-weight: 500;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# YARDIMCI FONKSİYONLAR: API ANAHTARI, GEMINI İSTEMCİSİ VE SAYAÇ
# ==============================================================================
def resolve_gemini_api_key() -> Optional[str]:
    """Yalnızca st.secrets veya ortam değişkenlerinden API anahtarını okur."""
    key: Optional[str] = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            key = str(st.secrets["GEMINI_API_KEY"]).strip()
    except Exception:
        key = None

    if not key:
        env_val = os.environ.get("GEMINI_API_KEY", "").strip()
        if env_val:
            key = env_val

    return key if key else None


def get_gemini_client() -> genai.Client:
    """API anahtarını doğrulayarak google.genai Client nesnesi döndürür."""
    api_key = resolve_gemini_api_key()
    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY bulunamadı. Lütfen Streamlit Secrets veya ortam "
            "değişkenlerinde GEMINI_API_KEY tanımlayın."
        )
    return genai.Client(api_key=api_key)


def call_gemini_text(prompt: str, system_instruction: str, model: str = "gemini-2.5-flash") -> str:
    """Gemini API üzerinden metin üretir."""
    client = get_gemini_client()
    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=system_instruction,
            temperature=0.7,
        ),
    )
    text = getattr(response, "text", None)
    if not text:
        raise RuntimeError("Modelden boş yanıt döndü.")
    return text.strip()


def format_countdown(target: datetime) -> Tuple[int, int, int, int]:
    """Hedef tarihe kalan gün, saat, dakika ve saniyeyi döndürür."""
    now = datetime.now()
    delta = target - now
    if delta.total_seconds() <= 0:
        return 0, 0, 0, 0
    total_seconds = int(delta.total_seconds())
    days = total_seconds // 86400
    hours = (total_seconds % 86400) // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    return days, hours, minutes, seconds


def render_countdown_boxes(target: datetime) -> None:
    """Gün, Saat, Dakika ve Saniye sayaç kutularını çizer."""
    d, h, m, s = format_countdown(target)
    c1, c2, c3, c4 = st.columns(4)
    for col, val, lbl in [
        (c1, f"{d:03d}", "GÜN"),
        (c2, f"{h:02d}", "SAAT"),
        (c3, f"{m:02d}", "DAKİKA"),
        (c4, f"{s:02d}", "SANİYE"),
    ]:
        col.markdown(
            f"""
            <div class="kp-metric-box">
                <div class="kp-metric-num">{val}</div>
                <div class="kp-metric-lbl">{lbl}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def build_test_docx(
    institution: str,
    branch: str,
    grade: str,
    learning_area: str,
    outcome: str,
    context_theme: str,
    generated_content: str,
) -> bytes:
    """Üretilen soruları kurumsal başlıklı DOCX dosyasına dönüştürür."""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"{institution.upper()} — DEĞERLENDİRME TESTİ")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.color.rgb = RGBColor(27, 42, 74)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(
        f"Branş: {branch}  |  Sınıf: {grade}  |  Öğrenme Alanı: {learning_area}"
    )
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run(
        f"Hedef Kazanım: {outcome}\nBağlam / Tema: {context_theme}\n"
        f"Tarih: {datetime.now().strftime('%d.%m.%Y')}"
    )
    meta_run.font.size = Pt(9.5)
    meta_run.font.color.rgb = RGBColor(51, 65, 85)

    doc.add_paragraph("─" * 78)

    for line in generated_content.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(cleaned)
        run.font.size = Pt(10.5)
        if re.match(r"^(Soru\s*\d+|Cevap\s*Anahtarı|Çözüm)", cleaned, re.IGNORECASE):
            run.bold = True
            run.font.color.rgb = RGBColor(127, 29, 29)

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run(f"\n{BRAND_NAME} • {DEDICATION_TEXT}")
    footer_run.italic = True
    footer_run.font.size = Pt(8.5)
    footer_run.font.color.rgb = RGBColor(100, 116, 139)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def render_iyzico_simulator(package_name: str, monthly_price: int) -> None:
    """iyzico taksit ve ödeme simülatörü."""
    st.markdown("### 💳 iyzico Taksit & Ödeme Simülatörü")
    c1, c2, c3 = st.columns(3)
    with c1:
        billing_cycle = st.selectbox(
            "Dönem Seçimi",
            ["Aylık Abonelik", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"],
            key=f"cycle_{package_name}",
        )
    with c2:
        installment = st.selectbox(
            "Taksit Sayısı (iyzico)",
            [1, 2, 3, 6, 9, 12],
            index=2,
            key=f"inst_{package_name}",
        )
    with c3:
        card_brand = st.selectbox(
            "Kart Programı",
            ["Bonus", "Maximum", "World", "Axess", "Paraf", "Diğer Banka Kartı"],
            key=f"card_{package_name}",
        )

    if "3 Aylık" in billing_cycle:
        total_amount = round(monthly_price * 3 * 0.90)
    elif "Yıllık" in billing_cycle:
        total_amount = round(monthly_price * 12 * 0.80)
    else:
        total_amount = monthly_price

    monthly_installment = total_amount / installment

    m1, m2, m3 = st.columns(3)
    m1.metric("Toplam Tutar", f"{total_amount:,} ₺".replace(",", "."))
    m2.metric(f"{installment} Taksit Aylık", f"{monthly_installment:,.2f} ₺".replace(",", " "))
    m3.metric("Ödeme Altyapısı", f"iyzico • {card_brand}")

    if st.button("iyzico Ödeme Talebi Oluştur (Simülasyon)", key=f"pay_btn_{package_name}"):
        st.success(
            f"**{package_name}** paketi için **{total_amount:,} ₺** ({installment} taksit, "
            f"{card_brand}) tutarında iyzico ödeme oturumu hazırlandı.".replace(",", ".")
        )


# ==============================================================================
# SESSION STATE BAŞLATMA
# ==============================================================================
if "active_page" not in st.session_state:
    st.session_state.active_page = PAGES[0]
if "generated_test_text" not in st.session_state:
    st.session_state.generated_test_text = ""
if "dijikoc_plan" not in st.session_state:
    st.session_state.dijikoc_plan = ""
if "selected_tier_id" not in st.session_state:
    st.session_state.selected_tier_id = SUBSCRIPTION_TIERS[1]["id"]
if "live_tick" not in st.session_state:
    st.session_state.live_tick = False

api_key_ready = resolve_gemini_api_key() is not None

# ==============================================================================
# KENAR ÇUBUĞU (CUSTOM BUTTON NAVİGASYON)
# ==============================================================================
with st.sidebar:
    st.markdown(
        f"""
        <div style="background:#1e293b;color:#ffffff;padding:0.85rem 1rem;border-radius:10px;margin-bottom:0.85rem;">
            <div style="font-weight:700;font-size:1rem;">🧭 Sistem Navigasyonu</div>
            <div style="font-size:0.78rem;color:#cbd5e1;margin-top:0.15rem;">{BRAND_NAME}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    page_icons = {
        "Ana Sayfa & Geri Sayım": "🏠",
        "Soru & Redaksiyon Fabrikası": "⚙️",
        "Öğrenci Portalı & Çözüm": "🎓",
        "Dijikoç": "🤖",
        "Veli Takip Paneli": "👨‍👩‍👧",
        "B2B Lisans & Kurum": "🏫",
        "Paketler & Abonelik Vitrini": "💎",
    }

    for p_name in PAGES:
        is_current = st.session_state.active_page == p_name
        btn_type = "primary" if is_current else "secondary"
        label = f"{page_icons.get(p_name, '•')} {p_name}"
        if st.button(label, key=f"nav_{p_name}", type=btn_type):
            st.session_state.active_page = p_name
            st.rerun()

    st.markdown("---")
    if api_key_ready:
        st.success("● Sistem Aktif • 2027 LGS & Maarif Modeli")
    else:
        st.warning("● API Anahtarı Bekleniyor (Secrets)")

    st.caption(DEDICATION_TEXT)

# ==============================================================================
# ÜST KURUMSAL BANNER
# ==============================================================================
st.markdown(
    f"""
    <div class="kp-hero">
        <h1>✏️ {BRAND_NAME.upper()}</h1>
        <p>{BRAND_TAGLINE}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

current_page = st.session_state.active_page

# ==============================================================================
# 1. ANA SAYFA & CANLI GERİ SAYIM
# ==============================================================================
if current_page == "Ana Sayfa & Geri Sayım":
    st.subheader("🚀 Kurşun Kalem Dijital Eğitim Platformuna Hoş Geldiniz")
    st.write(
        "İlkokuldan Liseye ve YKS'ye uzanan, yapay zekâ destekli Türkiye'nin akıllı "
        "soru fabrikası, Dijikoç ve ölçme-değerlendirme ekosistemi."
    )

    head_col, toggle_col = st.columns([3, 1])
    with head_col:
        st.markdown("#### 🏆 2027 LGS'ye Geri Sayım (13 Haziran 2027)")
    with toggle_col:
        live_counter = st.toggle("⏱️ Canlı Sayaç Akışı", value=st.session_state.live_tick)
        st.session_state.live_tick = live_counter

    counter_placeholder = st.empty()
    with counter_placeholder.container():
        render_countdown_boxes(LGS_2027_TARGET)

    yks_d, yks_h, _, _ = format_countdown(YKS_2027_TARGET)
    st.caption(f"2027 YKS (19 Haziran 2027) için kalan süre: **{yks_d} gün {yks_h} saat**")

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Üretilen Soru Havuzu", "142.500+", "+1.200 bugün")
    m2.metric("Aktif Öğrenci / Veli", "18.420", "+340 bu hafta")
    m3.metric("B2B Anlaşmalı Kurum", "48", "+4 Yeni")
    m4.metric("Çözülen Soru Sayısı", "2.4 Milyon", "%98 Başarı")

    st.markdown("### 💡 Platform Modülleri")
    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown(
            """
            <div class="kp-card">
                <h4>⚙️ Soru & Redaksiyon Fabrikası</h4>
                <p>Türkiye Yüzyılı Maarif Modeli kazanımlarına uygun, çeldirici analizi yapılmış
                özgün sorular üretin ve DOCX formatında indirin.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with f2:
        st.markdown(
            """
            <div class="kp-card">
                <h4>🤖 Dijikoç Akıllı Rehberlik</h4>
                <p>Öğrencinin hedef sınavına, zayıf kazanımlarına ve haftalık temposuna göre
                kişiselleştirilmiş çalışma ve deneme stratejisi sunar.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with f3:
        st.markdown(
            """
            <div class="kp-card">
                <h4>🏫 Veli & Kurum Ekosistemi</h4>
                <p>Veli takip raporları, zümre ve kurum lisansları ile 12 kademeli abonelik
                yapısını tek merkezden yönetin.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(f'<div class="kp-dedication">{DEDICATION_TEXT}</div>', unsafe_allow_html=True)

    if st.session_state.live_tick:
        time.sleep(1)
        st.rerun()

# ==============================================================================
# 2. SORU & REDAKSİYON FABRİKASI
# ==============================================================================
elif current_page == "Soru & Redaksiyon Fabrikası":
    st.subheader("⚙️ Kurşun Kalem Dijital Eğitim - Soru Üretim Stüdyosu")
    st.caption(
        "Türkiye Yüzyılı Maarif Modeli standartlarında, Otomatik Redaktör ve Kalite Kontrollü "
        "değerlendirme testleri oluşturun."
    )

    with st.expander("📋 Test ve Kurum Parametreleri", expanded=True):
        r1c1, r1c2, r1c3 = st.columns(3)
        with r1c1:
            institution = st.text_input("Kurum / Okul Adı", value=BRAND_NAME)
            branch = st.selectbox("Branş Seçimi", list(CURRICULUM_TOPICS.keys()), index=0)
            grade = st.selectbox(
                "Sınıf Seviyesi",
                ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS)", "9-10. Sınıf", "11-12. Sınıf (YKS)"],
                index=0,
            )
        with r1c2:
            areas = list(CURRICULUM_TOPICS[branch].keys())
            learning_area = st.selectbox("Öğrenme Alanı", areas, index=0)
            outcomes = CURRICULUM_TOPICS[branch][learning_area]
            outcome = st.selectbox("Hedef Kazanım", outcomes, index=0)
            context_theme = st.text_input("Bağlam / Tema / Senaryo", value="HAYAT")
        with r1c3:
            question_count = st.slider("Üretilecek Soru Adedi (Test Modu)", 1, 10, 3)
            difficulty = st.select_slider(
                "Zorluk Derecesi",
                options=["Kolay", "Orta", "Orta-Zor", "Zor (LGS/ÖSYM)"],
                value="Orta-Zor",
            )
            auto_review = st.checkbox(
                "Otomatik Soru Kalite & Redaksiyon Denetimi",
                value=True,
            )

    if st.button("✨ Kurumsal Soru Paketi Üret", type="primary"):
        sys_prompt = (
            "Sen Türkiye Yüzyılı Maarif Modeli, LGS ve YKS ölçme-değerlendirme standartlarında "
            "uzman bir Türkçe ve test yazarısın. Soruların özgün, açık, çeldiricileri güçlü, "
            "dil bilgisi ve yazım kurallarına kusursuz uygun olmasını sağla. "
            "Otomatik redaksiyon açıksa her sorunun sonunda kısa bir [Redaksiyon & Çeldirici Notu] ekle "
            "ve en sonda ayrı bir 'Cevap Anahtarı ve Kısa Gerekçeler' bölümü oluştur."
        )
        user_prompt = (
            f"Kurum: {institution}\n"
            f"Branş: {branch}\n"
            f"Sınıf Seviyesi: {grade}\n"
            f"Öğrenme Alanı: {learning_area}\n"
            f"Hedef Kazanım: {outcome}\n"
            f"Bağlam / Tema / Senaryo: {context_theme}\n"
            f"Zorluk Derecesi: {difficulty}\n"
            f"Soru Adedi: {question_count}\n"
            f"Otomatik Redaksiyon ve Kalite Denetimi: {'Aktif' if auto_review else 'Pasif'}\n\n"
            "Lütfen numaralandırılmış çoktan seçmeli (A, B, C, D) özgün sorular üret."
        )
        try:
            with st.spinner("Sorular üretiliyor ve kalite denetimi yapılıyor..."):
                result = call_gemini_text(prompt=user_prompt, system_instruction=sys_prompt)
            st.session_state.generated_test_text = result
            st.success("Soru paketi başarıyla oluşturuldu.")
        except Exception as exc:
            st.error(f"Soru üretim hatası oluştu: {exc}")

    if st.session_state.generated_test_text:
        st.markdown("---")
        st.markdown("### 📄 Üretilen Test ve Redaksiyon Çıktısı")
        st.markdown(st.session_state.generated_test_text)

        docx_bytes = build_test_docx(
            institution=institution,
            branch=branch,
            grade=grade,
            learning_area=learning_area,
            outcome=outcome,
            context_theme=context_theme,
            generated_content=st.session_state.generated_test_text,
        )
        st.download_button(
            label="📥 Testi DOCX Olarak İndir",
            data=docx_bytes,
            file_name=f"kursun_kalem_{branch.lower().replace(' ', '_')}_test.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# ==============================================================================
# 3. ÖĞRENCİ PORTALI & ÇÖZÜM
# ==============================================================================
elif current_page == "Öğrenci Portalı & Çözüm":
    st.subheader("🎓 Öğrenci Portalı & Adımlı Çözüm Asistanı")
    st.write("Takıldığın soruyu veya konuyu gir; adım adım yönlendirme ve çözüm stratejisi al.")

    col_l, col_r = st.columns([2, 1])
    with col_l:
        student_question = st.text_area(
            "Çözümünü İncelemek İstediğin Soru Metni",
            height=160,
            placeholder="Soru kökünü, verilen öncülleri ve şıkları buraya yapıştır...",
        )
    with col_r:
        hint_mode = st.radio(
            "Yardım Düzeyi",
            ["Önce İpucu Ver (Kendim Çözeyim)", "Adım Adım Çözümü Göster"],
            index=0,
        )
        student_attempt = st.text_input("Senin Tahminin / Şıkkın", placeholder="Örn: C şıkkı, çünkü...")

    if st.button("🔍 Çözüm Rehberi Oluştur", type="primary"):
        if not student_question.strip():
            st.warning("Lütfen önce bir soru metni girin.")
        else:
            sys_inst = (
                "Sen sabırlı, pedagojik ve teşvik edici bir eğitim koçusun. "
                "Öğrenciye doğrudan ezberletmek yerine mantığını kavratacak şekilde açıkla."
            )
            prompt = (
                f"Soru:\n{student_question}\n\n"
                f"Öğrencinin Denemesi: {student_attempt or 'Belirtilmedi'}\n"
                f"Yardım Düzeyi: {hint_mode}\n\n"
                "Lütfen Türkçe olarak net ve anlaşılır yanıt ver."
            )
            try:
                with st.spinner("Çözüm rehberi hazırlanıyor..."):
                    ans = call_gemini_text(prompt=prompt, system_instruction=sys_inst)
                st.markdown("#### 💡 Çözüm Analizi")
                st.markdown(ans)
            except Exception as exc:
                st.error(f"Çözüm üretilirken hata oluştu: {exc}")

# ==============================================================================
# 4. DİJİKOÇ (AYRI SAYFA VE KOÇLUK MODÜLÜ)
# ==============================================================================
elif current_page == "Dijikoç":
    st.subheader("🤖 Dijikoç — Kişisel Akıllı Sınav ve Çalışma Koçu")
    st.caption(
        "Hedef sınavına, haftalık müsait sürene ve zayıf konularına göre özel çalışma planı, "
        "deneme analizi ve günlük rutin oluşturur."
    )

    k1, k2, k3 = st.columns(3)
    with k1:
        target_exam = st.selectbox(
            "Hedef Sınav / Kademe",
            ["2027 LGS", "2027 TYT - AYT", "5-7. Sınıf Akademik Gelişim"],
            index=0,
        )
        weekly_hours = st.slider("Haftalık Çalışma Süresi (Saat)", 4, 45, 16)
    with k2:
        weak_topics = st.text_area(
            "Zorlandığın Dersler / Konular",
            value="Paragrafta ana düşünce, üslü ifadeler, dikkat kaynaklı işlem hataları",
            height=105,
        )
    with k3:
        coach_focus = st.selectbox(
            "Koçluk Odak Alanı",
            [
                "Haftalık Dengeli Çalışma Programı",
                "Deneme Net Artırma & Zaman Yönetimi",
                "Dikkat ve Çeldirici Tuzağı Azaltma",
                "Sınav Kaygısı & Motivasyon Rutini",
            ],
            index=0,
        )
        recent_score = st.text_input(
            "Son Deneme / Sınav Durumu",
            value="Orta seviye, süreyi yetiştirmede zorlanıyorum",
        )

    if st.button("🎯 Dijikoç Planı Oluştur", type="primary"):
        sys_coach = (
            "Sen 'Dijikoç' adlı, Türkiye'deki öğrenciler için çalışan disiplinli, empatik ve "
            "veri odaklı bir dijital eğitim koçusun. Uygulanabilir, günlere bölünmüş, "
            "mola ve tekrar döngüleri içeren net bir çalışma planı hazırla."
        )
        prompt_coach = (
            f"Hedef: {target_exam}\n"
            f"Haftalık Süre: {weekly_hours} saat\n"
            f"Zorlanılan Konular: {weak_topics}\n"
            f"Koçluk Odağı: {coach_focus}\n"
            f"Mevcut Durum: {recent_score}\n\n"
            "Lütfen haftalık planı, kritik uyarıları ve haftalık kontrol hedeflerini Türkçe yaz."
        )
        try:
            with st.spinner("Dijikoç çalışma planını hazırlıyor..."):
                plan = call_gemini_text(prompt=prompt_coach, system_instruction=sys_coach)
            st.session_state.dijikoc_plan = plan
            st.success("Dijikoç planı hazırlandı.")
        except Exception as exc:
            st.error(f"Dijikoç planı oluşturulurken hata oluştu: {exc}")

    if st.session_state.dijikoc_plan:
        st.markdown("---")
        st.markdown("### 📅 Kişiselleştirilmiş Dijikoç Planı")
        st.markdown(st.session_state.dijikoc_plan)

# ==============================================================================
# 5. VELİ TAKİP PANELİ
# ==============================================================================
elif current_page == "Veli Takip Paneli":
    st.subheader("👨‍👩‍👧 Veli Takip Paneli")
    st.write("Öğrencinin haftalık çalışma düzeni, soru çözümü ve kazanım ilerleme özeti.")

    v1, v2, v3 = st.columns(3)
    v1.metric("Bu Hafta Çözülen Soru", "340 Soru", "+18% geçen haftaya göre")
    v2.metric("Ortalama Doğruluk", "%81", "+4 puan")
    v3.metric("Dijikoç Plan Uyumu", "%88", "6 / 7 gün aktif")

    st.markdown("#### Kazanım Bazlı Görünüm")
    st.progress(0.84, text="Okuma ve Anlama — %84")
    st.progress(0.72, text="Sayısal Muhakeme ve Problem Çözme — %72")
    st.progress(0.90, text="Düzenli Tekrar ve Ödev Tamamlama — %90")

# ==============================================================================
# 6. B2B LİSANS & KURUM
# ==============================================================================
elif current_page == "B2B Lisans & Kurum":
    st.subheader("🏫 B2B Lisans & Kurum Yönetimi")
    st.write(
        "Okullar, kurs merkezleri ve eğitim kurumları için toplu soru üretimi, zümre "
        "kalite denetimi ve kurumsal lisanslama çözümleri."
    )

    with st.form("b2b_inquiry_form"):
        c1, c2 = st.columns(2)
        with c1:
            org_name = st.text_input("Kurum / Okul Adı")
            city = st.text_input("İl / İlçe", value="İstanbul")
            teacher_seats = st.number_input(
                "Öğretmen / Editör Koltuk Sayısı", min_value=1, max_value=500, value=15
            )
        with c2:
            student_seats = st.number_input(
                "Öğrenci Hesap Sayısı", min_value=10, max_value=10000, value=250
            )
            contact_name = st.text_input("Yetkili Adı Soyadı")
            contact_email = st.text_input("Kurumsal E-posta")

        submitted = st.form_submit_button("Kurumsal Lisans Teklifi Hazırla")
        if submitted:
            if not org_name.strip():
                st.warning("Lütfen kurum adını girin.")
            else:
                st.success(
                    f"**{org_name}** ({city}) için **{teacher_seats} öğretmen** ve "
                    f"**{student_seats} öğrenci** kapasiteli kurumsal lisans talebi kaydedildi."
                )

# ==============================================================================
# 7. PAKETLER & ABONELİK VİTRİNİ (12 PAKET, KARŞILAŞTIRMA MATRİSİ & IYZICO)
# ==============================================================================
elif current_page == "Paketler & Abonelik Vitrini":
    st.subheader("💎 Paketler & Abonelik Vitrini (12 Paket)")
    st.write("İhtiyacınıza uygun öğrenci, YKS, veli veya kurumsal paketi seçin ve iyzico taksitlerini inceleyin.")

    selected_segment = st.segmented_control(
        "Segment Filtresi",
        options=["Tümü", "Öğrenci", "YKS", "Veli", "Kurum / B2B"],
        default="Tümü",
    )

    filtered_tiers = (
        SUBSCRIPTION_TIERS
        if not selected_segment or selected_segment == "Tümü"
        else [t for t in SUBSCRIPTION_TIERS if t["segment"] == selected_segment]
    )

    for row_start in range(0, len(filtered_tiers), 3):
        cols = st.columns(3)
        for idx, tier in enumerate(filtered_tiers[row_start : row_start + 3]):
            with cols[idx]:
                st.markdown(
                    f"""
                    <div class="kp-card">
                        <div style="font-size:0.78rem;color:#64748b;font-weight:600;">
                            {tier["segment"].upper()} • {tier["badge"]}
                        </div>
                        <h4 style="margin:0.25rem 0;">{tier["name"]}</h4>
                        <div style="font-size:1.35rem;font-weight:700;color:#7f1d1d;margin-bottom:0.4rem;">
                            {tier["monthly_price"]:,} ₺ <span style="font-size:0.8rem;font-weight:400;color:#64748b;">/ ay</span>
                        </div>
                        <p style="font-size:0.86rem;color:#475569;">{tier["desc"]}</p>
                    </div>
                    """.replace(",", "."),
                    unsafe_allow_html=True,
                )
                for feat in tier["features"]:
                    st.markdown(f"- {feat}")
                if st.button("Paketi Seç", key=f"select_{tier['id']}"):
                    st.session_state.selected_tier_id = tier["id"]
                    st.toast(f"{tier['name']} seçildi.")

    st.markdown("---")
    st.markdown("### 📊 12 Paket Karşılaştırma Matrisi")
    matrix_rows = [
        {
            "Paket Adı": t["name"],
            "Segment": t["segment"],
            "Etiket": t["badge"],
            "Aylık Ücret (₺)": f"{t['monthly_price']:,}".replace(",", "."),
            "Öne Çıkan Kapsam": t["features"][0],
        }
        for t in SUBSCRIPTION_TIERS
    ]
    st.dataframe(matrix_rows, use_container_width=True, hide_index=True)

    st.markdown("---")
    chosen_tier = next(
        (t for t in SUBSCRIPTION_TIERS if t["id"] == st.session_state.selected_tier_id),
        SUBSCRIPTION_TIERS[1],
    )
    chosen_name = st.selectbox(
        "Simülasyon İçin Paket Seçin",
        options=[t["name"] for t in SUBSCRIPTION_TIERS],
        index=[t["name"] for t in SUBSCRIPTION_TIERS].index(chosen_tier["name"]),
    )
    active_tier = next(t for t in SUBSCRIPTION_TIERS if t["name"] == chosen_name)
    render_iyzico_simulator(active_tier["name"], active_tier["monthly_price"])
