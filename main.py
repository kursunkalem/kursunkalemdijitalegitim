import base64
import html
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
    page_title="KURŞUN KALEM DİJİTAL EĞİTİM",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "KURŞUN KALEM DİJİTAL EĞİTİM"
BRAND_TAGLINE = "KURŞUN KALEM DİJİTAL EĞİTİM- Geleceği Yazan Dijital Eğitim"
DEFAULT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-2.5-flash"

# Kalıcı Gömülü Kurumsal Logo (SVG -> Base64 / PNG uyumlu byte akışı)
EMBEDDED_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512">
  <defs>
    <linearGradient id="goldGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#f6d365"/>
      <stop offset="50%" stop-color="#d4af37"/>
      <stop offset="100%" stop-color="#9a6b14"/>
    </linearGradient>
    <linearGradient id="navyGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#172c56"/>
      <stop offset="100%" stop-color="#091328"/>
    </linearGradient>
    <linearGradient id="pencilBody" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#ffd23f"/>
      <stop offset="100%" stop-color="#e09f1f"/>
    </linearGradient>
  </defs>
  <circle cx="256" cy="256" r="244" fill="#ffffff"/>
  <path d="M 432 108 A 232 232 0 1 0 468 332 L 444 326 A 208 208 0 1 1 412 124 Z" fill="url(#navyGrad)"/>
  <path d="M 456 138 A 232 232 0 0 1 474 295 L 452 292 A 210 210 0 0 0 436 152 Z" fill="url(#goldGrad)"/>
  <path d="M 102 222 Q 178 214 256 254 Q 334 214 410 222 L 404 236 Q 332 228 256 266 Q 180 228 108 236 Z" fill="url(#navyGrad)"/>
  <path d="M 118 202 Q 184 196 256 238 Q 328 196 394 202 L 390 214 Q 326 208 256 248 Q 186 208 122 214 Z" fill="url(#goldGrad)"/>
  <polygon points="256,82 362,130 256,178 150,130" fill="url(#navyGrad)" stroke="url(#goldGrad)" stroke-width="5"/>
  <polygon points="190,148 256,178 322,148 322,188 256,216 190,188" fill="url(#navyGrad)"/>
  <path d="M 332 144 L 332 206" stroke="url(#goldGrad)" stroke-width="6" stroke-linecap="round"/>
  <circle cx="332" cy="214" r="9" fill="url(#goldGrad)"/>
  <text x="256" y="314" text-anchor="middle" font-family="Plus Jakarta Sans, Arial, sans-serif" font-size="48" font-weight="800" fill="#0b1836" letter-spacing="1">DİJİTAL OKUL</text>
  <line x1="92" y1="344" x2="152" y2="344" stroke="#64748b" stroke-width="2.5"/>
  <text x="256" y="351" text-anchor="middle" font-family="Plus Jakarta Sans, Arial, sans-serif" font-size="22" font-weight="700" fill="#0f172a" letter-spacing="5">KURŞUN KALEM</text>
  <line x1="360" y1="344" x2="420" y2="344" stroke="#64748b" stroke-width="2.5"/>
  <path d="M 122 396 Q 272 414 432 356 Q 282 404 122 390 Z" fill="url(#goldGrad)"/>
  <g transform="translate(432, 342) rotate(24)">
    <rect x="-12" y="-110" width="24" height="96" rx="3" fill="url(#pencilBody)" stroke="#1e293b" stroke-width="2"/>
    <rect x="-12" y="-124" width="24" height="14" fill="#cbd5e1" stroke="#1e293b" stroke-width="2"/>
    <rect x="-12" y="-138" width="24" height="14" rx="4" fill="#f472b6" stroke="#1e293b" stroke-width="2"/>
    <polygon points="-12,-14 12,-14 0,24" fill="#fde68a" stroke="#1e293b" stroke-width="2"/>
    <polygon points="-3.5,12 3.5,12 0,24" fill="#0f172a"/>
  </g>
</svg>"""

# Alışveriş Sepeti Mantığında Çalışan Okul Sırt Çantası Görseli (SVG)
SCHOOL_BACKPACK_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 300" width="260" height="300">
  <defs>
    <linearGradient id="bagBody" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#1e3a8a"/>
      <stop offset="60%" stop-color="#0f172a"/>
      <stop offset="100%" stop-color="#7f1d1d"/>
    </linearGradient>
    <linearGradient id="bagPocket" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#f59e0b"/>
      <stop offset="100%" stop-color="#b45309"/>
    </linearGradient>
    <linearGradient id="strapGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#334155"/>
      <stop offset="100%" stop-color="#0f172a"/>
    </linearGradient>
  </defs>
  <path d="M 98 56 Q 130 20 162 56" fill="none" stroke="#1e293b" stroke-width="10" stroke-linecap="round"/>
  <path d="M 88 64 C 58 95, 54 165, 66 245" fill="none" stroke="url(#strapGrad)" stroke-width="16" stroke-linecap="round"/>
  <path d="M 172 64 C 202 95, 206 165, 194 245" fill="none" stroke="url(#strapGrad)" stroke-width="16" stroke-linecap="round"/>
  <rect x="64" y="56" width="132" height="205" rx="34" fill="url(#bagBody)" stroke="#0f172a" stroke-width="4"/>
  <path d="M 68 92 Q 130 116 192 92 L 192 78 Q 130 56 68 78 Z" fill="#334155" opacity="0.55"/>
  <rect x="80" y="138" width="100" height="98" rx="22" fill="url(#bagPocket)" stroke="#fef3c7" stroke-width="2.5"/>
  <path d="M 88 152 Q 130 166 172 152" fill="none" stroke="#0f172a" stroke-width="3" stroke-dasharray="6,4"/>
  <circle cx="130" cy="186" r="22" fill="#ffffff" stroke="#0f172a" stroke-width="2.5"/>
  <text x="130" y="184" text-anchor="middle" font-family="Plus Jakarta Sans, Arial, sans-serif" font-size="9" font-weight="800" fill="#0f172a">ÇANTA</text>
  <text x="130" y="196" text-anchor="middle" font-family="Plus Jakarta Sans, Arial, sans-serif" font-size="8" font-weight="700" fill="#881337">PAKETLER</text>
  <rect x="72" y="242" width="116" height="20" rx="10" fill="#090d16"/>
</svg>"""


def svg_to_data_uri(svg_str: str) -> str:
    """SVG metnini güvenli Data URI formatına çevirir."""
    encoded = base64.b64encode(svg_str.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def get_embedded_logo_png_bytes() -> Optional[bytes]:
    """Gömülü logoyu DOCX belgesine eklemek için PNG baytı üretir (cairosvg veya Pillow fallback)."""
    try:
        import cairosvg  # type: ignore

        png = cairosvg.svg2png(bytestring=EMBEDDED_LOGO_SVG.encode("utf-8"), scale=2.0)
        if png:
            return png
    except Exception:
        pass

    try:
        from PIL import Image, ImageDraw  # type: ignore

        size = 320
        img = Image.new("RGBA", (size, size), (255, 255, 255, 255))
        draw = ImageDraw.Draw(img)

        draw.ellipse([14, 14, size - 14, size - 14], outline=(15, 23, 42, 255), width=10)
        draw.arc([20, 20, size - 20, size - 20], start=290, end=75, fill=(212, 175, 55, 255), width=8)

        draw.pieslice([68, 95, 252, 215], start=180, end=360, fill=(23, 44, 86, 255))
        draw.polygon(
            [(160, 56), (232, 88), (160, 120), (88, 88)],
            fill=(15, 23, 42, 255),
            outline=(212, 175, 55, 255),
        )
        draw.rectangle([122, 104, 198, 130], fill=(23, 44, 86, 255))

        draw.text((160, 182), "DİJİTAL OKUL", fill=(11, 24, 54, 255), anchor="mm")
        draw.text((160, 206), "— KURŞUN KALEM —", fill=(127, 29, 29, 255), anchor="mm")

        draw.line([(215, 252), (274, 206)], fill=(234, 179, 8, 255), width=12)
        draw.polygon([(274, 206), (284, 198), (281, 212)], fill=(15, 23, 42, 255))

        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return None


LGS_2027_TARGET = datetime(2027, 6, 13, 9, 30, 0)
YKS_2027_TARGET = datetime(2027, 6, 19, 10, 15, 0)

PAGES = [
    "Ana Sayfa & Geri Sayım",
    "Soru & Redaksiyon Fabrikası",
    "Öğrenci Portalı & Çözüm",
    "Dijimentor",
    "Veli Takip Paneli",
    "B2B Lisans & Kurum",
    "Paketler & Abonelik Vitrini",
]

DIJIMENTOR_PRESETS: Dict[str, Dict[str, Any]] = {
    "2027 LGS Derece & Net İvmesi": {
        "exam": "2027 LGS",
        "hours": 20,
        "goal": "Fen liseleri ve nitelikli Anadolu liseleri bandı (470+ puan ve 82+ net hedefi)",
        "focus": "Deneme Net Artırma & Zaman Yönetimi",
        "weak": "Sözel mantık ve muhakeme metinleri, cebirsel ifadeler ve çarpanlar-katlar yeni nesil problemleri, süre baskısı altında çeldirici ayıklama",
        "status": "Orta-üst seviye; Türkçe ve Matematikte son 15 soruda süre baskısı yaşıyor ve deneme analizini haftalık sistematik yapamıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "Sözel Mantık & Muhakeme ve Çeldirici Analizi",
    },
    "2027 YKS TYT-AYT Disiplin Rutini": {
        "exam": "2027 TYT - AYT",
        "hours": 30,
        "goal": "TYT 98+ net ve AYT alanında ilk 15 bin derece hedefi",
        "focus": "Haftalık Dengeli Çalışma Programı",
        "weak": "TYT Türkçe paragraf süresi ve anlam-yorum çeldiricileri, problem kurma hızı, AYT konu tekrar gecikmeleri ve deneme tempo dalgalanması",
        "status": "Düzenli çalışma alışkanlığı var ancak TYT süresini verimli bölme ve hafta sonu deneme analizlerini puan/net aksiyonuna çevirme konusunda gelişime açık",
        "branch": "Türkçe",
        "grade": "11-12. Sınıf (YKS)",
        "topic": "TYT Paragraf Hızı, Çeldirici Eleme & AYT Tekrar Döngüsü",
    },
    "Dikkat & Çeldirici Tuzağı Temizliği": {
        "exam": "2027 LGS",
        "hours": 16,
        "goal": "Dikkat ve okuma kaynaklı net kayıplarını %70 azaltmak ve istikrarlı net artışı sağlamak",
        "focus": "Dikkat ve Çeldirici Tuzağı Azaltma",
        "weak": "Soru kökündeki 'ulaşılamaz / çıkarılamaz / değildir' ifadelerini kaçırma, öncül atlama ve çeldirici şıklara erken yönelme",
        "status": "Konu bilgisi yeterli olmasına rağmen hızlı okuma ve aceleci işaretleme nedeniyle bildiği sorularda hata yapıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "Soru Kökü Analizi, Çeldirici Tuzağı ve Dikkat Odaklı Okuma",
    },
}

SUBSCRIPTION_TIERS: List[Dict[str, Any]] = [
    {
        "id": "ogrenci_temel",
        "segment": "Öğrenci",
        "name": "LGS Başlangıç Dijital",
        "monthly_price": 1490,
        "monthly_quota": 250,
        "coach_quota": 20,
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
        "monthly_price": 2490,
        "monthly_quota": 1000,
        "coach_quota": 60,
        "badge": "Popüler",
        "desc": "Yeni nesil muhakeme soruları, Dijimentor birebir çalışma planı ve veli raporu.",
        "features": [
            "Aylık 1.000 yapay zekâ soru ve redaksiyon hakkı",
            "Dijimentor haftalık akıllı çalışma programı",
            "Veli Takip Paneli canlı bildirimleri",
            "2 sütunlu LGS kitapçık DOCX çıktısı",
        ],
    },
    {
        "id": "ogrenci_elit",
        "segment": "Öğrenci",
        "name": "LGS Tam Donanım Elit",
        "monthly_price": 3790,
        "monthly_quota": 2500,
        "coach_quota": 150,
        "badge": "Kapsamlı",
        "desc": "Yüksek soru üretim kotası, yoğun redaksiyon denetimi ve özel mentorluk senaryoları.",
        "features": [
            "Aylık 2.500 soru üretimi + otomatik redaksiyon",
            "Dijimentor günlük deneme analizi & çeldirici raporu",
            "Veli Takip Paneli + haftalık özet bülteni",
            "Özel kurum/öğrenci başlıklı DOCX testler",
        ],
    },
    {
        "id": "yks_sayisal",
        "segment": "YKS",
        "name": "YKS MF (Sayısal) Akademi",
        "monthly_price": 2890,
        "monthly_quota": 1200,
        "coach_quota": 80,
        "badge": "TYT + AYT",
        "desc": "TYT ve AYT Matematik, Fizik, Kimya, Biyoloji ve Türkçe odaklı ileri hazırlık paketi.",
        "features": [
            "TYT & AYT kazanım ve ÖSYM tarzı soru üretimi",
            "Dijimentor net artırma ve süre yönetimi mentorluğu",
            "Konu bazlı hata defteri ve çözüm desteği",
            "2 sütunlu deneme/test DOCX çıktısı",
        ],
    },
    {
        "id": "yks_ea_sozel",
        "segment": "YKS",
        "name": "YKS EA & Sözel Atölye",
        "monthly_price": 2690,
        "monthly_quota": 1200,
        "coach_quota": 80,
        "badge": "TYT + AYT",
        "desc": "Türk Dili ve Edebiyatı, Tarih, Coğrafya, Matematik ve paragraf analizi paketi.",
        "features": [
            "Paragraf, edebiyat ve sosyal bilimler soru fabrikası",
            "Dijimentor deneme stratejisi ve okuma hızı mentorluğu",
            "Detaylı çeldirici ve kazanım eşleştirme",
            "Düzenlenebilir DOCX çalışma kağıtları",
        ],
    },
    {
        "id": "yks_vip",
        "segment": "YKS",
        "name": "YKS Derece VIP",
        "monthly_price": 4290,
        "monthly_quota": 3000,
        "coach_quota": 200,
        "badge": "VIP",
        "desc": "Yoğun deneme analizi, kişiselleştirilmiş Dijimentor ve yüksek soru üretim kotası.",
        "features": [
            "Aylık 3.000 soru üretimi ve kalite denetimi",
            "Dijimentor birebir hedef üniversite & net simülasyonu",
            "Veli/danışman görünümü ve gelişim grafikleri",
            "Kurumsal standartta DOCX kitapçık çıktısı",
        ],
    },
    {
        "id": "veli_aile",
        "segment": "Veli",
        "name": "Aile Takip & Rehberlik",
        "monthly_price": 1690,
        "monthly_quota": 400,
        "coach_quota": 40,
        "badge": "Veli Odaklı",
        "desc": "Öğrencinin çalışma düzenini, deneme ivmesini ve eksik kazanımlarını şeffaf izleyin.",
        "features": [
            "Veli Takip Paneli haftalık performans özeti",
            "Dijimentor çalışma disiplini ve motivasyon planı",
            "Eksik kazanım uyarıları ve öneri listesi",
            "1 öğrenci profili dahil",
        ],
    },
    {
        "id": "veli_coklu",
        "segment": "Veli",
        "name": "Çoklu Kardeş Aile Paketi",
        "monthly_price": 2790,
        "monthly_quota": 1500,
        "coach_quota": 90,
        "badge": "2 Öğrenci",
        "desc": "İki kardeş için ayrı sınıf/sınav seviyelerinde takip, soru üretimi ve Dijimentor desteği.",
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
        "monthly_price": 5900,
        "monthly_quota": 4000,
        "coach_quota": 150,
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
        "monthly_price": 14900,
        "monthly_quota": 12000,
        "coach_quota": 400,
        "badge": "Kurumsal",
        "desc": "Kurs merkezleri ve butik okullar için ölçeklenebilir soru fabrikası ve öğrenci portalı.",
        "features": [
            "Aylık 12.000 soru üretimi ve toplu test hazırlama",
            "Öğrenci Portalı, Dijimentor ve Veli Paneli erişimi",
            "Kurum logolu/başlıklı DOCX sınav çıktıları",
            "25 öğretmen + 250 öğrenci hesabı",
        ],
    },
    {
        "id": "b2b_kampus",
        "segment": "Kurum / B2B",
        "name": "Okul & Kampüs Lisansı",
        "monthly_price": 29900,
        "monthly_quota": 35000,
        "coach_quota": 1000,
        "badge": "Kampüs",
        "desc": "Tüm kademelerde standart ölçme-değerlendirme ve akıllı mentorluk altyapısı.",
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
        "monthly_price": 54900,
        "monthly_quota": 100000,
        "coach_quota": 2500,
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

LOGO_DATA_URI = svg_to_data_uri(EMBEDDED_LOGO_SVG)
BACKPACK_DATA_URI = svg_to_data_uri(SCHOOL_BACKPACK_SVG)
LOGO_PNG_BYTES = get_embedded_logo_png_bytes()

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }
        [data-testid="stHeader"] {
            background: rgba(255, 255, 255, 0.72);
            backdrop-filter: blur(8px);
        }
        [data-testid="stDecoration"] {
            display: none;
        }
        .block-container {
            padding-top: 1.15rem;
            padding-bottom: 4.5rem;
            max-width: 1320px;
        }
        .kp-sidebar-logo {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 14px;
            padding: 0.55rem;
            margin-bottom: 0.75rem;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
        }
        .kp-sidebar-logo img {
            width: 132px;
            height: 132px;
            object-fit: contain;
            display: block;
        }
        .kp-hero {
            background: linear-gradient(125deg, #090e1a 0%, #16233b 52%, #7f1d1d 100%);
            color: #ffffff;
            padding: 1.35rem 1.7rem;
            border-radius: 16px;
            margin-bottom: 1.3rem;
            box-shadow: 0 14px 30px rgba(9, 14, 26, 0.2);
            border: 1px solid rgba(255, 255, 255, 0.1);
            display: flex;
            align-items: center;
            gap: 1.25rem;
        }
        .kp-hero-badge {
            background: #ffffff;
            border-radius: 14px;
            padding: 0.35rem;
            flex-shrink: 0;
            box-shadow: 0 6px 18px rgba(0, 0, 0, 0.28);
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .kp-hero-badge img {
            width: 84px;
            height: 84px;
            object-fit: contain;
            display: block;
        }
        .kp-hero-content {
            flex: 1;
            min-width: 0;
        }
        .kp-hero h1 {
            margin: 0;
            font-size: 1.62rem;
            font-weight: 800;
            letter-spacing: -0.2px;
            color: #ffffff;
        }
        .kp-hero p {
            margin: 0.32rem 0 0 0;
            font-size: 0.92rem;
            color: #e2e8f0;
            font-weight: 400;
        }
        .kp-tech-strip {
            display: flex;
            flex-wrap: wrap;
            gap: 0.5rem;
            margin-top: 0.78rem;
        }
        .kp-tech-pill {
            background: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.16);
            color: #f8fafc;
            font-size: 0.75rem;
            font-weight: 600;
            padding: 0.24rem 0.65rem;
            border-radius: 999px;
        }
        .kp-studio-banner {
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 65%, #fff5f5 100%);
            border: 1px solid #e2e8f0;
            border-left: 5px solid #881337;
            border-radius: 14px;
            padding: 1.15rem 1.4rem;
            margin-bottom: 1rem;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
        }
        .kp-studio-title {
            font-size: 1.24rem;
            font-weight: 700;
            color: #0f172a;
            margin: 0;
            letter-spacing: -0.15px;
        }
        .kp-studio-sub {
            font-size: 0.89rem;
            color: #475569;
            margin: 0.28rem 0 0 0;
            line-height: 1.45;
        }
        .kp-kpi-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 0.9rem 1.05rem;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
            transition: transform 0.15s ease, border-color 0.15s ease;
        }
        .kp-kpi-card:hover {
            transform: translateY(-2px);
            border-color: #cbd5e1;
        }
        .kp-kpi-label {
            font-size: 0.71rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.6px;
            color: #64748b;
        }
        .kp-kpi-value {
            font-size: 1.26rem;
            font-weight: 800;
            color: #0f172a;
            margin-top: 0.2rem;
        }
        .kp-kpi-sub {
            font-size: 0.77rem;
            color: #881337;
            font-weight: 600;
            margin-top: 0.16rem;
        }
        .kp-dist-bar {
            display: flex;
            height: 10px;
            border-radius: 999px;
            overflow: hidden;
            background: #f1f5f9;
            margin: 0.55rem 0 0.45rem 0;
            border: 1px solid #e2e8f0;
        }
        .kp-dist-legend {
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            font-size: 0.8rem;
            color: #475569;
        }
        .kp-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1.15rem 1.25rem;
            margin-bottom: 0.7rem;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.035);
            transition: transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease;
        }
        .kp-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.07);
            border-color: #94a3b8;
        }
        .kp-bag-shell {
            background: linear-gradient(145deg, #fffbeb 0%, #ffffff 55%, #fff1f2 100%);
            border: 1.5px solid #e2e8f0;
            border-radius: 16px;
            padding: 1.2rem 1.35rem;
            box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
        }
        .kp-bag-visual {
            position: relative;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 14px;
            padding: 0.85rem;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            min-height: 250px;
        }
        .kp-bag-visual img {
            width: 150px;
            height: 175px;
            object-fit: contain;
        }
        .kp-bag-count-pill {
            margin-top: 0.4rem;
            background: #0f172a;
            color: #fbbf24;
            font-size: 0.78rem;
            font-weight: 800;
            padding: 0.25rem 0.75rem;
            border-radius: 999px;
            letter-spacing: 0.3px;
        }
        .kp-bag-item-chip {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-left: 4px solid #881337;
            border-radius: 10px;
            padding: 0.5rem 0.75rem;
            margin-bottom: 0.45rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.85rem;
        }
        .kp-metric-box {
            background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1rem;
            text-align: center;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
        }
        .kp-metric-num {
            font-size: 1.55rem;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.2;
        }
        .kp-metric-lbl {
            font-size: 0.78rem;
            color: #64748b;
            margin-top: 0.25rem;
            text-transform: uppercase;
            letter-spacing: 0.45px;
            font-weight: 600;
        }
        div[data-testid="stSidebar"] .stButton > button {
            width: 100%;
            text-align: left;
            justify-content: flex-start;
            border-radius: 8px;
            padding: 0.5rem 0.75rem;
            font-weight: 500;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .stButton > button {
            border-radius: 8px;
            font-weight: 600;
        }
        div[data-testid="stTextInput"] input,
        div[data-testid="stTextArea"] textarea {
            border-radius: 8px;
            border-color: #cbd5e1;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# YARDIMCI FONKSİYONLAR: LOGO, ÇANTA SEPETİ, API, GEMINI, SAYAÇ, KOTA VE DOCX
# ==============================================================================
def attach_logo_to_doc(doc: Document, width_in: float = 1.15) -> None:
    """Mümkünse gömülü logoyu DOCX belgesinin üst kısmına ortalanmış olarak ekler."""
    if not LOGO_PNG_BYTES:
        return
    try:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        p.add_run().add_picture(io.BytesIO(LOGO_PNG_BYTES), width=Inches(width_in))
    except Exception:
        pass


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


def call_gemini_text(
    prompt: str,
    system_instruction: str,
    model: str = DEFAULT_MODEL,
) -> str:
    """Gemini API üzerinden metin üretir; 503/429 yoğunluk ve 404 hatalarında tekrar dener ve yedek modele geçer."""
    client = get_gemini_client()
    candidate_models = [model, FALLBACK_MODEL, "gemini-2.5-pro"]
    deduped_models: List[str] = []
    for m in candidate_models:
        if m and m not in deduped_models:
            deduped_models.append(m)

    last_exc: Optional[Exception] = None

    for candidate_model in deduped_models:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=candidate_model,
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
            except Exception as exc:
                last_exc = exc
                err_str = str(exc)
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                if any(tok in err_str for tok in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "overloaded")):
                    time.sleep(1.2 * (attempt + 1))
                    continue
                raise

    raise RuntimeError(
        f"Model çağrısı geçici yoğunluk nedeniyle tamamlanamadı ({last_exc}). "
        "Lütfen birkaç saniye sonra tekrar deneyin."
    )


def navigate_to(page_name: str) -> None:
    """İstenen sayfaya geçiş yapar ve uygulamayı yeniler."""
    if page_name in PAGES:
        st.session_state.active_page = page_name
        st.rerun()


def record_activity(module_name: str, summary: str) -> None:
    """Oturum içi teknolojik aktivite akışına kayıt ekler."""
    logs: List[Dict[str, str]] = st.session_state.setdefault("activity_log", [])
    logs.insert(
        0,
        {
            "time": datetime.now().strftime("%H:%M:%S"),
            "module": module_name,
            "summary": summary,
        },
    )
    st.session_state["activity_log"] = logs[:12]


def get_tier_by_id(tier_id: str) -> Dict[str, Any]:
    """Paket kimliğine göre abonelik paketi kaydını döndürür."""
    return next((t for t in SUBSCRIPTION_TIERS if t["id"] == tier_id), SUBSCRIPTION_TIERS[1])


def ensure_bag_initialized() -> None:
    """Sepet sözlüğünü yalnızca ilk oturum açılışında başlatır; kullanıcı boşaltırsa yeniden doldurmaz."""
    if "backpack_cart" not in st.session_state or not isinstance(st.session_state["backpack_cart"], dict):
        default_id = st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
        st.session_state["backpack_cart"] = {default_id: 1}


def get_bag_items() -> Dict[str, int]:
    """Okul sırt çantasındaki {tier_id: adet} eşleşmesini döndürür."""
    ensure_bag_initialized()
    cleaned: Dict[str, int] = {}
    for k, v in st.session_state["backpack_cart"].items():
        try:
            qty = int(v)
        except Exception:
            qty = 0
        if qty > 0:
            cleaned[str(k)] = qty
    st.session_state["backpack_cart"] = cleaned
    return cleaned


def get_bag_summary() -> Dict[str, Any]:
    """Okul sırt çantasındaki tüm paketlerin toplam fiyat, kota ve adedini hesaplar."""
    bag = get_bag_items()
    entries: List[Dict[str, Any]] = []
    total_monthly_price = 0
    total_monthly_quota = 0
    total_coach_quota = 0
    total_count = 0

    for tid, qty in bag.items():
        tier = get_tier_by_id(tid)
        line_price = int(tier["monthly_price"]) * qty
        line_quota = int(tier["monthly_quota"]) * qty
        line_coach = int(tier["coach_quota"]) * qty

        total_monthly_price += line_price
        total_monthly_quota += line_quota
        total_coach_quota += line_coach
        total_count += qty

        entries.append(
            {
                "tier": tier,
                "qty": qty,
                "line_price": line_price,
                "line_quota": line_quota,
                "line_coach": line_coach,
            }
        )

    if entries:
        title = (
            entries[0]["tier"]["name"]
            if len(entries) == 1
            else f"Sırt Çantası ({total_count} Paket)"
        )
        segment = (
            entries[0]["tier"]["segment"]
            if len(entries) == 1
            else "Karma / Çoklu Paket"
        )
        badge = entries[0]["tier"]["badge"] if len(entries) == 1 else f"{len(entries)} Çeşit"
    else:
        title = "Boş Çanta"
        segment = "Seçim Bekleniyor"
        badge = "0 Paket"

    return {
        "id": "backpack_bundle",
        "name": title,
        "segment": segment,
        "badge": badge,
        "monthly_price": total_monthly_price,
        "monthly_quota": total_monthly_quota,
        "coach_quota": total_coach_quota,
        "total_count": total_count,
        "entries": entries,
    }


def sync_plan_quotas(summary: Dict[str, Any], force_reset: bool = False) -> None:
    """Çantadaki toplam kotalara göre oturum kotalarını eşitler."""
    sig = f"{summary['total_count']}:{summary['monthly_quota']}:{summary['coach_quota']}"
    if force_reset or st.session_state.get("quota_tier_id") != sig:
        st.session_state["quota_tier_id"] = sig
        st.session_state["remaining_question_quota"] = int(summary["monthly_quota"])
        st.session_state["remaining_coach_quota"] = int(summary["coach_quota"])


def cb_add_tier_to_bag(tier_id: str, qty: int = 1) -> None:
    """Callback: Paketi okul sırt çantasına (sepete) ekler ve kotaları günceller."""
    bag = dict(get_bag_items())
    bag[tier_id] = bag.get(tier_id, 0) + max(1, int(qty))
    st.session_state["backpack_cart"] = bag
    st.session_state["selected_tier_id"] = tier_id
    sync_plan_quotas(get_bag_summary(), force_reset=True)
    tier = get_tier_by_id(tier_id)
    record_activity("Sırt Çantası", f"{tier['name']} çantaya eklendi")


def cb_update_bag_qty(tier_id: str, delta: int) -> None:
    """Callback: Çantadaki paketin adedini artırır/azaltır veya 0 ise tamamen siler."""
    bag = dict(get_bag_items())
    if tier_id not in bag:
        return
    new_qty = bag[tier_id] + delta
    tier = get_tier_by_id(tier_id)
    if new_qty <= 0:
        del bag[tier_id]
        record_activity("Sırt Çantası", f"{tier['name']} çantadan çıkarıldı")
    else:
        bag[tier_id] = new_qty
        record_activity("Sırt Çantası", f"{tier['name']} adedi {new_qty} yapıldı")
    st.session_state["backpack_cart"] = bag
    if bag:
        st.session_state["selected_tier_id"] = next(iter(bag.keys()))
    sync_plan_quotas(get_bag_summary(), force_reset=True)


def cb_clear_bag() -> None:
    """Callback: Okul sırt çantasını tamamen boşaltır."""
    st.session_state["backpack_cart"] = {}
    sync_plan_quotas(get_bag_summary(), force_reset=True)
    record_activity("Sırt Çantası", "Çantadaki tüm paketler silindi")


def consume_quota(kind: str, amount: int = 1) -> Tuple[bool, str]:
    """Soru ('question') veya mentorluk ('coach') kotasından kullanım düşer."""
    summary = get_bag_summary()
    sync_plan_quotas(summary, force_reset=False)

    if kind == "question":
        rem = int(st.session_state.get("remaining_question_quota", 0))
        if rem < amount:
            return False, (
                f"Okul sırt çantanızda ({summary['name']}) yeterli soru üretim kotası kalmadı "
                f"(Kalan: {rem}, İstenen: {amount}). Paketler vitrininden paket ekleyebilirsiniz."
            )
        st.session_state["remaining_question_quota"] = rem - amount
        return True, ""

    rem_c = int(st.session_state.get("remaining_coach_quota", 0))
    if rem_c < amount:
        return False, (
            f"Okul sırt çantanızda ({summary['name']}) Dijimentor plan kotası doldu "
            f"(Kalan: {rem_c}). Paketler vitrininden paket ekleyebilirsiniz."
        )
    st.session_state["remaining_coach_quota"] = rem_c - amount
    return True, ""


def extract_answer_key(raw_text: str) -> Dict[int, str]:
    """Metindeki 'Cevap Anahtarı' bölümünden {soru_no: 'A'|'B'|'C'|'D'} haritası çıkarır."""
    if not raw_text.strip():
        return {}

    parts = re.split(r"(?:^|\n)#*\s*Cevap\s*Anahtar[^\n]*", raw_text, maxsplit=1, flags=re.IGNORECASE)
    search_space = parts[1] if len(parts) > 1 else raw_text
    key_map: Dict[int, str] = {}

    patterns = [
        r"(?:Soru\s*)?(\d+)\s*[\.\-\:\)]\s*\**\s*([A-D])\b",
        r"(\d+)\s*\.\s*Soru\s*[\:\-\)]\s*\**\s*([A-D])\b",
    ]
    for pat in patterns:
        for m in re.finditer(pat, search_space, flags=re.IGNORECASE):
            q_num = int(m.group(1))
            ans_letter = m.group(2).upper()
            if q_num not in key_map:
                key_map[q_num] = ans_letter

    return key_map


def parse_quiz_questions(raw_text: str) -> List[Dict[str, Any]]:
    """Üretilen test metninden etkileşimli mini deneme için soruları ve cevap anahtarını ayıklar."""
    if not raw_text.strip():
        return []

    answer_map = extract_answer_key(raw_text)
    main_body = re.split(r"(?:^|\n)#*\s*Cevap\s*Anahtar", raw_text, maxsplit=1, flags=re.IGNORECASE)[0]
    blocks = re.split(r"(?=(?:^|\n)#*\s*Soru\s*\d+)", main_body.strip(), flags=re.IGNORECASE)
    parsed: List[Dict[str, Any]] = []

    for block in blocks:
        text = block.strip()
        num_match = re.match(r"^#*\s*Soru\s*(\d+)", text, flags=re.IGNORECASE)
        if not num_match:
            continue
        q_num = int(num_match.group(1))
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        stem_lines: List[str] = []
        options: Dict[str, str] = {}
        current_opt: Optional[str] = None

        for ln in lines:
            if re.match(r"^\[?Redaksiyon", ln, flags=re.IGNORECASE):
                continue
            opt_match = re.match(r"^[\-\*\s]*([A-D])\s*[\)\.\:]\s*(.+)$", ln, flags=re.IGNORECASE)
            if opt_match:
                letter = opt_match.group(1).upper()
                options[letter] = opt_match.group(2).strip()
                current_opt = letter
            elif current_opt:
                options[current_opt] = f"{options[current_opt]} {ln}".strip()
            else:
                stem_lines.append(ln)

        if len(options) >= 2:
            parsed.append(
                {
                    "num": q_num,
                    "stem": "\n".join(stem_lines),
                    "options": options,
                    "correct": answer_map.get(q_num),
                    "full_text": text,
                }
            )

    return parsed


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

    attach_logo_to_doc(doc, width_in=1.15)

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
    footer_run = footer_p.add_run(f"\n{BRAND_NAME}")
    footer_run.italic = True
    footer_run.font.size = Pt(8.5)
    footer_run.font.color.rgb = RGBColor(100, 116, 139)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_coach_docx(title_meta: str, plan_text: str) -> bytes:
    """Dijimentor çalışma planını kurumsal DOCX belgesine çevirir."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} — DİJİMENTOR ÇALIŞMA PLANI")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(15, 23, 42)

    m = doc.add_paragraph()
    mr = m.add_run(f"{title_meta}\nTarih: {datetime.now().strftime('%d.%m.%Y')}")
    mr.font.size = Pt(9.5)
    mr.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph("─" * 78)

    for line in plan_text.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(cleaned)
        run.font.size = Pt(10.5)
        if cleaned.startswith("#") or cleaned.endswith(":"):
            run.bold = True
            run.font.color.rgb = RGBColor(153, 27, 27)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_commercial_brief_docx(
    title: str,
    summary: Dict[str, Any],
    billing_cycle: str,
    installment: int,
    total_amount: int,
    extra_notes: str,
) -> bytes:
    """Okul sırt çantasındaki paketler veya kurumsal lisans için DOCX teklif/özet belgesi üretir."""
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} — OKUL SIRT ÇANTASI & ABONELİK ÖZETİ")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(15, 23, 42)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(f"{title}  |  Tarih: {datetime.now().strftime('%d.%m.%Y')}").font.size = Pt(9.5)

    doc.add_paragraph("─" * 78)

    lines: List[str] = [
        f"Çanta Durumu: {summary['name']} ({summary['total_count']} Adet Paket)",
        f"Aylık Liste Toplamı: {summary['monthly_price']:,} ₺".replace(",", "."),
        f"Tercih Edilen Dönem: {billing_cycle}",
        f"iyzico Taksit Planı: {installment} Taksit",
        f"Hesaplanan Toplam Tutar: {total_amount:,} ₺".replace(",", "."),
        f"Toplam Aylık Soru Kotası: {summary['monthly_quota']:,} Soru".replace(",", "."),
        f"Toplam Dijimentor Plan Kotası: {summary['coach_quota']} Plan / Ay",
        "",
        "Çantadaki Paketler:",
    ]

    for entry in summary.get("entries", []):
        t = entry["tier"]
        qty = entry["qty"]
        lines.append(
            f"• {t['name']} x{qty} ({t['segment']}) — {entry['line_price']:,} ₺/ay | "
            f"{entry['line_quota']:,} soru/ay".replace(",", ".")
        )

    if extra_notes.strip():
        lines += ["", "Ek / Kurumsal Notlar:", extra_notes.strip()]

    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3.5)
        run = p.add_run(line)
        run.font.size = Pt(10.5)
        if line.startswith(("Çanta Durumu", "Çantadaki Paketler", "Ek / Kurumsal Notlar")):
            run.bold = True
            run.font.color.rgb = RGBColor(153, 27, 27)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_school_backpack_panel() -> Dict[str, Any]:
    """Seçilen paketleri alışveriş sepeti gibi Okul Sırt Çantası görseli içinde gösterir ve yönetir."""
    summary = get_bag_summary()
    entries = summary["entries"]

    st.markdown('<div class="kp-bag-shell">', unsafe_allow_html=True)
    st.markdown("### 🎒 Okul Sırt Çantam (Seçili Paket Sepeti)")
    st.caption(
        "Vitrin veya paket kartlarından seçtiğiniz paketler okul sırt çantanızda toplanır; "
        "adet güncelleyebilir, çıkarabilir veya çantadaki tüm paketler için iyzico ödemesi oluşturabilirsiniz."
    )

    v_col, list_col = st.columns([1.15, 2.25])

    with v_col:
        chips_html = ""
        for entry in entries:
            safe_name = html.escape(entry["tier"]["name"])
            safe_qty = int(entry["qty"])
            chips_html += (
                f'<div style="background:#fef3c7;color:#0f172a;border:1px solid #f59e0b;'
                f'border-radius:8px;padding:0.22rem 0.5rem;font-size:0.74rem;font-weight:700;'
                f'margin-top:0.25rem;width:100%;text-align:center;">'
                f'📘 {safe_name} (x{safe_qty})</div>'
            )

        if not chips_html:
            chips_html = (
                '<div style="color:#64748b;font-size:0.78rem;margin-top:0.35rem;text-align:center;">'
                'Çanta şu an boş</div>'
            )

        st.markdown(
            f"""
            <div class="kp-bag-visual">
                <img src="{BACKPACK_DATA_URI}" alt="Okul Sırt Çantası Sepeti"/>
                <div class="kp-bag-count-pill">Çantada {summary['total_count']} Paket</div>
                <div style="width:100%;margin-top:0.35rem;">{chips_html}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with list_col:
        if not entries:
            st.info(
                "Okul sırt çantanız boş. Aşağıdaki paket kartlarından **Çantaya Ekle** butonuna "
                "tıklayarak seçtiğiniz paketleri çantanın içine alabilirsiniz."
            )
        else:
            for entry in entries:
                t = entry["tier"]
                qty = entry["qty"]
                c_info, c_minus, c_plus, c_rem = st.columns([4.4, 0.85, 0.85, 1.1])
                with c_info:
                    st.markdown(
                        f"""
                        <div class="kp-bag-item-chip">
                            <div>
                                <b>{html.escape(t['name'])}</b>
                                <span style="color:#64748b;font-size:0.78rem;"> • {html.escape(t['segment'])}</span><br>
                                <span style="color:#475569;font-size:0.78rem;">
                                    Adet: <b>{qty}</b> • Aylık Kota: <b>{entry['line_quota']:,}</b> Soru
                                </span>
                            </div>
                            <div style="font-weight:800;color:#881337;white-space:nowrap;">
                                {entry['line_price']:,} ₺/ay
                            </div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )
                with c_minus:
                    st.button(
                        "−",
                        key=f"bag_minus_{t['id']}",
                        use_container_width=True,
                        on_click=cb_update_bag_qty,
                        args=(t["id"], -1),
                    )
                with c_plus:
                    st.button(
                        "+",
                        key=f"bag_plus_{t['id']}",
                        use_container_width=True,
                        on_click=cb_update_bag_qty,
                        args=(t["id"], +1),
                    )
                with c_rem:
                    st.button(
                        "Çıkar",
                        key=f"bag_del_{t['id']}",
                        use_container_width=True,
                        on_click=cb_update_bag_qty,
                        args=(t["id"], -qty),
                    )

            m1, m2, m3 = st.columns(3)
            m1.metric("Çanta Aylık Toplam", f"{summary['monthly_price']:,} ₺".replace(",", "."))
            m2.metric("Toplam Aylık Soru", f"{summary['monthly_quota']:,}".replace(",", "."))
            m3.metric("Toplam Dijimentor Plan", f"{summary['coach_quota']} / Ay")

            st.button(
                "🗑️ Çantayı Boşalt",
                key="bag_clear_all",
                on_click=cb_clear_bag,
            )

    st.markdown("</div>", unsafe_allow_html=True)
    return summary


def render_iyzico_simulator(
    package_name: str,
    monthly_price: int,
    summary: Optional[Dict[str, Any]] = None,
) -> None:
    """iyzico taksit ve ödeme simülatörü."""
    bag_summary = summary if summary is not None else get_bag_summary()
    effective_monthly = (
        int(bag_summary["monthly_price"]) if bag_summary["total_count"] > 0 else int(monthly_price)
    )
    display_label = (
        f"{package_name} ({bag_summary['total_count']} Paket)"
        if bag_summary["total_count"] > 0
        else package_name
    )

    st.markdown("### 💳 iyzico Taksit & Ödeme Simülatörü")
    c1, c2, c3 = st.columns(3)
    with c1:
        billing_cycle = st.selectbox(
            "Dönem Seçimi",
            ["Aylık Abonelik", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"],
            key=f"cycle_{display_label}",
        )
    with c2:
        installment = st.selectbox(
            "Taksit Sayısı (iyzico)",
            [1, 2, 3, 6, 9, 12],
            index=2,
            key=f"inst_{display_label}",
        )
    with c3:
        card_brand = st.selectbox(
            "Kart Programı",
            ["Bonus", "Maximum", "World", "Axess", "Paraf", "Diğer Banka Kartı"],
            key=f"card_{display_label}",
        )

    if "3 Aylık" in billing_cycle:
        total_amount = round(effective_monthly * 3 * 0.90)
    elif "Yıllık" in billing_cycle:
        total_amount = round(effective_monthly * 12 * 0.80)
    else:
        total_amount = effective_monthly

    monthly_installment = total_amount / installment

    m1, m2, m3 = st.columns(3)
    m1.metric("Toplam Tutar", f"{total_amount:,} ₺".replace(",", "."))
    m2.metric(f"{installment} Taksit Aylık", f"{monthly_installment:,.2f} ₺".replace(",", " "))
    m3.metric("Ödeme Altyapısı", f"iyzico • {card_brand}")

    btn_c1, btn_c2 = st.columns([2, 1])
    with btn_c1:
        if st.button("iyzico Ödeme Talebi Oluştur (Simülasyon)", key=f"pay_btn_{display_label}"):
            sync_plan_quotas(bag_summary, force_reset=True)
            record_activity(
                "iyzico Ödeme",
                f"{display_label} ({total_amount:,} ₺) aktiflendi".replace(",", "."),
            )
            st.success(
                f"**{display_label}** için **{total_amount:,} ₺** ({installment} taksit, "
                f"{card_brand}) tutarında iyzico ödeme oturumu hazırlandı ve toplam kota "
                f"({bag_summary['monthly_quota']:,} soru) aktifleşti.".replace(",", ".")
            )
    with btn_c2:
        brief_docx = build_commercial_brief_docx(
            title=f"{display_label} Özeti",
            summary=bag_summary,
            billing_cycle=billing_cycle,
            installment=int(installment),
            total_amount=int(total_amount),
            extra_notes=f"Kart Programı: {card_brand}",
        )
        st.download_button(
            label="📥 Teklif / Özet DOCX",
            data=brief_docx,
            file_name="kursun_kalem_sirt_cantasi_ozet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl_brief_{display_label}",
        )


# ==============================================================================
# SESSION STATE BAŞLATMA
# ==============================================================================
if "active_page" not in st.session_state:
    st.session_state.active_page = PAGES[0]
if "generated_test_text" not in st.session_state:
    st.session_state.generated_test_text = ""
if "dijimentor_plan" not in st.session_state:
    st.session_state.dijimentor_plan = ""
if "selected_tier_id" not in st.session_state:
    st.session_state.selected_tier_id = SUBSCRIPTION_TIERS[1]["id"]

ensure_bag_initialized()

if "live_tick" not in st.session_state:
    st.session_state.live_tick = False
if "portal_question_draft" not in st.session_state:
    st.session_state.portal_question_draft = ""
if "portal_exam_source" not in st.session_state:
    st.session_state.portal_exam_source = ""
if "activity_log" not in st.session_state:
    st.session_state.activity_log = [
        {"time": "09:15:00", "module": "Soru Fabrikası", "summary": "8. Sınıf LGS Muhakeme seti hazır"},
        {"time": "10:42:18", "module": "Dijimentor", "summary": "Haftalık çeldirici & tempo analizi aktif"},
    ]

bag_summary_global = get_bag_summary()
sync_plan_quotas(bag_summary_global, force_reset=False)
api_key_ready = resolve_gemini_api_key() is not None

# ==============================================================================
# KENAR ÇUBUĞU (KALICI LOGO + CUSTOM BUTTON NAVİGASYON + ÇANTA & KOTA ÖZETİ)
# ==============================================================================
with st.sidebar:
    st.markdown(
        f"""
        <div class="kp-sidebar-logo">
            <img src="{LOGO_DATA_URI}" alt="{BRAND_NAME} Logosu"/>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div style="background:#0f172a;color:#ffffff;padding:0.9rem 1rem;border-radius:10px;margin-bottom:0.85rem;">
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
        "Dijimentor": "🤖",
        "Veli Takip Paneli": "👨‍👩‍👧",
        "B2B Lisans & Kurum": "🏫",
        "Paketler & Abonelik Vitrini": "💎",
    }

    for p_name in PAGES:
        is_current = st.session_state.active_page == p_name
        btn_type = "primary" if is_current else "secondary"
        suffix = f" ({bag_summary_global['total_count']})" if p_name == "Paketler & Abonelik Vitrini" else ""
        label = f"{page_icons.get(p_name, '•')} {p_name}{suffix}"
        if st.button(label, key=f"nav_{p_name}", type=btn_type):
            st.session_state.active_page = p_name
            st.rerun()

    st.markdown("---")
    st.markdown(
        f"""
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:0.7rem 0.85rem;margin-bottom:0.75rem;">
            <div style="font-size:0.72rem;font-weight:700;color:#64748b;text-transform:uppercase;">🎒 Okul Sırt Çantası</div>
            <div style="font-size:0.88rem;font-weight:700;color:#0f172a;margin-top:0.1rem;">
                {bag_summary_global['total_count']} Paket • {bag_summary_global['monthly_price']:,} ₺/ay
            </div>
            <div style="font-size:0.78rem;color:#475569;margin-top:0.25rem;">
                Soru Kotası: <b>{int(st.session_state.remaining_question_quota):,}</b>/{bag_summary_global['monthly_quota']:,}<br>
                Dijimentor Plan: <b>{int(st.session_state.remaining_coach_quota)}</b>/{bag_summary_global['coach_quota']}
            </div>
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )

    if api_key_ready:
        st.success("● Sistem Aktif • 2027 LGS & Maarif Modeli")
    else:
        st.warning("● API Anahtarı Bekleniyor (Secrets)")

# ==============================================================================
# ÜST KURUMSAL BANNER (KALICI GÖMÜLÜ LOGO + BAŞLIK + SLOGAN)
# ==============================================================================
st.markdown(
    f"""
    <div class="kp-hero">
        <div class="kp-hero-badge">
            <img src="{LOGO_DATA_URI}" alt="{BRAND_NAME} Logosu"/>
        </div>
        <div class="kp-hero-content">
            <h1>✏️ {BRAND_NAME}</h1>
            <p>{BRAND_TAGLINE}</p>
            <div class="kp-tech-strip">
                <span class="kp-tech-pill">⚡ Gemini AI Motoru</span>
                <span class="kp-tech-pill">📐 Maarif Modeli Kazanım Uyumu</span>
                <span class="kp-tech-pill">🎯 Dijimentor v3 Analitik</span>
                <span class="kp-tech-pill">📄 Kurumsal DOCX Stüdyo</span>
                <span class="kp-tech-pill">💳 iyzico Altyapı</span>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

current_page = st.session_state.active_page

# ==============================================================================
# 1. ANA SAYFA & CANLI GERİ SAYIM + ANALİTİK KOMUTA MERKEZİ
# ==============================================================================
if current_page == "Ana Sayfa & Geri Sayım":
    st.subheader("🚀 Kurşun Kalem Dijital Eğitim Platformuna Hoş Geldiniz")
    st.write(
        "İlkokuldan Liseye ve YKS'ye uzanan, yapay zekâ destekli Türkiye'nin akıllı "
        "soru fabrikası, Dijimentor ve ölçme-değerlendirme ekosistemi."
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

    st.markdown("### 💡 Platform Modülleri & Hızlı Geçiş")
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
        if st.button("Soru Fabrikasına Git →", key="go_factory", use_container_width=True):
            navigate_to("Soru & Redaksiyon Fabrikası")
    with f2:
        st.markdown(
            """
            <div class="kp-card">
                <h4>🤖 Dijimentor Akıllı Rehberlik</h4>
                <p>Öğrencinin hedef sınavına, zayıf kazanımlarına ve haftalık temposuna göre
                kişiselleştirilmiş çalışma ve deneme stratejisi sunar.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Dijimentor'u Aç →", key="go_mentor", use_container_width=True):
            navigate_to("Dijimentor")
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
        if st.button("Paketler & Abonelik Vitrini →", key="go_plans", use_container_width=True):
            navigate_to("Paketler & Abonelik Vitrini")

    st.markdown("---")
    ana_l, ana_r = st.columns([2, 1])
    with ana_l:
        st.markdown("#### 📈 Haftalık Soru Üretim & Çözüm Telemetrisi")
        telemetry_data = {
            "Pzt": 16800,
            "Sal": 18200,
            "Çar": 19500,
            "Per": 21100,
            "Cum": 22400,
            "Cmt": 25900,
            "Paz": 24800,
        }
        st.area_chart(telemetry_data, height=205)
    with ana_r:
        st.markdown("#### ⚡ Canlı Oturum Akışı")
        for item in st.session_state.get("activity_log", [])[:5]:
            st.markdown(
                f"""
                <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:0.55rem 0.75rem;margin-bottom:0.45rem;">
                    <div style="font-size:0.72rem;font-weight:700;color:#881337;">{item['time']} • {item['module']}</div>
                    <div style="font-size:0.82rem;color:#0f172a;margin-top:0.1rem;">{item['summary']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

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

    if st.session_state.get("factory_focus_topic"):
        st.info(f"Dijimentor'dan aktarılan odak konu: **{st.session_state['factory_focus_topic']}**")

    grade_options = ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS)", "9-10. Sınıf", "11-12. Sınıf (YKS)"]
    default_branch = st.session_state.get("factory_branch", "Türkçe")
    if default_branch not in CURRICULUM_TOPICS:
        default_branch = "Türkçe"
    default_grade = st.session_state.get("factory_grade", "5. Sınıf")
    if default_grade not in grade_options:
        default_grade = "5. Sınıf"

    with st.expander("📋 Test ve Kurum Parametreleri", expanded=True):
        r1c1, r1c2, r1c3 = st.columns(3)
        with r1c1:
            institution = st.text_input("Kurum / Okul Adı", value=BRAND_NAME)
            branch = st.selectbox(
                "Branş Seçimi",
                list(CURRICULUM_TOPICS.keys()),
                index=list(CURRICULUM_TOPICS.keys()).index(default_branch),
            )
            grade = st.selectbox(
                "Sınıf Seviyesi",
                grade_options,
                index=grade_options.index(default_grade),
            )
        with r1c2:
            areas = list(CURRICULUM_TOPICS[branch].keys())
            learning_area = st.selectbox("Öğrenme Alanı", areas, index=0)
            outcomes = CURRICULUM_TOPICS[branch][learning_area]
            outcome = st.selectbox("Hedef Kazanım", outcomes, index=0)
            default_theme = st.session_state.get("factory_focus_topic") or "HAYAT"
            context_theme = st.text_input("Bağlam / Tema / Senaryo", value=default_theme)
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
        quota_ok, quota_msg = consume_quota("question", int(question_count))
        if not quota_ok:
            st.error(quota_msg)
        else:
            sys_prompt = (
                "Sen Türkiye Yüzyılı Maarif Modeli, LGS ve YKS ölçme-değerlendirme standartlarında "
                "uzman bir Türkçe ve test yazarısın. Soruların özgün, açık, çeldiricileri güçlü, "
                "dil bilgisi ve yazım kurallarına kusursuz uygun olmasını sağla. "
                "Otomatik redaksiyon açıksa her sorunun sonunda kısa bir [Redaksiyon & Çeldirici Notu] ekle "
                "ve en sonda mutlaka 'Cevap Anahtarı' başlığı altında '1-A', '2-C' formatında net anahtar ver."
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
                "Lütfen numaralandırılmış çoktan seçmeli (Soru 1, Soru 2... ve A, B, C, D şıkları) özgün sorular üret."
            )
            try:
                with st.spinner("Sorular üretiliyor ve kalite denetimi yapılıyor..."):
                    result = call_gemini_text(prompt=user_prompt, system_instruction=sys_prompt)
                st.session_state.generated_test_text = result
                record_activity("Soru Fabrikası", f"{branch} • {question_count} soru üretildi")
                st.success(
                    f"Soru paketi başarıyla oluşturuldu. Kalan soru kotası: "
                    f"{int(st.session_state.remaining_question_quota):,}".replace(",", ".")
                )
            except Exception as exc:
                st.session_state["remaining_question_quota"] = int(
                    st.session_state.get("remaining_question_quota", 0)
                ) + int(question_count)
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
        act_c1, act_c2, act_c3 = st.columns(3)
        with act_c1:
            st.download_button(
                label="📥 Testi DOCX Olarak İndir",
                data=docx_bytes,
                file_name=f"kursun_kalem_{branch.lower().replace(' ', '_')}_test.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with act_c2:
            if st.button("🎓 Mini Deneme Olarak Çöz", use_container_width=True):
                st.session_state["portal_exam_source"] = st.session_state.generated_test_text
                record_activity("Köprü", "Test Öğrenci Portalı mini denemesine aktarıldı")
                navigate_to("Öğrenci Portalı & Çözüm")
        with act_c3:
            if st.button("💡 İlk Soruyu Çözüm Rehberine Gönder", use_container_width=True):
                parsed_q = parse_quiz_questions(st.session_state.generated_test_text)
                st.session_state["portal_question_draft"] = (
                    parsed_q[0]["full_text"] if parsed_q else st.session_state.generated_test_text[:1000]
                )
                navigate_to("Öğrenci Portalı & Çözüm")

# ==============================================================================
# 3. ÖĞRENCİ PORTALI & ÇÖZÜM + İNTERAKTİF MİNİ DENEME (NET & DOĞRULUK MOTORU)
# ==============================================================================
elif current_page == "Öğrenci Portalı & Çözüm":
    st.subheader("🎓 Öğrenci Portalı, Mini Deneme & Adımlı Çözüm Asistanı")
    st.write(
        "Üretilen soruları otomatik cevap anahtarı kontrollü mini deneme olarak çözebilir veya "
        "takıldığın soruyu adım adım mentorluk rehberiyle analiz edebilirsin."
    )

    tab_exam, tab_hint = st.tabs(["📝 İnteraktif Mini Deneme & Net Analizi", "🔍 Adımlı Çözüm Asistanı"])

    with tab_exam:
        exam_source = st.text_area(
            "Deneme Soru Metni (Soru Fabrikasından otomatik gelir veya buraya yapıştır)",
            value=st.session_state.get("portal_exam_source") or st.session_state.get("generated_test_text", ""),
            height=130,
            placeholder="Soru 1) ... A) ... B) ... C) ... D) ... Cevap Anahtarı: 1-B",
        )
        st.session_state["portal_exam_source"] = exam_source
        questions = parse_quiz_questions(exam_source)

        if not questions:
            st.info(
                "Etkileşimli şık işaretleyebilmek için **Soru & Redaksiyon Fabrikası** sayfasından "
                "bir test üretin veya 'Soru 1 ... A) ... B) ... C) ... D) ...' biçiminde sorular yapıştırın."
            )
        else:
            st.markdown(f"#### 🎯 Aktif Mini Deneme ({len(questions)} Soru)")
            user_answers: Dict[int, str] = {}
            for idx, q in enumerate(questions):
                with st.container(border=True):
                    st.markdown(q["stem"])
                    opts = list(q["options"].keys())
                    chosen = st.radio(
                        f"Soru {q['num']} Yanıtın",
                        options=opts,
                        format_func=lambda k, o_map=q["options"]: f"{k}) {o_map[k]}",
                        horizontal=True,
                        key=f"mini_q_{idx}",
                    )
                    user_answers[idx] = chosen

            if st.button("✅ Denemeyi Bitir ve Netimi Hesapla", type="primary"):
                correct_cnt = 0
                wrong_cnt = 0
                ungraded_cnt = 0
                wrong_questions: List[Dict[str, Any]] = []
                result_rows: List[Dict[str, str]] = []

                for idx, q in enumerate(questions):
                    u_ans = user_answers[idx]
                    c_ans = q.get("correct")
                    if not c_ans:
                        ungraded_cnt += 1
                        status_txt = "Anahtar Yok"
                    elif u_ans == c_ans:
                        correct_cnt += 1
                        status_txt = "✅ Doğru"
                    else:
                        wrong_cnt += 1
                        wrong_questions.append(q)
                        status_txt = "❌ Yanlış"

                    result_rows.append(
                        {
                            "Soru": f"Soru {q['num']}",
                            "Senin Yanıtın": u_ans,
                            "Doğru Cevap": c_ans or "Belirtilmedi",
                            "Durum": status_txt,
                        }
                    )

                graded_total = correct_cnt + wrong_cnt
                net_score = max(0.0, correct_cnt - (wrong_cnt / 3.0))
                accuracy = (correct_cnt / graded_total * 100.0) if graded_total > 0 else 0.0

                st.session_state["last_exam_report"] = {
                    "correct": correct_cnt,
                    "wrong": wrong_cnt,
                    "net": net_score,
                    "accuracy": accuracy,
                    "rows": result_rows,
                    "first_wrong": wrong_questions[0]["full_text"] if wrong_questions else "",
                }
                record_activity(
                    "Mini Deneme",
                    f"{correct_cnt}D {wrong_cnt}Y • {net_score:.2f} Net (%{accuracy:.0f})",
                )

            report = st.session_state.get("last_exam_report")
            if report:
                st.markdown("---")
                r1, r2, r3, r4 = st.columns(4)
                r1.metric("Doğru Sayısı", f"{report['correct']} Soru")
                r2.metric("Yanlış Sayısı", f"{report['wrong']} Soru")
                r3.metric("Hesaplanan Net (3Y=1D)", f"{report['net']:.2f} Net")
                r4.metric("Başarı Oranı", f"%{report['accuracy']:.1f}")

                st.dataframe(report["rows"], use_container_width=True, hide_index=True)

                if report.get("first_wrong"):
                    if st.button("🔍 Yanlış Yaptığım İlk Soruyu Çözüm Rehberine Aktar"):
                        st.session_state["portal_question_draft"] = report["first_wrong"]
                        st.toast("Yanlış soru Adımlı Çözüm Asistanı sekmesine aktarıldı.")
                        st.rerun()

    with tab_hint:
        col_l, col_r = st.columns([2, 1])
        with col_l:
            default_q = st.session_state.get("portal_question_draft", "")
            student_question = st.text_area(
                "Çözümünü İncelemek İstediğin Soru Metni",
                value=default_q,
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

        if st.button("🔍 Çözüm Rehberi Oluştur", type="primary", key="btn_solve_guide"):
            if not student_question.strip():
                st.warning("Lütfen önce bir soru metni girin.")
            else:
                sys_inst = (
                    "Sen sabırlı, pedagojik ve teşvik edici bir eğitim mentorusun. "
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
                    record_activity("Çözüm Asistanı", f"{hint_mode} üretildi")
                    st.markdown("#### 💡 Çözüm Analizi")
                    st.markdown(ans)
                except Exception as exc:
                    st.error(f"Çözüm üretilirken hata oluştu: {exc}")

# ==============================================================================
# 4. DİJİMENTOR — KURUMSAL AKILLI MENTORLUK STÜDYOSU
# ==============================================================================
elif current_page == "Dijimentor":
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">🤖 Dijimentor — Kurumsal Akıllı Sınav & Performans Mentorluğu</div>
            <div class="kp-studio-sub">
                Hedef sınavına, haftalık çalışma kapasitene ve çeldirici/zaman analizine göre
                ölçülebilir haftalık aksiyon planı ve deneme stratejisi üretir.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    current_hours = int(st.session_state.get("dk_hours", 16))
    daily_avg = current_hours / 7.0
    k_c1, k_c2, k_c3, k_c4 = st.columns(4)
    k_c1.markdown(
        f"""
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">Mentorluk Motoru</div>
            <div class="kp-kpi-value">Dijimentor v3</div>
            <div class="kp-kpi-sub">Kalan Plan: {int(st.session_state.remaining_coach_quota)} Hak</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    k_c2.markdown(
        f"""
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">Günlük Ortalama Tempo</div>
            <div class="kp-kpi-value">{daily_avg:.1f} Saat</div>
            <div class="kp-kpi-sub">Haftalık Toplam: {current_hours} Saat</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    k_c3.markdown(
        """
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">Analiz Boyutu</div>
            <div class="kp-kpi-value">3 Katman</div>
            <div class="kp-kpi-sub">Hız • Çeldirici • Net İvmesi</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    k_c4.markdown(
        """
        <div class="kp-kpi-card">
            <div class="kp-kpi-label">Çıktı Formatı</div>
            <div class="kp-kpi-value">Plan + DOCX</div>
            <div class="kp-kpi-sub">İndirilebilir Çalışma Rehberi</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")

    st.markdown("#### ⚡ Hızlı Mentorluk Senaryoları")
    st.caption(
        "Herhangi bir senaryo butonuna tıkladığında tüm profil verileri dolu şekilde yüklenir "
        "ve doğrudan Soru & Redaksiyon Fabrikası sayfasına yönlendirilirsin."
    )
    p_cols = st.columns(len(DIJIMENTOR_PRESETS) + 1)
    for idx, (preset_title, preset_data) in enumerate(DIJIMENTOR_PRESETS.items()):
        with p_cols[idx]:
            if st.button(f"⚡ {preset_title}", key=f"preset_{idx}", use_container_width=True):
                st.session_state["dk_exam"] = preset_data["exam"]
                st.session_state["dk_hours"] = preset_data["hours"]
                st.session_state["dk_goal"] = preset_data["goal"]
                st.session_state["dk_focus"] = preset_data["focus"]
                st.session_state["dk_weak"] = preset_data["weak"]
                st.session_state["dk_status"] = preset_data["status"]
                st.session_state["dk_active_preset"] = preset_title
                st.session_state["factory_focus_topic"] = preset_data["topic"]
                st.session_state["factory_branch"] = preset_data["branch"]
                st.session_state["factory_grade"] = preset_data["grade"]
                record_activity(
                    "Dijimentor Yönlendirme",
                    f"{preset_title} seçildi ve Soru Fabrikasına yönlendirildi",
                )
                navigate_to("Soru & Redaksiyon Fabrikası")
    with p_cols[-1]:
        if st.button("🧹 Planı Temizle", key="clear_dk_plan", use_container_width=True):
            st.session_state.dijimentor_plan = ""
            st.session_state["dk_active_preset"] = ""
            st.rerun()

    st.markdown("---")
    st.markdown("#### 🎯 Öğrenci Hedef ve Çalışma Profili")

    r1_a, r1_b, r1_c = st.columns(3)
    with r1_a:
        target_exam = st.selectbox(
            "Hedef Sınav / Kademe",
            ["2027 LGS", "2027 TYT - AYT", "5-7. Sınıf Akademik Gelişim"],
            index=["2027 LGS", "2027 TYT - AYT", "5-7. Sınıf Akademik Gelişim"].index(
                st.session_state.get("dk_exam", "2027 LGS")
            ),
            key="dk_exam_sel",
        )
        st.session_state["dk_exam"] = target_exam

        weekly_hours = st.slider(
            "Haftalık Çalışma Süresi (Saat)",
            min_value=4,
            max_value=45,
            value=current_hours,
            key="dk_hours_sld",
        )
        st.session_state["dk_hours"] = weekly_hours

    with r1_b:
        target_goal = st.text_input(
            "Hedef Puan / Net / Okul Bandı",
            value=st.session_state.get(
                "dk_goal", "Nitelikli lise / yüksek net ve sıfıra yakın dikkat hatası"
            ),
            key="dk_goal_inp",
        )
        st.session_state["dk_goal"] = target_goal

        coach_focus = st.selectbox(
            "Mentorluk Odak Alanı",
            [
                "Haftalık Dengeli Çalışma Programı",
                "Deneme Net Artırma & Zaman Yönetimi",
                "Dikkat ve Çeldirici Tuzağı Azaltma",
                "Sınav Kaygısı & Motivasyon Rutini",
            ],
            index=[
                "Haftalık Dengeli Çalışma Programı",
                "Deneme Net Artırma & Zaman Yönetimi",
                "Dikkat ve Çeldirici Tuzağı Azaltma",
                "Sınav Kaygısı & Motivasyon Rutini",
            ].index(st.session_state.get("dk_focus", "Haftalık Dengeli Çalışma Programı")),
            key="dk_focus_sel",
        )
        st.session_state["dk_focus"] = coach_focus

    with r1_c:
        weak_topics = st.text_area(
            "Zorlandığın Dersler / Kazanımlar",
            value=st.session_state.get(
                "dk_weak", "Paragrafta ana düşünce, üslü ifadeler, dikkat kaynaklı işlem hataları"
            ),
            height=82,
            key="dk_weak_area",
        )
        st.session_state["dk_weak"] = weak_topics

        recent_score = st.text_input(
            "Mevcut Deneme / Süre Durumu",
            value=st.session_state.get(
                "dk_status", "Orta seviye, süreyi yetiştirmede zorlanıyorum"
            ),
            key="dk_status_inp",
        )
        st.session_state["dk_status"] = recent_score

    q_hrs = round(weekly_hours * 0.45, 1)
    r_hrs = round(weekly_hours * 0.30, 1)
    a_hrs = round(weekly_hours * 0.25, 1)
    st.markdown(
        f"""
        <div style="margin-top:0.65rem;padding-top:0.75rem;border-top:1px solid #f1f5f9;">
            <div style="font-size:0.76rem;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:0.4px;">
                Önerilen Haftalık Kapasite Dağılımı ({weekly_hours} Saat)
            </div>
            <div class="kp-dist-bar">
                <div style="width:45%;background:#0f172a;" title="Soru & Muhakeme Pratiği"></div>
                <div style="width:30%;background:#881337;" title="Konu Tekrarı & Kazanım"></div>
                <div style="width:25%;background:#475569;" title="Deneme & Çeldirici Analizi"></div>
            </div>
            <div class="kp-dist-legend">
                <span>■ Soru & Muhakeme: <b>{q_hrs} sa</b> (%45)</span>
                <span>■ Konu Tekrarı: <b>{r_hrs} sa</b> (%30)</span>
                <span>■ Deneme & Analiz: <b>{a_hrs} sa</b> (%25)</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.write("")
    include_routine = st.checkbox(
        "Günlük mikro rutin, deneme analiz protokolü ve hafta sonu kontrol listesi dahil et",
        value=True,
    )

    if st.button("🎯 Dijimentor Planı Oluştur", type="primary"):
        quota_ok, quota_msg = consume_quota("coach", 1)
        if not quota_ok:
            st.error(quota_msg)
        else:
            sys_coach = (
                "Sen 'Dijimentor' adlı, Türkiye'deki öğrenciler için çalışan üst düzey, net, "
                "empatik ve veri odaklı bir dijital eğitim mentorusun. Yanıtını başlıklarla, "
                "madde işaretleriyle ve uygulanabilir zaman bloklarıyla kurumsal bir mentorluk "
                "raporu gibi düzenle. Şu bölümleri net oluştur:\n"
                "1. Teşhis ve Stratejik Özet\n"
                "2. Pazartesi–Pazar Haftalık Çalışma Planı (saat ve odak ders dağılımı)\n"
                "3. Çeldirici & Zaman Yönetimi Protokolü\n"
                "4. Haftalık Başarı KPI'ları ve Kontrol Listesi"
            )
            prompt_coach = (
                f"Hedef Sınav / Kademe: {target_exam}\n"
                f"Hedef Bandı: {target_goal}\n"
                f"Haftalık Çalışma Süresi: {weekly_hours} saat\n"
                f"Zorlanılan Dersler / Kazanımlar: {weak_topics}\n"
                f"Mentorluk Odak Alanı: {coach_focus}\n"
                f"Mevcut Deneme / Süre Durumu: {recent_score}\n"
                f"Detaylı Rutin & Analiz Protokolü: {'Evet' if include_routine else 'Özet'}\n\n"
                "Lütfen öğrenciye doğrudan uygulanabilir, motive edici ve net bir Türkçe plan sun."
            )
            try:
                with st.spinner("Dijimentor çalışma ve deneme stratejisi hazırlanıyor..."):
                    plan = call_gemini_text(prompt=prompt_coach, system_instruction=sys_coach)
                st.session_state.dijimentor_plan = plan
                record_activity("Dijimentor", f"{target_exam} • {weekly_hours} saatlik plan üretildi")
                st.success("Dijimentor planı başarıyla oluşturuldu.")
            except Exception as exc:
                st.session_state["remaining_coach_quota"] = int(
                    st.session_state.get("remaining_coach_quota", 0)
                ) + 1
                st.error(f"Dijimentor planı oluşturulurken hata oluştu: {exc}")

    if st.session_state.dijimentor_plan:
        st.markdown("---")
        st.markdown("### 📅 Kişiselleştirilmiş Dijimentor Mentorluk Raporu")
        st.markdown(st.session_state.dijimentor_plan)

        meta_summary = (
            f"Hedef: {target_exam} | Hedef Bandı: {target_goal} | "
            f"Haftalık Süre: {weekly_hours} Saat | Odak: {coach_focus}"
        )
        docx_coach = build_coach_docx(meta_summary, st.session_state.dijimentor_plan)
        txt_report = (
            f"{BRAND_NAME}\n"
            f"{meta_summary}\n"
            f"Tarih: {datetime.now().strftime('%d.%m.%Y')}\n"
            f"{'=' * 60}\n\n"
            f"{st.session_state.dijimentor_plan}"
        )

        d_col1, d_col2, d_col3 = st.columns(3)
        with d_col1:
            st.download_button(
                label="📥 Mentorluk Planını DOCX İndir",
                data=docx_coach,
                file_name="dijimentor_calisma_plani.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with d_col2:
            st.download_button(
                label="📄 Planı Metin (.txt) İndir",
                data=txt_report.encode("utf-8"),
                file_name="dijimentor_calisma_plani.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with d_col3:
            if st.button("⚙️ Zayıf Konulardan Soru Üret →", use_container_width=True):
                st.session_state["factory_focus_topic"] = weak_topics.split(",")[0].strip()
                navigate_to("Soru & Redaksiyon Fabrikası")

# ==============================================================================
# 5. VELİ TAKİP PANELİ + PERFORMANS GRAFİKLERİ
# ==============================================================================
elif current_page == "Veli Takip Paneli":
    st.subheader("👨‍👩‍👧 Veli Takip Paneli")
    st.write("Öğrencinin haftalık çalışma düzeni, soru çözümü ve kazanım ilerleme özeti.")

    v1, v2, v3 = st.columns(3)
    v1.metric("Bu Hafta Çözülen Soru", "340 Soru", "+18% geçen haftaya göre")
    v2.metric("Ortalama Doğruluk", "%81", "+4 puan")
    v3.metric("Dijimentor Plan Uyumu", "%88", "6 / 7 gün aktif")

    st.markdown("#### Kazanım Bazlı Görünüm")
    st.progress(0.84, text="Okuma ve Anlama — %84")
    st.progress(0.72, text="Sayısal Muhakeme ve Problem Çözme — %72")
    st.progress(0.90, text="Düzenli Tekrar ve Ödev Tamamlama — %90")

    g_col1, g_col2 = st.columns(2)
    with g_col1:
        st.markdown("#### 📈 Haftalık Doğruluk İvmesi (%)")
        st.area_chart(
            {"1. Hafta": 71, "2. Hafta": 74, "3. Hafta": 77, "4. Hafta": 81},
            height=190,
        )
    with g_col2:
        st.markdown("#### 📊 Ders Bazlı Başarı Oranı (%)")
        st.bar_chart(
            {"Türkçe": 84, "Matematik": 72, "Fen": 82, "Sosyal": 86, "İngilizce": 79},
            height=190,
        )

# ==============================================================================
# 6. B2B LİSANS & KURUM + ZÜMRE SORU DAĞITIMI & RAPORLAMA SİMÜLASYONU
# ==============================================================================
elif current_page == "B2B Lisans & Kurum":
    st.subheader("🏫 B2B Lisans & Kurum Yönetimi")
    st.write(
        "Okullar, kurs merkezleri ve eğitim kurumları için toplu soru üretimi, zümre "
        "kapasite dağıtımı ve kurumsal lisanslama çözümleri."
    )

    tab_b2b_form, tab_b2b_zumre = st.tabs(
        ["📋 Kurumsal Lisans Teklif Stüdyosu", "📊 Zümre Soru Dağıtımı & Kurumsal Rapor"]
    )

    with tab_b2b_form:
        with st.form("b2b_inquiry_form"):
            c1, c2 = st.columns(2)
            with c1:
                org_name = st.text_input("Kurum / Okul Adı", value="Özel Kurşun Kalem Akademi")
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
                    st.session_state["last_b2b_summary"] = (
                        f"Kurum: {org_name} ({city}) | Yetkili: {contact_name or '-'} | "
                        f"E-posta: {contact_email or '-'} | Koltuk: {teacher_seats} Öğretmen, "
                        f"{student_seats} Öğrenci"
                    )
                    st.session_state["b2b_org_cache"] = {
                        "org": org_name,
                        "teachers": int(teacher_seats),
                        "students": int(student_seats),
                    }
                    record_activity("B2B Lisans", f"{org_name} teklif taslağı hazırlandı")
                    st.success(
                        f"**{org_name}** ({city}) için **{teacher_seats} öğretmen** ve "
                        f"**{student_seats} öğrenci** kapasiteli kurumsal lisans talebi kaydedildi."
                    )

        if st.session_state.get("last_b2b_summary"):
            b2b_tier = get_tier_by_id("b2b_kurs" if student_seats <= 300 else "b2b_kampus")
            b2b_summary = {
                "name": b2b_tier["name"],
                "total_count": 1,
                "monthly_price": b2b_tier["monthly_price"],
                "monthly_quota": b2b_tier["monthly_quota"],
                "coach_quota": b2b_tier["coach_quota"],
                "entries": [
                    {
                        "tier": b2b_tier,
                        "qty": 1,
                        "line_price": b2b_tier["monthly_price"],
                        "line_quota": b2b_tier["monthly_quota"],
                        "line_coach": b2b_tier["coach_quota"],
                    }
                ],
            }
            b2b_docx = build_commercial_brief_docx(
                title=f"{org_name or 'Kurumsal'} Lisans Teklif Özeti",
                summary=b2b_summary,
                billing_cycle="Yıllık (%20 İndirim)",
                installment=6,
                total_amount=round(b2b_tier["monthly_price"] * 12 * 0.80),
                extra_notes=st.session_state["last_b2b_summary"],
            )
            st.download_button(
                label="📥 Kurumsal Lisans Teklifini DOCX İndir",
                data=b2b_docx,
                file_name="kursun_kalem_kurumsal_teklif.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    with tab_b2b_zumre:
        org_cache = st.session_state.get(
            "b2b_org_cache", {"org": "Özel Kurşun Kalem Akademi", "teachers": 15, "students": 250}
        )
        recommended_tier = get_tier_by_id(
            "b2b_kurs" if org_cache["students"] <= 300 else "b2b_kampus"
        )
        total_q = recommended_tier["monthly_quota"]

        bk1, bk2, bk3, bk4 = st.columns(4)
        bk1.metric("Kurum Profili", org_cache["org"])
        bk2.metric("Aylık Soru Havuzu", f"{total_q:,}".replace(",", "."))
        bk3.metric("Öğretmen Başına Kota", f"{total_q // max(1, org_cache['teachers']):,} Soru".replace(",", "."))
        bk4.metric("Öğrenci Başına Test", f"{total_q // max(1, org_cache['students'])} Soru / Ay")

        st.markdown("#### 📐 Branş Zümreleri Arası Soru Kotası Dağılımı")
        dist_ratios = {
            "Türkçe & Edebiyat Zümresi": 0.28,
            "Matematik & Geometri Zümresi": 0.30,
            "Fen Bilimleri Zümresi": 0.22,
            "Sosyal Bilimler & İnkılap Zümresi": 0.12,
            "Yabancı Dil (İngilizce) Zümresi": 0.08,
        }
        zumre_rows = []
        chart_map: Dict[str, int] = {}
        for z_name, ratio in dist_ratios.items():
            allocated = int(round(total_q * ratio))
            short_lbl = z_name.split(" ")[0]
            chart_map[short_lbl] = allocated
            zumre_rows.append(
                {
                    "Zümre Birimi": z_name,
                    "Pay (%)": f"%{int(ratio * 100)}",
                    "Aylık Soru Kotası": f"{allocated:,}".replace(",", "."),
                    "Haftalık Deneme Hedefi": f"{max(1, allocated // 80)} Kitapçık",
                    "Redaksiyon Standardı": "Maarif Modeli • Çift Kontrol",
                }
            )

        zc1, zc2 = st.columns([3, 2])
        with zc1:
            st.dataframe(zumre_rows, use_container_width=True, hide_index=True)
        with zc2:
            st.bar_chart(chart_map, height=225)

# ==============================================================================
# 7. PAKETLER & ABONELİK VİTRİNİ (OKUL SIRT ÇANTASI SEPETİ, 12 PAKET & IYZICO)
# ==============================================================================
elif current_page == "Paketler & Abonelik Vitrini":
    st.subheader("💎 Paketler & Abonelik Vitrini (12 Paket)")
    st.write(
        "Seçtiğiniz paketler alışveriş sepeti gibi **Okul Sırt Çantası** görselinin içine eklenir; "
        "birden fazla paket toplayabilir, adet ayarlayabilir ve iyzico taksitlerini inceleyebilirsiniz."
    )

    current_bag = get_bag_items()
    active_bag_summary = render_school_backpack_panel()

    st.markdown("---")
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
                in_bag_qty = current_bag.get(tier["id"], 0)
                bag_tag = f" • 🎒 Çantada: {in_bag_qty}" if in_bag_qty > 0 else ""
                st.markdown(
                    f"""
                    <div class="kp-card">
                        <div style="font-size:0.78rem;color:#64748b;font-weight:600;">
                            {tier["segment"].upper()} • {tier["badge"]}{bag_tag}
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
                st.button(
                    "🎒 Çantaya Ekle",
                    key=f"select_{tier['id']}",
                    use_container_width=True,
                    on_click=cb_add_tier_to_bag,
                    args=(tier["id"], 1),
                )

    st.markdown("---")
    st.markdown("### 📊 12 Paket Karşılaştırma Matrisi")
    matrix_rows = [
        {
            "Paket Adı": t["name"],
            "Segment": t["segment"],
            "Etiket": t["badge"],
            "Çantadaki Adet": current_bag.get(t["id"], 0),
            "Aylık Soru Kotası": f"{t['monthly_quota']:,}".replace(",", "."),
            "Aylık Ücret (₺)": f"{t['monthly_price']:,}".replace(",", "."),
            "Öne Çıkan Kapsam": t["features"][0],
        }
        for t in SUBSCRIPTION_TIERS
    ]
    st.dataframe(matrix_rows, use_container_width=True, hide_index=True)

    st.markdown("---")
    render_iyzico_simulator(
        package_name=active_bag_summary["name"],
        monthly_price=active_bag_summary["monthly_price"],
        summary=active_bag_summary,
    )
