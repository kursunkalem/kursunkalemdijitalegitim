import base64
import html
import io
import json
import os
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import iyzipay
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
    page_title="KURŞUN KALEM DİJİTAL OKUL",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "KURŞUN KALEM DİJİTAL OKUL"
BRAND_SHORT = "Kurşun Kalem"
BRAND_TAGLINE = "Dijital Eğitim, Ölçme-Değerlendirme ve Akıllı Mentorluk Platformu"
DEFAULT_MODEL = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-2.5-flash"
STATE_FILE_PATH = Path(".kursun_kalem_state.json")

# Kalıcı Gömülü Kurumsal Logo
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
  <text x="256" y="314" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="48" font-weight="800" fill="#0b1836" letter-spacing="1">DİJİTAL OKUL</text>
  <line x1="92" y1="344" x2="152" y2="344" stroke="#64748b" stroke-width="2.5"/>
  <text x="256" y="351" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="22" font-weight="700" fill="#0f172a" letter-spacing="5">KURŞUN KALEM</text>
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

# Okul Sırt Çantası Görseli (SVG)
SCHOOL_BACKPACK_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 260 300" width="260" height="300">
  <defs>
    <linearGradient id="bagBody" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#172c56"/>
      <stop offset="100%" stop-color="#0f172a"/>
    </linearGradient>
    <linearGradient id="bagPocket" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="#334155"/>
      <stop offset="100%" stop-color="#1e293b"/>
    </linearGradient>
  </defs>
  <path d="M 98 56 Q 130 20 162 56" fill="none" stroke="#0f172a" stroke-width="10" stroke-linecap="round"/>
  <path d="M 88 64 C 58 95, 54 165, 66 245" fill="none" stroke="#334155" stroke-width="15" stroke-linecap="round"/>
  <path d="M 172 64 C 202 95, 206 165, 194 245" fill="none" stroke="#334155" stroke-width="15" stroke-linecap="round"/>
  <rect x="64" y="56" width="132" height="205" rx="32" fill="url(#bagBody)" stroke="#0f172a" stroke-width="3"/>
  <path d="M 68 92 Q 130 114 192 92 L 192 78 Q 130 56 68 78 Z" fill="#d4af37" opacity="0.28"/>
  <rect x="82" y="138" width="96" height="96" rx="20" fill="url(#bagPocket)" stroke="#d4af37" stroke-width="2"/>
  <circle cx="130" cy="186" r="22" fill="#ffffff" stroke="#0f172a" stroke-width="2"/>
  <text x="130" y="184" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="800" fill="#0f172a">ÇANTA</text>
  <text x="130" y="196" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="7.5" font-weight="700" fill="#475569">PAKETLER</text>
  <rect x="72" y="242" width="116" height="18" rx="9" fill="#090d16"/>
</svg>"""


def svg_to_data_uri(svg_str: str) -> str:
    """SVG metnini güvenli Data URI formatına çevirir."""
    encoded = base64.b64encode(svg_str.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def get_embedded_logo_png_bytes() -> Optional[bytes]:
    """Gömülü logoyu DOCX belgesine eklemek için PNG baytı üretir."""
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
        draw.text((160, 206), "— KURŞUN KALEM —", fill=(15, 23, 42, 255), anchor="mm")
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
    "Ana Sayfa",
    "Soru & Redaksiyon Fabrikası",
    "Öğrenci Portalı & Çözüm",
    "Dijimentor",
    "Veli Takip Paneli",
    "Öğrenci Röntgeni",
    "B2B Lisans & Kurum",
    "Paketler & Abonelik Vitrini",
]

PAGE_DESCRIPTIONS: Dict[str, Tuple[str, str]] = {
    "Ana Sayfa": (
        "Genel Bakış & Sınav Telemetrisi",
        "2027 LGS ve YKS geri sayımı, platform metrikleri, modül erişimi ve haftalık üretim akışı.",
    ),
    "Soru & Redaksiyon Fabrikası": (
        "Maarif Modeli Soru & Kalite Stüdyosu",
        "Kazanım, çeldirici ve soru kökü kalitesi odaklı özgün testler üretin, redaksiyon denetleyin ve DOCX olarak indirin.",
    ),
    "Öğrenci Portalı & Çözüm": (
        "İnteraktif Mini Deneme & Çözüm Rehberi",
        "Testleri şık işaretleyerek çözün, netinizi hesaplayın ve yanlış veya zorlandığınız sorularda adımlı rehber alın.",
    ),
    "Dijimentor": (
        "Akıllı Sınav & Performans Mentorluğu",
        "Hedef sınava, zayıf kazanımlara ve haftalık kapasiteye göre ölçülebilir çalışma ve deneme planı oluşturun.",
    ),
    "Veli Takip Paneli": (
        "Veli Gelişim & Bilgilendirme Merkezi",
        "Öğrencinin haftalık soru temposunu, kazanım başarılarını izleyin ve hazır veli mesajları üretin.",
    ),
    "Öğrenci Röntgeni": (
        "Öğrenci Röntgeni & Akademik Check-Up",
        "Kazanım bazlı röntgen taraması, hata kök neden analizi ve Dijimentor telafi reçetesi ile öğrencinin akademik röntgenini çekin.",
    ),
    "B2B Lisans & Kurum": (
        "Kurumsal Lisans & Zümre Yönetimi",
        "Okul ve kurs merkezleri için koltuk/kota planlaması, zümre dağılımı ve DOCX teklif çıktısı.",
    ),
    "Paketler & Abonelik Vitrini": (
        "Abonelik Paketleri & Okul Sırt Çantası",
        "Öğrenci, YKS, Veli ve Kurumsal paketleri çantanıza ekleyin, kotaları yönetin ve iyzico planını inceleyin.",
    ),
}

OGRENCI_RONTGEN_PROFILLERI: Dict[str, Dict[str, Any]] = {
    "Elif Öztürk (8-A | 2027 LGS)": {
        "basari_endeksi": "%76.4",
        "basari_delta": "+4.2% (Son 2 hafta)",
        "kritik_konu": "3 Kazanım",
        "kritik_delta": "-2 İyileşme",
        "hiz": "1.35 dk/soru",
        "hiz_delta": "İdeal Tempo",
        "hedef_puan": "454 / 480",
        "hedef_delta": "26 Puan Fark",
        "kazanimlar": [
            {
                "Ders / Alan": "Türkçe",
                "Kazanım / Konu": "Sözel Mantık & Muhakeme",
                "Çözülen": 58,
                "Başarı (%)": 46,
                "Teşhis Durumu": "🔴 Kritik Müdahale",
            },
            {
                "Ders / Alan": "Türkçe",
                "Kazanım / Konu": "Paragrafta Ana Düşünce",
                "Çözülen": 135,
                "Başarı (%)": 86,
                "Teşhis Durumu": "🟢 Tam Hakimiyet",
            },
            {
                "Ders / Alan": "Matematik",
                "Kazanım / Konu": "Sayısal Muhakeme & Cebirsel İfadeler",
                "Çözülen": 74,
                "Başarı (%)": 52,
                "Teşhis Durumu": "🔴 Kritik Müdahale",
            },
            {
                "Ders / Alan": "Fen Bilimleri",
                "Kazanım / Konu": "Basınç & DNA Genetik Kod",
                "Çözülen": 62,
                "Başarı (%)": 74,
                "Teşhis Durumu": "🟡 Geliştirilmeli",
            },
            {
                "Ders / Alan": "Sosyal / İnkılap",
                "Kazanım / Konu": "Millî Uyanış & Çağdaş Türkiye",
                "Çözülen": 48,
                "Başarı (%)": 88,
                "Teşhis Durumu": "🟢 Tam Hakimiyet",
            },
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 36,
            "Uzun Metin / Odak Kaybı": 30,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 22,
            "Süre Yetişmemesi": 12,
        },
        "recete_ozet": (
            "Sözel Mantık & Muhakeme (%46) ve Sayısal Muhakeme (%52) kazanımlarında tablo/şema kurma ve "
            "çok adımlı yönerge takibi zayıf. Yanlışların %22'si soru kökündeki olumsuz ifadelerin hızlı okunmasından kaynaklanıyor."
        ),
        "recete_aksiyon": "Soru & Redaksiyon Fabrikası üzerinden orta-zor seviyede 15 Sözel Mantık ve 15 Sayısal Muhakeme sorusu üretilip atanmaya hazır.",
        "hedef_konu": "Sözel Mantık & Muhakeme",
        "hedef_brans": "Türkçe",
        "hedef_sinif": "8. Sınıf (LGS)",
    },
    "Zeynep Kaya (8-B | 2027 LGS)": {
        "basari_endeksi": "%83.8",
        "basari_delta": "+5.6% (Son 2 hafta)",
        "kritik_konu": "2 Kazanım",
        "kritik_delta": "-1 İyileşme",
        "hiz": "1.18 dk/soru",
        "hiz_delta": "Hızlı ve Dengeli",
        "hedef_puan": "468 / 480",
        "hedef_delta": "12 Puan Fark",
        "kazanimlar": [
            {
                "Ders / Alan": "Türkçe",
                "Kazanım / Konu": "Metinler Arası Karşılaştırma",
                "Çözülen": 64,
                "Başarı (%)": 81,
                "Teşhis Durumu": "🟢 Tam Hakimiyet",
            },
            {
                "Ders / Alan": "Matematik",
                "Kazanım / Konu": "Kareköklü İfadeler & Olasılık",
                "Çözülen": 72,
                "Başarı (%)": 55,
                "Teşhis Durumu": "🔴 Kritik Müdahale",
            },
            {
                "Ders / Alan": "Fen Bilimleri",
                "Kazanım / Konu": "Mevsimler & İklim Hareketleri",
                "Çözülen": 54,
                "Başarı (%)": 89,
                "Teşhis Durumu": "🟢 Tam Hakimiyet",
            },
            {
                "Ders / Alan": "Türkçe",
                "Kazanım / Konu": "Fiilimsiler & Cümle Ögeleri",
                "Çözülen": 86,
                "Başarı (%)": 71,
                "Teşhis Durumu": "🟡 Geliştirilmeli",
            },
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 27,
            "Uzun Metin / Odak Kaybı": 21,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 35,
            "Süre Yetişmemesi": 17,
        },
        "recete_ozet": (
            "Kareköklü İfadeler (%55) alanında yaklaşık değer ve yeni nesil kurgu adımlarında sapma görülüyor. "
            "Hataların %35'i soru kökündeki sınırlandırıcı ifadelerin gözden kaçmasından kaynaklanıyor."
        ),
        "recete_aksiyon": "Olumsuz soru köklü 20 Kareköklü İfade ve dikkat odaklı muhakeme seti oluşturularak Öğrenci Portalı'na gönderilmeye hazır.",
        "hedef_konu": "Üslü ve kareköklü ifadelerde akıl yürütme ve tahmin",
        "hedef_brans": "Matematik",
        "hedef_sinif": "8. Sınıf (LGS)",
    },
    "Emir Demir (12-A | 2027 YKS)": {
        "basari_endeksi": "%71.2",
        "basari_delta": "+3.8% (Son 2 hafta)",
        "kritik_konu": "4 Kazanım",
        "kritik_delta": "-1 İyileşme",
        "hiz": "1.52 dk/soru",
        "hiz_delta": "Süre Baskısı Var",
        "hedef_puan": "442 / 500",
        "hedef_delta": "58 Puan Fark",
        "kazanimlar": [
            {
                "Ders / Alan": "Türkçe",
                "Kazanım / Konu": "Paragrafta Çıkarım & Yorum",
                "Çözülen": 145,
                "Başarı (%)": 66,
                "Teşhis Durumu": "🟡 Geliştirilmeli",
            },
            {
                "Ders / Alan": "Türkçe",
                "Kazanım / Konu": "Dil Bilgisi & Anlatım Bozukluğu",
                "Çözülen": 78,
                "Başarı (%)": 49,
                "Teşhis Durumu": "🔴 Kritik Müdahale",
            },
            {
                "Ders / Alan": "Matematik",
                "Kazanım / Konu": "Geometri & Veri Analizi",
                "Çözülen": 88,
                "Başarı (%)": 54,
                "Teşhis Durumu": "🔴 Kritik Müdahale",
            },
            {
                "Ders / Alan": "Fen Bilimleri",
                "Kazanım / Konu": "Basit Makineler & Enerji",
                "Çözülen": 56,
                "Başarı (%)": 84,
                "Teşhis Durumu": "🟢 Tam Hakimiyet",
            },
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 34,
            "Uzun Metin / Odak Kaybı": 24,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 16,
            "Süre Yetişmemesi": 26,
        },
        "recete_ozet": (
            "Dil Bilgisi (%49) ve Geometri/Veri Analizi (%54) alanlarında kural transferi ve süre yönetimi kritik öncelikte. "
            "Süre yetişmemesi (%26) denemenin ikinci yarısında dikkat düşüşüyle birleşiyor."
        ),
        "recete_aksiyon": "Zamanlı 15 Dil Bilgisi + 15 Analitik Muhakeme sorusundan oluşan hız ve doğruluk telafi paketi hazır.",
        "hedef_konu": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
        "hedef_brans": "Türkçe",
        "hedef_sinif": "11-12. Sınıf (YKS)",
    },
}

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
        "monthly_price": 1990,
        "monthly_quota": 400,
        "coach_quota": 30,
        "badge": "Temel",
        "desc": "5–8. sınıf öğrencileri için kazanım tekrarları, çeldirici farkındalığı ve adımlı çözüm desteği.",
        "features": [
            "Aylık 400 Maarif Modeli & LGS soru üretimi",
            "Öğrenci Portalı mini deneme ve ipucu/çözüm asistanı",
            "2027 LGS canlı geri sayım ve haftalık hedef takibi",
            "Çeldirici analizli standart DOCX test çıktısı",
            "Temel hata defteri ve kazanım özeti",
            "Mobil & masaüstü uyumlu kesintisiz çalışma arayüzü",
        ],
    },
    {
        "id": "ogrenci_pro",
        "segment": "Öğrenci",
        "name": "2027 LGS Hedef Pro",
        "monthly_price": 3290,
        "monthly_quota": 1500,
        "coach_quota": 90,
        "badge": "Popüler",
        "desc": "Yeni nesil muhakeme soruları, otomatik redaksiyon, Dijimentor planı ve veli raporu.",
        "features": [
            "Aylık 1.500 yapay zekâ soru ve redaksiyon hakkı",
            "Otomatik redaktör ile çeldirici ve soru kökü denetimi",
            "Dijimentor haftalık akıllı çalışma ve tempo programı",
            "Veli Takip Paneli canlı bildirim ve özet erişimi",
            "Soru köküne uygun başlıklandırılmış DOCX test çıktısı",
            "Yanlış soruları otomatik çözüm rehberine aktarma",
        ],
    },
    {
        "id": "ogrenci_elit",
        "segment": "Öğrenci",
        "name": "LGS Tam Donanım Elit",
        "monthly_price": 4890,
        "monthly_quota": 3500,
        "coach_quota": 200,
        "badge": "Kapsamlı",
        "desc": "Yüksek soru üretim kotası, yoğun kalite denetimi, çeldirici raporu ve özel mentorluk senaryoları.",
        "features": [
            "Aylık 3.500 soru üretimi + çift katmanlı redaksiyon",
            "Dijimentor günlük deneme analizi & çeldirici tuzağı raporu",
            "Veli Takip Paneli + haftalık veli bilgilendirme merkezi",
            "Özel kurum/öğrenci başlıklı düzenlenebilir DOCX testler",
            "Zayıf kazanımlardan anında telafi testi üretimi",
            "Süre yönetimi ve dikkat kaybı azaltma protokolü",
        ],
    },
    {
        "id": "ogrenci_kamp",
        "segment": "Öğrenci",
        "name": "LGS Final & Deneme Kampı",
        "monthly_price": 5690,
        "monthly_quota": 5000,
        "coach_quota": 260,
        "badge": "Kamp",
        "desc": "Sınava hazırlık maratonunda yüksek hacimli deneme üretimi, hız antrenmanı ve net ivme planı.",
        "features": [
            "Aylık 5.000 yeni nesil deneme ve kazanım sorusu",
            "Yoğun süre yönetimi ve son tur hızlandırma planları",
            "Dijimentor haftalık net artırma ve kritik kazanım takibi",
            "Ayrıntılı doğru/yanlış/net tablosu ve hata analizi",
            "Baskıya hazır DOCX deneme ve çalışma fasikülleri",
            "Veliye haftalık kamp ilerleme raporu paylaşımı",
        ],
    },
    {
        "id": "yks_sayisal",
        "segment": "YKS",
        "name": "YKS MF (Sayısal) Akademi",
        "monthly_price": 3690,
        "monthly_quota": 1800,
        "coach_quota": 110,
        "badge": "TYT + AYT",
        "desc": "TYT ve AYT Matematik, Fizik, Kimya, Biyoloji ve Türkçe odaklı ileri hazırlık paketi.",
        "features": [
            "Aylık 1.800 TYT & AYT kazanım ve analitik soru üretimi",
            "Dijimentor net artırma, süre bölümü ve tekrar döngüsü",
            "Adımlı çözüm asistanı ile mantık ve işlem adımı analizi",
            "Konu bazlı zorluk kademeli test ve deneme üretimi",
            "Düzenlenebilir DOCX soru ve çözüm çıktısı",
            "Haftalık çalışma kapasitesi (%45/%30/%25) optimizasyonu",
        ],
    },
    {
        "id": "yks_ea_sozel",
        "segment": "YKS",
        "name": "YKS EA & Sözel Atölye",
        "monthly_price": 3490,
        "monthly_quota": 1800,
        "coach_quota": 110,
        "badge": "TYT + AYT",
        "desc": "Türk Dili ve Edebiyatı, Tarih, Coğrafya, Matematik ve paragraf yorum-analiz paketi.",
        "features": [
            "Aylık 1.800 paragraf, edebiyat ve sosyal bilimler sorusu",
            "Çeldirici şık ayıklama ve metin analizi redaksiyonu",
            "Dijimentor okuma hızı ve deneme stratejisi mentorluğu",
            "Mini deneme modülü ile anlık net ve doğruluk hesabı",
            "Düzenlenebilir DOCX çalışma kağıdı ve test çıktısı",
            "Eksik kazanım odaklı haftalık tekrar planı",
        ],
    },
    {
        "id": "yks_vip",
        "segment": "YKS",
        "name": "YKS Derece VIP",
        "monthly_price": 5490,
        "monthly_quota": 4500,
        "coach_quota": 280,
        "badge": "VIP",
        "desc": "Derece hedefleyen öğrenciler için yüksek soru kotası, derin analiz ve özel mentorluk.",
        "features": [
            "Aylık 4.500 soru üretimi ve kalite redaksiyonu",
            "Dijimentor hedef derece simülasyonu ve haftalık rapor",
            "Zor (LGS/ÖSYM) seviye muhakeme ve çeldirici setleri",
            "Yanlış sorulardan kişiye özel pekiştirme akışı",
            "Kurumsal standartta DOCX test ve plan çıktıları",
            "Öncelikli kapsam ve yoğun deneme hazırlık desteği",
        ],
    },
    {
        "id": "yks_tyt_sprint",
        "segment": "YKS",
        "name": "TYT Hız & Paragraf Sprinti",
        "monthly_price": 3190,
        "monthly_quota": 1400,
        "coach_quota": 90,
        "badge": "TYT Odaklı",
        "desc": "TYT Türkçe ve temel matematikte okuma hızı, dikkat ve süre baskısını aşmaya yönelik paket.",
        "features": [
            "Aylık 1.400 TYT paragraf ve akıl yürütme sorusu",
            "Dikkat ve çeldirici tuzağı azaltma mentorluk planı",
            "İnteraktif mini deneme ile hız ve net takibi",
            "Zorluk seviyeli (Kolay–Zor) kademeli soru setleri",
            "DOCX formatında pratik test ve değerlendirme çıktısı",
            "Haftalık tempo ve dikkat KPI kontrol listesi",
        ],
    },
    {
        "id": "veli_aile",
        "segment": "Veli",
        "name": "Aile Takip & Rehberlik",
        "monthly_price": 2190,
        "monthly_quota": 600,
        "coach_quota": 60,
        "badge": "Veli Odaklı",
        "desc": "Öğrencinin çalışma düzenini, deneme ivmesini ve eksik kazanımlarını şeffaf izleyin.",
        "features": [
            "Veli Takip Paneli haftalık performans ve net özeti",
            "4 kategorili Veli Bilgilendirme & Otomatik Mesaj Merkezi",
            "WhatsApp, SMS ve E-posta için hazır veli mesaj şablonları",
            "Dijimentor çalışma disiplini ve motivasyon planı",
            "Aylık 600 soru üretimi ve 1 öğrenci profili dahil",
            "Veli mesajlarını ve raporları DOCX olarak indirme",
        ],
    },
    {
        "id": "veli_coklu",
        "segment": "Veli",
        "name": "Çoklu Kardeş Aile Paketi",
        "monthly_price": 3690,
        "monthly_quota": 2200,
        "coach_quota": 130,
        "badge": "2 Öğrenci",
        "desc": "İki kardeş için ayrı sınıf/sınav seviyelerinde takip, soru üretimi ve Dijimentor desteği.",
        "features": [
            "2 ayrı öğrenci profili (5–12. sınıf / LGS & YKS)",
            "Aylık 2.200 ortak soru üretim ve redaksiyon kotası",
            "Veli Takip Paneli karşılaştırmalı gelişim görünümü",
            "Otomatik veli bilgilendirme mesajları ve duyuru akışı",
            "İki kademe için ayrı DOCX test ve plan çıktısı",
            "Aile bütçesinde avantajlı çoklu paket kotası",
        ],
    },
    {
        "id": "veli_rehber_plus",
        "segment": "Veli",
        "name": "Aile Rehberlik & Deneme Plus",
        "monthly_price": 4390,
        "monthly_quota": 3000,
        "coach_quota": 160,
        "badge": "Plus",
        "desc": "Sınav yılındaki öğrencinin deneme, kazanım ve çalışma disiplini takibini üst düzeye taşıyan paket.",
        "features": [
            "Aylık 3.000 soru üretimi + otomatik redaksiyon",
            "Dijimentor haftalık çalışma ve sınav kaygısı desteği",
            "Veli Takip Paneli ders ve kazanım bazlı grafikler",
            "Düzenli veli bilgilendirme şablonları ve DOCX dışa aktarım",
            "Eksik kazanım telafi takibi ve özet bildirimleri",
            "Yüksek mentorluk kotası ile kesintisiz planlama",
        ],
    },
    {
        "id": "b2b_ogretmen",
        "segment": "Kurum / B2B",
        "name": "Öğretmen & Zümre Stüdyosu",
        "monthly_price": 7900,
        "monthly_quota": 6000,
        "coach_quota": 220,
        "badge": "Zümre",
        "desc": "Branş öğretmenleri ve zümre başkanları için kazanım testi, redaksiyon ve zümre üretim aracı.",
        "features": [
            "Aylık 6.000 soru üretimi + otomatik redaksiyon denetimi",
            "Türkiye Yüzyılı Maarif Modeli kazanım ve tema uyumu",
            "Kurum başlıklı düzenlenebilir DOCX sınav ve test çıktısı",
            "5 öğretmen/editör koltuğu ve zümre paylaşım akışı",
            "Çeldirici kalitesi ve dil/yazım doğruluğu kontrolü",
            "Hızlı kazanım tarama ve ölçme-değerlendirme setleri",
        ],
    },
    {
        "id": "b2b_kurs",
        "segment": "Kurum / B2B",
        "name": "Etüt Merkezi & Kurs Paketi",
        "monthly_price": 19900,
        "monthly_quota": 18000,
        "coach_quota": 600,
        "badge": "Kurumsal",
        "desc": "Kurs merkezleri ve butik okullar için ölçeklenebilir soru fabrikası, portal ve zümre yönetimi.",
        "features": [
            "Aylık 18.000 soru üretimi ve toplu test hazırlama",
            "Öğrenci Portalı, Dijimentor ve Veli Paneli kurumsal erişimi",
            "Zümre soru kotası dağıtımı ve kurum içi raporlama",
            "Kurum logolu/başlıklı DOCX sınav ve özet çıktıları",
            "25 öğretmen + 250 öğrenci hesabı kapasitesi",
            "Veli bilgilendirme şablonları ile kurumsal iletişim",
        ],
    },
    {
        "id": "b2b_kampus",
        "segment": "Kurum / B2B",
        "name": "Okul & Kampüs Lisansı",
        "monthly_price": 39900,
        "monthly_quota": 50000,
        "coach_quota": 1500,
        "badge": "Kampüs",
        "desc": "Tüm kademelerde standart ölçme-değerlendirme, zümre havuzu ve akıllı mentorluk altyapısı.",
        "features": [
            "Aylık 50.000 soru üretimi + öncelikli redaksiyon",
            "Tüm modüllerde kurumsal yetkilendirme ve takip",
            "Branş zümreleri arası otomatik kota ve deneme planı",
            "İdareci ve zümre performans/kapasite raporları",
            "Sınırsız DOCX test/teklif çıktısı ve kurumsal şablon",
            "Kampüs geneli standart sınav ve kazanım uyumu",
        ],
    },
    {
        "id": "b2b_ozel",
        "segment": "Kurum / B2B",
        "name": "Kurumsal Zincir & Yayınevi",
        "monthly_price": 74900,
        "monthly_quota": 140000,
        "coach_quota": 3500,
        "badge": "Enterprise",
        "desc": "Çok şubeli eğitim kurumları ve içerik ekipleri için yüksek kapasite ve özel editörlük.",
        "features": [
            "Aylık 140.000 yüksek hacimli soru üretim kotası",
            "Çok şubeli kampüs ve zümre kapasite yönetimi",
            "Sıkı redaksiyon, çeldirici ve kazanım kalite standardı",
            "Kurumsal teklif, özet ve sınav DOCX otomasyonu",
            "Yüksek hacimli Dijimentor mentorluk ve analiz kotası",
            "Kurumsal onboarding ve özel yapılandırma desteği",
        ],
    },
    {
        "id": "b2b_yayinevi_pro",
        "segment": "Kurum / B2B",
        "name": "Akademik İçerik & Soru Bankası Pro",
        "monthly_price": 98900,
        "monthly_quota": 220000,
        "coach_quota": 5000,
        "badge": "Ultra B2B",
        "desc": "Yayınevleri ve büyük eğitim grupları için tam kapasiteli soru fabrikası ve ölçme stüdyosu.",
        "features": [
            "Aylık 220.000 soru üretim ve redaksiyon kapasitesi",
            "Tüm sınıf kademeleri ve sınav türlerinde çoklu üretim",
            "Gelişmiş zümre dağılımı ve kurumsal raporlama",
            "5.000 Dijimentor plan kotası ile geniş mentorluk ağı",
            "Standartlaştırılmış DOCX kitapçık ve rapor çıktıları",
            "Kurumsal ölçekte kesintisiz içerik üretim altyapısı",
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

PARENT_MESSAGE_TOPICS: Dict[str, Dict[str, Any]] = {
    "1. Akademik İlerleme ve Başarı Takibi": {
        "subtopics": [
            "Haftalık Soru ve Net Özeti",
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi",
            "Deneme Sınavı ve LGS Değerlendirmesi",
        ],
        "templates": {
            "Haftalık Soru ve Net Özeti": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Okul platformunda öğrencimiz {student_name} ({grade}) "
                "bu hafta toplam {weekly_solved} soru çözmüştür. "
                "{correct_count} doğru, {wrong_count} yanlış ve {blank_count} boş ile "
                "%{accuracy} başarı oranı elde edilmiştir.\n\n"
                "Güçlü Kazanım: {strong_topic}\n"
                "Pekiştirilmesi Önerilen Alan: {weak_topic}\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in Türkiye Yüzyılı Maarif Modeli beceri temelli kazanım tablosu güncellenmiştir. "
                "'{strong_topic}' alanında güçlü bir ilerleme kaydedilirken, '{weak_topic}' kazanımında "
                "ek pekiştirme testleri tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Deneme Sınavı ve LGS Değerlendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in son deneme performansında doğruluk oranı %{accuracy} seviyesindedir. "
                "Bir sonraki deneme oturumumuz {next_exam_date} tarihinde '{next_exam_name}' adıyla yapılacaktır.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
        },
    },
    "2. Çalışma Disiplini ve Dijimentor Takibi": {
        "subtopics": [
            "Dijimentor Çalışma Planı Uyumu",
            "Aktif Kullanım ve Düzenlilik",
            "Eksik Telafi Uyarıları",
        ],
        "templates": {
            "Dijimentor Çalışma Planı Uyumu": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}, haftalık Dijimentor çalışma planının %{mentor_completion} "
                "kısmını tamamlamıştır. Düzenli çalışma temposu net artışını desteklemektedir.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Aktif Kullanım ve Düzenlilik": (
                "Sayın {parent_name},\n\n"
                "{student_name} bu hafta platformda düzenli soru ve mini test çözümü gerçekleştirmiştir. "
                "Günlük kısa tekrar rutinlerinin sürdürülmesi önerilmektedir.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Eksik Telafi Uyarıları": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in '{weak_topic}' kazanımındaki telafi görevlerinin tamamlanması "
                "konu bütünlüğü açısından önem taşımaktadır.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
        },
    },
    "3. Sınav, Ödev ve Takvim Duyuruları": {
        "subtopics": [
            "Yaklaşan Sınav ve Deneme Takvimi",
            "Ödev ve Etkinlik Bilgilendirmesi",
            "Özel Dönem Bilgilendirmeleri",
        ],
        "templates": {
            "Yaklaşan Sınav ve Deneme Takvimi": (
                "Sayın {parent_name},\n\n"
                "{next_exam_date} tarihinde '{next_exam_name}' sınav oturumu yapılacaktır. "
                "Öğrencimiz {student_name}'in sınav saatinden 15 dakika önce hazır bulunmasını rica ederiz.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Ödev ve Etkinlik Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} için '{weak_topic}' ve '{strong_topic}' odaklı "
                "haftalık çalışma görevi tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Özel Dönem Bilgilendirmeleri": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Okul dönemsel kamp ve kazanım tekrar programları aktif edilmiştir. "
                "Detaylara panel üzerinden ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
        },
    },
    "4. Abonelik, Paket ve Kurumsal İşlemler": {
        "subtopics": [
            "Okul Sırt Çantası & Paket Durumu",
            "Ödeme ve Fatura Bilgilendirmesi",
            "Destek ve Veli-Öğretmen İletişimi",
        ],
        "templates": {
            "Okul Sırt Çantası & Paket Durumu": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} hesabında {rem_questions} soru üretim kotası ve "
                "{rem_mentor} Dijimentor plan hakkı tanımlıdır.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Ödeme ve Fatura Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Abonelik ve ödeme bilgilendirmeleri {parent_email} adresine iletilmektedir. "
                "Her türlü soru için Veli Takip Paneli üzerinden bize ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
            "Destek ve Veli-Öğretmen İletişimi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in gelişim süreciyle ilgili rehberlik ve öğretmen görüşmesi "
                "taleplerinizi Veli Takip Paneli üzerinden iletebilirsiniz.\n\n"
                "Kurşun Kalem Dijital Okul"
            ),
        },
    },
}

LOGO_DATA_URI = svg_to_data_uri(EMBEDDED_LOGO_SVG)
BACKPACK_DATA_URI = svg_to_data_uri(SCHOOL_BACKPACK_SVG)
LOGO_PNG_BYTES = get_embedded_logo_png_bytes()

# ==============================================================================
# KURUMSAL SAAS TASARIM SİSTEMİ (LACİVERT + ALTIN + FERAH ZEMİN PALETİ)
# ==============================================================================
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        :root {
            --kp-navy: #162B54;
            --kp-navy-hover: #1E3A6E;
            --kp-navy-soft: #EBF0F8;
            --kp-gold: #C48F1B;
            --kp-gold-soft: #FFFBEB;
            --kp-gold-border: #E8C97A;
            --kp-bg: #F4F6FA;
            --kp-surface: #FFFFFF;
            --kp-surface-alt: #F8FAFC;
            --kp-border: #DCE3EC;
            --kp-text: #0F172A;
            --kp-muted: #475569;
        }

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            color: var(--kp-text);
        }
        .stApp {
            background-color: var(--kp-bg);
        }
        [data-testid="stHeader"] {
            background: rgba(244, 246, 250, 0.92);
            backdrop-filter: blur(8px);
            border-bottom: 1px solid var(--kp-border);
        }
        [data-testid="stDecoration"] {
            display: none;
        }
        .block-container {
            padding-top: 1.15rem;
            padding-bottom: 4.5rem;
            max-width: 1240px;
        }

        /* Kenar Çubuğu */
        section[data-testid="stSidebar"] {
            background-color: var(--kp-surface);
            border-right: 1px solid var(--kp-border);
        }
        section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.36rem;
        }
        .kp-sidebar-brand {
            display: flex;
            align-items: center;
            gap: 0.72rem;
            padding: 0.15rem 0.2rem 0.85rem 0.2rem;
            border-bottom: 1px solid var(--kp-border);
            margin-bottom: 0.65rem;
        }
        .kp-sidebar-brand img {
            width: 44px;
            height: 44px;
            object-fit: contain;
            border-radius: 10px;
            border: 1px solid var(--kp-gold-border);
            padding: 2px;
            background: #FFFFFF;
            flex-shrink: 0;
        }
        .kp-sidebar-brand-title {
            font-size: 0.95rem;
            font-weight: 700;
            color: var(--kp-navy);
            line-height: 1.2;
        }
        .kp-sidebar-brand-sub {
            font-size: 0.73rem;
            color: var(--kp-muted);
            margin-top: 2px;
        }

        div[data-testid="stSidebar"] .stButton {
            width: 100%;
        }
        div[data-testid="stSidebar"] .stButton > button {
            width: 100% !important;
            display: flex !important;
            text-align: left !important;
            justify-content: flex-start !important;
            align-items: center !important;
            border-radius: 8px !important;
            padding: 0.52rem 0.85rem !important;
            font-size: 0.84rem !important;
            font-weight: 500 !important;
            border: 1px solid #E2E8F0 !important;
            background: #FFFFFF !important;
            color: #1E293B !important;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03) !important;
            transition: all 0.14s ease !important;
        }
        div[data-testid="stSidebar"] .stButton > button:hover {
            background: var(--kp-navy-soft) !important;
            color: var(--kp-navy) !important;
            border-color: #BDD0E8 !important;
        }
        div[data-testid="stSidebar"] .stButton > button[kind="primary"] {
            background: var(--kp-navy) !important;
            color: #FFFFFF !important;
            border-color: var(--kp-navy) !important;
            font-weight: 600 !important;
            box-shadow: 0 2px 6px rgba(22, 43, 84, 0.18) !important;
        }

        .kp-sidebar-meta {
            background: var(--kp-surface-alt);
            border: 1px solid var(--kp-border);
            border-left: 3px solid var(--kp-gold);
            border-radius: 10px;
            padding: 0.82rem 0.9rem;
            margin-top: 0.7rem;
            font-size: 0.79rem;
            color: #1E293B;
            line-height: 1.5;
        }

        /* Üst Kurumsal Header */
        .kp-top-header {
            background: var(--kp-surface);
            border: 1px solid var(--kp-border);
            border-top: 3px solid var(--kp-navy);
            border-radius: 12px;
            padding: 1.05rem 1.4rem;
            margin-bottom: 1.15rem;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        }
        .kp-top-left {
            display: flex;
            align-items: center;
            gap: 0.95rem;
        }
        .kp-top-left img {
            width: 50px;
            height: 50px;
            object-fit: contain;
            border-radius: 10px;
            border: 1px solid var(--kp-gold-border);
            padding: 2px;
            background: #FFFFFF;
            flex-shrink: 0;
        }
        .kp-top-title {
            font-size: 1.24rem;
            font-weight: 800;
            color: var(--kp-navy);
            letter-spacing: -0.02em;
            line-height: 1.2;
            margin: 0;
        }
        .kp-top-sub {
            font-size: 0.83rem;
            color: #334155;
            margin: 0.2rem 0 0 0;
        }
        .kp-header-badges {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            flex-wrap: wrap;
        }
        .kp-status-chip {
            font-size: 0.74rem;
            font-weight: 600;
            color: var(--kp-navy);
            background: var(--kp-navy-soft);
            border: 1px solid #C6D4E8;
            padding: 0.32rem 0.76rem;
            border-radius: 999px;
            white-space: nowrap;
        }
        .kp-gold-chip {
            font-size: 0.74rem;
            font-weight: 700;
            color: #7D5606;
            background: var(--kp-gold-soft);
            border: 1px solid var(--kp-gold-border);
            padding: 0.32rem 0.76rem;
            border-radius: 999px;
            white-space: nowrap;
        }

        /* Sayfa Bağlam Şeridi */
        .kp-page-strip {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 1rem;
            padding: 0.1rem 0 0.78rem 0;
            margin-bottom: 1rem;
            border-bottom: 1px solid var(--kp-border);
        }
        .kp-page-strip-title {
            font-size: 1.32rem;
            font-weight: 700;
            color: var(--kp-navy);
            letter-spacing: -0.015em;
            margin: 0;
        }
        .kp-page-strip-desc {
            font-size: 0.84rem;
            color: var(--kp-muted);
            margin: 0.2rem 0 0 0;
        }

        /* Sekmeler (stTabs) Kurumsal Renkleri */
        button[data-baseweb="tab"] {
            color: #475569 !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
        }
        button[data-baseweb="tab"][aria-selected="true"] {
            color: var(--kp-navy) !important;
        }
        div[data-baseweb="tab-highlight"] {
            background-color: var(--kp-navy) !important;
            height: 2.5px !important;
        }

        /* Metrik Kartları (st.metric) */
        div[data-testid="stMetric"] {
            background: var(--kp-surface);
            border: 1px solid var(--kp-border);
            border-radius: 12px;
            padding: 0.95rem 1.1rem;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        }
        div[data-testid="stMetric"] [data-testid="stMetricLabel"] {
            font-size: 0.78rem;
            font-weight: 600;
            color: #475569;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            font-size: 1.55rem;
            font-weight: 800;
            color: var(--kp-navy);
            letter-spacing: -0.02em;
        }

        /* Kartlar & Paneller */
        .kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-kpi-card, .kp-metric-box, .kp-studio-banner {
            background: var(--kp-surface);
            border: 1px solid var(--kp-border);
            border-radius: 12px;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        }
        .kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-studio-banner {
            padding: 1.15rem 1.25rem;
        }
        .kp-card {
            margin-bottom: 0.65rem;
        }
        .kp-mod-card {
            background: var(--kp-surface);
            border: 1px solid var(--kp-border);
            border-radius: 12px 12px 6px 6px;
            padding: 1.15rem 1.25rem 1rem 1.25rem;
            min-height: 126px;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        }
        .kp-mod-title {
            font-weight: 700;
            font-size: 0.98rem;
            color: var(--kp-navy);
            margin-bottom: 0.35rem;
        }
        .kp-mod-desc {
            font-size: 0.83rem;
            color: var(--kp-muted);
            line-height: 1.45;
            margin: 0;
        }

        .kp-studio-banner {
            margin-bottom: 1.1rem;
            border-left: 4px solid var(--kp-gold);
        }
        .kp-studio-title {
            font-size: 1.16rem;
            font-weight: 700;
            color: var(--kp-navy);
            margin: 0;
        }
        .kp-studio-sub {
            font-size: 0.85rem;
            color: var(--kp-muted);
            margin: 0.25rem 0 0 0;
            line-height: 1.45;
        }
        .kp-kpi-card {
            padding: 0.95rem 1.05rem;
        }
        .kp-kpi-label {
            font-size: 0.71rem;
            font-weight: 700;
            color: #64748B;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }
        .kp-kpi-value {
            font-size: 1.28rem;
            font-weight: 700;
            color: var(--kp-navy);
            margin-top: 0.2rem;
        }
        .kp-kpi-sub {
            font-size: 0.77rem;
            color: var(--kp-muted);
            margin-top: 0.16rem;
        }

        .kp-metric-box {
            padding: 1.1rem 1rem;
            text-align: center;
            border-bottom: 3px solid var(--kp-navy);
            transition: transform 0.14s ease, border-color 0.14s ease;
        }
        .kp-metric-box:hover {
            transform: translateY(-1px);
            border-bottom-color: var(--kp-gold);
        }
        .kp-metric-num {
            font-size: 1.65rem;
            font-weight: 800;
            color: var(--kp-navy);
            line-height: 1.15;
            letter-spacing: -0.02em;
        }
        .kp-metric-lbl {
            font-size: 0.72rem;
            color: #64748B;
            margin-top: 0.32rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
        }

        /* Çanta Görseli & Satırları */
        .kp-bag-shell {
            margin-bottom: 1.2rem;
        }
        .kp-bag-visual {
            background: var(--kp-surface-alt);
            border: 1px solid var(--kp-border);
            border-radius: 12px;
            padding: 0.95rem;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        .kp-bag-visual img {
            width: 112px;
            height: 130px;
            object-fit: contain;
        }
        .kp-bag-count-pill {
            margin-top: 0.45rem;
            background: var(--kp-navy);
            color: #FFFFFF;
            font-size: 0.75rem;
            font-weight: 600;
            padding: 0.22rem 0.7rem;
            border-radius: 999px;
        }
        .kp-bag-item-chip {
            background: var(--kp-surface-alt);
            border: 1px solid var(--kp-border);
            border-radius: 10px;
            padding: 0.6rem 0.85rem;
            margin-bottom: 0.45rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: var(--kp-text);
        }

        /* Paket Kartları */
        .kp-pkg-card {
            margin-bottom: 0.65rem;
        }
        .kp-pkg-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 0.5rem;
        }
        .kp-pkg-seg {
            font-size: 0.72rem;
            font-weight: 600;
            color: var(--kp-navy);
            background: var(--kp-navy-soft);
            padding: 0.2rem 0.6rem;
            border-radius: 6px;
        }
        .kp-pkg-badge {
            font-size: 0.72rem;
            font-weight: 700;
            color: #7D5606;
            background: var(--kp-gold-soft);
            border: 1px solid var(--kp-gold-border);
            padding: 0.2rem 0.6rem;
            border-radius: 6px;
        }
        .kp-pkg-title {
            font-size: 1.03rem;
            font-weight: 700;
            color: var(--kp-navy);
            margin: 0.6rem 0 0.2rem 0;
        }
        .kp-pkg-price {
            font-size: 1.38rem;
            font-weight: 800;
            color: var(--kp-text);
            margin-bottom: 0.45rem;
        }
        .kp-pkg-price span {
            font-size: 0.82rem;
            font-weight: 500;
            color: #64748B;
        }
        .kp-pkg-desc {
            font-size: 0.83rem;
            color: var(--kp-muted);
            line-height: 1.45;
            margin-bottom: 0.75rem;
            min-height: 3.3rem;
        }
        .kp-pkg-feats {
            list-style: none;
            padding: 0.7rem 0 0 0;
            margin: 0;
            border-top: 1px solid var(--kp-border);
        }
        .kp-pkg-feats li {
            font-size: 0.8rem;
            color: #334155;
            padding: 0.18rem 0;
        }

        .kp-dist-bar {
            display: flex;
            height: 8px;
            border-radius: 999px;
            overflow: hidden;
            background: #E2E8F0;
            margin: 0.5rem 0 0.45rem 0;
        }
        .kp-dist-legend {
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
            font-size: 0.78rem;
            color: var(--kp-muted);
        }

        /* Yüksek Kontrastlı Form Alanları & Kurumsal Butonlar */
        div[data-testid="stTextInput"] input,
        div[data-testid="stNumberInput"] input,
        div[data-testid="stTextArea"] textarea {
            background-color: #FFFFFF !important;
            color: var(--kp-text) !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 8px !important;
            font-size: 0.88rem !important;
        }
        div[data-testid="stNumberInput"] button {
            background-color: #F8FAFC !important;
            color: var(--kp-navy) !important;
            border-color: #CBD5E1 !important;
        }
        div[data-testid="stTextInput"] input:focus,
        div[data-testid="stNumberInput"] input:focus,
        div[data-testid="stTextArea"] textarea:focus {
            border-color: var(--kp-navy) !important;
            box-shadow: 0 0 0 2px rgba(22, 43, 84, 0.12) !important;
        }
        div[data-testid="stSelectbox"] > div > div {
            background-color: #FFFFFF !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 8px !important;
            color: var(--kp-text) !important;
        }

        .stButton > button,
        div[data-testid="stFormSubmitButton"] > button {
            border-radius: 8px;
            font-weight: 600;
            font-size: 0.85rem;
            padding: 0.5rem 0.95rem;
            border: 1px solid #CBD5E1;
            background-color: #FFFFFF;
            color: var(--kp-navy);
            transition: all 0.14s ease;
        }
        .stButton > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover {
            border-color: var(--kp-navy);
            background-color: var(--kp-navy-soft);
        }
        .stButton > button[kind="primary"],
        div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
            background-color: var(--kp-navy) !important;
            border-color: var(--kp-navy) !important;
            color: #FFFFFF !important;
        }
        .stButton > button[kind="primary"]:hover,
        div[data-testid="stFormSubmitButton"] > button[kind="primary"]:hover {
            background-color: var(--kp-navy-hover) !important;
            border-color: var(--kp-navy-hover) !important;
            color: #FFFFFF !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# YARDIMCI FONKSİYONLAR & KALICI DURUM YÖNETİMİ
# ==============================================================================
def _default_parent_stats() -> Dict[str, Any]:
    return {
        "student_name": "Elif Öztürk",
        "grade": "8. Sınıf (LGS)",
        "parent_name": "Selin Öztürk",
        "parent_phone": "+905551234567",
        "parent_email": "veli@ornek.com",
        "weekly_solved": 340,
        "correct": 275,
        "wrong": 45,
        "blank": 20,
        "mentor_completion": 88,
        "strong_topic": "Okuma ve Anlama (%84)",
        "weak_topic": "Sayısal Muhakeme ve Problem Çözme",
        "next_exam_date": (datetime.now() + timedelta(days=5)).strftime("%d.%m.%Y"),
        "next_exam_name": "Türkiye Geneli LGS Deneme Sınavı",
    }


def _default_parent_announcements() -> List[Dict[str, str]]:
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    return [
        {
            "category": "1. Akademik İlerleme ve Başarı Takibi",
            "subtopic": "Haftalık Soru ve Net Özeti",
            "body": "Bu hafta 340 soru çözüldü ve %81 ortalama doğruluk sağlandı.",
            "date": now_str,
        },
        {
            "category": "2. Çalışma Disiplini ve Dijimentor Takibi",
            "subtopic": "Dijimentor Çalışma Planı Uyumu",
            "body": "Dijimentor haftalık çalışma planına %88 oranında uyum sağlandı.",
            "date": now_str,
        },
    ]


def _default_activity_log() -> List[Dict[str, str]]:
    return [
        {"time": "09:15:00", "module": "Soru Fabrikası", "summary": "8. Sınıf LGS Muhakeme seti hazır"},
        {"time": "10:42:18", "module": "Dijimentor", "summary": "Haftalık çeldirici & tempo analizi aktif"},
    ]


def load_persisted_state() -> Dict[str, Any]:
    """Diskteki platform durumu dosyasını güvenli biçimde okur."""
    if not STATE_FILE_PATH.exists():
        return {}
    try:
        raw = STATE_FILE_PATH.read_text(encoding="utf-8")
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_persisted_state() -> None:
    """Kritik kullanıcı/kurum verilerini diske yazar."""
    payload = {
        "backpack_cart": st.session_state.get("backpack_cart", {}),
        "selected_tier_id": st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"]),
        "remaining_question_quota": st.session_state.get("remaining_question_quota"),
        "remaining_coach_quota": st.session_state.get("remaining_coach_quota"),
        "quota_tier_id": st.session_state.get("quota_tier_id"),
        "parent_student_stats": st.session_state.get("parent_student_stats", _default_parent_stats()),
        "parent_announcements": st.session_state.get("parent_announcements", _default_parent_announcements()),
        "activity_log": st.session_state.get("activity_log", _default_activity_log()),
        "b2b_org_cache": st.session_state.get("b2b_org_cache"),
        "last_b2b_summary": st.session_state.get("last_b2b_summary", ""),
    }
    try:
        STATE_FILE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


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


def resolve_iyzico_config() -> Tuple[str, str, str]:
    """iyzico API anahtarlarını ve base_url değerini st.secrets veya ortam değişkenlerinden okur."""
    api_key = ""
    secret_key = ""
    base_url = "https://sandbox-api.iyzipay.com"

    try:
        if "iyzico" in st.secrets:
            sec = st.secrets["iyzico"]
            api_key = str(sec.get("api_key", "")).strip()
            secret_key = str(sec.get("secret_key", "")).strip()
            base_url = str(sec.get("base_url", base_url)).strip()
    except Exception:
        pass

    if not api_key:
        for k_name in ("IYZICO_API_KEY", "IYSU_API_KEY"):
            val = os.environ.get(k_name, "").strip()
            if val:
                api_key = val
                break
            try:
                if k_name in st.secrets:
                    api_key = str(st.secrets[k_name]).strip()
                    if api_key:
                        break
            except Exception:
                pass

    if not secret_key:
        for s_name in ("IYZICO_SECRET_KEY", "IYSU_SECRET_KEY"):
            val = os.environ.get(s_name, "").strip()
            if val:
                secret_key = val
                break
            try:
                if s_name in st.secrets:
                    secret_key = str(st.secrets[s_name]).strip()
                    if secret_key:
                        break
            except Exception:
                pass

    env_url = os.environ.get("IYZICO_BASE_URL", "").strip()
    if env_url:
        base_url = env_url
    else:
        try:
            if "IYZICO_BASE_URL" in st.secrets and str(st.secrets["IYZICO_BASE_URL"]).strip():
                base_url = str(st.secrets["IYZICO_BASE_URL"]).strip()
        except Exception:
            pass

    return api_key, secret_key, base_url


def initialize_iyzico_checkout(
    buyer_name: str,
    buyer_surname: str,
    buyer_email: str,
    buyer_phone: str,
    amount_try: float,
    basket_title: str,
) -> Dict[str, Any]:
    """iyzico CheckoutFormInitialize isteği oluşturur ve yanıt sözlüğünü döndürür."""
    api_key, secret_key, base_url = resolve_iyzico_config()
    if not api_key or not secret_key:
        return {
            "status": "failure",
            "errorMessage": (
                "IYZICO_API_KEY ve IYZICO_SECRET_KEY tanımlı değil. "
                "Streamlit Secrets veya ortam değişkenlerine anahtarlarınızı ekleyin."
            ),
        }

    options = {
        "api_key": api_key,
        "secret_key": secret_key,
        "base_url": base_url,
    }
    price_str = f"{max(1.0, float(amount_try)):.2f}"
    conv_id = f"KP-{int(time.time())}"
    basket_id = f"B-{int(time.time())}"

    request = {
        "locale": "tr",
        "conversationId": conv_id,
        "price": price_str,
        "paidPrice": price_str,
        "currency": "TRY",
        "basketId": basket_id,
        "paymentGroup": "PRODUCT",
        "callbackUrl": "https://kursunkalemdijitalokul.com.tr",
        "enabledInstallments": [1, 2, 3, 6, 9, 12],
        "buyer": {
            "id": "BY_KP_01",
            "name": buyer_name.strip() or "Kürşat Han",
            "surname": buyer_surname.strip() or "Kurşun",
            "gsmNumber": buyer_phone.strip() or "+905350000000",
            "email": buyer_email.strip() or "iletisim@kursunkalemdijitalokul.com.tr",
            "identityNumber": "74300864791",
            "registrationAddress": "İstanbul, Türkiye",
            "ip": "85.34.78.112",
            "city": "Istanbul",
            "country": "Turkey",
        },
        "shippingAddress": {
            "contactName": f"{buyer_name.strip() or 'Kürşat Han'} {buyer_surname.strip() or 'Kurşun'}",
            "city": "Istanbul",
            "country": "Turkey",
            "address": "İstanbul, Türkiye",
        },
        "billingAddress": {
            "contactName": f"{buyer_name.strip() or 'Kürşat Han'} {buyer_surname.strip() or 'Kurşun'}",
            "city": "Istanbul",
            "country": "Turkey",
            "address": "İstanbul, Türkiye",
        },
        "basketItems": [
            {
                "id": "BI_KP_101",
                "name": basket_title[:120],
                "category1": "Dijital Eğitim",
                "itemType": "VIRTUAL",
                "price": price_str,
            }
        ],
    }

    checkout_form_initialize = iyzipay.CheckoutFormInitialize().create(request, options)
    raw_res = checkout_form_initialize.read().decode("utf-8")
    return json.loads(raw_res)


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
    """Oturum içi aktivite akışına kayıt ekler ve diske kaydeder."""
    logs: List[Dict[str, str]] = st.session_state.setdefault("activity_log", [])
    logs.insert(
        0,
        {
            "time": datetime.now().strftime("%H:%M:%S"),
            "module": module_name,
            "summary": summary,
        },
    )
    st.session_state["activity_log"] = logs[:15]
    save_persisted_state()


def get_tier_by_id(tier_id: str) -> Dict[str, Any]:
    """Paket kimliğine göre abonelik paketi kaydını döndürür."""
    return next((t for t in SUBSCRIPTION_TIERS if t["id"] == tier_id), SUBSCRIPTION_TIERS[1])


def ensure_bag_initialized() -> None:
    """Sepet sözlüğünü yalnızca ilk oturum açılışında başlatır."""
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
    if (
        force_reset
        or st.session_state.get("quota_tier_id") != sig
        or "remaining_question_quota" not in st.session_state
        or "remaining_coach_quota" not in st.session_state
    ):
        st.session_state["quota_tier_id"] = sig
        st.session_state["remaining_question_quota"] = int(summary["monthly_quota"])
        st.session_state["remaining_coach_quota"] = int(summary["coach_quota"])
        save_persisted_state()


def cb_add_tier_to_bag(tier_id: str, qty: int = 1) -> None:
    """Callback: Paketi okul sırt çantasına ekler ve kotaları günceller."""
    bag = dict(get_bag_items())
    bag[tier_id] = bag.get(tier_id, 0) + max(1, int(qty))
    st.session_state["backpack_cart"] = bag
    st.session_state["selected_tier_id"] = tier_id
    sync_plan_quotas(get_bag_summary(), force_reset=True)
    tier = get_tier_by_id(tier_id)
    record_activity("Sırt Çantası", f"{tier['name']} çantaya eklendi")


def cb_update_bag_qty(tier_id: str, delta: int) -> None:
    """Callback: Çantadaki paketin adedini günceller veya çıkarır."""
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
    """Callback: Okul sırt çantasını boşaltır."""
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
        save_persisted_state()
        return True, ""

    rem_c = int(st.session_state.get("remaining_coach_quota", 0))
    if rem_c < amount:
        return False, (
            f"Okul sırt çantanızda ({summary['name']}) Dijimentor plan kotası doldu "
            f"(Kalan: {rem_c}). Paketler vitrininden paket ekleyebilirsiniz."
        )
    st.session_state["remaining_coach_quota"] = rem_c - amount
    save_persisted_state()
    return True, ""


def extract_answer_key(raw_text: str) -> Dict[int, str]:
    """Metindeki 'Cevap Anahtarı' veya satır içi cevap ifadelerinden {soru_no: 'A'|'B'|'C'|'D'} haritası çıkarır."""
    if not raw_text.strip():
        return {}

    parts = re.split(r"(?:^|\n)#*\s*Cevap\s*Anahtar[^\n]*", raw_text, maxsplit=1, flags=re.IGNORECASE)
    search_space = parts[1] if len(parts) > 1 else raw_text
    key_map: Dict[int, str] = {}

    patterns = [
        r"(?:Soru\s*)?(\d+)\s*[\.\-\:\)]\s*\**\s*([A-D])\b",
        r"(\d+)\s*\.\s*Soru\s*[\:\-\)]\s*\**\s*([A-D])\b",
        r"\b(\d+)\s*[-–—]\s*([A-D])\b",
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
    blocks = re.split(
        r"(?=(?:^|\n)#*\s*(?:Soru\s*\d+|(?:\d+\s*[\.\)])))",
        main_body.strip(),
        flags=re.IGNORECASE,
    )
    parsed: List[Dict[str, Any]] = []

    for block in blocks:
        text = block.strip()
        num_match = re.match(r"^#*\s*(?:Soru\s*)?(\d+)\b", text, flags=re.IGNORECASE)
        if not num_match:
            continue
        q_num = int(num_match.group(1))
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        stem_lines: List[str] = []
        options: Dict[str, str] = {}
        current_opt: Optional[str] = None

        for ln in lines:
            if re.match(r"^\[?(?:Redaksiyon|Çözüm|Cevap)", ln, flags=re.IGNORECASE):
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


def render_page_strip(page_name: str) -> None:
    """Aktif sayfa için üst header tekrarı olmadan sade bağlam başlığı çizer."""
    title, desc = PAGE_DESCRIPTIONS.get(
        page_name,
        (page_name, "Kurşun Kalem Dijital Okul çalışma alanı."),
    )
    st.markdown(
        f"""
        <div class="kp-page-strip">
            <div>
                <h2 class="kp-page-strip-title">{html.escape(title)}</h2>
                <p class="kp-page-strip-desc">{html.escape(desc)}</p>
            </div>
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
    title_run.font.color.rgb = RGBColor(22, 43, 84)

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
            run.font.color.rgb = RGBColor(22, 43, 84)

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
    r.font.color.rgb = RGBColor(22, 43, 84)

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
            run.font.color.rgb = RGBColor(22, 43, 84)

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
    r.font.color.rgb = RGBColor(22, 43, 84)

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
            run.font.color.rgb = RGBColor(22, 43, 84)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_school_backpack_panel() -> Dict[str, Any]:
    """Seçilen paketleri alışveriş sepeti gibi Okul Sırt Çantası içinde sade ve hizalı gösterir."""
    summary = get_bag_summary()
    entries = summary["entries"]

    st.markdown('<div class="kp-bag-shell">', unsafe_allow_html=True)
    h_col, clr_col = st.columns([4, 1])
    with h_col:
        st.markdown("#### Okul Sırt Çantası")
        st.caption("Seçili paketler, adet kontrolleri ve aylık toplam kota özeti.")
    with clr_col:
        if entries:
            st.button(
                "Çantayı Boşalt",
                key="bag_clear_all",
                use_container_width=True,
                on_click=cb_clear_bag,
            )

    v_col, list_col = st.columns([1.1, 2.9])

    with v_col:
        chips_html = ""
        for entry in entries:
            safe_name = html.escape(entry["tier"]["name"])
            safe_qty = int(entry["qty"])
            chips_html += (
                f'<div style="background:#FFFFFF;color:#0F172A;border:1px solid #CBD5E1;'
                f'border-radius:7px;padding:0.22rem 0.5rem;font-size:0.74rem;font-weight:600;'
                f'margin-top:0.28rem;width:100%;text-align:center;">'
                f'{safe_name} (x{safe_qty})</div>'
            )

        if not chips_html:
            chips_html = (
                '<div style="color:#64748B;font-size:0.78rem;margin-top:0.35rem;text-align:center;">'
                'Çanta boş</div>'
            )

        st.markdown(
            f"""
            <div class="kp-bag-visual">
                <img src="{BACKPACK_DATA_URI}" alt="Okul Sırt Çantası"/>
                <div class="kp-bag-count-pill">{summary['total_count']} Paket</div>
                <div style="width:100%;margin-top:0.3rem;">{chips_html}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with list_col:
        if not entries:
            st.info(
                "Okul sırt çantanız şu an boş. Aşağıdaki paketlerden **Çantaya Ekle** "
                "butonunu kullanarak paket ekleyebilirsiniz."
            )
        else:
            for entry in entries:
                t = entry["tier"]
                qty = entry["qty"]
                c_info, c_minus, c_plus, c_rem = st.columns([4.6, 0.75, 0.75, 1.15])
                with c_info:
                    st.markdown(
                        f"""
                        <div class="kp-bag-item-chip">
                            <div>
                                <div style="font-weight:700;font-size:0.88rem;color:#0F172A;">
                                    {html.escape(t['name'])}
                                    <span style="color:#64748B;font-size:0.76rem;font-weight:500;"> • {html.escape(t['segment'])}</span>
                                </div>
                                <div style="color:#475569;font-size:0.78rem;margin-top:0.12rem;">
                                    Adet: <b>{qty}</b> • Soru: <b>{entry['line_quota']:,}/ay</b> • Dijimentor: <b>{entry['line_coach']}/ay</b>
                                </div>
                            </div>
                            <div style="font-weight:700;font-size:0.92rem;color:#162B54;white-space:nowrap;">
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

            st.markdown("<div style='margin-top:0.45rem;'></div>", unsafe_allow_html=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Aylık Toplam", f"{summary['monthly_price']:,} ₺".replace(",", "."))
            m2.metric("Aylık Soru Kotası", f"{summary['monthly_quota']:,}".replace(",", "."))
            m3.metric("Dijimentor Plan", f"{summary['coach_quota']} / Ay")

    st.markdown("</div>", unsafe_allow_html=True)
    return summary


def render_iyzico_simulator(
    package_name: str,
    monthly_price: int,
    summary: Optional[Dict[str, Any]] = None,
) -> None:
    """Kurumsal görünümde iyzico taksit, dönem indirimi, gerçek API Checkout Form ve teklif çıktısı paneli."""
    bag_summary = summary if summary is not None else get_bag_summary()
    effective_monthly = (
        int(bag_summary["monthly_price"]) if bag_summary["total_count"] > 0 else int(monthly_price)
    )
    display_label = (
        f"{package_name} ({bag_summary['total_count']} Paket)"
        if bag_summary["total_count"] > 0
        else package_name
    )

    st.markdown('<div class="kp-pay-box">', unsafe_allow_html=True)
    st.markdown("#### iyzico Taksit & Ödeme Simülatörü")
    st.caption("Dönem avantajı ve taksit planına göre tutarı hesaplayın, iyzico ödeme formu başlatın veya DOCX özet indirin.")

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

    st.markdown("<div style='margin:0.65rem 0;'></div>", unsafe_allow_html=True)
    m1, m2, m3 = st.columns(3)
    m1.metric("Toplam Tutar", f"{total_amount:,} ₺".replace(",", "."))
    m2.metric(f"{installment} Taksit Aylık", f"{monthly_installment:,.2f} ₺".replace(",", " "))
    m3.metric("Ödeme Altyapısı", f"iyzico • {card_brand}")

    st.markdown("<div style='margin:0.65rem 0;'></div>", unsafe_allow_html=True)
    b_col1, b_col2, b_col3 = st.columns(3)
    with b_col1:
        buyer_name = st.text_input("Ad", value="Kürşat Han", key=f"iyz_name_{display_label}")
    with b_col2:
        buyer_surname = st.text_input("Soyad", value="Kurşun", key=f"iyz_surname_{display_label}")
    with b_col3:
        buyer_email = st.text_input("E-posta", value="iletisim@kursunkalemdijitalokul.com.tr", key=f"iyz_email_{display_label}")

    btn_c1, btn_c2, btn_c3 = st.columns([1.8, 1.8, 1.1])
    with btn_c1:
        if st.button(
            "iyzico Ödeme Formu Başlat (API)",
            key=f"iyz_live_btn_{display_label}",
            type="primary",
            use_container_width=True,
        ):
            if total_amount <= 0:
                st.warning("Ödeme başlatmak için önce çantanıza en az bir paket ekleyin.")
            else:
                with st.spinner("iyzico ödeme formu oluşturuluyor..."):
                    res_dict = initialize_iyzico_checkout(
                        buyer_name=buyer_name,
                        buyer_surname=buyer_surname,
                        buyer_email=buyer_email,
                        buyer_phone="+905350000000",
                        amount_try=float(total_amount),
                        basket_title=f"{BRAND_NAME} - {display_label} ({billing_cycle})",
                    )
                if res_dict.get("status") == "success":
                    sync_plan_quotas(bag_summary, force_reset=True)
                    record_activity(
                        "iyzico API",
                        f"{display_label} ({total_amount:,} ₺) ödeme formu açıldı".replace(",", "."),
                    )
                    st.session_state["iyzico_last_response"] = res_dict
                    st.success("iyzico ödeme formu başarıyla oluşturuldu!")
                else:
                    st.error(f"iyzico yanıtı: {res_dict.get('errorMessage', 'Bilinmeyen hata')}")

    with btn_c2:
        if st.button(
            "Simülasyon ile Kota Aktifleştir",
            key=f"pay_btn_{display_label}",
            use_container_width=True,
        ):
            sync_plan_quotas(bag_summary, force_reset=True)
            record_activity(
                "iyzico Simülasyon",
                f"{display_label} ({total_amount:,} ₺) aktiflendi".replace(",", "."),
            )
            st.success(
                f"**{display_label}** için **{total_amount:,} ₺** ({installment} taksit, "
                f"{card_brand}) tutarında kota ({bag_summary['monthly_quota']:,} soru) aktifleşti.".replace(",", ".")
            )

    with btn_c3:
        brief_docx = build_commercial_brief_docx(
            title=f"{display_label} Özeti",
            summary=bag_summary,
            billing_cycle=billing_cycle,
            installment=int(installment),
            total_amount=int(total_amount),
            extra_notes=f"Kart Programı: {card_brand}",
        )
        st.download_button(
            label="Özet DOCX",
            data=brief_docx,
            file_name="kursun_kalem_sirt_cantasi_ozet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl_brief_{display_label}",
            use_container_width=True,
        )

    iyz_res = st.session_state.get("iyzico_last_response")
    if isinstance(iyz_res, dict) and iyz_res.get("status") == "success":
        pay_url = iyz_res.get("paymentPageUrl")
        form_html = iyz_res.get("checkoutFormContent")
        if pay_url:
            st.markdown(
                f'<div style="margin-top:0.75rem;"><a href="{html.escape(str(pay_url))}" target="_blank">'
                f'<button style="background:#162B54;color:#FFFFFF;border:none;border-radius:8px;'
                f'padding:0.55rem 1.15rem;font-weight:600;cursor:pointer;">'
                f'Güvenli iyzico Ödeme Sayfasını Aç</button></a></div>',
                unsafe_allow_html=True,
            )
        if form_html:
            st.components.v1.html(str(form_html), height=460, scrolling=True)

    st.markdown("</div>", unsafe_allow_html=True)


def render_student_xray_panel() -> None:
    """Öğrencinin kazanım başarısını, hata kök nedenlerini ve Dijimentor telafi reçetesini sunan röntgen modülü."""
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Öğrenci Röntgeni — Kazanım, Hata Kök Neden &amp; Telafi Tarama</div>
            <div class="kp-studio-sub">
                Öğrencinin ders ve kazanım bazlı performans röntgenini çıkarın, hata türlerini ayrıştırın
                ve tek tıkla Soru Fabrikası'na telafi paketi aktarın.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    f_col1, f_col2, f_col3 = st.columns([2.2, 1.8, 1.0])
    with f_col1:
        sec_student = st.selectbox(
            "Öğrenci Profili Seçin",
            list(OGRENCI_RONTGEN_PROFILLERI.keys()),
            key="xray_student_sel",
        )
    with f_col2:
        sec_scope = st.selectbox(
            "Röntgen Kapsamı",
            ["Son 5 Deneme Ortalaması", "Son 30 Günlük Soru Çözümü", "Tüm Dönem Kümülatif"],
            key="xray_scope_sel",
        )
    with f_col3:
        st.markdown("<div style='height:1.65rem;'></div>", unsafe_allow_html=True)
        if st.button("Röntgen Çek", key="xray_scan_btn", type="primary", use_container_width=True):
            st.session_state["xray_last_scan"] = f"{sec_student} • {sec_scope} ({datetime.now().strftime('%H:%M:%S')})"
            record_activity("Öğrenci Röntgeni", f"{sec_student} ({sec_scope}) tarandı")
            st.toast("Öğrenci röntgeni güncellendi.")

    if st.session_state.get("xray_last_scan"):
        st.caption(f"Son Röntgen Çekimi: {st.session_state['xray_last_scan']}")

    prof = OGRENCI_RONTGEN_PROFILLERI[sec_student]

    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Genel Başarı Endeksi", prof["basari_endeksi"], prof["basari_delta"])
    r2.metric("Kritik Eksik Alan", prof["kritik_konu"], prof["kritik_delta"], delta_color="inverse")
    r3.metric("Çözüm Hızı & Tempo", prof["hiz"], prof["hiz_delta"])
    r4.metric("Hedef Puan Yakınlığı", prof["hedef_puan"], prof["hedef_delta"])

    st.write("")
    t_col, c_col = st.columns([3.1, 1.9])
    with t_col:
        st.markdown("#### Kazanım & Konu Röntgen Tablosu")
        st.dataframe(prof["kazanimlar"], use_container_width=True, hide_index=True)
    with c_col:
        st.markdown("#### Hata Kök Neden Dağılımı (%)")
        st.bar_chart(prof["hatalar"], color="#162B54", height=225)

    st.markdown("#### Dijimentor Otomatik Telafi Reçetesi")
    st.markdown(
        f"""
        <div class="kp-card">
            <div style="font-weight:700;color:#162B54;font-size:0.92rem;">
                {html.escape(sec_student)} — {html.escape(sec_scope)}
            </div>
            <div style="color:#334155;font-size:0.85rem;margin-top:0.4rem;line-height:1.5;">
                <b>Teşhis Özeti:</b> {html.escape(prof['recete_ozet'])}
            </div>
            <div style="color:#475569;font-size:0.84rem;margin-top:0.35rem;line-height:1.45;">
                <b>Önerilen Aksiyon:</b> {html.escape(prof['recete_aksiyon'])}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    a_col1, a_col2 = st.columns(2)
    with a_col1:
        if st.button("Kritik Kazanım İçin Soru Üret", key="xray_to_factory", type="primary", use_container_width=True):
            st.session_state["factory_focus_topic"] = prof["hedef_konu"]
            st.session_state["factory_branch"] = prof["hedef_brans"]
            st.session_state["factory_grade"] = prof["hedef_sinif"]
            record_activity("Röntgen -> Fabrika", f"{prof['hedef_konu']} Soru Fabrikasına aktarıldı")
            navigate_to("Soru & Redaksiyon Fabrikası")
    with a_col2:
        if st.button("Dijimentor Planına Aktar", key="xray_to_mentor", use_container_width=True):
            st.session_state["dk_weak"] = prof["recete_ozet"]
            st.session_state["dk_focus"] = "Dikkat ve Çeldirici Tuzağı Azaltma"
            record_activity("Röntgen -> Dijimentor", f"{sec_student} bulguları mentorluğa aktarıldı")
            navigate_to("Dijimentor")


def render_parent_notification_center() -> None:
    """Veli Takip Paneli içindeki 4 ana kategorili Veli Bilgilendirme & Otomatik Mesaj Merkezi."""
    st.markdown("---")
    st.markdown("### Veli Bilgilendirme & Otomatik Mesaj Merkezi")
    st.caption(
        "Veli bilgilendirme başlıkları altında otomatik mesaj taslağı oluşturun, "
        "duyuru akışına ekleyin ve kurumsal DOCX olarak indirin."
    )

    stats = st.session_state["parent_student_stats"]
    solved_safe = max(1, int(stats["weekly_solved"]))
    acc_pct = round((int(stats["correct"]) / solved_safe) * 100, 1)

    t_msg, t_feed, t_profile = st.tabs(
        ["Mesaj Üreticisi", "Veli Duyuru Akışı", "Öğrenci & Veli Verileri"]
    )

    with t_msg:
        c_left, c_right = st.columns([1.05, 1.35])
        with c_left:
            sel_cat = st.selectbox("Ana Konu Başlığı", list(PARENT_MESSAGE_TOPICS.keys()), key="pm_cat")
            subtopics = PARENT_MESSAGE_TOPICS[sel_cat]["subtopics"]
            sel_sub = st.selectbox("Alt Konu Başlığı", subtopics, key="pm_sub")
            channel = st.radio(
                "Mesaj Formatı",
                ["WhatsApp / Kurumsal Mesaj", "Kısa SMS"],
                horizontal=True,
                key="pm_ch",
            )

            raw_tpl = PARENT_MESSAGE_TOPICS[sel_cat]["templates"][sel_sub]
            msg_text = raw_tpl.format(
                parent_name=stats["parent_name"],
                student_name=stats["student_name"],
                grade=stats["grade"],
                parent_email=stats["parent_email"],
                weekly_solved=stats["weekly_solved"],
                correct_count=stats["correct"],
                wrong_count=stats["wrong"],
                blank_count=stats["blank"],
                accuracy=acc_pct,
                mentor_completion=stats["mentor_completion"],
                strong_topic=stats["strong_topic"],
                weak_topic=stats["weak_topic"],
                next_exam_date=stats["next_exam_date"],
                next_exam_name=stats["next_exam_name"],
                rem_questions=int(st.session_state.get("remaining_question_quota", 0)),
                rem_mentor=int(st.session_state.get("remaining_coach_quota", 0)),
            )

            if channel == "Kısa SMS":
                msg_text = (
                    f"{BRAND_NAME}: Sn. {stats['parent_name']}, {stats['student_name']} "
                    f"({sel_sub}) haftalık {stats['weekly_solved']} soru, %{acc_pct} doğruluk, "
                    f"Dijimentor %{stats['mentor_completion']} uyum."
                )

            if st.button("Veli Duyuru Akışına Ekle", key="pm_add_feed", use_container_width=True):
                st.session_state["parent_announcements"].insert(
                    0,
                    {
                        "category": sel_cat,
                        "subtopic": sel_sub,
                        "body": msg_text,
                        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    },
                )
                record_activity("Veli Mesajı", f"{sel_sub} duyurusu eklendi")
                st.success("Bilgilendirme mesajı Veli Duyuru Akışına eklendi.")

        with c_right:
            edited_msg = st.text_area("Hazır Veli Bilgilendirme Metni", value=msg_text, height=220, key="pm_body")
            clean_tel = re.sub(r"\D", "", stats["parent_phone"])
            encoded = edited_msg.replace("\n", "%0A").replace(" ", "%20")
            wa_url = f"https://wa.me/{clean_tel}?text={encoded}"
            mail_url = f"mailto:{stats['parent_email']}?subject={BRAND_NAME}%20-%20{sel_sub}&body={encoded}"

            l1, l2 = st.columns(2)
            l1.link_button("WhatsApp ile Gönder", wa_url, use_container_width=True)
            l2.link_button("E-Posta ile Gönder", mail_url, use_container_width=True)

            msg_docx = build_coach_docx(
                title_meta=f"Veli Bilgilendirme: {sel_cat} > {sel_sub} | Öğrenci: {stats['student_name']}",
                plan_text=edited_msg,
            )
            st.download_button(
                label="Veli Mesajını DOCX İndir",
                data=msg_docx,
                file_name="veli_bilgilendirme_mesaji.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    with t_feed:
        flt = st.selectbox("Kategoriye Göre Filtrele", ["Tümü"] + list(PARENT_MESSAGE_TOPICS.keys()), key="pm_flt")
        items = [
            a for a in st.session_state["parent_announcements"] if flt == "Tümü" or a["category"] == flt
        ]
        for ann in items:
            st.markdown(
                f"""
                <div class="kp-card">
                    <div style="font-size:0.74rem;color:#64748B;font-weight:600;">
                        {html.escape(ann['category'])} • {html.escape(ann['subtopic'])} • {html.escape(ann['date'])}
                    </div>
                    <div style="font-size:0.88rem;color:#0F172A;margin-top:0.35rem;white-space:pre-line;">
                        {html.escape(ann['body'])}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with t_profile:
        with st.form("parent_profile_form"):
            p1, p2 = st.columns(2)
            with p1:
                st_n = st.text_input("Öğrenci Adı Soyadı", value=stats["student_name"])
                st_g = st.text_input("Sınıf / Düzey", value=stats["grade"])
                pr_n = st.text_input("Veli Adı Soyadı", value=stats["parent_name"])
                pr_ph = st.text_input("Veli Telefon", value=stats["parent_phone"])
                pr_em = st.text_input("Veli E-Posta", value=stats["parent_email"])
            with p2:
                w_s = st.number_input("Haftalık Çözülen Soru", min_value=0, value=int(stats["weekly_solved"]))
                c_c = st.number_input("Doğru Sayısı", min_value=0, value=int(stats["correct"]))
                w_c = st.number_input("Yanlış Sayısı", min_value=0, value=int(stats["wrong"]))
                m_c = st.slider("Dijimentor Plan Uyumu (%)", 0, 100, int(stats["mentor_completion"]))
                wk_t = st.text_input("Pekiştirilmesi Gereken Kazanım", value=stats["weak_topic"])
            if st.form_submit_button("Veli & Öğrenci Verilerini Güncelle", type="primary"):
                stats.update(
                    {
                        "student_name": st_n,
                        "grade": st_g,
                        "parent_name": pr_n,
                        "parent_phone": pr_ph,
                        "parent_email": pr_em,
                        "weekly_solved": int(w_s),
                        "correct": int(c_c),
                        "wrong": int(w_c),
                        "blank": max(0, int(w_s) - (int(c_c) + int(w_c))),
                        "mentor_completion": int(m_c),
                        "weak_topic": wk_t,
                    }
                )
                save_persisted_state()
                record_activity("Veli Paneli", f"{st_n} verileri güncellendi")
                st.success("Öğrenci ve veli verileri güncellendi.")


# ==============================================================================
# SESSION STATE BAŞLATMA & KALICI VERİ YÜKLEME
# ==============================================================================
_persisted = load_persisted_state()

if "active_page" not in st.session_state:
    st.session_state.active_page = PAGES[0]
elif st.session_state.active_page == "Ana Sayfa & Geri Sayım":
    st.session_state.active_page = "Ana Sayfa"

if "generated_test_text" not in st.session_state:
    st.session_state.generated_test_text = ""
if "dijimentor_plan" not in st.session_state:
    st.session_state.dijimentor_plan = ""
if "selected_tier_id" not in st.session_state:
    st.session_state.selected_tier_id = _persisted.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
if "backpack_cart" not in st.session_state and isinstance(_persisted.get("backpack_cart"), dict):
    st.session_state["backpack_cart"] = dict(_persisted["backpack_cart"])

ensure_bag_initialized()

if "quota_tier_id" not in st.session_state and _persisted.get("quota_tier_id"):
    st.session_state["quota_tier_id"] = _persisted["quota_tier_id"]
if "remaining_question_quota" not in st.session_state and _persisted.get("remaining_question_quota") is not None:
    st.session_state["remaining_question_quota"] = int(_persisted["remaining_question_quota"])
if "remaining_coach_quota" not in st.session_state and _persisted.get("remaining_coach_quota") is not None:
    st.session_state["remaining_coach_quota"] = int(_persisted["remaining_coach_quota"])

if "live_tick" not in st.session_state:
    st.session_state.live_tick = False
if "portal_question_draft" not in st.session_state:
    st.session_state.portal_question_draft = ""
if "portal_exam_source" not in st.session_state:
    st.session_state.portal_exam_source = ""
if "activity_log" not in st.session_state:
    p_logs = _persisted.get("activity_log")
    st.session_state.activity_log = p_logs if isinstance(p_logs, list) and p_logs else _default_activity_log()
if "parent_student_stats" not in st.session_state:
    p_stats = _persisted.get("parent_student_stats")
    merged_stats = _default_parent_stats()
    if isinstance(p_stats, dict):
        merged_stats.update(p_stats)
    if merged_stats.get("student_name") == "Ege Yılmaz":
        merged_stats["student_name"] = "Elif Öztürk"
    if merged_stats.get("parent_name") == "Selin Yılmaz":
        merged_stats["parent_name"] = "Selin Öztürk"
    st.session_state.parent_student_stats = merged_stats
if "parent_announcements" not in st.session_state:
    p_ann = _persisted.get("parent_announcements")
    st.session_state.parent_announcements = p_ann if isinstance(p_ann, list) and p_ann else _default_parent_announcements()
if "b2b_org_cache" not in st.session_state and isinstance(_persisted.get("b2b_org_cache"), dict):
    st.session_state["b2b_org_cache"] = _persisted["b2b_org_cache"]
if "last_b2b_summary" not in st.session_state and _persisted.get("last_b2b_summary"):
    st.session_state["last_b2b_summary"] = _persisted["last_b2b_summary"]

bag_summary_global = get_bag_summary()
sync_plan_quotas(bag_summary_global, force_reset=False)
api_key_ready = resolve_gemini_api_key() is not None

# ==============================================================================
# KENAR ÇUBUĞU
# ==============================================================================
with st.sidebar:
    st.markdown(
        f"""
        <div class="kp-sidebar-brand">
            <img src="{LOGO_DATA_URI}" alt="{BRAND_NAME}"/>
            <div>
                <div class="kp-sidebar-brand-title">{BRAND_SHORT}</div>
                <div class="kp-sidebar-brand-sub">Dijital Okul Platformu</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for p_name in PAGES:
        is_current = st.session_state.active_page == p_name
        btn_type = "primary" if is_current else "secondary"
        if st.button(p_name, key=f"nav_{p_name}", type=btn_type, use_container_width=True):
            st.session_state.active_page = p_name
            st.rerun()

    st.markdown(
        f"""
        <div class="kp-sidebar-meta">
            <div style="font-weight:700;color:#162B54;margin-bottom:0.2rem;">Aktif Kota &amp; Çanta</div>
            <div>{bag_summary_global['total_count']} Paket • {bag_summary_global['monthly_price']:,} ₺/ay</div>
            <div style="color:#475569;margin-top:0.25rem;">
                Soru: <b>{int(st.session_state.remaining_question_quota):,}</b> / {bag_summary_global['monthly_quota']:,}<br>
                Dijimentor: <b>{int(st.session_state.remaining_coach_quota)}</b> / {bag_summary_global['coach_quota']}
            </div>
        </div>
        """.replace(",", "."),
        unsafe_allow_html=True,
    )

    st.markdown("<div style='margin-top:0.55rem;'></div>", unsafe_allow_html=True)
    if api_key_ready:
        st.caption("Durum: Sistem aktif")
    else:
        st.caption("Durum: API anahtarı bekleniyor")

    if st.button("Oturumu Sıfırla", key="reset_session_btn", use_container_width=True):
        try:
            if STATE_FILE_PATH.exists():
                STATE_FILE_PATH.unlink()
        except Exception:
            pass
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

# ==============================================================================
# ÜST KURUMSAL HEADER
# ==============================================================================
st.markdown(
    f"""
    <div class="kp-top-header">
        <div class="kp-top-left">
            <img src="{LOGO_DATA_URI}" alt="{BRAND_NAME}"/>
            <div>
                <h1 class="kp-top-title">{BRAND_NAME}</h1>
                <p class="kp-top-sub">{BRAND_TAGLINE}</p>
            </div>
        </div>
        <div class="kp-header-badges">
            <span class="kp-gold-chip">2027 LGS &amp; YKS</span>
            <span class="kp-status-chip">Maarif Modeli</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

current_page = st.session_state.active_page
render_page_strip(current_page)

# ==============================================================================
# 1. ANA SAYFA
# ==============================================================================
if current_page == "Ana Sayfa":
    head_col, toggle_col = st.columns([4, 1], vertical_alignment="center")
    with head_col:
        st.markdown("#### 2027 LGS Geri Sayım (13 Haziran 2027)")
    with toggle_col:
        live_counter = st.toggle("Canlı Sayaç", value=st.session_state.live_tick)
        st.session_state.live_tick = live_counter

    counter_placeholder = st.empty()
    with counter_placeholder.container():
        render_countdown_boxes(LGS_2027_TARGET)

    yks_d, yks_h, _, _ = format_countdown(YKS_2027_TARGET)
    st.caption(f"2027 YKS (19 Haziran 2027) için kalan süre: {yks_d} gün {yks_h} saat")

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Üretilen Soru Havuzu", "142.500+", "+1.200 bugün")
    m2.metric("Aktif Öğrenci / Veli", "18.420", "+340 bu hafta")
    m3.metric("B2B Anlaşmalı Kurum", "48", "+4 Yeni")
    m4.metric("Çözülen Soru Sayısı", "2.4 Milyon", "%98 Başarı")

    st.markdown("### Modüller")
    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown(
            """
            <div class="kp-mod-card">
                <div class="kp-mod-title">Soru &amp; Redaksiyon Fabrikası</div>
                <p class="kp-mod-desc">
                    Maarif Modeli kazanımlarına uygun, çeldirici analizi yapılmış
                    özgün sorular üretin ve DOCX formatında indirin.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Soru Fabrikasına Git", key="go_factory", use_container_width=True):
            navigate_to("Soru & Redaksiyon Fabrikası")
    with f2:
        st.markdown(
            """
            <div class="kp-mod-card">
                <div class="kp-mod-title">Öğrenci Röntgeni</div>
                <p class="kp-mod-desc">
                    Kazanım başarı röntgenini çıkarın, hata kök nedenlerini inceleyin
                    ve tek tıkla telafi soru seti oluşturun.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Öğrenci Röntgenini Aç", key="go_xray", use_container_width=True):
            navigate_to("Öğrenci Röntgeni")
    with f3:
        st.markdown(
            """
            <div class="kp-mod-card">
                <div class="kp-mod-title">Paketler &amp; Abonelik Vitrini</div>
                <p class="kp-mod-desc">
                    Öğrenci, YKS, Veli ve Kurumsal paketleri Okul Sırt Çantasına
                    ekleyin ve iyzico taksit seçeneklerini inceleyin.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Paketler Vitrinine Git", key="go_plans", use_container_width=True):
            navigate_to("Paketler & Abonelik Vitrini")

    st.markdown("---")
    ana_l, ana_r = st.columns([2, 1])
    with ana_l:
        st.markdown("#### Haftalık Soru Üretim & Çözüm Telemetrisi")
        telemetry_data = {
            "Pzt": 16800,
            "Sal": 18200,
            "Çar": 19500,
            "Per": 21100,
            "Cum": 22400,
            "Cmt": 25900,
            "Paz": 24800,
        }
        st.area_chart(telemetry_data, color="#162B54", height=210)
    with ana_r:
        st.markdown("#### Son Aktiviteler")
        for item in st.session_state.get("activity_log", [])[:5]:
            st.markdown(
                f"""
                <div style="background:#FFFFFF;border:1px solid #DCE3EC;border-left:3px solid #C48F1B;border-radius:10px;padding:0.58rem 0.8rem;margin-bottom:0.45rem;">
                    <div style="font-size:0.72rem;font-weight:600;color:#64748B;">{html.escape(item['time'])} • {html.escape(item['module'])}</div>
                    <div style="font-size:0.83rem;color:#0F172A;margin-top:0.12rem;">{html.escape(item['summary'])}</div>
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
    if st.session_state.get("factory_focus_topic"):
        st.info(f"Aktarılan odak konu: **{st.session_state['factory_focus_topic']}**")

    grade_options = ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS)", "9-10. Sınıf", "11-12. Sınıf (YKS)"]
    default_branch = st.session_state.get("factory_branch", "Türkçe")
    if default_branch not in CURRICULUM_TOPICS:
        default_branch = "Türkçe"
    default_grade = st.session_state.get("factory_grade", "5. Sınıf")
    if default_grade not in grade_options:
        default_grade = "5. Sınıf"

    with st.expander("Test ve Kurum Parametreleri", expanded=True):
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
            distractor_mode = st.selectbox(
                "Çeldirici & Kök Stratejisi",
                [
                    "Muhakeme & Çeldirici Tuzağı Dengeli",
                    "Olumsuz Kök ('Değildir / Ulaşılamaz') Odaklı",
                    "Öncüllü & Bağlamsal Analiz Odaklı",
                ],
                index=0,
            )
            auto_review = st.checkbox(
                "Otomatik Soru Kalite & Redaksiyon Denetimi",
                value=True,
            )

    if st.button("Kurumsal Soru Paketi Üret", type="primary"):
        quota_ok, quota_msg = consume_quota("question", int(question_count))
        if not quota_ok:
            st.error(quota_msg)
        else:
            sys_prompt = (
                "Sen Türkiye Yüzyılı Maarif Modeli, LGS ve YKS ölçme-değerlendirme standartlarında "
                "uzman bir kıdemli test editörüsün. Soruların özgün, dilbilgisi ve noktalama açısından "
                "kusursuz, çeldiricilerinin pedagojik olarak gerekçeli olmasını sağla. "
                "Her soruyu 'Soru 1', 'Soru 2'... şeklinde numaralandır ve şıkları 'A)', 'B)', 'C)', 'D)' "
                "olarak yaz. Otomatik redaksiyon açıksa her sorunun sonunda kısa bir [Redaksiyon & Çeldirici Notu] "
                "ekle ve en sonda mutlaka 'Cevap Anahtarı' başlığı altında '1-A', '2-C' formatında net anahtar ver."
            )
            user_prompt = (
                f"Kurum: {institution}\n"
                f"Branş: {branch}\n"
                f"Sınıf Seviyesi: {grade}\n"
                f"Öğrenme Alanı: {learning_area}\n"
                f"Hedef Kazanım: {outcome}\n"
                f"Bağlam / Tema / Senaryo: {context_theme}\n"
                f"Zorluk Derecesi: {difficulty}\n"
                f"Çeldirici & Kök Stratejisi: {distractor_mode}\n"
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
                    f"Soru paketi oluşturuldu. Kalan soru kotası: "
                    f"{int(st.session_state.remaining_question_quota):,}".replace(",", ".")
                )
            except Exception as exc:
                st.session_state["remaining_question_quota"] = int(
                    st.session_state.get("remaining_question_quota", 0)
                ) + int(question_count)
                save_persisted_state()
                st.error(f"Soru üretim hatası oluştu: {exc}")

    if st.session_state.generated_test_text:
        st.markdown("---")
        st.markdown("### Üretilen Test ve Redaksiyon Çıktısı")
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
                label="Testi DOCX Olarak İndir",
                data=docx_bytes,
                file_name=f"kursun_kalem_{branch.lower().replace(' ', '_')}_test.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with act_c2:
            if st.button("Mini Deneme Olarak Çöz", use_container_width=True):
                st.session_state["portal_exam_source"] = st.session_state.generated_test_text
                record_activity("Köprü", "Test Öğrenci Portalı mini denemesine aktarıldı")
                navigate_to("Öğrenci Portalı & Çözüm")
        with act_c3:
            if st.button("İlk Soruyu Çözüm Rehberine Gönder", use_container_width=True):
                parsed_q = parse_quiz_questions(st.session_state.generated_test_text)
                st.session_state["portal_question_draft"] = (
                    parsed_q[0]["full_text"] if parsed_q else st.session_state.generated_test_text[:1000]
                )
                navigate_to("Öğrenci Portalı & Çözüm")

# ==============================================================================
# 3. ÖĞRENCİ PORTALI & ÇÖZÜM
# ==============================================================================
elif current_page == "Öğrenci Portalı & Çözüm":
    tab_exam, tab_hint = st.tabs(["İnteraktif Mini Deneme & Net Analizi", "Adımlı Çözüm Asistanı"])

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
            st.markdown(f"#### Aktif Mini Deneme ({len(questions)} Soru)")
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

            if st.button("Denemeyi Bitir ve Netimi Hesapla", type="primary"):
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
                        status_txt = "Doğru"
                    else:
                        wrong_cnt += 1
                        wrong_questions.append(q)
                        status_txt = "Yanlış"

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
                    if st.button("Yanlış Yaptığım İlk Soruyu Çözüm Rehberine Aktar"):
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

        if st.button("Çözüm Rehberi Oluştur", type="primary", key="btn_solve_guide"):
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
                    st.markdown("#### Çözüm Analizi")
                    st.markdown(ans)
                except Exception as exc:
                    st.error(f"Çözüm üretilirken hata oluştu: {exc}")

# ==============================================================================
# 4. DİJİMENTOR
# ==============================================================================
elif current_page == "Dijimentor":
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Dijimentor — Akıllı Sınav &amp; Performans Mentorluğu</div>
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
    st.markdown("#### Hızlı Mentorluk Senaryoları")
    p_cols = st.columns(len(DIJIMENTOR_PRESETS) + 1)
    for idx, (preset_title, preset_data) in enumerate(DIJIMENTOR_PRESETS.items()):
        with p_cols[idx]:
            if st.button(preset_title, key=f"preset_{idx}", use_container_width=True):
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
        if st.button("Planı Temizle", key="clear_dk_plan", use_container_width=True):
            st.session_state.dijimentor_plan = ""
            st.session_state["dk_active_preset"] = ""
            st.rerun()

    st.markdown("---")
    st.markdown("#### Öğrenci Hedef ve Çalışma Profili")

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
        <div style="margin-top:0.65rem;padding-top:0.75rem;border-top:1px solid #DCE3EC;">
            <div style="font-size:0.75rem;font-weight:600;color:#475569;text-transform:uppercase;letter-spacing:0.04em;">
                Önerilen Haftalık Kapasite Dağılımı ({weekly_hours} Saat)
            </div>
            <div class="kp-dist-bar">
                <div style="width:45%;background:#162B54;"></div>
                <div style="width:30%;background:#C48F1B;"></div>
                <div style="width:25%;background:#94A3B8;"></div>
            </div>
            <div class="kp-dist-legend">
                <span>Soru &amp; Muhakeme: <b>{q_hrs} sa</b> (%45)</span>
                <span>Konu Tekrarı: <b>{r_hrs} sa</b> (%30)</span>
                <span>Deneme &amp; Analiz: <b>{a_hrs} sa</b> (%25)</span>
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

    if st.button("Dijimentor Planı Oluştur", type="primary"):
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
                save_persisted_state()
                st.error(f"Dijimentor planı oluşturulurken hata oluştu: {exc}")

    if st.session_state.dijimentor_plan:
        st.markdown("---")
        st.markdown("### Kişiselleştirilmiş Dijimentor Mentorluk Raporu")
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
                label="Mentorluk Planını DOCX İndir",
                data=docx_coach,
                file_name="dijimentor_calisma_plani.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with d_col2:
            st.download_button(
                label="Planı Metin (.txt) İndir",
                data=txt_report.encode("utf-8"),
                file_name="dijimentor_calisma_plani.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with d_col3:
            if st.button("Zayıf Konulardan Soru Üret", use_container_width=True):
                st.session_state["factory_focus_topic"] = weak_topics.split(",")[0].strip()
                navigate_to("Soru & Redaksiyon Fabrikası")

# ==============================================================================
# 5. VELİ TAKİP PANELİ
# ==============================================================================
elif current_page == "Veli Takip Paneli":
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
        st.markdown("#### Haftalık Doğruluk İvmesi (%)")
        st.area_chart(
            {"1. Hafta": 71, "2. Hafta": 74, "3. Hafta": 77, "4. Hafta": 81},
            color="#162B54",
            height=190,
        )
    with g_col2:
        st.markdown("#### Ders Bazlı Başarı Oranı (%)")
        st.bar_chart(
            {"Türkçe": 84, "Matematik": 72, "Fen": 82, "Sosyal": 86, "İngilizce": 79},
            color="#162B54",
            height=190,
        )

    render_parent_notification_center()

# ==============================================================================
# 6. ÖĞRENCİ RÖNTGENİ
# ==============================================================================
elif current_page == "Öğrenci Röntgeni":
    render_student_xray_panel()

# ==============================================================================
# 7. B2B LİSANS & KURUM
# ==============================================================================
elif current_page == "B2B Lisans & Kurum":
    tab_b2b_form, tab_b2b_zumre = st.tabs(
        ["Kurumsal Lisans Teklif Stüdyosu", "Zümre Soru Dağıtımı & Kurumsal Rapor"]
    )

    with tab_b2b_form:
        with st.form("b2b_inquiry_form"):
            c1, c2 = st.columns(2)
            with c1:
                org_name = st.text_input("Kurum / Okul Adı", value="Burak Aygören Akademi")
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

            submitted = st.form_submit_button("Kurumsal Lisans Teklifi Hazırla", type="primary")
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
                label="Kurumsal Lisans Teklifini DOCX İndir",
                data=b2b_docx,
                file_name="kursun_kalem_kurumsal_teklif.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    with tab_b2b_zumre:
        org_cache = st.session_state.get(
            "b2b_org_cache", {"org": "Burak Aygören Akademi", "teachers": 15, "students": 250}
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

        st.markdown("#### Branş Zümreleri Arası Soru Kotası Dağılımı")
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
            st.bar_chart(chart_map, color="#162B54", height=225)

# ==============================================================================
# 8. PAKETLER & ABONELİK VİTRİNİ
# ==============================================================================
elif current_page == "Paketler & Abonelik Vitrini":
    total_pkg_count = len(SUBSCRIPTION_TIERS)

    current_bag = get_bag_items()
    active_bag_summary = render_school_backpack_panel()

    st.markdown("#### Segment Filtresi")
    selected_segment = st.segmented_control(
        "Segment Filtresi",
        options=["Tümü", "Öğrenci", "YKS", "Veli", "Kurum / B2B"],
        default="Tümü",
        label_visibility="collapsed",
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
                bag_tag = f" • Çantada: {in_bag_qty}" if in_bag_qty > 0 else ""
                feats_li = "".join([f"<li>• {html.escape(feat)}</li>" for feat in tier["features"]])
                st.markdown(
                    f"""
                    <div class="kp-pkg-card">
                        <div class="kp-pkg-top">
                            <span class="kp-pkg-seg">{html.escape(tier["segment"])}</span>
                            <span class="kp-pkg-badge">{html.escape(tier["badge"])}{html.escape(bag_tag)}</span>
                        </div>
                        <div class="kp-pkg-title">{html.escape(tier["name"])}</div>
                        <div class="kp-pkg-price">
                            {tier["monthly_price"]:,} ₺ <span>/ ay</span>
                        </div>
                        <div class="kp-pkg-desc">{html.escape(tier["desc"])}</div>
                        <ul class="kp-pkg-feats">{feats_li}</ul>
                    </div>
                    """.replace(",", "."),
                    unsafe_allow_html=True,
                )
                st.button(
                    "Çantaya Ekle",
                    key=f"select_{tier['id']}",
                    type="primary" if in_bag_qty > 0 else "secondary",
                    use_container_width=True,
                    on_click=cb_add_tier_to_bag,
                    args=(tier["id"], 1),
                )

    st.markdown("---")
    st.markdown(f"#### {total_pkg_count} Paket Karşılaştırma Matrisi")
    matrix_rows = [
        {
            "Paket Adı": t["name"],
            "Segment": t["segment"],
            "Etiket": t["badge"],
            "Çantadaki Adet": current_bag.get(t["id"], 0),
            "Aylık Soru Kotası": f"{t['monthly_quota']:,}".replace(",", "."),
            "Dijimentor Plan": f"{t['coach_quota']}/ay",
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
