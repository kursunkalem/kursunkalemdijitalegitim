import base64
import html
import io
import json
import os
import re
import time
import uuid
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
# SAYFA YAPILANDIRMASI VE KURUMSAL SABİTLER (2028 KURGUSU)
# ==============================================================================
st.set_page_config(
    page_title="KURŞUN KALEM DİJİTAL OFİS • 2028",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "KURŞUN KALEM DİJİTAL OFİS"
BRAND_SHORT = "Kurşun Kalem"
BRAND_TAGLINE = "2028 Hedefli Soru & Redaksiyon Fabrikası • Dijimentor • Veli & B2B Dijital Ofis"
DEFAULT_MODEL = "gemini-2.5-flash"
FALLBACK_MODEL = "gemini-3.5-flash-lite"
STATE_FILE_PATH = Path(".kursun_kalem_ofis_2028_state.json")
STATE_BACKUP_PATH = Path(".kursun_kalem_ofis_2028_state.bak.json")
EXTERNAL_CSS_PATH = Path("assets/style.css")

GRADE_LEVELS_2028: List[str] = [
    "5. Sınıf",
    "6. Sınıf",
    "7. Sınıf",
    "8. Sınıf (LGS)",
    "9. Sınıf",
    "10. Sınıf",
    "11. Sınıf (YKS)",
    "12. Sınıf (YKS)",
]

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
  <polygon points="256,82 362,130 256,178 150,130" fill="url(#navyGrad)" stroke="url(#goldGrad)" stroke-width="5"/>
  <polygon points="190,148 256,178 322,148 322,188 256,216 190,188" fill="url(#navyGrad)"/>
  <text x="256" y="314" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="44" font-weight="800" fill="#0b1836">DİJİTAL OFİS</text>
  <text x="256" y="351" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="22" font-weight="700" fill="#0f172a" letter-spacing="4">KURŞUN KALEM</text>
  <g transform="translate(432, 342) rotate(24)">
    <rect x="-12" y="-110" width="24" height="96" rx="3" fill="url(#pencilBody)" stroke="#1e293b" stroke-width="2"/>
    <polygon points="-12,-14 12,-14 0,24" fill="#fde68a" stroke="#1e293b" stroke-width="2"/>
    <polygon points="-3.5,12 3.5,12 0,24" fill="#0f172a"/>
  </g>
</svg>"""

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
  <text x="130" y="184" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="800" fill="#0f172a">OFİS</text>
  <text x="130" y="196" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="7.5" font-weight="700" fill="#475569">ÇANTASI</text>
  <rect x="72" y="242" width="116" height="18" rx="9" fill="#090d16"/>
</svg>"""

HERO_SHOWCASE_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 440 330" width="440" height="330">
  <defs>
    <linearGradient id="bgCardGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#EEF2FF"/>
      <stop offset="100%" stop-color="#E0E7FF"/>
    </linearGradient>
    <linearGradient id="navyBanner" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#1E2952"/>
      <stop offset="100%" stop-color="#2563EB"/>
    </linearGradient>
    <linearGradient id="goldBadgeGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FACC15"/>
      <stop offset="100%" stop-color="#D97706"/>
    </linearGradient>
    <linearGradient id="pencilWoodGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FDE047"/>
      <stop offset="100%" stop-color="#EAB308"/>
    </linearGradient>
  </defs>

  <rect x="22" y="18" width="396" height="292" rx="24" fill="url(#bgCardGrad)" stroke="#C7D2FE" stroke-width="1.5"/>

  <rect x="44" y="42" width="214" height="244" rx="18" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="1.5"/>
  <rect x="44" y="42" width="214" height="58" rx="18" fill="url(#navyBanner)"/>
  <rect x="44" y="76" width="214" height="24" fill="url(#navyBanner)"/>
  <text x="62" y="68" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="800" fill="#FACC15" letter-spacing="1.2">KURŞUN KALEM • 2028</text>
  <text x="62" y="87" font-family="Inter, Arial, sans-serif" font-size="13" font-weight="800" fill="#FFFFFF">Ölçme &amp; Redaksiyon Ofisi</text>

  <rect x="60" y="114" width="182" height="42" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <circle cx="79" cy="135" r="8" fill="#2563EB"/>
  <text x="94" y="132" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#1E2952">Maarif Kazanım Uyumu</text>
  <text x="94" y="146" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">LGS &amp; YKS 2028 Müfredatı</text>

  <rect x="60" y="164" width="182" height="42" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <circle cx="79" cy="185" r="8" fill="#4F46E5"/>
  <text x="94" y="182" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#1E2952">Çeldirici &amp; Kök Denetimi</text>
  <text x="94" y="196" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">Çift Katmanlı Editör Kontrolü</text>

  <rect x="60" y="214" width="182" height="52" rx="10" fill="#EFF6FF" stroke="#BFDBFE"/>
  <text x="74" y="233" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="700" fill="#1D4ED8">HEDEF NET &amp; İSABET İVMESİ</text>
  <rect x="74" y="243" width="152" height="8" rx="4" fill="#DBEAFE"/>
  <rect x="74" y="243" width="124" height="8" rx="4" fill="#2563EB"/>

  <rect x="276" y="42" width="120" height="114" rx="18" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="1.5"/>
  <circle cx="336" cy="82" r="22" fill="url(#goldBadgeGrad)"/>
  <polygon points="336,68 340,77 350,78 342,85 345,95 336,90 327,95 330,85 322,78 332,77" fill="#1E2952"/>
  <text x="336" y="121" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="11" font-weight="800" fill="#1E2952">2028 HEDEF</text>
  <text x="336" y="136" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">Dijimentor &amp; Röntgen</text>

  <rect x="276" y="170" width="120" height="116" rx="18" fill="#1E2952"/>
  <g transform="translate(336, 214) rotate(-32)">
    <rect x="-8" y="-34" width="16" height="52" rx="2" fill="url(#pencilWoodGrad)"/>
    <rect x="-8" y="-38" width="16" height="6" rx="1.5" fill="#F43F5E"/>
    <polygon points="-8,18 8,18 0,34" fill="#FDE68A"/>
    <polygon points="-2.5,28 2.5,28 0,34" fill="#0F172A"/>
  </g>
  <text x="336" y="266" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="10" font-weight="800" fill="#FACC15" letter-spacing="0.8">KURŞUN KALEM</text>
  <text x="336" y="279" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="600" fill="#E2E8F0">Dijital Ofis Vitrini</text>
</svg>"""


def svg_to_data_uri(svg_str: str) -> str:
    encoded = base64.b64encode(svg_str.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def get_embedded_logo_png_bytes() -> Optional[bytes]:
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
        draw.text((160, 182), "DİJİTAL OFİS", fill=(11, 24, 54, 255), anchor="mm")
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return None


LGS_2028_TARGET = datetime(2028, 6, 11, 9, 30, 0)
YKS_2028_TARGET = datetime(2028, 6, 17, 10, 15, 0)

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
        "Genel Bakış, 2028 Kademe Vitrini & Kişisel Akış",
        "Kademe seçimi, eksik kazanım öncelik sırası, 2028 LGS ve YKS geri sayımı ile dijital ofis özeti.",
    ),
    "Soru & Redaksiyon Fabrikası": (
        "Maarif Modeli Soru & Kalite Stüdyosu (2028)",
        "5–12. sınıf müfredatı, kademeli zorluk serisi, çeldirici denetimi, soru havuzu ve DOCX çıktı.",
    ),
    "Öğrenci Portalı & Çözüm": (
        "İnteraktif Mini Deneme, Hata Defteri & Çözüm Rehberi",
        "Süre takipli mini deneme çözün, netinizi görün, yanlışları Hata Defteri'ne alın ve adımlı çözüm inceleyin.",
    ),
    "Dijimentor": (
        "Akıllı Sınav, Faz Planı & Performans Mentorluğu (2028)",
        "2028 sınav takvimine göre aşamalı kamp ve haftalık aksiyon planı üretin.",
    ),
    "Veli Takip Paneli": (
        "Veli Karne Özeti & Bilgilendirme Merkezi",
        "Haftalık değişim kartları, kazanım ivmesi ve hazır veli bilgilendirme mesajları.",
    ),
    "Öğrenci Röntgeni": (
        "Öğrenci Röntgeni & Akademik Check-Up",
        "Kazanım bazlı röntgen taraması, hata kök neden analizi ve modüller arası tek tık aktarım.",
    ),
    "B2B Lisans & Kurum": (
        "Kurumsal Lisans, Zümre & Deneme Planlayıcı",
        "Koltuk/kota planlaması, zümre dağılımı, kurum içi deneme takvimi ve DOCX teklif çıktısı.",
    ),
    "Paketler & Abonelik Vitrini": (
        "Abonelik Paketleri & Ofis Çantası",
        "Öğrenci, 2028 YKS, Veli ve Kurumsal paketleri çantanıza ekleyin, kotaları yönetin ve iyzico planını inceleyin.",
    ),
}

HERO_PILL_CHIPS: List[Dict[str, str]] = [
    {"label": "YKS 2028", "grade": "12. Sınıf (YKS)", "exam": "2028 TYT - AYT", "page": "Dijimentor"},
    {"label": "LGS 2028", "grade": "8. Sınıf (LGS)", "exam": "2028 LGS", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "11. Sınıf", "grade": "11. Sınıf (YKS)", "exam": "2028 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "10. Sınıf", "grade": "10. Sınıf", "exam": "2028 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "9. Sınıf", "grade": "9. Sınıf", "exam": "2028 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "7. Sınıf", "grade": "7. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "6. Sınıf", "grade": "6. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "5. Sınıf", "grade": "5. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "Veli & Kurum", "grade": "8. Sınıf (LGS)", "exam": "2028 LGS", "page": "Veli Takip Paneli"},
    {"label": "Zümre & B2B", "grade": "8. Sınıf (LGS)", "exam": "2028 LGS", "page": "B2B Lisans & Kurum"},
]

KADEME_VITRIN_KARTLARI: List[Dict[str, str]] = [
    {
        "tag": "5–7. SINIF • 2028",
        "cta": "5–7. Sınıf • Aç",
        "title": "Ortaokul & Maarif Temeli",
        "desc": "Kazanım odaklı sorular, okuma hızı ve düzenli çalışma alışkanlığı.",
        "accent": "#2563EB",
        "grade": "7. Sınıf",
        "exam": "5-7. Sınıf Akademik Gelişim",
        "target_page": "Soru & Redaksiyon Fabrikası",
    },
    {
        "tag": "2028 LGS",
        "cta": "2028 LGS • Aç",
        "title": "8. Sınıf LGS Hazırlık",
        "desc": "Sözel-sayısal muhakeme, çeldirici analizi ve net yükselten deneme akışı.",
        "accent": "#4F46E5",
        "grade": "8. Sınıf (LGS)",
        "exam": "2028 LGS",
        "target_page": "Soru & Redaksiyon Fabrikası",
    },
    {
        "tag": "2028 YKS",
        "cta": "2028 YKS • Aç",
        "title": "TYT & AYT Lise Kampı",
        "desc": "Süre yönetimi, paragraf-analitik denge ve derece hedefleyen mentorluk.",
        "accent": "#7C3AED",
        "grade": "12. Sınıf (YKS)",
        "exam": "2028 TYT - AYT",
        "target_page": "Dijimentor",
    },
    {
        "tag": "VELİ & KURUM • 2028",
        "cta": "Veli & Kurum • Aç",
        "title": "Veli Takip & B2B Ofis",
        "desc": "Haftalık çalışma-soru özeti, veli mesajları ve zümre kota yönetimi.",
        "accent": "#0284C7",
        "grade": "8. Sınıf (LGS)",
        "exam": "2028 LGS",
        "target_page": "Veli Takip Paneli",
    },
]

# Ana sayfadaki sınıf seçici: sekme -> o sekmedeki seçenekler (HERO_PILL_CHIPS etiketleri)
GRADE_GROUPS: Dict[str, List[str]] = {
    "Ortaokul": ["5. Sınıf", "6. Sınıf", "7. Sınıf", "LGS 2028"],
    "Lise": ["9. Sınıf", "10. Sınıf", "11. Sınıf", "YKS 2028"],
    "Veli & Kurum": ["Veli & Kurum", "Zümre & B2B"],
}
CHIP_BY_LABEL: Dict[str, Dict[str, str]] = {c["label"]: c for c in HERO_PILL_CHIPS}

# Sol menü: (başlık, [(sayfa adı, menüde görünen kısa ad, ikon), ...])
NAV_GROUPS: List[Tuple[Optional[str], List[Tuple[str, str, str]]]] = [
    (None, [("Ana Sayfa", "Ana Sayfa", ":material/home:")]),
    ("Modüller", [
        ("Soru & Redaksiyon Fabrikası", "Soru & Redaksiyon", ":material/edit_note:"),
        ("Öğrenci Portalı & Çözüm", "Öğrenci Portalı", ":material/quiz:"),
        ("Dijimentor", "Dijimentor", ":material/event_note:"),
        ("Öğrenci Röntgeni", "Öğrenci Röntgeni", ":material/monitor_heart:"),
    ]),
    ("Veli & Kurum", [
        ("Veli Takip Paneli", "Veli Takip Paneli", ":material/family_restroom:"),
        ("B2B Lisans & Kurum", "B2B Lisans & Kurum", ":material/apartment:"),
    ]),
    ("Hesap", [
        ("Paketler & Abonelik Vitrini", "Paketler", ":material/shopping_bag:"),
    ]),
]

# Ana sayfadaki modül kartları: (kısaltma, başlık, açıklama, hedef sayfa, arka plan, yazı rengi)
MODULE_CARDS: List[Tuple[str, str, str, str, str, str]] = [
    ("SR", "Soru & Redaksiyon", "Kazanım etiketli soru üretimi ve çift katmanlı editör kontrolü.", "Soru & Redaksiyon Fabrikası", "#EFF6FF", "#1D4ED8"),
    ("DM", "Dijimentor", "Seviyene göre her hafta yenilenen kişisel çalışma planı.", "Dijimentor", "#F0FDF4", "#15803D"),
    ("MD", "Mini Deneme", "Süre takipli denemeler ve otomatik hata defteri.", "Öğrenci Portalı & Çözüm", "#FEF3C7", "#B45309"),
    ("ÖR", "Öğrenci Röntgeni", "Konu konu güçlü ve zayıf yönlerin net haritası.", "Öğrenci Röntgeni", "#F5F3FF", "#6D28D9"),
    ("VT", "Veli Takip", "Haftalık özet, ilerleme ve veli mesajları tek yerde.", "Veli Takip Paneli", "#FFF1F2", "#BE123C"),
]

# Modül kartlarındaki çizgi ikonlar (kısaltma -> SVG içi)
_SVG = '<svg aria-hidden="true" width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{}</svg>'
MODULE_ICONS: Dict[str, str] = {
    "SR": _SVG.format('<path d="M17 3l4 4L8 20H4v-4L17 3z"/><path d="M14 6l4 4"/>'),
    "DM": _SVG.format('<rect x="3" y="4" width="18" height="17" rx="3"/><path d="M3 9h18M8 2v4M16 2v4M8 14l2 2 4-4"/>'),
    "MD": _SVG.format('<circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2.5M9 2h6"/>'),
    "ÖR": _SVG.format('<path d="M3 12h4l3-8 4 16 3-8h4"/>'),
    "VT": _SVG.format('<circle cx="9" cy="8" r="3"/><circle cx="17" cy="10" r="2.5"/><path d="M3 20c0-3 3-5 6-5s6 2 6 5M15 20c0-2 1.5-3.5 4-3.5"/>'),
}

# Eğlenceli tema açık/kapalı (False yaparsan sade görünüme döner)
FUN_THEME = True

# ==============================================================================
# 5–12. SINIF TAM MÜFREDAT MOTORU (MAARİF MODELİ, LGS & YKS)
# ==============================================================================
_MIDDLE_COMMON: Dict[str, Dict[str, List[str]]] = {
    "Türkçe": {
        "Okuma ve Anlama": [
            "Bağlamdan yararlanarak sözcük ve söz öbeklerinin anlamını tahmin etme",
            "Metindeki ana düşünce, yardımcı düşünceler ve başlık ilişkisini kurma",
            "Sözel mantık, tablo/grafik yorumlama ve muhakeme gerektiren metinleri çözümleme",
            "Metinler arası karşılaştırma yapma ve çıkarımda bulunma",
        ],
        "Dil Bilgisi ve Yazım": [
            "Sözcük türleri, fiilimsiler ve cümle ögeleri işlevini belirleme",
            "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            "Cümlede anlam ilişkileri ve anlatım akıcılığını denetleme",
        ],
    },
    "Matematik": {
        "Sayılar ve İşlemler": [
            "Doğal sayılar, tam sayılar, kesirler ve ondalık gösterimlerle çok adımlı işlemler",
            "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
            "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve cebirsel ifadeler",
        ],
        "Geometri ve Veri Analizi": [
            "Doğrular, açılar, çokgenler, üçgenler ve geometrik cisimler",
            "Veri toplama, grafik yorumlama, olasılık ve merkezi eğilim ölçüleri",
        ],
    },
    "Fen Bilimleri": {
        "Fiziksel Süreçler ve Kuvvet": [
            "Kuvvet, hareket, basınç, iş-enerji ve basit makineler",
            "Mevsimlerin oluşumu, ışık, ses ve elektrik enerjisi uygulamaları",
        ],
        "Madde, Doğa ve Canlılar": [
            "Maddenin yapısı, karışımlar, kimyasal tepkimeler ve asit-baz kavramları",
            "Hücre, sistemler, DNA-genetik kod, kalıtım ve ekoloji/sürdürülebilirlik",
        ],
    },
    "Sosyal Bilgiler / İnkılap Tarihi": {
        "Tarih ve Kültür": [
            "Birey, toplum, kültür ve miras; Millî Mücadele ve Atatürk İlkeleri",
            "Cumhuriyet kazanımları, çağdaş gelişmeler ve tarihsel kanıt okuma",
        ],
        "İnsan, Yer ve Çevre": [
            "Harita okuryazarlığı, nüfus, ekonomik faaliyetler ve çevre bilinci",
            "Haklar, sorumluluklar, demokratik süreçler ve küresel bağlantılar",
        ],
    },
    "İngilizce": {
        "Reading & Comprehension": [
            "Contextual vocabulary, inference and main idea in short texts",
            "Everyday communication, invitations, preferences and plans",
        ],
        "Language & Reasoning": [
            "Tenses, modals, comparatives and cohesive connectors in context",
            "Table/chart interpretation and situational dialogue completion",
        ],
    },
}

_HIGH_COMMON: Dict[str, Dict[str, List[str]]] = {
    "Türk Dili ve Edebiyatı": {
        "Anlama ve Yorumlama": [
            "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi",
            "Dönem/zihniyet ilişkisi, edebi sanatlar ve metinler arası etkileşim",
        ],
        "Dil Bilgisi ve Yazım": [
            "Sözcükte ve cümlede anlam, paragraf kurgusu ve çeldirici analizi",
            "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu",
        ],
    },
    "Matematik": {
        "Temel Kavramlar ve Cebir": [
            "Temel kavramlar, bölünebilme, rasyonel/üslü/köklü ifadeler ve mutlak değer",
            "Denklemler, eşitsizlikler, oran-orantı, problemler ve fonksiyonlar",
        ],
        "Analitik, Geometri ve İleri Konular": [
            "Üçgenler, çokgenler, çember-daire, katı cisimler ve analitik geometri",
            "Polinomlar, logaritma, diziler, trigonometri, limit, türev ve integral",
        ],
    },
    "Fizik": {
        "Mekanik ve Hareket": [
            "Fizik bilimine giriş, hareket, Newton hareket yasaları, iş-güç-enerji",
            "İtme-momentum, tork-denge, basit makineler, kaldırma kuvveti ve basınç",
        ],
        "Elektrik, Manyetizma ve Dalgalar": [
            "Elektrostatik, elektrik akımı, manyetizma, elektromanyetik indüksiyon",
            "Dalgalar, optik, çembersel hareket, harmonik hareket ve modern fizik",
        ],
    },
    "Kimya": {
        "Madde ve Tepkimeler": [
            "Kimya bilimi, atom ve periyodik sistem, kimyasal türler arası etkileşimler",
            "Maddenin halleri, mol kavramı, kimyasal hesaplamalar ve tepkime hızı/denge",
        ],
        "Çözeltiler ve Enerji": [
            "Karışımlar, asitler-bazlar-tuzlar, kimya ve enerji, elektrokimya ve organik kimya",
        ],
    },
    "Biyoloji": {
        "Canlıların Temelleri ve Sistemler": [
            "Yaşam bilimi biyoloji, hücre zarından madde geçişi, canlıların sınıflandırılması",
            "Hücre bölünmeleri, kalıtım, ekosistem ekolojisi ve insan fizyolojisi sistemleri",
        ],
        "Genden Proteine ve Enerji": [
            "Nükleik asitler, protein sentezi, fotosentez ve hücresel solunum",
        ],
    },
    "Tarih": {
        "Tarihsel Süreç ve İnkılap": [
            "Tarih ve zaman, insanlığın ilk dönemleri, İslam tarihi ve Türk-İslam devletleri",
            "Osmanlı kurumları ve diplomasi, Kurtuluş Savaşı ve Atatürk İlke ve İnkılapları",
        ],
    },
    "Coğrafya": {
        "Doğal ve Beşeri Sistemler": [
            "Doğa ve insan, coğrafi koordinat ve iklim bilgisi, iç-dış kuvvetler",
            "Nüfus, yerleşme, ekonomik faaliyetler, bölgeler ve küresel çevre sorunları",
        ],
    },
}

FULL_CURRICULUM_2028: Dict[str, Dict[str, Dict[str, List[str]]]] = {
    "5. Sınıf": _MIDDLE_COMMON,
    "6. Sınıf": _MIDDLE_COMMON,
    "7. Sınıf": _MIDDLE_COMMON,
    "8. Sınıf (LGS)": {
        "Türkçe": {
            "Okuma ve Muhakeme": [
                "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
                "Metindeki ana düşünce, çıkarım ve başlık/görsel ilişkisini kurma",
                "Metinler arası karşılaştırma yapma ve yargı/kanıt türlerini ayırt etme",
            ],
            "Dil Bilgisi ve Yazım": [
                "Fiilimsiler (isim-fiil, sıfat-fiil, zarf-fiil) ve cümle ögeleri",
                "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            ],
        },
        "Matematik": {
            "Sayılar ve Cebir": [
                "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
                "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler",
                "Cebirsel ifadeler, özdeşlikler ve birinci dereceden denklemler",
            ],
            "Geometri, Olasılık ve Veri": [
                "Üçgenler, kenarortay/açıortay/yükseklik, eşlik-benzerlik ve dönüşüm geometrisi",
                "Olasılık ve istatistiksel veri analizi / grafik dönüştürme",
            ],
        },
        "Fen Bilimleri": {
            "2028 LGS Fen Kazanımları": [
                "Mevsimlerin oluşumu ve iklim/hava hareketleri",
                "DNA, genetik kod, mutasyon, modifikasyon, adaptasyon ve biyoteknoloji",
                "Basınç (katı, sıvı, gaz) ve madde-endüstri uygulamaları",
                "Basit makineler, enerji dönüşümleri, elektrik yükleri ve elektrik enerjisi",
            ],
        },
        "T.C. İnkılap Tarihi ve Atatürkçülük": {
            "Millî Mücadele ve Çağdaş Türkiye": [
                "Bir Kahraman Doğuyor ve Millî Uyanış: Bağımsızlık Yolunda Atılan Adımlar",
                "Millî Bir Destan: Ya İstiklal Ya Ölüm ve Atatürkçülük/Çağdaşlaşan Türkiye",
                "Demokratikleşme çabaları ve Atatürk'ten Sonraki Dış Politika",
            ],
        },
        "İngilizce": {
            "2028 LGS İngilizce Temaları": [
                "Friendship, Teen Life, In the Kitchen, On the Phone, The Internet",
                "Tourism, Chores, Science, Natural Forces & Reading Comprehension",
            ],
        },
    },
    "9. Sınıf": _HIGH_COMMON,
    "10. Sınıf": _HIGH_COMMON,
    "11. Sınıf (YKS)": _HIGH_COMMON,
    "12. Sınıf (YKS)": _HIGH_COMMON,
}


def get_branches_for_grade_2028(grade: str) -> List[str]:
    g_map = FULL_CURRICULUM_2028.get(grade)
    if g_map:
        return list(g_map.keys())
    return list(_MIDDLE_COMMON.keys())


def get_areas_for_grade_branch_2028(grade: str, branch: str) -> List[str]:
    g_map = FULL_CURRICULUM_2028.get(grade, {})
    if branch in g_map:
        return list(g_map[branch].keys())
    for other_g in FULL_CURRICULUM_2028.values():
        if branch in other_g:
            return list(other_g[branch].keys())
    return ["Genel Kazanımlar"]


def get_outcomes_for_2028(grade: str, branch: str, area: str) -> List[str]:
    g_map = FULL_CURRICULUM_2028.get(grade, {})
    if branch in g_map and area in g_map[branch]:
        return list(g_map[branch][area])
    for other_g in FULL_CURRICULUM_2028.values():
        if branch in other_g and area in other_g[branch]:
            return list(other_g[branch][area])
    return ["Kazanım odaklı beceri ve muhakeme sorusu üretimi"]


# ==============================================================================
# DİĞER SABİTLER (RÖNTGEN, DİJİMENTOR, ABONELİK, VELİ MESAJLARI)
# ==============================================================================
OGRENCI_RONTGEN_PROFILLERI: Dict[str, Dict[str, Any]] = {
    "Elif Öztürk (8-A | 2028 LGS)": {
        "basari_endeksi": "%76.4",
        "basari_delta": "+4.2% (Son 2 hafta)",
        "kritik_konu": "3 Kazanım",
        "kritik_delta": "-2 İyileşme",
        "hiz": "1.35 dk/soru",
        "hiz_delta": "İdeal Tempo",
        "hedef_puan": "454 / 480",
        "hedef_delta": "26 Puan Fark",
        "kazanimlar": [
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme", "Çözülen": 58, "Başarı (%)": 46, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "Metindeki ana düşünce, çıkarım ve başlık/görsel ilişkisini kurma", "Çözülen": 135, "Başarı (%)": 86, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
            {"Ders / Alan": "Matematik", "Kazanım / Konu": "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri", "Çözülen": 74, "Başarı (%)": 52, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Fen Bilimleri", "Kazanım / Konu": "Basınç (katı, sıvı, gaz) ve madde-endüstri uygulamaları", "Çözülen": 62, "Başarı (%)": 74, "Teşhis Durumu": "🟡 Geliştirilmeli"},
            {"Ders / Alan": "T.C. İnkılap Tarihi ve Atatürkçülük", "Kazanım / Konu": "Millî Mücadele ve Atatürk İlkeleri", "Çözülen": 48, "Başarı (%)": 88, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 36,
            "Uzun Metin / Odak Kaybı": 30,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 22,
            "Süre Yetişmemesi": 12,
        },
        "recete_ozet": (
            "Sözel Mantık & Muhakeme (%46) ve EBOB-EKOK modellemeleri (%52) kazanımlarında tablo/şema kurma ve "
            "çok adımlı yönerge takibi zayıf. Yanlışların %22'si soru kökündeki olumsuz ifadelerin hızlı okunmasından kaynaklanıyor."
        ),
        "recete_aksiyon": "Soru & Redaksiyon Fabrikası üzerinden orta-zor seviyede Sözel Mantık ve Matematik telafi soruları üretilip atanmaya hazır.",
        "hedef_konu": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
        "hedef_brans": "Türkçe",
        "hedef_sinif": "8. Sınıf (LGS)",
    },
    "Zeynep Kaya (8-B | 2028 LGS)": {
        "basari_endeksi": "%83.8",
        "basari_delta": "+5.6% (Son 2 hafta)",
        "kritik_konu": "2 Kazanım",
        "kritik_delta": "-1 İyileşme",
        "hiz": "1.18 dk/soru",
        "hiz_delta": "Hızlı ve Dengeli",
        "hedef_puan": "468 / 480",
        "hedef_delta": "12 Puan Fark",
        "kazanimlar": [
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "Metinler arası karşılaştırma yapma ve yargı/kanıt türlerini ayırt etme", "Çözülen": 64, "Başarı (%)": 81, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
            {"Ders / Alan": "Matematik", "Kazanım / Konu": "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler", "Çözülen": 72, "Başarı (%)": 55, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Fen Bilimleri", "Kazanım / Konu": "Mevsimlerin oluşumu ve iklim/hava hareketleri", "Çözülen": 54, "Başarı (%)": 89, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "Fiilimsiler (isim-fiil, sıfat-fiil, zarf-fiil) ve cümle ögeleri", "Çözülen": 86, "Başarı (%)": 71, "Teşhis Durumu": "🟡 Geliştirilmeli"},
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 27,
            "Uzun Metin / Odak Kaybı": 21,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 35,
            "Süre Yetişmemesi": 17,
        },
        "recete_ozet": (
            "Üslü ve kareköklü ifadeler (%55) alanında yaklaşık değer ve yeni nesil kurgu adımlarında sapma görülüyor. "
            "Hataların %35'i soru kökündeki sınırlandırıcı ifadelerin gözden kaçmasından kaynaklanıyor."
        ),
        "recete_aksiyon": "Olumsuz soru köklü Kareköklü İfade ve dikkat odaklı muhakeme seti oluşturularak Öğrenci Portalı'na gönderilmeye hazır.",
        "hedef_konu": "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler",
        "hedef_brans": "Matematik",
        "hedef_sinif": "8. Sınıf (LGS)",
    },
    "Emir Demir (12-A | 2028 YKS)": {
        "basari_endeksi": "%71.2",
        "basari_delta": "+3.8% (Son 2 hafta)",
        "kritik_konu": "4 Kazanım",
        "kritik_delta": "-1 İyileşme",
        "hiz": "1.52 dk/soru",
        "hiz_delta": "Süre Baskısı Var",
        "hedef_puan": "442 / 500",
        "hedef_delta": "58 Puan Fark",
        "kazanimlar": [
            {"Ders / Alan": "Türk Dili ve Edebiyatı", "Kazanım / Konu": "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi", "Çözülen": 145, "Başarı (%)": 66, "Teşhis Durumu": "🟡 Geliştirilmeli"},
            {"Ders / Alan": "Türk Dili ve Edebiyatı", "Kazanım / Konu": "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu", "Çözülen": 78, "Başarı (%)": 49, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Matematik", "Kazanım / Konu": "Polinomlar, logaritma, diziler, trigonometri, limit, türev ve integral", "Çözülen": 88, "Başarı (%)": 54, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Fizik", "Kazanım / Konu": "İtme-momentum, tork-denge, basit makineler, kaldırma kuvveti ve basınç", "Çözülen": 56, "Başarı (%)": 84, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 34,
            "Uzun Metin / Odak Kaybı": 24,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 16,
            "Süre Yetişmemesi": 26,
        },
        "recete_ozet": (
            "Dil Bilgisi (%49) ve İleri Matematik (%54) alanlarında kural transferi ve süre yönetimi kritik öncelikte. "
            "Süre yetişmemesi (%26) denemenin ikinci yarısında dikkat düşüşüyle birleşiyor."
        ),
        "recete_aksiyon": "Zamanlı Dil Bilgisi + Analitik Muhakeme sorularından oluşan hız ve doğruluk telafi paketi hazır.",
        "hedef_konu": "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu",
        "hedef_brans": "Türk Dili ve Edebiyatı",
        "hedef_sinif": "12. Sınıf (YKS)",
    },
}

ONCELIKLI_EKSIK_SIRASI: List[Dict[str, str]] = [
    {
        "priority": "1. ÖNCELİK • %46 İSABET",
        "pct": 46,
        "topic": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "action": "Tablo/şema kurma ve telafi sorusu",
    },
    {
        "priority": "2. ÖNCELİK • %52 İSABET",
        "pct": 52,
        "topic": "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
        "branch": "Matematik",
        "grade": "8. Sınıf (LGS)",
        "action": "Yeni nesil kurgu & çeldirici analizi",
    },
    {
        "priority": "3. ÖNCELİK • DİKKAT",
        "topic": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "action": "Olumsuz soru kökü hız antrenmanı",
    },
]

DIJIMENTOR_PRESETS: Dict[str, Dict[str, Any]] = {
    "2028 LGS Derece & Net İvmesi": {
        "exam": "2028 LGS",
        "hours": 20,
        "goal": "Fen liseleri ve nitelikli Anadolu liseleri bandı (470+ puan ve 82+ net hedefi)",
        "focus": "Deneme Net Artırma & Zaman Yönetimi",
        "weak": "Sözel mantık ve muhakeme metinleri, cebirsel ifadeler ve çarpanlar-katlar yeni nesil problemleri, süre baskısı altında çeldirici ayıklama",
        "status": "Orta-üst seviye; Türkçe ve Matematikte son 15 soruda süre baskısı yaşıyor ve deneme analizini haftalık sistematik yapamıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
    },
    "2028 YKS TYT-AYT Disiplin Rutini": {
        "exam": "2028 TYT - AYT",
        "hours": 30,
        "goal": "TYT 98+ net ve AYT alanında ilk 15 bin derece hedefi",
        "focus": "Haftalık Dengeli Çalışma Programı",
        "weak": "TYT Türkçe paragraf süresi ve anlam-yorum çeldiricileri, problem kurma hızı, AYT konu tekrar gecikmeleri ve deneme tempo dalgalanması",
        "status": "Düzenli çalışma alışkanlığı var ancak TYT süresini verimli bölme ve hafta sonu deneme analizlerini puan/net aksiyonuna çevirme konusunda gelişime açık",
        "branch": "Türk Dili ve Edebiyatı",
        "grade": "12. Sınıf (YKS)",
        "topic": "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi",
    },
    "Dikkat & Çeldirici Tuzağı Temizliği": {
        "exam": "2028 LGS",
        "hours": 16,
        "goal": "Dikkat ve okuma kaynaklı net kayıplarını %70 azaltmak ve istikrarlı net artışı sağlamak",
        "focus": "Dikkat ve Çeldirici Tuzağı Azaltma",
        "weak": "Soru kökündeki 'ulaşılamaz / çıkarılamaz / değildir' ifadelerini kaçırma, öncül atlama ve çeldirici şıklara erken yönelme",
        "status": "Konu bilgisi yeterli olmasına rağmen hızlı okuma ve aceleci işaretleme nedeniyle bildiği sorularda hata yapıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
    },
}

SUBSCRIPTION_TIERS: List[Dict[str, Any]] = [
    {
        "id": "ogrenci_temel",
        "segment": "Öğrenci",
        "name": "2028 LGS Başlangıç Dijital",
        "monthly_price": 1990,
        "monthly_quota": 400,
        "coach_quota": 30,
        "badge": "Temel",
        "desc": "5–8. sınıf öğrencileri için kazanım tekrarları, çeldirici farkındalığı ve adımlı çözüm desteği.",
        "features": [
            "Aylık 400 Maarif Modeli & LGS soru üretimi",
            "Öğrenci Portalı mini deneme ve ipucu/çözüm asistanı",
            "2028 LGS canlı geri sayım ve haftalık hedef takibi",
            "Çeldirici analizli standart DOCX test çıktısı",
            "Hata defteri ve kazanım özeti",
            "Mobil & masaüstü uyumlu kesintisiz çalışma arayüzü",
        ],
    },
    {
        "id": "ogrenci_pro",
        "segment": "Öğrenci",
        "name": "2028 LGS Hedef Pro",
        "monthly_price": 3290,
        "monthly_quota": 1500,
        "coach_quota": 90,
        "badge": "En Çok Tercih Edilen",
        "desc": "Yeni nesil muhakeme soruları, otomatik redaksiyon, Dijimentor planı ve veli raporu.",
        "features": [
            "Aylık 1.500 yapay zekâ soru ve redaksiyon hakkı",
            "Otomatik redaktör ile çeldirici ve soru kökü denetimi",
            "Dijimentor haftalık akıllı çalışma ve tempo programı",
            "Veli Takip Paneli canlı bildirim ve özet erişimi",
            "Soru köküne uygun başlıklandırılmış DOCX test çıktısı",
            "Yanlış soruları otomatik Hata Defteri'ne aktarma",
        ],
    },
    {
        "id": "ogrenci_elit",
        "segment": "Öğrenci",
        "name": "2028 LGS Tam Donanım Elit",
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
        "name": "2028 LGS Final & Deneme Kampı",
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
        "name": "2028 YKS MF (Sayısal) Akademi",
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
        "name": "2028 YKS EA & Sözel Atölye",
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
        "name": "2028 YKS Derece VIP",
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
        "name": "2028 TYT Hız & Paragraf Sprinti",
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
        "name": "Aile Takip & Rehberlik (2028)",
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
        "name": "Çoklu Kardeş Aile Paketi (2028)",
        "monthly_price": 3690,
        "monthly_quota": 2200,
        "coach_quota": 130,
        "badge": "2 Öğrenci",
        "desc": "İki kardeş için ayrı sınıf/sınav seviyelerinde takip, soru üretimi ve Dijimentor desteği.",
        "features": [
            "2 ayrı öğrenci profili (5–12. sınıf / 2028 LGS & YKS)",
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
        "name": "Aile Rehberlik & Deneme Plus (2028)",
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
        "name": "Öğretmen & Zümre Stüdyosu (2028)",
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
        "name": "Etüt Merkezi & Kurs Paketi (2028)",
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
        "name": "Okul & Kampüs Lisansı (2028)",
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
        "name": "Kurumsal Zincir & Yayınevi (2028)",
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
        "name": "Akademik İçerik & Soru Bankası Pro (2028)",
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

PARENT_MESSAGE_TOPICS: Dict[str, Dict[str, Any]] = {
    "1. Akademik İlerleme ve Başarı Takibi": {
        "subtopics": [
            "Haftalık Soru ve Net Özeti",
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi",
            "Deneme Sınavı ve 2028 LGS Değerlendirmesi",
        ],
        "templates": {
            "Haftalık Soru ve Net Özeti": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Ofis (2028) platformunda öğrencimiz {student_name} ({grade}) "
                "bu hafta toplam {weekly_solved} soru çözmüştür. "
                "{correct_count} doğru, {wrong_count} yanlış ve {blank_count} boş ile "
                "%{accuracy} başarı oranı elde edilmiştir.\n\n"
                "Güçlü Kazanım: {strong_topic}\n"
                "Pekiştirilmesi Önerilen Alan: {weak_topic}\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in Türkiye Yüzyılı Maarif Modeli beceri temelli kazanım tablosu güncellenmiştir. "
                "'{strong_topic}' alanında güçlü bir ilerleme kaydedilirken, '{weak_topic}' kazanımında "
                "ek pekiştirme testleri tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Deneme Sınavı ve 2028 LGS Değerlendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in son deneme performansında doğruluk oranı %{accuracy} seviyesindedir. "
                "Bir sonraki deneme oturumumuz {next_exam_date} tarihinde '{next_exam_name}' adıyla yapılacaktır.\n\n"
                "Kurşun Kalem Dijital Ofis"
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
                "Kurşun Kalem Dijital Ofis"
            ),
            "Aktif Kullanım ve Düzenlilik": (
                "Sayın {parent_name},\n\n"
                "{student_name} bu hafta platformda düzenli soru ve mini test çözümü gerçekleştirmiştir. "
                "Günlük kısa tekrar rutinlerinin sürdürülmesi önerilmektedir.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Eksik Telafi Uyarıları": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in '{weak_topic}' kazanımındaki telafi görevlerinin tamamlanması "
                "konu bütünlüğü açısından önem taşımaktadır.\n\n"
                "Kurşun Kalem Dijital Ofis"
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
                "Kurşun Kalem Dijital Ofis"
            ),
            "Ödev ve Etkinlik Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} için '{weak_topic}' ve '{strong_topic}' odaklı "
                "haftalık çalışma görevi tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Özel Dönem Bilgilendirmeleri": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Ofis 2028 kamp ve kazanım tekrar programları aktif edilmiştir. "
                "Detaylara panel üzerinden ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
    "4. Abonelik, Paket ve Kurumsal İşlemler": {
        "subtopics": [
            "Ofis Çantası & Paket Durumu",
            "Ödeme ve Fatura Bilgilendirmesi",
            "Destek ve Veli-Öğretmen İletişimi",
        ],
        "templates": {
            "Ofis Çantası & Paket Durumu": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} hesabında {rem_questions} soru üretim kotası ve "
                "{rem_mentor} Dijimentor plan hakkı tanımlıdır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Ödeme ve Fatura Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Abonelik ve ödeme bilgilendirmeleri {parent_email} adresine iletilmektedir. "
                "Her türlü soru için Veli Takip Paneli üzerinden bize ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Destek ve Veli-Öğretmen İletişimi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in gelişim süreciyle ilgili rehberlik ve öğretmen görüşmesi "
                "taleplerinizi Veli Takip Paneli üzerinden iletebilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
}

LOGO_DATA_URI = svg_to_data_uri(EMBEDDED_LOGO_SVG)
BACKPACK_DATA_URI = svg_to_data_uri(SCHOOL_BACKPACK_SVG)
HERO_SHOWCASE_URI = svg_to_data_uri(HERO_SHOWCASE_SVG)
LOGO_PNG_BYTES = get_embedded_logo_png_bytes()

# ==============================================================================
# KUNDUZ İLHAMLI KURUMSAL TASARIM SİSTEMİ + GİRİŞ ANİMASYONLARI
# ==============================================================================
REFINED_OFFICE_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');

:root {
    --kp-navy: #1E2952;
    --kp-navy-2: #2E3F8F;
    --kp-indigo: #5A52E8;
    --kp-purple-bar: #6E47E5;
    --kp-blue-accent: #2563EB;
    --kp-navy-soft: #EEF2FF;
    --kp-gold: #D97706;
    --kp-gold-bright: #FACC15;
    --kp-bg: #FFFFFF;
    --kp-surface: #FFFFFF;
    --kp-surface-alt: #F8FAFC;
    --kp-border: #E2E8F0;
    --kp-text: #1E293B;
    --kp-muted: #64748B;
}

@keyframes kpSoftEnter {
    from { opacity: 0.72; transform: translateY(5px); }
    to { opacity: 1; transform: translateY(0); }
}

@keyframes kpHeroEntrance {
    0% { opacity: 0; transform: translateY(10px) scale(0.99); }
    100% { opacity: 1; transform: translateY(0) scale(1); }
}

@keyframes kpStaggerCard {
    0% { opacity: 0; transform: translateY(8px); }
    100% { opacity: 1; transform: translateY(0); }
}

@keyframes kpTopProgressOnce {
    0% { width: 0%; opacity: 0.9; }
    70% { width: 85%; opacity: 0.9; }
    100% { width: 100%; opacity: 0; }
}

@keyframes kpSkeletonShimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

.kp-top-loader {
    position: fixed;
    top: 0;
    left: 0;
    height: 3px;
    z-index: 9998;
    pointer-events: none;
    background: linear-gradient(90deg, var(--kp-purple-bar) 0%, var(--kp-blue-accent) 60%, var(--kp-gold-bright) 100%);
    animation: kpTopProgressOnce 0.45s ease-out forwards;
}

.kp-skeleton-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 14px;
    padding: 1rem 1.15rem;
    margin-bottom: 0.85rem;
}
.kp-skel-line {
    height: 12px;
    border-radius: 8px;
    margin-bottom: 0.55rem;
    background: linear-gradient(90deg, #F1F5F9 25%, #E2E8F0 50%, #F1F5F9 75%);
    background-size: 200% 100%;
    animation: kpSkeletonShimmer 1.35s ease-in-out infinite;
}

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: var(--kp-text);
}
h1, h2, h3, .kp-kd-hero-title, .kp-page-strip-title, .kp-pkg-title, .kp-mod-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
}
.stApp { background: #FFFFFF; }
[data-testid="stHeader"] {
    background: rgba(255, 255, 255, 0.96);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--kp-border);
}
[data-testid="stDecoration"] { display: none; }
.block-container {
    padding-top: 0.75rem;
    padding-bottom: 4.5rem;
    max-width: 1280px;
    animation: kpSoftEnter 0.24s cubic-bezier(0.22, 1, 0.36, 1) both;
}

section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid var(--kp-border);
}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
    gap: 0.34rem;
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
    width: 42px;
    height: 42px;
    object-fit: contain;
    border-radius: 10px;
    border: 1px solid #CBD5E1;
    padding: 2px;
    background: #FFFFFF;
}
.kp-sidebar-brand-title {
    font-size: 0.98rem;
    font-weight: 800;
    color: #0F172A;
    line-height: 1.2;
}
.kp-sidebar-brand-sub {
    font-size: 0.78rem;
    color: var(--kp-muted);
    margin-top: 2px;
}
/* ---------- Sol menü: çerçevesiz, ikonlu linkler ---------- */
[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    display: flex !important;
    justify-content: flex-start !important;
    align-items: center !important;
    gap: 0.55rem;
    border-radius: 8px !important;
    padding: 0.5rem 0.75rem !important;
    min-height: 2.5rem;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    border: 0 !important;
    background: transparent !important;
    color: #334155 !important;
    box-shadow: none !important;
    transition: background-color 0.14s ease, color 0.14s ease !important;
}
[data-testid="stSidebar"] .stButton > button > div { justify-content: flex-start !important; }
[data-testid="stSidebar"] .stButton > button:hover {
    background: #F1F5F9 !important;
    color: #0F172A !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #EFF6FF !important;
    color: #1D4ED8 !important;
    font-weight: 600 !important;
}
.kp-nav-title {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #64748B;
    padding: 0.9rem 0.75rem 0.2rem 0.75rem;
}
.kp-sidebar-meta {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 0.85rem 0.9rem;
    margin-top: 1rem;
    font-size: 0.8rem;
    color: #334155;
    line-height: 1.45;
}
.kp-meta-head { display: flex; justify-content: space-between; align-items: center; font-weight: 700; color: #0F172A; }
.kp-meta-sub { color: #475569; margin: 0.15rem 0 0.6rem 0; }
.kp-meter-label { display: flex; justify-content: space-between; font-size: 0.76rem; color: #475569; margin: 0.45rem 0 0.2rem 0; }
.kp-meter { height: 6px; background: #E2E8F0; border-radius: 3px; overflow: hidden; }
.kp-meter > span { display: block; height: 100%; background: #1D4ED8; border-radius: 3px; }
.kp-status { display: flex; align-items: center; gap: 0.4rem; font-size: 0.76rem; color: #475569; margin-top: 0.7rem; padding-top: 0.6rem; border-top: 1px solid #E2E8F0; }
.kp-status-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

/* ---------- Üst şerit ---------- */
.kp-topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    flex-wrap: wrap;
    padding: 0.2rem 0 0.9rem 0;
    margin-bottom: 1.5rem;
    border-bottom: 1px solid #E2E8F0;
    font-size: 0.85rem;
    color: #475569;
}
.kp-topbar-brand { font-weight: 700; color: #0F172A; }
.kp-topbar a { color: #1D4ED8; text-decoration: none; }

/* ---------- Hero ---------- */
.kp-eyebrow {
    display: inline-block;
    font-size: 0.82rem;
    font-weight: 600;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.3rem 0.75rem;
    border-radius: 999px;
    margin-bottom: 1rem;
}
.kp-hero-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
    font-size: clamp(2rem, 3.6vw, 3.1rem);
    font-weight: 800;
    line-height: 1.1;
    letter-spacing: -0.02em;
    color: #0F172A;
    margin: 0 0 1rem 0;
    padding: 0;
}
.kp-hero-title em { font-style: normal; color: #1D4ED8; }
.kp-hero-lead { font-size: 1.1rem; line-height: 1.6; color: #475569; margin: 0 0 1.4rem 0; max-width: 34rem; }
.kp-stats { display: flex; gap: 2.2rem; flex-wrap: wrap; margin-top: 1.4rem; }
.kp-stat-num { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.7rem; font-weight: 800; color: #0F172A; line-height: 1.2; }
.kp-stat-lbl { font-size: 0.82rem; color: #64748B; }
.st-key-hero_cta_primary button, .st-key-hero_cta_secondary button {
    min-height: 3rem;
    font-size: 0.95rem !important;
    border-radius: 10px !important;
}

.kp-preview {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 16px;
    padding: 1.3rem;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.08);
}
.kp-preview-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 0.5rem; }
.kp-preview-head small { display: block; font-size: 0.78rem; color: #64748B; }
.kp-preview-head strong { font-size: 1.05rem; color: #0F172A; }
.kp-pill { font-size: 0.75rem; font-weight: 600; color: #1D4ED8; background: #EFF6FF; padding: 0.2rem 0.6rem; border-radius: 999px; white-space: nowrap; }
.kp-preview-stats { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 0.6rem; margin: 1.1rem 0; }
.kp-preview-stat { background: #F8FAFC; border-radius: 10px; padding: 0.7rem; }
.kp-preview-stat b { display: block; font-size: 1.25rem; color: #0F172A; }
.kp-preview-stat span { font-size: 0.75rem; color: #64748B; }
.kp-preview-row { display: flex; justify-content: space-between; gap: 0.75rem; font-size: 0.86rem; color: #334155; padding: 0.6rem 0; border-top: 1px solid #E2E8F0; }
.kp-preview-row span:last-child { color: #0F172A; font-weight: 600; text-align: right; }

/* ---------- Bölüm başlıkları ---------- */
.kp-section-title { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.45rem; font-weight: 700; color: #0F172A; margin: 0; letter-spacing: -0.01em; }
.kp-section-sub { font-size: 0.95rem; color: #475569; margin: 0.3rem 0 0.9rem 0; }
.kp-gap { height: 3rem; }

/* ---------- Kartlı kapsayıcılar (st.container(border=True)) ---------- */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 14px !important; }

/* Sınıf seçici: hap şeklinde butonlar */
[class*="st-key-grade_chip_"] button { border-radius: 999px !important; min-height: 2.75rem; font-weight: 500 !important; font-size: 0.9rem !important; }

/* Modül kartları */
.kp-mod-icon { width: 40px; height: 40px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 0.85rem; margin-bottom: 0.7rem; }
.kp-mod-title { font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 700; font-size: 1rem; color: #0F172A; }
.kp-mod-desc { font-size: 0.86rem; line-height: 1.5; color: #475569; margin-top: 0.3rem; min-height: 3.9rem; }
[class*="st-key-mod_open_"] button, .st-key-prio_all button {
    border: 0 !important;
    background: transparent !important;
    color: #1D4ED8 !important;
    padding: 0 !important;
    min-height: 0 !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    justify-content: flex-start !important;
}
[class*="st-key-mod_open_"] button:hover, .st-key-prio_all button:hover { color: #1E3A8A !important; text-decoration: underline; }

/* Öncelik satırları */
.kp-prio-num { width: 32px; height: 32px; border-radius: 50%; background: #F1F5F9; font-weight: 700; font-size: 0.9rem; display: flex; align-items: center; justify-content: center; color: #0F172A; }
.kp-prio-title { font-weight: 600; font-size: 0.95rem; color: #0F172A; line-height: 1.4; }
.kp-prio-sub { font-size: 0.8rem; color: #64748B; margin-top: 0.15rem; }
.kp-prio-divider { border-top: 1px solid #E2E8F0; margin: 0.35rem 0; }
.kp-tag-warn { font-size: 0.75rem; font-weight: 600; color: #92400E; background: #FEF3C7; padding: 0.2rem 0.6rem; border-radius: 999px; }
.kp-page-strip {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 1rem;
    padding: 0.05rem 0 0.7rem 0;
    margin-bottom: 1rem;
    border-bottom: 1px solid var(--kp-border);
    animation: kpHeroEntrance 0.32s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-page-strip-title {
    font-size: 1.28rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0;
}
.kp-page-strip-desc {
    font-size: 0.83rem;
    color: var(--kp-muted);
    margin: 0.2rem 0 0 0;
}

.kp-kademe-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 16px;
    padding: 1rem 1.1rem;
    height: 146px;
    display: flex;
    flex-direction: column;
    justify-content: flex-start;
    margin-bottom: 0.55rem;
    box-shadow: 0 3px 12px rgba(37, 99, 235, 0.04);
    transition: transform 0.16s ease, box-shadow 0.16s ease;
    animation: kpStaggerCard 0.44s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-kademe-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 18px rgba(37, 99, 235, 0.09);
}
.kp-kademe-tag {
    display: inline-block;
    width: fit-content;
    font-size: 0.7rem;
    font-weight: 800;
    color: #FFFFFF;
    padding: 0.22rem 0.65rem;
    border-radius: 999px;
}
.kp-kademe-title {
    font-size: 1rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.55rem 0 0.25rem 0;
}
.kp-kademe-desc {
    font-size: 0.8rem;
    color: var(--kp-muted);
    line-height: 1.42;
    margin: 0;
}

.kp-priority-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-left: 4px solid #2563EB;
    border-radius: 14px;
    padding: 0.95rem 1.05rem;
    height: 132px;
    margin-bottom: 0.55rem;
    animation: kpStaggerCard 0.46s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-priority-badge {
    font-size: 0.7rem;
    font-weight: 800;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.18rem 0.58rem;
    border-radius: 999px;
    display: inline-block;
}
.kp-priority-title {
    font-size: 0.94rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.45rem 0 0.2rem 0;
}
.kp-priority-sub {
    font-size: 0.79rem;
    color: var(--kp-muted);
    margin: 0;
}

.kp-value-strip {
    background: linear-gradient(90deg, #1E2952 0%, #2563EB 65%, #5A52E8 100%);
    color: #FFFFFF;
    border-radius: 14px;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1.05rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
    flex-wrap: wrap;
}
.kp-value-strip-item {
    display: flex;
    align-items: center;
    gap: 0.48rem;
    font-size: 0.82rem;
    font-weight: 600;
}
.kp-value-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #FACC15;
    display: inline-block;
}

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 14px;
    padding: 0.92rem 1.05rem;
    box-shadow: 0 3px 10px rgba(15, 23, 42, 0.04);
    animation: kpStaggerCard 0.42s cubic-bezier(0.22, 1, 0.36, 1) both;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    font-size: 1.52rem;
    font-weight: 800;
    color: var(--kp-navy);
}

.kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-kpi-card, .kp-metric-box, .kp-studio-banner {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 16px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
    animation: kpStaggerCard 0.42s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-studio-banner {
    padding: 1.15rem 1.25rem;
}
.kp-card { margin-bottom: 0.65rem; }

.kp-studio-banner {
    margin-bottom: 1.05rem;
    border-left: 4px solid #2563EB;
}
.kp-studio-title {
    font-size: 1.14rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0;
}
.kp-studio-sub {
    font-size: 0.84rem;
    color: var(--kp-muted);
    margin: 0.22rem 0 0 0;
}
.kp-kpi-card { padding: 0.92rem 1.02rem; }
.kp-kpi-label {
    font-size: 0.7rem;
    font-weight: 700;
    color: #64748B;
    text-transform: uppercase;
}
.kp-kpi-value {
    font-size: 1.25rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin-top: 0.2rem;
}
.kp-kpi-sub {
    font-size: 0.76rem;
    color: var(--kp-muted);
    margin-top: 0.15rem;
}

.kp-metric-box {
    padding: 1.05rem 0.95rem;
    text-align: center;
    border-bottom: 4px solid #2563EB;
}
.kp-metric-num {
    font-size: 1.62rem;
    font-weight: 800;
    color: var(--kp-navy);
    line-height: 1.15;
    font-variant-numeric: tabular-nums;
}
.kp-metric-lbl {
    font-size: 0.71rem;
    color: #64748B;
    margin-top: 0.3rem;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 700;
}

.kp-bag-shell {
    margin-bottom: 1.2rem;
    border-top: 4px solid #2563EB;
}
.kp-bag-visual {
    background: linear-gradient(160deg, #1E2952 0%, #2563EB 100%);
    border-radius: 14px;
    padding: 1rem 0.85rem;
    display: flex;
    flex-direction: column;
    align-items: center;
    color: #FFFFFF;
}
.kp-bag-visual img {
    width: 102px;
    height: 118px;
    object-fit: contain;
}
.kp-bag-count-pill {
    margin-top: 0.5rem;
    background: #FACC15;
    color: #0F172A;
    font-size: 0.74rem;
    font-weight: 800;
    padding: 0.2rem 0.7rem;
    border-radius: 999px;
}
.kp-bag-item-chip {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-left: 3.5px solid #2563EB;
    border-radius: 10px;
    padding: 0.62rem 0.85rem;
    margin-bottom: 0.45rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.kp-pkg-card { margin-bottom: 0.65rem; }
.kp-pkg-featured {
    border: 2px solid #2563EB !important;
    background: linear-gradient(180deg, #EFF6FF 0%, #FFFFFF 35%) !important;
}
.kp-pkg-top {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 0.5rem;
}
.kp-pkg-seg {
    font-size: 0.71rem;
    font-weight: 700;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.2rem 0.58rem;
    border-radius: 6px;
}
.kp-pkg-badge {
    font-size: 0.71rem;
    font-weight: 800;
    color: #92400E;
    background: #FEF3C7;
    border: 1px solid #F59E0B;
    padding: 0.2rem 0.58rem;
    border-radius: 6px;
}
.kp-pkg-title {
    font-size: 1.04rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.55rem 0 0.2rem 0;
}
.kp-pkg-price {
    font-size: 1.38rem;
    font-weight: 800;
    color: var(--kp-text);
    margin-bottom: 0.4rem;
}
.kp-pkg-price span {
    font-size: 0.8rem;
    font-weight: 500;
    color: #64748B;
}
.kp-pkg-desc {
    font-size: 0.82rem;
    color: var(--kp-muted);
    line-height: 1.42;
    margin-bottom: 0.7rem;
    min-height: 3.1rem;
}
.kp-pkg-feats {
    list-style: none;
    padding: 0.65rem 0 0 0;
    margin: 0;
    border-top: 1px solid var(--kp-border);
}
.kp-pkg-feats li {
    font-size: 0.79rem;
    color: #334155;
    padding: 0.16rem 0;
}

.kp-dist-bar {
    display: flex;
    height: 9px;
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

.stButton > button,
div[data-testid="stFormSubmitButton"] > button {
    border-radius: 8px;
    font-weight: 600;
    font-size: 0.9rem;
    border: 1px solid #CBD5E1;
    background-color: #FFFFFF;
    color: #1E293B;
}
.stButton > button:hover,
div[data-testid="stFormSubmitButton"] > button:hover {
    border-color: #2563EB;
    background-color: #EFF6FF;
    color: #1D4ED8;
}
.stButton > button[kind="primary"],
div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
    background: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFFFFF !important;
}
.stButton > button[kind="primary"]:hover,
div[data-testid="stFormSubmitButton"] > button[kind="primary"]:hover {
    background: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
}

@media (prefers-reduced-motion: reduce) {
    .block-container,
    .kp-top-loader,
    .kp-top-compact-bar,
    .kp-kd-kicker,
    .kp-kd-hero-title,
    .kp-kd-hero-sub,
    .kp-kd-social-row,
    .kp-feature-strip-card,
    .kp-showcase-frame,
    .kp-page-strip,
    .kp-kademe-card,
    .kp-priority-card,
    div[data-testid="stMetric"],
    .kp-card,
    .kp-bag-shell,
    .kp-pkg-card,
    .kp-pay-box,
    .kp-kpi-card,
    .kp-metric-box,
    .kp-studio-banner {
        animation: none !important;
        opacity: 1 !important;
        transform: none !important;
    }
}
"""

# ==============================================================================
# EĞLENCELİ TEMA KATMANI: defter zemini, fosforlu vurgu, "pop" gölgeli kartlar
# ==============================================================================
FUN_THEME_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@600;700;800&display=swap');

:root {
    --fun-ink: #0F172A;
    --fun-yellow: #FDE047;
    --fun-yellow-deep: #FACC15;
    --fun-blue: #1D4ED8;
    --fun-paper: #FFFDF6;
}

/* Kareli defter zemini */
.stApp {
    background-color: var(--fun-paper);
    background-image:
        linear-gradient(rgba(29, 78, 216, 0.06) 1px, transparent 1px),
        linear-gradient(90deg, rgba(29, 78, 216, 0.06) 1px, transparent 1px);
    background-size: 26px 26px;
}
[data-testid="stHeader"] { background: rgba(255, 253, 246, 0.9); }

/* Başlıklarda yuvarlak, samimi yazı tipi */
.kp-hero-title, .kp-section-title, .kp-stat-num, .kp-mod-title, .kp-metric-num, .kp-page-strip-title {
    font-family: 'Baloo 2', 'Plus Jakarta Sans', sans-serif !important;
    letter-spacing: 0 !important;
}
.kp-hero-title { font-size: clamp(2.3rem, 4vw, 3.4rem); line-height: 1.05; }
.kp-section-title { font-size: 1.7rem; }

/* Fosforlu kalem vurgusu */
.kp-hero-title em {
    color: var(--fun-ink);
    background: linear-gradient(100deg, transparent 2%, var(--fun-yellow) 4%, var(--fun-yellow) 96%, transparent 98%);
    background-size: 100% 42%;
    background-repeat: no-repeat;
    background-position: 0 88%;
    padding: 0 0.15em;
}

/* Çıkartma gibi etiket */
.kp-eyebrow {
    background: var(--fun-yellow);
    color: var(--fun-ink);
    border: 2px solid var(--fun-ink);
    box-shadow: 2px 2px 0 var(--fun-ink);
    transform: rotate(-2deg);
    font-weight: 700;
}

/* Rakamlar: her biri kendi rengiyle */
.kp-stats > div:nth-child(1) .kp-stat-num { color: #1D4ED8; }
.kp-stats > div:nth-child(2) .kp-stat-num { color: #DB2777; }
.kp-stats > div:nth-child(3) .kp-stat-num { color: #059669; }
.kp-stat-num { font-size: 2.1rem; }

/* Butonlar: kalın çerçeve + bastıkça gömülen gölge */
.stButton > button, div[data-testid="stFormSubmitButton"] > button {
    border: 2px solid var(--fun-ink) !important;
    border-radius: 12px !important;
    box-shadow: 3px 3px 0 var(--fun-ink);
    font-weight: 700 !important;
    transition: transform 0.08s ease, box-shadow 0.08s ease, background-color 0.12s ease !important;
}
.stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover {
    transform: translate(-1px, -1px);
    box-shadow: 4px 4px 0 var(--fun-ink);
    background-color: #FEF9C3 !important;
    color: var(--fun-ink) !important;
    border-color: var(--fun-ink) !important;
}
.stButton > button:active { transform: translate(2px, 2px); box-shadow: 1px 1px 0 var(--fun-ink); }
.stButton > button[kind="primary"]:hover { background: #1E40AF !important; color: #FFFFFF !important; }

/* Sol menü ve link görünümlü butonlar bu stilden muaf */
[data-testid="stSidebar"] .stButton > button,
[class*="st-key-mod_open_"] button, .st-key-prio_all button {
    box-shadow: none !important;
    transform: none !important;
}
[data-testid="stSidebar"] .stButton > button { border: 0 !important; }
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: var(--fun-yellow) !important;
    color: var(--fun-ink) !important;
}
[class*="st-key-mod_open_"] button, .st-key-prio_all button { border: 0 !important; }
[class*="st-key-mod_open_"] button:hover, .st-key-prio_all button:hover { background: transparent !important; }

/* Kartlar: kalın çerçeve + renkli pop gölge */
[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"]) {
    border: 2px solid var(--fun-ink) !important;
    border-radius: 18px !important;
    background: #FFFFFF;
    box-shadow: 5px 5px 0 var(--fun-ink);
}
.kp-preview {
    position: relative;
    border: 2px solid var(--fun-ink);
    border-radius: 20px;
    box-shadow: 7px 7px 0 var(--fun-yellow-deep), 7px 7px 0 2px var(--fun-ink);
    transform: rotate(1.2deg);
    transition: transform 0.2s ease;
}
.kp-preview:hover { transform: rotate(0deg); }
.kp-preview-stat { border: 2px solid #E2E8F0; border-radius: 12px; }
.kp-preview-stat:nth-child(1) { background: #EFF6FF; }
.kp-preview-stat:nth-child(2) { background: #FDF2F8; }
.kp-preview-stat:nth-child(3) { background: #ECFDF5; }
.kp-meter { height: 10px; border-radius: 999px; border: 1.5px solid var(--fun-ink); background: #FFFFFF; }
.kp-meter > span { border-radius: 999px; }
.kp-sticker {
    position: absolute;
    top: -18px;
    right: -14px;
    background: #F472B6;
    color: #FFFFFF;
    font-weight: 800;
    font-size: 0.85rem;
    padding: 0.45rem 0.8rem;
    border: 2px solid var(--fun-ink);
    border-radius: 999px;
    box-shadow: 2px 2px 0 var(--fun-ink);
    transform: rotate(8deg);
    white-space: nowrap;
}

/* Modül ikonları: hover'da zıplama */
.kp-mod-icon {
    width: 52px; height: 52px; border-radius: 14px;
    border: 2px solid var(--fun-ink);
    box-shadow: 2px 2px 0 var(--fun-ink);
    transition: transform 0.2s ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover .kp-mod-icon { transform: rotate(-8deg) scale(1.08); }

/* Sınıf hapları */
[class*="st-key-grade_chip_"] button { background: #FFFFFF; }

/* Öncelik numaraları */
.kp-prio-num { background: var(--fun-yellow); border: 2px solid var(--fun-ink); }

/* Geri sayım kutuları */
.kp-metric-box {
    border: 2px solid var(--fun-ink) !important;
    border-radius: 16px !important;
    box-shadow: 4px 4px 0 var(--fun-ink) !important;
}
.kp-metric-num { font-size: 2.2rem !important; color: var(--fun-blue) !important; }

/* Sol menü kota kutusu */
.kp-sidebar-meta { border: 2px solid var(--fun-ink); box-shadow: 3px 3px 0 var(--fun-ink); background: #FFFFFF; }

@media (prefers-reduced-motion: reduce) {
    .kp-preview, .kp-eyebrow, .kp-sticker, .kp-mod-icon { transform: none !important; transition: none !important; }
}
"""

# Her modül kartına kendi pastel zemini
FUN_THEME_CSS += "".join(
    f'[data-testid="stVerticalBlockBorderWrapper"]:has(.kp-mod-{i}) {{ background: {bg} !important; }}\n'
    for i, (_c, _t, _d, _p, bg, _fg) in enumerate(MODULE_CARDS)
)

css_to_inject = REFINED_OFFICE_CSS + ("\n" + FUN_THEME_CSS if FUN_THEME else "")
if EXTERNAL_CSS_PATH.exists():
    try:
        ext_css = EXTERNAL_CSS_PATH.read_text(encoding="utf-8").strip()
        if ext_css:
            css_to_inject = f"{REFINED_OFFICE_CSS}\n\n{ext_css}"
    except Exception:
        pass

st.markdown(f"<style>{css_to_inject}</style>", unsafe_allow_html=True)


def trigger_loader_bar() -> None:
    st.session_state["loader_nonce"] = uuid.uuid4().hex[:8]


def render_top_loader_bar() -> None:
    nonce = st.session_state.get("loader_nonce", "init")
    st.markdown(
        f'<div class="kp-top-loader" data-load-id="{html.escape(str(nonce))}"></div>',
        unsafe_allow_html=True,
    )


def render_loading_skeleton(placeholder: Any, lines: int = 3) -> None:
    bars = "".join(['<div class="kp-skel-line" style="width:94%;"></div>' for _ in range(max(2, lines))])
    placeholder.markdown(f'<div class="kp-skeleton-card">{bars}</div>', unsafe_allow_html=True)


# ==============================================================================
# YARDIMCI FONKSİYONLAR, KOTA YÖNETİMİ VE KALICI DURUM
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
        "weak_topic": "Sözel Mantık ve Muhakeme",
        "next_exam_date": (datetime.now() + timedelta(days=5)).strftime("%d.%m.%Y"),
        "next_exam_name": "Türkiye Geneli 2028 LGS Deneme Sınavı",
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
        {"time": "09:15:00", "module": "Dijital Ofis 2028 • Fabrika", "summary": "8. Sınıf 2028 LGS Muhakeme seti hazır"},
        {"time": "10:42:18", "module": "Dijital Ofis 2028 • Dijimentor", "summary": "Haftalık çeldirici & tempo analizi aktif"},
    ]


def load_persisted_state() -> Dict[str, Any]:
    if not STATE_FILE_PATH.exists():
        return {}
    try:
        raw = STATE_FILE_PATH.read_text(encoding="utf-8")
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_persisted_state() -> None:
    payload = {
        "backpack_cart": st.session_state.get("backpack_cart", {}),
        "selected_tier_id": st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"]),
        "used_question_quota": int(st.session_state.get("used_question_quota", 0)),
        "used_coach_quota": int(st.session_state.get("used_coach_quota", 0)),
        "remaining_question_quota": int(st.session_state.get("remaining_question_quota", 0)),
        "remaining_coach_quota": int(st.session_state.get("remaining_coach_quota", 0)),
        "parent_student_stats": st.session_state.get("parent_student_stats", _default_parent_stats()),
        "parent_announcements": st.session_state.get("parent_announcements", _default_parent_announcements()),
        "activity_log": st.session_state.get("activity_log", _default_activity_log()),
        "question_bank_2028": st.session_state.get("question_bank_2028", []),
        "error_notebook_2028": st.session_state.get("error_notebook_2028", []),
        "b2b_org_cache": st.session_state.get("b2b_org_cache"),
        "last_b2b_summary": st.session_state.get("last_b2b_summary", ""),
    }
    try:
        STATE_FILE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def attach_logo_to_doc(doc: Document, width_in: float = 1.15) -> None:
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
    api_key, secret_key, base_url = resolve_iyzico_config()
    if not api_key or not secret_key:
        return {
            "status": "failure",
            "errorMessage": (
                "IYZICO_API_KEY ve IYZICO_SECRET_KEY tanımlı değil. "
                "Streamlit Secrets veya ortam değişkenlerine anahtarlarınızı ekleyin."
            ),
        }
    options = {"api_key": api_key, "secret_key": secret_key, "base_url": base_url}
    price_str = f"{max(1.0, float(amount_try)):.2f}"
    request = {
        "locale": "tr",
        "conversationId": f"KP28-{int(time.time())}",
        "price": price_str,
        "paidPrice": price_str,
        "currency": "TRY",
        "basketId": f"B28-{int(time.time())}",
        "paymentGroup": "PRODUCT",
        "callbackUrl": "https://kursunkalemdijitalofis.com.tr",
        "enabledInstallments": [1, 2, 3, 6, 9, 12],
        "buyer": {
            "id": "BY_KP28_01",
            "name": buyer_name.strip() or "Sinan Gürbüz",
            "surname": buyer_surname.strip() or "Kurşun",
            "gsmNumber": buyer_phone.strip() or "+905350000000",
            "email": buyer_email.strip() or "iletisim@kursunkalemdijitalofis.com.tr",
            "identityNumber": "74300864791",
            "registrationAddress": "İstanbul, Türkiye",
            "ip": "85.34.78.112",
            "city": "Istanbul",
            "country": "Turkey",
        },
        "shippingAddress": {
            "contactName": f"{buyer_name.strip() or 'Sinan Gürbüz'} {buyer_surname.strip() or 'Kurşun'}",
            "city": "Istanbul",
            "country": "Turkey",
            "address": "İstanbul, Türkiye",
        },
        "billingAddress": {
            "contactName": f"{buyer_name.strip() or 'Sinan Gürbüz'} {buyer_surname.strip() or 'Kurşun'}",
            "city": "Istanbul",
            "country": "Turkey",
            "address": "İstanbul, Türkiye",
        },
        "basketItems": [
            {
                "id": "BI_KP28_101",
                "name": basket_title[:120],
                "category1": "Dijital Eğitim 2028",
                "itemType": "VIRTUAL",
                "price": price_str,
            }
        ],
    }
    checkout_form_initialize = iyzipay.CheckoutFormInitialize().create(request, options)
    raw_res = checkout_form_initialize.read().decode("utf-8")
    return json.loads(raw_res)


def get_gemini_client() -> genai.Client:
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
    client = get_gemini_client()
    candidate_models = [DEFAULT_MODEL, FALLBACK_MODEL]
    if model and model not in (
        "gemini-2.5-pro",
        "gemini-3.1-pro-preview",
        "gemini-2.5-flash-lite",
    ):
        candidate_models.insert(0, model)

    deduped_models: List[str] = []
    for m in candidate_models:
        if m and m not in deduped_models:
            deduped_models.append(m)

    errors_by_model: List[str] = []
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
                err_str = str(exc)
                errors_by_model.append(f"[{candidate_model}] {err_str}")
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                if any(tok in err_str for tok in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "overloaded")):
                    time.sleep(1.2 * (attempt + 1))
                    continue
                break

    raise RuntimeError(f"Model çağrısı başarısız oldu: {' | '.join(errors_by_model[-2:])}")


def clear_page_ephemeral_state(target_page: str) -> None:
    if target_page != "Öğrenci Portalı & Çözüm":
        st.session_state.pop("last_exam_report", None)
    if target_page != "Öğrenci Röntgeni":
        st.session_state.pop("xray_last_scan", None)


def navigate_to(page_name: str) -> None:
    if page_name in PAGES:
        clear_page_ephemeral_state(page_name)
        st.session_state.active_page = page_name
        trigger_loader_bar()
        st.rerun()


def record_activity(module_name: str, summary: str) -> None:
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
    return next((t for t in SUBSCRIPTION_TIERS if t["id"] == tier_id), SUBSCRIPTION_TIERS[1])


def ensure_bag_initialized() -> None:
    if "backpack_cart" not in st.session_state or not isinstance(st.session_state["backpack_cart"], dict):
        default_id = st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
        st.session_state["backpack_cart"] = {default_id: 1}


def get_bag_items() -> Dict[str, int]:
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
        title = entries[0]["tier"]["name"] if len(entries) == 1 else f"Ofis Çantası ({total_count} Paket)"
        segment = entries[0]["tier"]["segment"] if len(entries) == 1 else "Karma / Çoklu Paket"
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


def sync_plan_quotas(summary: Dict[str, Any], reset_usage: bool = False) -> None:
    if reset_usage:
        st.session_state["used_question_quota"] = 0
        st.session_state["used_coach_quota"] = 0

    used_q = max(0, int(st.session_state.get("used_question_quota", 0)))
    used_c = max(0, int(st.session_state.get("used_coach_quota", 0)))
    total_q = max(0, int(summary["monthly_quota"]))
    total_c = max(0, int(summary["coach_quota"]))

    if used_q > total_q:
        used_q = total_q
        st.session_state["used_question_quota"] = used_q
    if used_c > total_c:
        used_c = total_c
        st.session_state["used_coach_quota"] = used_c

    st.session_state["remaining_question_quota"] = max(0, total_q - used_q)
    st.session_state["remaining_coach_quota"] = max(0, total_c - used_c)
    save_persisted_state()


def cb_add_tier_to_bag(tier_id: str, qty: int = 1) -> None:
    bag = dict(get_bag_items())
    bag[tier_id] = bag.get(tier_id, 0) + max(1, int(qty))
    st.session_state["backpack_cart"] = bag
    st.session_state["selected_tier_id"] = tier_id
    sync_plan_quotas(get_bag_summary(), reset_usage=False)
    tier = get_tier_by_id(tier_id)
    record_activity("Ofis Çantası", f"{tier['name']} çantaya eklendi")


def cb_update_bag_qty(tier_id: str, delta: int) -> None:
    bag = dict(get_bag_items())
    if tier_id not in bag:
        return
    new_qty = bag[tier_id] + delta
    tier = get_tier_by_id(tier_id)
    if new_qty <= 0:
        del bag[tier_id]
        record_activity("Ofis Çantası", f"{tier['name']} çantadan çıkarıldı")
    else:
        bag[tier_id] = new_qty
        record_activity("Ofis Çantası", f"{tier['name']} adedi {new_qty} yapıldı")
    st.session_state["backpack_cart"] = bag
    if bag:
        st.session_state["selected_tier_id"] = next(iter(bag.keys()))
    sync_plan_quotas(get_bag_summary(), reset_usage=False)


def cb_clear_bag() -> None:
    st.session_state["backpack_cart"] = {}
    sync_plan_quotas(get_bag_summary(), reset_usage=False)
    record_activity("Ofis Çantası", "Çantadaki tüm paketler silindi")


def consume_quota(kind: str, amount: int = 1) -> Tuple[bool, str]:
    summary = get_bag_summary()
    sync_plan_quotas(summary, reset_usage=False)

    if kind == "question":
        rem = int(st.session_state.get("remaining_question_quota", 0))
        if rem < amount:
            return False, (
                f"Ofis çantanızda ({summary['name']}) yeterli soru üretim kotası kalmadı "
                f"(Kalan: {rem}, İstenen: {amount}). Paketler vitrininden paket ekleyebilirsiniz."
            )
        st.session_state["used_question_quota"] = int(st.session_state.get("used_question_quota", 0)) + amount
        sync_plan_quotas(summary, reset_usage=False)
        return True, ""

    rem_c = int(st.session_state.get("remaining_coach_quota", 0))
    if rem_c < amount:
        return False, (
            f"Ofis çantanızda ({summary['name']}) Dijimentor plan kotası doldu "
            f"(Kalan: {rem_c}). Paketler vitrininden paket ekleyebilirsiniz."
        )
    st.session_state["used_coach_quota"] = int(st.session_state.get("used_coach_quota", 0)) + amount
    sync_plan_quotas(summary, reset_usage=False)
    return True, ""


def refund_quota(kind: str, amount: int = 1) -> None:
    if kind == "question":
        st.session_state["used_question_quota"] = max(
            0, int(st.session_state.get("used_question_quota", 0)) - amount
        )
    else:
        st.session_state["used_coach_quota"] = max(
            0, int(st.session_state.get("used_coach_quota", 0)) - amount
        )
    sync_plan_quotas(get_bag_summary(), reset_usage=False)


def add_to_question_bank(
    grade: str,
    branch: str,
    outcome: str,
    difficulty: str,
    content_text: str,
) -> None:
    bank: List[Dict[str, str]] = st.session_state.setdefault("question_bank_2028", [])
    bank.insert(
        0,
        {
            "id": uuid.uuid4().hex[:6].upper(),
            "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "grade": grade,
            "branch": branch,
            "outcome": outcome,
            "difficulty": difficulty,
            "content": content_text,
        },
    )
    st.session_state["question_bank_2028"] = bank[:40]
    save_persisted_state()


def add_to_error_notebook(
    question_num: int,
    question_text: str,
    user_ans: str,
    correct_ans: str,
    elapsed_sec: float,
) -> None:
    nb: List[Dict[str, Any]] = st.session_state.setdefault("error_notebook_2028", [])
    now = datetime.now()
    diag = "Süre / Dikkat Baskısı" if elapsed_sec > 90 else "Çeldirici / Kavram İncelemesi"
    nb.insert(
        0,
        {
            "id": uuid.uuid4().hex[:6].upper(),
            "num": question_num,
            "date": now.strftime("%d.%m.%Y %H:%M"),
            "review_1d": (now + timedelta(days=1)).strftime("%d.%m.%Y"),
            "review_3d": (now + timedelta(days=3)).strftime("%d.%m.%Y"),
            "review_7d": (now + timedelta(days=7)).strftime("%d.%m.%Y"),
            "user_ans": user_ans,
            "correct_ans": correct_ans,
            "elapsed_sec": round(elapsed_sec, 1),
            "diagnosis": diag,
            "question_text": question_text,
        },
    )
    st.session_state["error_notebook_2028"] = nb[:50]
    save_persisted_state()


# ==============================================================================
# GÜÇLÜ SORU AYRIŞTIRICI & DOCX ÜRETİCİLERİ
# ==============================================================================
def extract_answer_key(raw_text: str) -> Dict[int, str]:
    if not raw_text.strip():
        return {}

    parts = re.split(
        r"(?:^|\n)#*\s*(?:Cevap\s*Anahtar\w*|Yanıt\s*Anahtar\w*)[^\n]*",
        raw_text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )
    key_map: Dict[int, str] = {}

    patterns = [
        r"(?:Soru\s*)?(\d+)\s*[\.\-\:\)\]]\s*\**\s*([A-D])\b",
        r"(\d+)\s*\.\s*Soru\s*[\:\-\)\]]\s*\**\s*([A-D])\b",
        r"\b(\d+)\s*[-–—]\s*\**\s*([A-D])\b",
    ]

    if len(parts) > 1:
        for pat in patterns:
            for m in re.finditer(pat, parts[1], flags=re.IGNORECASE):
                q_num = int(m.group(1))
                ans_letter = m.group(2).upper()
                if q_num not in key_map:
                    key_map[q_num] = ans_letter

    q_blocks = re.split(
        r"(?=(?:^|\n)\s*(?:#+\s*)?(?:\*{0,2})\s*(?:Soru\s*\d+|\d+\s*[\.\)]))",
        parts[0],
        flags=re.IGNORECASE,
    )
    for blk in q_blocks:
        num_m = re.match(r"^\s*(?:#+\s*)?(?:\*{0,2})\s*(?:Soru\s*)?(\d+)\b", blk, flags=re.IGNORECASE)
        if not num_m:
            continue
        q_num = int(num_m.group(1))
        if q_num in key_map:
            continue
        ans_m = re.search(
            r"(?:Doğru\s*(?:Cevap|Yanıt)|Cevap|Yanıt)\s*[\:\-–]\s*\**\s*([A-D])\b",
            blk,
            flags=re.IGNORECASE,
        )
        if ans_m:
            key_map[q_num] = ans_m.group(1).upper()

    return key_map


def parse_quiz_questions(raw_text: str) -> List[Dict[str, Any]]:
    if not raw_text.strip():
        return []

    answer_map = extract_answer_key(raw_text)
    main_body = re.split(
        r"(?:^|\n)\s*#*\s*(?:Cevap\s*Anahtar\w*|Yanıt\s*Anahtar\w*)",
        raw_text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip()

    if not main_body:
        return []

    header_split_pattern = (
        r"(?=(?:^|\n)\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s+\d+|\d+\s*(?:\.\s*Soru|\)|\.\s+(?=[A-ZÇĞİÖŞÜ\"“'\*]))))"
    )
    raw_blocks = re.split(header_split_pattern, main_body, flags=re.IGNORECASE | re.MULTILINE)

    parsed: List[Dict[str, Any]] = []
    seen_nums: set[int] = set()

    for idx_b, block in enumerate(raw_blocks, start=1):
        text = block.strip()
        if len(text) < 15:
            continue

        num_match = re.match(
            r"^\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s+)?(\d+)\b",
            text,
            flags=re.IGNORECASE,
        )
        q_num = int(num_match.group(1)) if num_match else idx_b
        if q_num in seen_nums:
            q_num = max(seen_nums) + 1

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        stem_lines: List[str] = []
        options: Dict[str, str] = {}
        current_opt: Optional[str] = None

        opt_pattern = re.compile(
            r"^\s*(?:[-*•]\s*)?(?:\*\*)?\s*([A-D])\s*[\)\.\:]\s*(?:\*\*)?\s*(.+)$",
            flags=re.IGNORECASE,
        )

        for ln in lines:
            if re.match(r"^\s*\[?(?:Redaksiyon|Çözüm|Cevap\s*Anahtar|Doğru\s*Cevap)", ln, flags=re.IGNORECASE):
                current_opt = None
                continue
            if re.match(r"^\s*[-_=]{3,}\s*$", ln):
                continue

            opt_m = opt_pattern.match(ln)
            if opt_m:
                letter = opt_m.group(1).upper()
                opt_val = re.sub(r"\*\*$", "", opt_m.group(2).strip()).strip()
                options[letter] = opt_val
                current_opt = letter
            elif current_opt:
                extra = re.sub(r"\*\*$", "", ln).strip()
                if extra:
                    options[current_opt] = f"{options[current_opt]} {extra}".strip()
            else:
                stem_lines.append(ln)

        if len(options) >= 2:
            ordered_opts = {k: options[k] for k in ["A", "B", "C", "D"] if k in options}
            for k, v in options.items():
                if k not in ordered_opts:
                    ordered_opts[k] = v

            seen_nums.add(q_num)
            parsed.append(
                {
                    "num": q_num,
                    "stem": "\n".join(stem_lines).strip(),
                    "options": ordered_opts,
                    "correct": answer_map.get(q_num),
                    "full_text": text,
                }
            )

    return parsed


def format_countdown(target: datetime) -> Tuple[int, int, int, int]:
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
    title, desc = PAGE_DESCRIPTIONS.get(page_name, (page_name, "Kurşun Kalem Dijital Ofis 2028 çalışma alanı."))
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
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"{institution.upper()} (2028) — DEĞERLENDİRME TESTİ")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.color.rgb = RGBColor(30, 41, 82)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(f"Branş: {branch}  |  Sınıf: {grade}  |  Öğrenme Alanı: {learning_area}")
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run(
        f"Hedef Kazanım: {outcome}\nBağlam / Tema: {context_theme}\nTarih: {datetime.now().strftime('%d.%m.%Y')}"
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
            run.font.color.rgb = RGBColor(30, 41, 82)

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run(f"\n{BRAND_NAME} • 2028")
    footer_run.italic = True
    footer_run.font.size = Pt(8.5)
    footer_run.font.color.rgb = RGBColor(100, 116, 139)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_coach_docx(title_meta: str, plan_text: str) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} (2028) — DİJİMENTOR ÇALIŞMA PLANI")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(30, 41, 82)

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
            run.font.color.rgb = RGBColor(30, 41, 82)

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
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} (2028) — OFİS ÇANTASI & ABONELİK ÖZETİ")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(30, 41, 82)

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
            run.font.color.rgb = RGBColor(30, 41, 82)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_school_backpack_panel() -> Dict[str, Any]:
    summary = get_bag_summary()
    entries = summary["entries"]

    st.markdown('<div class="kp-bag-shell">', unsafe_allow_html=True)
    h_col, clr_col = st.columns([4, 1])
    with h_col:
        st.markdown("#### Kurşun Kalem Ofis Çantası (2028)")
        st.caption("Seçili paketler, adet kontrolleri, kota çubukları ve aylık toplam özet.")
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
            chips_html = '<div style="color:#E2E8F0;font-size:0.78rem;margin-top:0.35rem;text-align:center;">Çanta boş</div>'

        st.markdown(
            f"""
            <div class="kp-bag-visual">
                <img src="{BACKPACK_DATA_URI}" alt="Ofis Çantası"/>
                <div class="kp-bag-count-pill">{summary['total_count']} Paket</div>
                <div style="width:100%;margin-top:0.3rem;">{chips_html}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with list_col:
        if not entries:
            st.info("Ofis çantanız şu an boş. Aşağıdaki paketlerden **Çantaya Ekle** butonunu kullanabilirsiniz.")
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
                            <div style="font-weight:700;font-size:0.92rem;color:#2563EB;white-space:nowrap;">
                                {entry['line_price']:,} ₺/ay
                            </div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )
                with c_minus:
                    st.button("−", key=f"bag_minus_{t['id']}", use_container_width=True, on_click=cb_update_bag_qty, args=(t["id"], -1))
                with c_plus:
                    st.button("+", key=f"bag_plus_{t['id']}", use_container_width=True, on_click=cb_update_bag_qty, args=(t["id"], +1))
                with c_rem:
                    st.button("Çıkar", key=f"bag_del_{t['id']}", use_container_width=True, on_click=cb_update_bag_qty, args=(t["id"], -qty))

            st.markdown("<div style='margin-top:0.45rem;'></div>", unsafe_allow_html=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Aylık Toplam", f"{summary['monthly_price']:,} ₺".replace(",", "."))
            m2.metric("Aylık Soru Kotası", f"{summary['monthly_quota']:,}".replace(",", "."))
            m3.metric("Dijimentor Plan", f"{summary['coach_quota']} / Ay")

            rem_q = int(st.session_state.get("remaining_question_quota", summary["monthly_quota"]))
            rem_c = int(st.session_state.get("remaining_coach_quota", summary["coach_quota"]))
            q_pct = min(1.0, max(0.0, rem_q / max(1, int(summary["monthly_quota"]))))
            c_pct = min(1.0, max(0.0, rem_c / max(1, int(summary["coach_quota"]))))

            st.markdown("<div style='margin-top:0.55rem;'></div>", unsafe_allow_html=True)
            pb1, pb2 = st.columns(2)
            with pb1:
                st.progress(q_pct, text=f"Kullanılabilir Soru Kotası: {rem_q:,} / {summary['monthly_quota']:,}".replace(",", "."))
            with pb2:
                st.progress(c_pct, text=f"Kullanılabilir Dijimentor: {rem_c} / {summary['coach_quota']} Plan")

    st.markdown("</div>", unsafe_allow_html=True)
    return summary


def render_iyzico_simulator(
    package_name: str,
    monthly_price: int,
    summary: Optional[Dict[str, Any]] = None,
) -> None:
    bag_summary = summary if summary is not None else get_bag_summary()
    effective_monthly = int(bag_summary["monthly_price"]) if bag_summary["total_count"] > 0 else int(monthly_price)
    display_label = f"{package_name} ({bag_summary['total_count']} Paket)" if bag_summary["total_count"] > 0 else package_name

    st.markdown('<div class="kp-pay-box">', unsafe_allow_html=True)
    st.markdown("#### iyzico Taksit & Ödeme Simülatörü")
    st.caption("Dönem avantajı ve taksit planına göre tutarı hesaplayın, iyzico ödeme formu başlatın veya DOCX özet indirin.")

    c1, c2, c3 = st.columns(3)
    with c1:
        billing_cycle = st.selectbox("Dönem Seçimi", ["Aylık Abonelik", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"], key=f"cycle_{display_label}")
    with c2:
        installment = st.selectbox("Taksit Sayısı (iyzico)", [1, 2, 3, 6, 9, 12], index=2, key=f"inst_{display_label}")
    with c3:
        card_brand = st.selectbox("Kart Programı", ["Bonus", "Maximum", "World", "Axess", "Paraf", "Diğer Banka Kartı"], key=f"card_{display_label}")

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
        buyer_name = st.text_input("Ad", value="Sinan Gürbüz", key=f"iyz_name_{display_label}")
    with b_col2:
        buyer_surname = st.text_input("Soyad", value="Kurşun", key=f"iyz_surname_{display_label}")
    with b_col3:
        buyer_email = st.text_input("E-posta", value="iletisim@kursunkalemdijitalofis.com.tr", key=f"iyz_email_{display_label}")

    btn_c1, btn_c2, btn_c3 = st.columns([1.8, 1.8, 1.1])
    with btn_c1:
        if st.button("iyzico Ödeme Formu Başlat (API)", key=f"iyz_live_btn_{display_label}", type="primary", use_container_width=True):
            if total_amount <= 0:
                st.warning("Ödeme başlatmak için önce çantanıza en az bir paket ekleyin.")
            else:
                load_slot = st.empty()
                render_loading_skeleton(load_slot, lines=2)
                res_dict = initialize_iyzico_checkout(
                    buyer_name=buyer_name,
                    buyer_surname=buyer_surname,
                    buyer_email=buyer_email,
                    buyer_phone="+905350000000",
                    amount_try=float(total_amount),
                    basket_title=f"{BRAND_NAME} 2028 - {display_label} ({billing_cycle})",
                )
                load_slot.empty()
                if res_dict.get("status") == "success":
                    sync_plan_quotas(bag_summary, reset_usage=True)
                    record_activity("iyzico API", f"{display_label} ({total_amount:,} ₺) ödeme formu açıldı".replace(",", "."))
                    st.session_state["iyzico_last_response"] = res_dict
                    st.success("iyzico ödeme formu başarıyla oluşturuldu ve kota yenilendi!")
                else:
                    st.error(f"iyzico yanıtı: {res_dict.get('errorMessage', 'Bilinmeyen hata')}")

    with btn_c2:
        if st.button("Simülasyon ile Kota Aktifleştir", key=f"pay_btn_{display_label}", use_container_width=True):
            sync_plan_quotas(bag_summary, reset_usage=True)
            record_activity("iyzico Simülasyon", f"{display_label} ({total_amount:,} ₺) aktiflendi".replace(",", "."))
            st.success(
                f"**{display_label}** için **{total_amount:,} ₺** ({installment} taksit, "
                f"{card_brand}) tutarında kota ({bag_summary['monthly_quota']:,} soru) aktifleşti.".replace(",", ".")
            )

    with btn_c3:
        brief_docx = build_commercial_brief_docx(
            title=f"{display_label} Özeti (2028)",
            summary=bag_summary,
            billing_cycle=billing_cycle,
            installment=int(installment),
            total_amount=int(total_amount),
            extra_notes=f"Kart Programı: {card_brand}",
        )
        st.download_button(
            label="Özet DOCX",
            data=brief_docx,
            file_name="kursun_kalem_ofis_2028_cantasi_ozet.docx",
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
                f'<button style="background:#2563EB;color:#FFFFFF;border:none;border-radius:8px;'
                f'padding:0.55rem 1.15rem;font-weight:600;cursor:pointer;">'
                f'Güvenli iyzico Ödeme Sayfasını Aç</button></a></div>',
                unsafe_allow_html=True,
            )
        if form_html:
            st.components.v1.html(str(form_html), height=460, scrolling=True)

    st.markdown("</div>", unsafe_allow_html=True)


def render_student_xray_panel() -> None:
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Öğrenci Röntgeni (2028) — Kazanım, Hata Kök Neden &amp; Telafi Tarama</div>
            <div class="kp-studio-sub">
                Öğrencinin ders ve kazanım bazlı performans röntgenini çıkarın, hata türlerini ayrıştırın
                ve tek tıkla Soru Fabrikası, Dijimentor veya Veli Paneli'ne aktarın.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    f_col1, f_col2, f_col3 = st.columns([2.2, 1.8, 1.0])
    with f_col1:
        sec_student = st.selectbox("Öğrenci Profili Seçin", list(OGRENCI_RONTGEN_PROFILLERI.keys()), key="xray_student_sel")
    with f_col2:
        sec_scope = st.selectbox("Röntgen Kapsamı", ["Son 5 Deneme Ortalaması", "Son 30 Günlük Soru Çözümü", "Tüm Dönem Kümülatif"], key="xray_scope_sel")
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
        st.bar_chart(prof["hatalar"], color="#2563EB", height=225)

    st.markdown("#### Dijimentor Otomatik Telafi Reçetesi")
    st.markdown(
        f"""
        <div class="kp-card">
            <div style="font-weight:700;color:#1E2952;font-size:0.92rem;">
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

    a_col1, a_col2, a_col3 = st.columns(3)
    with a_col1:
        if st.button("Kritik Kazanım İçin Soru Üret", key="xray_to_factory", type="primary", use_container_width=True):
            st.session_state["factory_focus_topic"] = prof["hedef_konu"]
            st.session_state["factory_branch"] = prof["hedef_brans"]
            st.session_state["factory_grade"] = prof["hedef_sinif"]
            st.session_state["fab_theme_inp"] = prof["hedef_konu"]
            st.session_state["factory_sync_pending"] = True
            st.session_state["factory_grade_sync"] = True
            record_activity("Röntgen -> Fabrika", f"{prof['hedef_konu']} Soru Fabrikasına aktarıldı")
            navigate_to("Soru & Redaksiyon Fabrikası")
    with a_col2:
        if st.button("Dijimentor Planına Aktar", key="xray_to_mentor", use_container_width=True):
            st.session_state["dk_weak"] = prof["recete_ozet"]
            st.session_state["dk_focus"] = "Dikkat ve Çeldirici Tuzağı Azaltma"
            record_activity("Röntgen -> Dijimentor", f"{sec_student} bulguları mentorluğa aktarıldı")
            navigate_to("Dijimentor")
    with a_col3:
        if st.button("Veli Özeti olarak Aktar", key="xray_to_parent", use_container_width=True):
            st.session_state["parent_student_stats"]["weak_topic"] = prof["hedef_konu"]
            save_persisted_state()
            record_activity("Röntgen -> Veli", f"{sec_student} kritik kazanımı Veli Paneline aktarıldı")
            navigate_to("Veli Takip Paneli")


def render_parent_notification_center() -> None:
    st.markdown("---")
    st.markdown("### Veli Bilgilendirme & Otomatik Mesaj Merkezi")
    st.caption(
        "Veli bilgilendirme başlıkları altında otomatik mesaj taslağı oluşturun, "
        "duyuru akışına ekleyin ve kurumsal DOCX olarak indirin."
    )

    stats = st.session_state["parent_student_stats"]
    solved_safe = max(1, int(stats["weekly_solved"]))
    acc_pct = round((int(stats["correct"]) / solved_safe) * 100, 1)

    t_msg, t_feed, t_profile = st.tabs(["Mesaj Üreticisi", "Veli Duyuru Akışı", "Öğrenci & Veli Verileri"])

    with t_msg:
        c_left, c_right = st.columns([1.05, 1.35])
        with c_left:
            sel_cat = st.selectbox("Ana Konu Başlığı", list(PARENT_MESSAGE_TOPICS.keys()), key="pm_cat")
            subtopics = PARENT_MESSAGE_TOPICS[sel_cat]["subtopics"]
            sel_sub = st.selectbox("Alt Konu Başlığı", subtopics, key="pm_sub")
            channel = st.radio("Mesaj Formatı", ["WhatsApp / Kurumsal Mesaj", "Kısa SMS"], horizontal=True, key="pm_ch")

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
                    f"{BRAND_NAME} 2028: Sn. {stats['parent_name']}, {stats['student_name']} "
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
            mail_url = f"mailto:{stats['parent_email']}?subject={BRAND_NAME}%202028%20-%20{sel_sub}&body={encoded}"

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
                file_name="veli_bilgilendirme_mesaji_2028.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    with t_feed:
        flt = st.selectbox("Kategoriye Göre Filtrele", ["Tümü"] + list(PARENT_MESSAGE_TOPICS.keys()), key="pm_flt")
        items = [a for a in st.session_state["parent_announcements"] if flt == "Tümü" or a["category"] == flt]
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

if "loader_nonce" not in st.session_state:
    trigger_loader_bar()

render_top_loader_bar()

if "generated_test_text" not in st.session_state:
    st.session_state.generated_test_text = ""
if "dijimentor_plan" not in st.session_state:
    st.session_state.dijimentor_plan = ""
if "selected_tier_id" not in st.session_state:
    st.session_state.selected_tier_id = _persisted.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
if "backpack_cart" not in st.session_state and isinstance(_persisted.get("backpack_cart"), dict):
    st.session_state["backpack_cart"] = dict(_persisted["backpack_cart"])

ensure_bag_initialized()
bag_summary_global = get_bag_summary()

if "used_question_quota" not in st.session_state:
    if _persisted.get("used_question_quota") is not None:
        st.session_state["used_question_quota"] = int(_persisted["used_question_quota"])
    elif _persisted.get("remaining_question_quota") is not None:
        st.session_state["used_question_quota"] = max(
            0, int(bag_summary_global["monthly_quota"]) - int(_persisted["remaining_question_quota"])
        )
    else:
        st.session_state["used_question_quota"] = 0

if "used_coach_quota" not in st.session_state:
    if _persisted.get("used_coach_quota") is not None:
        st.session_state["used_coach_quota"] = int(_persisted["used_coach_quota"])
    elif _persisted.get("remaining_coach_quota") is not None:
        st.session_state["used_coach_quota"] = max(
            0, int(bag_summary_global["coach_quota"]) - int(_persisted["remaining_coach_quota"])
        )
    else:
        st.session_state["used_coach_quota"] = 0

sync_plan_quotas(bag_summary_global, reset_usage=False)

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
    st.session_state.parent_student_stats = merged_stats
if "parent_announcements" not in st.session_state:
    p_ann = _persisted.get("parent_announcements")
    st.session_state.parent_announcements = p_ann if isinstance(p_ann, list) and p_ann else _default_parent_announcements()
if "question_bank_2028" not in st.session_state:
    st.session_state["question_bank_2028"] = _persisted.get("question_bank_2028", []) if isinstance(_persisted.get("question_bank_2028"), list) else []
if "error_notebook_2028" not in st.session_state:
    st.session_state["error_notebook_2028"] = _persisted.get("error_notebook_2028", []) if isinstance(_persisted.get("error_notebook_2028"), list) else []
if "b2b_org_cache" not in st.session_state and isinstance(_persisted.get("b2b_org_cache"), dict):
    st.session_state["b2b_org_cache"] = _persisted["b2b_org_cache"]
if "last_b2b_summary" not in st.session_state and _persisted.get("last_b2b_summary"):
    st.session_state["last_b2b_summary"] = _persisted["last_b2b_summary"]

api_key_ready = resolve_gemini_api_key() is not None

# ==============================================================================
# KENAR ÇUBUĞU
# ==============================================================================
def _fmt(n: int) -> str:
    """1500 -> '1.500'"""
    return f"{int(n):,}".replace(",", ".")


with st.sidebar:
    st.markdown(
        f"""
        <div class="kp-sidebar-brand">
            <img src="{LOGO_DATA_URI}" alt=""/>
            <div>
                <div class="kp-sidebar-brand-title">{BRAND_SHORT}</div>
                <div class="kp-sidebar-brand-sub">Dijital Ofis</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Gruplanmış, ikonlu menü
    for group_title, items in NAV_GROUPS:
        if group_title:
            st.markdown(f'<div class="kp-nav-title">{html.escape(group_title)}</div>', unsafe_allow_html=True)
        for p_name, p_label, p_icon in items:
            is_current = st.session_state.active_page == p_name
            if st.button(
                p_label,
                key=f"nav_{p_name}",
                icon=p_icon,
                type="primary" if is_current else "secondary",
                use_container_width=True,
            ):
                clear_page_ephemeral_state(p_name)
                st.session_state.active_page = p_name
                trigger_loader_bar()
                st.rerun()

    # Paket & kota kartı (ilerleme çubuklu)
    total_q = int(bag_summary_global["monthly_quota"])
    total_c = int(bag_summary_global["coach_quota"])
    rem_q = int(st.session_state.remaining_question_quota)
    rem_c = int(st.session_state.remaining_coach_quota)
    q_pct = round(100 * rem_q / total_q) if total_q else 0
    c_pct = round(100 * rem_c / total_c) if total_c else 0
    status_color = "#16A34A" if api_key_ready else "#D97706"
    status_text = "Sistem aktif" if api_key_ready else "API anahtarı bekleniyor"

    st.markdown(
        f"""
        <div class="kp-sidebar-meta">
            <div class="kp-meta-head"><span>Aktif paketin</span></div>
            <div class="kp-meta-sub">{bag_summary_global['total_count']} paket · {_fmt(bag_summary_global['monthly_price'])} ₺/ay</div>
            <div class="kp-meter-label"><span>Soru kotası</span><span>{_fmt(rem_q)} / {_fmt(total_q)}</span></div>
            <div class="kp-meter"><span style="width:{q_pct}%"></span></div>
            <div class="kp-meter-label"><span>Dijimentor</span><span>{rem_c} / {total_c}</span></div>
            <div class="kp-meter"><span style="width:{c_pct}%"></span></div>
            <div class="kp-meter-label" style="margin-top:0.6rem;">
                <span>Soru havuzu: <b>{len(st.session_state['question_bank_2028'])}</b></span>
                <span>Hata defteri: <b>{len(st.session_state['error_notebook_2028'])}</b></span>
            </div>
            <div class="kp-status"><span class="kp-status-dot" style="background:{status_color}"></span>{status_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Tehlikeli işlem: göz önünden kaldırıldı, "Ayarlar" altında
    with st.expander("Ayarlar", icon=":material/settings:"):
        if not st.session_state.get("confirm_reset_session", False):
            if st.button("Oturumu sıfırla", key="reset_session_btn", use_container_width=True):
                st.session_state["confirm_reset_session"] = True
                st.rerun()
        else:
            st.warning("Yerel durum sıfırlansın mı? (Yedek alınır)")
            rc1, rc2 = st.columns(2)
            with rc1:
                if st.button("Evet", key="confirm_reset_yes", type="primary", use_container_width=True):
                    try:
                        if STATE_FILE_PATH.exists():
                            STATE_BACKUP_PATH.write_text(STATE_FILE_PATH.read_text(encoding="utf-8"), encoding="utf-8")
                            STATE_FILE_PATH.unlink()
                    except Exception:
                        pass
                    for k in list(st.session_state.keys()):
                        del st.session_state[k]
                    st.rerun()
            with rc2:
                if st.button("İptal", key="confirm_reset_no", use_container_width=True):
                    st.session_state["confirm_reset_session"] = False
                    st.rerun()

# ==============================================================================
# ÜST ŞERİT
# ==============================================================================
st.markdown(
    """
    <div class="kp-topbar">
        <span class="kp-topbar-brand">Kurşun Kalem Dijital Ofis</span>
        <span>Destek: <a href="tel:08505321629">0850 532 16 29</a> · <a href="mailto:destek@kursunkalemdijitalofis.com.tr">destek@kursunkalemdijitalofis.com.tr</a></span>
    </div>
    """,
    unsafe_allow_html=True,
)

current_page = st.session_state.active_page
if current_page != "Ana Sayfa":
    render_page_strip(current_page)

# ==============================================================================
# 1. ANA SAYFA
# ==============================================================================
if current_page == "Ana Sayfa":
    stats_now = st.session_state["parent_student_stats"]
    solved_now = max(1, int(stats_now["weekly_solved"]))
    acc_now = round(int(stats_now["correct"]) / solved_now * 100)

    # ---------- HERO ----------
    hero_l, hero_r = st.columns([1.2, 1.0], gap="large", vertical_alignment="center")

    with hero_l:
        st.markdown(
            """
            <span class="kp-eyebrow">2028 müfredatına tam uyumlu</span>
            <h1 class="kp-hero-title">Hayalindeki <em>2028 hedefi</em> için ihtiyacın olan her şey, tek platformda.</h1>
            <p class="kp-hero-lead">
                Soru bankası, haftalık dijital mentorluk, mini denemeler ve kişisel eksik analizi.
                Öğrenci, veli ve okul aynı ekranda.
            </p>
            """,
            unsafe_allow_html=True,
        )
        cta1, cta2, _ = st.columns([1, 1.15, 0.6])
        with cta1:
            if st.button("Ücretsiz dene", key="hero_cta_primary", type="primary", use_container_width=True):
                navigate_to("Paketler & Abonelik Vitrini")
        with cta2:
            if st.button("Kurumlar için demo", key="hero_cta_secondary", use_container_width=True):
                navigate_to("B2B Lisans & Kurum")
        st.markdown(
            """
            <div class="kp-stats">
                <div><div class="kp-stat-num">142.500+</div><div class="kp-stat-lbl">soru havuzu</div></div>
                <div><div class="kp-stat-num">18.420</div><div class="kp-stat-lbl">ön kayıtlı öğrenci ve veli</div></div>
                <div><div class="kp-stat-num">48</div><div class="kp-stat-lbl">anlaşmalı kurum</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with hero_r:
        # Eski SVG vitrin yerine öğrencinin gerçek haftalık özeti
        st.markdown(
            f"""
            <div class="kp-preview">
                <div class="kp-preview-head">
                    <div><small>Bu hafta</small><strong>{html.escape(str(stats_now['student_name']))}</strong></div>
                    <span class="kp-pill">{html.escape(str(stats_now['grade']))}</span>
                </div>
                <div class="kp-preview-stats">
                    <div class="kp-preview-stat"><b>{_fmt(stats_now['weekly_solved'])}</b><span>çözülen soru</span></div>
                    <div class="kp-preview-stat"><b>%{acc_now}</b><span>doğruluk</span></div>
                    <div class="kp-preview-stat"><b>%{int(stats_now['mentor_completion'])}</b><span>plan uyumu</span></div>
                </div>
                <div class="kp-meter" style="margin-bottom:0.9rem;"><span style="width:{int(stats_now['mentor_completion'])}%"></span></div>
                <div class="kp-preview-row"><span>Güçlü alan</span><span>{html.escape(str(stats_now['strong_topic']))}</span></div>
                <div class="kp-preview-row"><span>Pekiştirilecek</span><span>{html.escape(str(stats_now['weak_topic']))}</span></div>
                <div class="kp-preview-row"><span>Sıradaki deneme</span><span>{html.escape(str(stats_now['next_exam_date']))}</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- SINIF SEÇİCİ ----------
    with st.container(border=True):
        sel_l, sel_r = st.columns([1.3, 1.0], vertical_alignment="center")
        with sel_l:
            st.markdown(
                '<h2 class="kp-section-title">Sınıfını ya da sınavını seç</h2>'
                '<p class="kp-section-sub" style="margin-bottom:0;">Sana uygun modül otomatik açılır.</p>',
                unsafe_allow_html=True,
            )
        with sel_r:
            group_names = list(GRADE_GROUPS.keys())
            selected_group = st.segmented_control(
                "Kademe",
                options=group_names,
                default=group_names[0],
                key="grade_group_sel",
                label_visibility="collapsed",
            ) or group_names[0]

        group_items = GRADE_GROUPS[selected_group]
        g_idx = group_names.index(selected_group)
        chip_cols = st.columns(4)
        for idx, chip_label in enumerate(group_items):
            chip = CHIP_BY_LABEL[chip_label]
            with chip_cols[idx % 4]:
                if st.button(chip["label"], key=f"grade_chip_{g_idx}_{idx}", use_container_width=True):
                    st.session_state["factory_grade"] = chip["grade"]
                    st.session_state["factory_grade_sync"] = True
                    st.session_state["dk_exam"] = chip["exam"]
                    record_activity("Hızlı Kademe (2028)", f"{chip['label']} seçildi")
                    navigate_to(chip["page"])

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- MODÜLLER ----------
    st.markdown(
        '<h2 class="kp-section-title">Modüller</h2>'
        '<p class="kp-section-sub">Her biri tek başına çalışır, birlikte daha güçlüdür.</p>',
        unsafe_allow_html=True,
    )
    mod_cols = st.columns(len(MODULE_CARDS))
    for idx, (code, m_title, m_desc, m_page, m_bg, m_fg) in enumerate(MODULE_CARDS):
        with mod_cols[idx]:
            with st.container(border=True):
                st.markdown(
                    f"""
                    <div class="kp-mod-icon" style="background:{m_bg};color:{m_fg};">{html.escape(code)}</div>
                    <div class="kp-mod-title">{html.escape(m_title)}</div>
                    <div class="kp-mod-desc">{html.escape(m_desc)}</div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Aç →", key=f"mod_open_{idx}"):
                    navigate_to(m_page)

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- ÖNCELİKLİ EKSİKLER ----------
    pr_l, pr_r = st.columns([4, 1], vertical_alignment="bottom")
    with pr_l:
        st.markdown(
            '<h2 class="kp-section-title">Öncelikli eksiklerin</h2>'
            '<p class="kp-section-sub">En çok puan kazandıracak konular, öncelik sırasıyla.</p>',
            unsafe_allow_html=True,
        )
    with pr_r:
        if st.button("Tüm analizi gör →", key="prio_all"):
            navigate_to("Öğrenci Röntgeni")

    with st.container(border=True):
        for idx, p_item in enumerate(ONCELIKLI_EKSIK_SIRASI):
            if idx:
                st.markdown('<div class="kp-prio-divider"></div>', unsafe_allow_html=True)
            c_num, c_txt, c_meter, c_btn = st.columns([0.45, 5, 2.2, 1.2], vertical_alignment="center")
            with c_num:
                st.markdown(f'<div class="kp-prio-num">{idx + 1}</div>', unsafe_allow_html=True)
            with c_txt:
                st.markdown(
                    f'<div class="kp-prio-title">{html.escape(p_item["topic"])}</div>'
                    f'<div class="kp-prio-sub">{html.escape(p_item["branch"])} · {html.escape(p_item["action"])}</div>',
                    unsafe_allow_html=True,
                )
            with c_meter:
                pct = p_item.get("pct")
                if pct is not None:
                    bar_color = "#DC2626" if pct < 50 else "#EA580C" if pct < 60 else "#CA8A04"
                    st.markdown(
                        f'<div class="kp-meter-label"><span>Doğru oranı</span><b style="color:#0F172A">%{pct}</b></div>'
                        f'<div class="kp-meter"><span style="width:{pct}%;background:{bar_color}"></span></div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown('<span class="kp-tag-warn">Dikkat gerektiriyor</span>', unsafe_allow_html=True)
            with c_btn:
                if st.button("Çalış", key=f"prio_btn_{idx}", use_container_width=True):
                    st.session_state["factory_focus_topic"] = p_item["topic"]
                    st.session_state["factory_branch"] = p_item["branch"]
                    st.session_state["factory_grade"] = p_item["grade"]
                    st.session_state["fab_theme_inp"] = p_item["topic"]
                    st.session_state["factory_sync_pending"] = True
                    st.session_state["factory_grade_sync"] = True
                    record_activity("Öncelikli Akış (2028)", f"{p_item['topic']} seçildi")
                    navigate_to("Soru & Redaksiyon Fabrikası")

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- GERİ SAYIM ----------
    head_col, toggle_col = st.columns([4, 1], vertical_alignment="center")
    with head_col:
        st.markdown('<h2 class="kp-section-title">2028 LGS\'ye kalan süre</h2>', unsafe_allow_html=True)
    with toggle_col:
        live_counter = st.toggle("Canlı sayaç", value=st.session_state.live_tick)
        st.session_state.live_tick = live_counter

    counter_placeholder = st.empty()
    with counter_placeholder.container():
        render_countdown_boxes(LGS_2028_TARGET)

    yks_d, yks_h, _, _ = format_countdown(YKS_2028_TARGET)
    st.caption(f"LGS: 11 Haziran 2028 · YKS (17 Haziran 2028) için kalan süre: {yks_d} gün {yks_h} saat")

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- TELEMETRİ & AKTİVİTE ----------
    ana_l, ana_r = st.columns([2, 1], gap="large")
    with ana_l:
        st.markdown('<h2 class="kp-section-title" style="font-size:1.15rem;">Haftalık soru üretimi</h2>', unsafe_allow_html=True)
        telemetry_data = {"Pzt": 16800, "Sal": 18200, "Çar": 19500, "Per": 21100, "Cum": 22400, "Cmt": 25900, "Paz": 24800}
        st.area_chart(telemetry_data, color="#1D4ED8", height=210)
    with ana_r:
        st.markdown('<h2 class="kp-section-title" style="font-size:1.15rem;">Son aktiviteler</h2>', unsafe_allow_html=True)
        for item in st.session_state.get("activity_log", [])[:5]:
            st.markdown(
                f"""
                <div style="padding:0.6rem 0;border-bottom:1px solid #E2E8F0;">
                    <div style="font-size:0.78rem;color:#64748B;">{html.escape(item['time'])} · {html.escape(item['module'])}</div>
                    <div style="font-size:0.9rem;color:#0F172A;margin-top:0.1rem;">{html.escape(item['summary'])}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    if st.session_state.live_tick:
        time.sleep(1)
        st.rerun()

# ==============================================================================
# 2. SORU & REDAKSİYON FABRİKASI (5–12. SINIF MÜFREDAT + KADEMELİ SERİ + HAVUZ)
# ==============================================================================
elif current_page == "Soru & Redaksiyon Fabrikası":
    tab_gen, tab_bank = st.tabs(["Soru Üretim & Redaksiyon Stüdyosu", f"Soru Havuzu ({len(st.session_state['question_bank_2028'])})"])

    with tab_gen:
        if st.session_state.get("factory_grade") in GRADE_LEVELS_2028 and st.session_state.get("factory_grade_sync"):
            st.session_state["fab_grade_sel"] = st.session_state["factory_grade"]
            st.session_state["factory_grade_sync"] = False

        cur_grade_init = st.session_state.get("fab_grade_sel", "8. Sınıf (LGS)")
        if cur_grade_init not in GRADE_LEVELS_2028:
            cur_grade_init = "8. Sınıf (LGS)"

        if st.session_state.get("factory_sync_pending"):
            st.session_state["fab_grade_sel"] = st.session_state.get("factory_grade", cur_grade_init)
            g_tmp = st.session_state["fab_grade_sel"]
            b_list = get_branches_for_grade_2028(g_tmp)
            target_b = st.session_state.get("factory_branch")
            st.session_state["fab_branch_sel"] = target_b if target_b in b_list else b_list[0]
            a_list = get_areas_for_grade_branch_2028(g_tmp, st.session_state["fab_branch_sel"])
            st.session_state["fab_area_sel"] = a_list[0]
            o_list = get_outcomes_for_2028(g_tmp, st.session_state["fab_branch_sel"], a_list[0])
            focus_t = st.session_state.get("factory_focus_topic")
            if focus_t and focus_t not in o_list:
                st.session_state["fab_custom_outcome_chk"] = True
                st.session_state["fab_custom_outcome_inp"] = focus_t
            else:
                st.session_state["fab_outcome_sel"] = focus_t if focus_t in o_list else o_list[0]
            if focus_t:
                st.session_state["fab_theme_inp"] = focus_t
            st.session_state["factory_sync_pending"] = False

        if st.session_state.get("factory_focus_topic"):
            st.info(f"Aktarılan odak konu: **{st.session_state['factory_focus_topic']}**")

        with st.expander("5–12. Sınıf Müfredat, Zorluk & Kalite Parametreleri (2028)", expanded=True):
            r1c1, r1c2, r1c3 = st.columns(3)
            with r1c1:
                institution = st.text_input("Kurum / Ofis Adı", value=BRAND_NAME, key="fab_inst_inp")
                grade = st.selectbox("Sınıf Seviyesi (5–12)", GRADE_LEVELS_2028, key="fab_grade_sel")
                branches = get_branches_for_grade_2028(grade)
                if st.session_state.get("fab_branch_sel") not in branches:
                    st.session_state["fab_branch_sel"] = branches[0]
                branch = st.selectbox("Branş Seçimi", branches, key="fab_branch_sel")
            with r1c2:
                areas = get_areas_for_grade_branch_2028(grade, branch)
                if st.session_state.get("fab_area_sel") not in areas:
                    st.session_state["fab_area_sel"] = areas[0]
                learning_area = st.selectbox("Öğrenme Alanı / Tema", areas, key="fab_area_sel")

                use_custom_outcome = st.checkbox("Özel / Manuel Kazanım veya Konu Gir", key="fab_custom_outcome_chk")
                if use_custom_outcome:
                    outcome = st.text_input(
                        "Özel Kazanım / Zümre Konusu",
                        value=st.session_state.get("fab_custom_outcome_inp", ""),
                        placeholder="Örn: Paragrafta yardımcı düşünce & çeldirici kurgusu",
                        key="fab_custom_outcome_inp",
                    ) or "Özel Kazanım Odaklı Soru Üretimi"
                else:
                    outcomes = get_outcomes_for_2028(grade, branch, learning_area)
                    if st.session_state.get("fab_outcome_sel") not in outcomes:
                        st.session_state["fab_outcome_sel"] = outcomes[0]
                    outcome = st.selectbox("Hedef Kazanım", outcomes, key="fab_outcome_sel")

                if "fab_theme_inp" not in st.session_state:
                    st.session_state["fab_theme_inp"] = st.session_state.get("factory_focus_topic") or "2028 DİJİTAL OFİS & HAYAT"
                context_theme = st.text_input("Bağlam / Tema / Senaryo", key="fab_theme_inp")
            with r1c3:
                production_mode = st.radio(
                    "Üretim Modu",
                    ["Standart Test Paketi", "Kademeli Zorluk Serisi (Kolay → Orta → Zor)"],
                    key="fab_prod_mode",
                )
                question_count = st.slider("Soru Adedi", 1, 10, 3, key="fab_q_cnt")
                difficulty = st.select_slider(
                    "Temel Zorluk Düzeyi",
                    options=["Kolay", "Orta", "Orta-Zor", "Zor (2028 LGS/ÖSYM)"],
                    value="Orta-Zor",
                    key="fab_diff_sld",
                )
                distractor_mode = st.selectbox(
                    "Çeldirici & Kök Stratejisi",
                    [
                        "Muhakeme & Çeldirici Tuzağı Dengeli",
                        "Olumsuz Kök ('Değildir / Ulaşılamaz') Odaklı",
                        "Öncüllü & Bağlamsal Analiz Odaklı",
                    ],
                    index=0,
                    key="fab_dist_sel",
                )
                auto_review = st.checkbox("Çift Katmanlı Redaksiyon & Çeldirici Kalite Denetimi", value=True, key="fab_rev_chk")

        if st.button("Kurumsal Soru Paketi Üret (2028)", type="primary", key="fab_gen_btn"):
            quota_ok, quota_msg = consume_quota("question", int(question_count))
            if not quota_ok:
                st.error(quota_msg)
            else:
                mode_instruction = (
                    f"Her soruyu '{difficulty}' seviyesinde üret."
                    if production_mode == "Standart Test Paketi"
                    else "Soruları Kolay (pekiştirme), Orta (kazanım) ve Zor (2028 LGS/YKS muhakeme) olarak kademelendir."
                )
                sys_prompt = (
                    "Sen Türkiye Yüzyılı Maarif Modeli, 2028 LGS ve 2028 YKS ölçme-değerlendirme standartlarında "
                    "uzman bir kıdemli test editörüsün. Soruların özgün, dilbilgisi ve noktalama açısından "
                    "kusursuz, çeldiricilerinin uzunluk ve dilbilgisi paralelliği açısından dengeli olmasını sağla. "
                    "Her soruyu 'Soru 1', 'Soru 2'... şeklinde numaralandır ve şıkları 'A)', 'B)', 'C)', 'D)' "
                    "olarak yaz. Her sorunun altında kısa bir [Çözüm & Şık Gerekçesi] ve redaksiyon açıksa "
                    "[Redaksiyon & Çeldirici Kalite Notu] ekle; en sonda mutlaka 'Cevap Anahtarı' başlığı altında "
                    "'1-A', '2-C' formatında net anahtar ver."
                )
                user_prompt = (
                    f"Kurum: {institution} (2028)\nSınıf Seviyesi: {grade}\nBranş: {branch}\n"
                    f"Öğrenme Alanı: {learning_area}\nHedef Kazanım: {outcome}\n"
                    f"Bağlam / Tema / Senaryo: {context_theme}\nZorluk Kurgusu: {mode_instruction}\n"
                    f"Çeldirici & Kök Stratejisi: {distractor_mode}\nSoru Adedi: {question_count}\n"
                    f"Redaksiyon & Kalite Denetimi: {'Aktif' if auto_review else 'Pasif'}\n\n"
                    "Lütfen numaralandırılmış çoktan seçmeli (Soru 1, Soru 2... ve A, B, C, D şıkları) özgün sorular üret."
                )
                skel_slot = st.empty()
                render_loading_skeleton(skel_slot, lines=4)
                try:
                    result = call_gemini_text(prompt=user_prompt, system_instruction=sys_prompt)
                    skel_slot.empty()
                    st.session_state.generated_test_text = result
                    st.session_state["last_fab_meta"] = {
                        "grade": grade,
                        "branch": branch,
                        "outcome": outcome,
                        "difficulty": difficulty if production_mode == "Standart Test Paketi" else "Kademeli Seri",
                    }
                    record_activity("Soru Fabrikası (2028)", f"{grade} • {branch} • {question_count} soru üretildi")
                    st.success(
                        f"Soru paketi oluşturuldu. Kalan soru kotası: "
                        f"{int(st.session_state.remaining_question_quota):,}".replace(",", ".")
                    )
                except Exception as exc:
                    skel_slot.empty()
                    refund_quota("question", int(question_count))
                    st.error(f"Soru üretim hatası oluştu: {exc}")

        if st.session_state.generated_test_text:
            st.markdown("---")
            st.markdown("### Üretilen Test, Çözüm Gerekçeleri ve Redaksiyon Çıktısı")
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
            act_c1, act_c2, act_c3, act_c4 = st.columns(4)
            with act_c1:
                st.download_button(
                    label="Testi DOCX İndir",
                    data=docx_bytes,
                    file_name=f"kursun_kalem_ofis_2028_{branch.lower().replace(' ', '_')}_test.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            with act_c2:
                if st.button("Soru Havuzuna Kaydet", use_container_width=True):
                    add_to_question_bank(
                        grade=grade,
                        branch=branch,
                        outcome=outcome,
                        difficulty=difficulty,
                        content_text=st.session_state.generated_test_text,
                    )
                    record_activity("Soru Havuzu", f"{grade} • {branch} testi havuza kaydedildi")
                    st.toast("Test Soru Havuzuna kaydedildi.")
                    st.rerun()
            with act_c3:
                if st.button("Mini Deneme Olarak Çöz", use_container_width=True):
                    st.session_state["portal_exam_source"] = st.session_state.generated_test_text
                    record_activity("Köprü", "Test Öğrenci Portalı mini denemesine aktarıldı")
                    navigate_to("Öğrenci Portalı & Çözüm")
            with act_c4:
                if st.button("İlk Soruyu Çözüm Rehberine Gönder", use_container_width=True):
                    parsed_q = parse_quiz_questions(st.session_state.generated_test_text)
                    st.session_state["portal_question_draft"] = (
                        parsed_q[0]["full_text"] if parsed_q else st.session_state.generated_test_text[:1000]
                    )
                    navigate_to("Öğrenci Portalı & Çözüm")

    with tab_bank:
        bank_items: List[Dict[str, str]] = st.session_state.get("question_bank_2028", [])
        if not bank_items:
            st.info("Soru havuzunda henüz kayıtlı test bulunmuyor. Üretim sekmesinden **Soru Havuzuna Kaydet** butonunu kullanabilirsiniz.")
        else:
            combined_bank_text = "\n\n".join(
                [f"=== [{b['grade']} | {b['branch']} | {b['difficulty']}] {b['outcome']} ===\n{b['content']}" for b in bank_items]
            )
            bundle_docx = build_test_docx(
                institution=BRAND_NAME,
                branch="Karma Zümre Havuzu",
                grade="5–12. Sınıf",
                learning_area="Kayıtlı Soru Havuzu",
                outcome="Havuzdaki Seçili Kazanımlar",
                context_theme="2028 Kurumsal Soru Kitapçığı",
                generated_content=combined_bank_text,
            )
            st.download_button(
                label=f"Tüm Havuzu ({len(bank_items)} Kayıt) Toplu DOCX İndir",
                data=bundle_docx,
                file_name="kursun_kalem_2028_soru_havuzu.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            for b_item in bank_items:
                with st.expander(f"[{b_item['id']}] {b_item['grade']} • {b_item['branch']} ({b_item['difficulty']}) — {b_item['date']}"):
                    st.caption(f"Kazanım: {b_item['outcome']}")
                    st.markdown(b_item["content"])
                    if st.button("Bu Testi Portala Aktar", key=f"load_bank_{b_item['id']}"):
                        st.session_state["portal_exam_source"] = b_item["content"]
                        navigate_to("Öğrenci Portalı & Çözüm")

# ==============================================================================
# 3. ÖĞRENCİ PORTALI & ÇÖZÜM (+ SÜRE TAKİBİ + HATA DEFTERİ)
# ==============================================================================
elif current_page == "Öğrenci Portalı & Çözüm":
    tab_exam, tab_hint, tab_err = st.tabs([
        "İnteraktif Mini Deneme & Süre Analizi",
        "Adımlı Çözüm Asistanı",
        f"Hata Defteri ({len(st.session_state['error_notebook_2028'])})",
    ])

    with tab_exam:
        exam_source = st.text_area(
            "Deneme Soru Metni (Soru Fabrikasından otomatik gelir veya buraya yapıştır)",
            value=st.session_state.get("portal_exam_source") or st.session_state.get("generated_test_text", ""),
            height=120,
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
            if "exam_start_ts" not in st.session_state:
                st.session_state["exam_start_ts"] = time.time()

            st.markdown(f"#### Aktif Mini Deneme ({len(questions)} Soru)")
            st.caption("Boş bırakmak isterseniz **Boş Bırak** seçeneğini işaretleyebilir veya seçim yapmayabilirsiniz; boş sorular yanlışa sayılmaz.")
            user_answers: Dict[int, Optional[str]] = {}
            q_durations: Dict[int, float] = {}

            for idx, q in enumerate(questions):
                with st.container(border=True):
                    st.markdown(q["stem"])
                    c_opt, c_sec = st.columns([3.4, 1.0])
                    with c_opt:
                        opts = list(q["options"].keys()) + ["BOŞ"]
                        chosen = st.radio(
                            f"Soru {q['num']} Yanıtın",
                            options=opts,
                            index=None,
                            format_func=lambda k, o_map=q["options"]: "Boş Bırak" if k == "BOŞ" else f"{k}) {o_map[k]}",
                            horizontal=True,
                            key=f"mini_q_{q['num']}_{idx}",
                        )
                        user_answers[idx] = None if chosen in (None, "BOŞ") else chosen
                    with c_sec:
                        q_durations[idx] = float(
                            st.number_input(
                                f"Süre (sn) • S{q['num']}",
                                min_value=5,
                                max_value=300,
                                value=65,
                                step=5,
                                key=f"mini_dur_{q['num']}_{idx}",
                            )
                        )

            if st.button("Denemeyi Bitir ve Netimi Hesapla", type="primary"):
                correct_cnt = 0
                wrong_cnt = 0
                blank_cnt = 0
                ungraded_cnt = 0
                wrong_questions: List[Dict[str, Any]] = []
                result_rows: List[Dict[str, str]] = []
                total_sec = 0.0

                for idx, q in enumerate(questions):
                    u_ans = user_answers[idx]
                    c_ans = q.get("correct")
                    sec_val = q_durations.get(idx, 60.0)
                    total_sec += sec_val

                    if u_ans is None:
                        blank_cnt += 1
                        status_txt = "Boş"
                    elif not c_ans:
                        ungraded_cnt += 1
                        status_txt = "Anahtar Yok"
                    elif u_ans == c_ans:
                        correct_cnt += 1
                        status_txt = "Doğru"
                    else:
                        wrong_cnt += 1
                        wrong_questions.append(q)
                        status_txt = "Yanlış"
                        add_to_error_notebook(
                            question_num=q["num"],
                            question_text=q["full_text"],
                            user_ans=u_ans,
                            correct_ans=c_ans,
                            elapsed_sec=sec_val,
                        )

                    result_rows.append(
                        {
                            "Soru": f"Soru {q['num']}",
                            "Senin Yanıtın": u_ans or "Boş",
                            "Doğru Cevap": c_ans or "Belirtilmedi",
                            "Süre (sn)": f"{int(sec_val)} sn",
                            "Durum": status_txt,
                        }
                    )

                graded_total = correct_cnt + wrong_cnt
                net_score = max(0.0, correct_cnt - (wrong_cnt / 3.0))
                accuracy = (correct_cnt / graded_total * 100.0) if graded_total > 0 else 0.0
                avg_sec = total_sec / max(1, len(questions))

                st.session_state["last_exam_report"] = {
                    "correct": correct_cnt,
                    "wrong": wrong_cnt,
                    "blank": blank_cnt,
                    "net": net_score,
                    "accuracy": accuracy,
                    "avg_sec": avg_sec,
                    "rows": result_rows,
                    "first_wrong": wrong_questions[0]["full_text"] if wrong_questions else "",
                }
                record_activity(
                    "Mini Deneme",
                    f"{correct_cnt}D {wrong_cnt}Y {blank_cnt}B • {net_score:.2f} Net • Ort {avg_sec:.0f} sn/soru",
                )

            report = st.session_state.get("last_exam_report")
            if report:
                st.markdown("---")
                r1, r2, r3, r4 = st.columns(4)
                r1.metric("Doğru / Boş", f"{report['correct']}D / {report['blank']}B")
                r2.metric("Yanlış Sayısı", f"{report['wrong']} Soru")
                r3.metric("Hesaplanan Net (3Y=1D)", f"{report['net']:.2f} Net")
                r4.metric("Ort. Soru Süresi", f"{report['avg_sec']:.0f} sn (%{report['accuracy']:.1f})")

                st.dataframe(report["rows"], use_container_width=True, hide_index=True)

                if report.get("first_wrong"):
                    wa_c1, wa_c2 = st.columns(2)
                    with wa_c1:
                        if st.button("Yanlış Soruyu Çözüm Rehberine Aktar", use_container_width=True):
                            st.session_state["portal_question_draft"] = report["first_wrong"]
                            st.toast("Yanlış soru Adımlı Çözüm Asistanı sekmesine aktarıldı.")
                            st.rerun()
                    with wa_c2:
                        if st.button("Yanlış Sorudan Benzer Soru Üret", type="primary", use_container_width=True):
                            snippet = re.sub(r"\s+", " ", report["first_wrong"])[:110]
                            st.session_state["factory_focus_topic"] = f"Telafi: {snippet}"
                            st.session_state["fab_theme_inp"] = f"Telafi: {snippet}"
                            record_activity("Telafi Akışı", "Yanlış sorudan benzer soru üretimine geçildi")
                            navigate_to("Soru & Redaksiyon Fabrikası")

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
                skel_hint = st.empty()
                render_loading_skeleton(skel_hint, lines=3)
                try:
                    ans = call_gemini_text(prompt=prompt, system_instruction=sys_inst)
                    skel_hint.empty()
                    record_activity("Çözüm Asistanı", f"{hint_mode} üretildi")
                    st.markdown("#### Çözüm Analizi")
                    st.markdown(ans)
                except Exception as exc:
                    skel_hint.empty()
                    st.error(f"Çözüm üretilirken hata oluştu: {exc}")

    with tab_err:
        err_list: List[Dict[str, Any]] = st.session_state.get("error_notebook_2028", [])
        if not err_list:
            st.info("Hata defterinde henüz kayıt bulunmuyor. Mini denemede yanlış işaretlenen sorular otomatik buraya eklenir.")
        else:
            st.caption("Aralıklı tekrar takvimi: 1. Gün • 3. Gün • 7. Gün")
            for item in err_list:
                with st.container(border=True):
                    st.markdown(
                        f"**[{item['id']}] Soru {item['num']}** • Teşhis: **{item['diagnosis']}** • "
                        f"Senin: `{item['user_ans']}` / Doğru: `{item['correct_ans']}` • Süre: `{item['elapsed_sec']} sn`"
                    )
                    st.caption(
                        f"Tekrar Hatırlatmaları → 1. Gün: {item['review_1d']} | "
                        f"3. Gün: {item['review_3d']} | 7. Gün: {item['review_7d']}"
                    )
                    st.markdown(item["question_text"])
                    if st.button("Çözüm Rehberinde Aç", key=f"err_solve_{item['id']}"):
                        st.session_state["portal_question_draft"] = item["question_text"]
                        st.toast("Soru Adımlı Çözüm Asistanına aktarıldı.")

# ==============================================================================
# 4. DİJİMENTOR (2028 + AŞAMALI KAMP PLANI)
# ==============================================================================
elif current_page == "Dijimentor":
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Dijimentor (2028) — Akıllı Sınav, Faz Planı &amp; Performans Mentorluğu</div>
            <div class="kp-studio-sub">
                2028 hedef sınavına, haftalık çalışma kapasitene ve çeldirici/zaman analizine göre
                aşamalı kamp ve haftalık aksiyon planı üretir.
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
            <div class="kp-kpi-value">Dijimentor 2028</div>
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
            <div class="kp-kpi-label">2028 Kamp Fazları</div>
            <div class="kp-kpi-value">3 Aşama</div>
            <div class="kp-kpi-sub">Kazanım • Branş Hızı • Genel Deneme</div>
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
    st.markdown("#### Hızlı 2028 Mentorluk Senaryoları")
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
                st.session_state["fab_theme_inp"] = preset_data["topic"]
                st.session_state["factory_sync_pending"] = True
                st.session_state["factory_grade_sync"] = True
                record_activity("Dijimentor Yönlendirme", f"{preset_title} seçildi ve Soru Fabrikasına yönlendirildi")
                navigate_to("Soru & Redaksiyon Fabrikası")
    with p_cols[-1]:
        if st.button("Planı Temizle", key="clear_dk_plan", use_container_width=True):
            st.session_state.dijimentor_plan = ""
            st.session_state["dk_active_preset"] = ""
            st.rerun()

    st.markdown("---")
    st.markdown("#### Öğrenci 2028 Hedef ve Çalışma Profili")

    exam_opts = ["2028 LGS", "2028 TYT - AYT", "5-7. Sınıf Akademik Gelişim"]
    stored_exam = st.session_state.get("dk_exam", "2028 LGS")
    if stored_exam not in exam_opts:
        stored_exam = "2028 LGS"

    r1_a, r1_b, r1_c = st.columns(3)
    with r1_a:
        target_exam = st.selectbox(
            "Hedef Sınav / Kademe",
            exam_opts,
            index=exam_opts.index(stored_exam),
            key="dk_exam_sel",
        )
        st.session_state["dk_exam"] = target_exam

        weekly_hours = st.slider("Haftalık Çalışma Süresi (Saat)", min_value=4, max_value=45, value=current_hours, key="dk_hours_sld")
        st.session_state["dk_hours"] = weekly_hours

    with r1_b:
        target_goal = st.text_input(
            "Hedef Puan / Net / Okul Bandı",
            value=st.session_state.get("dk_goal", "Nitelikli lise / yüksek net ve sıfıra yakın dikkat hatası"),
            key="dk_goal_inp",
        )
        st.session_state["dk_goal"] = target_goal

        focus_opts = [
            "Haftalık Dengeli Çalışma Programı",
            "Deneme Net Artırma & Zaman Yönetimi",
            "Dikkat ve Çeldirici Tuzağı Azaltma",
            "Sınav Kaygısı & Motivasyon Rutini",
        ]
        stored_focus = st.session_state.get("dk_focus", focus_opts[0])
        if stored_focus not in focus_opts:
            stored_focus = focus_opts[0]
        coach_focus = st.selectbox(
            "Mentorluk Odak Alanı",
            focus_opts,
            index=focus_opts.index(stored_focus),
            key="dk_focus_sel",
        )
        st.session_state["dk_focus"] = coach_focus

    with r1_c:
        weak_topics = st.text_area(
            "Zorlandığın Dersler / Kazanımlar",
            value=st.session_state.get("dk_weak", "Sözel mantık ve muhakeme, üslü ifadeler, dikkat kaynaklı işlem hataları"),
            height=82,
            key="dk_weak_area",
        )
        st.session_state["dk_weak"] = weak_topics

        recent_score = st.text_input(
            "Mevcut Deneme / Süre Durumu",
            value=st.session_state.get("dk_status", "Orta seviye, süreyi yetiştirmede zorlanıyorum"),
            key="dk_status_inp",
        )
        st.session_state["dk_status"] = recent_score

    q_hrs = round(weekly_hours * 0.45, 1)
    r_hrs = round(weekly_hours * 0.30, 1)
    a_hrs = round(weekly_hours * 0.25, 1)
    st.markdown(
        f"""
        <div style="margin-top:0.65rem;padding-top:0.75rem;border-top:1px solid #E2E8F0;">
            <div style="font-size:0.75rem;font-weight:600;color:#475569;text-transform:uppercase;letter-spacing:0.04em;">
                Önerilen Haftalık Kapasite Dağılımı ({weekly_hours} Saat)
            </div>
            <div class="kp-dist-bar">
                <div style="width:45%;background:#2563EB;"></div>
                <div style="width:30%;background:#5A52E8;"></div>
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
    include_routine = st.checkbox("2028 Aşamalı Kamp Fazları (Kazanım → Branş → Genel Deneme) ve hafta sonu kontrol listesi dahil et", value=True)

    if st.button("Dijimentor 2028 Planı Oluştur", type="primary"):
        quota_ok, quota_msg = consume_quota("coach", 1)
        if not quota_ok:
            st.error(quota_msg)
        else:
            sys_coach = (
                "Sen 'Dijimentor' adlı, 2028 sınav dönemine hazırlanan öğrenciler için çalışan üst düzey, "
                "net, empatik ve veri odaklı bir dijital eğitim mentorusun. Yanıtını başlıklarla, "
                "madde işaretleriyle ve uygulanabilir zaman bloklarıyla kurumsal bir mentorluk "
                "raporu gibi düzenle. Şu bölümleri net oluştur:\n"
                "1. Teşhis ve 2028 Aşamalı Kamp Yol Haritası (Faz 1 Kazanım, Faz 2 Branş Hızı, Faz 3 Genel Deneme)\n"
                "2. Pazartesi–Pazar Haftalık Çalışma Planı (saat ve odak ders dağılımı)\n"
                "3. Çeldirici & Zaman Yönetimi Protokolü\n"
                "4. Haftalık Başarı KPI'ları ve Kontrol Listesi"
            )
            prompt_coach = (
                f"Hedef Sınav / Kademe: {target_exam}\nHedef Bandı: {target_goal}\n"
                f"Haftalık Çalışma Süresi: {weekly_hours} saat\n"
                f"Zorlanılan Dersler / Kazanımlar: {weak_topics}\n"
                f"Mentorluk Odak Alanı: {coach_focus}\n"
                f"Mevcut Deneme / Süre Durumu: {recent_score}\n"
                f"Aşamalı Kamp & Analiz Protokolü: {'Evet' if include_routine else 'Özet'}\n\n"
                "Lütfen öğrenciye doğrudan uygulanabilir, motive edici ve net bir Türkçe 2028 planı sun."
            )
            skel_coach = st.empty()
            render_loading_skeleton(skel_coach, lines=4)
            try:
                plan = call_gemini_text(prompt=prompt_coach, system_instruction=sys_coach)
                skel_coach.empty()
                st.session_state.dijimentor_plan = plan
                record_activity("Dijimentor (2028)", f"{target_exam} • {weekly_hours} saatlik plan üretildi")
                st.success("Dijimentor 2028 planı başarıyla oluşturuldu.")
            except Exception as exc:
                skel_coach.empty()
                refund_quota("coach", 1)
                st.error(f"Dijimentor planı oluşturulurken hata oluştu: {exc}")

    if st.session_state.dijimentor_plan:
        st.markdown("---")
        st.markdown("### Kişiselleştirilmiş Dijimentor 2028 Mentorluk Raporu")
        st.markdown(st.session_state.dijimentor_plan)

        meta_summary = (
            f"Hedef: {target_exam} | Hedef Bandı: {target_goal} | "
            f"Haftalık Süre: {weekly_hours} Saat | Odak: {coach_focus}"
        )
        docx_coach = build_coach_docx(meta_summary, st.session_state.dijimentor_plan)
        txt_report = (
            f"{BRAND_NAME} (2028)\n{meta_summary}\nTarih: {datetime.now().strftime('%d.%m.%Y')}\n"
            f"{'=' * 60}\n\n{st.session_state.dijimentor_plan}"
        )

        d_col1, d_col2, d_col3 = st.columns(3)
        with d_col1:
            st.download_button(
                label="Mentorluk Planını DOCX İndir",
                data=docx_coach,
                file_name="dijimentor_2028_calisma_plani.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with d_col2:
            st.download_button(
                label="Planı Metin (.txt) İndir",
                data=txt_report.encode("utf-8"),
                file_name="dijimentor_2028_calisma_plani.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with d_col3:
            if st.button("Zayıf Konulardan Soru Üret", use_container_width=True):
                first_w = weak_topics.split(",")[0].strip()
                st.session_state["factory_focus_topic"] = first_w
                st.session_state["fab_theme_inp"] = first_w
                navigate_to("Soru & Redaksiyon Fabrikası")

# ==============================================================================
# 5. VELİ TAKİP PANELİ (+ BU HAFTA NE DEĞİŞTİ KARTLARI)
# ==============================================================================
elif current_page == "Veli Takip Paneli":
    stats_now = st.session_state["parent_student_stats"]
    v1, v2, v3 = st.columns(3)
    v1.metric("Haftalık Çalışma Süresi", f"{int(st.session_state.get('dk_hours', 16))} Saat", "Düzenli tempo")
    v2.metric("Çözülen Soru Adedi", f"{int(stats_now['weekly_solved'])} Soru", "+18% geçen haftaya göre")
    v3.metric("Kazanım İvmesi & Uyum", f"%{int(stats_now['mentor_completion'])}", "+4 puan doğruluk ivmesi")

    st.markdown("#### Bu Hafta Ne Değişti?")
    w_c1, w_c2, w_c3 = st.columns(3)
    w_c1.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#2563EB;">GÜÇLÜ KAZANIM</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['strong_topic']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">İsabet oranı istikrarlı biçimde yükseldi.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    w_c2.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#D97706;">RİSKLİ KAZANIM</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['weak_topic']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">Hata Defteri ve telafi seti ile takip ediliyor.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    w_c3.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#059669;">SIRADAKİ DENEME</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['next_exam_date']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">{html.escape(str(stats_now['next_exam_name']))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("#### Kazanım Bazlı Görünüm")
    st.progress(0.84, text="Okuma ve Anlama — %84")
    st.progress(0.72, text="Sözel Mantık & Sayısal Muhakeme — %72")
    st.progress(0.90, text="Düzenli Tekrar ve Ödev Tamamlama — %90")

    g_col1, g_col2 = st.columns(2)
    with g_col1:
        st.markdown("#### Haftalık Doğruluk İvmesi (%)")
        st.area_chart({"1. Hafta": 71, "2. Hafta": 74, "3. Hafta": 77, "4. Hafta": 81}, color="#2563EB", height=190)
    with g_col2:
        st.markdown("#### Ders Bazlı Başarı Oranı (%)")
        st.bar_chart({"Türkçe": 84, "Matematik": 72, "Fen": 82, "Sosyal": 86, "İngilizce": 79}, color="#2563EB", height=190)

    render_parent_notification_center()

# ==============================================================================
# 6. ÖĞRENCİ RÖNTGENİ
# ==============================================================================
elif current_page == "Öğrenci Röntgeni":
    render_student_xray_panel()

# ==============================================================================
# 7. B2B LİSANS & KURUM (+ DENEME TAKVİMİ)
# ==============================================================================
elif current_page == "B2B Lisans & Kurum":
    tab_b2b_form, tab_b2b_zumre = st.tabs(["Kurumsal Lisans Teklif Stüdyosu (2028)", "Zümre Soru Dağıtımı & Kurumsal Rapor"])

    with tab_b2b_form:
        with st.form("b2b_inquiry_form"):
            c1, c2 = st.columns(2)
            with c1:
                org_name = st.text_input("Kurum / Ofis Adı", value="Kurşun Kalem Akademi Ofisi")
                city = st.text_input("İl / İlçe", value="İstanbul")
                teacher_seats = st.number_input("Öğretmen / Editör Koltuk Sayısı", min_value=1, max_value=500, value=15)
            with c2:
                student_seats = st.number_input("Öğrenci Hesap Sayısı", min_value=10, max_value=10000, value=250)
                contact_name = st.text_input("Yetkili Adı Soyadı")
                contact_email = st.text_input("Kurumsal E-posta")

            submitted = st.form_submit_button("2028 Kurumsal Lisans Teklifi Hazırla", type="primary")
            if submitted:
                if not org_name.strip():
                    st.warning("Lütfen kurum adını girin.")
                else:
                    st.session_state["last_b2b_summary"] = (
                        f"Dönem: 2028 | Kurum: {org_name} ({city}) | Yetkili: {contact_name or '-'} | "
                        f"E-posta: {contact_email or '-'} | Koltuk: {teacher_seats} Öğretmen, "
                        f"{student_seats} Öğrenci"
                    )
                    st.session_state["b2b_org_cache"] = {
                        "org": org_name,
                        "teachers": int(teacher_seats),
                        "students": int(student_seats),
                    }
                    record_activity("B2B Lisans (2028)", f"{org_name} teklif taslağı hazırlandı")
                    st.success(
                        f"**{org_name}** ({city}) için 2028 dönemi **{teacher_seats} öğretmen** ve "
                        f"**{student_seats} öğrenci** kapasiteli kurumsal lisans talebi kaydedildi."
                    )

        if st.session_state.get("last_b2b_summary"):
            cache_for_doc = st.session_state.get("b2b_org_cache", {"org": org_name, "students": int(student_seats)})
            b2b_tier = get_tier_by_id("b2b_kurs" if cache_for_doc.get("students", 250) <= 300 else "b2b_kampus")
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
                title=f"{cache_for_doc.get('org', 'Kurumsal')} 2028 Lisans Teklif Özeti",
                summary=b2b_summary,
                billing_cycle="Yıllık (%20 İndirim)",
                installment=6,
                total_amount=round(b2b_tier["monthly_price"] * 12 * 0.80),
                extra_notes=st.session_state["last_b2b_summary"],
            )
            st.download_button(
                label="Kurumsal Lisans Teklifini DOCX İndir",
                data=b2b_docx,
                file_name="kursun_kalem_ofis_2028_kurumsal_teklif.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    with tab_b2b_zumre:
        org_cache = st.session_state.get("b2b_org_cache", {"org": "Kurşun Kalem Akademi Ofisi", "teachers": 15, "students": 250})
        recommended_tier = get_tier_by_id("b2b_kurs" if org_cache["students"] <= 300 else "b2b_kampus")
        total_q = recommended_tier["monthly_quota"]

        bk1, bk2, bk3, bk4 = st.columns(4)
        bk1.metric("Kurum Profili", org_cache["org"])
        bk2.metric("Aylık Soru Havuzu", f"{total_q:,}".replace(",", "."))
        bk3.metric("Öğretmen Başına Kota", f"{total_q // max(1, org_cache['teachers']):,} Soru".replace(",", "."))
        bk4.metric("Öğrenci Başına Test", f"{total_q // max(1, org_cache['students'])} Soru / Ay")

        st.markdown("#### Branş Zümreleri Arası Soru Kotası Dağılımı (2028)")
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
            st.bar_chart(chart_map, color="#2563EB", height=225)

        st.markdown("#### Kurum İçi Deneme & Kazanım Tarama Planı")
        st.dataframe(
            [
                {"Hafta": "1. Hafta", "Kademe": "8. Sınıf (LGS)", "Odak": "Sözel-Sayısal Muhakeme Tarama", "Durum": "Planlandı"},
                {"Hafta": "2. Hafta", "Kademe": "5–7. Sınıf", "Odak": "Maarif Modeli Ara Kazanım Testi", "Durum": "Planlandı"},
                {"Hafta": "3. Hafta", "Kademe": "9–12. Sınıf (YKS)", "Odak": "TYT-AYT Hız & Çeldirici Denemesi", "Durum": "Taslak"},
            ],
            use_container_width=True,
            hide_index=True,
        )

# ==============================================================================
# 8. PAKETLER & ABONELİK VİTRİNİ (2028)
# ==============================================================================
elif current_page == "Paketler & Abonelik Vitrini":
    total_pkg_count = len(SUBSCRIPTION_TIERS)

    st.markdown(
        """
        <div class="kp-value-strip">
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> Türkiye Yüzyılı Maarif Modeli 2028 Soru Havuzu</div>
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> Dijimentor Haftalık Akıllı Çalışma Planı</div>
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> iyzico Güvencesiyle 12 Taksite Varan Ödeme</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
                is_featured = tier["badge"] in ("En Çok Tercih Edilen", "VIP", "Kampüs")
                card_cls = "kp-pkg-card kp-pkg-featured" if is_featured else "kp-pkg-card"
                feats_li = "".join([f"<li>• {html.escape(feat)}</li>" for feat in tier["features"]])
                st.markdown(
                    f"""
                    <div class="{card_cls}">
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
    st.markdown(f"#### {total_pkg_count} Paket Karşılaştırma Matrisi (2028)")
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
    )import base64
import html
import io
import json
import os
import re
import time
import uuid
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
# SAYFA YAPILANDIRMASI VE KURUMSAL SABİTLER (2028 KURGUSU)
# ==============================================================================
st.set_page_config(
    page_title="KURŞUN KALEM DİJİTAL OFİS • 2028",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "KURŞUN KALEM DİJİTAL OFİS"
BRAND_SHORT = "Kurşun Kalem"
BRAND_TAGLINE = "2028 Hedefli Soru & Redaksiyon Fabrikası • Dijimentor • Veli & B2B Dijital Ofis"
DEFAULT_MODEL = "gemini-2.5-flash"
FALLBACK_MODEL = "gemini-3.5-flash-lite"
STATE_FILE_PATH = Path(".kursun_kalem_ofis_2028_state.json")
STATE_BACKUP_PATH = Path(".kursun_kalem_ofis_2028_state.bak.json")
EXTERNAL_CSS_PATH = Path("assets/style.css")

GRADE_LEVELS_2028: List[str] = [
    "5. Sınıf",
    "6. Sınıf",
    "7. Sınıf",
    "8. Sınıf (LGS)",
    "9. Sınıf",
    "10. Sınıf",
    "11. Sınıf (YKS)",
    "12. Sınıf (YKS)",
]

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
  <polygon points="256,82 362,130 256,178 150,130" fill="url(#navyGrad)" stroke="url(#goldGrad)" stroke-width="5"/>
  <polygon points="190,148 256,178 322,148 322,188 256,216 190,188" fill="url(#navyGrad)"/>
  <text x="256" y="314" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="44" font-weight="800" fill="#0b1836">DİJİTAL OFİS</text>
  <text x="256" y="351" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="22" font-weight="700" fill="#0f172a" letter-spacing="4">KURŞUN KALEM</text>
  <g transform="translate(432, 342) rotate(24)">
    <rect x="-12" y="-110" width="24" height="96" rx="3" fill="url(#pencilBody)" stroke="#1e293b" stroke-width="2"/>
    <polygon points="-12,-14 12,-14 0,24" fill="#fde68a" stroke="#1e293b" stroke-width="2"/>
    <polygon points="-3.5,12 3.5,12 0,24" fill="#0f172a"/>
  </g>
</svg>"""

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
  <text x="130" y="184" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="800" fill="#0f172a">OFİS</text>
  <text x="130" y="196" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="7.5" font-weight="700" fill="#475569">ÇANTASI</text>
  <rect x="72" y="242" width="116" height="18" rx="9" fill="#090d16"/>
</svg>"""

HERO_SHOWCASE_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 440 330" width="440" height="330">
  <defs>
    <linearGradient id="bgCardGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#EEF2FF"/>
      <stop offset="100%" stop-color="#E0E7FF"/>
    </linearGradient>
    <linearGradient id="navyBanner" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#1E2952"/>
      <stop offset="100%" stop-color="#2563EB"/>
    </linearGradient>
    <linearGradient id="goldBadgeGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FACC15"/>
      <stop offset="100%" stop-color="#D97706"/>
    </linearGradient>
    <linearGradient id="pencilWoodGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FDE047"/>
      <stop offset="100%" stop-color="#EAB308"/>
    </linearGradient>
  </defs>

  <rect x="22" y="18" width="396" height="292" rx="24" fill="url(#bgCardGrad)" stroke="#C7D2FE" stroke-width="1.5"/>

  <rect x="44" y="42" width="214" height="244" rx="18" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="1.5"/>
  <rect x="44" y="42" width="214" height="58" rx="18" fill="url(#navyBanner)"/>
  <rect x="44" y="76" width="214" height="24" fill="url(#navyBanner)"/>
  <text x="62" y="68" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="800" fill="#FACC15" letter-spacing="1.2">KURŞUN KALEM • 2028</text>
  <text x="62" y="87" font-family="Inter, Arial, sans-serif" font-size="13" font-weight="800" fill="#FFFFFF">Ölçme &amp; Redaksiyon Ofisi</text>

  <rect x="60" y="114" width="182" height="42" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <circle cx="79" cy="135" r="8" fill="#2563EB"/>
  <text x="94" y="132" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#1E2952">Maarif Kazanım Uyumu</text>
  <text x="94" y="146" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">LGS &amp; YKS 2028 Müfredatı</text>

  <rect x="60" y="164" width="182" height="42" rx="10" fill="#F8FAFC" stroke="#E2E8F0"/>
  <circle cx="79" cy="185" r="8" fill="#4F46E5"/>
  <text x="94" y="182" font-family="Inter, Arial, sans-serif" font-size="10.5" font-weight="700" fill="#1E2952">Çeldirici &amp; Kök Denetimi</text>
  <text x="94" y="196" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">Çift Katmanlı Editör Kontrolü</text>

  <rect x="60" y="214" width="182" height="52" rx="10" fill="#EFF6FF" stroke="#BFDBFE"/>
  <text x="74" y="233" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="700" fill="#1D4ED8">HEDEF NET &amp; İSABET İVMESİ</text>
  <rect x="74" y="243" width="152" height="8" rx="4" fill="#DBEAFE"/>
  <rect x="74" y="243" width="124" height="8" rx="4" fill="#2563EB"/>

  <rect x="276" y="42" width="120" height="114" rx="18" fill="#FFFFFF" stroke="#E2E8F0" stroke-width="1.5"/>
  <circle cx="336" cy="82" r="22" fill="url(#goldBadgeGrad)"/>
  <polygon points="336,68 340,77 350,78 342,85 345,95 336,90 327,95 330,85 322,78 332,77" fill="#1E2952"/>
  <text x="336" y="121" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="11" font-weight="800" fill="#1E2952">2028 HEDEF</text>
  <text x="336" y="136" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="9.5" font-weight="600" fill="#64748B">Dijimentor &amp; Röntgen</text>

  <rect x="276" y="170" width="120" height="116" rx="18" fill="#1E2952"/>
  <g transform="translate(336, 214) rotate(-32)">
    <rect x="-8" y="-34" width="16" height="52" rx="2" fill="url(#pencilWoodGrad)"/>
    <rect x="-8" y="-38" width="16" height="6" rx="1.5" fill="#F43F5E"/>
    <polygon points="-8,18 8,18 0,34" fill="#FDE68A"/>
    <polygon points="-2.5,28 2.5,28 0,34" fill="#0F172A"/>
  </g>
  <text x="336" y="266" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="10" font-weight="800" fill="#FACC15" letter-spacing="0.8">KURŞUN KALEM</text>
  <text x="336" y="279" text-anchor="middle" font-family="Inter, Arial, sans-serif" font-size="8.5" font-weight="600" fill="#E2E8F0">Dijital Ofis Vitrini</text>
</svg>"""


def svg_to_data_uri(svg_str: str) -> str:
    encoded = base64.b64encode(svg_str.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{encoded}"


def get_embedded_logo_png_bytes() -> Optional[bytes]:
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
        draw.text((160, 182), "DİJİTAL OFİS", fill=(11, 24, 54, 255), anchor="mm")
        out = io.BytesIO()
        img.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return None


LGS_2028_TARGET = datetime(2028, 6, 11, 9, 30, 0)
YKS_2028_TARGET = datetime(2028, 6, 17, 10, 15, 0)

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
        "Genel Bakış, 2028 Kademe Vitrini & Kişisel Akış",
        "Kademe seçimi, eksik kazanım öncelik sırası, 2028 LGS ve YKS geri sayımı ile dijital ofis özeti.",
    ),
    "Soru & Redaksiyon Fabrikası": (
        "Maarif Modeli Soru & Kalite Stüdyosu (2028)",
        "5–12. sınıf müfredatı, kademeli zorluk serisi, çeldirici denetimi, soru havuzu ve DOCX çıktı.",
    ),
    "Öğrenci Portalı & Çözüm": (
        "İnteraktif Mini Deneme, Hata Defteri & Çözüm Rehberi",
        "Süre takipli mini deneme çözün, netinizi görün, yanlışları Hata Defteri'ne alın ve adımlı çözüm inceleyin.",
    ),
    "Dijimentor": (
        "Akıllı Sınav, Faz Planı & Performans Mentorluğu (2028)",
        "2028 sınav takvimine göre aşamalı kamp ve haftalık aksiyon planı üretin.",
    ),
    "Veli Takip Paneli": (
        "Veli Karne Özeti & Bilgilendirme Merkezi",
        "Haftalık değişim kartları, kazanım ivmesi ve hazır veli bilgilendirme mesajları.",
    ),
    "Öğrenci Röntgeni": (
        "Öğrenci Röntgeni & Akademik Check-Up",
        "Kazanım bazlı röntgen taraması, hata kök neden analizi ve modüller arası tek tık aktarım.",
    ),
    "B2B Lisans & Kurum": (
        "Kurumsal Lisans, Zümre & Deneme Planlayıcı",
        "Koltuk/kota planlaması, zümre dağılımı, kurum içi deneme takvimi ve DOCX teklif çıktısı.",
    ),
    "Paketler & Abonelik Vitrini": (
        "Abonelik Paketleri & Ofis Çantası",
        "Öğrenci, 2028 YKS, Veli ve Kurumsal paketleri çantanıza ekleyin, kotaları yönetin ve iyzico planını inceleyin.",
    ),
}

HERO_PILL_CHIPS: List[Dict[str, str]] = [
    {"label": "YKS 2028", "grade": "12. Sınıf (YKS)", "exam": "2028 TYT - AYT", "page": "Dijimentor"},
    {"label": "LGS 2028", "grade": "8. Sınıf (LGS)", "exam": "2028 LGS", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "11. Sınıf", "grade": "11. Sınıf (YKS)", "exam": "2028 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "10. Sınıf", "grade": "10. Sınıf", "exam": "2028 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "9. Sınıf", "grade": "9. Sınıf", "exam": "2028 TYT - AYT", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "7. Sınıf", "grade": "7. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "6. Sınıf", "grade": "6. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "5. Sınıf", "grade": "5. Sınıf", "exam": "5-7. Sınıf Akademik Gelişim", "page": "Soru & Redaksiyon Fabrikası"},
    {"label": "Veli & Kurum", "grade": "8. Sınıf (LGS)", "exam": "2028 LGS", "page": "Veli Takip Paneli"},
    {"label": "Zümre & B2B", "grade": "8. Sınıf (LGS)", "exam": "2028 LGS", "page": "B2B Lisans & Kurum"},
]

KADEME_VITRIN_KARTLARI: List[Dict[str, str]] = [
    {
        "tag": "5–7. SINIF • 2028",
        "cta": "5–7. Sınıf • Aç",
        "title": "Ortaokul & Maarif Temeli",
        "desc": "Kazanım odaklı sorular, okuma hızı ve düzenli çalışma alışkanlığı.",
        "accent": "#2563EB",
        "grade": "7. Sınıf",
        "exam": "5-7. Sınıf Akademik Gelişim",
        "target_page": "Soru & Redaksiyon Fabrikası",
    },
    {
        "tag": "2028 LGS",
        "cta": "2028 LGS • Aç",
        "title": "8. Sınıf LGS Hazırlık",
        "desc": "Sözel-sayısal muhakeme, çeldirici analizi ve net yükselten deneme akışı.",
        "accent": "#4F46E5",
        "grade": "8. Sınıf (LGS)",
        "exam": "2028 LGS",
        "target_page": "Soru & Redaksiyon Fabrikası",
    },
    {
        "tag": "2028 YKS",
        "cta": "2028 YKS • Aç",
        "title": "TYT & AYT Lise Kampı",
        "desc": "Süre yönetimi, paragraf-analitik denge ve derece hedefleyen mentorluk.",
        "accent": "#7C3AED",
        "grade": "12. Sınıf (YKS)",
        "exam": "2028 TYT - AYT",
        "target_page": "Dijimentor",
    },
    {
        "tag": "VELİ & KURUM • 2028",
        "cta": "Veli & Kurum • Aç",
        "title": "Veli Takip & B2B Ofis",
        "desc": "Haftalık çalışma-soru özeti, veli mesajları ve zümre kota yönetimi.",
        "accent": "#0284C7",
        "grade": "8. Sınıf (LGS)",
        "exam": "2028 LGS",
        "target_page": "Veli Takip Paneli",
    },
]

# Ana sayfadaki sınıf seçici: sekme -> o sekmedeki seçenekler (HERO_PILL_CHIPS etiketleri)
GRADE_GROUPS: Dict[str, List[str]] = {
    "Ortaokul": ["5. Sınıf", "6. Sınıf", "7. Sınıf", "LGS 2028"],
    "Lise": ["9. Sınıf", "10. Sınıf", "11. Sınıf", "YKS 2028"],
    "Veli & Kurum": ["Veli & Kurum", "Zümre & B2B"],
}
CHIP_BY_LABEL: Dict[str, Dict[str, str]] = {c["label"]: c for c in HERO_PILL_CHIPS}

# Sol menü: (başlık, [(sayfa adı, menüde görünen kısa ad, ikon), ...])
NAV_GROUPS: List[Tuple[Optional[str], List[Tuple[str, str, str]]]] = [
    (None, [("Ana Sayfa", "Ana Sayfa", ":material/home:")]),
    ("Modüller", [
        ("Soru & Redaksiyon Fabrikası", "Soru & Redaksiyon", ":material/edit_note:"),
        ("Öğrenci Portalı & Çözüm", "Öğrenci Portalı", ":material/quiz:"),
        ("Dijimentor", "Dijimentor", ":material/event_note:"),
        ("Öğrenci Röntgeni", "Öğrenci Röntgeni", ":material/monitor_heart:"),
    ]),
    ("Veli & Kurum", [
        ("Veli Takip Paneli", "Veli Takip Paneli", ":material/family_restroom:"),
        ("B2B Lisans & Kurum", "B2B Lisans & Kurum", ":material/apartment:"),
    ]),
    ("Hesap", [
        ("Paketler & Abonelik Vitrini", "Paketler", ":material/shopping_bag:"),
    ]),
]

# Ana sayfadaki modül kartları: (kısaltma, başlık, açıklama, hedef sayfa, arka plan, yazı rengi)
MODULE_CARDS: List[Tuple[str, str, str, str, str, str]] = [
    ("SR", "Soru & Redaksiyon", "Kazanım etiketli soru üretimi ve çift katmanlı editör kontrolü.", "Soru & Redaksiyon Fabrikası", "#EFF6FF", "#1D4ED8"),
    ("DM", "Dijimentor", "Seviyene göre her hafta yenilenen kişisel çalışma planı.", "Dijimentor", "#F0FDF4", "#15803D"),
    ("MD", "Mini Deneme", "Süre takipli denemeler ve otomatik hata defteri.", "Öğrenci Portalı & Çözüm", "#FEF3C7", "#B45309"),
    ("ÖR", "Öğrenci Röntgeni", "Konu konu güçlü ve zayıf yönlerin net haritası.", "Öğrenci Röntgeni", "#F5F3FF", "#6D28D9"),
    ("VT", "Veli Takip", "Haftalık özet, ilerleme ve veli mesajları tek yerde.", "Veli Takip Paneli", "#FFF1F2", "#BE123C"),
]

# ==============================================================================
# 5–12. SINIF TAM MÜFREDAT MOTORU (MAARİF MODELİ, LGS & YKS)
# ==============================================================================
_MIDDLE_COMMON: Dict[str, Dict[str, List[str]]] = {
    "Türkçe": {
        "Okuma ve Anlama": [
            "Bağlamdan yararlanarak sözcük ve söz öbeklerinin anlamını tahmin etme",
            "Metindeki ana düşünce, yardımcı düşünceler ve başlık ilişkisini kurma",
            "Sözel mantık, tablo/grafik yorumlama ve muhakeme gerektiren metinleri çözümleme",
            "Metinler arası karşılaştırma yapma ve çıkarımda bulunma",
        ],
        "Dil Bilgisi ve Yazım": [
            "Sözcük türleri, fiilimsiler ve cümle ögeleri işlevini belirleme",
            "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            "Cümlede anlam ilişkileri ve anlatım akıcılığını denetleme",
        ],
    },
    "Matematik": {
        "Sayılar ve İşlemler": [
            "Doğal sayılar, tam sayılar, kesirler ve ondalık gösterimlerle çok adımlı işlemler",
            "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
            "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve cebirsel ifadeler",
        ],
        "Geometri ve Veri Analizi": [
            "Doğrular, açılar, çokgenler, üçgenler ve geometrik cisimler",
            "Veri toplama, grafik yorumlama, olasılık ve merkezi eğilim ölçüleri",
        ],
    },
    "Fen Bilimleri": {
        "Fiziksel Süreçler ve Kuvvet": [
            "Kuvvet, hareket, basınç, iş-enerji ve basit makineler",
            "Mevsimlerin oluşumu, ışık, ses ve elektrik enerjisi uygulamaları",
        ],
        "Madde, Doğa ve Canlılar": [
            "Maddenin yapısı, karışımlar, kimyasal tepkimeler ve asit-baz kavramları",
            "Hücre, sistemler, DNA-genetik kod, kalıtım ve ekoloji/sürdürülebilirlik",
        ],
    },
    "Sosyal Bilgiler / İnkılap Tarihi": {
        "Tarih ve Kültür": [
            "Birey, toplum, kültür ve miras; Millî Mücadele ve Atatürk İlkeleri",
            "Cumhuriyet kazanımları, çağdaş gelişmeler ve tarihsel kanıt okuma",
        ],
        "İnsan, Yer ve Çevre": [
            "Harita okuryazarlığı, nüfus, ekonomik faaliyetler ve çevre bilinci",
            "Haklar, sorumluluklar, demokratik süreçler ve küresel bağlantılar",
        ],
    },
    "İngilizce": {
        "Reading & Comprehension": [
            "Contextual vocabulary, inference and main idea in short texts",
            "Everyday communication, invitations, preferences and plans",
        ],
        "Language & Reasoning": [
            "Tenses, modals, comparatives and cohesive connectors in context",
            "Table/chart interpretation and situational dialogue completion",
        ],
    },
}

_HIGH_COMMON: Dict[str, Dict[str, List[str]]] = {
    "Türk Dili ve Edebiyatı": {
        "Anlama ve Yorumlama": [
            "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi",
            "Dönem/zihniyet ilişkisi, edebi sanatlar ve metinler arası etkileşim",
        ],
        "Dil Bilgisi ve Yazım": [
            "Sözcükte ve cümlede anlam, paragraf kurgusu ve çeldirici analizi",
            "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu",
        ],
    },
    "Matematik": {
        "Temel Kavramlar ve Cebir": [
            "Temel kavramlar, bölünebilme, rasyonel/üslü/köklü ifadeler ve mutlak değer",
            "Denklemler, eşitsizlikler, oran-orantı, problemler ve fonksiyonlar",
        ],
        "Analitik, Geometri ve İleri Konular": [
            "Üçgenler, çokgenler, çember-daire, katı cisimler ve analitik geometri",
            "Polinomlar, logaritma, diziler, trigonometri, limit, türev ve integral",
        ],
    },
    "Fizik": {
        "Mekanik ve Hareket": [
            "Fizik bilimine giriş, hareket, Newton hareket yasaları, iş-güç-enerji",
            "İtme-momentum, tork-denge, basit makineler, kaldırma kuvveti ve basınç",
        ],
        "Elektrik, Manyetizma ve Dalgalar": [
            "Elektrostatik, elektrik akımı, manyetizma, elektromanyetik indüksiyon",
            "Dalgalar, optik, çembersel hareket, harmonik hareket ve modern fizik",
        ],
    },
    "Kimya": {
        "Madde ve Tepkimeler": [
            "Kimya bilimi, atom ve periyodik sistem, kimyasal türler arası etkileşimler",
            "Maddenin halleri, mol kavramı, kimyasal hesaplamalar ve tepkime hızı/denge",
        ],
        "Çözeltiler ve Enerji": [
            "Karışımlar, asitler-bazlar-tuzlar, kimya ve enerji, elektrokimya ve organik kimya",
        ],
    },
    "Biyoloji": {
        "Canlıların Temelleri ve Sistemler": [
            "Yaşam bilimi biyoloji, hücre zarından madde geçişi, canlıların sınıflandırılması",
            "Hücre bölünmeleri, kalıtım, ekosistem ekolojisi ve insan fizyolojisi sistemleri",
        ],
        "Genden Proteine ve Enerji": [
            "Nükleik asitler, protein sentezi, fotosentez ve hücresel solunum",
        ],
    },
    "Tarih": {
        "Tarihsel Süreç ve İnkılap": [
            "Tarih ve zaman, insanlığın ilk dönemleri, İslam tarihi ve Türk-İslam devletleri",
            "Osmanlı kurumları ve diplomasi, Kurtuluş Savaşı ve Atatürk İlke ve İnkılapları",
        ],
    },
    "Coğrafya": {
        "Doğal ve Beşeri Sistemler": [
            "Doğa ve insan, coğrafi koordinat ve iklim bilgisi, iç-dış kuvvetler",
            "Nüfus, yerleşme, ekonomik faaliyetler, bölgeler ve küresel çevre sorunları",
        ],
    },
}

FULL_CURRICULUM_2028: Dict[str, Dict[str, Dict[str, List[str]]]] = {
    "5. Sınıf": _MIDDLE_COMMON,
    "6. Sınıf": _MIDDLE_COMMON,
    "7. Sınıf": _MIDDLE_COMMON,
    "8. Sınıf (LGS)": {
        "Türkçe": {
            "Okuma ve Muhakeme": [
                "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
                "Metindeki ana düşünce, çıkarım ve başlık/görsel ilişkisini kurma",
                "Metinler arası karşılaştırma yapma ve yargı/kanıt türlerini ayırt etme",
            ],
            "Dil Bilgisi ve Yazım": [
                "Fiilimsiler (isim-fiil, sıfat-fiil, zarf-fiil) ve cümle ögeleri",
                "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
            ],
        },
        "Matematik": {
            "Sayılar ve Cebir": [
                "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
                "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler",
                "Cebirsel ifadeler, özdeşlikler ve birinci dereceden denklemler",
            ],
            "Geometri, Olasılık ve Veri": [
                "Üçgenler, kenarortay/açıortay/yükseklik, eşlik-benzerlik ve dönüşüm geometrisi",
                "Olasılık ve istatistiksel veri analizi / grafik dönüştürme",
            ],
        },
        "Fen Bilimleri": {
            "2028 LGS Fen Kazanımları": [
                "Mevsimlerin oluşumu ve iklim/hava hareketleri",
                "DNA, genetik kod, mutasyon, modifikasyon, adaptasyon ve biyoteknoloji",
                "Basınç (katı, sıvı, gaz) ve madde-endüstri uygulamaları",
                "Basit makineler, enerji dönüşümleri, elektrik yükleri ve elektrik enerjisi",
            ],
        },
        "T.C. İnkılap Tarihi ve Atatürkçülük": {
            "Millî Mücadele ve Çağdaş Türkiye": [
                "Bir Kahraman Doğuyor ve Millî Uyanış: Bağımsızlık Yolunda Atılan Adımlar",
                "Millî Bir Destan: Ya İstiklal Ya Ölüm ve Atatürkçülük/Çağdaşlaşan Türkiye",
                "Demokratikleşme çabaları ve Atatürk'ten Sonraki Dış Politika",
            ],
        },
        "İngilizce": {
            "2028 LGS İngilizce Temaları": [
                "Friendship, Teen Life, In the Kitchen, On the Phone, The Internet",
                "Tourism, Chores, Science, Natural Forces & Reading Comprehension",
            ],
        },
    },
    "9. Sınıf": _HIGH_COMMON,
    "10. Sınıf": _HIGH_COMMON,
    "11. Sınıf (YKS)": _HIGH_COMMON,
    "12. Sınıf (YKS)": _HIGH_COMMON,
}


def get_branches_for_grade_2028(grade: str) -> List[str]:
    g_map = FULL_CURRICULUM_2028.get(grade)
    if g_map:
        return list(g_map.keys())
    return list(_MIDDLE_COMMON.keys())


def get_areas_for_grade_branch_2028(grade: str, branch: str) -> List[str]:
    g_map = FULL_CURRICULUM_2028.get(grade, {})
    if branch in g_map:
        return list(g_map[branch].keys())
    for other_g in FULL_CURRICULUM_2028.values():
        if branch in other_g:
            return list(other_g[branch].keys())
    return ["Genel Kazanımlar"]


def get_outcomes_for_2028(grade: str, branch: str, area: str) -> List[str]:
    g_map = FULL_CURRICULUM_2028.get(grade, {})
    if branch in g_map and area in g_map[branch]:
        return list(g_map[branch][area])
    for other_g in FULL_CURRICULUM_2028.values():
        if branch in other_g and area in other_g[branch]:
            return list(other_g[branch][area])
    return ["Kazanım odaklı beceri ve muhakeme sorusu üretimi"]


# ==============================================================================
# DİĞER SABİTLER (RÖNTGEN, DİJİMENTOR, ABONELİK, VELİ MESAJLARI)
# ==============================================================================
OGRENCI_RONTGEN_PROFILLERI: Dict[str, Dict[str, Any]] = {
    "Elif Öztürk (8-A | 2028 LGS)": {
        "basari_endeksi": "%76.4",
        "basari_delta": "+4.2% (Son 2 hafta)",
        "kritik_konu": "3 Kazanım",
        "kritik_delta": "-2 İyileşme",
        "hiz": "1.35 dk/soru",
        "hiz_delta": "İdeal Tempo",
        "hedef_puan": "454 / 480",
        "hedef_delta": "26 Puan Fark",
        "kazanimlar": [
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme", "Çözülen": 58, "Başarı (%)": 46, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "Metindeki ana düşünce, çıkarım ve başlık/görsel ilişkisini kurma", "Çözülen": 135, "Başarı (%)": 86, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
            {"Ders / Alan": "Matematik", "Kazanım / Konu": "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri", "Çözülen": 74, "Başarı (%)": 52, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Fen Bilimleri", "Kazanım / Konu": "Basınç (katı, sıvı, gaz) ve madde-endüstri uygulamaları", "Çözülen": 62, "Başarı (%)": 74, "Teşhis Durumu": "🟡 Geliştirilmeli"},
            {"Ders / Alan": "T.C. İnkılap Tarihi ve Atatürkçülük", "Kazanım / Konu": "Millî Mücadele ve Atatürk İlkeleri", "Çözülen": 48, "Başarı (%)": 88, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 36,
            "Uzun Metin / Odak Kaybı": 30,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 22,
            "Süre Yetişmemesi": 12,
        },
        "recete_ozet": (
            "Sözel Mantık & Muhakeme (%46) ve EBOB-EKOK modellemeleri (%52) kazanımlarında tablo/şema kurma ve "
            "çok adımlı yönerge takibi zayıf. Yanlışların %22'si soru kökündeki olumsuz ifadelerin hızlı okunmasından kaynaklanıyor."
        ),
        "recete_aksiyon": "Soru & Redaksiyon Fabrikası üzerinden orta-zor seviyede Sözel Mantık ve Matematik telafi soruları üretilip atanmaya hazır.",
        "hedef_konu": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
        "hedef_brans": "Türkçe",
        "hedef_sinif": "8. Sınıf (LGS)",
    },
    "Zeynep Kaya (8-B | 2028 LGS)": {
        "basari_endeksi": "%83.8",
        "basari_delta": "+5.6% (Son 2 hafta)",
        "kritik_konu": "2 Kazanım",
        "kritik_delta": "-1 İyileşme",
        "hiz": "1.18 dk/soru",
        "hiz_delta": "Hızlı ve Dengeli",
        "hedef_puan": "468 / 480",
        "hedef_delta": "12 Puan Fark",
        "kazanimlar": [
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "Metinler arası karşılaştırma yapma ve yargı/kanıt türlerini ayırt etme", "Çözülen": 64, "Başarı (%)": 81, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
            {"Ders / Alan": "Matematik", "Kazanım / Konu": "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler", "Çözülen": 72, "Başarı (%)": 55, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Fen Bilimleri", "Kazanım / Konu": "Mevsimlerin oluşumu ve iklim/hava hareketleri", "Çözülen": 54, "Başarı (%)": 89, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
            {"Ders / Alan": "Türkçe", "Kazanım / Konu": "Fiilimsiler (isim-fiil, sıfat-fiil, zarf-fiil) ve cümle ögeleri", "Çözülen": 86, "Başarı (%)": 71, "Teşhis Durumu": "🟡 Geliştirilmeli"},
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 27,
            "Uzun Metin / Odak Kaybı": 21,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 35,
            "Süre Yetişmemesi": 17,
        },
        "recete_ozet": (
            "Üslü ve kareköklü ifadeler (%55) alanında yaklaşık değer ve yeni nesil kurgu adımlarında sapma görülüyor. "
            "Hataların %35'i soru kökündeki sınırlandırıcı ifadelerin gözden kaçmasından kaynaklanıyor."
        ),
        "recete_aksiyon": "Olumsuz soru köklü Kareköklü İfade ve dikkat odaklı muhakeme seti oluşturularak Öğrenci Portalı'na gönderilmeye hazır.",
        "hedef_konu": "Üslü ve kareköklü ifadelerde akıl yürütme, tahmin ve işlemler",
        "hedef_brans": "Matematik",
        "hedef_sinif": "8. Sınıf (LGS)",
    },
    "Emir Demir (12-A | 2028 YKS)": {
        "basari_endeksi": "%71.2",
        "basari_delta": "+3.8% (Son 2 hafta)",
        "kritik_konu": "4 Kazanım",
        "kritik_delta": "-1 İyileşme",
        "hiz": "1.52 dk/soru",
        "hiz_delta": "Süre Baskısı Var",
        "hedef_puan": "442 / 500",
        "hedef_delta": "58 Puan Fark",
        "kazanimlar": [
            {"Ders / Alan": "Türk Dili ve Edebiyatı", "Kazanım / Konu": "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi", "Çözülen": 145, "Başarı (%)": 66, "Teşhis Durumu": "🟡 Geliştirilmeli"},
            {"Ders / Alan": "Türk Dili ve Edebiyatı", "Kazanım / Konu": "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu", "Çözülen": 78, "Başarı (%)": 49, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Matematik", "Kazanım / Konu": "Polinomlar, logaritma, diziler, trigonometri, limit, türev ve integral", "Çözülen": 88, "Başarı (%)": 54, "Teşhis Durumu": "🔴 Kritik Müdahale"},
            {"Ders / Alan": "Fizik", "Kazanım / Konu": "İtme-momentum, tork-denge, basit makineler, kaldırma kuvveti ve basınç", "Çözülen": 56, "Başarı (%)": 84, "Teşhis Durumu": "🟢 Tam Hakimiyet"},
        ],
        "hatalar": {
            "Kavram / Bilgi Eksikliği": 34,
            "Uzun Metin / Odak Kaybı": 24,
            "Soru Kökü ('Değildir/Ulaşılamaz')": 16,
            "Süre Yetişmemesi": 26,
        },
        "recete_ozet": (
            "Dil Bilgisi (%49) ve İleri Matematik (%54) alanlarında kural transferi ve süre yönetimi kritik öncelikte. "
            "Süre yetişmemesi (%26) denemenin ikinci yarısında dikkat düşüşüyle birleşiyor."
        ),
        "recete_aksiyon": "Zamanlı Dil Bilgisi + Analitik Muhakeme sorularından oluşan hız ve doğruluk telafi paketi hazır.",
        "hedef_konu": "Ses bilgisi, yazım kuralları, noktalama işaretleri ve anlatım bozukluğu",
        "hedef_brans": "Türk Dili ve Edebiyatı",
        "hedef_sinif": "12. Sınıf (YKS)",
    },
}

ONCELIKLI_EKSIK_SIRASI: List[Dict[str, str]] = [
    {
        "priority": "1. ÖNCELİK • %46 İSABET",
        "pct": 46,
        "topic": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "action": "Tablo/şema kurma ve telafi sorusu",
    },
    {
        "priority": "2. ÖNCELİK • %52 İSABET",
        "pct": 52,
        "topic": "Çarpanlar ve katlar, EBOB-EKOK problem modellemeleri",
        "branch": "Matematik",
        "grade": "8. Sınıf (LGS)",
        "action": "Yeni nesil kurgu & çeldirici analizi",
    },
    {
        "priority": "3. ÖNCELİK • DİKKAT",
        "topic": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "action": "Olumsuz soru kökü hız antrenmanı",
    },
]

DIJIMENTOR_PRESETS: Dict[str, Dict[str, Any]] = {
    "2028 LGS Derece & Net İvmesi": {
        "exam": "2028 LGS",
        "hours": 20,
        "goal": "Fen liseleri ve nitelikli Anadolu liseleri bandı (470+ puan ve 82+ net hedefi)",
        "focus": "Deneme Net Artırma & Zaman Yönetimi",
        "weak": "Sözel mantık ve muhakeme metinleri, cebirsel ifadeler ve çarpanlar-katlar yeni nesil problemleri, süre baskısı altında çeldirici ayıklama",
        "status": "Orta-üst seviye; Türkçe ve Matematikte son 15 soruda süre baskısı yaşıyor ve deneme analizini haftalık sistematik yapamıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "T.8.1.9. Sözel mantık ve muhakeme gerektiren metinleri analiz etme",
    },
    "2028 YKS TYT-AYT Disiplin Rutini": {
        "exam": "2028 TYT - AYT",
        "hours": 30,
        "goal": "TYT 98+ net ve AYT alanında ilk 15 bin derece hedefi",
        "focus": "Haftalık Dengeli Çalışma Programı",
        "weak": "TYT Türkçe paragraf süresi ve anlam-yorum çeldiricileri, problem kurma hızı, AYT konu tekrar gecikmeleri ve deneme tempo dalgalanması",
        "status": "Düzenli çalışma alışkanlığı var ancak TYT süresini verimli bölme ve hafta sonu deneme analizlerini puan/net aksiyonuna çevirme konusunda gelişime açık",
        "branch": "Türk Dili ve Edebiyatı",
        "grade": "12. Sınıf (YKS)",
        "topic": "Nitelikli metinlerde yapı, bakış açısı, anlatım biçimleri ve anlam çözümlemesi",
    },
    "Dikkat & Çeldirici Tuzağı Temizliği": {
        "exam": "2028 LGS",
        "hours": 16,
        "goal": "Dikkat ve okuma kaynaklı net kayıplarını %70 azaltmak ve istikrarlı net artışı sağlamak",
        "focus": "Dikkat ve Çeldirici Tuzağı Azaltma",
        "weak": "Soru kökündeki 'ulaşılamaz / çıkarılamaz / değildir' ifadelerini kaçırma, öncül atlama ve çeldirici şıklara erken yönelme",
        "status": "Konu bilgisi yeterli olmasına rağmen hızlı okuma ve aceleci işaretleme nedeniyle bildiği sorularda hata yapıyor",
        "branch": "Türkçe",
        "grade": "8. Sınıf (LGS)",
        "topic": "Yazım kuralları ve noktalama işaretlerinin anlamı belirlemedeki işlevi",
    },
}

SUBSCRIPTION_TIERS: List[Dict[str, Any]] = [
    {
        "id": "ogrenci_temel",
        "segment": "Öğrenci",
        "name": "2028 LGS Başlangıç Dijital",
        "monthly_price": 1990,
        "monthly_quota": 400,
        "coach_quota": 30,
        "badge": "Temel",
        "desc": "5–8. sınıf öğrencileri için kazanım tekrarları, çeldirici farkındalığı ve adımlı çözüm desteği.",
        "features": [
            "Aylık 400 Maarif Modeli & LGS soru üretimi",
            "Öğrenci Portalı mini deneme ve ipucu/çözüm asistanı",
            "2028 LGS canlı geri sayım ve haftalık hedef takibi",
            "Çeldirici analizli standart DOCX test çıktısı",
            "Hata defteri ve kazanım özeti",
            "Mobil & masaüstü uyumlu kesintisiz çalışma arayüzü",
        ],
    },
    {
        "id": "ogrenci_pro",
        "segment": "Öğrenci",
        "name": "2028 LGS Hedef Pro",
        "monthly_price": 3290,
        "monthly_quota": 1500,
        "coach_quota": 90,
        "badge": "En Çok Tercih Edilen",
        "desc": "Yeni nesil muhakeme soruları, otomatik redaksiyon, Dijimentor planı ve veli raporu.",
        "features": [
            "Aylık 1.500 yapay zekâ soru ve redaksiyon hakkı",
            "Otomatik redaktör ile çeldirici ve soru kökü denetimi",
            "Dijimentor haftalık akıllı çalışma ve tempo programı",
            "Veli Takip Paneli canlı bildirim ve özet erişimi",
            "Soru köküne uygun başlıklandırılmış DOCX test çıktısı",
            "Yanlış soruları otomatik Hata Defteri'ne aktarma",
        ],
    },
    {
        "id": "ogrenci_elit",
        "segment": "Öğrenci",
        "name": "2028 LGS Tam Donanım Elit",
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
        "name": "2028 LGS Final & Deneme Kampı",
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
        "name": "2028 YKS MF (Sayısal) Akademi",
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
        "name": "2028 YKS EA & Sözel Atölye",
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
        "name": "2028 YKS Derece VIP",
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
        "name": "2028 TYT Hız & Paragraf Sprinti",
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
        "name": "Aile Takip & Rehberlik (2028)",
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
        "name": "Çoklu Kardeş Aile Paketi (2028)",
        "monthly_price": 3690,
        "monthly_quota": 2200,
        "coach_quota": 130,
        "badge": "2 Öğrenci",
        "desc": "İki kardeş için ayrı sınıf/sınav seviyelerinde takip, soru üretimi ve Dijimentor desteği.",
        "features": [
            "2 ayrı öğrenci profili (5–12. sınıf / 2028 LGS & YKS)",
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
        "name": "Aile Rehberlik & Deneme Plus (2028)",
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
        "name": "Öğretmen & Zümre Stüdyosu (2028)",
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
        "name": "Etüt Merkezi & Kurs Paketi (2028)",
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
        "name": "Okul & Kampüs Lisansı (2028)",
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
        "name": "Kurumsal Zincir & Yayınevi (2028)",
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
        "name": "Akademik İçerik & Soru Bankası Pro (2028)",
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

PARENT_MESSAGE_TOPICS: Dict[str, Dict[str, Any]] = {
    "1. Akademik İlerleme ve Başarı Takibi": {
        "subtopics": [
            "Haftalık Soru ve Net Özeti",
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi",
            "Deneme Sınavı ve 2028 LGS Değerlendirmesi",
        ],
        "templates": {
            "Haftalık Soru ve Net Özeti": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Ofis (2028) platformunda öğrencimiz {student_name} ({grade}) "
                "bu hafta toplam {weekly_solved} soru çözmüştür. "
                "{correct_count} doğru, {wrong_count} yanlış ve {blank_count} boş ile "
                "%{accuracy} başarı oranı elde edilmiştir.\n\n"
                "Güçlü Kazanım: {strong_topic}\n"
                "Pekiştirilmesi Önerilen Alan: {weak_topic}\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in Türkiye Yüzyılı Maarif Modeli beceri temelli kazanım tablosu güncellenmiştir. "
                "'{strong_topic}' alanında güçlü bir ilerleme kaydedilirken, '{weak_topic}' kazanımında "
                "ek pekiştirme testleri tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Deneme Sınavı ve 2028 LGS Değerlendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in son deneme performansında doğruluk oranı %{accuracy} seviyesindedir. "
                "Bir sonraki deneme oturumumuz {next_exam_date} tarihinde '{next_exam_name}' adıyla yapılacaktır.\n\n"
                "Kurşun Kalem Dijital Ofis"
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
                "Kurşun Kalem Dijital Ofis"
            ),
            "Aktif Kullanım ve Düzenlilik": (
                "Sayın {parent_name},\n\n"
                "{student_name} bu hafta platformda düzenli soru ve mini test çözümü gerçekleştirmiştir. "
                "Günlük kısa tekrar rutinlerinin sürdürülmesi önerilmektedir.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Eksik Telafi Uyarıları": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in '{weak_topic}' kazanımındaki telafi görevlerinin tamamlanması "
                "konu bütünlüğü açısından önem taşımaktadır.\n\n"
                "Kurşun Kalem Dijital Ofis"
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
                "Kurşun Kalem Dijital Ofis"
            ),
            "Ödev ve Etkinlik Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} için '{weak_topic}' ve '{strong_topic}' odaklı "
                "haftalık çalışma görevi tanımlanmıştır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Özel Dönem Bilgilendirmeleri": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Ofis 2028 kamp ve kazanım tekrar programları aktif edilmiştir. "
                "Detaylara panel üzerinden ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
    "4. Abonelik, Paket ve Kurumsal İşlemler": {
        "subtopics": [
            "Ofis Çantası & Paket Durumu",
            "Ödeme ve Fatura Bilgilendirmesi",
            "Destek ve Veli-Öğretmen İletişimi",
        ],
        "templates": {
            "Ofis Çantası & Paket Durumu": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name} hesabında {rem_questions} soru üretim kotası ve "
                "{rem_mentor} Dijimentor plan hakkı tanımlıdır.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Ödeme ve Fatura Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Abonelik ve ödeme bilgilendirmeleri {parent_email} adresine iletilmektedir. "
                "Her türlü soru için Veli Takip Paneli üzerinden bize ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
            "Destek ve Veli-Öğretmen İletişimi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in gelişim süreciyle ilgili rehberlik ve öğretmen görüşmesi "
                "taleplerinizi Veli Takip Paneli üzerinden iletebilirsiniz.\n\n"
                "Kurşun Kalem Dijital Ofis"
            ),
        },
    },
}

LOGO_DATA_URI = svg_to_data_uri(EMBEDDED_LOGO_SVG)
BACKPACK_DATA_URI = svg_to_data_uri(SCHOOL_BACKPACK_SVG)
HERO_SHOWCASE_URI = svg_to_data_uri(HERO_SHOWCASE_SVG)
LOGO_PNG_BYTES = get_embedded_logo_png_bytes()

# ==============================================================================
# KUNDUZ İLHAMLI KURUMSAL TASARIM SİSTEMİ + GİRİŞ ANİMASYONLARI
# ==============================================================================
REFINED_OFFICE_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');

:root {
    --kp-navy: #1E2952;
    --kp-navy-2: #2E3F8F;
    --kp-indigo: #5A52E8;
    --kp-purple-bar: #6E47E5;
    --kp-blue-accent: #2563EB;
    --kp-navy-soft: #EEF2FF;
    --kp-gold: #D97706;
    --kp-gold-bright: #FACC15;
    --kp-bg: #FFFFFF;
    --kp-surface: #FFFFFF;
    --kp-surface-alt: #F8FAFC;
    --kp-border: #E2E8F0;
    --kp-text: #1E293B;
    --kp-muted: #64748B;
}

@keyframes kpSoftEnter {
    from { opacity: 0.72; transform: translateY(5px); }
    to { opacity: 1; transform: translateY(0); }
}

@keyframes kpHeroEntrance {
    0% { opacity: 0; transform: translateY(10px) scale(0.99); }
    100% { opacity: 1; transform: translateY(0) scale(1); }
}

@keyframes kpStaggerCard {
    0% { opacity: 0; transform: translateY(8px); }
    100% { opacity: 1; transform: translateY(0); }
}

@keyframes kpTopProgressOnce {
    0% { width: 0%; opacity: 0.9; }
    70% { width: 85%; opacity: 0.9; }
    100% { width: 100%; opacity: 0; }
}

@keyframes kpSkeletonShimmer {
    0% { background-position: -200% 0; }
    100% { background-position: 200% 0; }
}

.kp-top-loader {
    position: fixed;
    top: 0;
    left: 0;
    height: 3px;
    z-index: 9998;
    pointer-events: none;
    background: linear-gradient(90deg, var(--kp-purple-bar) 0%, var(--kp-blue-accent) 60%, var(--kp-gold-bright) 100%);
    animation: kpTopProgressOnce 0.45s ease-out forwards;
}

.kp-skeleton-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 14px;
    padding: 1rem 1.15rem;
    margin-bottom: 0.85rem;
}
.kp-skel-line {
    height: 12px;
    border-radius: 8px;
    margin-bottom: 0.55rem;
    background: linear-gradient(90deg, #F1F5F9 25%, #E2E8F0 50%, #F1F5F9 75%);
    background-size: 200% 100%;
    animation: kpSkeletonShimmer 1.35s ease-in-out infinite;
}

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    color: var(--kp-text);
}
h1, h2, h3, .kp-kd-hero-title, .kp-page-strip-title, .kp-pkg-title, .kp-mod-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
}
.stApp { background: #FFFFFF; }
[data-testid="stHeader"] {
    background: rgba(255, 255, 255, 0.96);
    backdrop-filter: blur(8px);
    border-bottom: 1px solid var(--kp-border);
}
[data-testid="stDecoration"] { display: none; }
.block-container {
    padding-top: 0.75rem;
    padding-bottom: 4.5rem;
    max-width: 1280px;
    animation: kpSoftEnter 0.24s cubic-bezier(0.22, 1, 0.36, 1) both;
}

section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid var(--kp-border);
}
section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
    gap: 0.34rem;
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
    width: 42px;
    height: 42px;
    object-fit: contain;
    border-radius: 10px;
    border: 1px solid #CBD5E1;
    padding: 2px;
    background: #FFFFFF;
}
.kp-sidebar-brand-title {
    font-size: 0.98rem;
    font-weight: 800;
    color: #0F172A;
    line-height: 1.2;
}
.kp-sidebar-brand-sub {
    font-size: 0.78rem;
    color: var(--kp-muted);
    margin-top: 2px;
}
/* ---------- Sol menü: çerçevesiz, ikonlu linkler ---------- */
[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    display: flex !important;
    justify-content: flex-start !important;
    align-items: center !important;
    gap: 0.55rem;
    border-radius: 8px !important;
    padding: 0.5rem 0.75rem !important;
    min-height: 2.5rem;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    border: 0 !important;
    background: transparent !important;
    color: #334155 !important;
    box-shadow: none !important;
    transition: background-color 0.14s ease, color 0.14s ease !important;
}
[data-testid="stSidebar"] .stButton > button > div { justify-content: flex-start !important; }
[data-testid="stSidebar"] .stButton > button:hover {
    background: #F1F5F9 !important;
    color: #0F172A !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #EFF6FF !important;
    color: #1D4ED8 !important;
    font-weight: 600 !important;
}
.kp-nav-title {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #64748B;
    padding: 0.9rem 0.75rem 0.2rem 0.75rem;
}
.kp-sidebar-meta {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
    padding: 0.85rem 0.9rem;
    margin-top: 1rem;
    font-size: 0.8rem;
    color: #334155;
    line-height: 1.45;
}
.kp-meta-head { display: flex; justify-content: space-between; align-items: center; font-weight: 700; color: #0F172A; }
.kp-meta-sub { color: #475569; margin: 0.15rem 0 0.6rem 0; }
.kp-meter-label { display: flex; justify-content: space-between; font-size: 0.76rem; color: #475569; margin: 0.45rem 0 0.2rem 0; }
.kp-meter { height: 6px; background: #E2E8F0; border-radius: 3px; overflow: hidden; }
.kp-meter > span { display: block; height: 100%; background: #1D4ED8; border-radius: 3px; }
.kp-status { display: flex; align-items: center; gap: 0.4rem; font-size: 0.76rem; color: #475569; margin-top: 0.7rem; padding-top: 0.6rem; border-top: 1px solid #E2E8F0; }
.kp-status-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

/* ---------- Üst şerit ---------- */
.kp-topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    flex-wrap: wrap;
    padding: 0.2rem 0 0.9rem 0;
    margin-bottom: 1.5rem;
    border-bottom: 1px solid #E2E8F0;
    font-size: 0.85rem;
    color: #475569;
}
.kp-topbar-brand { font-weight: 700; color: #0F172A; }
.kp-topbar a { color: #1D4ED8; text-decoration: none; }

/* ---------- Hero ---------- */
.kp-eyebrow {
    display: inline-block;
    font-size: 0.82rem;
    font-weight: 600;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.3rem 0.75rem;
    border-radius: 999px;
    margin-bottom: 1rem;
}
.kp-hero-title {
    font-family: 'Plus Jakarta Sans', 'Inter', sans-serif;
    font-size: clamp(2rem, 3.6vw, 3.1rem);
    font-weight: 800;
    line-height: 1.1;
    letter-spacing: -0.02em;
    color: #0F172A;
    margin: 0 0 1rem 0;
    padding: 0;
}
.kp-hero-title em { font-style: normal; color: #1D4ED8; }
.kp-hero-lead { font-size: 1.1rem; line-height: 1.6; color: #475569; margin: 0 0 1.4rem 0; max-width: 34rem; }
.kp-stats { display: flex; gap: 2.2rem; flex-wrap: wrap; margin-top: 1.4rem; }
.kp-stat-num { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.7rem; font-weight: 800; color: #0F172A; line-height: 1.2; }
.kp-stat-lbl { font-size: 0.82rem; color: #64748B; }
.st-key-hero_cta_primary button, .st-key-hero_cta_secondary button {
    min-height: 3rem;
    font-size: 0.95rem !important;
    border-radius: 10px !important;
}

.kp-preview {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 16px;
    padding: 1.3rem;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.08);
}
.kp-preview-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 0.5rem; }
.kp-preview-head small { display: block; font-size: 0.78rem; color: #64748B; }
.kp-preview-head strong { font-size: 1.05rem; color: #0F172A; }
.kp-pill { font-size: 0.75rem; font-weight: 600; color: #1D4ED8; background: #EFF6FF; padding: 0.2rem 0.6rem; border-radius: 999px; white-space: nowrap; }
.kp-preview-stats { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 0.6rem; margin: 1.1rem 0; }
.kp-preview-stat { background: #F8FAFC; border-radius: 10px; padding: 0.7rem; }
.kp-preview-stat b { display: block; font-size: 1.25rem; color: #0F172A; }
.kp-preview-stat span { font-size: 0.75rem; color: #64748B; }
.kp-preview-row { display: flex; justify-content: space-between; gap: 0.75rem; font-size: 0.86rem; color: #334155; padding: 0.6rem 0; border-top: 1px solid #E2E8F0; }
.kp-preview-row span:last-child { color: #0F172A; font-weight: 600; text-align: right; }

/* ---------- Bölüm başlıkları ---------- */
.kp-section-title { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.45rem; font-weight: 700; color: #0F172A; margin: 0; letter-spacing: -0.01em; }
.kp-section-sub { font-size: 0.95rem; color: #475569; margin: 0.3rem 0 0.9rem 0; }
.kp-gap { height: 3rem; }

/* ---------- Kartlı kapsayıcılar (st.container(border=True)) ---------- */
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 14px !important; }

/* Sınıf seçici: hap şeklinde butonlar */
[class*="st-key-grade_chip_"] button { border-radius: 999px !important; min-height: 2.75rem; font-weight: 500 !important; font-size: 0.9rem !important; }

/* Modül kartları */
.kp-mod-icon { width: 40px; height: 40px; border-radius: 10px; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 0.85rem; margin-bottom: 0.7rem; }
.kp-mod-title { font-family: 'Plus Jakarta Sans', sans-serif; font-weight: 700; font-size: 1rem; color: #0F172A; }
.kp-mod-desc { font-size: 0.86rem; line-height: 1.5; color: #475569; margin-top: 0.3rem; min-height: 3.9rem; }
[class*="st-key-mod_open_"] button, .st-key-prio_all button {
    border: 0 !important;
    background: transparent !important;
    color: #1D4ED8 !important;
    padding: 0 !important;
    min-height: 0 !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    justify-content: flex-start !important;
}
[class*="st-key-mod_open_"] button:hover, .st-key-prio_all button:hover { color: #1E3A8A !important; text-decoration: underline; }

/* Öncelik satırları */
.kp-prio-num { width: 32px; height: 32px; border-radius: 50%; background: #F1F5F9; font-weight: 700; font-size: 0.9rem; display: flex; align-items: center; justify-content: center; color: #0F172A; }
.kp-prio-title { font-weight: 600; font-size: 0.95rem; color: #0F172A; line-height: 1.4; }
.kp-prio-sub { font-size: 0.8rem; color: #64748B; margin-top: 0.15rem; }
.kp-prio-divider { border-top: 1px solid #E2E8F0; margin: 0.35rem 0; }
.kp-tag-warn { font-size: 0.75rem; font-weight: 600; color: #92400E; background: #FEF3C7; padding: 0.2rem 0.6rem; border-radius: 999px; }
.kp-page-strip {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 1rem;
    padding: 0.05rem 0 0.7rem 0;
    margin-bottom: 1rem;
    border-bottom: 1px solid var(--kp-border);
    animation: kpHeroEntrance 0.32s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-page-strip-title {
    font-size: 1.28rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0;
}
.kp-page-strip-desc {
    font-size: 0.83rem;
    color: var(--kp-muted);
    margin: 0.2rem 0 0 0;
}

.kp-kademe-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 16px;
    padding: 1rem 1.1rem;
    height: 146px;
    display: flex;
    flex-direction: column;
    justify-content: flex-start;
    margin-bottom: 0.55rem;
    box-shadow: 0 3px 12px rgba(37, 99, 235, 0.04);
    transition: transform 0.16s ease, box-shadow 0.16s ease;
    animation: kpStaggerCard 0.44s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-kademe-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 18px rgba(37, 99, 235, 0.09);
}
.kp-kademe-tag {
    display: inline-block;
    width: fit-content;
    font-size: 0.7rem;
    font-weight: 800;
    color: #FFFFFF;
    padding: 0.22rem 0.65rem;
    border-radius: 999px;
}
.kp-kademe-title {
    font-size: 1rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.55rem 0 0.25rem 0;
}
.kp-kademe-desc {
    font-size: 0.8rem;
    color: var(--kp-muted);
    line-height: 1.42;
    margin: 0;
}

.kp-priority-card {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-left: 4px solid #2563EB;
    border-radius: 14px;
    padding: 0.95rem 1.05rem;
    height: 132px;
    margin-bottom: 0.55rem;
    animation: kpStaggerCard 0.46s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-priority-badge {
    font-size: 0.7rem;
    font-weight: 800;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.18rem 0.58rem;
    border-radius: 999px;
    display: inline-block;
}
.kp-priority-title {
    font-size: 0.94rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.45rem 0 0.2rem 0;
}
.kp-priority-sub {
    font-size: 0.79rem;
    color: var(--kp-muted);
    margin: 0;
}

.kp-value-strip {
    background: linear-gradient(90deg, #1E2952 0%, #2563EB 65%, #5A52E8 100%);
    color: #FFFFFF;
    border-radius: 14px;
    padding: 0.85rem 1.2rem;
    margin-bottom: 1.05rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 1rem;
    flex-wrap: wrap;
}
.kp-value-strip-item {
    display: flex;
    align-items: center;
    gap: 0.48rem;
    font-size: 0.82rem;
    font-weight: 600;
}
.kp-value-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #FACC15;
    display: inline-block;
}

div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 14px;
    padding: 0.92rem 1.05rem;
    box-shadow: 0 3px 10px rgba(15, 23, 42, 0.04);
    animation: kpStaggerCard 0.42s cubic-bezier(0.22, 1, 0.36, 1) both;
}
div[data-testid="stMetric"] [data-testid="stMetricValue"] {
    font-size: 1.52rem;
    font-weight: 800;
    color: var(--kp-navy);
}

.kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-kpi-card, .kp-metric-box, .kp-studio-banner {
    background: #FFFFFF;
    border: 1px solid var(--kp-border);
    border-radius: 16px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
    animation: kpStaggerCard 0.42s cubic-bezier(0.22, 1, 0.36, 1) both;
}
.kp-card, .kp-bag-shell, .kp-pkg-card, .kp-pay-box, .kp-studio-banner {
    padding: 1.15rem 1.25rem;
}
.kp-card { margin-bottom: 0.65rem; }

.kp-studio-banner {
    margin-bottom: 1.05rem;
    border-left: 4px solid #2563EB;
}
.kp-studio-title {
    font-size: 1.14rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0;
}
.kp-studio-sub {
    font-size: 0.84rem;
    color: var(--kp-muted);
    margin: 0.22rem 0 0 0;
}
.kp-kpi-card { padding: 0.92rem 1.02rem; }
.kp-kpi-label {
    font-size: 0.7rem;
    font-weight: 700;
    color: #64748B;
    text-transform: uppercase;
}
.kp-kpi-value {
    font-size: 1.25rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin-top: 0.2rem;
}
.kp-kpi-sub {
    font-size: 0.76rem;
    color: var(--kp-muted);
    margin-top: 0.15rem;
}

.kp-metric-box {
    padding: 1.05rem 0.95rem;
    text-align: center;
    border-bottom: 4px solid #2563EB;
}
.kp-metric-num {
    font-size: 1.62rem;
    font-weight: 800;
    color: var(--kp-navy);
    line-height: 1.15;
    font-variant-numeric: tabular-nums;
}
.kp-metric-lbl {
    font-size: 0.71rem;
    color: #64748B;
    margin-top: 0.3rem;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 700;
}

.kp-bag-shell {
    margin-bottom: 1.2rem;
    border-top: 4px solid #2563EB;
}
.kp-bag-visual {
    background: linear-gradient(160deg, #1E2952 0%, #2563EB 100%);
    border-radius: 14px;
    padding: 1rem 0.85rem;
    display: flex;
    flex-direction: column;
    align-items: center;
    color: #FFFFFF;
}
.kp-bag-visual img {
    width: 102px;
    height: 118px;
    object-fit: contain;
}
.kp-bag-count-pill {
    margin-top: 0.5rem;
    background: #FACC15;
    color: #0F172A;
    font-size: 0.74rem;
    font-weight: 800;
    padding: 0.2rem 0.7rem;
    border-radius: 999px;
}
.kp-bag-item-chip {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-left: 3.5px solid #2563EB;
    border-radius: 10px;
    padding: 0.62rem 0.85rem;
    margin-bottom: 0.45rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.kp-pkg-card { margin-bottom: 0.65rem; }
.kp-pkg-featured {
    border: 2px solid #2563EB !important;
    background: linear-gradient(180deg, #EFF6FF 0%, #FFFFFF 35%) !important;
}
.kp-pkg-top {
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 0.5rem;
}
.kp-pkg-seg {
    font-size: 0.71rem;
    font-weight: 700;
    color: #1D4ED8;
    background: #EFF6FF;
    padding: 0.2rem 0.58rem;
    border-radius: 6px;
}
.kp-pkg-badge {
    font-size: 0.71rem;
    font-weight: 800;
    color: #92400E;
    background: #FEF3C7;
    border: 1px solid #F59E0B;
    padding: 0.2rem 0.58rem;
    border-radius: 6px;
}
.kp-pkg-title {
    font-size: 1.04rem;
    font-weight: 800;
    color: var(--kp-navy);
    margin: 0.55rem 0 0.2rem 0;
}
.kp-pkg-price {
    font-size: 1.38rem;
    font-weight: 800;
    color: var(--kp-text);
    margin-bottom: 0.4rem;
}
.kp-pkg-price span {
    font-size: 0.8rem;
    font-weight: 500;
    color: #64748B;
}
.kp-pkg-desc {
    font-size: 0.82rem;
    color: var(--kp-muted);
    line-height: 1.42;
    margin-bottom: 0.7rem;
    min-height: 3.1rem;
}
.kp-pkg-feats {
    list-style: none;
    padding: 0.65rem 0 0 0;
    margin: 0;
    border-top: 1px solid var(--kp-border);
}
.kp-pkg-feats li {
    font-size: 0.79rem;
    color: #334155;
    padding: 0.16rem 0;
}

.kp-dist-bar {
    display: flex;
    height: 9px;
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

.stButton > button,
div[data-testid="stFormSubmitButton"] > button {
    border-radius: 8px;
    font-weight: 600;
    font-size: 0.9rem;
    border: 1px solid #CBD5E1;
    background-color: #FFFFFF;
    color: #1E293B;
}
.stButton > button:hover,
div[data-testid="stFormSubmitButton"] > button:hover {
    border-color: #2563EB;
    background-color: #EFF6FF;
    color: #1D4ED8;
}
.stButton > button[kind="primary"],
div[data-testid="stFormSubmitButton"] > button[kind="primary"] {
    background: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFFFFF !important;
}
.stButton > button[kind="primary"]:hover,
div[data-testid="stFormSubmitButton"] > button[kind="primary"]:hover {
    background: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
}

@media (prefers-reduced-motion: reduce) {
    .block-container,
    .kp-top-loader,
    .kp-top-compact-bar,
    .kp-kd-kicker,
    .kp-kd-hero-title,
    .kp-kd-hero-sub,
    .kp-kd-social-row,
    .kp-feature-strip-card,
    .kp-showcase-frame,
    .kp-page-strip,
    .kp-kademe-card,
    .kp-priority-card,
    div[data-testid="stMetric"],
    .kp-card,
    .kp-bag-shell,
    .kp-pkg-card,
    .kp-pay-box,
    .kp-kpi-card,
    .kp-metric-box,
    .kp-studio-banner {
        animation: none !important;
        opacity: 1 !important;
        transform: none !important;
    }
}
"""

css_to_inject = REFINED_OFFICE_CSS
if EXTERNAL_CSS_PATH.exists():
    try:
        ext_css = EXTERNAL_CSS_PATH.read_text(encoding="utf-8").strip()
        if ext_css:
            css_to_inject = f"{REFINED_OFFICE_CSS}\n\n{ext_css}"
    except Exception:
        pass

st.markdown(f"<style>{css_to_inject}</style>", unsafe_allow_html=True)


def trigger_loader_bar() -> None:
    st.session_state["loader_nonce"] = uuid.uuid4().hex[:8]


def render_top_loader_bar() -> None:
    nonce = st.session_state.get("loader_nonce", "init")
    st.markdown(
        f'<div class="kp-top-loader" data-load-id="{html.escape(str(nonce))}"></div>',
        unsafe_allow_html=True,
    )


def render_loading_skeleton(placeholder: Any, lines: int = 3) -> None:
    bars = "".join(['<div class="kp-skel-line" style="width:94%;"></div>' for _ in range(max(2, lines))])
    placeholder.markdown(f'<div class="kp-skeleton-card">{bars}</div>', unsafe_allow_html=True)


# ==============================================================================
# YARDIMCI FONKSİYONLAR, KOTA YÖNETİMİ VE KALICI DURUM
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
        "weak_topic": "Sözel Mantık ve Muhakeme",
        "next_exam_date": (datetime.now() + timedelta(days=5)).strftime("%d.%m.%Y"),
        "next_exam_name": "Türkiye Geneli 2028 LGS Deneme Sınavı",
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
        {"time": "09:15:00", "module": "Dijital Ofis 2028 • Fabrika", "summary": "8. Sınıf 2028 LGS Muhakeme seti hazır"},
        {"time": "10:42:18", "module": "Dijital Ofis 2028 • Dijimentor", "summary": "Haftalık çeldirici & tempo analizi aktif"},
    ]


def load_persisted_state() -> Dict[str, Any]:
    if not STATE_FILE_PATH.exists():
        return {}
    try:
        raw = STATE_FILE_PATH.read_text(encoding="utf-8")
        data = json.loads(raw)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_persisted_state() -> None:
    payload = {
        "backpack_cart": st.session_state.get("backpack_cart", {}),
        "selected_tier_id": st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"]),
        "used_question_quota": int(st.session_state.get("used_question_quota", 0)),
        "used_coach_quota": int(st.session_state.get("used_coach_quota", 0)),
        "remaining_question_quota": int(st.session_state.get("remaining_question_quota", 0)),
        "remaining_coach_quota": int(st.session_state.get("remaining_coach_quota", 0)),
        "parent_student_stats": st.session_state.get("parent_student_stats", _default_parent_stats()),
        "parent_announcements": st.session_state.get("parent_announcements", _default_parent_announcements()),
        "activity_log": st.session_state.get("activity_log", _default_activity_log()),
        "question_bank_2028": st.session_state.get("question_bank_2028", []),
        "error_notebook_2028": st.session_state.get("error_notebook_2028", []),
        "b2b_org_cache": st.session_state.get("b2b_org_cache"),
        "last_b2b_summary": st.session_state.get("last_b2b_summary", ""),
    }
    try:
        STATE_FILE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def attach_logo_to_doc(doc: Document, width_in: float = 1.15) -> None:
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
    api_key, secret_key, base_url = resolve_iyzico_config()
    if not api_key or not secret_key:
        return {
            "status": "failure",
            "errorMessage": (
                "IYZICO_API_KEY ve IYZICO_SECRET_KEY tanımlı değil. "
                "Streamlit Secrets veya ortam değişkenlerine anahtarlarınızı ekleyin."
            ),
        }
    options = {"api_key": api_key, "secret_key": secret_key, "base_url": base_url}
    price_str = f"{max(1.0, float(amount_try)):.2f}"
    request = {
        "locale": "tr",
        "conversationId": f"KP28-{int(time.time())}",
        "price": price_str,
        "paidPrice": price_str,
        "currency": "TRY",
        "basketId": f"B28-{int(time.time())}",
        "paymentGroup": "PRODUCT",
        "callbackUrl": "https://kursunkalemdijitalofis.com.tr",
        "enabledInstallments": [1, 2, 3, 6, 9, 12],
        "buyer": {
            "id": "BY_KP28_01",
            "name": buyer_name.strip() or "Sinan Gürbüz",
            "surname": buyer_surname.strip() or "Kurşun",
            "gsmNumber": buyer_phone.strip() or "+905350000000",
            "email": buyer_email.strip() or "iletisim@kursunkalemdijitalofis.com.tr",
            "identityNumber": "74300864791",
            "registrationAddress": "İstanbul, Türkiye",
            "ip": "85.34.78.112",
            "city": "Istanbul",
            "country": "Turkey",
        },
        "shippingAddress": {
            "contactName": f"{buyer_name.strip() or 'Sinan Gürbüz'} {buyer_surname.strip() or 'Kurşun'}",
            "city": "Istanbul",
            "country": "Turkey",
            "address": "İstanbul, Türkiye",
        },
        "billingAddress": {
            "contactName": f"{buyer_name.strip() or 'Sinan Gürbüz'} {buyer_surname.strip() or 'Kurşun'}",
            "city": "Istanbul",
            "country": "Turkey",
            "address": "İstanbul, Türkiye",
        },
        "basketItems": [
            {
                "id": "BI_KP28_101",
                "name": basket_title[:120],
                "category1": "Dijital Eğitim 2028",
                "itemType": "VIRTUAL",
                "price": price_str,
            }
        ],
    }
    checkout_form_initialize = iyzipay.CheckoutFormInitialize().create(request, options)
    raw_res = checkout_form_initialize.read().decode("utf-8")
    return json.loads(raw_res)


def get_gemini_client() -> genai.Client:
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
    client = get_gemini_client()
    candidate_models = [DEFAULT_MODEL, FALLBACK_MODEL]
    if model and model not in (
        "gemini-2.5-pro",
        "gemini-3.1-pro-preview",
        "gemini-2.5-flash-lite",
    ):
        candidate_models.insert(0, model)

    deduped_models: List[str] = []
    for m in candidate_models:
        if m and m not in deduped_models:
            deduped_models.append(m)

    errors_by_model: List[str] = []
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
                err_str = str(exc)
                errors_by_model.append(f"[{candidate_model}] {err_str}")
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                if any(tok in err_str for tok in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED", "overloaded")):
                    time.sleep(1.2 * (attempt + 1))
                    continue
                break

    raise RuntimeError(f"Model çağrısı başarısız oldu: {' | '.join(errors_by_model[-2:])}")


def clear_page_ephemeral_state(target_page: str) -> None:
    if target_page != "Öğrenci Portalı & Çözüm":
        st.session_state.pop("last_exam_report", None)
    if target_page != "Öğrenci Röntgeni":
        st.session_state.pop("xray_last_scan", None)


def navigate_to(page_name: str) -> None:
    if page_name in PAGES:
        clear_page_ephemeral_state(page_name)
        st.session_state.active_page = page_name
        trigger_loader_bar()
        st.rerun()


def record_activity(module_name: str, summary: str) -> None:
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
    return next((t for t in SUBSCRIPTION_TIERS if t["id"] == tier_id), SUBSCRIPTION_TIERS[1])


def ensure_bag_initialized() -> None:
    if "backpack_cart" not in st.session_state or not isinstance(st.session_state["backpack_cart"], dict):
        default_id = st.session_state.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
        st.session_state["backpack_cart"] = {default_id: 1}


def get_bag_items() -> Dict[str, int]:
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
        title = entries[0]["tier"]["name"] if len(entries) == 1 else f"Ofis Çantası ({total_count} Paket)"
        segment = entries[0]["tier"]["segment"] if len(entries) == 1 else "Karma / Çoklu Paket"
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


def sync_plan_quotas(summary: Dict[str, Any], reset_usage: bool = False) -> None:
    if reset_usage:
        st.session_state["used_question_quota"] = 0
        st.session_state["used_coach_quota"] = 0

    used_q = max(0, int(st.session_state.get("used_question_quota", 0)))
    used_c = max(0, int(st.session_state.get("used_coach_quota", 0)))
    total_q = max(0, int(summary["monthly_quota"]))
    total_c = max(0, int(summary["coach_quota"]))

    if used_q > total_q:
        used_q = total_q
        st.session_state["used_question_quota"] = used_q
    if used_c > total_c:
        used_c = total_c
        st.session_state["used_coach_quota"] = used_c

    st.session_state["remaining_question_quota"] = max(0, total_q - used_q)
    st.session_state["remaining_coach_quota"] = max(0, total_c - used_c)
    save_persisted_state()


def cb_add_tier_to_bag(tier_id: str, qty: int = 1) -> None:
    bag = dict(get_bag_items())
    bag[tier_id] = bag.get(tier_id, 0) + max(1, int(qty))
    st.session_state["backpack_cart"] = bag
    st.session_state["selected_tier_id"] = tier_id
    sync_plan_quotas(get_bag_summary(), reset_usage=False)
    tier = get_tier_by_id(tier_id)
    record_activity("Ofis Çantası", f"{tier['name']} çantaya eklendi")


def cb_update_bag_qty(tier_id: str, delta: int) -> None:
    bag = dict(get_bag_items())
    if tier_id not in bag:
        return
    new_qty = bag[tier_id] + delta
    tier = get_tier_by_id(tier_id)
    if new_qty <= 0:
        del bag[tier_id]
        record_activity("Ofis Çantası", f"{tier['name']} çantadan çıkarıldı")
    else:
        bag[tier_id] = new_qty
        record_activity("Ofis Çantası", f"{tier['name']} adedi {new_qty} yapıldı")
    st.session_state["backpack_cart"] = bag
    if bag:
        st.session_state["selected_tier_id"] = next(iter(bag.keys()))
    sync_plan_quotas(get_bag_summary(), reset_usage=False)


def cb_clear_bag() -> None:
    st.session_state["backpack_cart"] = {}
    sync_plan_quotas(get_bag_summary(), reset_usage=False)
    record_activity("Ofis Çantası", "Çantadaki tüm paketler silindi")


def consume_quota(kind: str, amount: int = 1) -> Tuple[bool, str]:
    summary = get_bag_summary()
    sync_plan_quotas(summary, reset_usage=False)

    if kind == "question":
        rem = int(st.session_state.get("remaining_question_quota", 0))
        if rem < amount:
            return False, (
                f"Ofis çantanızda ({summary['name']}) yeterli soru üretim kotası kalmadı "
                f"(Kalan: {rem}, İstenen: {amount}). Paketler vitrininden paket ekleyebilirsiniz."
            )
        st.session_state["used_question_quota"] = int(st.session_state.get("used_question_quota", 0)) + amount
        sync_plan_quotas(summary, reset_usage=False)
        return True, ""

    rem_c = int(st.session_state.get("remaining_coach_quota", 0))
    if rem_c < amount:
        return False, (
            f"Ofis çantanızda ({summary['name']}) Dijimentor plan kotası doldu "
            f"(Kalan: {rem_c}). Paketler vitrininden paket ekleyebilirsiniz."
        )
    st.session_state["used_coach_quota"] = int(st.session_state.get("used_coach_quota", 0)) + amount
    sync_plan_quotas(summary, reset_usage=False)
    return True, ""


def refund_quota(kind: str, amount: int = 1) -> None:
    if kind == "question":
        st.session_state["used_question_quota"] = max(
            0, int(st.session_state.get("used_question_quota", 0)) - amount
        )
    else:
        st.session_state["used_coach_quota"] = max(
            0, int(st.session_state.get("used_coach_quota", 0)) - amount
        )
    sync_plan_quotas(get_bag_summary(), reset_usage=False)


def add_to_question_bank(
    grade: str,
    branch: str,
    outcome: str,
    difficulty: str,
    content_text: str,
) -> None:
    bank: List[Dict[str, str]] = st.session_state.setdefault("question_bank_2028", [])
    bank.insert(
        0,
        {
            "id": uuid.uuid4().hex[:6].upper(),
            "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "grade": grade,
            "branch": branch,
            "outcome": outcome,
            "difficulty": difficulty,
            "content": content_text,
        },
    )
    st.session_state["question_bank_2028"] = bank[:40]
    save_persisted_state()


def add_to_error_notebook(
    question_num: int,
    question_text: str,
    user_ans: str,
    correct_ans: str,
    elapsed_sec: float,
) -> None:
    nb: List[Dict[str, Any]] = st.session_state.setdefault("error_notebook_2028", [])
    now = datetime.now()
    diag = "Süre / Dikkat Baskısı" if elapsed_sec > 90 else "Çeldirici / Kavram İncelemesi"
    nb.insert(
        0,
        {
            "id": uuid.uuid4().hex[:6].upper(),
            "num": question_num,
            "date": now.strftime("%d.%m.%Y %H:%M"),
            "review_1d": (now + timedelta(days=1)).strftime("%d.%m.%Y"),
            "review_3d": (now + timedelta(days=3)).strftime("%d.%m.%Y"),
            "review_7d": (now + timedelta(days=7)).strftime("%d.%m.%Y"),
            "user_ans": user_ans,
            "correct_ans": correct_ans,
            "elapsed_sec": round(elapsed_sec, 1),
            "diagnosis": diag,
            "question_text": question_text,
        },
    )
    st.session_state["error_notebook_2028"] = nb[:50]
    save_persisted_state()


# ==============================================================================
# GÜÇLÜ SORU AYRIŞTIRICI & DOCX ÜRETİCİLERİ
# ==============================================================================
def extract_answer_key(raw_text: str) -> Dict[int, str]:
    if not raw_text.strip():
        return {}

    parts = re.split(
        r"(?:^|\n)#*\s*(?:Cevap\s*Anahtar\w*|Yanıt\s*Anahtar\w*)[^\n]*",
        raw_text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )
    key_map: Dict[int, str] = {}

    patterns = [
        r"(?:Soru\s*)?(\d+)\s*[\.\-\:\)\]]\s*\**\s*([A-D])\b",
        r"(\d+)\s*\.\s*Soru\s*[\:\-\)\]]\s*\**\s*([A-D])\b",
        r"\b(\d+)\s*[-–—]\s*\**\s*([A-D])\b",
    ]

    if len(parts) > 1:
        for pat in patterns:
            for m in re.finditer(pat, parts[1], flags=re.IGNORECASE):
                q_num = int(m.group(1))
                ans_letter = m.group(2).upper()
                if q_num not in key_map:
                    key_map[q_num] = ans_letter

    q_blocks = re.split(
        r"(?=(?:^|\n)\s*(?:#+\s*)?(?:\*{0,2})\s*(?:Soru\s*\d+|\d+\s*[\.\)]))",
        parts[0],
        flags=re.IGNORECASE,
    )
    for blk in q_blocks:
        num_m = re.match(r"^\s*(?:#+\s*)?(?:\*{0,2})\s*(?:Soru\s*)?(\d+)\b", blk, flags=re.IGNORECASE)
        if not num_m:
            continue
        q_num = int(num_m.group(1))
        if q_num in key_map:
            continue
        ans_m = re.search(
            r"(?:Doğru\s*(?:Cevap|Yanıt)|Cevap|Yanıt)\s*[\:\-–]\s*\**\s*([A-D])\b",
            blk,
            flags=re.IGNORECASE,
        )
        if ans_m:
            key_map[q_num] = ans_m.group(1).upper()

    return key_map


def parse_quiz_questions(raw_text: str) -> List[Dict[str, Any]]:
    if not raw_text.strip():
        return []

    answer_map = extract_answer_key(raw_text)
    main_body = re.split(
        r"(?:^|\n)\s*#*\s*(?:Cevap\s*Anahtar\w*|Yanıt\s*Anahtar\w*)",
        raw_text,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip()

    if not main_body:
        return []

    header_split_pattern = (
        r"(?=(?:^|\n)\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s+\d+|\d+\s*(?:\.\s*Soru|\)|\.\s+(?=[A-ZÇĞİÖŞÜ\"“'\*]))))"
    )
    raw_blocks = re.split(header_split_pattern, main_body, flags=re.IGNORECASE | re.MULTILINE)

    parsed: List[Dict[str, Any]] = []
    seen_nums: set[int] = set()

    for idx_b, block in enumerate(raw_blocks, start=1):
        text = block.strip()
        if len(text) < 15:
            continue

        num_match = re.match(
            r"^\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:Soru\s+)?(\d+)\b",
            text,
            flags=re.IGNORECASE,
        )
        q_num = int(num_match.group(1)) if num_match else idx_b
        if q_num in seen_nums:
            q_num = max(seen_nums) + 1

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        stem_lines: List[str] = []
        options: Dict[str, str] = {}
        current_opt: Optional[str] = None

        opt_pattern = re.compile(
            r"^\s*(?:[-*•]\s*)?(?:\*\*)?\s*([A-D])\s*[\)\.\:]\s*(?:\*\*)?\s*(.+)$",
            flags=re.IGNORECASE,
        )

        for ln in lines:
            if re.match(r"^\s*\[?(?:Redaksiyon|Çözüm|Cevap\s*Anahtar|Doğru\s*Cevap)", ln, flags=re.IGNORECASE):
                current_opt = None
                continue
            if re.match(r"^\s*[-_=]{3,}\s*$", ln):
                continue

            opt_m = opt_pattern.match(ln)
            if opt_m:
                letter = opt_m.group(1).upper()
                opt_val = re.sub(r"\*\*$", "", opt_m.group(2).strip()).strip()
                options[letter] = opt_val
                current_opt = letter
            elif current_opt:
                extra = re.sub(r"\*\*$", "", ln).strip()
                if extra:
                    options[current_opt] = f"{options[current_opt]} {extra}".strip()
            else:
                stem_lines.append(ln)

        if len(options) >= 2:
            ordered_opts = {k: options[k] for k in ["A", "B", "C", "D"] if k in options}
            for k, v in options.items():
                if k not in ordered_opts:
                    ordered_opts[k] = v

            seen_nums.add(q_num)
            parsed.append(
                {
                    "num": q_num,
                    "stem": "\n".join(stem_lines).strip(),
                    "options": ordered_opts,
                    "correct": answer_map.get(q_num),
                    "full_text": text,
                }
            )

    return parsed


def format_countdown(target: datetime) -> Tuple[int, int, int, int]:
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
    title, desc = PAGE_DESCRIPTIONS.get(page_name, (page_name, "Kurşun Kalem Dijital Ofis 2028 çalışma alanı."))
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
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"{institution.upper()} (2028) — DEĞERLENDİRME TESTİ")
    title_run.bold = True
    title_run.font.size = Pt(14)
    title_run.font.color.rgb = RGBColor(30, 41, 82)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_p.add_run(f"Branş: {branch}  |  Sınıf: {grade}  |  Öğrenme Alanı: {learning_area}")
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(71, 85, 105)

    meta_p = doc.add_paragraph()
    meta_run = meta_p.add_run(
        f"Hedef Kazanım: {outcome}\nBağlam / Tema: {context_theme}\nTarih: {datetime.now().strftime('%d.%m.%Y')}"
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
            run.font.color.rgb = RGBColor(30, 41, 82)

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer_p.add_run(f"\n{BRAND_NAME} • 2028")
    footer_run.italic = True
    footer_run.font.size = Pt(8.5)
    footer_run.font.color.rgb = RGBColor(100, 116, 139)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_coach_docx(title_meta: str, plan_text: str) -> bytes:
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} (2028) — DİJİMENTOR ÇALIŞMA PLANI")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(30, 41, 82)

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
            run.font.color.rgb = RGBColor(30, 41, 82)

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
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    attach_logo_to_doc(doc, width_in=1.15)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"{BRAND_NAME.upper()} (2028) — OFİS ÇANTASI & ABONELİK ÖZETİ")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(30, 41, 82)

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
            run.font.color.rgb = RGBColor(30, 41, 82)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_school_backpack_panel() -> Dict[str, Any]:
    summary = get_bag_summary()
    entries = summary["entries"]

    st.markdown('<div class="kp-bag-shell">', unsafe_allow_html=True)
    h_col, clr_col = st.columns([4, 1])
    with h_col:
        st.markdown("#### Kurşun Kalem Ofis Çantası (2028)")
        st.caption("Seçili paketler, adet kontrolleri, kota çubukları ve aylık toplam özet.")
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
            chips_html = '<div style="color:#E2E8F0;font-size:0.78rem;margin-top:0.35rem;text-align:center;">Çanta boş</div>'

        st.markdown(
            f"""
            <div class="kp-bag-visual">
                <img src="{BACKPACK_DATA_URI}" alt="Ofis Çantası"/>
                <div class="kp-bag-count-pill">{summary['total_count']} Paket</div>
                <div style="width:100%;margin-top:0.3rem;">{chips_html}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with list_col:
        if not entries:
            st.info("Ofis çantanız şu an boş. Aşağıdaki paketlerden **Çantaya Ekle** butonunu kullanabilirsiniz.")
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
                            <div style="font-weight:700;font-size:0.92rem;color:#2563EB;white-space:nowrap;">
                                {entry['line_price']:,} ₺/ay
                            </div>
                        </div>
                        """.replace(",", "."),
                        unsafe_allow_html=True,
                    )
                with c_minus:
                    st.button("−", key=f"bag_minus_{t['id']}", use_container_width=True, on_click=cb_update_bag_qty, args=(t["id"], -1))
                with c_plus:
                    st.button("+", key=f"bag_plus_{t['id']}", use_container_width=True, on_click=cb_update_bag_qty, args=(t["id"], +1))
                with c_rem:
                    st.button("Çıkar", key=f"bag_del_{t['id']}", use_container_width=True, on_click=cb_update_bag_qty, args=(t["id"], -qty))

            st.markdown("<div style='margin-top:0.45rem;'></div>", unsafe_allow_html=True)
            m1, m2, m3 = st.columns(3)
            m1.metric("Aylık Toplam", f"{summary['monthly_price']:,} ₺".replace(",", "."))
            m2.metric("Aylık Soru Kotası", f"{summary['monthly_quota']:,}".replace(",", "."))
            m3.metric("Dijimentor Plan", f"{summary['coach_quota']} / Ay")

            rem_q = int(st.session_state.get("remaining_question_quota", summary["monthly_quota"]))
            rem_c = int(st.session_state.get("remaining_coach_quota", summary["coach_quota"]))
            q_pct = min(1.0, max(0.0, rem_q / max(1, int(summary["monthly_quota"]))))
            c_pct = min(1.0, max(0.0, rem_c / max(1, int(summary["coach_quota"]))))

            st.markdown("<div style='margin-top:0.55rem;'></div>", unsafe_allow_html=True)
            pb1, pb2 = st.columns(2)
            with pb1:
                st.progress(q_pct, text=f"Kullanılabilir Soru Kotası: {rem_q:,} / {summary['monthly_quota']:,}".replace(",", "."))
            with pb2:
                st.progress(c_pct, text=f"Kullanılabilir Dijimentor: {rem_c} / {summary['coach_quota']} Plan")

    st.markdown("</div>", unsafe_allow_html=True)
    return summary


def render_iyzico_simulator(
    package_name: str,
    monthly_price: int,
    summary: Optional[Dict[str, Any]] = None,
) -> None:
    bag_summary = summary if summary is not None else get_bag_summary()
    effective_monthly = int(bag_summary["monthly_price"]) if bag_summary["total_count"] > 0 else int(monthly_price)
    display_label = f"{package_name} ({bag_summary['total_count']} Paket)" if bag_summary["total_count"] > 0 else package_name

    st.markdown('<div class="kp-pay-box">', unsafe_allow_html=True)
    st.markdown("#### iyzico Taksit & Ödeme Simülatörü")
    st.caption("Dönem avantajı ve taksit planına göre tutarı hesaplayın, iyzico ödeme formu başlatın veya DOCX özet indirin.")

    c1, c2, c3 = st.columns(3)
    with c1:
        billing_cycle = st.selectbox("Dönem Seçimi", ["Aylık Abonelik", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"], key=f"cycle_{display_label}")
    with c2:
        installment = st.selectbox("Taksit Sayısı (iyzico)", [1, 2, 3, 6, 9, 12], index=2, key=f"inst_{display_label}")
    with c3:
        card_brand = st.selectbox("Kart Programı", ["Bonus", "Maximum", "World", "Axess", "Paraf", "Diğer Banka Kartı"], key=f"card_{display_label}")

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
        buyer_name = st.text_input("Ad", value="Sinan Gürbüz", key=f"iyz_name_{display_label}")
    with b_col2:
        buyer_surname = st.text_input("Soyad", value="Kurşun", key=f"iyz_surname_{display_label}")
    with b_col3:
        buyer_email = st.text_input("E-posta", value="iletisim@kursunkalemdijitalofis.com.tr", key=f"iyz_email_{display_label}")

    btn_c1, btn_c2, btn_c3 = st.columns([1.8, 1.8, 1.1])
    with btn_c1:
        if st.button("iyzico Ödeme Formu Başlat (API)", key=f"iyz_live_btn_{display_label}", type="primary", use_container_width=True):
            if total_amount <= 0:
                st.warning("Ödeme başlatmak için önce çantanıza en az bir paket ekleyin.")
            else:
                load_slot = st.empty()
                render_loading_skeleton(load_slot, lines=2)
                res_dict = initialize_iyzico_checkout(
                    buyer_name=buyer_name,
                    buyer_surname=buyer_surname,
                    buyer_email=buyer_email,
                    buyer_phone="+905350000000",
                    amount_try=float(total_amount),
                    basket_title=f"{BRAND_NAME} 2028 - {display_label} ({billing_cycle})",
                )
                load_slot.empty()
                if res_dict.get("status") == "success":
                    sync_plan_quotas(bag_summary, reset_usage=True)
                    record_activity("iyzico API", f"{display_label} ({total_amount:,} ₺) ödeme formu açıldı".replace(",", "."))
                    st.session_state["iyzico_last_response"] = res_dict
                    st.success("iyzico ödeme formu başarıyla oluşturuldu ve kota yenilendi!")
                else:
                    st.error(f"iyzico yanıtı: {res_dict.get('errorMessage', 'Bilinmeyen hata')}")

    with btn_c2:
        if st.button("Simülasyon ile Kota Aktifleştir", key=f"pay_btn_{display_label}", use_container_width=True):
            sync_plan_quotas(bag_summary, reset_usage=True)
            record_activity("iyzico Simülasyon", f"{display_label} ({total_amount:,} ₺) aktiflendi".replace(",", "."))
            st.success(
                f"**{display_label}** için **{total_amount:,} ₺** ({installment} taksit, "
                f"{card_brand}) tutarında kota ({bag_summary['monthly_quota']:,} soru) aktifleşti.".replace(",", ".")
            )

    with btn_c3:
        brief_docx = build_commercial_brief_docx(
            title=f"{display_label} Özeti (2028)",
            summary=bag_summary,
            billing_cycle=billing_cycle,
            installment=int(installment),
            total_amount=int(total_amount),
            extra_notes=f"Kart Programı: {card_brand}",
        )
        st.download_button(
            label="Özet DOCX",
            data=brief_docx,
            file_name="kursun_kalem_ofis_2028_cantasi_ozet.docx",
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
                f'<button style="background:#2563EB;color:#FFFFFF;border:none;border-radius:8px;'
                f'padding:0.55rem 1.15rem;font-weight:600;cursor:pointer;">'
                f'Güvenli iyzico Ödeme Sayfasını Aç</button></a></div>',
                unsafe_allow_html=True,
            )
        if form_html:
            st.components.v1.html(str(form_html), height=460, scrolling=True)

    st.markdown("</div>", unsafe_allow_html=True)


def render_student_xray_panel() -> None:
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Öğrenci Röntgeni (2028) — Kazanım, Hata Kök Neden &amp; Telafi Tarama</div>
            <div class="kp-studio-sub">
                Öğrencinin ders ve kazanım bazlı performans röntgenini çıkarın, hata türlerini ayrıştırın
                ve tek tıkla Soru Fabrikası, Dijimentor veya Veli Paneli'ne aktarın.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    f_col1, f_col2, f_col3 = st.columns([2.2, 1.8, 1.0])
    with f_col1:
        sec_student = st.selectbox("Öğrenci Profili Seçin", list(OGRENCI_RONTGEN_PROFILLERI.keys()), key="xray_student_sel")
    with f_col2:
        sec_scope = st.selectbox("Röntgen Kapsamı", ["Son 5 Deneme Ortalaması", "Son 30 Günlük Soru Çözümü", "Tüm Dönem Kümülatif"], key="xray_scope_sel")
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
        st.bar_chart(prof["hatalar"], color="#2563EB", height=225)

    st.markdown("#### Dijimentor Otomatik Telafi Reçetesi")
    st.markdown(
        f"""
        <div class="kp-card">
            <div style="font-weight:700;color:#1E2952;font-size:0.92rem;">
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

    a_col1, a_col2, a_col3 = st.columns(3)
    with a_col1:
        if st.button("Kritik Kazanım İçin Soru Üret", key="xray_to_factory", type="primary", use_container_width=True):
            st.session_state["factory_focus_topic"] = prof["hedef_konu"]
            st.session_state["factory_branch"] = prof["hedef_brans"]
            st.session_state["factory_grade"] = prof["hedef_sinif"]
            st.session_state["fab_theme_inp"] = prof["hedef_konu"]
            st.session_state["factory_sync_pending"] = True
            st.session_state["factory_grade_sync"] = True
            record_activity("Röntgen -> Fabrika", f"{prof['hedef_konu']} Soru Fabrikasına aktarıldı")
            navigate_to("Soru & Redaksiyon Fabrikası")
    with a_col2:
        if st.button("Dijimentor Planına Aktar", key="xray_to_mentor", use_container_width=True):
            st.session_state["dk_weak"] = prof["recete_ozet"]
            st.session_state["dk_focus"] = "Dikkat ve Çeldirici Tuzağı Azaltma"
            record_activity("Röntgen -> Dijimentor", f"{sec_student} bulguları mentorluğa aktarıldı")
            navigate_to("Dijimentor")
    with a_col3:
        if st.button("Veli Özeti olarak Aktar", key="xray_to_parent", use_container_width=True):
            st.session_state["parent_student_stats"]["weak_topic"] = prof["hedef_konu"]
            save_persisted_state()
            record_activity("Röntgen -> Veli", f"{sec_student} kritik kazanımı Veli Paneline aktarıldı")
            navigate_to("Veli Takip Paneli")


def render_parent_notification_center() -> None:
    st.markdown("---")
    st.markdown("### Veli Bilgilendirme & Otomatik Mesaj Merkezi")
    st.caption(
        "Veli bilgilendirme başlıkları altında otomatik mesaj taslağı oluşturun, "
        "duyuru akışına ekleyin ve kurumsal DOCX olarak indirin."
    )

    stats = st.session_state["parent_student_stats"]
    solved_safe = max(1, int(stats["weekly_solved"]))
    acc_pct = round((int(stats["correct"]) / solved_safe) * 100, 1)

    t_msg, t_feed, t_profile = st.tabs(["Mesaj Üreticisi", "Veli Duyuru Akışı", "Öğrenci & Veli Verileri"])

    with t_msg:
        c_left, c_right = st.columns([1.05, 1.35])
        with c_left:
            sel_cat = st.selectbox("Ana Konu Başlığı", list(PARENT_MESSAGE_TOPICS.keys()), key="pm_cat")
            subtopics = PARENT_MESSAGE_TOPICS[sel_cat]["subtopics"]
            sel_sub = st.selectbox("Alt Konu Başlığı", subtopics, key="pm_sub")
            channel = st.radio("Mesaj Formatı", ["WhatsApp / Kurumsal Mesaj", "Kısa SMS"], horizontal=True, key="pm_ch")

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
                    f"{BRAND_NAME} 2028: Sn. {stats['parent_name']}, {stats['student_name']} "
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
            mail_url = f"mailto:{stats['parent_email']}?subject={BRAND_NAME}%202028%20-%20{sel_sub}&body={encoded}"

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
                file_name="veli_bilgilendirme_mesaji_2028.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    with t_feed:
        flt = st.selectbox("Kategoriye Göre Filtrele", ["Tümü"] + list(PARENT_MESSAGE_TOPICS.keys()), key="pm_flt")
        items = [a for a in st.session_state["parent_announcements"] if flt == "Tümü" or a["category"] == flt]
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

if "loader_nonce" not in st.session_state:
    trigger_loader_bar()

render_top_loader_bar()

if "generated_test_text" not in st.session_state:
    st.session_state.generated_test_text = ""
if "dijimentor_plan" not in st.session_state:
    st.session_state.dijimentor_plan = ""
if "selected_tier_id" not in st.session_state:
    st.session_state.selected_tier_id = _persisted.get("selected_tier_id", SUBSCRIPTION_TIERS[1]["id"])
if "backpack_cart" not in st.session_state and isinstance(_persisted.get("backpack_cart"), dict):
    st.session_state["backpack_cart"] = dict(_persisted["backpack_cart"])

ensure_bag_initialized()
bag_summary_global = get_bag_summary()

if "used_question_quota" not in st.session_state:
    if _persisted.get("used_question_quota") is not None:
        st.session_state["used_question_quota"] = int(_persisted["used_question_quota"])
    elif _persisted.get("remaining_question_quota") is not None:
        st.session_state["used_question_quota"] = max(
            0, int(bag_summary_global["monthly_quota"]) - int(_persisted["remaining_question_quota"])
        )
    else:
        st.session_state["used_question_quota"] = 0

if "used_coach_quota" not in st.session_state:
    if _persisted.get("used_coach_quota") is not None:
        st.session_state["used_coach_quota"] = int(_persisted["used_coach_quota"])
    elif _persisted.get("remaining_coach_quota") is not None:
        st.session_state["used_coach_quota"] = max(
            0, int(bag_summary_global["coach_quota"]) - int(_persisted["remaining_coach_quota"])
        )
    else:
        st.session_state["used_coach_quota"] = 0

sync_plan_quotas(bag_summary_global, reset_usage=False)

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
    st.session_state.parent_student_stats = merged_stats
if "parent_announcements" not in st.session_state:
    p_ann = _persisted.get("parent_announcements")
    st.session_state.parent_announcements = p_ann if isinstance(p_ann, list) and p_ann else _default_parent_announcements()
if "question_bank_2028" not in st.session_state:
    st.session_state["question_bank_2028"] = _persisted.get("question_bank_2028", []) if isinstance(_persisted.get("question_bank_2028"), list) else []
if "error_notebook_2028" not in st.session_state:
    st.session_state["error_notebook_2028"] = _persisted.get("error_notebook_2028", []) if isinstance(_persisted.get("error_notebook_2028"), list) else []
if "b2b_org_cache" not in st.session_state and isinstance(_persisted.get("b2b_org_cache"), dict):
    st.session_state["b2b_org_cache"] = _persisted["b2b_org_cache"]
if "last_b2b_summary" not in st.session_state and _persisted.get("last_b2b_summary"):
    st.session_state["last_b2b_summary"] = _persisted["last_b2b_summary"]

api_key_ready = resolve_gemini_api_key() is not None

# ==============================================================================
# KENAR ÇUBUĞU
# ==============================================================================
def _fmt(n: int) -> str:
    """1500 -> '1.500'"""
    return f"{int(n):,}".replace(",", ".")


with st.sidebar:
    st.markdown(
        f"""
        <div class="kp-sidebar-brand">
            <img src="{LOGO_DATA_URI}" alt=""/>
            <div>
                <div class="kp-sidebar-brand-title">{BRAND_SHORT}</div>
                <div class="kp-sidebar-brand-sub">Dijital Ofis</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Gruplanmış, ikonlu menü
    for group_title, items in NAV_GROUPS:
        if group_title:
            st.markdown(f'<div class="kp-nav-title">{html.escape(group_title)}</div>', unsafe_allow_html=True)
        for p_name, p_label, p_icon in items:
            is_current = st.session_state.active_page == p_name
            if st.button(
                p_label,
                key=f"nav_{p_name}",
                icon=p_icon,
                type="primary" if is_current else "secondary",
                use_container_width=True,
            ):
                clear_page_ephemeral_state(p_name)
                st.session_state.active_page = p_name
                trigger_loader_bar()
                st.rerun()

    # Paket & kota kartı (ilerleme çubuklu)
    total_q = int(bag_summary_global["monthly_quota"])
    total_c = int(bag_summary_global["coach_quota"])
    rem_q = int(st.session_state.remaining_question_quota)
    rem_c = int(st.session_state.remaining_coach_quota)
    q_pct = round(100 * rem_q / total_q) if total_q else 0
    c_pct = round(100 * rem_c / total_c) if total_c else 0
    status_color = "#16A34A" if api_key_ready else "#D97706"
    status_text = "Sistem aktif" if api_key_ready else "API anahtarı bekleniyor"

    st.markdown(
        f"""
        <div class="kp-sidebar-meta">
            <div class="kp-meta-head"><span>Aktif paketin</span></div>
            <div class="kp-meta-sub">{bag_summary_global['total_count']} paket · {_fmt(bag_summary_global['monthly_price'])} ₺/ay</div>
            <div class="kp-meter-label"><span>Soru kotası</span><span>{_fmt(rem_q)} / {_fmt(total_q)}</span></div>
            <div class="kp-meter"><span style="width:{q_pct}%"></span></div>
            <div class="kp-meter-label"><span>Dijimentor</span><span>{rem_c} / {total_c}</span></div>
            <div class="kp-meter"><span style="width:{c_pct}%"></span></div>
            <div class="kp-meter-label" style="margin-top:0.6rem;">
                <span>Soru havuzu: <b>{len(st.session_state['question_bank_2028'])}</b></span>
                <span>Hata defteri: <b>{len(st.session_state['error_notebook_2028'])}</b></span>
            </div>
            <div class="kp-status"><span class="kp-status-dot" style="background:{status_color}"></span>{status_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Tehlikeli işlem: göz önünden kaldırıldı, "Ayarlar" altında
    with st.expander("Ayarlar", icon=":material/settings:"):
        if not st.session_state.get("confirm_reset_session", False):
            if st.button("Oturumu sıfırla", key="reset_session_btn", use_container_width=True):
                st.session_state["confirm_reset_session"] = True
                st.rerun()
        else:
            st.warning("Yerel durum sıfırlansın mı? (Yedek alınır)")
            rc1, rc2 = st.columns(2)
            with rc1:
                if st.button("Evet", key="confirm_reset_yes", type="primary", use_container_width=True):
                    try:
                        if STATE_FILE_PATH.exists():
                            STATE_BACKUP_PATH.write_text(STATE_FILE_PATH.read_text(encoding="utf-8"), encoding="utf-8")
                            STATE_FILE_PATH.unlink()
                    except Exception:
                        pass
                    for k in list(st.session_state.keys()):
                        del st.session_state[k]
                    st.rerun()
            with rc2:
                if st.button("İptal", key="confirm_reset_no", use_container_width=True):
                    st.session_state["confirm_reset_session"] = False
                    st.rerun()

# ==============================================================================
# ÜST ŞERİT
# ==============================================================================
st.markdown(
    """
    <div class="kp-topbar">
        <span class="kp-topbar-brand">Kurşun Kalem Dijital Ofis</span>
        <span>Destek: <a href="tel:08505321629">0850 532 16 29</a> · <a href="mailto:destek@kursunkalemdijitalofis.com.tr">destek@kursunkalemdijitalofis.com.tr</a></span>
    </div>
    """,
    unsafe_allow_html=True,
)

current_page = st.session_state.active_page
if current_page != "Ana Sayfa":
    render_page_strip(current_page)

# ==============================================================================
# 1. ANA SAYFA
# ==============================================================================
if current_page == "Ana Sayfa":
    stats_now = st.session_state["parent_student_stats"]
    solved_now = max(1, int(stats_now["weekly_solved"]))
    acc_now = round(int(stats_now["correct"]) / solved_now * 100)

    # ---------- HERO ----------
    hero_l, hero_r = st.columns([1.2, 1.0], gap="large", vertical_alignment="center")

    with hero_l:
        st.markdown(
            """
            <span class="kp-eyebrow">2028 müfredatına tam uyumlu</span>
            <h1 class="kp-hero-title">Hayalindeki <em>2028 hedefi</em> için ihtiyacın olan her şey, tek platformda.</h1>
            <p class="kp-hero-lead">
                Soru bankası, haftalık dijital mentorluk, mini denemeler ve kişisel eksik analizi.
                Öğrenci, veli ve okul aynı ekranda.
            </p>
            """,
            unsafe_allow_html=True,
        )
        cta1, cta2, _ = st.columns([1, 1.15, 0.6])
        with cta1:
            if st.button("Ücretsiz dene", key="hero_cta_primary", type="primary", use_container_width=True):
                navigate_to("Paketler & Abonelik Vitrini")
        with cta2:
            if st.button("Kurumlar için demo", key="hero_cta_secondary", use_container_width=True):
                navigate_to("B2B Lisans & Kurum")
        st.markdown(
            """
            <div class="kp-stats">
                <div><div class="kp-stat-num">142.500+</div><div class="kp-stat-lbl">soru havuzu</div></div>
                <div><div class="kp-stat-num">18.420</div><div class="kp-stat-lbl">ön kayıtlı öğrenci ve veli</div></div>
                <div><div class="kp-stat-num">48</div><div class="kp-stat-lbl">anlaşmalı kurum</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with hero_r:
        # Eski SVG vitrin yerine öğrencinin gerçek haftalık özeti
        st.markdown(
            f"""
            <div class="kp-preview">
                <div class="kp-preview-head">
                    <div><small>Bu hafta</small><strong>{html.escape(str(stats_now['student_name']))}</strong></div>
                    <span class="kp-pill">{html.escape(str(stats_now['grade']))}</span>
                </div>
                <div class="kp-preview-stats">
                    <div class="kp-preview-stat"><b>{_fmt(stats_now['weekly_solved'])}</b><span>çözülen soru</span></div>
                    <div class="kp-preview-stat"><b>%{acc_now}</b><span>doğruluk</span></div>
                    <div class="kp-preview-stat"><b>%{int(stats_now['mentor_completion'])}</b><span>plan uyumu</span></div>
                </div>
                <div class="kp-meter" style="margin-bottom:0.9rem;"><span style="width:{int(stats_now['mentor_completion'])}%"></span></div>
                <div class="kp-preview-row"><span>Güçlü alan</span><span>{html.escape(str(stats_now['strong_topic']))}</span></div>
                <div class="kp-preview-row"><span>Pekiştirilecek</span><span>{html.escape(str(stats_now['weak_topic']))}</span></div>
                <div class="kp-preview-row"><span>Sıradaki deneme</span><span>{html.escape(str(stats_now['next_exam_date']))}</span></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- SINIF SEÇİCİ ----------
    with st.container(border=True):
        sel_l, sel_r = st.columns([1.3, 1.0], vertical_alignment="center")
        with sel_l:
            st.markdown(
                '<h2 class="kp-section-title">Sınıfını ya da sınavını seç</h2>'
                '<p class="kp-section-sub" style="margin-bottom:0;">Sana uygun modül otomatik açılır.</p>',
                unsafe_allow_html=True,
            )
        with sel_r:
            group_names = list(GRADE_GROUPS.keys())
            selected_group = st.segmented_control(
                "Kademe",
                options=group_names,
                default=group_names[0],
                key="grade_group_sel",
                label_visibility="collapsed",
            ) or group_names[0]

        group_items = GRADE_GROUPS[selected_group]
        g_idx = group_names.index(selected_group)
        chip_cols = st.columns(4)
        for idx, chip_label in enumerate(group_items):
            chip = CHIP_BY_LABEL[chip_label]
            with chip_cols[idx % 4]:
                if st.button(chip["label"], key=f"grade_chip_{g_idx}_{idx}", use_container_width=True):
                    st.session_state["factory_grade"] = chip["grade"]
                    st.session_state["factory_grade_sync"] = True
                    st.session_state["dk_exam"] = chip["exam"]
                    record_activity("Hızlı Kademe (2028)", f"{chip['label']} seçildi")
                    navigate_to(chip["page"])

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- MODÜLLER ----------
    st.markdown(
        '<h2 class="kp-section-title">Modüller</h2>'
        '<p class="kp-section-sub">Her biri tek başına çalışır, birlikte daha güçlüdür.</p>',
        unsafe_allow_html=True,
    )
    mod_cols = st.columns(len(MODULE_CARDS))
    for idx, (code, m_title, m_desc, m_page, m_bg, m_fg) in enumerate(MODULE_CARDS):
        with mod_cols[idx]:
            with st.container(border=True):
                st.markdown(
                    f"""
                    <div class="kp-mod-icon" style="background:{m_bg};color:{m_fg};">{html.escape(code)}</div>
                    <div class="kp-mod-title">{html.escape(m_title)}</div>
                    <div class="kp-mod-desc">{html.escape(m_desc)}</div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Aç →", key=f"mod_open_{idx}"):
                    navigate_to(m_page)

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- ÖNCELİKLİ EKSİKLER ----------
    pr_l, pr_r = st.columns([4, 1], vertical_alignment="bottom")
    with pr_l:
        st.markdown(
            '<h2 class="kp-section-title">Öncelikli eksiklerin</h2>'
            '<p class="kp-section-sub">En çok puan kazandıracak konular, öncelik sırasıyla.</p>',
            unsafe_allow_html=True,
        )
    with pr_r:
        if st.button("Tüm analizi gör →", key="prio_all"):
            navigate_to("Öğrenci Röntgeni")

    with st.container(border=True):
        for idx, p_item in enumerate(ONCELIKLI_EKSIK_SIRASI):
            if idx:
                st.markdown('<div class="kp-prio-divider"></div>', unsafe_allow_html=True)
            c_num, c_txt, c_meter, c_btn = st.columns([0.45, 5, 2.2, 1.2], vertical_alignment="center")
            with c_num:
                st.markdown(f'<div class="kp-prio-num">{idx + 1}</div>', unsafe_allow_html=True)
            with c_txt:
                st.markdown(
                    f'<div class="kp-prio-title">{html.escape(p_item["topic"])}</div>'
                    f'<div class="kp-prio-sub">{html.escape(p_item["branch"])} · {html.escape(p_item["action"])}</div>',
                    unsafe_allow_html=True,
                )
            with c_meter:
                pct = p_item.get("pct")
                if pct is not None:
                    bar_color = "#DC2626" if pct < 50 else "#EA580C" if pct < 60 else "#CA8A04"
                    st.markdown(
                        f'<div class="kp-meter-label"><span>Doğru oranı</span><b style="color:#0F172A">%{pct}</b></div>'
                        f'<div class="kp-meter"><span style="width:{pct}%;background:{bar_color}"></span></div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown('<span class="kp-tag-warn">Dikkat gerektiriyor</span>', unsafe_allow_html=True)
            with c_btn:
                if st.button("Çalış", key=f"prio_btn_{idx}", use_container_width=True):
                    st.session_state["factory_focus_topic"] = p_item["topic"]
                    st.session_state["factory_branch"] = p_item["branch"]
                    st.session_state["factory_grade"] = p_item["grade"]
                    st.session_state["fab_theme_inp"] = p_item["topic"]
                    st.session_state["factory_sync_pending"] = True
                    st.session_state["factory_grade_sync"] = True
                    record_activity("Öncelikli Akış (2028)", f"{p_item['topic']} seçildi")
                    navigate_to("Soru & Redaksiyon Fabrikası")

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- GERİ SAYIM ----------
    head_col, toggle_col = st.columns([4, 1], vertical_alignment="center")
    with head_col:
        st.markdown('<h2 class="kp-section-title">2028 LGS\'ye kalan süre</h2>', unsafe_allow_html=True)
    with toggle_col:
        live_counter = st.toggle("Canlı sayaç", value=st.session_state.live_tick)
        st.session_state.live_tick = live_counter

    counter_placeholder = st.empty()
    with counter_placeholder.container():
        render_countdown_boxes(LGS_2028_TARGET)

    yks_d, yks_h, _, _ = format_countdown(YKS_2028_TARGET)
    st.caption(f"LGS: 11 Haziran 2028 · YKS (17 Haziran 2028) için kalan süre: {yks_d} gün {yks_h} saat")

    st.markdown('<div class="kp-gap"></div>', unsafe_allow_html=True)

    # ---------- TELEMETRİ & AKTİVİTE ----------
    ana_l, ana_r = st.columns([2, 1], gap="large")
    with ana_l:
        st.markdown('<h2 class="kp-section-title" style="font-size:1.15rem;">Haftalık soru üretimi</h2>', unsafe_allow_html=True)
        telemetry_data = {"Pzt": 16800, "Sal": 18200, "Çar": 19500, "Per": 21100, "Cum": 22400, "Cmt": 25900, "Paz": 24800}
        st.area_chart(telemetry_data, color="#1D4ED8", height=210)
    with ana_r:
        st.markdown('<h2 class="kp-section-title" style="font-size:1.15rem;">Son aktiviteler</h2>', unsafe_allow_html=True)
        for item in st.session_state.get("activity_log", [])[:5]:
            st.markdown(
                f"""
                <div style="padding:0.6rem 0;border-bottom:1px solid #E2E8F0;">
                    <div style="font-size:0.78rem;color:#64748B;">{html.escape(item['time'])} · {html.escape(item['module'])}</div>
                    <div style="font-size:0.9rem;color:#0F172A;margin-top:0.1rem;">{html.escape(item['summary'])}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    if st.session_state.live_tick:
        time.sleep(1)
        st.rerun()

# ==============================================================================
# 2. SORU & REDAKSİYON FABRİKASI (5–12. SINIF MÜFREDAT + KADEMELİ SERİ + HAVUZ)
# ==============================================================================
elif current_page == "Soru & Redaksiyon Fabrikası":
    tab_gen, tab_bank = st.tabs(["Soru Üretim & Redaksiyon Stüdyosu", f"Soru Havuzu ({len(st.session_state['question_bank_2028'])})"])

    with tab_gen:
        if st.session_state.get("factory_grade") in GRADE_LEVELS_2028 and st.session_state.get("factory_grade_sync"):
            st.session_state["fab_grade_sel"] = st.session_state["factory_grade"]
            st.session_state["factory_grade_sync"] = False

        cur_grade_init = st.session_state.get("fab_grade_sel", "8. Sınıf (LGS)")
        if cur_grade_init not in GRADE_LEVELS_2028:
            cur_grade_init = "8. Sınıf (LGS)"

        if st.session_state.get("factory_sync_pending"):
            st.session_state["fab_grade_sel"] = st.session_state.get("factory_grade", cur_grade_init)
            g_tmp = st.session_state["fab_grade_sel"]
            b_list = get_branches_for_grade_2028(g_tmp)
            target_b = st.session_state.get("factory_branch")
            st.session_state["fab_branch_sel"] = target_b if target_b in b_list else b_list[0]
            a_list = get_areas_for_grade_branch_2028(g_tmp, st.session_state["fab_branch_sel"])
            st.session_state["fab_area_sel"] = a_list[0]
            o_list = get_outcomes_for_2028(g_tmp, st.session_state["fab_branch_sel"], a_list[0])
            focus_t = st.session_state.get("factory_focus_topic")
            if focus_t and focus_t not in o_list:
                st.session_state["fab_custom_outcome_chk"] = True
                st.session_state["fab_custom_outcome_inp"] = focus_t
            else:
                st.session_state["fab_outcome_sel"] = focus_t if focus_t in o_list else o_list[0]
            if focus_t:
                st.session_state["fab_theme_inp"] = focus_t
            st.session_state["factory_sync_pending"] = False

        if st.session_state.get("factory_focus_topic"):
            st.info(f"Aktarılan odak konu: **{st.session_state['factory_focus_topic']}**")

        with st.expander("5–12. Sınıf Müfredat, Zorluk & Kalite Parametreleri (2028)", expanded=True):
            r1c1, r1c2, r1c3 = st.columns(3)
            with r1c1:
                institution = st.text_input("Kurum / Ofis Adı", value=BRAND_NAME, key="fab_inst_inp")
                grade = st.selectbox("Sınıf Seviyesi (5–12)", GRADE_LEVELS_2028, key="fab_grade_sel")
                branches = get_branches_for_grade_2028(grade)
                if st.session_state.get("fab_branch_sel") not in branches:
                    st.session_state["fab_branch_sel"] = branches[0]
                branch = st.selectbox("Branş Seçimi", branches, key="fab_branch_sel")
            with r1c2:
                areas = get_areas_for_grade_branch_2028(grade, branch)
                if st.session_state.get("fab_area_sel") not in areas:
                    st.session_state["fab_area_sel"] = areas[0]
                learning_area = st.selectbox("Öğrenme Alanı / Tema", areas, key="fab_area_sel")

                use_custom_outcome = st.checkbox("Özel / Manuel Kazanım veya Konu Gir", key="fab_custom_outcome_chk")
                if use_custom_outcome:
                    outcome = st.text_input(
                        "Özel Kazanım / Zümre Konusu",
                        value=st.session_state.get("fab_custom_outcome_inp", ""),
                        placeholder="Örn: Paragrafta yardımcı düşünce & çeldirici kurgusu",
                        key="fab_custom_outcome_inp",
                    ) or "Özel Kazanım Odaklı Soru Üretimi"
                else:
                    outcomes = get_outcomes_for_2028(grade, branch, learning_area)
                    if st.session_state.get("fab_outcome_sel") not in outcomes:
                        st.session_state["fab_outcome_sel"] = outcomes[0]
                    outcome = st.selectbox("Hedef Kazanım", outcomes, key="fab_outcome_sel")

                if "fab_theme_inp" not in st.session_state:
                    st.session_state["fab_theme_inp"] = st.session_state.get("factory_focus_topic") or "2028 DİJİTAL OFİS & HAYAT"
                context_theme = st.text_input("Bağlam / Tema / Senaryo", key="fab_theme_inp")
            with r1c3:
                production_mode = st.radio(
                    "Üretim Modu",
                    ["Standart Test Paketi", "Kademeli Zorluk Serisi (Kolay → Orta → Zor)"],
                    key="fab_prod_mode",
                )
                question_count = st.slider("Soru Adedi", 1, 10, 3, key="fab_q_cnt")
                difficulty = st.select_slider(
                    "Temel Zorluk Düzeyi",
                    options=["Kolay", "Orta", "Orta-Zor", "Zor (2028 LGS/ÖSYM)"],
                    value="Orta-Zor",
                    key="fab_diff_sld",
                )
                distractor_mode = st.selectbox(
                    "Çeldirici & Kök Stratejisi",
                    [
                        "Muhakeme & Çeldirici Tuzağı Dengeli",
                        "Olumsuz Kök ('Değildir / Ulaşılamaz') Odaklı",
                        "Öncüllü & Bağlamsal Analiz Odaklı",
                    ],
                    index=0,
                    key="fab_dist_sel",
                )
                auto_review = st.checkbox("Çift Katmanlı Redaksiyon & Çeldirici Kalite Denetimi", value=True, key="fab_rev_chk")

        if st.button("Kurumsal Soru Paketi Üret (2028)", type="primary", key="fab_gen_btn"):
            quota_ok, quota_msg = consume_quota("question", int(question_count))
            if not quota_ok:
                st.error(quota_msg)
            else:
                mode_instruction = (
                    f"Her soruyu '{difficulty}' seviyesinde üret."
                    if production_mode == "Standart Test Paketi"
                    else "Soruları Kolay (pekiştirme), Orta (kazanım) ve Zor (2028 LGS/YKS muhakeme) olarak kademelendir."
                )
                sys_prompt = (
                    "Sen Türkiye Yüzyılı Maarif Modeli, 2028 LGS ve 2028 YKS ölçme-değerlendirme standartlarında "
                    "uzman bir kıdemli test editörüsün. Soruların özgün, dilbilgisi ve noktalama açısından "
                    "kusursuz, çeldiricilerinin uzunluk ve dilbilgisi paralelliği açısından dengeli olmasını sağla. "
                    "Her soruyu 'Soru 1', 'Soru 2'... şeklinde numaralandır ve şıkları 'A)', 'B)', 'C)', 'D)' "
                    "olarak yaz. Her sorunun altında kısa bir [Çözüm & Şık Gerekçesi] ve redaksiyon açıksa "
                    "[Redaksiyon & Çeldirici Kalite Notu] ekle; en sonda mutlaka 'Cevap Anahtarı' başlığı altında "
                    "'1-A', '2-C' formatında net anahtar ver."
                )
                user_prompt = (
                    f"Kurum: {institution} (2028)\nSınıf Seviyesi: {grade}\nBranş: {branch}\n"
                    f"Öğrenme Alanı: {learning_area}\nHedef Kazanım: {outcome}\n"
                    f"Bağlam / Tema / Senaryo: {context_theme}\nZorluk Kurgusu: {mode_instruction}\n"
                    f"Çeldirici & Kök Stratejisi: {distractor_mode}\nSoru Adedi: {question_count}\n"
                    f"Redaksiyon & Kalite Denetimi: {'Aktif' if auto_review else 'Pasif'}\n\n"
                    "Lütfen numaralandırılmış çoktan seçmeli (Soru 1, Soru 2... ve A, B, C, D şıkları) özgün sorular üret."
                )
                skel_slot = st.empty()
                render_loading_skeleton(skel_slot, lines=4)
                try:
                    result = call_gemini_text(prompt=user_prompt, system_instruction=sys_prompt)
                    skel_slot.empty()
                    st.session_state.generated_test_text = result
                    st.session_state["last_fab_meta"] = {
                        "grade": grade,
                        "branch": branch,
                        "outcome": outcome,
                        "difficulty": difficulty if production_mode == "Standart Test Paketi" else "Kademeli Seri",
                    }
                    record_activity("Soru Fabrikası (2028)", f"{grade} • {branch} • {question_count} soru üretildi")
                    st.success(
                        f"Soru paketi oluşturuldu. Kalan soru kotası: "
                        f"{int(st.session_state.remaining_question_quota):,}".replace(",", ".")
                    )
                except Exception as exc:
                    skel_slot.empty()
                    refund_quota("question", int(question_count))
                    st.error(f"Soru üretim hatası oluştu: {exc}")

        if st.session_state.generated_test_text:
            st.markdown("---")
            st.markdown("### Üretilen Test, Çözüm Gerekçeleri ve Redaksiyon Çıktısı")
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
            act_c1, act_c2, act_c3, act_c4 = st.columns(4)
            with act_c1:
                st.download_button(
                    label="Testi DOCX İndir",
                    data=docx_bytes,
                    file_name=f"kursun_kalem_ofis_2028_{branch.lower().replace(' ', '_')}_test.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            with act_c2:
                if st.button("Soru Havuzuna Kaydet", use_container_width=True):
                    add_to_question_bank(
                        grade=grade,
                        branch=branch,
                        outcome=outcome,
                        difficulty=difficulty,
                        content_text=st.session_state.generated_test_text,
                    )
                    record_activity("Soru Havuzu", f"{grade} • {branch} testi havuza kaydedildi")
                    st.toast("Test Soru Havuzuna kaydedildi.")
                    st.rerun()
            with act_c3:
                if st.button("Mini Deneme Olarak Çöz", use_container_width=True):
                    st.session_state["portal_exam_source"] = st.session_state.generated_test_text
                    record_activity("Köprü", "Test Öğrenci Portalı mini denemesine aktarıldı")
                    navigate_to("Öğrenci Portalı & Çözüm")
            with act_c4:
                if st.button("İlk Soruyu Çözüm Rehberine Gönder", use_container_width=True):
                    parsed_q = parse_quiz_questions(st.session_state.generated_test_text)
                    st.session_state["portal_question_draft"] = (
                        parsed_q[0]["full_text"] if parsed_q else st.session_state.generated_test_text[:1000]
                    )
                    navigate_to("Öğrenci Portalı & Çözüm")

    with tab_bank:
        bank_items: List[Dict[str, str]] = st.session_state.get("question_bank_2028", [])
        if not bank_items:
            st.info("Soru havuzunda henüz kayıtlı test bulunmuyor. Üretim sekmesinden **Soru Havuzuna Kaydet** butonunu kullanabilirsiniz.")
        else:
            combined_bank_text = "\n\n".join(
                [f"=== [{b['grade']} | {b['branch']} | {b['difficulty']}] {b['outcome']} ===\n{b['content']}" for b in bank_items]
            )
            bundle_docx = build_test_docx(
                institution=BRAND_NAME,
                branch="Karma Zümre Havuzu",
                grade="5–12. Sınıf",
                learning_area="Kayıtlı Soru Havuzu",
                outcome="Havuzdaki Seçili Kazanımlar",
                context_theme="2028 Kurumsal Soru Kitapçığı",
                generated_content=combined_bank_text,
            )
            st.download_button(
                label=f"Tüm Havuzu ({len(bank_items)} Kayıt) Toplu DOCX İndir",
                data=bundle_docx,
                file_name="kursun_kalem_2028_soru_havuzu.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            for b_item in bank_items:
                with st.expander(f"[{b_item['id']}] {b_item['grade']} • {b_item['branch']} ({b_item['difficulty']}) — {b_item['date']}"):
                    st.caption(f"Kazanım: {b_item['outcome']}")
                    st.markdown(b_item["content"])
                    if st.button("Bu Testi Portala Aktar", key=f"load_bank_{b_item['id']}"):
                        st.session_state["portal_exam_source"] = b_item["content"]
                        navigate_to("Öğrenci Portalı & Çözüm")

# ==============================================================================
# 3. ÖĞRENCİ PORTALI & ÇÖZÜM (+ SÜRE TAKİBİ + HATA DEFTERİ)
# ==============================================================================
elif current_page == "Öğrenci Portalı & Çözüm":
    tab_exam, tab_hint, tab_err = st.tabs([
        "İnteraktif Mini Deneme & Süre Analizi",
        "Adımlı Çözüm Asistanı",
        f"Hata Defteri ({len(st.session_state['error_notebook_2028'])})",
    ])

    with tab_exam:
        exam_source = st.text_area(
            "Deneme Soru Metni (Soru Fabrikasından otomatik gelir veya buraya yapıştır)",
            value=st.session_state.get("portal_exam_source") or st.session_state.get("generated_test_text", ""),
            height=120,
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
            if "exam_start_ts" not in st.session_state:
                st.session_state["exam_start_ts"] = time.time()

            st.markdown(f"#### Aktif Mini Deneme ({len(questions)} Soru)")
            st.caption("Boş bırakmak isterseniz **Boş Bırak** seçeneğini işaretleyebilir veya seçim yapmayabilirsiniz; boş sorular yanlışa sayılmaz.")
            user_answers: Dict[int, Optional[str]] = {}
            q_durations: Dict[int, float] = {}

            for idx, q in enumerate(questions):
                with st.container(border=True):
                    st.markdown(q["stem"])
                    c_opt, c_sec = st.columns([3.4, 1.0])
                    with c_opt:
                        opts = list(q["options"].keys()) + ["BOŞ"]
                        chosen = st.radio(
                            f"Soru {q['num']} Yanıtın",
                            options=opts,
                            index=None,
                            format_func=lambda k, o_map=q["options"]: "Boş Bırak" if k == "BOŞ" else f"{k}) {o_map[k]}",
                            horizontal=True,
                            key=f"mini_q_{q['num']}_{idx}",
                        )
                        user_answers[idx] = None if chosen in (None, "BOŞ") else chosen
                    with c_sec:
                        q_durations[idx] = float(
                            st.number_input(
                                f"Süre (sn) • S{q['num']}",
                                min_value=5,
                                max_value=300,
                                value=65,
                                step=5,
                                key=f"mini_dur_{q['num']}_{idx}",
                            )
                        )

            if st.button("Denemeyi Bitir ve Netimi Hesapla", type="primary"):
                correct_cnt = 0
                wrong_cnt = 0
                blank_cnt = 0
                ungraded_cnt = 0
                wrong_questions: List[Dict[str, Any]] = []
                result_rows: List[Dict[str, str]] = []
                total_sec = 0.0

                for idx, q in enumerate(questions):
                    u_ans = user_answers[idx]
                    c_ans = q.get("correct")
                    sec_val = q_durations.get(idx, 60.0)
                    total_sec += sec_val

                    if u_ans is None:
                        blank_cnt += 1
                        status_txt = "Boş"
                    elif not c_ans:
                        ungraded_cnt += 1
                        status_txt = "Anahtar Yok"
                    elif u_ans == c_ans:
                        correct_cnt += 1
                        status_txt = "Doğru"
                    else:
                        wrong_cnt += 1
                        wrong_questions.append(q)
                        status_txt = "Yanlış"
                        add_to_error_notebook(
                            question_num=q["num"],
                            question_text=q["full_text"],
                            user_ans=u_ans,
                            correct_ans=c_ans,
                            elapsed_sec=sec_val,
                        )

                    result_rows.append(
                        {
                            "Soru": f"Soru {q['num']}",
                            "Senin Yanıtın": u_ans or "Boş",
                            "Doğru Cevap": c_ans or "Belirtilmedi",
                            "Süre (sn)": f"{int(sec_val)} sn",
                            "Durum": status_txt,
                        }
                    )

                graded_total = correct_cnt + wrong_cnt
                net_score = max(0.0, correct_cnt - (wrong_cnt / 3.0))
                accuracy = (correct_cnt / graded_total * 100.0) if graded_total > 0 else 0.0
                avg_sec = total_sec / max(1, len(questions))

                st.session_state["last_exam_report"] = {
                    "correct": correct_cnt,
                    "wrong": wrong_cnt,
                    "blank": blank_cnt,
                    "net": net_score,
                    "accuracy": accuracy,
                    "avg_sec": avg_sec,
                    "rows": result_rows,
                    "first_wrong": wrong_questions[0]["full_text"] if wrong_questions else "",
                }
                record_activity(
                    "Mini Deneme",
                    f"{correct_cnt}D {wrong_cnt}Y {blank_cnt}B • {net_score:.2f} Net • Ort {avg_sec:.0f} sn/soru",
                )

            report = st.session_state.get("last_exam_report")
            if report:
                st.markdown("---")
                r1, r2, r3, r4 = st.columns(4)
                r1.metric("Doğru / Boş", f"{report['correct']}D / {report['blank']}B")
                r2.metric("Yanlış Sayısı", f"{report['wrong']} Soru")
                r3.metric("Hesaplanan Net (3Y=1D)", f"{report['net']:.2f} Net")
                r4.metric("Ort. Soru Süresi", f"{report['avg_sec']:.0f} sn (%{report['accuracy']:.1f})")

                st.dataframe(report["rows"], use_container_width=True, hide_index=True)

                if report.get("first_wrong"):
                    wa_c1, wa_c2 = st.columns(2)
                    with wa_c1:
                        if st.button("Yanlış Soruyu Çözüm Rehberine Aktar", use_container_width=True):
                            st.session_state["portal_question_draft"] = report["first_wrong"]
                            st.toast("Yanlış soru Adımlı Çözüm Asistanı sekmesine aktarıldı.")
                            st.rerun()
                    with wa_c2:
                        if st.button("Yanlış Sorudan Benzer Soru Üret", type="primary", use_container_width=True):
                            snippet = re.sub(r"\s+", " ", report["first_wrong"])[:110]
                            st.session_state["factory_focus_topic"] = f"Telafi: {snippet}"
                            st.session_state["fab_theme_inp"] = f"Telafi: {snippet}"
                            record_activity("Telafi Akışı", "Yanlış sorudan benzer soru üretimine geçildi")
                            navigate_to("Soru & Redaksiyon Fabrikası")

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
                skel_hint = st.empty()
                render_loading_skeleton(skel_hint, lines=3)
                try:
                    ans = call_gemini_text(prompt=prompt, system_instruction=sys_inst)
                    skel_hint.empty()
                    record_activity("Çözüm Asistanı", f"{hint_mode} üretildi")
                    st.markdown("#### Çözüm Analizi")
                    st.markdown(ans)
                except Exception as exc:
                    skel_hint.empty()
                    st.error(f"Çözüm üretilirken hata oluştu: {exc}")

    with tab_err:
        err_list: List[Dict[str, Any]] = st.session_state.get("error_notebook_2028", [])
        if not err_list:
            st.info("Hata defterinde henüz kayıt bulunmuyor. Mini denemede yanlış işaretlenen sorular otomatik buraya eklenir.")
        else:
            st.caption("Aralıklı tekrar takvimi: 1. Gün • 3. Gün • 7. Gün")
            for item in err_list:
                with st.container(border=True):
                    st.markdown(
                        f"**[{item['id']}] Soru {item['num']}** • Teşhis: **{item['diagnosis']}** • "
                        f"Senin: `{item['user_ans']}` / Doğru: `{item['correct_ans']}` • Süre: `{item['elapsed_sec']} sn`"
                    )
                    st.caption(
                        f"Tekrar Hatırlatmaları → 1. Gün: {item['review_1d']} | "
                        f"3. Gün: {item['review_3d']} | 7. Gün: {item['review_7d']}"
                    )
                    st.markdown(item["question_text"])
                    if st.button("Çözüm Rehberinde Aç", key=f"err_solve_{item['id']}"):
                        st.session_state["portal_question_draft"] = item["question_text"]
                        st.toast("Soru Adımlı Çözüm Asistanına aktarıldı.")

# ==============================================================================
# 4. DİJİMENTOR (2028 + AŞAMALI KAMP PLANI)
# ==============================================================================
elif current_page == "Dijimentor":
    st.markdown(
        """
        <div class="kp-studio-banner">
            <div class="kp-studio-title">Dijimentor (2028) — Akıllı Sınav, Faz Planı &amp; Performans Mentorluğu</div>
            <div class="kp-studio-sub">
                2028 hedef sınavına, haftalık çalışma kapasitene ve çeldirici/zaman analizine göre
                aşamalı kamp ve haftalık aksiyon planı üretir.
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
            <div class="kp-kpi-value">Dijimentor 2028</div>
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
            <div class="kp-kpi-label">2028 Kamp Fazları</div>
            <div class="kp-kpi-value">3 Aşama</div>
            <div class="kp-kpi-sub">Kazanım • Branş Hızı • Genel Deneme</div>
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
    st.markdown("#### Hızlı 2028 Mentorluk Senaryoları")
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
                st.session_state["fab_theme_inp"] = preset_data["topic"]
                st.session_state["factory_sync_pending"] = True
                st.session_state["factory_grade_sync"] = True
                record_activity("Dijimentor Yönlendirme", f"{preset_title} seçildi ve Soru Fabrikasına yönlendirildi")
                navigate_to("Soru & Redaksiyon Fabrikası")
    with p_cols[-1]:
        if st.button("Planı Temizle", key="clear_dk_plan", use_container_width=True):
            st.session_state.dijimentor_plan = ""
            st.session_state["dk_active_preset"] = ""
            st.rerun()

    st.markdown("---")
    st.markdown("#### Öğrenci 2028 Hedef ve Çalışma Profili")

    exam_opts = ["2028 LGS", "2028 TYT - AYT", "5-7. Sınıf Akademik Gelişim"]
    stored_exam = st.session_state.get("dk_exam", "2028 LGS")
    if stored_exam not in exam_opts:
        stored_exam = "2028 LGS"

    r1_a, r1_b, r1_c = st.columns(3)
    with r1_a:
        target_exam = st.selectbox(
            "Hedef Sınav / Kademe",
            exam_opts,
            index=exam_opts.index(stored_exam),
            key="dk_exam_sel",
        )
        st.session_state["dk_exam"] = target_exam

        weekly_hours = st.slider("Haftalık Çalışma Süresi (Saat)", min_value=4, max_value=45, value=current_hours, key="dk_hours_sld")
        st.session_state["dk_hours"] = weekly_hours

    with r1_b:
        target_goal = st.text_input(
            "Hedef Puan / Net / Okul Bandı",
            value=st.session_state.get("dk_goal", "Nitelikli lise / yüksek net ve sıfıra yakın dikkat hatası"),
            key="dk_goal_inp",
        )
        st.session_state["dk_goal"] = target_goal

        focus_opts = [
            "Haftalık Dengeli Çalışma Programı",
            "Deneme Net Artırma & Zaman Yönetimi",
            "Dikkat ve Çeldirici Tuzağı Azaltma",
            "Sınav Kaygısı & Motivasyon Rutini",
        ]
        stored_focus = st.session_state.get("dk_focus", focus_opts[0])
        if stored_focus not in focus_opts:
            stored_focus = focus_opts[0]
        coach_focus = st.selectbox(
            "Mentorluk Odak Alanı",
            focus_opts,
            index=focus_opts.index(stored_focus),
            key="dk_focus_sel",
        )
        st.session_state["dk_focus"] = coach_focus

    with r1_c:
        weak_topics = st.text_area(
            "Zorlandığın Dersler / Kazanımlar",
            value=st.session_state.get("dk_weak", "Sözel mantık ve muhakeme, üslü ifadeler, dikkat kaynaklı işlem hataları"),
            height=82,
            key="dk_weak_area",
        )
        st.session_state["dk_weak"] = weak_topics

        recent_score = st.text_input(
            "Mevcut Deneme / Süre Durumu",
            value=st.session_state.get("dk_status", "Orta seviye, süreyi yetiştirmede zorlanıyorum"),
            key="dk_status_inp",
        )
        st.session_state["dk_status"] = recent_score

    q_hrs = round(weekly_hours * 0.45, 1)
    r_hrs = round(weekly_hours * 0.30, 1)
    a_hrs = round(weekly_hours * 0.25, 1)
    st.markdown(
        f"""
        <div style="margin-top:0.65rem;padding-top:0.75rem;border-top:1px solid #E2E8F0;">
            <div style="font-size:0.75rem;font-weight:600;color:#475569;text-transform:uppercase;letter-spacing:0.04em;">
                Önerilen Haftalık Kapasite Dağılımı ({weekly_hours} Saat)
            </div>
            <div class="kp-dist-bar">
                <div style="width:45%;background:#2563EB;"></div>
                <div style="width:30%;background:#5A52E8;"></div>
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
    include_routine = st.checkbox("2028 Aşamalı Kamp Fazları (Kazanım → Branş → Genel Deneme) ve hafta sonu kontrol listesi dahil et", value=True)

    if st.button("Dijimentor 2028 Planı Oluştur", type="primary"):
        quota_ok, quota_msg = consume_quota("coach", 1)
        if not quota_ok:
            st.error(quota_msg)
        else:
            sys_coach = (
                "Sen 'Dijimentor' adlı, 2028 sınav dönemine hazırlanan öğrenciler için çalışan üst düzey, "
                "net, empatik ve veri odaklı bir dijital eğitim mentorusun. Yanıtını başlıklarla, "
                "madde işaretleriyle ve uygulanabilir zaman bloklarıyla kurumsal bir mentorluk "
                "raporu gibi düzenle. Şu bölümleri net oluştur:\n"
                "1. Teşhis ve 2028 Aşamalı Kamp Yol Haritası (Faz 1 Kazanım, Faz 2 Branş Hızı, Faz 3 Genel Deneme)\n"
                "2. Pazartesi–Pazar Haftalık Çalışma Planı (saat ve odak ders dağılımı)\n"
                "3. Çeldirici & Zaman Yönetimi Protokolü\n"
                "4. Haftalık Başarı KPI'ları ve Kontrol Listesi"
            )
            prompt_coach = (
                f"Hedef Sınav / Kademe: {target_exam}\nHedef Bandı: {target_goal}\n"
                f"Haftalık Çalışma Süresi: {weekly_hours} saat\n"
                f"Zorlanılan Dersler / Kazanımlar: {weak_topics}\n"
                f"Mentorluk Odak Alanı: {coach_focus}\n"
                f"Mevcut Deneme / Süre Durumu: {recent_score}\n"
                f"Aşamalı Kamp & Analiz Protokolü: {'Evet' if include_routine else 'Özet'}\n\n"
                "Lütfen öğrenciye doğrudan uygulanabilir, motive edici ve net bir Türkçe 2028 planı sun."
            )
            skel_coach = st.empty()
            render_loading_skeleton(skel_coach, lines=4)
            try:
                plan = call_gemini_text(prompt=prompt_coach, system_instruction=sys_coach)
                skel_coach.empty()
                st.session_state.dijimentor_plan = plan
                record_activity("Dijimentor (2028)", f"{target_exam} • {weekly_hours} saatlik plan üretildi")
                st.success("Dijimentor 2028 planı başarıyla oluşturuldu.")
            except Exception as exc:
                skel_coach.empty()
                refund_quota("coach", 1)
                st.error(f"Dijimentor planı oluşturulurken hata oluştu: {exc}")

    if st.session_state.dijimentor_plan:
        st.markdown("---")
        st.markdown("### Kişiselleştirilmiş Dijimentor 2028 Mentorluk Raporu")
        st.markdown(st.session_state.dijimentor_plan)

        meta_summary = (
            f"Hedef: {target_exam} | Hedef Bandı: {target_goal} | "
            f"Haftalık Süre: {weekly_hours} Saat | Odak: {coach_focus}"
        )
        docx_coach = build_coach_docx(meta_summary, st.session_state.dijimentor_plan)
        txt_report = (
            f"{BRAND_NAME} (2028)\n{meta_summary}\nTarih: {datetime.now().strftime('%d.%m.%Y')}\n"
            f"{'=' * 60}\n\n{st.session_state.dijimentor_plan}"
        )

        d_col1, d_col2, d_col3 = st.columns(3)
        with d_col1:
            st.download_button(
                label="Mentorluk Planını DOCX İndir",
                data=docx_coach,
                file_name="dijimentor_2028_calisma_plani.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with d_col2:
            st.download_button(
                label="Planı Metin (.txt) İndir",
                data=txt_report.encode("utf-8"),
                file_name="dijimentor_2028_calisma_plani.txt",
                mime="text/plain",
                use_container_width=True,
            )
        with d_col3:
            if st.button("Zayıf Konulardan Soru Üret", use_container_width=True):
                first_w = weak_topics.split(",")[0].strip()
                st.session_state["factory_focus_topic"] = first_w
                st.session_state["fab_theme_inp"] = first_w
                navigate_to("Soru & Redaksiyon Fabrikası")

# ==============================================================================
# 5. VELİ TAKİP PANELİ (+ BU HAFTA NE DEĞİŞTİ KARTLARI)
# ==============================================================================
elif current_page == "Veli Takip Paneli":
    stats_now = st.session_state["parent_student_stats"]
    v1, v2, v3 = st.columns(3)
    v1.metric("Haftalık Çalışma Süresi", f"{int(st.session_state.get('dk_hours', 16))} Saat", "Düzenli tempo")
    v2.metric("Çözülen Soru Adedi", f"{int(stats_now['weekly_solved'])} Soru", "+18% geçen haftaya göre")
    v3.metric("Kazanım İvmesi & Uyum", f"%{int(stats_now['mentor_completion'])}", "+4 puan doğruluk ivmesi")

    st.markdown("#### Bu Hafta Ne Değişti?")
    w_c1, w_c2, w_c3 = st.columns(3)
    w_c1.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#2563EB;">GÜÇLÜ KAZANIM</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['strong_topic']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">İsabet oranı istikrarlı biçimde yükseldi.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    w_c2.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#D97706;">RİSKLİ KAZANIM</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['weak_topic']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">Hata Defteri ve telafi seti ile takip ediliyor.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    w_c3.markdown(
        f"""
        <div class="kp-card">
            <div style="font-size:0.72rem;font-weight:700;color:#059669;">SIRADAKİ DENEME</div>
            <div style="font-weight:800;color:#1E2952;margin-top:0.25rem;">{html.escape(str(stats_now['next_exam_date']))}</div>
            <div style="font-size:0.8rem;color:#64748B;margin-top:0.2rem;">{html.escape(str(stats_now['next_exam_name']))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("#### Kazanım Bazlı Görünüm")
    st.progress(0.84, text="Okuma ve Anlama — %84")
    st.progress(0.72, text="Sözel Mantık & Sayısal Muhakeme — %72")
    st.progress(0.90, text="Düzenli Tekrar ve Ödev Tamamlama — %90")

    g_col1, g_col2 = st.columns(2)
    with g_col1:
        st.markdown("#### Haftalık Doğruluk İvmesi (%)")
        st.area_chart({"1. Hafta": 71, "2. Hafta": 74, "3. Hafta": 77, "4. Hafta": 81}, color="#2563EB", height=190)
    with g_col2:
        st.markdown("#### Ders Bazlı Başarı Oranı (%)")
        st.bar_chart({"Türkçe": 84, "Matematik": 72, "Fen": 82, "Sosyal": 86, "İngilizce": 79}, color="#2563EB", height=190)

    render_parent_notification_center()

# ==============================================================================
# 6. ÖĞRENCİ RÖNTGENİ
# ==============================================================================
elif current_page == "Öğrenci Röntgeni":
    render_student_xray_panel()

# ==============================================================================
# 7. B2B LİSANS & KURUM (+ DENEME TAKVİMİ)
# ==============================================================================
elif current_page == "B2B Lisans & Kurum":
    tab_b2b_form, tab_b2b_zumre = st.tabs(["Kurumsal Lisans Teklif Stüdyosu (2028)", "Zümre Soru Dağıtımı & Kurumsal Rapor"])

    with tab_b2b_form:
        with st.form("b2b_inquiry_form"):
            c1, c2 = st.columns(2)
            with c1:
                org_name = st.text_input("Kurum / Ofis Adı", value="Kurşun Kalem Akademi Ofisi")
                city = st.text_input("İl / İlçe", value="İstanbul")
                teacher_seats = st.number_input("Öğretmen / Editör Koltuk Sayısı", min_value=1, max_value=500, value=15)
            with c2:
                student_seats = st.number_input("Öğrenci Hesap Sayısı", min_value=10, max_value=10000, value=250)
                contact_name = st.text_input("Yetkili Adı Soyadı")
                contact_email = st.text_input("Kurumsal E-posta")

            submitted = st.form_submit_button("2028 Kurumsal Lisans Teklifi Hazırla", type="primary")
            if submitted:
                if not org_name.strip():
                    st.warning("Lütfen kurum adını girin.")
                else:
                    st.session_state["last_b2b_summary"] = (
                        f"Dönem: 2028 | Kurum: {org_name} ({city}) | Yetkili: {contact_name or '-'} | "
                        f"E-posta: {contact_email or '-'} | Koltuk: {teacher_seats} Öğretmen, "
                        f"{student_seats} Öğrenci"
                    )
                    st.session_state["b2b_org_cache"] = {
                        "org": org_name,
                        "teachers": int(teacher_seats),
                        "students": int(student_seats),
                    }
                    record_activity("B2B Lisans (2028)", f"{org_name} teklif taslağı hazırlandı")
                    st.success(
                        f"**{org_name}** ({city}) için 2028 dönemi **{teacher_seats} öğretmen** ve "
                        f"**{student_seats} öğrenci** kapasiteli kurumsal lisans talebi kaydedildi."
                    )

        if st.session_state.get("last_b2b_summary"):
            cache_for_doc = st.session_state.get("b2b_org_cache", {"org": org_name, "students": int(student_seats)})
            b2b_tier = get_tier_by_id("b2b_kurs" if cache_for_doc.get("students", 250) <= 300 else "b2b_kampus")
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
                title=f"{cache_for_doc.get('org', 'Kurumsal')} 2028 Lisans Teklif Özeti",
                summary=b2b_summary,
                billing_cycle="Yıllık (%20 İndirim)",
                installment=6,
                total_amount=round(b2b_tier["monthly_price"] * 12 * 0.80),
                extra_notes=st.session_state["last_b2b_summary"],
            )
            st.download_button(
                label="Kurumsal Lisans Teklifini DOCX İndir",
                data=b2b_docx,
                file_name="kursun_kalem_ofis_2028_kurumsal_teklif.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    with tab_b2b_zumre:
        org_cache = st.session_state.get("b2b_org_cache", {"org": "Kurşun Kalem Akademi Ofisi", "teachers": 15, "students": 250})
        recommended_tier = get_tier_by_id("b2b_kurs" if org_cache["students"] <= 300 else "b2b_kampus")
        total_q = recommended_tier["monthly_quota"]

        bk1, bk2, bk3, bk4 = st.columns(4)
        bk1.metric("Kurum Profili", org_cache["org"])
        bk2.metric("Aylık Soru Havuzu", f"{total_q:,}".replace(",", "."))
        bk3.metric("Öğretmen Başına Kota", f"{total_q // max(1, org_cache['teachers']):,} Soru".replace(",", "."))
        bk4.metric("Öğrenci Başına Test", f"{total_q // max(1, org_cache['students'])} Soru / Ay")

        st.markdown("#### Branş Zümreleri Arası Soru Kotası Dağılımı (2028)")
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
            st.bar_chart(chart_map, color="#2563EB", height=225)

        st.markdown("#### Kurum İçi Deneme & Kazanım Tarama Planı")
        st.dataframe(
            [
                {"Hafta": "1. Hafta", "Kademe": "8. Sınıf (LGS)", "Odak": "Sözel-Sayısal Muhakeme Tarama", "Durum": "Planlandı"},
                {"Hafta": "2. Hafta", "Kademe": "5–7. Sınıf", "Odak": "Maarif Modeli Ara Kazanım Testi", "Durum": "Planlandı"},
                {"Hafta": "3. Hafta", "Kademe": "9–12. Sınıf (YKS)", "Odak": "TYT-AYT Hız & Çeldirici Denemesi", "Durum": "Taslak"},
            ],
            use_container_width=True,
            hide_index=True,
        )

# ==============================================================================
# 8. PAKETLER & ABONELİK VİTRİNİ (2028)
# ==============================================================================
elif current_page == "Paketler & Abonelik Vitrini":
    total_pkg_count = len(SUBSCRIPTION_TIERS)

    st.markdown(
        """
        <div class="kp-value-strip">
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> Türkiye Yüzyılı Maarif Modeli 2028 Soru Havuzu</div>
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> Dijimentor Haftalık Akıllı Çalışma Planı</div>
            <div class="kp-value-strip-item"><span class="kp-value-dot"></span> iyzico Güvencesiyle 12 Taksite Varan Ödeme</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
                is_featured = tier["badge"] in ("En Çok Tercih Edilen", "VIP", "Kampüs")
                card_cls = "kp-pkg-card kp-pkg-featured" if is_featured else "kp-pkg-card"
                feats_li = "".join([f"<li>• {html.escape(feat)}</li>" for feat in tier["features"]])
                st.markdown(
                    f"""
                    <div class="{card_cls}">
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
    st.markdown(f"#### {total_pkg_count} Paket Karşılaştırma Matrisi (2028)")
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
