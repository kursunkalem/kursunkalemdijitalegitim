import base64
import io
import json
import os
import re
from datetime import datetime, timedelta
from typing import Dict, List, Optional

import streamlit as st

# İsteğe bağlı kütüphaneler (yüklü değilse zarifçe yedek mekanizmaya geçer)
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from PIL import Image, ImageDraw, ImageFont

    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import cairosvg

    CAIRO_AVAILABLE = True
except ImportError:
    CAIRO_AVAILABLE = False

try:
    import iyzipay

    IYZIPAY_AVAILABLE = True
except ImportError:
    IYZIPAY_AVAILABLE = False


# ==============================================================================
# 1. SAYFA YAPILANDIRMASI VE KURUMSAL KİMLİK (SVG LOGO & STİL)
# ==============================================================================
st.set_page_config(
    page_title="KURŞUN KALEM DİJİTAL EĞİTİM",
    page_icon="✏️",
    layout="wide",
    initial_sidebar_state="expanded",
)

BRAND_NAME = "KURŞUN KALEM DİJİTAL EĞİTİM"
BRAND_TAGLINE = "Türkiye Yüzyılı Maarif Modeli Uyumlu Yeni Nesil Eğitim & Ölçme Platformu"
DEDICATION_TEXT = (
    "Bu program Sevgili Annem Ayşe Kurşun'a ve Sevgili Babam Sinan Gürbüz Kurşun'a ithaf edilmiştir."
)

# Özel Vektör Logo (SVG)
BRAND_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 360 90" width="360" height="90">
  <defs>
    <linearGradient id="kkGrad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#0F172A"/>
      <stop offset="100%" stop-color="#1E3A8A"/>
    </linearGradient>
    <linearGradient id="goldGrad" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#F59E0B"/>
      <stop offset="100%" stop-color="#D97706"/>
    </linearGradient>
  </defs>
  <rect x="4" y="8" width="74" height="74" rx="18" fill="url(#kkGrad)"/>
  <!-- Stilize Kurşun Kalem İkonu -->
  <g transform="translate(22, 20) rotate(35 18 24)">
    <rect x="12" y="4" width="12" height="32" rx="2" fill="#F8FAFC"/>
    <rect x="12" y="4" width="12" height="6" rx="1" fill="url(#goldGrad)"/>
    <polygon points="12,36 24,36 18,47" fill="#FDE68A"/>
    <polygon points="16,43 20,43 18,47" fill="#0F172A"/>
  </g>
  <text x="92" y="40" font-family="Segoe UI, Arial, sans-serif" font-weight="800" font-size="21" fill="#0F172A" letter-spacing="0.6">KURŞUN KALEM</text>
  <text x="92" y="62" font-family="Segoe UI, Arial, sans-serif" font-weight="700" font-size="13" fill="#1D4ED8" letter-spacing="2.2">DİJİTAL EĞİTİM</text>
  <rect x="92" y="70" width="120" height="3" rx="1.5" fill="url(#goldGrad)"/>
</svg>"""


def get_svg_data_uri(svg_text: str) -> str:
    b64 = base64.b64encode(svg_text.encode("utf-8")).decode("utf-8")
    return f"data:image/svg+xml;base64,{b64}"


def generate_logo_png_bytes() -> Optional[bytes]:
    """DOCX belgelerine gömmek için vektör logoyu PNG byte dizisine dönüştürür."""
    if CAIRO_AVAILABLE:
        try:
            return cairosvg.svg2png(bytestring=BRAND_LOGO_SVG.encode("utf-8"), output_width=520)
        except Exception:
            pass

    if PIL_AVAILABLE:
        img = Image.new("RGBA", (540, 120), (255, 255, 255, 255))
        draw = ImageDraw.Draw(img)
        draw.rounded_rectangle([10, 14, 102, 106], radius=20, fill=(15, 23, 42, 255))
        # Kalem gövdesi
        draw.rectangle([48, 32, 64, 76], fill=(248, 250, 252, 255))
        draw.rectangle([48, 32, 64, 40], fill=(245, 158, 11, 255))
        draw.polygon([(48, 76), (64, 76), (56, 92)], fill=(253, 230, 138, 255))
        draw.polygon([(53, 86), (59, 86), (56, 92)], fill=(15, 23, 42, 255))
        try:
            font_big = ImageFont.truetype("DejaVuSans-Bold.ttf", 28)
            font_sub = ImageFont.truetype("DejaVuSans-Bold.ttf", 17)
        except Exception:
            font_big = ImageFont.load_default()
            font_sub = ImageFont.load_default()

        draw.text((118, 30), "KURŞUN KALEM", fill=(15, 23, 42, 255), font=font_big)
        draw.text((118, 66), "DİJİTAL EĞİTİM", fill=(29, 78, 216, 255), font=font_sub)
        draw.rectangle([118, 92, 290, 96], fill=(245, 158, 11, 255))

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    return None


def inject_custom_css() -> None:
    st.markdown(
        """
        <style>
        .main-header-card {
            background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
            border: 1px solid #E2E8F0;
            border-left: 6px solid #1E3A8A;
            border-radius: 14px;
            padding: 16px 22px;
            margin-bottom: 18px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
        }
        .dedication-banner {
            background: #FFFBEB;
            border: 1px solid #FDE68A;
            color: #92400E;
            padding: 8px 14px;
            border-radius: 8px;
            font-size: 0.86rem;
            font-weight: 500;
            margin-top: 8px;
        }
        .tier-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 12px;
            transition: all 0.2s ease;
        }
        .tier-card:hover {
            border-color: #2563EB;
            box-shadow: 0 6px 18px rgba(37, 99, 235, 0.08);
        }
        .parent-alert-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-left: 5px solid #2563EB;
            border-radius: 10px;
            padding: 14px 16px;
            margin-bottom: 10px;
        }
        .parent-alert-warn {
            border-left-color: #F59E0B;
            background: #FFFDF7;
        }
        .parent-alert-ok {
            border-left-color: #10B981;
            background: #F6FEF9;
        }
        .badge-segment {
            display: inline-block;
            padding: 3px 10px;
            border-radius: 999px;
            font-size: 0.75rem;
            font-weight: 700;
            background: #EFF6FF;
            color: #1D4ED8;
            margin-bottom: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# 2. ABONELİK KATALOĞU (12 PAKET - ÖĞRENCİ, YKS, VELİ, B2B)
# ==============================================================================
SUBSCRIPTION_TIERS: List[Dict] = [
    # ÖĞRENCİ SEGMENTİ (5-8. SINIF / LGS)
    {
        "id": "OGR_TEMEL",
        "segment": "Öğrenci (5–8. Sınıf)",
        "name": "Kalem Ucu Başlangıç",
        "monthly_price": 249.0,
        "question_quota": 500,
        "mentor_quota": 15,
        "features": ["Aylık 500 Yeni Nesil Soru", "15 Dijimentor Planı", "Temel Kazanım Raporu"],
    },
    {
        "id": "OGR_PRO",
        "segment": "Öğrenci (5–8. Sınıf)",
        "name": "Maarif Plus Ortaokul",
        "monthly_price": 449.0,
        "question_quota": 1500,
        "mentor_quota": 45,
        "features": ["Aylık 1.500 Soru", "45 Dijimentor Planı", "Haftalık Veli SMS/Rapor Özeti"],
    },
    {
        "id": "OGR_LGS",
        "segment": "Öğrenci (5–8. Sınıf)",
        "name": "LGS Şampiyonlar Kulübü",
        "monthly_price": 699.0,
        "question_quota": 4000,
        "mentor_quota": 120,
        "features": ["4.000 Beceri Temelli Soru", "Sınırsıza Yakın Dijimentor", "Türkiye Geneli Deneme Analizi"],
    },
    # YKS SEGMENTİ
    {
        "id": "YKS_TYT",
        "segment": "YKS Hazırlık",
        "name": "YKS TYT Odak Paketi",
        "monthly_price": 549.0,
        "question_quota": 2000,
        "mentor_quota": 60,
        "features": ["TYT Paragraf & Problem Stüdyosu", "60 Dijimentor Planı", "Hız & Süre Analizi"],
    },
    {
        "id": "YKS_SAY_EA",
        "segment": "YKS Hazırlık",
        "name": "YKS TYT + AYT Tam Paket",
        "monthly_price": 849.0,
        "question_quota": 5000,
        "mentor_quota": 150,
        "features": ["TYT + AYT Tüm Dersler", "150 Dijimentor Koçluk", "Kişisel Net Simülatörü"],
    },
    {
        "id": "YKS_DERECE",
        "segment": "YKS Hazırlık",
        "name": "YKS İlk 1000 Derece",
        "monthly_price": 1190.0,
        "question_quota": 10000,
        "mentor_quota": 300,
        "features": ["10.000 Seçkin Soru", "300 Dijimentor Analizi", "Birebir Haftalık Gelişim Karnesi"],
    },
    # VELİ SEGMENTİ
    {
        "id": "VELI_TAKIP",
        "segment": "Veli Kontrol",
        "name": "Veli Bilinçli Takip",
        "monthly_price": 199.0,
        "question_quota": 300,
        "mentor_quota": 20,
        "features": ["Anlık Veli Bildirim Paneli", "Haftalık Kazanım Karnesi", "WhatsApp/SMS Mesaj Şablonları"],
    },
    {
        "id": "VELI_AILE",
        "segment": "Veli Kontrol",
        "name": "Aile Boyu Eğitim (2 Kardeş)",
        "monthly_price": 749.0,
        "question_quota": 3000,
        "mentor_quota": 90,
        "features": ["2 Öğrenci Hesabı", "Ortak Veli Mesaj & Takip Merkezi", "Aylık 3.000 Soru Havuzu"],
    },
    {
        "id": "VELI_VIP",
        "segment": "Veli Kontrol",
        "name": "Veli VIP Rehberlik",
        "monthly_price": 990.0,
        "question_quota": 4500,
        "mentor_quota": 150,
        "features": ["Sınav Kaygısı & Çalışma Takibi", "Otomatik Veli Bilgilendirme Akışı", "Öncelikli Destek"],
    },
    # KURUMSAL / B2B SEGMENTİ
    {
        "id": "B2B_ZUMRE",
        "segment": "Kurumsal B2B",
        "name": "Öğretmen & Zümre Stüdyosu",
        "monthly_price": 1490.0,
        "question_quota": 12000,
        "mentor_quota": 250,
        "features": ["Logolu DOCX/TXT Sınav Üretimi", "Kazanım Tablosu Eşleştirme", "Toplu Veli Bilgilendirme Üretici"],
    },
    {
        "id": "B2B_KURS",
        "segment": "Kurumsal B2B",
        "name": "Kurs & Etüt Merkezi (100 Öğrenci)",
        "monthly_price": 4950.0,
        "question_quota": 50000,
        "mentor_quota": 1000,
        "features": ["100 Öğrenci Lisansı", "Kurumsal Veli Duyuru Merkezi", "Çoklu Şube Raporlama"],
    },
    {
        "id": "B2B_KOLEJ",
        "segment": "Kurumsal B2B",
        "name": "Kolej & Okul Kampüs Lisansı",
        "monthly_price": 9900.0,
        "question_quota": 150000,
        "mentor_quota": 5000,
        "features": ["500+ Öğrenci Kapasitesi", "Özel Okul Logolu Çıktılar", "Tam API & Veli Bilgilendirme Entegrasyonu"],
    },
]

TIER_MAP: Dict[str, Dict] = {t["id"]: t for t in SUBSCRIPTION_TIERS}


# ==============================================================================
# 3. OTURUM DURUMU (SESSION STATE) & SEPET CALLBACK FONKSİYONLARI
# ==============================================================================
def init_session_state() -> None:
    defaults = {
        "backpack_cart": {},  # {tier_id: qty}
        "billing_period": "Aylık",
        "installment_count": 1,
        "active_quotas": {
            "question_quota": 250,
            "mentor_quota": 10,
            "active_plans": ["Ücretsiz Deneme Paketi"],
        },
        "generated_questions": [],
        "student_stats": {
            "student_name": "Ege Yılmaz",
            "grade": "8. Sınıf (LGS)",
            "parent_name": "Selin Yılmaz",
            "parent_phone": "+90 555 123 45 67",
            "parent_email": "selin.yilmaz@ornek.com",
            "weekly_solved": 340,
            "correct_count": 286,
            "wrong_count": 38,
            "blank_count": 16,
            "mentor_completion": 88,
            "strong_topic": "Fiilimsiler ve Cümlenin Ögeleri",
            "weak_topic": "Paragrafta Yapı ve Sözel Mantık",
            "next_exam_date": (datetime.now() + timedelta(days=6)).strftime("%d.%m.%Y"),
            "next_exam_name": "Türkiye Geneli LGS Deneme Sınavı - 4",
        },
        "parent_announcements": [
            {
                "id": "ANN-1",
                "category": "Akademik İlerleme ve Başarı Takibi",
                "subtopic": "Haftalık Soru ve Net Özeti",
                "title": "Haftalık Çözüm Hedefi %113 Oranında Tamamlandı",
                "body": "Öğrencimiz bu hafta toplam 340 yeni nesil soru çözerek %84,1 doğruluk oranına ulaşmıştır.",
                "level": "ok",
                "date": datetime.now().strftime("%d.%m.%Y"),
            },
            {
                "id": "ANN-2",
                "category": "Çalışma Disiplini ve Dijimentor Takibi",
                "subtopic": "Eksik Telafi Uyarıları",
                "title": "Sözel Mantık & Paragraf Yapısı Telafi Etüdü Önerisi",
                "body": "Son 3 mini testte 'Sözel Mantık' kazanımında %62 başarı görülmüştür. Dijimentor programına 2 adet telafi testi eklenmiştir.",
                "level": "warn",
                "date": datetime.now().strftime("%d.%m.%Y"),
            },
            {
                "id": "ANN-3",
                "category": "Sınav, Ödev ve Takvim Duyuruları",
                "subtopic": "Yaklaşan Sınav ve Deneme Takvimi",
                "title": "Türkiye Geneli LGS Deneme Sınavı Hatırlatması",
                "body": "6 gün sonra saat 10.00'da çevrim içi deneme sınavı oturumu başlayacaktır. Öğrencimizin 15 dakika önce sistemde hazır bulunması önerilir.",
                "level": "info",
                "date": datetime.now().strftime("%d.%m.%Y"),
            },
        ],
        "payment_history": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def cb_add_to_bag(tier_id: str) -> None:
    cart = st.session_state["backpack_cart"]
    cart[tier_id] = cart.get(tier_id, 0) + 1


def cb_update_bag_qty(tier_id: str, delta: int) -> None:
    cart = st.session_state["backpack_cart"]
    if tier_id in cart:
        new_qty = cart[tier_id] + delta
        if new_qty <= 0:
            del cart[tier_id]
        else:
            cart[tier_id] = new_qty


def cb_remove_bag_item(tier_id: str) -> None:
    cart = st.session_state["backpack_cart"]
    if tier_id in cart:
        del cart[tier_id]


def cb_clear_bag() -> None:
    st.session_state["backpack_cart"] = {}


def get_bag_summary() -> Dict:
    period = st.session_state.get("billing_period", "Aylık")
    multiplier = {"Aylık": 1, "3 Aylık (%10 İndirim)": 3, "Yıllık (%20 İndirim)": 12}[period]
    discount = {"Aylık": 1.0, "3 Aylık (%10 İndirim)": 0.90, "Yıllık (%20 İndirim)": 0.80}[period]

    items = []
    subtotal = 0.0
    total_q = 0
    total_m = 0

    for tier_id, qty in st.session_state["backpack_cart"].items():
        tier = TIER_MAP.get(tier_id)
        if not tier or qty <= 0:
            continue
        unit_period_price = round(tier["monthly_price"] * multiplier * discount, 2)
        line_total = round(unit_period_price * qty, 2)
        subtotal += line_total
        total_q += tier["question_quota"] * qty * multiplier
        total_m += tier["mentor_quota"] * qty * multiplier
        items.append(
            {
                "tier": tier,
                "qty": qty,
                "unit_price": unit_period_price,
                "line_total": line_total,
            }
        )

    return {
        "items": items,
        "item_count": sum(i["qty"] for i in items),
        "total_amount": round(subtotal, 2),
        "period": period,
        "installment": st.session_state.get("installment_count", 1),
        "total_questions": total_q,
        "total_mentor": total_m,
    }


def sync_plan_quotas(bag_summary: Dict, force_reset: bool = False) -> None:
    if force_reset:
        st.session_state["active_quotas"] = {
            "question_quota": bag_summary["total_questions"],
            "mentor_quota": bag_summary["total_mentor"],
            "active_plans": [f"{i['tier']['name']} (x{i['qty']})" for i in bag_summary["items"]],
        }
    else:
        st.session_state["active_quotas"]["question_quota"] += bag_summary["total_questions"]
        st.session_state["active_quotas"]["mentor_quota"] += bag_summary["total_mentor"]
        for i in bag_summary["items"]:
            st.session_state["active_quotas"]["active_plans"].append(f"{i['tier']['name']} (x{i['qty']})")


# ==============================================================================
# 4. LOGOLU DOCX VE TXT BELGE ÜRETİCİLERİ
# ==============================================================================
def build_branded_txt(title: str, sections: List[tuple]) -> bytes:
    lines = [
        "====================================================================",
        f"  {BRAND_NAME}",
        f"  {BRAND_TAGLINE}",
        f"  Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
        "====================================================================",
        f"\nBAŞLIK: {title.upper()}\n",
    ]
    for sec_title, sec_body in sections:
        lines.append(f"--- {sec_title} ---")
        lines.append(sec_body.strip())
        lines.append("")

    lines.append("--------------------------------------------------------------------")
    lines.append(DEDICATION_TEXT)
    lines.append("--------------------------------------------------------------------")
    return "\n".join(lines).encode("utf-8")


def build_branded_docx(title: str, sections: List[tuple]) -> Optional[bytes]:
    if not DOCX_AVAILABLE:
        return None

    doc = Document()
    logo_bytes = generate_logo_png_bytes()

    # Üst Antet: Kurumsal Vektör Logo
    header_para = doc.add_paragraph()
    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if logo_bytes:
        run_img = header_para.add_run()
        run_img.add_picture(io.BytesIO(logo_bytes), width=Inches(2.5))
    else:
        run_txt = header_para.add_run(f"{BRAND_NAME}\n")
        run_txt.bold = True
        run_txt.font.size = Pt(16)
        run_txt.font.color.rgb = RGBColor(15, 23, 42)

    sub_p = doc.add_paragraph()
    r_sub = sub_p.add_run(f"{BRAND_TAGLINE}  |  Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    title_p = doc.add_heading(title, level=1)
    for r in title_p.runs:
        r.font.color.rgb = RGBColor(30, 58, 138)

    for sec_title, sec_body in sections:
        h = doc.add_heading(sec_title, level=2)
        for r in h.runs:
            r.font.color.rgb = RGBColor(15, 23, 42)
        p = doc.add_paragraph(sec_body.strip())
        p.paragraph_format.space_after = Pt(8)

    footer_p = doc.add_paragraph()
    footer_p.paragraph_format.space_before = Pt(16)
    r_ded = footer_p.add_run(DEDICATION_TEXT)
    r_ded.italic = True
    r_ded.font.size = Pt(9)
    r_ded.font.color.rgb = RGBColor(146, 64, 14)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def render_download_buttons(doc_title: str, sections: List[tuple], file_prefix: str) -> None:
    c1, c2 = st.columns(2)
    txt_data = build_branded_txt(doc_title, sections)
    docx_data = build_branded_docx(doc_title, sections)

    with c1:
        if docx_data:
            st.download_button(
                label="📄 Logolu Word (.DOCX) İndir",
                data=docx_data,
                file_name=f"{file_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        else:
            st.info("DOCX için `pip install python-docx pillow` kurulu olmalıdır.")

    with c2:
        st.download_button(
            label="📝 Kurumsal Metin (.TXT) İndir",
            data=txt_data,
            file_name=f"{file_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            mime="text/plain",
            use_container_width=True,
        )


# ==============================================================================
# 5. IYZICO ÖDEME SERVİSİ (CANLI SDK + SANDBOX HAZIR ENTEGRASYON)
# ==============================================================================
def get_iyzico_options() -> Optional[Dict[str, str]]:
    try:
        api_key = st.secrets.get("IYZICO_API_KEY", os.getenv("IYZICO_API_KEY", ""))
        secret_key = st.secrets.get("IYZICO_SECRET_KEY", os.getenv("IYZICO_SECRET_KEY", ""))
        base_url = st.secrets.get(
            "IYZICO_BASE_URL", os.getenv("IYZICO_BASE_URL", "https://sandbox-api.iyzipi.com")
        )
        if api_key and secret_key:
            return {
                "api_key": api_key,
                "secret_key": secret_key,
                "base_url": base_url.replace("https://", "").replace("http://", ""),
            }
    except Exception:
        pass
    return None


def create_iyzico_checkout_session(bag: Dict, buyer_info: Dict) -> Dict:
    """
    Okul Sırt Çantası sepetini gerçek iyzico CheckoutFormInitialize isteğine dönüştürür.
    Secrets tanımlı değilse geliştiriciye yol gösteren güvenli yanıt döner.
    """
    options = get_iyzico_options()
    if not options:
        return {
            "status": "missing_keys",
            "message": "Streamlit Secrets içinde IYZICO_API_KEY ve IYZICO_SECRET_KEY tanımlı değil.",
        }

    if not IYZIPAY_AVAILABLE:
        return {
            "status": "missing_sdk",
            "message": "Sunucu ortamında 'iyzipay' paketi yüklü değil (pip install iyzipay).",
        }

    basket_items = []
    for idx, entry in enumerate(bag["items"], start=1):
        tier = entry["tier"]
        basket_items.append(
            {
                "id": f"{tier['id']}_{idx}",
                "name": f"{tier['name']} ({bag['period']}) x{entry['qty']}",
                "category1": "Dijital Eğitim",
                "category2": tier["segment"],
                "itemType": "VIRTUAL",
                "price": f"{entry['line_total']:.2f}",
            }
        )

    conversation_id = f"KKDE-{int(datetime.now().timestamp())}"
    total_str = f"{bag['total_amount']:.2f}"

    request_payload = {
        "locale": "tr",
        "conversationId": conversation_id,
        "price": total_str,
        "paidPrice": total_str,
        "currency": "TRY",
        "basketId": f"BAG-{conversation_id}",
        "paymentGroup": "SUBSCRIPTION",
        "callbackUrl": buyer_info.get("callback_url", "https://kursunkalemdijitalokul.com.tr/odeme-sonuc"),
        "enabledInstallments": ["1", "2", "3", "6", "9"],
        "buyer": {
            "id": "BYR-KKDE-001",
            "name": buyer_info["name"],
            "surname": buyer_info["surname"],
            "gsmNumber": buyer_info["phone"],
            "email": buyer_info["email"],
            "identityNumber": buyer_info["identity_no"],
            "registrationAddress": buyer_info["address"],
            "ip": "85.34.78.112",
            "city": buyer_info["city"],
            "country": "Turkey",
        },
        "shippingAddress": {
            "contactName": f"{buyer_info['name']} {buyer_info['surname']}",
            "city": buyer_info["city"],
            "country": "Turkey",
            "address": buyer_info["address"],
        },
        "billingAddress": {
            "contactName": f"{buyer_info['name']} {buyer_info['surname']}",
            "city": buyer_info["city"],
            "country": "Turkey",
            "address": buyer_info["address"],
        },
        "basketItems": basket_items,
    }

    try:
        checkout_init = iyzipay.CheckoutFormInitialize().create(request_payload, options)
        raw_res = checkout_init.read().decode("utf-8")
        parsed = json.loads(raw_res)
        return {"status": "ok", "data": parsed, "conversation_id": conversation_id}
    except Exception as exc:
        return {"status": "error", "message": str(exc)}


# ==============================================================================
# 6. VELİ BİLGİLENDİRME VE DUYURU MERKEZİ (4 ANA KATEGORİ + OTOMATİK MESAJ)
# ==============================================================================
PARENT_MESSAGE_TOPICS: Dict[str, Dict] = {
    "1. Akademik İlerleme ve Başarı Takibi": {
        "subtopics": [
            "Haftalık Soru ve Net Özeti",
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi",
            "Deneme Sınavı ve LGS Değerlendirmesi",
        ],
        "templates": {
            "Haftalık Soru ve Net Özeti": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Eğitim platformunda öğrencimiz {student_name} ({grade}) "
                "bu hafta toplam {weekly_solved} yeni nesil soru çözmüştür. "
                "Çözümlerde {correct_count} doğru, {wrong_count} yanlış ve {blank_count} boş ile "
                "%{accuracy} başarı oranı yakalanmıştır.\n\n"
                "Öne Çıkan Güçlü Kazanım: {strong_topic}\n"
                "Desteklenmesi Önerilen Alan: {weak_topic}\n\n"
                "Bilgilerinize sunar, verimli bir hafta dileriz.\n"
                "Kurşun Kalem Dijital Eğitim Akademik Koordinatörlüğü"
            ),
            "Türkiye Yüzyılı Maarif Modeli Kazanım Analizi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in Türkiye Yüzyılı Maarif Modeli beceri temelli kazanım haritası güncellenmiştir. "
                "'{strong_topic}' alanında üst düzey hakimiyet sağlanırken, '{weak_topic}' kazanımında "
                "ek pekiştirme testleri öğrencimizin paneline tanımlanmıştır.\n\n"
                "Detaylı kazanım karnesini Veli Takip Paneli üzerinden inceleyebilirsiniz.\n"
                "Kurşun Kalem Dijital Eğitim"
            ),
            "Deneme Sınavı ve LGS Değerlendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in son deneme sınavı analizine göre genel doğruluk oranı %{accuracy} seviyesindedir. "
                "Sınav süresi yönetimi ve paragraf-mantık sorularındaki net artışı hedeflerimiz doğrultusunda ilerlemektedir. "
                "Bir sonraki deneme oturumumuz {next_exam_date} tarihinde '{next_exam_name}' adıyla yapılacaktır.\n\n"
                "Saygılarımızla,\nKurşun Kalem Dijital Eğitim"
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
                "Öğrencimiz {student_name}, bu haftaki Dijimentor kişisel çalışma programının %{mentor_completion} "
                "bölümünü zamanında tamamlamıştır. Düzenli tekrar disiplini sınav başarısını doğrudan desteklemektedir.\n\n"
                "Kurşun Kalem Dijital Eğitim Rehberlik Birimi"
            ),
            "Aktif Kullanım ve Düzenlilik": (
                "Sayın {parent_name},\n\n"
                "{student_name} bu hafta platformumuzda düzenli soru çözümü ve mini kazanım tarama testlerine katılım göstermiştir. "
                "Günlük çalışma rutininin korunması adına akşam saatlerinde 30 dakikalık tekrar oturumu önerilmektedir.\n\n"
                "Kurşun Kalem Dijital Eğitim"
            ),
            "Eksik Telafi Uyarıları": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in '{weak_topic}' ünitesindeki telafi görevlerinin henüz tamamlanmadığı görülmüştür. "
                "Konu eksiğinin birikmemesi için bu hafta sonuna kadar ilgili mini testlerin çözülmesini rica ederiz.\n\n"
                "Kurşun Kalem Dijital Eğitim"
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
                "{next_exam_date} tarihinde '{next_exam_name}' oturumu gerçekleştirilecektir. "
                "Öğrencimiz {student_name}'in sınav başlangıç saatinden 15 dakika önce öğrenci portalına giriş yapması önemlidir.\n\n"
                "Başarılar dileriz.\nKurşun Kalem Dijital Eğitim"
            ),
            "Ödev ve Etkinlik Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in paneline '{weak_topic}' ve '{strong_topic}' konularını kapsayan "
                "haftalık dijital ölçme ödevi atanmıştır. Son tamamlanma tarihi Pazar saat 21.00'dir.\n\n"
                "Kurşun Kalem Dijital Eğitim"
            ),
            "Özel Dönem Bilgilendirmeleri": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Eğitim 'Hızlandırılmış Kazanım ve Kamp Programı' kayıtları açılmıştır. "
                "Öğrencimiz {student_name} için özel hazırlanan tekrar modüllerine panel üzerinden erişebilirsiniz.\n\n"
                "Kurşun Kalem Dijital Eğitim"
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
                "Öğrencimiz {student_name} adına tanımlı aktif abonelik paketi: {active_plans}. "
                "Hesabınızda şu an {rem_questions} adet soru çözüm hakkı ve {rem_mentor} adet Dijimentor plan kredisi bulunmaktadır.\n\n"
                "Kurşun Kalem Dijital Eğitim"
            ),
            "Ödeme ve Fatura Bilgilendirmesi": (
                "Sayın {parent_name},\n\n"
                "Kurşun Kalem Dijital Eğitim abonelik ödemeniz iyzico güvenli ödeme altyapısı üzerinden başarıyla onaylanmış "
                "ve e-Arşiv faturanız kayıtlı e-posta adresinize ({parent_email}) iletilmiştir.\n\n"
                "Teşekkür ederiz.\nKurşun Kalem Dijital Eğitim"
            ),
            "Destek ve Veli-Öğretmen İletişimi": (
                "Sayın {parent_name},\n\n"
                "Öğrencimiz {student_name}'in akademik gelişimiyle ilgili birebir değerlendirme görüşmesi talep etmek "
                "veya teknik destek almak için Veli Takip Paneli üzerinden bize dilediğiniz zaman ulaşabilirsiniz.\n\n"
                "Kurşun Kalem Dijital Eğitim"
            ),
        },
    },
}


def format_parent_message(template: str) -> str:
    s = st.session_state["student_stats"]
    q = st.session_state["active_quotas"]
    solved = max(1, s["weekly_solved"])
    accuracy = round((s["correct_count"] / solved) * 100, 1)

    return template.format(
        parent_name=s["parent_name"],
        student_name=s["student_name"],
        grade=s["grade"],
        parent_email=s["parent_email"],
        weekly_solved=s["weekly_solved"],
        correct_count=s["correct_count"],
        wrong_count=s["wrong_count"],
        blank_count=s["blank_count"],
        accuracy=accuracy,
        mentor_completion=s["mentor_completion"],
        strong_topic=s["strong_topic"],
        weak_topic=s["weak_topic"],
        next_exam_date=s["next_exam_date"],
        next_exam_name=s["next_exam_name"],
        active_plans=", ".join(q["active_plans"]),
        rem_questions=q["question_quota"],
        rem_mentor=q["mentor_quota"],
    )


def render_parent_notification_center() -> None:
    st.subheader("👨‍👩‍👧 Veli Bilgilendirme, Duyuru ve Otomatik Mesaj Merkezi")
    st.caption(
        "Öğrenci verilerine göre otomatik veli mesajı (WhatsApp / SMS / E-posta) oluşturun, "
        "veli paneline canlı duyuru yayınlayın ve kurumsal logolu rapor indirin."
    )

    tab_generator, tab_feed, tab_student_cfg = st.tabs(
        [
            "📲 Otomatik Veli Mesaj Üreticisi",
            "📢 Veli Paneli Canlı Duyuru Akışı",
            "⚙️ Öğrenci & Veli Profil Verileri",
        ]
    )

    # TAB 1: 4 ANA BAŞLIK ALTINDA OTOMATİK MESAJ ÜRETİCİ
    with tab_generator:
        col_left, col_right = st.columns([1.05, 1.35])

        with col_left:
            selected_cat = st.selectbox(
                "1. Ana Konu Kategorisi Seçin",
                list(PARENT_MESSAGE_TOPICS.keys()),
            )
            subtopics = PARENT_MESSAGE_TOPICS[selected_cat]["subtopics"]
            selected_sub = st.selectbox("2. Alt Konu Başlığı Seçin", subtopics)

            channel = st.radio(
                "3. Gönderim Kanalı Formatı",
                ["WhatsApp Mesajı", "Kısa SMS Özeti", "Kurumsal E-Posta"],
                horizontal=True,
            )

            raw_template = PARENT_MESSAGE_TOPICS[selected_cat]["templates"][selected_sub]
            auto_filled_msg = format_parent_message(raw_template)

            if channel == "Kısa SMS Özeti":
                s = st.session_state["student_stats"]
                acc = round((s["correct_count"] / max(1, s["weekly_solved"])) * 100, 1)
                auto_filled_msg = (
                    f"KURŞUN KALEM EĞİTİM: Sn. {s['parent_name']}, öğrencimiz {s['student_name']} "
                    f"({selected_sub}) haftalık {s['weekly_solved']} soru, %{acc} başarı, "
                    f"Dijimentor %{s['mentor_completion']} uyum sağladı. Detaylar Veli Panelinde."
                )

            if st.button("📌 Bu Mesajı Veli Paneline Canlı Duyuru Olarak Ekle", use_container_width=True):
                st.session_state["parent_announcements"].insert(
                    0,
                    {
                        "id": f"ANN-{int(datetime.now().timestamp())}",
                        "category": selected_cat,
                        "subtopic": selected_sub,
                        "title": f"{selected_sub} Bilgilendirmesi",
                        "body": auto_filled_msg,
                        "level": "warn" if "Uyarı" in selected_sub else "ok",
                        "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    },
                )
                st.success("Bilgilendirme mesajı Veli Takip Paneli duyuru akışına eklendi!")

        with col_right:
            st.markdown("**4. Oluşturulan Veli Bilgilendirme Metni (Düzenlenebilir)**")
            final_msg = st.text_area(
                "Mesaj İçeriği",
                value=auto_filled_msg,
                height=240,
                label_visibility="collapsed",
            )

            s_info = st.session_state["student_stats"]
            clean_phone = re.sub(r"\D", "", s_info["parent_phone"])
            encoded_msg = final_msg.replace("\n", "%0A").replace(" ", "%20")
            wa_link = f"https://wa.me/{clean_phone}?text={encoded_msg}"
            mail_link = (
                f"mailto:{s_info['parent_email']}?subject=Kurşun Kalem Dijital Eğitim - {selected_sub}"
                f"&body={encoded_msg}"
            )

            b1, b2 = st.columns(2)
            with b1:
                st.link_button("💬 WhatsApp ile Gönder", wa_link, use_container_width=True)
            with b2:
                st.link_button("✉️ E-Posta İstemcisinde Aç", mail_link, use_container_width=True)

            st.markdown("---")
            render_download_buttons(
                doc_title=f"Veli Bilgilendirme Raporu - {selected_sub}",
                sections=[
                    ("Kategori ve Konu", f"{selected_cat} > {selected_sub}"),
                    ("Öğrenci ve Veli Bilgisi", f"Öğrenci: {s_info['student_name']} | Veli: {s_info['parent_name']}"),
                    ("Resmi Bilgilendirme Metni", final_msg),
                ],
                file_prefix="veli_bilgilendirme",
            )

    # TAB 2: VELİ TAKİP PANELİ CANLI DUYURU AKIŞI
    with tab_feed:
        filter_cat = st.selectbox(
            "Kategoriye Göre Filtrele",
            ["Tümü"] + list(PARENT_MESSAGE_TOPICS.keys()),
        )
        anns = st.session_state["parent_announcements"]
        filtered = [a for a in anns if filter_cat == "Tümü" or a["category"] == filter_cat]

        if not filtered:
            st.info("Bu kategoride henüz yayınlanmış bir veli bilgilendirmesi bulunmuyor.")
        else:
            for item in filtered:
                css_cls = "parent-alert-card"
                if item["level"] == "warn":
                    css_cls += " parent-alert-warn"
                elif item["level"] == "ok":
                    css_cls += " parent-alert-ok"

                st.markdown(
                    f"""
                    <div class="{css_cls}">
                        <div style="font-size:0.78rem; color:#475569; font-weight:600;">
                            {item['category']} • {item['subtopic']} • 📅 {item['date']}
                        </div>
                        <div style="font-size:1.02rem; font-weight:700; color:#0F172A; margin:4px 0;">
                            {item['title']}
                        </div>
                        <div style="font-size:0.92rem; color:#1E293B; white-space:pre-line;">
                            {item['body']}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    # TAB 3: ÖĞRENCİ VE VELİ VERİ GÜNCELLEME
    with tab_student_cfg:
        s = st.session_state["student_stats"]
        with st.form("student_parent_form"):
            c1, c2 = st.columns(2)
            with c1:
                st_name = st.text_input("Öğrenci Adı Soyadı", value=s["student_name"])
                st_grade = st.text_input("Sınıf / Düzey", value=s["grade"])
                p_name = st.text_input("Veli Adı Soyadı", value=s["parent_name"])
                p_phone = st.text_input("Veli Telefon (WhatsApp)", value=s["parent_phone"])
                p_email = st.text_input("Veli E-Posta", value=s["parent_email"])
            with c2:
                w_solved = st.number_input("Haftalık Çözülen Soru", min_value=0, value=int(s["weekly_solved"]))
                c_cnt = st.number_input("Doğru Sayısı", min_value=0, value=int(s["correct_count"]))
                w_cnt = st.number_input("Yanlış Sayısı", min_value=0, value=int(s["wrong_count"]))
                m_comp = st.slider("Dijimentor Plan Uyumu (%)", 0, 100, int(s["mentor_completion"]))
                strong_t = st.text_input("Güçlü Kazanım", value=s["strong_topic"])
                weak_t = st.text_input("Desteklenecek Kazanım", value=s["weak_topic"])

            if st.form_submit_button("💾 Öğrenci ve Veli Bilgilerini Güncelle", use_container_width=True):
                s.update(
                    {
                        "student_name": st_name,
                        "grade": st_grade,
                        "parent_name": p_name,
                        "parent_phone": p_phone,
                        "parent_email": p_email,
                        "weekly_solved": w_solved,
                        "correct_count": c_cnt,
                        "wrong_count": w_cnt,
                        "blank_count": max(0, w_solved - (c_cnt + w_cnt)),
                        "mentor_completion": m_comp,
                        "strong_topic": strong_t,
                        "weak_topic": weak_t,
                    }
                )
                st.success("Bilgiler güncellendi! Tüm mesaj şablonları yeni verilere göre otomatik yenilendi.")


# ==============================================================================
# 7. OKUL SIRT ÇANTASI (SEPET) VE IYZICO ÖDEME EKRANI
# ==============================================================================
def render_school_backpack_sidebar() -> None:
    bag = get_bag_summary()
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"### 🎒 Okul Sırt Çantası ({bag['item_count']} Paket)")

    if not bag["items"]:
        st.sidebar.info("Sırt çantanız şu an boş. Paketler bölümünden abonelik ekleyebilirsiniz.")
        return

    for entry in bag["items"]:
        tier = entry["tier"]
        st.sidebar.markdown(f"**{tier['name']}** — `{entry['line_total']:.2f} ₺`")
        c1, c2, c3, c4 = st.sidebar.columns([1, 1, 1, 1])
        c1.button("➖", key=f"sb_dec_{tier['id']}", on_click=cb_update_bag_qty, args=(tier["id"], -1))
        c2.markdown(f"<div style='text-align:center;padding-top:5px;'><b>{entry['qty']}</b></div>", unsafe_allow_html=True)
        c3.button("➕", key=f"sb_inc_{tier['id']}", on_click=cb_update_bag_qty, args=(tier["id"], 1))
        c4.button("🗑️", key=f"sb_del_{tier['id']}", on_click=cb_remove_bag_item, args=(tier["id"],))

    st.sidebar.markdown(f"**Toplam Tutar:** `{bag['total_amount']:.2f} ₺` ({bag['period']})")
    st.sidebar.button("🧹 Çantayı Tamamen Boşalt", key="sb_clear_all", on_click=cb_clear_bag, use_container_width=True)


def render_packages_and_iyzico() -> None:
    st.subheader("🎒 Abonelik Paketleri, Okul Sırt Çantası & iyzico Ödeme")

    seg_filter = st.selectbox(
        "Segment Filtrele",
        ["Tümü", "Öğrenci (5–8. Sınıf)", "YKS Hazırlık", "Veli Kontrol", "Kurumsal B2B"],
    )
    visible_tiers = [t for t in SUBSCRIPTION_TIERS if seg_filter == "Tümü" or t["segment"] == seg_filter]

    cols = st.columns(3)
    for idx, tier in enumerate(visible_tiers):
        with cols[idx % 3]:
            feats_html = "".join([f"<li>{f}</li>" for f in tier["features"]])
            st.markdown(
                f"""
                <div class="tier-card">
                    <span class="badge-segment">{tier['segment']}</span>
                    <h4 style="margin:2px 0 6px 0; color:#0F172A;">{tier['name']}</h4>
                    <div style="font-size:1.35rem; font-weight:800; color:#1D4ED8; margin-bottom:8px;">
                        {tier['monthly_price']:.0f} ₺ <span style="font-size:0.8rem; color:#64748B;">/ ay</span>
                    </div>
                    <ul style="padding-left:18px; margin:0; font-size:0.85rem; color:#334155;">
                        {feats_html}
                    </ul>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.button(
                "🎒 Sırt Çantasına Ekle",
                key=f"add_{tier['id']}",
                on_click=cb_add_to_bag,
                args=(tier["id"],),
                use_container_width=True,
            )

    st.markdown("---")
    st.markdown("### 💳 iyzico Güvenli Ödeme ve Kota Aktivasyonu")

    c_p1, c_p2 = st.columns(2)
    with c_p1:
        st.session_state["billing_period"] = st.selectbox(
            "Abonelik Dönemi",
            ["Aylık", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"],
            index=["Aylık", "3 Aylık (%10 İndirim)", "Yıllık (%20 İndirim)"].index(
                st.session_state["billing_period"]
            ),
        )
    with c_p2:
        st.session_state["installment_count"] = st.selectbox("Taksit Seçeneği", [1, 2, 3, 6, 9])

    bag = get_bag_summary()
    if not bag["items"]:
        st.warning("Ödeme adımına geçmek için yukarıdan en az bir abonelik paketini Okul Sırt Çantasına ekleyin.")
        return

    st.info(
        f"**Sepet Özeti:** {bag['item_count']} paket | Dönem: **{bag['period']}** | "
        f"Tanımlanacak Kota: **+{bag['total_questions']} Soru** & **+{bag['total_mentor']} Dijimentor** | "
        f"Ödenecek Toplam: **{bag['total_amount']:.2f} ₺**"
    )

    with st.expander("🧾 Fatura ve Alıcı Bilgileri (iyzico Zorunlu Alanlar)", expanded=True):
        fc1, fc2, fc3 = st.columns(3)
        b_name = fc1.text_input("Ad", value="Kürşat Han")
        b_surname = fc2.text_input("Soyad", value="Kurşun")
        b_tc = fc3.text_input("T.C. Kimlik / Vergi No", value="11111111111")

        fc4, fc5, fc6 = st.columns(3)
        b_phone = fc4.text_input("GSM", value="+905551112233")
        b_email = fc5.text_input("E-Posta", value="bilgi@kursunkalemdijitalokul.com.tr")
        b_city = fc6.text_input("Şehir", value="İstanbul")
        b_addr = st.text_input("Fatura Adresi", value="Kadıköy / İstanbul, Türkiye")

    col_live, col_sim = st.columns(2)
    with col_live:
        if st.button("🔐 Gerçek iyzico Checkout Formu Başlat", type="primary", use_container_width=True):
            buyer = {
                "name": b_name,
                "surname": b_surname,
                "identity_no": b_tc,
                "phone": b_phone,
                "email": b_email,
                "city": b_city,
                "address": b_addr,
            }
            res = create_iyzico_checkout_session(bag, buyer)
            if res["status"] == "ok" and res["data"].get("status") == "success":
                st.success(f"iyzico ödeme oturumu açıldı! (İşlem No: {res['conversation_id']})")
                st.components.v1.html(res["data"].get("checkoutFormContent", ""), height=560, scrolling=True)
            else:
                st.warning(
                    f"Canlı iyzico formu başlatılamadı: {res.get('message', res.get('data', {}))}\n\n"
                    "Streamlit Secrets bölümüne `IYZICO_API_KEY` ve `IYZICO_SECRET_KEY` eklediğinizde "
                    "bu buton doğrudan canlı/sandbox iyzico kart ekranını açacaktır."
                )

    with col_sim:
        if st.button("✅ Test Ödemesini Onayla ve Kotaları Tanımla", use_container_width=True):
            sync_plan_quotas(bag, force_reset=False)
            st.session_state["payment_history"].append(
                {
                    "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
                    "amount": bag["total_amount"],
                    "period": bag["period"],
                    "items": [i["tier"]["name"] for i in bag["items"]],
                }
            )
            cb_clear_bag()
            st.success("Ödeme onaylandı! Paket kotaları hesabınıza tanımlandı ve Okul Sırt Çantası güncellendi.")
            st.rerun()


# ==============================================================================
# 8. DİĞER ÇEKİRDEK EĞİTİM MODÜLLERİ (SORU FABRİKASI, ÖĞRENCİ, DİJİMENTOR)
# ==============================================================================
def render_question_factory() -> None:
    st.subheader("🏭 Türkiye Yüzyılı Maarif Modeli Soru & İçerik Fabrikası")
    c1, c2, c3 = st.columns(3)
    grade = c1.selectbox("Sınıf Düzeyi", ["5. Sınıf", "6. Sınıf", "7. Sınıf", "8. Sınıf (LGS)"])
    subject = c2.selectbox("Ders", ["Türkçe", "Matematik", "Fen Bilimleri", "T.C. İnkılap Tarihi"])
    topic = c3.text_input("Kazanım / Konu Başlığı", value="Paragrafta Anlam ve Metin Yorumlama")

    if st.button("✨ Yeni Nesil Beceri Temelli Soru Üret", type="primary"):
        if st.session_state["active_quotas"]["question_quota"] <= 0:
            st.error("Soru üretim kotanız dolmuştur. Lütfen Okul Sırt Çantasından paket ekleyin.")
            return

        st.session_state["active_quotas"]["question_quota"] -= 1
        q_text = (
            f"[{grade} - {subject} | {topic}]\n"
            f"Soru: Günlük yaşamda dijital okuryazarlık becerisi gelişmiş bireyler, okudukları metinlerdeki "
            f"örtülü anlamları ve neden-sonuç ilişkilerini hızla ayırt edebilirler.\n"
            f"Bu açıklamaya göre aşağıdakilerden hangisi metnin ana düşüncesini en doğru biçimde yansıtır?\n"
            f"A) Dijital araçlar yalnızca bilgiye hızlı erişim sağlar.\n"
            f"B) Eleştirel okuma becerisi, metinler arası bağ kurmayı güçlendirir.\n"
            f"C) Okuma hızı arttıkça metnin derinliği azalır.\n"
            f"D) Geleneksel metinler dijital metinlerden daha kalıcıdır.\n"
            f"Doğru Cevap: B"
        )
        st.session_state["generated_questions"].insert(0, q_text)
        st.success("Soru üretildi ve havuza eklendi!")

    if st.session_state["generated_questions"]:
        combined = "\n\n----------------------------------------\n\n".join(st.session_state["generated_questions"])
        st.text_area("Üretilen Soru Havuzu", value=combined, height=220)
        render_download_buttons(
            doc_title=f"{grade} {subject} - Kazanım Değerlendirme Testi",
            sections=[("Sınav Soruları ve Cevap Anahtarı", combined)],
            file_prefix="kursun_kalem_sinav",
        )


def render_dijimentor_and_student() -> None:
    st.subheader("🧭 Dijimentor Kişisel Çalışma Planı & Öğrenci Mini Portalı")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📅 Haftalık Dijimentor Plan Üretici")
        daily_hours = st.slider("Günlük Çalışma Süresi (Saat)", 1, 6, 3)
        focus_lesson = st.selectbox("Öncelikli Ders", ["Türkçe & Paragraf", "Matematik", "Fen Bilimleri", "Karma LGS"])
        if st.button("🚀 Haftalık Plan Oluştur", use_container_width=True):
            if st.session_state["active_quotas"]["mentor_quota"] <= 0:
                st.error("Dijimentor kotanız dolmuştur.")
            else:
                st.session_state["active_quotas"]["mentor_quota"] -= 1
                plan_txt = (
                    f"Haftalık {focus_lesson} Odaklı Dijimentor Programı (Günlük {daily_hours} Saat):\n"
                    f"• Pazartesi - Çarşamba: Konu Tekrarı (45 dk) + 40 Yeni Nesil Soru + Yanlış Analizi\n"
                    f"• Perşembe - Cuma: Süreli Paragraf & Mantık Etüdü + 50 Karma Soru\n"
                    f"• Cumartesi: Tam Süreli Genel Deneme Sınavı + Veli Bilgilendirme Raporu\n"
                    f"• Pazar: Eksik Kazanım Telafisi ve Dinlenme"
                )
                st.info(plan_txt)
                render_download_buttons(
                    doc_title="Dijimentor Kişisel Haftalık Çalışma Planı",
                    sections=[("Haftalık Çalışma Programı", plan_txt)],
                    file_prefix="dijimentor_plan",
                )

    with c2:
        st.markdown("#### 🎯 Hızlı Kazanım Mini Testi")
        ans = st.radio(
            "Soru: 'Kitap okuyan insanlar olaylara çok yönlü bakar.' cümlesinde altı çizili fiilimsinin türü nedir?",
            ["İsim-Fiil", "Sıfat-Fiil", "Zarf-Fiil"],
        )
        if st.button("Cevabı Kontrol Et", use_container_width=True):
            if ans == "Sıfat-Fiil":
                st.success("Tebrikler! '-an' ekiyle kurulmuş sıfat-fiil (ortaç) doğru tespit edildi.")
            else:
                st.warning("İpucu: '-an, -ası, -mez, -ar, -dik, -ecek, -miş' ekleri sıfat-fiil oluşturur.")


# ==============================================================================
# 9. ANA UYGULAMA AKIŞI
# ==============================================================================
def main() -> None:
    init_session_state()
    inject_custom_css()

    # Üst Kurumsal Başlık ve Kalıcı Vektör Logo
    logo_uri = get_svg_data_uri(BRAND_LOGO_SVG)
    st.markdown(
        f"""
        <div class="main-header-card">
            <div style="display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px;">
                <div>
                    <img src="{logo_uri}" alt="Kurşun Kalem Dijital Eğitim Logo" style="height:64px;"/>
                    <div style="color:#475569; font-size:0.92rem; margin-top:4px;">{BRAND_TAGLINE}</div>
                </div>
                <div style="text-align:right; font-size:0.85rem; color:#1E293B;">
                    <div><b>Aktif Soru Kotası:</b> {st.session_state['active_quotas']['question_quota']}</div>
                    <div><b>Dijimentor Kredisi:</b> {st.session_state['active_quotas']['mentor_quota']}</div>
                </div>
            </div>
            <div class="dedication-banner">💛 {DEDICATION_TEXT}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Sol Menü Navigasyonu
    st.sidebar.image(logo_uri, use_container_width=True)
    menu = st.sidebar.radio(
        "Platform Modülleri",
        [
            "👨‍👩‍👧 Veli Bilgilendirme & Duyuru Merkezi",
            "🎒 Paketler, Sırt Çantası & iyzico",
            "🏭 Soru & İçerik Fabrikası",
            "🧭 Dijimentor & Öğrenci Portalı",
        ],
    )
    render_school_backpack_sidebar()

    if menu == "👨‍👩‍👧 Veli Bilgilendirme & Duyuru Merkezi":
        render_parent_notification_center()
    elif menu == "🎒 Paketler, Sırt Çantası & iyzico":
        render_packages_and_iyzico()
    elif menu == "🏭 Soru & İçerik Fabrikası":
        render_question_factory()
    elif menu == "🧭 Dijimentor & Öğrenci Portalı":
        render_dijimentor_and_student()


if __name__ == "__main__":
    main()
