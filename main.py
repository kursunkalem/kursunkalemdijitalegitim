import streamlit as st
import iyzipay
import json

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Kurşun Kalem Dijital Eğitim Platformu",
    page_icon="📚",
    layout="wide"
)

# Başlık ve Karşılama
st.title("📚 Kurşun Kalem Dijital Eğitim Platformu")
st.markdown("LGS ve yeni nesil beceri temelli soru paketleri, akademik içerikler ve dijital eğitim yönetim paneli.")

# Ana Sekmeler (Eski harika içeriklerin ve ödeme sistemi bir arada)
tab1, tab2, tab3 = st.tabs(["🚀 Soru Üretim & Eğitim Paneli", "📦 Dijital Paketler & Mağaza", "💳 Güvenli Ödeme Ekranı"])

with tab1:
    st.header("Eğitim Modülleri ve Soru Üretim Merkezi")
    st.write("Burada kendi hazırladığın kazanım tabanlı soru üretim araçları, içerikler ve yönetim panellerin yer alır.")
    
    # Örnek çalışma alanı alanı
    with st.expander("Soru Üretim Modülünü Görüntüle"):
        st.info("Yapay zeka destekli soru havuzu ve kazanım eşleştirme sistemi aktif durumda.")
        kriter = st.text_input("Kazanım / Konu Başlığı Girin", "Örn: Sözel Mantık ve Muhakeme")
        if st.button("Soruları Listele / Üret"):
            st.success(f"'{kriter}' için içerikler ve test modülleri hazırlandı.")

with tab2:
    st.header("Dijital Soru Paketleri")
    st.write("Öğrenciler ve öğretmenler için hazırlanan özel lisanslı içerikler.")
    
    col1, col2 = st.cloud_col = st.columns(2) if hasattr(st, "columns") else (st, st)
    with col1:
        st.subheader("LGS Beceri Temelli Paket")
        st.write("Tamamı özgün, yeni nesil sorulardan oluşan kapsamlı dijital arşiv.")
        st.markdown("**Fiyat: 150.00 TL**")
        
with tab3:
    st.header("İyzico Güvenli Ödeme Ekranı")
    st.write("Seçtiğiniz dijital eğitim paketi için güvenli ödemenizi tamamlayın.")
    
    # Ödeme Bilgileri Formu
    with st.form("payment_form"):
        col_a, col_b = st.columns(2)
        with col_a:
            name = st.text_input("Adınız", "Kürşat Han")
        with col_b:
            surname = st.text_input("Soyadınız", "Kurşun")
            
        email = st.text_input("E-posta Adresiniz", "kullanici@ornek.com")
        submit_btn = st.form_submit_button("Ödeme Sayfasını Oluştur (150 TL)")

    if submit_btn:
        # Vercel Environment Variables'tan anahtarları güvenle çekiyoruz
        options = {
            'api_key': st.secrets.get("IYSU_API_KEY", "sandbox-api-key"),
            'secret_key': st.secrets.get("IYSU_SECRET_KEY", "sandbox-secret-key"),
            'base_url': 'https://sandbox-api.iyzipay.com' # Canlıya geçince https://api.iyzipay.com olur
        }

        request = {
            'locale': 'tr',
            'conversation_id': '123456789',
            'price': '150.00',
            'paid_price': '150.00',
            'currency': 'TRY',
            'basket_id': 'B67890',
            'payment_group': 'PRODUCT',
            'callback_url': 'https://kursunkalemdijitalokul.com.tr',
            'buyer': {
                'id': 'BY789',
                'name': name,
                'surname': surname,
                'gsm_number': '+905350000000',
                'email': email,
                'identity_number': '74300864791',
                'registration_address': 'İstanbul, Türkiye',
                'ip': '85.34.78.112',
                'city': 'Istanbul',
                'country': 'Turkey',
            },
            'shipping_address': {
                'contact_name': f"{name} {surname}",
                'city': 'Istanbul',
                'country': 'Turkey',
                'address': 'İstanbul, Türkiye',
            },
            'billing_address': {
                'contact_name': f"{name} {surname}",
                'city': 'Istanbul',
                'country': 'Turkey',
                'address': 'İstanbul, Türkiye',
            },
            'basket_items': [
                {
                    'id': 'BI101',
                    'name': 'Dijital Soru Bankası Paketi',
                    'category1': 'Eğitim',
                    'item_type': 'VIRTUAL',
                    'price': '150.00'
                }
            ]
        }

        try:
            checkout_form_initialize = iyzipay.CheckoutFormInitialize().create(request, options)
            response_str = checkout_form_initialize.read().decode('utf-8')
            response_dict = json.loads(response_str)

            if response_dict.get('status') == 'success':
                payment_page_url = response_dict.get('paymentPageUrl')
                st.success("Ödeme sayfası başarıyla oluşturuldu!")
                st.markdown(f'<a href="{payment_page_url}" target="_blank"><button style="background-color:#0080ff;color:white;padding:12px 24px;border:none;border-radius:6px;cursor:pointer;font-size:16px;">İyzico ile Güvenli Öde (150 TL)</button></a>', unsafe_allow_html=True)
            else:
                st.error(f"Ödeme başlatılamadı: {response_dict.get('errorMessage')}")
        except Exception as e:
            st.error(f"Bir hata oluştu: {str(e)}")
