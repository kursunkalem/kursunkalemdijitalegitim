import streamlit as st
import iyzipay
import json

# Sayfa Yapılandırması
st.set_page_config(page_title="Kurşun Kalem Dijital Okul", page_icon="📚", layout="centered")

st.title("📚 Kurşun Kalem Dijital Eğitim Platformu")
st.write("Soru bankaları, dijital içerikler ve yapay zeka destekli eğitim modülleri.")

# Sekmeler
tab1, tab2 = st.tabs(["Ana Sayfa / Soru Paketleri", "Güvenli Ödeme Ekranı"])

with tab1:
    st.header("Dijital Soru Paketi")
    st.write("LGS ve yeni nesil beceri temelli soru paketine anında erişin.")
    st.info("Fiyat: 150.00 TL")

with tab2:
    st.header("İyzico Güvenli Ödeme")
    
    # Kullanıcı Bilgileri Formu
    with st.form("payment_form"):
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Ad", "Kürşat Han")
        with col2:
            surname = st.text_input("Soyad", "Kurşun")
            
        email = st.text_input("E-posta Adresi", "kullanici@ornek.com")
        submit_btn = st.form_submit_button("Ödeme Sayfasını Oluştur")

    if submit_btn:
        # İyzico Bilgilerini Vercel Environment Variables'tan Çekiyoruz
        options = {
            'api_key': st.secrets.get("IYSU_API_KEY", "sandbox-api-key"),
            'secret_key': st.secrets.get("IYSU_SECRET_KEY", "sandbox-secret-key"),
            'base_url': 'https://sandbox-api.iyzipay.com' # Canlıya geçince https://api.iyzipay.com yapılır
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
                st.success("Ödeme sayfası hazırlandı! Aşağıdaki butona tıklayarak güvenli ödeme ekranına geçebilirsiniz.")
                st.markdown(f'<a href="{payment_page_url}" target="_blank"><button style="background-color:#0080ff;color:white;padding:10px 20px;border:none;border-radius:5px;cursor:pointer;">İyzico ile Güvenli Öde (150 TL)</button></a>', unsafe_allow_html=True)
            else:
                st.error(f"Ödeme başlatılamadı: {response_dict.get('errorMessage')}")
        except Exception as e:
            st.error(f"Bir hata oluştu: {str(e)}")
