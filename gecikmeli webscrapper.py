from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import random

# Excel dosyasından depo isimlerini çekme
excel_file_path = '/Users/gucluceyhan/Documents/EczaDepoları/Ecza Depoları İsimleri.xlsx'
depo_df = pd.read_excel(excel_file_path)

# Depo isimlerini liste olarak alma
depo_list = depo_df['Depo'].tolist()

# Selenium için driver kurulum
chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
service = Service('/usr/local/bin/chromedriver')
driver = webdriver.Chrome(service=service, options=chrome_options)

# Web adresleri için boş liste
websites = []

# Her bir depo için Google'da arama yap
for depo in depo_list:
    driver.get("https://www.google.com")
    search_box = driver.find_element("name", "q")
    search_box.send_keys(f"{depo} web adresi")
    search_box.send_keys(Keys.RETURN)
    
    # Rastgele gecikme ekleme
    time.sleep(random.uniform(3, 6))  # 3-6 saniye arasında rastgele bir bekleme

    # İlk sonucu almak için deneme
    try:
        # Eğer CAPTCHA çıkarsa manuel müdahale için durdurma
        input("Eğer CAPTCHA ile karşılaşırsanız, manuel olarak geçin ve devam etmek için ENTER tuşuna basın...")

        # İlk sonuçtaki URL'yi alma
        first_result = driver.find_element("css selector", 'div.yuRUbf > a')
        website = first_result.get_attribute('href')
        websites.append(website)
    except Exception as e:
        websites.append('Not Found')

# Sonuçları DataFrame'e aktar
depo_df['Web Adresi'] = websites

# DataFrame'i Excel dosyasına yazma
excel_output_path = '/Users/gucluceyhan/Documents/EczaDepoları/EczaDepolari_Web_Adresleri.xlsx'
depo_df.to_excel(excel_output_path, index=False)

# Tarayıcıyı kapatma
driver.quit()
