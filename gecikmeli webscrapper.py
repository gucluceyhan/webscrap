from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import random
import os

# Excel dosyasından depo isimlerini çekme
excel_file_path = '/Users/gucluceyhan/Documents/EczaDepoları/Ecza Depoları İsimleri.xlsx'
depo_df = pd.read_excel(excel_file_path)

# Depo isimlerini liste olarak alma
depo_list = depo_df['Depo'].tolist()

# Selenium için driver kurulum
chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
chromedriver_path = os.getenv('CHROMEDRIVER_PATH', '/usr/local/bin/chromedriver')
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Web adresleri için boş liste
websites = []

# Her bir depo için Bing'de arama yap
for depo in depo_list:
    driver.get("https://www.bing.com")
    search_box = driver.find_element(By.NAME, "q")
    search_box.send_keys(f"{depo} web adresi")
    search_box.send_keys(Keys.RETURN)
    
    # Sayfanın yüklenmesini beklemek için WebDriverWait kullanma
    try:
        first_result = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'li.b_algo h2 a'))
        )
        website = first_result.get_attribute('href')
        websites.append(website)
    except Exception as e:
        websites.append('Not Found')
        print("CAPTCHA tespit edildi veya sayfa yüklenemedi. CAPTCHA'ı manuel olarak geçin ve devam etmek için ENTER tuşuna basın.")
        input("Devam etmek için ENTER tuşuna basın...")  # Kullanıcıdan ENTER tuşuna basmasını bekler

# Sonuçları DataFrame'e aktar
depo_df['Web Adresi'] = websites

# DataFrame'i Excel dosyasına yazma
excel_output_path = '/Users/gucluceyhan/Documents/EczaDepoları/EczaDepolari_Web_Adresleri.xlsx'
depo_df.to_excel(excel_output_path, index=False)

# Tarayıcıyı kapatma
driver.quit()
