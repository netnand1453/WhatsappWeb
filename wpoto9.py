# Bu kod, yapay zeka araçları kullanılarak geliştirilmiştir ve yazarın daha önce herhangi bir kod yazma deneyimi bulunmamaktadır.
# Kodlama kusurları veya optimizasyon eksiklikleri içerebilir.
# İşletim Sistemi: Windows 10
# Python Versiyonu: 3.13
# Gerekli Kütüphaneler: tkinter, pandas, selenium, webdriver_manager, openpyxl, pillow, pywin32

# This code was developed using artificial intelligence tools, and the author has no prior coding experience.
# It may contain coding imperfections or lack of optimization.
# Operating System: Windows 10
# Python Version: 3.13
# Required Libraries: tkinter, pandas, selenium, webdriver_manager, openpyxl, pillow, pywin32

# Yüklemek için / To install:
# pip install pandas openpyxl selenium webdriver_manager pillow pywin32

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import time
import os
import threading
import numpy as np
from datetime import datetime
import json  # Ayarları kaydetmek/yüklemek için
import tempfile  # Geçici dosya oluşturmak için

# Selenium kütüphaneleri
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException, \
    StaleElementReferenceException
from selenium.webdriver.common.keys import Keys

# WebDriver Manager kütüphaneleri
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# COM otomasyonu için (Excel'den resim kopyalama - Sadece Windows)
try:
    import win32com.client

    WINDOWS_COM_AVAILABLE = True
except ImportError:
    WINDOWS_COM_AVAILABLE = False
    # Bu uyarıyı sadece kütüphane eksikse göster
    # print("Uyarı: 'pywin32' kütüphanesi bulunamadı. Excel'den resim kopyalama özelliği kullanılamayacaktır.")
    # print("Bu özelliği kullanmak için 'pip install pywin32' komutunu çalıştırın.")


# Excel dosyasının varsayılan adı
DEFAULT_EXCEL_FILE = 'rehber.xlsx'
# Ayarlar dosyası adı
SETTINGS_FILE = 'settings.json'

# Kişi listesini saklamak için bir liste
contact_list = []

# Gönderme işlemini durdurmak için bir bayrak
stop_sending_flag = False
# Gönderme işlemini tamamen iptal etmek için bir bayrak
cancel_sending_flag = False

# Selenium WebDriver instance
driver = None

# WebDriver durumunu kontrol eden thread
check_driver_thread = None
# WebDriver kontrol thread'ini durdurmak için bayrak
stop_check_driver_thread = False

# Düzenlenmekte olan öğenin Treeview ID'si
editing_item_id = None

# Gönderme sonuçlarını saklamak için liste
sending_results = []

# Son oluşturulan raporun dosya yolu
last_report_path = None

# Gönderme işlemini yürüten thread
sending_thread = None

# Rapor dosyalarının kaydedileceği klasör (Ayarlardan yüklenecek)
report_directory = None

# Gönderim sırasında işlenen son kişinin indeksi (Devam Etmek için)
last_processed_index = -1

# Varsayılan XPath değerleri
DEFAULT_XPATH_SETTINGS = {
    'whatsapp_main_panel': '//*[@id="pane-side"]',  # Genel Ayarlar
    'search_button': '//*[@id="app"]/div/div[3]/div/div[3]/header/header/div/span/div/div[1]/button/span',
    # Genel Ayarlar
    'search_input': '//*[@id="app"]/div/div[3]/div/div[2]/div[1]/span/div/span/div/div[1]/div[2]/div/div/div[1]/p',
    # Genel Ayarlar
    'message_input_box': '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[3]/div[1]/p',  # Mesaj Gönderme
    'send_message_button': '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[4]/button/span',  # Mesaj Gönder Butonu
    'attach_button': '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[1]/button/span',
    # Doküman/Medya Gönderme
    'document_option': '//*[@id="app"]/div/span[6]/div/ul/div/div/div[1]/li/div',  # Doküman Gönderme
    'document_file_input': '//*[@id="app"]/div/span[6]/div/ul/div/div/div[1]/li/div//input[@type="file"]',
    # Doküman Gönderme (Otomatik oluşturulan)
    'send_file_button': '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span',
    # Doküman Gönderme
    'media_option_li': '//*[@id="app"]/div/span[6]/div/ul/div/div/div[2]/li',  # Medya Gönderme
    'media_file_input': '//*[@id="app"]/div/span[6]/div/ul/div/div/div[2]/li//input[@type="file"]',
    # Medya Gönderme (Kullanıcı tarafından sağlanan li XPath'inin içindeki input)
    'media_preview_message_input': '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[1]/div[3]/div/div/div/div[1]/div[1]/p',
    # Medya Gönderme (Yeni önizleme mesaj input XPath'i)
    'send_media_button': '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span'
    # Medya Gönderme (Yeni gönder butonu XPath'i)
}

# Uygulama tarafından kullanılacak XPath değerleri (Ayarlardan yüklenecek)
xpath_settings = DEFAULT_XPATH_SETTINGS.copy()

# XPath Grupları Tanımı (Arayüzde kullanılacak)
XPATH_GROUPS = {
    "Genel Ayarlar": ['whatsapp_main_panel', 'search_button', 'search_input'],
    "Mesaj Gönderme": ['message_input_box', 'send_message_button'],
    "Doküman Gönderme": ['attach_button', 'document_option', 'document_file_input', 'send_file_button'],
    "Medya Gönderme": ['media_option_li', 'media_file_input', 'media_preview_message_input', 'send_media_button']
}

# XPath Anahtarları için Türkçe Açıklamalar
XPATH_TURKISH_LABELS = {
    'whatsapp_main_panel': 'WhatsApp Ana Panel',
    'search_button': 'Arama Butonu',
    'search_input': 'Arama Giriş Alanı',
    'message_input_box': 'Mesaj Giriş Alanı',
    'send_message_button': 'Mesaj Gönder Butonu',
    'attach_button': 'Ataç Butonu (Dosya/Medya)',
    'document_option': 'Doküman Seçeneği',
    'document_file_input': 'Doküman Dosya Girişi',
    'send_file_button': 'Doküman Gönder Butonu',
    'media_option_li': 'Medya Seçeneği (Liste Öğesi)',
    'media_file_input': 'Medya Dosya Girişi',
    'media_preview_message_input': 'Medya Önizleme Mesaj Girişi',
    'send_media_button': 'Medya Gönder Butonu'
}


# Ayarları yükleme fonksiyonu
def load_settings():
    """Ayarları settings.json dosyasından yükler."""
    global report_directory, xpath_settings, last_report_path
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                report_directory = settings.get('report_directory')
                loaded_xpaths = settings.get('xpath_settings', {})
                last_report_path = settings.get('last_report_path') # Son rapor yolunu yükle

                # Yüklenen XPath'leri mevcut ayarlarla birleştir
                # Yeni eklenen XPath'ler için varsayılan değerleri koru
                for key, default_value in DEFAULT_XPATH_SETTINGS.items():
                    xpath_settings[key] = loaded_xpaths.get(key, default_value)

                # Arayüzdeki alanları güncelle (Eğer arayüz elementleri oluşturulduysa)
                # Bu kontrol, load_settings'in arayüz oluşturulmadan önce çağrılması durumunda hata vermesini önler.
                if 'entry_report_directory' in globals() and entry_report_directory.winfo_exists():
                    entry_report_directory.delete(0, tk.END)
                    if report_directory:
                        entry_report_directory.insert(0, report_directory)

                # xpath_entries global sözlüğünü kullanarak tüm giriş alanlarını güncelle
                if 'xpath_entries' in globals():
                    for key, entry in xpath_entries.items():
                        if entry.winfo_exists():
                            entry.delete(0, tk.END)
                            entry.insert(0, xpath_settings.get(key, ''))

                if report_directory and os.path.isdir(report_directory):
                    root.after(0, update_status, f"Ayarlar yüklendi. Rapor klasörü: {report_directory}")
                elif report_directory:
                    root.after(0, update_status,
                               f"Ayarlar yüklendi ancak rapor klasörü bulunamadı veya geçersiz: {report_directory}. Lütfen yeni bir klasör seçin.")
                    report_directory = None  # Geçersiz klasör yolunu sıfırla
                else:
                    root.after(0, update_status, "Ayarlar yüklendi ancak rapor klasörü belirtilmemiş.")

                # Son raporu aç butonunun durumunu güncelle
                root.after(0, lambda: btn_open_last_report.config(state=tk.NORMAL if last_report_path and os.path.exists(last_report_path) else tk.DISABLED))


        except Exception as e:
            root.after(0, update_status, f"Hata: Ayarlar yüklenirken sorun oluştu: {e}")


# Ayarları kaydetme fonksiyonu
def save_settings():
    """Güncel ayarları settings.json dosyasına kaydeder."""
    global report_directory, xpath_settings, last_report_path
    # Arayüzdeki güncel değerleri al ve kaydet
    if 'entry_report_directory' in globals() and entry_report_directory.winfo_exists():
        report_directory = entry_report_directory.get().strip()

    # xpath_entries global sözlüğünü kullanarak tüm giriş alanlarındaki değerleri al
    if 'xpath_entries' in globals():
        for key, entry in xpath_entries.items():
            if entry.winfo_exists():
                xpath_settings[key] = entry.get().strip()

    settings = {
        'report_directory': report_directory,
        'xpath_settings': xpath_settings,
        'last_report_path': last_report_path # Son rapor yolunu kaydet
    }
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=4)  # Daha okunabilir olması için indent ekle
        root.after(0, update_status, "Ayarlar kaydedildi.")
    except Exception as e:
        root.after(0, update_status, f"Hata: Ayarlar kaydedilirken sorun oluştu: {e}")


# XPath ayarlarını varsayılana sıfırlama fonksiyonu
def reset_xpath_to_default():
    """XPath ayarlarını varsayılan değerlere sıfırlar ve arayüzü günceller."""
    global xpath_settings
    if messagebox.askokcancel("Varsayılana Sıfırla",
                              "XPath ayarlarını varsayılan değerlere sıfırlamak istediğinizden emin misiniz?"):
        xpath_settings = DEFAULT_XPATH_SETTINGS.copy()
        # Arayüzdeki alanları güncelle
        # xpath_entries global sözlüğünü kullanarak tüm giriş alanlarını güncelle
        if 'xpath_entries' in globals():
            for key, entry in xpath_entries.items():
                if entry.winfo_exists():
                    entry.delete(0, tk.END)
                    entry.insert(0, xpath_settings.get(key, ''))
        root.after(0, update_status,
                   "XPath ayarları varsayılan değerlere sıfırlandı. Kaydetmek için 'Ayarları Kaydet' düğmesine tıklayın.")


# Rapor klasörünü seçme fonksiyonu
def browse_report_directory():
    """Kullanıcının rapor klasörünü seçmesini sağlar ve arayüzü günceller."""
    global report_directory
    directory = filedialog.askdirectory(title="Rapor Klasörünü Seçin")
    if directory:
        report_directory = directory
        if 'entry_report_directory' in globals() and entry_report_directory.winfo_exists():
            entry_report_directory.delete(0, tk.END)
            entry_report_directory.insert(0, report_directory)
        root.after(0, update_status, f"Rapor klasörü seçildi: {report_directory}")
        # Ayarları kaydetmek için save_settings() çağrılmalı, ancak kullanıcı "Ayarları Kaydet" butonuna basana kadar kaydetmeyelim.
        # save_settings() # Ayarı kaydet
    else:
        root.after(0, update_status, "Rapor klasörü seçimi iptal edildi.")


# WebDriver'ı başlatan ve WhatsApp Web'e bağlanan fonksiyon
def start_driver_and_connect_whatsapp():
    """Seçilen tarayıcı için WebDriver'ı başlatır ve WhatsApp Web'e bağlanır."""
    global driver, stop_check_driver_thread, check_driver_thread
    if driver:
        root.after(0, update_status, "WebDriver zaten çalışıyor.")
        return

    selected_browser = combo_browser.get()

    root.after(0, update_status, f"{selected_browser} WebDriver başlatılıyor ve otomatik olarak indiriliyor...")

    # Buton durumlarını güncelle (WebDriver başlatılırken)
    btn_start_and_connect.config(state=tk.DISABLED)
    combo_browser.config(state=tk.DISABLED)
    btn_start_sending.config(state=tk.DISABLED)  # Gönder butonu pasif
    btn_continue_sending.config(state=tk.DISABLED)  # Devam Et butonu pasif
    btn_stop.config(state=tk.DISABLED)
    btn_cancel.config(state=tk.DISABLED)
    btn_generate_report.config(state=tk.DISABLED)
    btn_open_last_report.config(state=tk.DISABLED)

    connect_thread = threading.Thread(target=_start_driver_and_connect_whatsapp_thread, args=(selected_browser,))
    connect_thread.start()


# ChromiumDriverManager sınıfı zaten tanımlı, tekrar tanımlamaya gerek yok.
# class ChromiumDriverManager:
#     pass


def _start_driver_and_connect_whatsapp_thread(selected_browser):
    """WebDriver başlatma ve WhatsApp Web'e bağlanma işlemini ayrı bir thread'de yürütür."""
    global driver, stop_check_driver_thread, check_driver_thread
    try:
        service = None
        options = webdriver.ChromeOptions() if selected_browser == "Chrome" else \
            webdriver.FirefoxOptions() if selected_browser == "Firefox" else \
                webdriver.EdgeOptions() if selected_browser == "Edge" else None

        if options is None:
            root.after(0, update_status, "Geçersiz tarayıcı seçimi.")
            root.after(0, messagebox.showwarning, "Tarayıcı Hatası", "Lütfen geçerli bir tarayıcı seçin.")
            # Buton durumlarını sıfırla
            root.after(0, lambda: btn_start_and_connect.config(state=tk.NORMAL))
            root.after(0, lambda: combo_browser.config(state="readonly"))
            root.after(0, lambda: btn_start_sending.config(
                state=tk.NORMAL if contact_list else tk.DISABLED))  # Gönder butonu aktif (liste boş değilse)
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
            return

        try:
            if selected_browser == "Chrome":
                service = ChromeService(ChromeDriverManager().install())
            elif selected_browser == "Firefox":
                service = FirefoxService(GeckoDriverManager().install())
            elif selected_browser == "Edge":
                # EdgeChromiumDriverManager'ı doğrudan kullan
                service = EdgeService(EdgeChromiumDriverManager().install())
        except Exception as e:
            root.after(0, update_status,
                       f"Hata: {selected_browser} WebDriver indirilirken/başlatılırken sorun oluştu: {e}")
            root.after(0, messagebox.showerror, "WebDriver Hatası",
                       f"{selected_browser} WebDriver indirilirken/başlatılırken sorun oluştu:\n{e}\nLütfen internet bağlantınızı kontrol edin ve güvenlik duvarınızın engellemediğinden emin olun.")
            # Buton durumlarını sıfırla
            root.after(0, lambda: btn_start_and_connect.config(state=tk.NORMAL))
            root.after(0, lambda: combo_browser.config(state="readonly"))
            root.after(0, lambda: btn_start_sending.config(
                state=tk.NORMAL if contact_list else tk.DISABLED))  # Gönder butonu aktif (liste boş değilse)
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
            return

        user_data_dir = os.path.join(os.path.expanduser("~"),
                                     f".whatsapp_automation_profile_{selected_browser.lower()}")
        options.add_argument(f"user-data-dir={user_data_dir}")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox") # Docker gibi ortamlarda gerekebilir
        options.add_argument("--disable-dev-shm-usage") # Docker gibi ortamlarda gerekebilir


        if selected_browser == "Chrome":
            driver = webdriver.Chrome(service=service, options=options)
        elif selected_browser == "Firefox":
            driver = webdriver.Firefox(service=service, options=options)
        elif selected_browser == "Edge":
            driver = webdriver.Edge(service=service, options=options)

        root.after(0, update_status,
                   f"{selected_browser} WebDriver başarıyla başlatıldı. WhatsApp Web'e yönlendiriliyor...")

        # WhatsApp Web'e gitmeye çalış, hata olursa yakala
        try:
            driver.get('https://web.whatsapp.com/')
            root.after(0, update_status,
                       "WhatsApp Web tarayıcı penceresi açıldı. Lütfen tarayıcıdan QR kodu okutarak giriş yapın. Giriş yaptıktan sonra 'Mesajları Gönder' düğmesine tıklayabilirsiniz.")

            # WebDriver başarıyla başlatıldı ve siteye gidildi, buton durumlarını güncelle
            root.after(0, lambda: btn_start_and_connect.config(
                state=tk.NORMAL))  # Başlat butonu tekrar aktif (yeniden başlatmak için)
            root.after(0, lambda: combo_browser.config(state="readonly"))  # Tarayıcı seçimi aktif
            root.after(0, lambda: btn_start_sending.config(
                state=tk.NORMAL if contact_list else tk.DISABLED))  # Gönder butonu aktif (liste boş değilse)
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif

            # WebDriver durumunu kontrol eden thread'i başlat
            stop_check_driver_thread = False
            check_driver_thread = threading.Thread(target=_check_driver_thread, name="check_driver_thread")
            check_driver_thread.start()


        except WebDriverException as e:
            root.after(0, update_status, f"WhatsApp Web'e bağlanılırken hata oluştu: {e}")
            root.after(0, messagebox.showerror, "Bağlantı Hatası",
                       f"WhatsApp Web'e bağlanılırken hata oluştu:\n{e}\nLütfen internet bağlantınızı kontrol edin ve WhatsApp Web'in erişilebilir olduğundan emin olun.")
            # Hata durumunda driver'ı kapat ve buton durumlarını sıfırla
            if driver:
                driver.quit()
                driver = None
            root.after(0, lambda: btn_start_and_connect.config(state=tk.NORMAL))
            root.after(0, lambda: combo_browser.config(state="readonly"))
            root.after(0, lambda: btn_start_sending.config(
                state=tk.NORMAL if contact_list else tk.DISABLED))  # Gönder butonu aktif (liste boş değilse)
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif


    except WebDriverException as e:
        root.after(0, update_status, f"{selected_browser} WebDriver başlatılırken hata oluştu: {e}")
        root.after(0, messagebox.showerror, "WebDriver Hatası",
                   f"{selected_browser} WebDriver başlatılırken hata oluştu:\n{e}\nLütfen internet bağlantınızı kontrol edin, güvenlik duvarınızın indirmeyi engellemediğinden, seçtiğiniz tarayıcının sisteminizde yüklü olduğundan ve WhatsApp Web'in erişilebilir olduğundan emin olun.\n\nDeneyebileceğiniz Çözümler:\n1. Tarayıcınızın güncel olduğundan emin olun.\n2. Uygulamayı yönetici olarak çalıştırmayı deneyin.\n3. Antivirüs/güvenlik duvarı yazılımınızı geçici olarak devre dışı bırakarak tekrar deneyin (dikkatli olun).\n4. Farklı bir tarayıcı seçerek deneyin.\n5. Eğer sorun devam ederse, {selected_browser} WebDriver cache'ini manuel olarak temizlemeyi deneyin. Bu genellikle kullanıcı dizininizdeki .wdm klasöründedir.")
        # Hata durumunda buton durumlarını sıfırla
        root.after(0, lambda: btn_start_and_connect.config(state=tk.NORMAL))
        root.after(0, lambda: combo_browser.config(state="readonly"))
        root.after(0, lambda: btn_start_sending.config(
            state=tk.NORMAL if contact_list else tk.DISABLED))  # Gönder butonu aktif (liste boş değilse)
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif


    except Exception as e:
        root.after(0, update_status, f"Beklenmeyen bir hata oluştu: {e}")
        root.after(0, messagebox.showerror, "Hata", f"Beklenmeyen bir hata oluştu:\n{e}")
        # Hata durumunda buton durumlarını sıfırla
        root.after(0, lambda: btn_start_and_connect.config(state=tk.NORMAL))
        root.after(0, lambda: combo_browser.config(state="readonly"))
        root.after(0, lambda: btn_start_sending.config(
            state=tk.NORMAL if contact_list else tk.DISABLED))  # Gönder butonu aktif (liste boş değilse)
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif


# WebDriver durumunu periyodik olarak kontrol eden thread fonksiyonu
def _check_driver_thread():
    """WebDriver'ın hala aktif olup olmadığını periyodik olarak kontrol eder."""
    global driver, stop_check_driver_thread, sending_thread, stop_sending_flag, cancel_sending_flag
    root.after(0, update_status, "WebDriver durum kontrolü başlatıldı.")
    while not stop_check_driver_thread:
        time.sleep(5)  # Her 5 saniyede bir kontrol et
        if driver:
            try:
                # WebDriver'ın hala aktif olup olmadığını kontrol et
                driver.current_url  # Bu bir kontrol mekanizmasıdır.

            except WebDriverException:
                root.after(0, update_status, "WebDriver bağlantısı kesildi veya tarayıcı kapatıldı.")
                # Eğer bir gönderim işlemi devam ediyorsa, onu da durdur/iptal et
                if sending_thread and sending_thread.is_alive():
                    cancel_sending_flag = True  # İşlemi iptal et
                    stop_sending_flag = True  # Durdurma bayrağını set et
                    root.after(0, update_status,
                               "Devam eden gönderim işlemi WebDriver bağlantısı kesildiği için iptal ediliyor.")
                    # Thread'in bitmesini beklemeye gerek yok, kendi kendine bitecektir.

                # Driver instance'ını temizle
                driver = None
                root.after(0, update_status, "WebDriver kapatıldı.")
                # Buton durumlarını güncelle (Bağlantı kesildi)
                root.after(0, lambda: btn_start_and_connect.config(state=tk.NORMAL))  # Başlat butonu aktif
                root.after(0, lambda: combo_browser.config(state="readonly"))  # Tarayıcı seçimi aktif
                root.after(0, lambda: btn_start_sending.config(
                    state=tk.DISABLED))  # Gönder butonu pasif (Yeni gönderim başlatılamaz)
                root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
                root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
                root.after(0, lambda: btn_cancel.config(state=tk.DISABLED))
                root.after(0, lambda: btn_generate_report.config(
                    state=tk.NORMAL if sending_results else tk.DISABLED))  # Eğer sonuç varsa rapor aktif
                root.after(0, lambda: btn_open_last_report.config(
                    state=tk.NORMAL if last_report_path and os.path.exists(
                        last_report_path) else tk.DISABLED))  # Eğer son rapor varsa ve mevcutsa aktif

                break  # Kontrol thread'ini sonlandır

            except Exception as e:
                root.after(0, update_status, f"WebDriver durum kontrolünde beklenmeyen hata: {e}")
                pass  # Kontrole devam et


        elif not driver and not stop_check_driver_thread:
            # Eğer driver None olmuşsa (kapatılmış veya hata olmuşsa) ve thread durdurulmadıysa,
            # bu thread'i de sonlandır.
            root.after(0, update_status, "WebDriver durumu kontrol thread'i sonlandırılıyor.")
            break


# Arayüzü güncellemek için Treeview'ı dolduran fonksiyon
def update_contact_treeview():
    """Kişi listesi Treeview'ını günceller."""
    for item in tree_contacts.get_children():
        tree_contacts.delete(item)

    for i, contact in enumerate(contact_list):
        if isinstance(contact, dict):
            # Treeview'da gösterilecek değerleri İşlem Türüne göre belirle
            display_mesaj = ''
            display_hedef_hucre = ''
            display_dosya_yolu = ''


            islem_turu = contact.get('İşlem Türü', '')

            if islem_turu == 'Mesaj':
                display_mesaj = contact.get('Mesaj', '')
            elif islem_turu == 'Doküman':
                display_dosya_yolu = contact.get('Dosya Yolu', '')
                display_mesaj = contact.get('Mesaj', '') # Doküman için opsiyonel mesaj
            elif islem_turu == 'Medya':
                display_dosya_yolu = contact.get('Dosya Yolu', '')
                display_mesaj = contact.get('Mesaj', '') # Medya için alt yazı Mesaj alanından alınacak
            elif islem_turu == 'Excel to Media':
                display_dosya_yolu = contact.get('Dosya Yolu', '') # Excel Dosya Yolu
                display_hedef_hucre = contact.get('Hedef Hücre', '')
                display_mesaj = contact.get('Mesaj', '') # Excel görseli için alt yazı buradan alınacak


            tree_contacts.insert('', tk.END, values=(
                i + 1,
                contact.get('Ünvan', ''),
                contact.get('Alan', ''),
                contact.get('Numara', ''), # Numara veya Grup Adı
                islem_turu,
                display_mesaj,
                display_hedef_hucre,
                display_dosya_yolu,
            ))
        else:
            root.after(0, update_status,
                       f"Uyarı: contact_list içinde sözlük olmayan bir öğe bulundu: {contact} (Tip: {type(contact)})")


# Rapor Treeview'ını güncelleyen fonksiyon
def update_report_treeview():
    """Rapor sonuçları Treeview'ını günceller."""
    for item in tree_report.get_children():
        tree_report.delete(item)

    for i, result in enumerate(sending_results):
        if isinstance(result, dict):
            tree_report.insert('', tk.END, values=(
                i + 1,
                result.get('Ünvan', ''),
                result.get('Numara', ''), # Numara veya Grup Adı
                result.get('İşlem Türü', ''),
                result.get('Durum', ''),
                result.get('Hata Nedeni', '')
            ))
        else:
            root.after(0, update_status,
                       f"Uyarı: sending_results içinde sözlük olmayan bir öğe bulundu: {result} (Tip: {type(result)})")


# Giriş alanlarını temizleyen fonksiyon
def clear_input_fields():
    """Kişi bilgi giriş alanlarını temizler."""
    entry_unvan.delete(0, tk.END)
    entry_alan.delete(0, tk.END)
    entry_numara.delete(0, tk.END)
    combo_islem_turu.set('')
    entry_mesaj.delete("1.0", tk.END)  # Mesaj / Excel Alt Yazısı
    entry_hedef_hucre.delete(0, tk.END)  # Hedef Hücre
    entry_dosya_yolu.delete(0, tk.END)  # Dosya Yolu (Excel veya Medya)
    update_input_field_visibility() # Alanları temizledikten sonra görünürlüklerini de sıfırla


# İşlem Türü seçimine göre giriş alanlarının görünürlüğünü güncelleyen fonksiyon
def update_input_field_visibility(*args):
    """
    Seçilen işlem türüne göre ilgili giriş alanlarının görünürlüğünü ve etiketlerini günceller.
    """
    selected_type = combo_islem_turu.get()

    # Tüm alanları başlangıçta gizle
    lbl_mesaj.grid_remove()
    entry_mesaj.grid_remove()
    lbl_hedef_hucre.grid_remove()
    entry_hedef_hucre.grid_remove()
    lbl_dosya_yolu.grid_remove()
    entry_dosya_yolu.grid_remove()
    btn_browse_file.grid_remove()

    # İlgili alanları göster ve etiketleri güncelle
    if selected_type == "Mesaj":
        lbl_mesaj.config(text="Mesaj:")
        lbl_mesaj.grid(row=4, column=0, padx=5, pady=2, sticky="w")
        entry_mesaj.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
    elif selected_type == "Doküman":
        lbl_mesaj.config(text="Mesaj (Opsiyonel Başlık):")
        lbl_mesaj.grid(row=4, column=0, padx=5, pady=2, sticky="w")
        entry_mesaj.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        lbl_dosya_yolu.config(text="Doküman Dosya Yolu:")
        lbl_dosya_yolu.grid(row=5, column=0, padx=5, pady=2, sticky="w")
        entry_dosya_yolu.grid(row=5, column=1, padx=5, pady=2, sticky="ew")
        btn_browse_file.grid(row=5, column=2, padx=5, pady=2)
    elif selected_type == "Medya":
        lbl_mesaj.config(text="Medya Alt Yazısı (Opsiyonel):")
        lbl_mesaj.grid(row=4, column=0, padx=5, pady=2, sticky="w")
        entry_mesaj.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        lbl_dosya_yolu.config(text="Medya Dosya Yolu:")
        lbl_dosya_yolu.grid(row=5, column=0, padx=5, pady=2, sticky="w")
        entry_dosya_yolu.grid(row=5, column=1, padx=5, pady=2, sticky="ew")
        btn_browse_file.grid(row=5, column=2, padx=5, pady=2)
    elif selected_type == "Excel to Media":
        lbl_mesaj.config(text="Görsel Alt Yazısı (Opsiyonel):")
        lbl_mesaj.grid(row=4, column=0, padx=5, pady=2, sticky="w")
        entry_mesaj.grid(row=4, column=1, padx=5, pady=2, sticky="ew")
        lbl_hedef_hucre.grid(row=5, column=0, padx=5, pady=2, sticky="w")
        entry_hedef_hucre.grid(row=5, column=1, padx=5, pady=2, sticky="ew")
        lbl_dosya_yolu.config(text="Excel Dosya Yolu:")
        lbl_dosya_yolu.grid(row=6, column=0, padx=5, pady=2, sticky="w")
        entry_dosya_yolu.grid(row=6, column=1, padx=5, pady=2, sticky="ew")
        btn_browse_file.grid(row=6, column=2, padx=5, pady=2)

    # Grid konumlarını yeniden düzenle (boşlukları kapatmak için)
    # Bu, dinamik görünürlük sonrası boş satırları engeller.
    # Tüm widget'ları yeniden konumlandırmak yerine, sadece görünür olanları doğru sıraya koymak daha verimli olabilir.
    # Ancak burada basitlik adına, görünür olanları yeniden gridliyoruz.
    current_row = 4
    if lbl_mesaj.winfo_ismapped():
        lbl_mesaj.grid(row=current_row, column=0, padx=5, pady=2, sticky="w")
        entry_mesaj.grid(row=current_row, column=1, padx=5, pady=2, sticky="ew")
        current_row += 1
    if lbl_hedef_hucre.winfo_ismapped():
        lbl_hedef_hucre.grid(row=current_row, column=0, padx=5, pady=2, sticky="w")
        entry_hedef_hucre.grid(row=current_row, column=1, padx=5, pady=2, sticky="ew")
        current_row += 1
    if lbl_dosya_yolu.winfo_ismapped():
        lbl_dosya_yolu.grid(row=current_row, column=0, padx=5, pady=2, sticky="w")
        entry_dosya_yolu.grid(row=current_row, column=1, padx=5, pady=2, sticky="ew")
        btn_browse_file.grid(row=current_row, column=2, padx=5, pady=2)
        current_row += 1


# Kişi ekleme veya kaydetme fonksiyonu
def save_contact():
    """Giriş alanlarındaki bilgiyi kişi listesine ekler veya günceller."""
    global editing_item_id

    unvan = entry_unvan.get().strip()
    alan = entry_alan.get().strip()
    numara = entry_numara.get().strip() # Bu alan numara veya grup adı olabilir
    islem_turu = combo_islem_turu.get().strip()
    mesaj_input = entry_mesaj.get("1.0", tk.END).strip()  # Mesaj alanındaki girdi
    hedef_hucre_input = entry_hedef_hucre.get().strip()  # Hedef Hücre alanındaki girdi
    dosya_yolu_input = entry_dosya_yolu.get().strip()  # Dosya Yolu alanındaki girdi


    # Ünvan ve Numara/Grup Adı boş bırakılamaz (Alan boş bırakılabilir)
    if not unvan or not numara:
        messagebox.showwarning("Eksik Bilgi", "Ünvan ve Numara/Grup Adı alanları boş bırakılamaz.")
        return

    if not islem_turu:
        messagebox.showwarning("Eksik Bilgi",
                               "Lütfen bir İşlem Türü seçin (Mesaj, Doküman, Medya veya Excel to Media).")
        return

    # İşlem türüne göre ek kontroller ve veri hazırlığı
    contact_data = {
        'Ünvan': unvan,
        'Alan': alan, # Alan kodu boş bırakılabilir
        'Numara': numara, # Numara veya Grup Adı
        'İşlem Türü': islem_turu,
        'Mesaj': '', # Varsayılan olarak boş
        'Hedef Hücre': '', # Varsayılan olarak boş
        'Dosya Yolu': '', # Varsayılan olarak boş
    }

    if islem_turu == "Mesaj":
        if not mesaj_input:
             messagebox.showwarning("Eksik Bilgi", "İşlem Türü 'Mesaj' seçildi ancak Mesaj alanı boş.")
             return
        contact_data['Mesaj'] = mesaj_input

    elif islem_turu == "Doküman":
        if not dosya_yolu_input:
             messagebox.showwarning("Eksik Bilgi", "İşlem Türü 'Doküman' seçildi ancak Dosya Yolu alanı boş.")
             return
        contact_data['Dosya Yolu'] = dosya_yolu_input
        contact_data['Mesaj'] = mesaj_input # Doküman için opsiyonel mesaj

    elif islem_turu == "Medya":
        if not dosya_yolu_input:
             messagebox.showwarning("Eksik Bilgi", "İşlem Türü 'Medya' seçildi ancak Dosya Yolu alanı boş.")
             return
        contact_data['Dosya Yolu'] = dosya_yolu_input
        contact_data['Mesaj'] = mesaj_input # Medya için alt yazı Mesaj alanından alınacak

    elif islem_turu == "Excel to Media":
        if not WINDOWS_COM_AVAILABLE:
             messagebox.showerror("Uyarı", "'Excel to Media' işlemi için 'pywin32' kütüphanesi gereklidir ve bulunamadı. Lütfen bu kütüphaneyi yükleyin (pip install pywin32).")
             return
        if not dosya_yolu_input: # Excel Dosya Yolu
             messagebox.showwarning("Eksik Bilgi", "İşlem Türü 'Excel to Media' seçildi ancak Excel Dosya Yolu alanı boş.")
             return
        if not hedef_hucre_input: # Hedef Hücre
             messagebox.showwarning("Eksik Bilgi", "İşlem Türü 'Excel to Media' seçildi ancak Hedef Hücre alanı boş.")
             return
        # Temel hücre formatı kontrolü (örn: A1, B10)
        if not (hedef_hucre_input and hedef_hucre_input[0].isalpha() and hedef_hucre_input[1:].isdigit()):
             messagebox.showwarning("Geçersiz Hücre Formatı", "Hedef Hücre formatı geçersiz görünüyor. Lütfen 'A1' gibi geçerli bir hücre adresi girin.")
             return

        contact_data['Dosya Yolu'] = dosya_yolu_input # Excel Dosya Yolu
        contact_data['Hedef Hücre'] = hedef_hucre_input # Hedef Hücre
        contact_data['Mesaj'] = mesaj_input # Excel görseli için alt yazı buradan alınacak


    if editing_item_id:
        item_index = tree_contacts.index(editing_item_id)
        if 0 <= item_index < len(contact_list):
            contact_list[item_index] = contact_data
            root.after(0, update_status, f"{unvan} kişisi güncellendi.")
        else:
            root.after(0, update_status, f"Hata: Düzenlenecek kişi listede bulunamadı.")
            messagebox.showerror("Düzenleme Hatası", "Düzenlenecek kişi listede bulunamadı.")

        editing_item_id = None
        btn_save_contact.config(text="Kişi Ekle")
        btn_cancel_edit.grid_forget()

    else:
        contact_list.append(contact_data)
        root.after(0, update_status, f"{unvan} kişisi eklendi.")

    update_contact_treeview()
    clear_input_fields()
    # Kişi eklendiğinde veya güncellendiğinde gönder butonu aktif olabilir (WebDriver bağlıysa)
    if driver:
        root.after(0, lambda: btn_start_sending.config(state=tk.NORMAL))


# Kişi silme fonksiyonu
def delete_contact():
    """Seçili kişileri listeden siler."""
    selected_item = tree_contacts.selection()
    if not selected_item:
        messagebox.showwarning("Seçim Yok", "Lütfen silmek istediğiniz kişiyi listeden seçin.")
        return

    indices_to_delete = []
    for item in selected_item:
        item_index_in_treeview = tree_contacts.index(item)
        if 0 <= item_index_in_treeview < len(contact_list):
            indices_to_delete.append(item_index_in_treeview)

    indices_to_delete.sort(reverse=True)
    for i in indices_to_delete:
        del contact_list[i]

    update_contact_treeview()

    messagebox.showinfo("Silindi", "Seçilen kişi(ler) silindi.")
    cancel_edit()
    # Liste boşaldığında gönder butonu pasif olmalı
    if not contact_list and driver:
        root.after(0, lambda: btn_start_sending.config(state=tk.DISABLED))


# Listeyi temizleme fonksiyonu (Yeni)
def clear_contact_list():
    """Mevcut kişi listesini tamamen temizler."""
    global contact_list, last_processed_index
    if messagebox.askokcancel("Listeyi Temizle",
                              "Mevcut kişi listesini tamamen temizlemek istediğinizden emin misiniz?"):
        contact_list.clear()
        update_contact_treeview()
        last_processed_index = -1  # Liste temizlenince devam etme indeksi sıfırlanmalı
        root.after(0, update_status, "Kişi listesi temizlendi.")
        messagebox.showinfo("Liste Temizlendi", "Kişi listesi başarıyla temizlendi.")
        cancel_edit()  # Düzenleme modunu sıfırla
        # Liste temizlenince Gönder aktif, Devam Et pasif olmalı (Eğer WebDriver bağlıysa)
        if driver:
            root.after(0, lambda: btn_start_sending.config(
                state=tk.NORMAL if contact_list else tk.DISABLED))  # Liste boşsa gönder pasif
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))
        else:
            root.after(0, lambda: btn_start_sending.config(state=tk.DISABLED))
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))


# Listeyi içe aktırma fonksiyonu (Excel veya CSV) - Mevcut liste temizleme eklendi
def import_list():
    """Excel veya CSV dosyasından kişi listesini içe aktarır."""
    global last_processed_index  # İçe aktarma yapıldığında devam etme indeksi sıfırlanmalı
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Dosyaları", "*.xlsx *.xlsm"), ("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*")]
    )
    if not file_path:
        return

    try:
        if file_path.lower().endswith(('.xlsx', '.xlsm')):
            # Excel dosyasını olduğu gibi oku, boş hücreleri NaN olarak bırak
            df = pd.read_excel(file_path, dtype=str, keep_default_na=False) # dtype=str ve keep_default_na=False ekle
        elif file_path.lower().endswith('.csv'):
            # CSV dosyasını olduğu gibi oku, boş hücreleri boş string olarak al
            df = pd.read_csv(file_path, dtype=str, keep_default_na=False) # dtype=str ve keep_default_na=False ekle
        else:
            messagebox.showerror("Format Hatası", "Desteklenmeyen dosya formatı.")
            root.after(0, update_status, "Hata: Desteklenmeyen içe aktarma dosya formatı.")
            return

        # Beklenen sütunlar (Yeni yapıyı yansıtacak şekilde güncellendi)
        expected_columns = ['Ünvan', 'Alan', 'Numara', 'İşlem Türü', 'Mesaj', 'Hedef Hücre', 'Dosya Yolu']

        # Dosyada olmayan sütunları boş olarak ekle
        for col in expected_columns:
            if col not in df.columns:
                df[col] = ''

        # Sadece beklenen sütunları al ve sıralamayı ayarla
        df = df[expected_columns]

        # NaN değerlerini ve None'ları boş string'e dönüştür (daha sonra kontrol edilecek)
        # dtype=str ve keep_default_na=False ile okuduğumuz için bu adımın çoğu gereksizleşir
        # Ancak yine de ekstra güvenlik için yapılabilir.
        df = df.fillna('') # NaN değerlerini boş string ile doldur

        # Mevcut listeyi temizle
        contact_list.clear()
        update_contact_treeview()  # Arayüzü de temizle
        last_processed_index = -1  # Liste temizlenince devam etme indeksi sıfırlanmalı
        root.after(0, update_status, "Mevcut kişi listesi içe aktarma öncesi temizlendi.")

        # DataFrame'i sözlük listesine dönüştür
        new_contacts = df.to_dict('records')

        # Ekstra kontrol: Tüm değerlerin string olduğundan emin ol
        for contact in new_contacts:
            for key, value in contact.items():
                 contact[key] = str(value).strip() # Her değeri string'e çevir ve boşlukları temizle


        contact_list.extend(new_contacts)
        update_contact_treeview()

        messagebox.showinfo("İçe Aktarıldı",
                            f"{len(new_contacts)} kişi listeye eklendi. Lütfen 'Kişileri Kontrol Et' butonuna tıklayarak olası hataları kontrol edin.")
        root.after(0, update_status,
                   f"{len(new_contacts)} kişi başarıyla içe aktarıldı. Lütfen 'Kişileri Kontrol Et' butonuna tıklayarak olası hataları kontrol edin.")
        cancel_edit()
        # İçe aktarma sonrası Gönder aktif, Devam Et pasif olmalı (Eğer WebDriver bağlıysa ve liste boş değilse)
        if driver and contact_list:
            root.after(0, lambda: btn_start_sending.config(state=tk.NORMAL))
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))
        else:
            root.after(0, lambda: btn_start_sending.config(state=tk.DISABLED))
            root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))


    except FileNotFoundError:
        messagebox.showerror("Hata", "Dosya bulunamadı.")
        root.after(0, update_status, "Hata: İçe aktarılacak dosya bulunamadı.")
    except Exception as e:
        messagebox.showerror("Hata", f"Dosya okunurken bir hata oluştu: {e}")
        root.after(0, update_status, f"Hata: Dosya okunurken bir hata oluştu: {e}")


# Kişileri kontrol eden fonksiyon (Yeni)
def check_contacts():
    """Kişi listesindeki girişleri temel geçerlilik kurallarına göre kontrol eder."""
    if not contact_list:
        messagebox.showwarning("Boş Liste", "Kontrol edilecek kişi bulunamadı.")
        root.after(0, update_status, "Uyarı: Kontrol edilecek kişi bulunamadı.")
        return

    root.after(0, update_status, "Kişi listesi kontrol ediliyor...")
    checked_results = []  # Bu kontrolün sonuçları için geçici liste

    for i, contact in enumerate(contact_list):
        unvan = contact.get('Ünvan', '').strip()
        alan = str(contact.get('Alan', '')).strip()  # String'e çevir ve strip yap
        numara_veya_grup = str(contact.get('Numara', '')).strip()  # String'e çevir ve strip yap
        islem_turu = contact.get('İşlem Türü', '').strip()
        mesaj = contact.get('Mesaj', '').strip()  # Mesaj (Excel alt yazısı buradan alınacak)
        hedef_hucre = contact.get('Hedef Hücre', '').strip()  # Hedef Hücre
        dosya_yolu = contact.get('Dosya Yolu', '').strip()  # Dosya Yolu (Excel veya Medya)


        is_valid = True
        validation_errors = []

        # Temel alanların boş olup olmadığını kontrol et (Alan boş olabilir)
        if not unvan:
            is_valid = False
            validation_errors.append("Ünvan boş")
        if not numara_veya_grup:
            is_valid = False
            validation_errors.append("Numara/Grup Adı boş")

        # Alan ve Numara/Grup Adının format kontrolü
        if alan and not alan.lstrip('+').isdigit():  # Alan kodu '+' ile başlayabilir ve sayısal olmalı (boşsa kontrol yok)
            is_valid = False
            validation_errors.append(f"Geçersiz Alan Kodu: '{alan}' (Sayısal olmalı veya boş bırakılmalı)")
        # Numara/Grup Adı için sadece boş kontrolü yapıldı, sayısal olma zorunluluğu kaldırıldı.


        # İşlem türü ve ilgili alanların kontrolü
        valid_islem_turleri = ["Mesaj", "Doküman", "Medya", "Excel to Media"]
        if not islem_turu:
            is_valid = False
            validation_errors.append("İşlem Türü boş veya geçersiz")
        elif islem_turu not in valid_islem_turleri:
            is_valid = False
            validation_errors.append(f"Geçersiz İşlem Türü: '{islem_turu}'")


        if islem_turu == "Mesaj" and not mesaj:
            is_valid = False
            validation_errors.append("İşlem Türü 'Mesaj' seçildi ancak Mesaj alanı boş.")

        if islem_turu == "Doküman":
            if not dosya_yolu:
                is_valid = False
                validation_errors.append("İşlem Türü 'Doküman' ancak Dosya Yolu boş.")
            elif dosya_yolu and not os.path.exists(dosya_yolu):
                is_valid = False
                validation_errors.append(f"Dosya bulunamadı: '{dosya_yolu}'")
            # Mesaj (opsiyonel başlık) kontrolü yapılmıyor, boş olabilir

        if islem_turu == "Medya":
            if not dosya_yolu:
                is_valid = False
                validation_errors.append("İşlem Türü 'Medya' ancak Dosya Yolu boş.")
            elif dosya_yolu and not os.path.exists(dosya_yolu):
                is_valid = False
                validation_errors.append(f"Dosya bulunamadı: '{dosya_yolu}'")
            # Mesaj (alt yazı) opsiyoneldir, kontrol etmeye gerek yok

        if islem_turu == "Excel to Media":
            if not WINDOWS_COM_AVAILABLE:
                is_valid = False
                validation_errors.append("'Excel to Media' için 'pywin32' kütüphanesi eksik.")
            if not dosya_yolu:  # Excel Dosya Yolu
                is_valid = False
                validation_errors.append("İşlem Türü 'Excel to Media' ancak Excel Dosya Yolu boş.")
            elif dosya_yolu and not os.path.exists(dosya_yolu):
                is_valid = False
                validation_errors.append(f"Excel dosyası bulunamadı: '{dosya_yolu}'")
            if not hedef_hucre:  # Hedef Hücre
                is_valid = False
                validation_errors.append("İşlem Türü 'Excel to Media' seçildi ancak Hedef Hücre boş.")
            elif not (hedef_hucre and hedef_hucre[0].isalpha() and hedef_hucre[1:].isdigit()):
                is_valid = False
                validation_errors.append(f"Geçersiz Hedef Hücre formatı: '{hedef_hucre}' (örn: A1)")
            # Mesaj (Excel görseli alt yazısı) opsiyoneldir, kontrol etmeye gerek yok


        if not is_valid:
            checked_results.append({
                'Ünvan': unvan,
                'Alan': alan,
                'Numara': numara_veya_grup, # Raporlamada orijinal değeri göster
                'İşlem Türü': islem_turu,
                'Mesaj': mesaj,
                'Hedef Hücre': hedef_hucre,
                'Dosya Yolu': dosya_yolu,
                'Durum': 'Kontrol Başarısız',
                'Hata Nedeni': ', '.join(validation_errors)
            })
            root.after(0, update_status,
                       f"[{i + 1}/{len(contact_list)}] {unvan} ({alan}{numara_veya_grup}) kontrol başarısız: {', '.join(validation_errors)}")


    # Kontrol sonuçlarını kullanıcıya göstermek için Durum Raporu'na yazıldı.
    # Gönderim sonuçları listesine eklenmiyor.

    if checked_results:
        root.after(0, update_status,
                   f"Kişi kontrolü tamamlandı. {len(checked_results)} kişide hata bulundu. Detaylar Durum Raporu'nda görülebilir.")
        messagebox.showwarning("Kontrol Tamamlandı",
                               f"Kişi kontrolü tamamlandı. {len(checked_results)} kişide hata bulundu. Lütfen Durum Raporu'nu kontrol edin.")
    else:
        root.after(0, update_status, "Kişi kontrolü tamamlandı. Tüm kişiler geçerli görünüyor.")
        messagebox.showinfo("Kontrol Tamamlandı", "Kişi kontrolü tamamlandı. Tüm kişiler geçerli görünüyor.")


# Listeyi dışa aktırma fonksiyonu (Excel veya CSV)
def export_list():
    """Mevcut kişi listesini Excel veya CSV dosyası olarak dışa aktarır."""
    if not contact_list:
        messagebox.showwarning("Boş Liste", "Dışa aktarılacak kişi bulunamadı.")
        root.after(0, update_status, "Uyarı: Dışa aktarılacak kişi bulunamadı.")
        return

    # Rapor formatı seçimi için popup
    export_format = messagebox.askquestion("Dışa Aktırma Formatı", "Listeyi Excel olarak mı dışa aktarmak istersiniz?",
                                           type='yesnocancel', default='yes')

    if export_format == 'cancel':
        root.after(0, update_status, "Dışa aktarma iptal edildi.")
        return

    if export_format == 'yes':
        file_types = [("Excel Dosyaları", "*.xlsx"), ("Tüm Dosyalar", "*.*")]
        default_ext = ".xlsx"
    else:  # export_format == 'no' (CSV)
        file_types = [("CSV Dosyaları", "*.csv"), ("Tüm Dosyalar", "*.*")]
        default_ext = ".csv"

    file_path = filedialog.asksaveasfilename(
        defaultextension=default_ext,
        filetypes=file_types,
        title="Listeyi Kaydet"
    )
    if not file_path:
        root.after(0, update_status, "Dışa aktarma kayıt konumu seçilmedi.")
        return

    try:
        df = pd.DataFrame(contact_list)

        if file_path.lower().endswith('.xlsx'):
            df.to_excel(file_path, index=False)
        elif file_path.lower().endswith('.csv'):
            df.to_csv(file_path, index=False, quoting=1)
        else:
            # Eğer kullanıcı farklı bir uzantı girdiyse, seçilen formata göre kaydet
            if default_ext == ".xlsx":
                df.to_excel(file_path + ".xlsx", index=False)
                file_path = file_path + ".xlsx"
            elif default_ext == ".csv":
                df.to_csv(file_path + ".csv", index=False, quoting=1)
                file_path = file_path + ".csv"
            root.after(0, update_status,
                       f"Uyarı: Desteklenmeyen dosya uzantısı girildi. '{default_ext}' olarak kaydedildi.")

        messagebox.showinfo("Dışa Aktarıldı", f"Liste başarıyla dışa aktarıldı: {file_path}")
        root.after(0, update_status, f"Liste başarıyla dışa aktarıldı: {file_path}")

    except Exception as e:
        messagebox.showerror("Hata", f"Dosya yazılırken bir hata oluştu: {e}")
        root.after(0, update_status, f"Hata: Dosya yazılırken bir hata oluştu: {e}")


# Dosya yolu seçme fonksiyonu (Hem medya/doküman hem de Excel için kullanılacak)
def browse_file():
    """Dosya seçme penceresini açar ve seçilen yolu dosya yolu giriş alanına yazar."""
    # Seçili işlem türüne göre dosya tipi filtrelemesi yapılabilir
    selected_islem_turu = combo_islem_turu.get().strip()
    if selected_islem_turu == "Excel to Media":
        file_types = [("Excel Dosyaları", "*.xlsx *.xlsm"), ("Tüm Dosyalar", "*.*")]
        title = "Excel Dosyasını Seçin"
    elif selected_islem_turu in ["Doküman", "Medya"]:
        file_types = [("Tüm Dosyalar", "*.*"), ("Resim Dosyaları", "*.jpg *.jpeg *.png *.gif"),
                      ("PDF Dosyaları", "*.pdf")]
        title = f"{selected_islem_turu} Dosyasını Seçin"
    else:
        file_types = [("Tüm Dosyalar", "*.*")]
        title = "Dosya Seçin"

    file_path = filedialog.askopenfilename(filetypes=file_types, title=title)
    if file_path:
        entry_dosya_yolu.delete(0, tk.END)
        entry_dosya_yolu.insert(0, file_path)


# Kişiyi düzenlemek için seçme fonksiyonu
def select_contact_for_edit(event):
    """Treeview'da bir kişi seçildiğinde, bilgilerini giriş alanlarına yükler."""
    global editing_item_id
    selected_items = tree_contacts.selection()
    if not selected_items:
        cancel_edit()
        return

    editing_item_id = selected_items[0]
    # Treeview'daki değerler, update_contact_treeview'da belirlenen display değerleridir.
    # Gerçek contact_list objesinden değerleri almalıyız.
    item_index = tree_contacts.index(editing_item_id)
    if 0 <= item_index < len(contact_list):
        contact = contact_list[item_index]
    else:
        root.after(0, update_status, f"Hata: Düzenlenecek kişi listede bulunamadı.")
        messagebox.showerror("Düzenleme Hatası", "Düzenlenecek kişi listede bulunamadı.")
        cancel_edit()
        return


    clear_input_fields()
    entry_unvan.insert(0, contact.get('Ünvan', ''))
    entry_alan.insert(0, contact.get('Alan', ''))
    entry_numara.insert(0, contact.get('Numara', '')) # Numara veya Grup Adı
    islem_turu = contact.get('İşlem Türü', '')
    combo_islem_turu.set(islem_turu)
    update_input_field_visibility() # İşlem türü seçildikten sonra alanları güncelle

    # İşlem Türüne göre alanları doldur
    if islem_turu == 'Mesaj':
        entry_mesaj.insert("1.0", contact.get('Mesaj', ''))
    elif islem_turu == 'Doküman':
        entry_dosya_yolu.insert(0, contact.get('Dosya Yolu', ''))
        entry_mesaj.insert("1.0", contact.get('Mesaj', '')) # Doküman için opsiyonel mesaj
    elif islem_turu == 'Medya':
        entry_dosya_yolu.insert(0, contact.get('Dosya Yolu', ''))
        entry_mesaj.insert("1.0", contact.get('Mesaj', '')) # Medya için alt yazı Mesaj alanından alınacak
    elif islem_turu == 'Excel to Media':
        entry_dosya_yolu.insert(0, contact.get('Dosya Yolu', '')) # Excel Dosya Yolu
        entry_hedef_hucre.insert(0, contact.get('Hedef Hücre', ''))
        entry_mesaj.insert("1.0", contact.get('Mesaj', '')) # Excel görseli için alt yazı buradan alınacak


    btn_save_contact.config(text="Değişiklikleri Kaydet")
    btn_cancel_edit.grid(row=0, column=1, padx=5, pady=5)


# Düzenleme modunu iptal etme fonksiyonu
def cancel_edit():
    """Kişi düzenleme modunu iptal eder ve giriş alanlarını temizler."""
    global editing_item_id
    editing_item_id = None
    clear_input_fields()
    btn_save_contact.config(text="Kişi Ekle")
    btn_cancel_edit.grid_forget()
    tree_contacts.selection_remove(tree_contacts.selection())


# Durum raporu alanını güncelleyen fonksiyon (thread-safe)
def update_status(message):
    """Durum raporu metin alanına mesaj ekler (thread güvenli)."""
    status_report.insert(tk.END, message + "\n")
    status_report.see(tk.END)
    # Tek satırlık durum çubuğunu da güncelle
    status_bar_label.config(text=message)


# Excel'deki bir aralığı JPEG olarak kaydeden fonksiyon (Sadece Windows ve Excel gerektirir)
def _excel_range_to_jpeg(excel_path, target_cell_address, output_jpeg_path):
    """
    Belirtilen Excel dosyasındaki hedef hücrenin CurrentRegion'ını
    JPEG formatında belirtilen yola kaydeder.
    Sadece Windows işletim sisteminde ve Microsoft Excel yüklüyse çalışır.
    """
    if not WINDOWS_COM_AVAILABLE:
        return False, "pywin32 kütüphanesi yüklü değil veya Windows değil."

    excel_app = None
    workbook = None
    try:
        # Excel uygulamasını başlat
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False  # Excel penceresini gizle
        excel_app.DisplayAlerts = False  # Uyarıları kapat

        # Çalışma kitabını aç
        workbook = excel_app.Workbooks.Open(os.path.abspath(excel_path), ReadOnly=True)
        sheet = workbook.Sheets(1)  # İlk sayfayı al (isteğe bağlı olarak sayfa adı parametresi eklenebilir)

        # Hedef hücreyi ve CurrentRegion'ını al
        # target_range = sheet.Range(target_cell_address).CurrentRegion # Orijinal satır
        # Kullanıcının girdiği hücre adresini doğrudan kullan
        target_range = sheet.Range(target_cell_address)
        # Sonra CurrentRegion'ı al
        target_range = target_range.CurrentRegion

        # Aralığı resim olarak kopyala
        target_range.CopyPicture(Appearance=1, Format=2)  # Appearance=xlScreen, Format=xlPicture

        # Geçici bir grafik nesnesi oluştur
        # Boyutları ayarlamak VBA'daki gibi karmaşık olabilir, basit bir başlangıç boyutu verelim
        chart_object = sheet.ChartObjects().Add(Left=0, Top=0, Width=target_range.Width, Height=target_range.Height)
        chart_object.Activate()

        # Kopyalanan resmi grafiğe yapıştır
        chart = chart_object.Chart
        chart.Paste()

        # Grafiği JPEG olarak dışa aktar
        chart.Export(output_jpeg_path, "JPEG")

        # Grafik nesnesini sil
        chart_object.Delete()

        return True, "Başarılı"

    except Exception as e:
        return False, f"Excel'den resim kopyalama hatası: {e}"

    finally:
        # Temizlik
        if workbook:
            try:
                workbook.Close(False)  # Değişiklikleri kaydetmeden kapat
            except Exception as e:
                print(f"Workbook kapatılırken hata: {e}")
        if excel_app:
            try:
                excel_app.Quit()
            except Exception as e:
                print(f"Excel uygulaması kapatılırken hata: {e}")
        # COM objelerini serbest bırak (bellek sızıntısını önlemek için)
        excel_app = None
        workbook = None
        import gc
        gc.collect()


# Mesaj gönderme işlemini başlatan fonksiyon (Yeni gönderim başlatır)
def start_sending():
    """Mesaj gönderme işlemini başlatır (yeni bir gönderim)."""
    global stop_sending_flag, cancel_sending_flag, sending_results, sending_thread, last_processed_index

    if not driver:
        root.after(0, update_status,
                   "WebDriver başlatılmadı. Lütfen önce 'WebDriver Başlat ve Bağlan' düğmesine tıklayın.")
        messagebox.showwarning("WebDriver Hatası", "WebDriver başlatılmadı veya WhatsApp Web'e bağlanılmadı.")
        return

    # XPath ayarlarının yüklendiğinden emin ol
    if not xpath_settings:
        root.after(0, update_status,
                   "Hata: XPath ayarları yüklenemedi veya boş. Lütfen Ayarlar sekmesini kontrol edin.")
        messagebox.showerror("Ayarlar Hatası",
                             "XPath ayarları yüklenemedi veya boş. Lütfen Ayarlar sekmesini kontrol edin ve ayarları kaydedin.")
        return

    whatsapp_main_panel_xpath = xpath_settings.get('whatsapp_main_panel')
    if not whatsapp_main_panel_xpath:
        root.after(0, update_status, "Hata: WhatsApp ana panel XPath ayarı boş. Lütfen Ayarlar sekmesini kontrol edin.")
        messagebox.showerror("Ayarlar Hatası",
                             "WhatsApp ana panel XPath ayarı boş. Lütfen Ayarlar sekmesini kontrol edin.")
        return

    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, whatsapp_main_panel_xpath))
        )
        root.after(0, update_status, "WhatsApp Web ana arayüzü hazır.")
    except TimeoutException:
        root.after(0, update_status,
                   "Hata: WhatsApp Web ana arayüzü henüz yüklenmedi. Lütfen tarayıcıdan QR kodu okutarak giriş yapın ve tekrar deneyin.")
        messagebox.showwarning("WhatsApp Web Hazır Değil",
                               "WhatsApp Web ana arayüzü henüz yüklenmedi. Lütfen tarayıcıdan QR kodu okutarak giriş yapın ve tekrar deneyin.")
        return

    if not contact_list:
        messagebox.showwarning("Boş Liste", "Gönderilecek kişi bulunamadı.")
        return

    # Eğer gönderim thread'i zaten çalışıyorsa, yeni bir işlem başlatma
    if sending_thread and sending_thread.is_alive():
        messagebox.showwarning("İşlem Devam Ediyor", "Mesaj gönderme işlemi zaten devam ediyor.")
        return

    # Yeni bir gönderim başlat
    stop_sending_flag = False
    cancel_sending_flag = False
    sending_results = []  # Yeni gönderim için sonuç listesini sıfırla
    update_report_treeview()  # Rapor Treeview'ını temizle
    last_processed_index = -1  # Yeni gönderim için son işlenen indeksi sıfırla
    root.after(0, update_status, "Mesaj gönderme işlemi başlatıldı...")

    sending_thread = threading.Thread(target=_send_messages_thread, name="sending_thread")
    sending_thread.start()

    # Buton durumlarını güncelle (Yeni gönderim başlarken)
    btn_start_sending.config(state=tk.DISABLED)  # Gönder butonu pasif
    btn_continue_sending.config(state=tk.DISABLED)  # Devam Et butonu pasif
    btn_stop.config(state=tk.NORMAL)  # Durdur butonu aktif
    btn_cancel.config(state=tk.NORMAL)  # İptal Et butonu aktif
    btn_generate_report.config(state=tk.DISABLED)
    btn_open_last_report.config(state=tk.DISABLED)


# Mesaj gönderme işlemini duraklatılan yerden devam ettiren fonksiyon
def continue_sending():
    """Duraklatılan mesaj gönderme işlemine devam eder."""
    global stop_sending_flag, cancel_sending_flag, sending_thread, last_processed_index

    if not driver:
        root.after(0, update_status,
                   "WebDriver başlatılmadı. Lütfen önce 'WebDriver Başlat ve Bağlan' düğmesine tıklayın.")
        messagebox.showwarning("WebDriver Hatası", "WebDriver başlatılmadı veya WhatsApp Web'e bağlanılmadı.")
        return

    # XPath ayarlarının yüklendiğinden emin ol
    if not xpath_settings:
        root.after(0, update_status,
                   "Hata: XPath ayarları yüklenemedi veya boş. Lütfen Ayarlar sekmesini kontrol edin.")
        messagebox.showerror("Ayarlar Hatası",
                             "XPath ayarları yüklenemedi veya boş. Lütfen Ayarlar sekmesini kontrol edin.")
        return

    if sending_thread and sending_thread.is_alive():
        messagebox.showwarning("İşlem Devam Ediyor", "Mesaj gönderme işlemi zaten devam ediyor.")
        return

    # last_processed_index -1 ise veya liste sonuna ulaşıldysa devam edilemez
    if last_processed_index == -1 or last_processed_index >= len(contact_list) - 1:
        messagebox.showwarning("Devam Edilemez", "Devam edilecek bir işlem bulunamadı veya liste sonuna ulaşıldı.")
        root.after(0, update_status, "Uyarı: Devam edilecek bir işlem bulunamadı veya liste sonuna ulaşıldı.")
        # Duraklatılmış durumda kalmışsa butonları sıfırla
        root.after(0, lambda: btn_start_sending.config(
            state=tk.NORMAL if driver and contact_list else tk.DISABLED))  # Gönder butonu aktif
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
        root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
        root.after(0, lambda: btn_cancel.config(state=tk.DISABLED))
        root.after(0, lambda: btn_generate_report.config(
            state=tk.NORMAL if sending_results else tk.DISABLED))  # Rapor butonu aktif
        root.after(0, lambda: btn_open_last_report.config(state=tk.NORMAL if last_report_path and os.path.exists(
            last_report_path) else tk.DISABLED))  # Son Raporu Aç aktif (mevcutsa)
        return

    stop_sending_flag = False  # Duraklatma bayrağını kaldır
    cancel_sending_flag = False  # İptal bayrağını kaldır
    root.after(0, update_status,
               f"Mesaj gönderme işlemine {last_processed_index + 2}. kişiden devam ediliyor...")  # Kullanıcıya 1 tabanlı indeks göster

    sending_thread = threading.Thread(target=_send_messages_thread, name="sending_thread")
    sending_thread.start()

    # Buton durumlarını güncelle (Devam ederken)
    btn_start_sending.config(state=tk.DISABLED)  # Gönder butonu pasif
    btn_continue_sending.config(state=tk.DISABLED)  # Devam Et butonu pasif
    btn_stop.config(state=tk.NORMAL)  # Durdur butonu aktif
    btn_cancel.config(state=tk.NORMAL)  # İptal Et butonu aktif
    btn_generate_report.config(state=tk.DISABLED)
    btn_open_last_report.config(state=tk.DISABLED)


# Gönderme işlemini duraklatan fonksiyon
def stop_sending():
    """Mesaj gönderme işlemini duraklatır."""
    global stop_sending_flag
    stop_sending_flag = True
    root.after(0, update_status,
               "Mesaj gönderme işlemi duraklatılıyor... Lütfen mevcut işlemin tamamlanmasını bekleyin.")
    # Buton durumlarını güncelle (Duraklatılırken)
    btn_stop.config(state=tk.DISABLED)
    btn_cancel.config(state=tk.NORMAL)
    # Devam Et butonu thread durduktan sonra aktif edilecek


# Gönderme işlemini tamamen iptal eden fonksiyon
def cancel_sending():
    """Mesaj gönderme işlemini tamamen iptal eder."""
    global cancel_sending_flag, stop_sending_flag
    if messagebox.askokcancel("Gönderimi İptal Et",
                              "Gönderim işlemini tamamen iptal etmek istiyor musunuz? Mevcut duruma göre rapor oluşturulacaktır."):
        cancel_sending_flag = True
        stop_sending_flag = True  # İptal ederken duraklatma bayrağını da set et
        root.after(0, update_status,
                   "Mesaj gönderme işlemi iptal ediliyor... Lütfen mevcut işlemin tamamlanmasını ve raporun oluşturulmasını bekleyin.")
        # Buton durumlarını güncelle (İptal edilirken)
        btn_stop.config(state=tk.DISABLED)
        btn_cancel.config(state=tk.DISABLED)
        # Gönder butonu thread durduktan sonra aktif edilecek, Devam Et butonu pasif kalacak.


# Rapor oluşturan fonksiyon (Otomatik kayıt ve klasör seçimi eklendi)
def generate_report():
    """Mevcut gönderim sonuçlarından Excel formatında rapor oluşturur."""
    global sending_results, last_report_path, report_directory

    if not sending_results:
        messagebox.showwarning("Boş Rapor", "Oluşturulacak rapor verisi bulunamadı.")
        root.after(0, update_status, "Uyarı: Oluşturulacak rapor verisi bulunamadı.")
        return

    # Rapor klasörü belirlenmediyse, uygulamanın çalıştığı dizinde 'Reports' klasörünü kullan
    if report_directory is None or not os.path.isdir(report_directory):
        default_report_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Reports")
        # Klasör yoksa oluştur
        if not os.path.exists(default_report_dir):
            try:
                os.makedirs(default_report_dir)
                root.after(0, update_status, f"Varsayılan rapor klasörü oluşturuldu: {default_report_dir}")
            except Exception as e:
                root.after(0, update_status, f"Hata: Varsayılan rapor klasörü oluşturulamadı: {e}")
                messagebox.showerror("Klasör Hatası", f"Varsayılan rapor klasörü oluşturulamadı: {e}")
                return # Klasör oluşturulamazsa rapor oluşturma

        report_directory = default_report_dir
        root.after(0, update_status,
                   f"Rapor klasörü seçilmediği için varsayılan klasör kullanılıyor: {report_directory}")
        # Arayüzdeki giriş alanını da güncelle
        if 'entry_report_directory' in globals() and entry_report_directory.winfo_exists():
            entry_report_directory.delete(0, tk.END)
            entry_report_directory.insert(0, report_directory)

    # Rapor dosyası adı için tarih ve saat
    now = datetime.now()
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")

    filename = f"Gönderim_Raporu_{timestamp}.xlsx"  # Varsayılan Excel formatı
    file_path = os.path.join(report_directory, filename)

    try:
        df = pd.DataFrame(sending_results)
        # Excel formatına kaydet
        df.to_excel(file_path, index=False)
        last_report_path = file_path
        root.after(0, update_status, f"Rapor başarıyla oluşturuldu: {file_path}")

        # Ayarları kaydet (son rapor yolunu güncellemek için)
        save_settings()

        root.after(0, lambda: btn_generate_report.config(state=tk.NORMAL))
        root.after(0, lambda: btn_open_last_report.config(
            state=tk.NORMAL if last_report_path and os.path.exists(last_report_path) else tk.DISABLED))

    except Exception as e:
        messagebox.showerror("Rapor Hatası", f"Rapor oluşturulurken bir hata oluştu: {e}")
        root.after(0, update_status, f"Hata: Rapor oluşturulurken hata oluştu: {e}")
        root.after(0, lambda: btn_generate_report.config(
            state=tk.NORMAL if sending_results else tk.DISABLED))  # Hata olsa da eğer sonuç varsa rapor butonu aktif kalsın


# Son oluşturulan raporu açan fonksiyon
def open_last_report():
    """Son oluşturulan rapor dosyasını açar."""
    global last_report_path
    # last_report_path değeri load_settings veya generate_report tarafından belirlenir ve save_settings ile kaydedilir.
    if last_report_path and os.path.exists(last_report_path):
        try:
            os.startfile(last_report_path)
            root.after(0, update_status, f"Son rapor açılıyor: {last_report_path}")
        except Exception as e:
            messagebox.showerror("Dosya Açma Hatası", f"Rapor dosyası açılamadı: {e}")
            root.after(0, update_status, f"Hata: Son rapor dosyası açılamadı: {e}")
    else:
        messagebox.showwarning("Rapor Bulunamadı", "Açılacak bir rapor bulunamadı.")
        root.after(0, update_status, "Uyarı: Açılacak bir rapor bulunamadı.")
        # Eğer dosya yolu kayıtlı ama dosya yoksa butonu pasif yap (normalde load_settings yapmalı)
        root.after(0, lambda: btn_open_last_report.config(state=tk.DISABLED))


# Rapor sonuçlarını temizleme fonksiyonu (Yeni)
def clear_report_results():
    """Mevcut rapor sonuçlarını temizler."""
    global sending_results, last_processed_index
    if messagebox.askokcancel("Raporu Temizle",
                              "Mevcut rapor sonuçlarını temizlemek istediğinizden emin misiniz? Bu işlem kaydedilmemiş rapor verilerini siler."):
        sending_results.clear()
        update_report_treeview()
        last_processed_index = -1  # Rapor temizlenince devam etme indeksi sıfırlanmalı (mantıksal olarak yeni bir gönderim başlatılacak)
        root.after(0, update_status, "Rapor sonuçları temizlendi.")
        messagebox.showinfo("Rapor Temizlendi", "Rapor sonuçları başarıyla temizlendi.")
        # Rapor temizlenince rapor oluştur butonu pasif olmalı (çünkü gönderim sonucu yok)
        root.after(0, lambda: btn_generate_report.config(state=tk.DISABLED))
        # Son Raporu Aç butonu pasif OLMAMALI, çünkü son kaydedilen rapor hala diskte olabilir.
        # Durumu load_settings veya generate_report tarafından belirlenmeli.
        # root.after(0, lambda: btn_open_last_report.config(state=tk.DISABLED)) # Bu satır kaldırıldı

# Mesaj gönderme işlemini gerçekleştiren fonksiyon (XPath değerleri settings'ten alınacak)
def _send_messages_thread():
    """Mesaj gönderme işlemini ayrı bir thread'de yürütür."""
    global driver, sending_results, stop_sending_flag, cancel_sending_flag, sending_thread, last_processed_index, xpath_settings
    wait = WebDriverWait(driver, 20)

    # XPath değerlerini global xpath_settings sözlüğünden al
    search_button_xpath = xpath_settings.get('search_button')
    search_input_xpath = xpath_settings.get('search_input')
    message_input_box_xpath = xpath_settings.get('message_input_box')
    send_message_button_xpath = xpath_settings.get('send_message_button')
    attach_button_xpath = xpath_settings.get('attach_button')
    document_option_xpath = xpath_settings.get('document_option')
    document_file_input_xpath = xpath_settings.get('document_file_input')
    send_file_button_xpath = xpath_settings.get('send_file_button')
    media_option_li_xpath = xpath_settings.get('media_option_li')
    media_file_input_xpath = xpath_settings.get('media_file_input')
    media_preview_message_input_xpath = xpath_settings.get('media_preview_message_input')
    send_media_button_xpath = xpath_settings.get('send_media_button')
    whatsapp_main_panel_xpath = xpath_settings.get('whatsapp_main_panel')  # Ana panel XPath'i burada da kullanılabilir

    # XPath değerlerinin geçerli olup olmadığını kontrol et (Basit kontrol)
    if not all([search_button_xpath, search_input_xpath, message_input_box_xpath, send_message_button_xpath,
                attach_button_xpath, document_option_xpath, document_file_input_xpath, send_file_button_xpath,
                media_option_li_xpath, media_file_input_xpath, media_preview_message_input_xpath,
                send_media_button_xpath,
                whatsapp_main_panel_xpath]):
        root.after(0, update_status,
                   "Hata: Bir veya daha fazla XPath ayarı boş. Lütfen Ayarlar sekmesini kontrol edin.")
        root.after(0, messagebox.showerror("Ayarlar Hatası",
                                           "Bir veya daha fazla XPath ayarı boş. Lütfen Ayarlar sekmesini kontrol edin ve ayarları kaydedin."))
        # Hata durumunda buton durumlarını sıfırla
        root.after(0, lambda: btn_start_sending.config(state=tk.NORMAL if driver and contact_list else tk.DISABLED))
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))
        root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
        root.after(0, lambda: btn_cancel.config(state=tk.DISABLED))
        root.after(0, lambda: btn_generate_report.config(state=tk.NORMAL if sending_results else tk.DISABLED))
        root.after(0, lambda: btn_open_last_report.config(
            state=tk.NORMAL if last_report_path and os.path.exists(last_report_path) else tk.DISABLED))
        return

    # Arama butonunu bul ve tıkla (İlk arama için veya hata sonrası dönmek için)
    # Bu butona tıklamak, arama giriş alanını görünür yapar.
    try:
        search_button = wait.until(EC.presence_of_element_located((By.XPATH, search_button_xpath)))
        # Eğer arama alanı zaten açıksa tekrar tıklamak kapatabilir, bu yüzden kontrol edilebilir.
        # Ancak basitlik için her zaman tıklayalım ve hata olursa yakalayalım.
        search_button.click()
        root.after(0, update_status, "Arama butonu tıklandı.")
        time.sleep(1)  # Bekleme süresi 1 saniye
    except TimeoutException:
        root.after(0, update_status,
                   "Uyarı: Arama butonu bulunamadı veya tıklanamadı. Arama alanı zaten açık olabilir.")
        time.sleep(1)  # Yine de bekleme süresi
    except Exception as e:
        root.after(0, update_status, f"Beklenmeyen hata oluştu: Arama butonu bulunurken hata - {e}")
        # Hata durumunda rapor oluştur ve thread'i sonlandır
        root.after(0, generate_report)
        root.after(0, update_status, "Gönderim işlemi hata nedeniyle sonlandı.")
        # Buton durumlarını sıfırla
        root.after(0, lambda: btn_start_sending.config(
            state=tk.NORMAL if driver and contact_list else tk.DISABLED))  # Eğer driver bağlıysa ve liste boş değilse gönder aktif
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
        root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
        root.after(0, lambda: btn_cancel.config(state=tk.DISABLED))
        root.after(0, lambda: btn_generate_report.config(state=tk.NORMAL if sending_results else tk.DISABLED))
        root.after(0, lambda: btn_open_last_report.config(
            state=tk.NORMAL if last_report_path and os.path.exists(last_report_path) else tk.DISABLED))
        return

    # Gönderime başlanacak indeks belirleme
    # Eğer last_processed_index -1 ise 0'dan başla, aksi takdirde last_processed_index + 1'den başla
    start_index = last_processed_index + 1

    for i in range(start_index, len(contact_list)):
        # Durdurma veya İptal bayrağını kontrol et
        if stop_sending_flag:
            last_processed_index = i - 1  # Duraklatılan kişinin indeksi (döngü bir sonraki adımda artacağı için i-1)
            if cancel_sending_flag:
                root.after(0, update_status, "Mesaj gönderme işlemi kullanıcı tarafından iptal edildi.")
            else:
                root.after(0, update_status,
                           f"Mesaj gönderme işlemi kullanıcı tarafından duraklatıldı. {i}. kişide durdu.")
            print(f"Mesaj gönderme thread'i durduruldu/iptal edildi. Son işlenen indeks: {last_processed_index}")
            break  # Döngüyü kır

        contact = contact_list[i]

        # Hata kontrolü: contact'ın bir sözlük olduğundan emin olalım
        if not isinstance(contact, dict):
            root.after(0, update_status,
                       f"Uyarı: contact_list içinde sözlük olmayan bir öğe bulundu. Atlanıyor: {contact} (Tip: {type(contact)})")
            sending_results.append({'Ünvan': 'Hata', 'Alan': '', 'Numara': '', 'İşlem Türü': '', 'Durum': 'Başarısız',
                                    'Hata Nedeni': 'Geçersiz kişi formatı'})
            update_report_treeview()
            last_processed_index = i  # Bu kişiyi işledik (atlanmış olsa da)
            continue

        unvan = contact.get('Ünvan', 'Bilinmeyen Ünvan')
        alan = str(contact.get('Alan', '')).strip()  # String'e çevir ve strip yap
        numara_veya_grup = str(contact.get('Numara', '')).strip()  # String'e çevir ve strip yap
        # Mesaj, Hedef Hücre, Dosya Yolu değerlerini contact_list'ten alıyoruz
        mesaj_data = str(contact.get('Mesaj', '')).strip() # Mesaj veya Alt Yazı buradan alınacak
        hedef_hucre_data = str(contact.get('Hedef Hücre', '')).strip()
        dosya_yolu_data = str(contact.get('Dosya Yolu', '')).strip()
        islem_turu_original = str(
            contact.get('İşlem Türü', '')).strip()  # Orijinal işlem türü (String'e çevir ve strip yap)

        # Arama terimini belirle: Alan boşsa sadece Numara/Grup Adı, doluysa Alan + Numara
        search_term = numara_veya_grup if not alan else f"{alan}{numara_veya_grup}".strip()

        # İşlem türüne göre ön kontrol (Gönderme öncesi temel kontrol)
        # Bu kontroller check_contacts fonksiyonunda yapılıyor, burada tekrar yapmak gereksiz olabilir
        # ancak gönderim anında bir kez daha kontrol etmek hata riskini azaltır.
        # Basitçe, eğer işlem türü geçerli değilse veya gerekli alanlar boşsa atla.
        valid_islem_turleri = ["Mesaj", "Doküman", "Medya", "Excel to Media"]
        if islem_turu_original not in valid_islem_turleri:
            root.after(0, update_status,
                       f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Geçersiz işlem türü)")
            sending_results.append(
                {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                 'Durum': 'Başarısız', 'Hata Nedeni': 'Geçersiz veya boş işlem türü'})
            update_report_treeview()
            last_processed_index = i  # Bu kişiyi işledik (atlanmış olsa da)
            continue  # Bu kişiyi atla

        if islem_turu_original == "Mesaj" and not mesaj_data:
            root.after(0, update_status,
                       f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Mesaj boş)")
            sending_results.append(
                {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                 'Durum': 'Başarısız', 'Hata Nedeni': 'Mesaj alanı boş'})
            update_report_treeview()
            last_processed_index = i  # Bu kişiyi işledik (atlanmış olsa da)
            continue  # Bu kişiyi atla

        if islem_turu_original in ["Doküman", "Medya"]:
            if not dosya_yolu_data or not os.path.exists(dosya_yolu_data):
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Dosya bulunamadı veya yol boş)")
                sending_results.append(
                    {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                     'Durum': 'Başarısız', 'Hata Nedeni': 'Dosya yolu boş veya dosya bulunamadı'})
                update_report_treeview()
                last_processed_index = i  # Bu kişiyi işledik (atlanmış olsa da)
                continue  # Bu kişiyi atla

        if islem_turu_original == "Excel to Media":
            if not WINDOWS_COM_AVAILABLE:
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız ('Excel to Media' için pywin32 eksik)")
                sending_results.append(
                    {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                     'Durum': 'Başarısız',
                     'Hata Nedeni': "'Excel to Media' işlemi için 'pywin32' kütüphanesi gereklidir."})
                update_report_treeview()
                last_processed_index = i
                continue
            if not dosya_yolu_data or not os.path.exists(dosya_yolu_data):  # Excel Dosya Yolu
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Excel dosyası bulunamadı veya yol boş)")
                sending_results.append(
                    {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                     'Durum': 'Başarısız', 'Hata Nedeni': 'Excel Dosya yolu boş veya dosya bulunamadı'})
                update_report_treeview()
                last_processed_index = i
                continue
            if not hedef_hucre_data:  # Hedef Hücre
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Hedef Hücre boş)")
                sending_results.append(
                    {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                     'Durum': 'Başarısız', 'Hata Nedeni': 'Hedef Hücre boş'})
                update_report_treeview()
                last_processed_index = i
                continue
            elif not (hedef_hucre_data and hedef_hucre_data[0].isalpha() and hedef_hucre_data[1:].isdigit()):
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Geçersiz Hedef Hücre formatı)")
                sending_results.append(
                    {'Ünvan': unvan, 'Alan': alan, 'Numara': numara_veya_grup, 'İşlem Türü': islem_turu_original,
                     'Durum': 'Başarısız', 'Hata Nedeni': 'Geçersiz Hedef Hücre formatı'})
                update_report_treeview()
                last_processed_index = i
                continue
            # Mesaj (Excel görseli alt yazısı) opsiyoneldir, kontrol etmeye gerek yok


        # --- Her kişi işlemi başlangıcında kısa bekleme ---
        time.sleep(2)  # UI'ın stabil hale gelmesi için ek bekleme
        # --- Bekleme Sonu ---

        root.after(0, update_status,
                   f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı.")  # Başlangıç durumu raporu

        temp_jpeg_path = None  # Excel to Media için geçici dosya yolu
        current_dosya_yolu_to_send = dosya_yolu_data  # Gönderilecek dosya yolu (Doküman/Medya/Excel)
        current_message_to_to_send = mesaj_data  # Gönderilecek mesaj (Mesaj, Doküman başlık, Medya/Excel alt yazı)

        # Her kişi için bir sonuç sözlüğü oluştur
        current_result = {
            'Ünvan': unvan,
            'Alan': alan,
            'Numara': numara_veya_grup,
            'İşlem Türü': islem_turu_original,
            'Durum': 'Başarısız', # Varsayılan olarak başarısız
            'Hata Nedeni': ''
        }

        try:
            # Eğer işlem türü Excel to Media ise, önce Excel'den resmi oluştur
            if islem_turu_original == "Excel to Media":
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] Excel'den görsel oluşturuluyor...") # Daha kısa duyuru
                # Geçici bir dosya oluştur
                temp_dir = tempfile.gettempdir()
                temp_filename = f"excel_img_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.jpg"
                temp_jpeg_path = os.path.join(temp_dir, temp_filename)

                success, error_msg = _excel_range_to_jpeg(dosya_yolu_data, hedef_hucre_data, temp_jpeg_path)

                if not success:
                    root.after(0, update_status,
                               f"[{i + 1}/{len(contact_list)}] Görsel oluşturma başarısız.") # Daha kısa duyuru
                    current_result['Hata Nedeni'] = f"Excel'den görsel oluşturma hatası: {error_msg}"
                    # Geçici dosyayı temizle (oluştuysa)
                    if temp_jpeg_path and os.path.exists(temp_jpeg_path):
                        try:
                            os.remove(temp_jpeg_path)
                        except Exception as cleanup_e:
                            root.after(0, update_status, f"Uyarı: Geçici dosya silinirken hata: {cleanup_e}")
                    raise Exception("Excel to Media oluşturma hatası") # Hata fırlat ki aşağıdaki except bloğuna düşsün


                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] Görsel başarıyla oluşturuldu, gönderime hazırlanılıyor.") # Daha kısa duyuru
                current_dosya_yolu_to_send = temp_jpeg_path  # Medya gönderme için geçici dosya yolunu kullan
                # Excel görseli için alt yazı Mesaj alanından alınacak ve aşağıda kullanılacak


            # Arama giriş alanını bul ve numarayı/grup adını yaz
            try:
                search_input = wait.until(EC.presence_of_element_located((By.XPATH, search_input_xpath)))
            except (TimeoutException, NoSuchElementException) as e:
                # Kişi veya grup bulunamadı hatası için özel mesaj
                current_result['Hata Nedeni'] = 'Kişi veya grup bulunamadı'
                root.after(0, update_status, f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Kişi veya grup bulunamadı)")
                raise Exception("Kişi/Grup bulunamadı hatası") # Hata fırlat


            search_input.clear()  # Önceki aramayı temizle
            search_input.send_keys(search_term)
            time.sleep(1.5)

            search_input.send_keys(Keys.ENTER)

            wait_for_chat = WebDriverWait(driver, 15)
            # Mesaj giriş kutusunun görünmesini bekle, bu sohbetin açıldığını gösterir.
            message_input_box = wait_for_chat.until(EC.presence_of_element_located((By.XPATH, message_input_box_xpath)))


            contact_found = True
            attachment_sent_successful = False
            message_sent_successful = False
            message_written_successful = False
            caption_written_successful = False  # Alt yazı bayrağı


            # --- İŞLEM TÜRÜNE GÖRE GÖNDERME ---

            if islem_turu_original == 'Doküman':
                if current_message_to_to_send:  # Doküman gönderirken de mesaj (başlık) olabilir
                    try:
                        current_message_input_box = wait.until(
                            EC.presence_of_element_located((By.XPATH, message_input_box_xpath)))
                        current_message_input_box.send_keys(current_message_to_to_send)
                        time.sleep(1)
                        message_written_successful = True
                    except TimeoutException:
                        current_result['Hata Nedeni'] += 'Mesaj giriş alanı bulunamadı (yazma öncesi); '
                        message_written_successful = False
                    except NoSuchElementException:
                        current_result['Hata Nedeni'] += 'Mesaj giriş alanı bulunamadı (NoSuchElementException); '
                        message_written_successful = False
                    except Exception as e:
                        current_result['Hata Nedeni'] += f'Mesaj yazma hatası: {e}; '
                        message_written_successful = False

                try:
                    attach_button = wait.until(EC.presence_of_element_located((By.XPATH, attach_button_xpath)))
                    attach_button.click()
                    time.sleep(1)

                    # Belge input elementini bulmak için document_option_xpath kullanılıyor
                    file_input_element = wait.until(
                        EC.presence_of_element_located((By.XPATH, document_file_input_xpath)))
                    file_input_element.send_keys(current_dosya_yolu_to_send)

                    wait_for_file_ready = WebDriverWait(driver, 15)
                    wait_for_file_ready.until(EC.presence_of_element_located((By.XPATH, send_file_button_xpath)))

                    send_file_button = wait.until(EC.presence_of_element_located((By.XPATH, send_file_button_xpath)))
                    send_file_button.click()
                    root.after(0, update_status,
                               f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için Doküman gönderimi tamamlandı.")
                    time.sleep(2)
                    attachment_sent_successful = True

                    if current_message_to_to_send:
                        message_sent_successful = message_written_successful
                    else:
                        message_sent_successful = True


                except TimeoutException:
                    current_result['Hata Nedeni'] += 'Doküman gönderme zaman aşımı; '
                except NoSuchElementException:
                    current_result['Hata Nedeni'] += 'Doküman gönderme elementi bulunamadı; '
                except Exception as file_send_e:
                    current_result['Hata Nedeni'] += f'Doküman gönderme hatası: {file_send_e}; '


            elif islem_turu_original == 'Medya': # Medya
                 try:
                     attach_button = wait.until(EC.presence_of_element_located((By.XPATH, attach_button_xpath)))
                     attach_button.click()
                     time.sleep(1)

                     # Medya input elementini bulmak için yeni XPath kullanılıyor
                     file_input_element = wait.until(EC.presence_of_element_located((By.XPATH, media_file_input_xpath)))

                     file_input_element.send_keys(current_dosya_yolu_to_send)

                     wait_for_attachment_ready = WebDriverWait(driver, 15)
                     # Medya önizleme mesaj giriş alanını ve gönder butonunu bekliyoruz.
                     wait_for_attachment_ready.until(EC.presence_of_element_located(
                         (By.XPATH, media_preview_message_input_xpath)))  # Yeni önizleme mesaj input XPath'i
                     wait_for_attachment_ready.until(EC.presence_of_element_located(
                         (By.XPATH, send_media_button_xpath)))  # Yeni gönder butonu XPath'i


                     caption_written_successful = True
                     if current_message_to_to_send:  # Medya gönderirken alt yazı olabilir (Mesaj alanından alınan)
                         try:
                             # Alt yazıyı yazmak için yeni önizleme mesaj input XPath'i kullanılıyor
                             preview_message_input = wait.until(
                                 EC.presence_of_element_located((By.XPATH, media_preview_message_input_xpath)))
                             preview_message_input.send_keys(current_message_to_to_send)
                             time.sleep(1)
                             caption_written_successful = True
                         except TimeoutException:
                             current_result['Hata Nedeni'] += 'Önizleme alt yazı giriş alanı bulunamadı (Medya); '
                             caption_written_successful = False
                         except NoSuchElementException:
                              current_result['Hata Nedeni'] += 'Önizleme alt yazı giriş alanı bulunamadı (Medya - NoSuchElementException); '
                              caption_written_successful = False
                         except Exception as e:
                              current_result['Hata Nedeni'] += f'Önizleme alt yazı yazma hatası (Medya): {e}; '
                              caption_written_successful = False


                     # Medya gönder butonu için yeni XPath kullanılıyor
                     send_button = wait.until(EC.presence_of_element_located((By.XPATH, send_media_button_xpath)))
                     send_button.click()
                     root.after(0, update_status,
                               f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için Medya gönderimi tamamlandı.")
                     time.sleep(2) # Medya gönderimi sonrası bekleme süresi
                     attachment_sent_successful = True

                     if current_message_to_to_send:
                          message_sent_successful = caption_written_successful # Alt yazı gönderimi başarılıysa mesaj gönderimi de başarılı sayılır (mantıksal olarak)
                     else:
                          message_sent_successful = True # Alt yazı yoksa ve dosya gönderimi başarılıysa başarılı sayılır

                 except TimeoutException:
                      current_result['Hata Nedeni'] += f"'Medya' gönderme zaman aşımı; "
                 except NoSuchElementException:
                      current_result['Hata Nedeni'] += f"'Medya' gönderme elementi bulunamadı; "
                 except Exception as file_send_e:
                      current_result['Hata Nedeni'] += f"'Medya' gönderme hatası: {file_send_e}; "


            elif islem_turu_original == 'Excel to Media': # Excel to Media
                 root.after(0, update_status, f"[{i + 1}/{len(contact_list)}] Görsel gönderiliyor...") # Daha kısa duyuru
                 try:
                     attach_button = wait.until(EC.presence_of_element_located((By.XPATH, attach_button_xpath)))
                     attach_button.click()
                     time.sleep(1)

                     # Medya input elementini bulmak için yeni XPath kullanılıyor
                     # Excel to Media için current_dosya_yolu_to_send geçici JPEG yolu olacak
                     file_input_element = wait.until(EC.presence_of_element_located((By.XPATH, media_file_input_xpath)))

                     file_input_element.send_keys(current_dosya_yolu_to_send)

                     wait_for_attachment_ready = WebDriverWait(driver, 15)
                     # Medya önizleme mesaj giriş alanını ve gönder butonunu bekliyoruz.
                     wait_for_attachment_ready.until(EC.presence_of_element_located(
                         (By.XPATH, media_preview_message_input_xpath)))  # Yeni önizleme mesaj input XPath'i
                     wait_for_attachment_ready.until(EC.presence_of_element_located(
                         (By.XPATH, send_media_button_xpath)))  # Yeni gönder butonu XPath'i


                     caption_written_successful = True
                     if current_message_to_to_send:  # Excel görseli gönderirken alt yazı olabilir (Mesaj alanından alınan)
                         try:
                             # Alt yazıyı yazmak için yeni önizleme mesaj input XPath'i kullanılıyor
                             preview_message_input = wait.until(
                                 EC.presence_of_element_located((By.XPATH, media_preview_message_input_xpath)))
                             preview_message_input.send_keys(current_message_to_to_send)
                             time.sleep(1)
                             caption_written_successful = True
                         except TimeoutException:
                             current_result['Hata Nedeni'] += 'Önizleme alt yazı giriş alanı bulunamadı (Excel to Media); '
                             caption_written_successful = False
                         except NoSuchElementException:
                              current_result['Hata Nedeni'] += 'Önizleme alt yazı giriş alanı bulunamadı (Excel to Media - NoSuchElementException); '
                              caption_written_successful = False
                         except Exception as e:
                              current_result['Hata Nedeni'] += f'Önizleme alt yazı yazma hatası (Excel to Media): {e}; '
                              caption_written_successful = False


                     # Medya gönder butonu için yeni XPath kullanılıyor
                     send_button = wait.until(EC.presence_of_element_located((By.XPATH, send_media_button_xpath)))
                     send_button.click()
                     root.after(0, update_status,
                               f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için Excel to Media gönderimi tamamlandı.")
                     time.sleep(2) # Medya gönderimi sonrası bekleme süresi
                     attachment_sent_successful = True

                     if current_message_to_to_send:
                          message_sent_successful = caption_written_successful # Alt yazı gönderimi başarılıysa mesaj gönderimi de başarılı sayılır (mantıksal olarak)
                     else:
                          message_sent_successful = True # Alt yazı yoksa ve dosya gönderimi başarılıysa başarılı sayılır

                 except TimeoutException:
                      current_result['Hata Nedeni'] += f"'Excel to Media' gönderme zaman aşımı; "
                 except NoSuchElementException:
                      current_result['Hata Nedeni'] += f"'Excel to Media' gönderme elementi bulunamadı; "
                 except Exception as file_send_e:
                      current_result['Hata Nedeni'] += f"'Excel to Media' gönderme hatası: {file_send_e}; "


            elif islem_turu_original == 'Mesaj':
                if current_message_to_to_send:
                    try:
                        current_message_input_box = wait.until(
                            EC.presence_of_element_located((By.XPATH, message_input_box_xpath)))
                        current_message_input_box.send_keys(current_message_to_to_send)
                        time.sleep(1)

                        send_button = wait.until(EC.presence_of_element_located((By.XPATH, send_message_button_xpath)))
                        send_button.click()
                        root.after(0, update_status,
                                   f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için Mesaj gönderimi tamamlandı.")
                        time.sleep(1.5)  # Mesaj gönderimi sonrası bekleme süresi
                        message_sent_successful = True
                        attachment_sent_successful = True  # Mesaj gönderimi başarılıysa attachment da başarılı sayılır (mantıksal olarak)

                    except TimeoutException:
                        current_result['Hata Nedeni'] += 'Mesaj gönder butonu bulunamadı veya zaman aşımı oluştu (Sadece Mesaj); '
                    except NoSuchElementException:
                        current_result['Hata Nedeni'] += 'Mesaj gönderme elementi bulunamadı (Sadece Mesaj); '
                    except Exception as send_e:
                        current_result['Hata Nedeni'] += f'Mesaj gönderme hatası (Sadece Mesaj): {send_e}; '
                # Mesaj boş durumu artık gönderim öncesi ön kontrolde ele alınıyor.

            # --- İŞLEM TÜRÜNE GÖRE GÖNDERME SONU ---

            # Gönderim başarılı mı kontrol et
            is_successful = False
            if islem_turu_original == 'Mesaj':
                is_successful = message_sent_successful
            elif islem_turu_original in ['Doküman']:
                is_successful = attachment_sent_successful and (not current_message_to_to_send or message_written_successful)
            elif islem_turu_original == 'Medya':
                 is_successful = attachment_sent_successful and (not current_message_to_to_send or caption_written_successful)
            elif islem_turu_original == 'Excel to Media':
                 is_successful = attachment_sent_successful and (not current_message_to_to_send or caption_written_successful)


            # Sonuç sözlüğünü güncelle
            current_result['Durum'] = 'Başarılı' if is_successful else 'Başarısız'
            if not is_successful and not current_result['Hata Nedeni']:
                 current_result['Hata Nedeni'] = "Gönderim elementi bulunamadı veya gönderim hatası" # Daha spesifik hata yoksa genel hata

            # Sonucu listeye ekle
            sending_results.append(current_result)
            update_report_treeview()  # Rapor Treeview'ını güncelle

            last_processed_index = i  # Bu kişiyi başarıyla veya hatayla işledik

            # --- İki kez ESC tuşuna bas ---
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(0.5)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(1)
            except Exception as esc_e:
                root.after(0, update_status,
                           f"Hata: [{i + 1}/{len(contact_list)}] ESC tuşlarına basılırken hata oluştu: {esc_e}")
                # ESC hatasını rapora ekle (eğer zaten bir hata yoksa veya farklı bir hata türüyse)
                if current_result['Durum'] == 'Başarılı': # Eğer gönderim başarılı olduysa bu hatayı ekleyelim
                     current_result['Durum'] = 'Kısmen Başarılı' # ESC hatası kısmen başarılı sayılabilir
                     current_result['Hata Nedeni'] += f'; ESC tuşlarına basılırken hata: {esc_e}'
                     update_report_treeview() # Raporu tekrar güncelle


            # --- Arama butonuna tekrar tıkla ---
            try:
                search_button_after_esc = wait.until(EC.presence_of_element_located((By.XPATH, search_button_xpath)))
                search_button_after_esc.click()
                time.sleep(1)
            except TimeoutException:
                root.after(0, update_status,
                           f"Hata: [{i + 1}/{len(contact_list)}] {unvan} ({search_term}) işlemi sonrası arama butonuna tıklanırken zaman aşımı oluştu.")
                # Arama butonu hatasını rapora ekle (eğer zaten bir hata yoksa veya farklı bir hata türüyse)
                if current_result['Durum'] == 'Başarılı': # Eğer gönderim başarılı olduysa bu hatayı ekleyelim
                     current_result['Durum'] = 'Kısmen Başarılı' # Arama butonu hatası kısmen başarılı sayılabilir
                     current_result['Hata Nedeni'] += '; İşlem sonrası arama ekranına dönme zaman aşımı'
                     update_report_treeview() # Raporu tekrar güncelle
                elif 'İşlem sonrası arama ekranına dönme zaman aşımı' not in current_result['Hata Nedeni']:
                     current_result['Hata Nedeni'] += '; İşlem sonrası arama ekranına dönme zaman aşımı'
                     update_report_treeview() # Raporu tekrar güncelle


            except Exception as e:
                root.after(0, update_status,
                           f"Hata: [{i + 1}/{len(contact_list)}] {unvan} ({search_term}) işlemi sonrası arama butonuna tıklanırken beklenmeyen hata oluştu: {e}")
                # Arama butonu hatasını rapora ekle (eğer zaten bir hata yoksa veya farklı bir hata türüyse)
                if current_result['Durum'] == 'Başarılı': # Eğer gönderim başarılı olduysa bu hatayı ekleyelim
                     current_result['Durum'] = 'Kısmen Başarılı' # Arama butonu hatası kısmen başarılı sayılabilir
                     current_result['Hata Nedeni'] += f'; İşlem sonrası arama ekranına dönme hatası: {e}'
                     update_report_treeview() # Raporu tekrar güncelle
                elif f'İşlem sonrası arama ekranına dönme hatası: {e}' not in current_result['Hata Nedeni']:
                     current_result['Hata Nedeni'] += f'; İşlem sonrası arama ekranına dönme hatası: {e}'
                     update_report_treeview() # Raporu tekrar güncelle


        except Exception as e: # Genel hata yakalama bloğu (TimeoutException dahil)
            # Hata zaten yukarıda spesifik olarak ele alınmadıysa (örn: Kişi/Grup bulunamadı)
            if 'Kişi/Grup bulunamadı hatası' in str(e):
                # Bu hata zaten yukarıda işlendi, tekrar işlemeye gerek yok
                pass
            else:
                root.after(0, update_status,
                           f"[{i + 1}/{len(contact_list)}] {unvan} ({search_term}) için gönderim başladı. Durum: Başarısız (Beklenmeyen hata)")
                if not current_result['Hata Nedeni']: # Eğer daha spesifik bir hata nedeni yoksa
                     current_result['Hata Nedeni'] = f'Beklenmeyen hata: {e}'
                # Hata durumunda da sonucu listeye ekle (eğer daha önce eklenmediyse)
                if current_result not in sending_results:
                    sending_results.append(current_result)
                    update_report_treeview()
                last_processed_index = i # Bu kişiyi işledik (atlanmış olsa da)

                # Hata durumunda bir sonraki kişiye geçmek için arama ekranına dönmeye çalışalım.
                try:
                    search_button_after_send = wait.until(EC.presence_of_element_located((By.XPATH, search_button_xpath)))
                    search_button_after_send.click()
                    time.sleep(1)
                except:
                     root.after(0, update_status, f"Hata: [{i+1}/{len(contact_list)}] {unvan} ({search_term}) hatası sonrası arama ekranına dönülürken hata oluştu.")

                # Hata durumunda da ESC tuşlarına basalım
                try:
                    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                    time.sleep(0.5)
                    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                    time.sleep(1)
                except Exception as esc_e:
                     root.after(0, update_status, f"Hata: [{i+1}/{len(contact_list)}] Hata sonrası ESC tuşlarına basılırken hata oluştu: {esc_e}")

                # Hata durumunda da arama butonuna tekrar tıklamaya çalışalım
                try:
                    search_button_after_esc = wait.until(EC.presence_of_element_located((By.XPATH, search_button_xpath)))
                    search_button_after_esc.click()
                    time.sleep(1)
                except:
                     root.after(0, update_status, f"Hata: [{i+1}/{len(contact_list)}] {unvan} ({search_term}) hatası sonrası arama butonuna tekrar tıklanırken hata oluştu.")


        finally:
            # Excel to Media işlemi sonrası geçici dosyayı temizle
            if islem_turu_original == "Excel to Media" and temp_jpeg_path and os.path.exists(temp_jpeg_path):
                try:
                    os.remove(temp_jpeg_path)
                    root.after(0, update_status,
                               f"[{i + 1}/{len(contact_list)}] Geçici Excel resmi dosyası silindi: {temp_jpeg_path}")
                except Exception as cleanup_e:
                    root.after(0, update_status,
                               f"Uyarı: [{i + 1}/{len(contact_list)}] Geçici dosya silinirken hata: {cleanup_e}")

    # Döngü bittiğinde (tamamlandı veya iptal edildi/duraklatıldı)
    if cancel_sending_flag:
        root.after(0, update_status, "Mesaj gönderme işlemi kullanıcı tarafından iptal edildi. Rapor oluşturuluyor...")
        # İptal sonrası buton durumları
        root.after(0, lambda: btn_start_sending.config(
            state=tk.NORMAL if driver and contact_list else tk.DISABLED))  # Eğer driver bağlıysa ve liste boş değilse gönder aktif
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
        root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
        root.after(0, lambda: btn_cancel.config(state=tk.DISABLED))
        root.after(0, generate_report)  # Raporu otomatik oluştur
        # Rapor ve Son Raporu Aç butonları generate_report içinde aktif ediliyor.

    elif stop_sending_flag:
        root.after(0, update_status, f"Mesaj gönderme işlemi duraklatıldı. {last_processed_index + 1}. kişide durdu.")
        # Duraklatma sonrası buton durumları
        root.after(0, lambda: btn_start_sending.config(state=tk.DISABLED))  # Gönder butonu pasif
        root.after(0, lambda: btn_continue_sending.config(state=tk.NORMAL if driver and last_processed_index < len(
            contact_list) - 1 else tk.DISABLED))  # Eğer driver bağlıysa ve liste sonu değilse devam et aktif
        root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
        root.after(0, lambda: btn_cancel.config(state=tk.NORMAL))  # İptal Et hala aktif olabilir
        root.after(0, lambda: btn_generate_report.config(
            state=tk.NORMAL if sending_results else tk.DISABLED))  # Rapor oluştur aktif
        root.after(0, lambda: btn_open_last_report.config(state=tk.NORMAL if last_report_path and os.path.exists(
            last_report_path) else tk.DISABLED))  # Son Raporu Aç aktif (mevcutsa)


    else:  # Gönderim tamamlandı
        root.after(0, update_status, "Mesaj gönderme işlemi tamamlandı. Rapor oluşturuluyor...")
        last_processed_index = -1  # İşlem tamamlandı, indeksi sıfırla
        # Tamamlanma sonrası buton durumları
        root.after(0, lambda: btn_start_sending.config(
            state=tk.NORMAL if driver and contact_list else tk.DISABLED))  # Eğer driver bağlıysa ve liste boş değilse gönder aktif
        root.after(0, lambda: btn_continue_sending.config(state=tk.DISABLED))  # Devam Et butonu pasif
        root.after(0, lambda: btn_stop.config(state=tk.DISABLED))
        root.after(0, lambda: btn_cancel.config(state=tk.DISABLED))
        root.after(0, generate_report)  # Raporu otomatik oluştur
        # Rapor ve Son Raporu Aç butonları generate_report içinde aktif ediliyor.


# Uygulama kapatıldığında WebDriver'ı kapatma ve rapor oluşturma
def on_closing():
    """Uygulama kapatılırken gerekli temizlik işlemlerini yapar."""
    global driver, sending_thread, cancel_sending_flag, stop_sending_flag, stop_check_driver_thread, check_driver_thread
    if messagebox.askokcancel("Çıkış", "Uygulamadan çıkmak istiyor musunuz? WebDriver kapatılacaktır."):
        # WebDriver kontrol thread'ini durdur
        stop_check_driver_thread = True
        if check_driver_thread and check_driver_thread.is_alive():
            # Kontrol thread'inin bitmesi için kısa bir süre bekle
            check_driver_thread.join(timeout=1)

        # Eğer gönderme thread'i çalışıyorsa ve henüz iptal edilmediyse, iptal et ve rapor oluştur
        if sending_thread and sending_thread.is_alive() and not cancel_sending_flag:
            cancel_sending_flag = True
            stop_sending_flag = True
            # Thread'in bitmesi için kısa bir süre bekle (opsiyonel, uygulamanın donmasına neden olabilir)
            # sending_thread.join(timeout=5)
            root.after(0, update_status,
                       "Uygulama kapatılıyor, mevcut gönderim iptal ediliyor ve rapor oluşturuluyor...")
            # Thread'in bitmesini beklemek yerine, kapatma işlemine devam ediyoruz.
            # Rapor, thread kendi döngüsünü bitirdiğinde oluşturulacaktır.

        elif sending_thread and sending_thread.is_alive() and cancel_sending_flag:
            # Eğer zaten iptal edildiyse (İptal Et butonuna basıldıysa), sadece kapatmaya devam et
            root.after(0, update_status, "Uygulama kapatılıyor.")

        elif sending_results and not (sending_thread and sending_thread.is_alive()):
            # Eğer gönderim tamamlandıysa veya duraklatıldıysa ve rapor henüz oluşturulmadıysa
            # (normalde otomatik oluşturulmalı ama bir aksilik olursa diye)
            # Bu durum normalde oluşmamalı, ama ek güvenlik için kontrol edilebilir.
            # İstenen senaryo "durdura basılıp uygulama kapatılırsa raporu yine hazırlıcak" olduğu için,
            # eğer stop_sending_flag True ise ve thread çalışmıyorsa (duraklatılmış ve sonra kapatılmış), rapor oluşturalım.
            if stop_sending_flag and not (sending_thread and sending_thread.is_alive()):
                root.after(0, update_status, "Uygulama kapatılıyor, duraklatılmış gönderim için rapor oluşturuluyor...")
                root.after(0, generate_report)  # Duraklatılmış durumda kapatılırsa rapor oluştur

        # Ayarları kaydet
        save_settings()

        if driver:
            try:
                driver.quit()
                driver = None
                root.after(0, update_status, "WebDriver kapatıldı.")
            except Exception as e:
                root.after(0, update_status, f"WebDriver kapatılırken hata oluştu: {e}")

        root.destroy()


# Arayüzü oluştur
root = tk.Tk()
root.title("WhatsApp Kişi Yönetimi ve Mesaj/Dosya Gönderme")
root.geometry("1200x800") # Başlangıç boyutu
root.minsize(900, 600) # Minimum boyut

# Uygulama kapatılırken on_closing fonksiyonunu çağır
root.protocol("WM_DELETE_WINDOW", on_closing)

# Buton stillerini tanımla
style = ttk.Style()

# Genel stil ayarları
style.theme_use('clam') # Modern bir tema seçimi (alternatifler: 'default', 'alt', 'clam', 'xpnative', 'winnative')

# Font ayarı
default_font = ("Segoe UI", 10)
bold_font = ("Segoe UI", 10, "bold")
style.configure('.', font=default_font)
style.configure('TButton', font=bold_font, padding=6, borderwidth=2, relief="raised")
style.configure('TLabel', font=default_font)
style.configure('TEntry', font=default_font, padding=3)
style.configure('TCombobox', font=default_font, padding=3)
style.configure('TText', font=default_font) # Text widget için stil
style.configure('TLabelframe.Label', font=bold_font) # Labelframe başlıkları

# Renkli buton stilleri
style.configure('Blue.TButton', background='#3498db', foreground='white')
style.map('Blue.TButton', background=[('active', '#2980b9')])

style.configure('Green.TButton', background='#2ecc71', foreground='white')
style.map('Green.TButton', background=[('active', '#27ae60')])

style.configure('Orange.TButton', background='#f39c12', foreground='white')
style.map('Orange.TButton', background=[('active', '#e67e22')])

style.configure('Red.TButton', background='#e74c3c', foreground='white')
style.map('Red.TButton', background=[('active', '#c0392b')])

style.configure('Gray.TButton', background='#bdc3c7', foreground='black')
style.map('Gray.TButton', background=[('active', '#95a5a6')])

# Notebook (Sekmeler) Oluştur
notebook = ttk.Notebook(root)
notebook.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")

# Kişi Yönetimi Sekmesi
contact_management_frame = ttk.Frame(notebook, padding="10 10 10 10") # Daha fazla padding
notebook.add(contact_management_frame, text='Kişi Yönetimi')

# Ayarlar Sekmesi
settings_frame = ttk.Frame(notebook, padding="10 10 10 10")
notebook.add(settings_frame, text='Ayarlar')

# Raporlar Sekmesi
reports_frame = ttk.Frame(notebook, padding="10 10 10 10")
notebook.add(reports_frame, text='Raporlar')

# --- Kişi Yönetimi Sekmesi İçeriği ---

# Üst Butonlar Frame (WebDriver Başlat ve Bağlan, Tarayıcı Seçimi, Gönder, Devam Et, Durdur, İptal Et)
top_buttons_frame = ttk.Frame(contact_management_frame)
top_buttons_frame.grid(row=0, column=0, columnspan=2, pady=10, sticky="ew") # Daha fazla pady

# Tarayıcı Seçimi ve WebDriver Butonu Grubu
driver_control_frame = ttk.Frame(top_buttons_frame)
driver_control_frame.pack(side=tk.LEFT, padx=5)

ttk.Label(driver_control_frame, text="Tarayıcı Seç:").pack(side=tk.LEFT, padx=5)
browser_options = ["Chrome", "Firefox", "Edge"]
combo_browser = ttk.Combobox(driver_control_frame, values=browser_options, state="readonly", width=10)
combo_browser.set("Chrome")
combo_browser.pack(side=tk.LEFT, padx=5)

btn_start_and_connect = ttk.Button(driver_control_frame, text="WebDriver Başlat ve Bağlan",
                                   command=start_driver_and_connect_whatsapp, style='Blue.TButton')
btn_start_and_connect.pack(side=tk.LEFT, padx=5)

# Gönderim Kontrol Butonları Grubu
sending_control_frame = ttk.Frame(top_buttons_frame)
sending_control_frame.pack(side=tk.RIGHT, padx=5)

btn_start_sending = ttk.Button(sending_control_frame, text="Gönder", command=start_sending, state=tk.DISABLED, style='Green.TButton')
btn_start_sending.pack(side=tk.LEFT, padx=5)

btn_continue_sending = ttk.Button(sending_control_frame, text="Devam Et", command=continue_sending, state=tk.DISABLED, style='Green.TButton')
btn_continue_sending.pack(side=tk.LEFT, padx=5)

btn_stop = ttk.Button(sending_control_frame, text="Durdur", command=stop_sending, state=tk.DISABLED, style='Orange.TButton')
btn_stop.pack(side=tk.LEFT, padx=5)

btn_cancel = ttk.Button(sending_control_frame, text="İptal Et", command=cancel_sending, state=tk.DISABLED, style='Red.TButton')
btn_cancel.pack(side=tk.LEFT, padx=5)


# Kişi Listesi Frame (Treeview)
list_frame = ttk.LabelFrame(contact_management_frame, text="Kişi Listesi")
list_frame.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

# Treeview (Liste görünümü)
columns = ('No', 'Ünvan', 'Alan', 'Numara', 'İşlem Türü', 'Mesaj', 'Hedef Hücre', 'Dosya Yolu')
tree_contacts = ttk.Treeview(list_frame, columns=columns, show='headings')

# Scrollbar'lar
vsb_contacts = ttk.Scrollbar(list_frame, orient="vertical", command=tree_contacts.yview)
hsb_contacts = ttk.Scrollbar(list_frame, orient="horizontal", command=tree_contacts.xview)
tree_contacts.configure(yscrollcommand=vsb_contacts.set, xscrollcommand=hsb_contacts.set)

vsb_contacts.pack(side="right", fill="y")
hsb_contacts.pack(side="bottom", fill="x")
tree_contacts.pack(expand=True, fill='both')


for col in columns:
    tree_contacts.heading(col, text=col)
    if col == 'No':
        tree_contacts.column(col, width=40, minwidth=40, anchor='center')
    elif col == 'Ünvan':
        tree_contacts.column(col, width=120, minwidth=80)
    elif col == 'Alan':
        tree_contacts.column(col, width=60, minwidth=50, anchor='center')
    elif col == 'Numara':
        tree_contacts.column(col, width=120, minwidth=80)
    elif col == 'İşlem Türü':
        tree_contacts.column(col, width=100, minwidth=80)
    elif col == 'Mesaj':
        tree_contacts.column(col, width=200, minwidth=100)
    elif col == 'Hedef Hücre':
        tree_contacts.column(col, width=90, minwidth=70, anchor='center')
    elif col == 'Dosya Yolu':
        tree_contacts.column(col, width=250, minwidth=150)
    else:
        tree_contacts.column(col, width=100, minwidth=80)

tree_contacts.bind('<<TreeviewSelect>>', select_contact_for_edit)

# Giriş Alanları Frame
input_frame = ttk.LabelFrame(contact_management_frame, text="Kişi Bilgileri")
input_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# Giriş alanlarının genişlemesini sağla
input_frame.columnconfigure(1, weight=1)
input_frame.columnconfigure(2, weight=0) # Göz At butonu için sabit genişlik

ttk.Label(input_frame, text="Ünvan:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
entry_unvan = ttk.Entry(input_frame)
entry_unvan.grid(row=0, column=1, padx=5, pady=2, sticky="ew", columnspan=2) # 2 sütunu kapla

ttk.Label(input_frame, text="Alan Kodu:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
entry_alan = ttk.Entry(input_frame)
entry_alan.grid(row=1, column=1, padx=5, pady=2, sticky="ew", columnspan=2)

ttk.Label(input_frame, text="Numara / Grup Adı:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
entry_numara = ttk.Entry(input_frame)
entry_numara.grid(row=2, column=1, padx=5, pady=2, sticky="ew", columnspan=2)

ttk.Label(input_frame, text="İşlem Türü:").grid(row=3, column=0, padx=5, pady=2, sticky="w")
islem_turu_options = ["Mesaj", "Doküman", "Medya", "Excel to Media"]
combo_islem_turu = ttk.Combobox(input_frame, values=islem_turu_options, state="readonly")
combo_islem_turu.grid(row=3, column=1, padx=5, pady=2, sticky="ew", columnspan=2)
combo_islem_turu.bind("<<ComboboxSelected>>", update_input_field_visibility) # Seçim değiştiğinde fonksiyonu çağır

# Dinamik olarak gösterilecek/gizlenecek alanlar için etiket ve widget referansları
lbl_mesaj = ttk.Label(input_frame, text="Mesaj / Alt Yazı:")
entry_mesaj = tk.Text(input_frame, height=5) # Yükseklik artırıldı

lbl_hedef_hucre = ttk.Label(input_frame, text="Hedef Hücre (Excel):")
entry_hedef_hucre = ttk.Entry(input_frame)

lbl_dosya_yolu = ttk.Label(input_frame, text="Dosya Yolu:")
entry_dosya_yolu = ttk.Entry(input_frame)
btn_browse_file = ttk.Button(input_frame, text="Göz At", command=browse_file, style='Gray.TButton')


# Kişi Yönetim Butonları Frame (Ekle/Kaydet, Sil, İçe/Dışa Aktar, Kontrol Et, Listeyi Temizle)
contact_buttons_frame = ttk.Frame(contact_management_frame)
contact_buttons_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky="ew")

# Butonların eşit dağılması için columnconfigure
for i in range(7): # Toplam 7 buton var (Ekle/Kaydet, İptal, Sil, İçe Aktar, Dışa Aktar, Kontrol Et, Temizle)
    contact_buttons_frame.columnconfigure(i, weight=1)

btn_save_contact = ttk.Button(contact_buttons_frame, text="Kişi Ekle", command=save_contact, style='Green.TButton')
btn_save_contact.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

btn_cancel_edit = ttk.Button(contact_buttons_frame, text="İptal", command=cancel_edit, style='Red.TButton')
btn_cancel_edit.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
btn_cancel_edit.grid_forget()  # Başlangıçta gizle

btn_delete = ttk.Button(contact_buttons_frame, text="Seçileni Sil", command=delete_contact, style='Red.TButton')
btn_delete.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

btn_import = ttk.Button(contact_buttons_frame, text="Listeyi İçe Aktar", command=import_list, style='Gray.TButton')
btn_import.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

btn_export = ttk.Button(contact_buttons_frame, text="Listeyi Dışa Aktar", command=export_list, style='Gray.TButton')
btn_export.grid(row=0, column=4, padx=5, pady=5, sticky="ew")

btn_check_contacts = ttk.Button(contact_buttons_frame, text="Kişileri Kontrol Et", command=check_contacts, style='Gray.TButton')
btn_check_contacts.grid(row=0, column=5, padx=5, pady=5, sticky="ew")

btn_clear_list = ttk.Button(contact_buttons_frame, text="Listeyi Temizle", command=clear_contact_list, style='Red.TButton')
btn_clear_list.grid(row=0, column=6, padx=5, pady=5, sticky="ew")


# Durum Raporu Frame (Kişi Yönetimi sekmesinde kalıyor)
status_frame = ttk.LabelFrame(contact_management_frame, text="Durum Raporu")
status_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

status_report = tk.Text(status_frame, height=10, state='normal')
status_report_vsb = ttk.Scrollbar(status_frame, orient="vertical", command=status_report.yview)
status_report.configure(yscrollcommand=status_report_vsb.set)
status_report_vsb.pack(side="right", fill="y")
status_report.pack(expand=True, fill='both')


# Kişi Yönetimi Sekmesi Grid Ayarları
contact_management_frame.columnconfigure(0, weight=1) # Sol panel (girişler) genişlesin
contact_management_frame.columnconfigure(1, weight=2) # Sağ panel (liste) daha çok genişlesin
contact_management_frame.rowconfigure(1, weight=2) # Giriş ve Liste alanları genişlesin
contact_management_frame.rowconfigure(3, weight=1) # Durum raporu da genişlesin


# --- Ayarlar Sekmesi İçeriği ---

# XPath Ayarları Frame (Ana Frame)
xpath_settings_main_frame = ttk.Frame(settings_frame)
xpath_settings_main_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
xpath_settings_main_frame.columnconfigure(0, weight=1) # İçindeki labelframe'lerin genişlemesini sağla

# Gruplandırılmış XPath Ayarları Frame'leri
xpath_group_frames = {}
current_row_in_settings = 0

# Global xpath_entries sözlüğünü burada tanımla
xpath_entries = {}

for group_name, xpath_keys in XPATH_GROUPS.items():
    group_frame = ttk.LabelFrame(xpath_settings_main_frame, text=group_name)
    group_frame.grid(row=current_row_in_settings, column=0, padx=5, pady=5, sticky="nsew")
    xpath_group_frames[group_name] = group_frame

    group_frame.columnconfigure(1, weight=1)  # XPath giriş alanlarının genişlemesini sağla

    row_index_in_group = 0
    for key in xpath_keys:
        # Türkçe açıklamayı al, yoksa anahtarın kendisini kullan
        label_text = XPATH_TURKISH_LABELS.get(key, key)
        ttk.Label(group_frame, text=f"{label_text}:").grid(row=row_index_in_group, column=0, padx=5, pady=2, sticky="w")
        entry = ttk.Entry(group_frame) # Width kaldırıldı, sticky ile genişleyecek
        entry.grid(row=row_index_in_group, column=1, padx=5, pady=2, sticky="ew")
        # Global xpath_entries sözlüğüne referansı sakla
        xpath_entries[key] = entry
        row_index_in_group += 1

    current_row_in_settings += 1

# Rapor Klasörü Seçimi Frame (Ayarlar sekmesinde)
report_folder_frame_settings = ttk.LabelFrame(settings_frame, text="Rapor Klasörü")
report_folder_frame_settings.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

report_folder_frame_settings.columnconfigure(0, weight=1) # Giriş alanının genişlemesini sağla

entry_report_directory = ttk.Entry(report_folder_frame_settings) # Width kaldırıldı
entry_report_directory.grid(row=0, column=0, padx=5, pady=2, sticky="ew")

btn_browse_report_directory = ttk.Button(report_folder_frame_settings, text="Seç", command=browse_report_directory, style='Gray.TButton')
btn_browse_report_directory.grid(row=0, column=1, padx=5, pady=2)


# Ayarlar Kaydet Butonları Frame
settings_buttons_frame = ttk.Frame(settings_frame)
settings_buttons_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

settings_buttons_frame.columnconfigure(0, weight=1)
settings_buttons_frame.columnconfigure(1, weight=1)

btn_save_settings = ttk.Button(settings_buttons_frame, text="Ayarları Kaydet", command=save_settings, style='Green.TButton')
btn_save_settings.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

btn_reset_xpath = ttk.Button(settings_buttons_frame, text="XPath Varsayılana Sıfırla", command=reset_xpath_to_default, style='Orange.TButton')
btn_reset_xpath.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

# Ayarlar Sekmesi Grid Ayarları
settings_frame.columnconfigure(0, weight=1)
settings_frame.rowconfigure(0, weight=1)  # XPath ayarları ana frame'inin genişlemesini sağla


# --- Raporlar Sekmesi İçeriği ---

# Rapor Sonuçları Frame (Treeview)
report_results_frame = ttk.LabelFrame(reports_frame, text="Gönderim Raporu Sonuçları")
report_results_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Rapor Treeview'ı
report_columns = ('No', 'Ünvan', 'Numara', 'İşlem Türü', 'Durum', 'Hata Nedeni')
tree_report = ttk.Treeview(report_results_frame, columns=report_columns, show='headings')

# Scrollbar'lar
vsb_report = ttk.Scrollbar(report_results_frame, orient="vertical", command=tree_report.yview)
hsb_report = ttk.Scrollbar(report_results_frame, orient="horizontal", command=tree_report.xview)
tree_report.configure(yscrollcommand=vsb_report.set, xscrollcommand=hsb_report.set)

vsb_report.pack(side="right", fill="y")
hsb_report.pack(side="bottom", fill="x")
tree_report.pack(expand=True, fill='both')

for col in report_columns:
    tree_report.heading(col, text=col)
    if col == 'No':
        tree_report.column(col, width=40, minwidth=40, anchor='center')
    elif col == 'Ünvan':
        tree_report.column(col, width=120, minwidth=80)
    elif col == 'Numara':
        tree_report.column(col, width=120, minwidth=80)
    elif col == 'İşlem Türü' or col == 'Durum':
        tree_report.column(col, width=100, minwidth=80)
    elif col == 'Hata Nedeni':
        tree_report.column(col, width=300, minwidth=150)
    else:
        tree_report.column(col, width=150, minwidth=100)


# Rapor Butonları Frame
report_buttons_frame = ttk.Frame(reports_frame)
report_buttons_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

report_buttons_frame.columnconfigure(0, weight=1)
report_buttons_frame.columnconfigure(1, weight=1)
report_buttons_frame.columnconfigure(2, weight=1)

btn_generate_report = ttk.Button(report_buttons_frame, text="Raporu Excel Oluştur", command=generate_report,
                                 state=tk.DISABLED, style='Green.TButton')  # Başlangıçta pasif
btn_generate_report.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

btn_open_last_report = ttk.Button(report_buttons_frame, text="Son Raporu Aç", command=open_last_report,
                                  state=tk.DISABLED, style='Gray.TButton')  # Başlangıçta pasif (load_settings güncelleyecek)
btn_open_last_report.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

btn_clear_report = ttk.Button(report_buttons_frame, text="Raporu Temizle", command=clear_report_results, style='Red.TButton')
btn_clear_report.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

# Raporlar Sekmesi Grid Ayarları
reports_frame.columnconfigure(0, weight=1)
reports_frame.rowconfigure(0, weight=1)  # Rapor sonuçları frame'inin genişlemesini sağla

# Ana Pencere Grid Ayarları
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Durum Çubuğu (En altta)
status_bar_frame = ttk.Frame(root, relief=tk.SUNKEN, borderwidth=1)
status_bar_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
status_bar_frame.columnconfigure(0, weight=1)

status_bar_label = ttk.Label(status_bar_frame, text="Uygulama hazır.", anchor="w", font=("Segoe UI", 9))
status_bar_label.grid(row=0, column=0, sticky="ew", padx=5, pady=2)


# Uygulama başlangıcında ayarları yükle
load_settings()
# Uygulama başlangıcında giriş alanlarının görünürlüğünü ayarla
update_input_field_visibility()

# Uygulamayı çalıştır
root.mainloop()

pip install pandas selenium webdriver-manager numpy pywin32