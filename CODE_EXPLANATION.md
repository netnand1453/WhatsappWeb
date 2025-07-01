# Kod Açıklaması

Bu belge, `wpoto2.py` dosyasındaki Python kodunun ana bileşenlerini ve işlevlerini açıklamaktadır.

## Genel Yapı

Kod, `tkinter` kütüphanesi kullanılarak bir grafik kullanıcı arayüzü (GUI) sağlar. Bu arayüz, kullanıcıların WhatsApp Web otomasyonunu kolayca yönetmesine olanak tanır. `selenium` kütüphanesi, web tarayıcısıyla etkileşim için kullanılırken, `pandas` Excel dosyalarını okumak için kullanılır. İşlemlerin kullanıcı arayüzünü kilitlemesini önlemek için `threading` yoğun olarak kullanılır.

## Ana Bölümler

### 1. Kütüphane İçe Aktarımları ve Global Değişkenler
Kodun başında gerekli tüm kütüphaneler içe aktarılır. `tkinter` (GUI), `pandas` (Excel), `selenium` (Web otomasyonu) ve `webdriver_manager` (Tarayıcı sürücüsü yönetimi) başlıca kütüphanelerdir. Ayrıca, uygulamanın farklı fonksiyonları arasında durum bilgisini paylaşmak için çeşitli global değişkenler (örn. `contact_list`, `driver`, `stop_sending_flag`, `xpath_settings`) tanımlanır. Özellikle `DEFAULT_XPATH_SETTINGS`, WhatsApp Web arayüz elementlerinin konumlarını belirlemek için kullanılan temel XPath'leri içerir.

### 2. Ayar Yönetimi
* **`load_settings()`**: `settings.json` dosyasından kaydedilmiş ayarları (rapor klasörü, özel XPath'ler) yükler. Eğer ayarlar dosyası yoksa veya bozuksa, varsayılan değerleri kullanır.
* **`save_settings()`**: Mevcut uygulama ayarlarını (`settings.json` dosyasına) kaydeder, böylece kullanıcı uygulamayı bir dahaki sefere açtığında ayarları tekrar girmesine gerek kalmaz.
* **`reset_xpath_to_default()`**: XPath ayarlarını uygulamanın varsayılan değerlerine sıfırlayan bir fonksiyondur.

### 3. WebDriver Başlatma ve WhatsApp Bağlantısı
* **`start_driver_and_connect_whatsapp()`**: Seçilen tarayıcıya (Chrome, Firefox, Edge) göre uygun WebDriver'ı başlatır. Bu işlem ayrı bir iş parçacığında (`_start_driver_and_connect_whatsapp_thread`) çalışır, böylece arayüz donmaz. WebDriver Manager kütüphaneleri (`ChromeDriverManager`, `GeckoDriverManager`, `EdgeChromiumDriverManager`) sayesinde tarayıcı sürücüleri otomatik olarak indirilir.
* **`_start_driver_and_connect_whatsapp_thread()`**: Asıl WebDriver başlatma ve WhatsApp Web'e yönlendirme işlemini gerçekleştirir. Kullanıcı profili için özel bir dizin (`user-data-dir`) kullanarak oturum bilgilerini kalıcı hale getirir, böylece her seferde QR kodu okutmaya gerek kalmaz.
* **`_check_driver_thread()`**: WebDriver'ın (tarayıcı penceresinin) hala açık ve aktif olup olmadığını periyodik olarak kontrol eden arka plan iş parçacığıdır. Eğer bağlantı kesilirse, ilgili değişkenleri günceller ve devam eden gönderim işlemlerini iptal eder.

### 4. Kişi Listesi ve Veri Yönetimi
* **`update_contact_treeview()`**: Kişi listesi `Treeview` widget'ını (`contact_list` global değişkeninden) güncelleyerek arayüzde gösterir.
* **`update_report_treeview()`**: Gönderim sonuçları `Treeview` widget'ını (`sending_results` global değişkeninden) günceller.
* **`clear_input_fields()`**: Kişi ekleme/düzenleme giriş alanlarını temizler.
* **`update_input_field_visibility()`**: Seçilen "İşlem Türü"ne (Mesaj, Doküman, Medya, Excel to Media) göre ilgili giriş alanlarının (mesaj, dosya yolu, hedef hücre) görünürlüğünü dinamik olarak ayarlar.
* **`save_contact()`**: Giriş alanlarındaki bilgiyi alarak `contact_list`'e yeni bir kişi ekler veya mevcut bir kişiyi günceller. Gerekli alanların doldurulup doldurulmadığını kontrol eder.
* **`delete_contact()`**: Seçili kişileri `contact_list`'ten ve arayüzden siler.
* **`edit_contact()`**: Seçili bir kişinin bilgilerini giriş alanlarına doldurarak düzenlemeye olanak tanır.
* **`cancel_edit_contact()`**: Kişi düzenleme modunu iptal eder ve "Kişi Ekle" butonunu tekrar etkinleştirir.
* **`load_contacts_from_excel()`**: Bir Excel dosyasından (varsayılan: `rehber.xlsx`) kişi bilgilerini okuyarak `contact_list`'e ekler. Gerekli sütunların (Ünvan, Numara, İşlem Türü vb.) varlığını kontrol eder.
* **`clear_contact_list()`**: Tüm kişi listesini temizler.

### 5. Gönderim Fonksiyonları
* **`start_sending_messages()`**: Gönderim işlemini başlatan ana fonksiyondur. Ayrı bir iş parçacığında (`_send_messages_thread`) çalışır ve gönderim durumunu yönetir.
* **`_send_messages_thread()`**: Asıl mesaj gönderme mantığını içerir. `contact_list`'teki her kişi için WhatsApp Web üzerinde arama yapar, mesaj kutusunu bulur ve mesaj/doküman/medya gönderme işlemlerini Selenium kullanarak gerçekleştirir.
    * **Numara/Grup Arama:** WhatsApp'taki arama çubuğuna numarayı veya grup adını yazarak sohbeti açmaya çalışır.
    * **Mesaj/Doküman/Medya Gönderimi:** XPath'ler aracılığıyla ilgili elementleri (mesaj kutusu, ataç butonu, dosya yükleme input'u, gönder butonu) bulup etkileşime girer.
    * **"Excel to Media" Özel İşlevi:** `win32com.client` (pywin32) kütüphanesini kullanarak Excel dosyasından belirli bir hücrenin ekran görüntüsünü alır, geçici bir PNG dosyası olarak kaydeder ve bunu WhatsApp'a medya olarak gönderir. **Bu özellik sadece Windows'ta çalışır.**
* **`stop_sending()`**: Gönderim işlemini geçici olarak duraklatır.
* **`cancel_sending()`**: Gönderim işlemini tamamen iptal eder.
* **`continue_sending()`**: Duraklatılmış gönderim işlemini, kaldığı yerden devam ettirir.

### 6. Raporlama
* **`generate_report()`**: `sending_results` listesindeki verileri kullanarak detaylı bir Excel raporu oluşturur ve belirlenen rapor klasörüne kaydeder.
* **`open_last_report()`**: En son oluşturulan rapor dosyasını açar.
* **`clear_report_results()`**: Rapor sonuçlarını (`sending_results`) ve rapor `Treeview`'ını temizler.

### 7. Yardımcı Fonksiyonlar ve Arayüz Güncellemeleri
* **`update_status()`**: Arayüzdeki durum çubuğunu günceller.
* **`update_progress_bar()`**: Gönderim ilerleme çubuğunu günceller.
* **`on_closing()`**: Uygulama penceresi kapatıldığında çalışan, WebDriver'ı güvenli bir şekilde kapatma ve çalışan iş parçacıklarını durdurma işlevini sağlar.
* **`handle_error()`**: Hata mesajlarını kullanıcıya anlaşılır bir şekilde gösterir ve durum çubuğunu günceller.
* **`highlight_element()`**: Geçici olarak bir web elementini görsel olarak vurgulayarak hata ayıklama sırasında yardımcı olur.

## Kullanılan XPath'ler

Kod, WhatsApp Web arayüzündeki elementleri bulmak için XPath'leri kullanır. Bu XPath'ler `DEFAULT_XPATH_SETTINGS` sözlüğünde tanımlanır ve `settings.json` üzerinden özelleştirilebilir. Eğer WhatsApp Web'in arayüzü değişirse, bu XPath'lerin güncellenmesi gerekebilir.

## Stil ve Tema

`ttk.Style` kullanılarak arayüz için özel düğme stilleri (Yeşil, Kırmızı, Gri) tanımlanmıştır. Bu, arayüze daha modern ve işlevsel bir görünüm kazandırır.

---

# Code Explanation

This document explains the main components and functionalities of the Python code in `wpoto2.py`.

## General Structure

The code provides a graphical user interface (GUI) using the `tkinter` library. This interface allows users to easily manage WhatsApp Web automation. The `selenium` library is used for web browser interaction, while `pandas` is used for reading Excel files. `threading` is extensively used to prevent operations from freezing the user interface.

## Main Sections

### 1. Library Imports and Global Variables
At the beginning of the code, all necessary libraries are imported. `tkinter` (GUI), `pandas` (Excel), `selenium` (Web automation), and `webdriver_manager` (Browser driver management) are the primary libraries. Additionally, various global variables (e.g., `contact_list`, `driver`, `stop_sending_flag`, `xpath_settings`) are defined to share state information among different functions of the application. Notably, `DEFAULT_XPATH_SETTINGS` contains the fundamental XPaths used to locate WhatsApp Web interface elements.

### 2. Settings Management
* **`load_settings()`**: Loads saved settings (report directory, custom XPaths) from the `settings.json` file. If the settings file is missing or corrupted, it uses default values.
* **`save_settings()`**: Saves the current application settings (to the `settings.json` file) so that the user does not have to re-enter them the next time they open the application.
* **`reset_xpath_to_default()`**: A function that resets XPath settings to the application's default values.

### 3. WebDriver Başlatma ve WhatsApp Bağlantısı
* **`start_driver_and_connect_whatsapp()`**: Seçilen tarayıcıya (Chrome, Firefox, Edge) göre uygun WebDriver'ı başlatır. Bu işlem ayrı bir iş parçacığında (`_start_driver_and_connect_whatsapp_thread`) çalışır, böylece arayüz donmaz. WebDriver Manager kütüphaneleri (`ChromeDriverManager`, `GeckoDriverManager`, `EdgeChromiumDriverManager`) sayesinde tarayıcı sürücüleri otomatik olarak indirilir.
* **`_start_driver_and_connect_whatsapp_thread()`**: Asıl WebDriver başlatma ve WhatsApp Web'e yönlendirme işlemini gerçekleştirir. Kullanıcı profili için özel bir dizin (`user-data-dir`) kullanarak oturum bilgilerini kalıcı hale getirir, böylece her seferde QR kodu okutmaya gerek kalmaz.
* **`_check_driver_thread()`**: This is a background thread that periodically checks if the WebDriver (browser window) is still open and active. If the connection is lost, it updates the relevant variables and cancels any ongoing sending operations.

### 4. Contact List and Data Management
* **`update_contact_treeview()`**: Updates and displays the contact list `Treeview` widget from the `contact_list` global variable.
* **`update_report_treeview()`**: Updates the sending results `Treeview` widget from the `sending_results` global variable.
* **`clear_input_fields()`**: Clears the input fields for adding/editing contacts.
* **`update_input_field_visibility()`**: Dynamically adjusts the visibility of relevant input fields (message, file path, target cell) based on the selected "Action Type" (Message, Document, Media, Excel to Media).
* **`save_contact()`**: Takes information from input fields and adds a new contact or updates an existing one in `contact_list`. It checks if required fields are filled.
* **`delete_contact()`**: Deletes selected contacts from `contact_list` and the interface.
* **`edit_contact()`**: Allows editing a selected contact by populating its information into the input fields.
* **`cancel_edit_contact()`**: Cancels contact editing mode and re-enables the "Add Contact" button.
* **`load_contacts_from_excel()`**: Reads contact information from an Excel file (default: `rehber.xlsx`) and adds it to `contact_list`. It checks for the presence of required columns (Title, Area, Number, Action Type, etc.).
* **`clear_contact_list()`**: Clears the entire contact list.

### 5. Sending Functions
* **`start_sending_messages()`**: The main function that initiates the sending process. It runs in a separate thread (`_send_messages_thread`) and manages the sending status.
* **`_send_messages_thread()`**: Contains the core message sending logic. For each contact in `contact_list`, it searches on WhatsApp Web, finds the message box, and performs message/document/media sending operations using Selenium.
    * **Number/Group Search:** Attempts to open the chat by typing the number or group name into the WhatsApp search bar.
    * **Message/Document/Media Sending:** Interacts with relevant elements (message box, attach button, file upload input, send button) by locating them via XPaths.
    * **"Excel to Media" Special Function:** Uses the `win32com.client` (pywin32) library to take a screenshot of a specific cell from an Excel file, saves it as a temporary PNG file, and sends it as media to WhatsApp. **This feature only works on Windows.**
* **`stop_sending()`**: Temporarily pauses the sending process.
* **`cancel_sending()`**: Completely cancels the sending process.
* **`continue_sending()`**: Resumes a paused sending process from where it left off.

### 6. Reporting
* **`generate_report()`**: Creates a detailed Excel report using the data in the `sending_results` list and saves it to the specified report directory.
* **`open_last_report()`**: Opens the most recently generated report file.
* **`clear_report_results()`**: Clears the report results (`sending_results`) and the report `Treeview`().

### 7. Helper Functions and Arayüz Güncellemeleri
* **`update_status()`**: Updates the status bar in the GUI.
* **`update_progress_bar()`**: Updates the sending progress bar.
* **`on_closing()`**: Executes when the application window is closed, ensuring the WebDriver is safely quit and running threads are stopped.
* **`handle_error()`**: Displays error messages to the user in an understandable way and updates the status bar.
* **`highlight_element()`**: Temporarily highlights a web element visually, aiding in debugging.

## Kullanılan XPath'ler

The code uses XPaths to locate elements within the WhatsApp Web interface. These XPaths are defined in the `DEFAULT_XPATH_SETTINGS` dictionary and can be customized via `settings.json`. If WhatsApp Web's interface changes, these XPaths may need to be updated.

## Stil ve Tema

Custom button styles (Green, Red, Gray) are defined for the interface using `ttk.Style`. This gives the interface a more modern and functional appearance.