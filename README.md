# WhatsApp Otomasyon Aracı

Bu proje, hiç kodlama deneyimi olmayan biri tarafından yapay zeka araçları kullanılarak geliştirilmiş, WhatsApp Web üzerinden otomatik mesaj, doküman ve medya gönderimi yapmaya yarayan bir otomasyon aracıdır. Windows 10 işletim sistemi ve Python 3.13 versiyonu kullanılarak test edilmiştir.

## Özellikler

* **Kullanıcı Dostu Arayüz:** `tkinter` ile oluşturulmuş basit ve anlaşılır grafik arayüz.
* **Çoklu Gönderim Tipleri:** Metin mesajları, dokümanlar, görseller/videolar ve hatta Excel'deki belirli hücrelerden alınan görselleri gönderebilme.
* **Excel Entegrasyonu:** Kişi listesini Excel dosyasından (numara, işlem türü, mesaj vb. bilgilerle) içeri aktarabilme.
* **Tarayıcı Desteği:** Chrome, Firefox ve Edge tarayıcıları ile uyumluluk.
* **Otomatik WebDriver Yönetimi:** WebDriver'lar (tarayıcı sürücüleri) otomatik olarak indirilir ve güncellenir.
* **İlerleme ve Durum Takibi:** Gönderim süreci boyunca arayüzde durum güncellemeleri ve ilerleme çubuğu.
* **Raporlama:** Gönderim sonuçlarını detaylı bir rapor olarak Excel'e aktarabilme.
* **Durdurma ve İptal Etme:** Devam eden gönderim işlemini duraklatma, devam ettirme veya tamamen iptal etme yeteneği.
* **Ayarları Kaydetme/Yükleme:** Kullanıcı tanımlı XPath ayarlarını ve rapor klasörünü kaydedip yükleyebilme.

## Kurulum

Bu aracı çalıştırmak için Python'ın sisteminizde yüklü olması gerekmektedir. Eğer yoksa, Python'ın resmi web sitesinden (python.org) 3.13 veya uyumlu bir sürümünü indirip kurmanız önerilir.

1.  **Depoyu Klonlayın:**
    ```bash
    git clone [https://github.com/netnand1453/WhatsappWeb.git](https://github.com/netnand1453/WhatsappWeb.git)
    cd WhatsappWeb
    ```

2.  **Gerekli Kütüphaneleri Yükleyin:**
    ```bash
    pip install pandas openpyxl selenium webdriver_manager pillow pywin32
    ```
    * **Not:** `pywin32` kütüphanesi sadece Windows işletim sisteminde "Excel'den Medya" gönderme özelliği için gereklidir. Diğer işletim sistemlerinde kurulmasına gerek yoktur, ancak kurulması hata vermez.

## Kullanım

1.  **`wpoto2.py`'yi Çalıştırın:**
    Terminalde veya komut istemcisinde `python wpoto2.py` komutunu kullanarak uygulamayı başlatın.

2.  **Ayarları Yapılandırın:**
    * **"Ayarlar"** sekmesine gidin.
    * **"Rapor Klasörü Seç"** düğmesi ile raporların kaydedileceği bir klasör belirleyin.
    * Uygulamanın doğru çalışması için **XPath ayarlarının doğru olduğundan emin olun**. WhatsApp Web arayüzünde önemli değişiklikler olduğunda bu XPath'lerin güncellenmesi gerekebilir. Varsayılanlar çoğu durumda çalışmalıdır.
    * **"Ayarları Kaydet"** düğmesine tıklayarak ayarlarınızı `settings.json` dosyasına kaydedin.

3.  **Kişi Listesini Oluşturun:**
    * **"Kişi Listesi"** sekmesine gidin.
    * Excel'den içe aktarmak için **"Excel'den Kişi Yükle"** düğmesini kullanın. Excel dosyanızın belirli sütun başlıklarına (Ünvan, Alan, Numara, İşlem Türü, Mesaj, Hedef Hücre, Dosya Yolu) sahip olması gerekmektedir.
    * Manuel olarak kişi eklemek için giriş alanlarını kullanın ve **"Kişi Ekle"** düğmesine tıklayın.
    * Kişi Listesi sekmesindeki "İşlem Türü" seçimi, ilgili alanların görünürlüğünü dinamik olarak ayarlayacaktır.
        * **Mesaj:** Sadece metin mesajı gönderir.
        * **Doküman:** Belirtilen doküman dosyasını (PDF, Word vb.) gönderir. Opsiyonel olarak bir başlık metni ekleyebilirsiniz.
        * **Medya:** Belirtilen görsel veya video dosyasını gönderir. Opsiyonel olarak bir alt yazı metni ekleyebilirsiniz.
        * **Excel to Media:** Belirtilen Excel dosyasındaki belirli bir hücreden (örn. `A1`) görsel alır ve bunu medya olarak gönderir. **Bu özellik sadece Windows'ta `pywin32` yüklüyse çalışır.** Opsiyonel olarak bir alt yazı metni ekleyebilirsiniz.

4.  **WhatsApp Web'e Bağlanın:**
    * **"Ana Sayfa"** sekmesine dönün.
    * Kullanmak istediğiniz tarayıcıyı (Chrome, Firefox, Edge) seçin.
    * **"WhatsApp Web'i Başlat ve Bağlan"** düğmesine tıklayın. Seçtiğiniz tarayıcı otomatik olarak açılacak ve WhatsApp Web'e yönlendirileceksiniz.
    * Açılan tarayıcıda WhatsApp Web QR kodunu telefonunuzla tarayarak giriş yapın.

5.  **Mesaj Göndermeyi Başlatın:**
    * WhatsApp Web'e başarıyla giriş yaptıktan sonra, **"Mesajları Gönder"** düğmesi etkin hale gelecektir.
    * Bu düğmeye tıklayarak listedeki kişilere mesajları/medyaları/dokümanları göndermeye başlayın.
    * İlerleme çubuğundan ve durum çubuğundan süreci takip edebilirsiniz.
    * İşlemi duraklatmak için **"Durdur"**, tamamen iptal etmek için **"İptal Et"** düğmelerini kullanın.

6.  **Rapor Oluşturun:**
    * Gönderim tamamlandıktan veya iptal edildikten sonra **"Rapor Oluştur"** düğmesi etkinleşecektir.
    * Bu düğmeye tıklayarak gönderim sonuçlarını Excel raporu olarak kaydedin. Rapor, Ayarlar sekmesinde belirttiğiniz klasöre kaydedilecektir.

## Sorun Giderme

* **Tarayıcı veya WebDriver Hatası:**
    * İnternet bağlantınızın aktif olduğundan emin olun.
    * Antivirüs veya güvenlik duvarı yazılımınızın WebDriver indirme veya tarayıcı başlatma işlemini engellemediğinden emin olun.
    * Seçtiğiniz tarayıcının (Chrome, Firefox, Edge) sisteminizde güncel bir şekilde kurulu olduğundan emin olun.
    * Uygulamayı yönetici olarak çalıştırmayı deneyin.
    * WebDriver Manager'ın önbelleğini manuel olarak temizlemeyi deneyin (genellikle kullanıcı dizininizdeki `.wdm` klasöründedir).
* **XPath Hataları:** WhatsApp Web'in arayüzü sık sık güncellendiği için, mevcut XPath'ler geçersiz hale gelebilir. Eğer uygulama WhatsApp Web öğelerini bulmakta zorlanıyorsa, Ayarlar sekmesindeki XPath değerlerini manuel olarak güncellemeniz gerekebilir. Doğru XPath'leri bulmak için tarayıcınızın geliştirici araçlarını (F12) kullanabilirsiniz.
* **`pywin32` Hatası:** Eğer "Excel to Media" özelliğini kullanırken hata alıyorsanız, `pip install pywin32` komutunu çalıştırdığınızdan emin olun. Bu kütüphane sadece Windows içindir.
* **Excel Dosyası Formatı:** Excel dosyanızdaki sütun başlıklarının (Ünvan, Alan, Numara, İşlem Türü, Mesaj, Hedef Hücre, Dosya Yolu) doğru yazıldığından ve gerekli tüm bilgileri içerdiğinden emin olun.

## Yasal Uyarı

Bu araç, kişisel veya yasal işleriniz için bir otomasyon çözümü sunmak amacıyla geliştirilmiştir. WhatsApp'ın kullanım koşullarını ihlal edecek şekilde kötüye kullanımı tamamen kullanıcının sorumluluğundadır. Bu aracın yasa dışı veya istenmeyen mesajlaşma (spam) faaliyetleri için kullanılması kesinlikle önerilmez ve yasa dışıdır. Geliştirici, aracın kötüye kullanımından doğacak hiçbir sorumluluğu kabul etmez.

## Katkıda Bulunma

Bu proje açık kaynaklıdır ve katkılara açıktır. Her türlü hata düzeltmesi, yeni özellik veya iyileştirme önerisi memnuniyetle karşılanır. Lütfen bir Pull Request açmadan önce Issues (Sorunlar) bölümünde tartışın.

## Lisans

Bu proje MIT Lisansı altında lisanslanmıştır. Daha fazla bilgi için `LICENSE` dosyasına bakınız.

---

# WhatsApp Automation Tool

This project is an automation tool developed by a non-coder using artificial intelligence tools. It enables automated sending of messages, documents, and media via WhatsApp Web. It has been tested on Windows 10 and Python 3.13.

## Features

* **User-Friendly Interface:** Simple and intuitive graphical interface built with `tkinter`.
* **Multiple Sending Types:** Ability to send text messages, documents, images/videos, and even images extracted from specific cells in Excel.
* **Excel Integration:** Import contact lists from Excel files (including number, action type, message, etc.).
* **Browser Support:** Compatibility with Chrome, Firefox, and Edge browsers.
* **Automatic WebDriver Management:** Browser WebDrivers are automatically downloaded and managed.
* **Progress and Status Tracking:** Real-time status updates and a progress bar in the interface during the sending process.
* **Reporting:** Export sending results as a detailed Excel report.
* **Pause and Cancel:** Ability to pause, resume, or completely cancel an ongoing sending process.
* **Save/Load Settings:** Save and load user-defined XPath settings and report directory.

## Installation

To run this tool, Python must be installed on your system. If not, it is recommended to download and install version 3.13 or a compatible version from the official Python website (python.org).

1.  **Clone the Repository:**
    ```bash
    git clone [https://github.com/netnand1453/WhatsappWeb.git](https://github.com/netnand1453/WhatsappWeb.git)
    cd WhatsappWeb
    ```

2.  **Gerekli Kütüphaneleri Yükleyin:**
    ```bash
    pip install pandas openpyxl selenium webdriver_manager pillow pywin32
    ```
    * **Note:** The `pywin32` library is only required on Windows for the "Excel to Media" sending feature. It is not necessary to install it on other operating systems, but installing it will not cause errors.

## Usage

1.  **Run `wpoto2.py`:**
    Start the application by running `python wpoto2.py` in your terminal or command prompt.

2.  **Configure Settings:**
    * Go to the **"Settings"** tab.
    * Use the **"Select Report Directory"** button to specify a folder where reports will be saved.
    * Ensure **XPath settings are correct** for the application to function properly. These XPaths may need to be updated if WhatsApp Web's interface undergoes significant changes. Defaults should work in most cases.
    * Click the **"Save Settings"** button to save your configurations to the `settings.json` file.

3.  **Create Contact List:**
    * Go to the **"Contact List"** tab.
    * Use the **"Load Contacts from Excel"** button to import contacts from an Excel file. Your Excel file should have specific column headers (Title, Area, Number, Action Type, Message, Target Cell, File Path).
    * To add contacts manually, use the input fields and click **"Add Contact"**.
    * The "Action Type" selection in the Contact List tab will dynamically adjust the visibility of relevant input fields:
        * **Message:** Sends a plain text message.
        * **Document:** Sends the specified document file (PDF, Word, etc.). Optionally, you can add a caption.
        * **Media:** Sends the specified image or video file. Optionally, you can add a caption.
        * **Excel to Media:** Extracts an image from a specific cell (e.g., `A1`) in the specified Excel file and sends it as media. **This feature works only on Windows if `pywin32` is installed.** Optionally, you can add a caption.

4.  **Connect to WhatsApp Web:**
    * Return to the **"Home"** tab.
    * Select your desired browser (Chrome, Firefox, Edge).
    * Click the **"Start and Connect WhatsApp Web"** button. Your chosen browser will automatically open and navigate to WhatsApp Web.
    * Scan the WhatsApp Web QR code with your phone to log in.

5.  **Start Sending Messages:**
    * Once you have successfully logged into WhatsApp Web, the **"Start Sending Messages"** button will become active.
    * Click this button to begin sending messages/media/documents to the contacts in your list.
    * Monitor the process via the progress bar and status bar.
    * Use the **"Stop"** button to pause and **"Cancel"** to abort the operation completely.

6.  **Rapor Oluşturun:**
    * Gönderim tamamlandıktan veya iptal edildikten sonra **"Rapor Oluştur"** düğmesi etkinleşecektir.
    * Bu düğmeye tıklayarak gönderim sonuçlarını Excel raporu olarak kaydedin. Rapor, Ayarlar sekmesinde belirttiğiniz klasöre kaydedilecektir.

## Sorun Giderme

* **Tarayıcı veya WebDriver Hatası:**
    * İnternet bağlantınızın aktif olduğundan emin olun.
    * Antivirüs veya güvenlik duvarı yazılımınızın WebDriver indirme veya tarayıcı başlatma işlemini engellemediğinden emin olun.
    * Seçtiğiniz tarayıcının (Chrome, Firefox, Edge) sisteminizde güncel bir şekilde kurulu olduğundan emin olun.
    * Uygulamayı yönetici olarak çalıştırmayı deneyin.
    * WebDriver Manager'ın önbelleğini manuel olarak temizlemeyi deneyin (genellikle kullanıcı dizininizdeki `.wdm` klasöründedir).
* **XPath Hataları:** WhatsApp Web'in arayüzü sık sık güncellendiği için, mevcut XPath'ler geçersiz hale gelebilir. Eğer uygulama WhatsApp Web öğelerini bulmakta zorlanıyorsa, Ayarlar sekmesindeki XPath değerlerini manuel olarak güncellemeniz gerekebilir. Doğru XPath'leri bulmak için tarayıcınızın geliştirici araçlarını (F12) kullanabilirsiniz.
* **`pywin32` Hatası:** Eğer "Excel to Media" özelliğini kullanırken hata alıyorsanız, `pip install pywin32` komutunu çalıştırdığınızdan emin olun. Bu kütüphane sadece Windows içindir.
* **Excel Dosyası Formatı:** Excel dosyanızdaki sütun başlıklarının (Ünvan, Alan, Numara, İşlem Türü, Mesaj, Hedef Hücre, Dosya Yolu) doğru yazıldığından ve gerekli tüm bilgileri içerdiğinden emin olun.

## Yasal Uyarı

Bu araç, kişisel veya yasal işleriniz için bir otomasyon çözümü sunmak amacıyla geliştirilmiştir. WhatsApp'ın kullanım koşullarını ihlal edecek şekilde kötüye kullanımı tamamen kullanıcının sorumluluğundadır. Bu aracın yasa dışı veya istenmeyen mesajlaşma (spam) faaliyetleri için kullanılması kesinlikle önerilmez ve yasa dışıdır. Geliştirici, aracın kötüye kullanımından doğacak hiçbir sorumluluğu kabul etmez.

## Katkıda Bulunma

Bu proje açık kaynaklıdır ve katkılara açıktır. Her türlü hata düzeltmesi, yeni özellik veya iyileştirme önerisi memnuniyetle karşılanır. Lütfen bir Pull Request açmadan önce Issues (Sorunlar) bölümünde tartışın.

## Lisans

Bu proje MIT Lisansı altında lisanslanmıştır. Daha fazla bilgi için `LICENSE` dosyasına bakınız.