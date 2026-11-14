# USDA WASDE Excel otomasyonu

Mac kullanıyorsanız [adım adım macOS kurulumuna](MAC_KURULUM.md) geçin. Bu sayfa Windows kurulumu içindir.

Windows kurulum paketi, 8 Kasım 2026'da çalışan beş ürünlü sürümün **11 Kasım 2026 aylık sayfa yamasını** temel alır. Her rapor ayı, her çalışma kitabında iki sayfa oluşturur. Buğday, mısır, soya fasulyesi ve soya yağında ABD ve dünya tabloları, şekerde ABD ve Meksika tabloları yazılır. ABD sayfasının biçimi bir önceki ayın sayfasından alınır. Ayın dosyaları önce geçici kopyalarda hazırlanır ve doğrulanır. Güncelleme öncesi beş dosyanın yedeği alınır.

## 1. Windows kurulumu

Windows 10/11, etkinleştirilmiş masaüstü Microsoft Excel, Python 3.13 ve `py` başlatıcısı gerekir. Excel'i görev için kullanacağınız Windows hesabında bir kez açıp ilk kullanım ekranlarını tamamlayın. Git kullanacaksanız Git for Windows ve hedef depoya erişim de gerekir. ZIP üzerinden kurulumda arşivi yerel bir klasöre açıp aşağıdaki kurulum komutuyla başlayın.

Depo hazır olduğunda, PowerShell'de `ORG` ve `REPO` değerlerini gerçek depo adıyla değiştirin.

```powershell
git clone https://github.com/ORG/REPO.git "$env:LOCALAPPDATA\WASDE-Automation"
if ($LASTEXITCODE -ne 0) { throw 'Git clone failed.' }
Set-Location "$env:LOCALAPPDATA\WASDE-Automation"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

Kurulum `.venv` açar, `requirements.txt` içindeki `pywin32==312` paketini kurar ve üç regresyon testini çalıştırır. Var olan `config.json` dosyası korunur. Klasörü daha sonra taşırsanız yeni konumda sanal ortamı yeniden kurun ve görevi yeniden kaydedin.

## 2. Arkadaşınızın bilgisayarına kopyalanacak dosyalar

**Paket Excel kitabı içermez.** Uyumlu kitaplarınız yoksa [beş Excel kitabını sıfırdan hazırlama kılavuzunu](EXCEL_HAZIRLAMA.md) uygulayın. Eski bilgisayarda hazır kitaplarınız varsa aşağıdaki taşıma adımlarını izleyin.

Eski bilgisayarda Excel'i kapatın ve OneDrive kullanılıyorsa eşitlemenin bitmesini bekleyin. Eski bilgisayardan **en son doğrulanmış beş çalışma kitabının** birer kopyasını alın. Bunları yeni kurulumun `workbooks\` klasörüne koyun.

| Ürün | Varsayılan dosya adı |
|---|---|
| Buğday | `Wheat Wasde.xlsx` |
| Mısır | `Corn Wasde(1).xlsx` |
| Soya fasulyesi | `Soybean Wasde(1).xlsx` |
| Soya yağı | `Soyoil Wasde(1).xlsx` |
| Şeker | `Sugar Wasde(1).xlsx` |

İşlem geçmişini de taşıyacaksanız, **aynı beş dosyaya ait** `wasde_state.json` dosyasını `runtime\wasde_state.json` olarak kopyalayın. Excel dosyaları ve state farklı yedeklerden geldiyse önce aylık sayfaları kontrol edin. Önceki yedekleri arşivlemek isterseniz eski `backups\` klasörünü ayrıca kopyalayabilirsiniz. Kod ve testler paketle gelir. Yeni bilgisayar `.venv`, günlük ve Task Scheduler kaydını kendi oluşturur.

Excel dosya adları farklıysa `config.json` içindeki ilgili değerleri düzenleyin. Dosyalar başka bir klasörde kalacaksa `workbook_directory` değerini değiştirin. Göreli yol `config.json` konumuna göre çözülür. Mutlak yollar ve Windows ortam değişkenleri kullanılabilir. Örneğin `%USERPROFILE%/Documents/WASDE` veya ilgili Windows hesabında tanımlıysa `%OneDrive%/WASDE`. JSON içinde ileri eğik çizgi kullanabilirsiniz. Bir OneDrive klasörü kullanırsanız beş dosyayı **Her zaman bu cihazda tut** olarak işaretleyin. Tek bir aktif zamanlayıcıyı aynı çalışma kitabı kopyalarına yönlendirin.

## 3. İki bilgisayar arasında geçişi doğrulama

Önce komutları beş Excel'in **deneme kopyalarıyla** çalıştırın. İlk kez yazılacak ayın bir önceki ayına ait ABD Balance Sheet sayfası her kitapta bulunmalıdır. Kasım 2026 testi için Ekim sayfalarını içeren kopyaları kullanın.

```powershell
$Python = '.\.venv\Scripts\python.exe'
& $Python .\wasde_automation.py --check-install
& $Python .\wasde_automation.py --report 2026-11 --dry-run
```

`--check-install` dosya konumları ve Excel COM'u, `--dry-run` USDA raporunun beş ürün için ayrıştırılmasını kontrol eder. Kasım raporu için beklenen buğday dünya üretimi **827.51 MMT**. USDA'ya doğrudan erişim sorununda [USDA arşivinden](https://esmis.nal.usda.gov/publication/world-agricultural-supply-and-demand-estimates) ilgili yayının TXT dosyasını indirip şu komutu kullanabilirsiniz.

```powershell
& $Python .\wasde_automation.py --report 2026-11 --file 'C:\Temp\wasde1126.txt' --dry-run
```

Kasım sayfaları deneme kopyalarında henüz yoksa gerçek yazma testi için aşağıdaki komutu çalıştırın. Eski durum kaydını bu teste taşımayın. Durum kaydı aynı raporun zaten işlendiğini belirtiyorsa normal komut `NO CHANGE` döndürür.

```powershell
& $Python .\wasde_automation.py --report 2026-11
```

`UPDATED five Excel workbooks` mesajını, beş kitapta toplam on Kasım sayfasını, Ekim sayfalarını ve `runtime\backups\` içindeki beş yedeği kontrol edin. Kasım zaten varsa ve özellikle Kasım sayfalarını yeniden oluşturmanız gerekiyorsa, önce deneme kopyalarında `--report 2026-11 --repair-current` kullanın. Rapor saatini beklemeden son yayımlanmış ayı elle kontrol etmek için `--now` kullanabilirsiniz.

Üretim kopyalarına geçtiğinizde testin `runtime\wasde_state.json` dosyasını arşivleyip çalışma dizininden çıkarın. Üretim için varsa yalnızca aynı beş üretim kitabına ait durum dosyasını yerleştirin. Yeni görev devreye alınmadan önce aynı çalışma kitaplarını güncelleyen eski görevi durdurun veya devre dışı bırakın.

## 4. Task Scheduler

Komutları Excel'i kullanacak Windows hesabında repo klasöründen çalıştırın. Görev, Windows saat dilimi `Turkey Standard Time` iken kaydedilir ve kullanıcı oturumu açıkken çalışır. `release-dates.json` içindeki USDA yayın günlerinde 18.55–21.55 Türkiye saati arasında iki dakikada bir kontrol için tetiklenir. USDA'nın öğlen Eastern Time yayını Türkiye'de ABD yaz saatinde 19.00, kış saatinde 20.00'a denk gelir. Windows oturumu ve bilgisayar açık olmalıdır. Yayın penceresi kaçırılırsa `--now` ile elle kontrol edin.

```powershell
# Gecici gorevle Excel COM ve Python'u Scheduler icinden test et
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\task.ps1 -Action Test

# Uretim gorevini olustur
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\task.ps1 -Action Register

# Son calisma sonucu ve sonraki tetikleme
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\task.ps1 -Action Status

# Devre disi birak veya yeniden etkinlestir
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\task.ps1 -Action Disable
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\task.ps1 -Action Enable

# Kaldir
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\task.ps1 -Action Remove
```

Test görevi Excel dosyalarına yazmadan kendi çalışmasını doğrular ve kendini kaldırır. Aynı isimde bir üretim görevi varsa önce eski görevin bitmesini bekleyin, sonra `Remove` ve `Register` çalıştırın. Takvim dosyası 2026–2027 tarihlerini içerir. [USDA yayın takvimi](https://esmis.nal.usda.gov/sites/default/release-files/796142/wasde1126.pdf) değiştiğinde yeni tarihleri `YYYY-MM: YYYY-MM-DD` eşlemeleri olarak düzenleyin, görevi kaldırıp yeniden kaydedin.

Hata ayrıntıları `runtime\wasde_automation.log` içinde. Açık veya salt okunur çalışma kitabı hatasında Excel'i kapatın, eşitlemenin tamamlanmasını bekleyin ve elle tekrar çalıştırın. Hatalı ya da yarıda kalan bir aktarımdan geri dönmek için görevi devre dışı bırakın, güncelleme öncesi yedekten **beş kitabı birlikte** geri yükleyin ve aynı sürüme ait state dosyasını geri koyun. Eşleşen state yedeği yoksa mevcut state'i arşivleyip çalışma dizininden çıkarın.

## Dosya yapısı

| Dosya | İşlev |
|---|---|
| `wasde_automation.py` | USDA TXT indirme, doğrulama ve Excel güncellemesi |
| `requirements.txt` | Windows Python bağımlılığı |
| `config.example.json` | Yerel Excel dosya konumu ve beş dosya adı |
| `release-dates.json` | USDA yayın günleri |
| `install.ps1` | Sanal ortam, bağımlılık ve çevrimdışı testler |
| `task.ps1` | Task Scheduler yönetimi ve Excel COM testi |
| `test_*.py` | Ayrıştırma, aylık sayfa ve veri doğrulama regresyonları |
| `.gitignore` | Yerel yapılandırma, Excel'ler ve çalışma çıktıları için Git kuralları |

Windows'ta 11 Kasım yamasının bu taşınabilir sürümüyle yapılacak ilk gerçek yazma denemesi, üretim görevi açılmadan önce tamamlanmalıdır.


## English overview

This project downloads the monthly USDA World Agricultural Supply and Demand Estimates (WASDE) report and updates five local Microsoft Excel workbooks for wheat, corn, soybeans, soybean oil, and sugar. Each release adds a U.S. balance sheet and a second worksheet for world data or, for sugar, Mexico. The change column compares the previous and current projections in the same USDA report. The script backs up the workbooks before applying an update.

### Getting started

1. Install Python 3.13 and activate the desktop version of Microsoft Excel. Clone or download this repository, then open PowerShell (Windows) or Terminal (macOS) in the project folder.
2. On Windows, run `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1`. On macOS, run `bash install-mac.sh`. The installer creates the local Python environment, installs dependencies, and runs regression tests.
3. Prepare your own five `.xlsx` workbooks and place them in `workbooks/`. See the [workbook preparation guide](EXCEL_HAZIRLAMA.md) for exact filenames and cell labels. Each workbook needs a U.S. Balance Sheet tab for the month before your first update. For a November 2026 update, name that seed tab `October'26 Balance Sheet`. You can leave its historical values blank. The script fills the new release tabs from the USDA report. If your filenames differ, edit `config.json`.
4. On Windows use `.\.venv\Scripts\python.exe` and on macOS use `.venv/bin/python` to run `wasde_automation.py --check-install`, followed by `wasde_automation.py --report 2026-11 --dry-run`. The dry run parses the report without writing to your workbooks. Use the month you are testing and verify a real update on copies of all five workbooks before enabling the scheduler.
5. Use Windows Task Scheduler or the macOS `launchd` agent to run the automation on USDA release days. Test the scheduled job under the account that will run it. On macOS, allow access to Excel under System Settings → Privacy & Security → Automation when prompted. See the [macOS setup guide](MAC_KURULUM.md) or the Turkish Windows instructions above for scheduler commands.

Workbooks, local settings, logs, and backups are stored on your computer. Keep them out of commits.
