# WASDE otomasyonunu Mac’e kurma

Bu paket her USDA WASDE raporunda buğday, mısır, soya fasulyesi, soya yağı ve şeker için mevcut Excel kitaplarına aylık sayfalar ekler. İlk kurulumu aşağıdaki sırayla yapın. Otomatik çalışma en son açılır.

## 1. Gerekenleri hazırlayın

1. **Mac için Microsoft Excel** kurulu ve etkin olmalı. Finder’da **Uygulamalar → Microsoft Excel** yolundan Excel’i bir kez açın, giriş ve lisans adımlarını bitirin. Sonra **⌘Q** ile Excel’den çıkın. Kırmızı pencere düğmesi yalnızca pencereyi kapatır.
2. [Python 3.13 macOS yükleyicisini](https://www.python.org/downloads/macos/) indirin. Python 3.13 sürümleri arasından **macOS 64-bit universal2 installer** seçin. İndirilen `.pkg` dosyasını açıp kurulumu tamamlayın.
3. WASDE ZIP dosyasına Finder’da çift tıklayın. İçinden çıkan `wasde-desktop` klasörünü Finder’da **Ev klasörünüze** taşıyın. Finder’da Ev klasörünü açmak için **⇧⌘H** tuşlarına basabilirsiniz.

## 2. Kurulum komutunu çalıştırın

**⌘Boşluk** tuşlarına basın, **Terminal** yazın ve Enter’a basın. Terminal komut yazdığınız uygulamadır. Aşağıdaki satırları tek tek yapıştırın ve her satırdan sonra Enter’a basın. Yeni bir Terminal açtığınızda tekrar ilk satırla başlayın.

```bash
cd "$HOME/wasde-desktop"
python3.13 --version
bash install-mac.sh
```

İkinci komutun `Python 3.13.x`, son komutun `PASS: Installation complete` yazması beklenir. Kurulum `.venv` adlı proje içi Python ortamını, `config.json` ayar dosyasını ve boş `workbooks` klasörünü hazırlar. Python komutu bulunamazsa kurulumdan sonra Terminal’i kapatıp yeniden açın. `cd` klasörü bulamazsa Finder’da klasörün Ev klasörü altında olduğunu ve adının `wasde-desktop` olduğunu kontrol edin. Alternatif olarak Terminal’e `cd ` yazıp Finder’daki klasörü Terminal penceresine sürükleyin ve Enter’a basın.

## 3. Beş Excel dosyasını getirin

**Paket Excel kitabı içermez.** Uyumlu kitaplarınız yoksa [beş Excel kitabını sıfırdan hazırlama kılavuzunu](EXCEL_HAZIRLAMA.md) uygulayın. Eski bilgisayarda hazır kitaplarınız varsa aşağıdaki taşıma adımlarını izleyin.

İlk deneme için eski bilgisayardaki dosyaların **kopyalarını** kullanın. Finder’da `wasde-desktop → workbooks` klasörüne aşağıdaki beş dosyayı ekleyin.

| Ürün | Varsayılan dosya adı |
|---|---|
| Buğday | `Wheat Wasde.xlsx` |
| Mısır | `Corn Wasde(1).xlsx` |
| Soya fasulyesi | `Soybean Wasde(1).xlsx` |
| Soya yağı | `Soyoil Wasde(1).xlsx` |
| Şeker | `Sugar Wasde(1).xlsx` |

Dosya adları farklıysa `config.json` dosyasını TextEdit’te düz metin olarak açıp ilgili adları güncelleyin. Tırnakları ve virgülleri koruyun. Varsayılan `workbooks` klasörünü kullanıyorsanız klasör ayarını düzenlemeniz gerekmez. Daha sonra başka bir klasör kullanmak isterseniz `workbook_directory` alanına `~/WASDE-Excel` gibi bir yol yazabilirsiniz. OneDrive kullanacaksanız Finder’dan Mac’teki gerçek klasör yolunu alın, dosyaları cihazda tutun ve eşitlemenin bitmesini bekleyin.

Windows’taki görev kaydı ve Python ortamı Mac kurulumu sırasında yeniden oluşturulur. Önceki işlem kaydını korumak istiyorsanız yalnızca **aynı beş üretim dosyasına ait** `wasde_state.json` dosyasını üretime geçerken `runtime` klasörüne koyun.

## 4. Önce Excel’e ve rapora erişimi kontrol edin

Excel’den **⌘Q** ile çıkın. Terminal’de çalıştırın:

```bash
.venv/bin/python wasde_automation.py --check-install
.venv/bin/python wasde_automation.py --report 2026-11 --dry-run
```

İlk komut Excel’i açar, geçici bir çalışma kitabında formül hesaplar, kitabı kaydeder ve açarak sonucu doğrular. Mac, Terminal veya Python’un Excel’i kontrol etmesi için izin isterse **İzin Ver** seçin. Yanlışlıkla reddettiyseniz **Sistem Ayarları → Gizlilik ve Güvenlik → Otomasyon** altında listelenen uygulamanın **Microsoft Excel** iznini açıp tekrar deneyin. [Apple yardım sayfası](https://support.apple.com/guide/mac-help/mchl07817563/mac)

İkinci komut USDA raporunu indirip beş ürünü kontrol eder. Çalışma kitaplarına yazmaz. Kasım örneğinde `VALIDATED USDA WASDE 2026-11` ve buğday dünya üretimi için `827.51 MMT` görmelisiniz. İndirme başarısızsa [USDA arşivinden Kasım TXT dosyasını](https://esmis.nal.usda.gov/sites/default/release-files/796142/wasde1126.txt) İndirilenler klasörüne kaydedip aşağıdaki komutu deneyin.

```bash
.venv/bin/python wasde_automation.py --report 2026-11 --file "$HOME/Downloads/wasde1126.txt" --dry-run
```

## 5. Deneme kopyalarına yazdırın

Bu örnek için beş kitabın her birinde Ekim 2026 ABD Balance Sheet sayfası bulunmalı. Kasım sayfaları henüz yoksa:

```bash
.venv/bin/python wasde_automation.py --report 2026-11
```

İndirmeyi önceki adımda TXT dosyasıyla yaptıysanız aynı komuta `--file "$HOME/Downloads/wasde1126.txt"` ekleyebilirsiniz. Kasım sayfaları zaten varsa, bu sayfaların yeniden oluşturulmasını **yalnızca deneme kopyalarında** sınamak için sonuna `--repair-current` ekleyin. Bu seçenek seçilen ayın mevcut sayfalarını yeniden oluşturur.

`UPDATED five Excel workbooks` mesajından sonra Excel’i açıp her kitapta iki Kasım sayfası ve önceki Ekim sayfalarını kontrol edin. Yedekleri Finder’da `runtime → backups` altında bulabilirsiniz. Üretime geçerken deneme kitaplarını kaldırıp güncel beş üretim kitabını kopyalayın. Denemenin `runtime/wasde_state.json` dosyasını çalışma klasöründen çıkarın ve varsa aynı üretim kitaplarına ait eski durum dosyasını yerleştirin.

## 6. Otomatik çalışmayı açın

Mac, oturum açıldığında `launchd` adlı yerleşik zamanlayıcıyı kullanır. Önce üretim dosyalarını yerleştirin, Excel’den **⌘Q** ile çıkın ve zamanlayıcının kendi Excel erişimini test edin:

```bash
.venv/bin/python task-mac.py test
```

Bu komut geçici bir görev oluşturur ve Excel testinden sonra görevi kaldırır. İzin sorulursa Excel erişimine izin verin. Terminal’deki test geçmiş olsa bile bu adım ayrıca gereklidir, çünkü macOS zamanlayıcı için farklı bir uygulama kimliği kullanabilir. **Zamanlayıcı testi PASS verdikten sonra** normal görevi açın:

```bash
.venv/bin/python task-mac.py register
.venv/bin/python task-mac.py status
```

Mac açık, kullanıcı oturumu açık ve Excel kapalı olmalıdır. Görev her iki dakikada bir uyandırılır, USDA’ya yalnızca `release-dates.json` içindeki yayın günlerinde **18.55–21.55 Türkiye saati** aralığında bakar. Mac’in saat dilimi farklı olsa da Türkiye saati kullanılır. Mac uyuduysa veya Excel açıksa, uygun olduğunda Excel’i kapatıp elle güncelleyebilirsiniz.

| İşlem | Terminal komutu |
|---|---|
| Son yayımlanan raporu elle kontrol et | `.venv/bin/python wasde_automation.py --now` |
| Zamanlayıcının durumuna bak | `.venv/bin/python task-mac.py status` |
| Geçici olarak durdur | `.venv/bin/python task-mac.py disable` |
| Tekrar çalıştır | `.venv/bin/python task-mac.py enable` |
| Zamanlayıcı kaydını kaldır | `.venv/bin/python task-mac.py remove` |

Görevi kaldırmak Excel kitaplarını ve yedekleri silmez. Klasörün yerini değiştirirseniz önce görevi kaldırın, yeni yerde kurulumu ve zamanlayıcı testini tekrarlayın. Aynı üretim dosyalarına yazan eski Windows görevini Mac görevini açmadan önce devre dışı bırakın.

## Bir şey çalışmazsa

| Gördüğünüz durum | Yapılacak işlem |
|---|---|
| Excel açık uyarısı | Microsoft Excel’i öne getirip **⌘Q** yapın, sonra komutu yeniden çalıştırın. |
| Dosya eksik uyarısı | Beş dosyanın `workbooks` klasöründeki adlarını kontrol edin. |
| Excel izni hatası veya zamanlayıcı testi başarısız | **Gizlilik ve Güvenlik → Otomasyon** ayarlarını ve `runtime/launchd-test.log` dosyasını kontrol edin. Zamanlayıcı izni verilemiyorsa Terminal’den elle çalıştırın. |
| Rapor ayı zaten var | `NO CHANGE` önceki işlem kaydının eşleştiğini gösterir. Sayfalar zaten varsa ve düzeltilmesi gerekiyorsa yalnızca seçili ayı önce deneme kopyalarında `--repair-current` ile sınayın. |
| Çalıştırma hatası | Finder’da `runtime/wasde_automation.log` ve zamanlayıcı için `runtime/launchd.log` dosyalarını açın. |

Windows için [Windows kurulum kılavuzuna](README.md) bakın.
