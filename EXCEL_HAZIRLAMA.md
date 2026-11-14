# Kendi WASDE Excel dosyalarınızı hazırlayın

Paket kodu ve kurulum adımlarını içerir. **Beş Excel çalışma kitabını siz oluşturup `workbooks` klasörüne koymalısınız.** Uyumlu kitaplarınız zaten varsa ilk denemede bunların kopyalarını kullanabilirsiniz.

## 1. Beş boş çalışma kitabı oluşturun

1. Kurulum kılavuzundaki yükleme adımını tamamlayın. Bu adım proje klasörü içinde `workbooks` klasörünü oluşturur.
2. Microsoft Excel’de **Dosya → Yeni → Boş çalışma kitabı** seçin.
3. **Dosya → Farklı Kaydet** ile dosyayı **Excel Çalışma Kitabı (.xlsx)** biçiminde `workbooks` klasörüne kaydedin. Her ürün için ayrı dosya oluşturun.

| Ürün | Dosya adı |
|---|---|
| Buğday | `Wheat Wasde.xlsx` |
| Mısır | `Corn Wasde(1).xlsx` |
| Soya fasulyesi | `Soybean Wasde(1).xlsx` |
| Soya yağı | `Soyoil Wasde(1).xlsx` |
| Şeker | `Sugar Wasde(1).xlsx` |

Farklı dosya adları kullanıyorsanız `config.json` içindeki ürün eşleştirmelerini düzenleyin.

## 2. Her dosyada bir başlangıç sayfası hazırlayın

1. İlk işleyeceğiniz rapor ayını seçin. Örneğin `2026-11` ile başlayacaksanız beş kitapta da alttaki **Sayfa1** sekmesine çift tıklayın ve adını **`October'26 Balance Sheet`** yapın. Aralık 2026 ile başlayacaksanız **`November'26 Balance Sheet`** kullanın. Ocak 2027 için **`December'26 Balance Sheet`** kullanın. Ayın İngilizce tam adını ve normal düz kesme işaretini yazın.
2. Aşağıdaki tablodan ilgili ürünün satırlarını bulun. Etiketleri **B sütununa** tam karşılarındaki satır numarasıyla yazın. Örneğin şeker için `B6` hücresine **Beginning Stocks**, `B7` hücresine **Production** yazın. Atlanan satırları boş bırakın. Excel’in sol üstündeki ad kutusuna `B6` yazıp Enter’a basarak hücreye doğrudan gidebilirsiniz.
3. **C:G sütunlarını** programın dolduracağı alan olarak boş bırakın. Program yeni ayın sayfasında C ve D’ye tarihsel dönem sütunlarını, E ve F’ye aynı USDA raporundaki önceki ve güncel projeksiyonları, G’ye de **F − E** değişimini yazar. Soya fasulyesinde başlık alanı **B2:G2**, diğer dört üründe **B3:G3**. Başlık ve veri alanlarındaki hücreleri birleştirmeyin. Sütun genişliği, renk ve yazı tipi gibi biçimleri ayarlayabilirsiniz.
4. Beş dosyayı kaydedip Excel’i kapatın. Mac’te **⌘Q** ile Excel’den çıkın.

Başlangıç sayfası biçim ve satır etiketi sağlar. Önceki aya ait rakamları elle girmeniz gerekmez. Başlangıç sayfası, geçmiş ayın gerçek verisi olarak kullanılacaksa o aya ait rakamları ayrıca doldurun. İlk güncelleme, seçilen ay için **Balance Sheet** ile **World** veya **Mexico Sugar** sayfasını otomatik oluşturur.

## 3. İlk çalıştırmayı kontrol edin

Kendi işletim sisteminizin kılavuzundaki `--check-install` ve `--dry-run` komutlarını çalıştırın. Ardından aynı rapor ayı için beş deneme kitabında gerçek yazma adımını uygulayın. `UPDATED five Excel workbooks` mesajını ve her dosyadaki iki yeni sekmeyi kontrol edin. İlk ay başarıyla yazıldıktan sonra bir sonraki ayın başlangıç sayfası olarak bu yeni Balance Sheet sayfası kullanılır.

- [Mac kurulum kılavuzu](MAC_KURULUM.md)
- [Windows kurulum kılavuzu](README.md)

## B sütununa yazılacak etiketler

### Buğday

| Hücre | Etiket |
|---|---|
| `B6` | Area Planted |
| `B7` | Area Harvested |
| `B9` | Yield per Harvested Acre |
| `B11` | Beginning Stocks |
| `B12` | Production |
| `B13` | Imports |
| `B14` | Supply Total |
| `B15` | Food |
| `B16` | Seed |
| `B17` | Feed and Residual |
| `B18` | Domestic Total |
| `B19` | Exports |
| `B20` | Use Total |
| `B21` | Ending Stocks |
| `B22` | Average Farm Price |

### Mısır

| Hücre | Etiket |
|---|---|
| `B6` | Area Planted |
| `B7` | Area Harvested |
| `B9` | Yield per Harvested Acre |
| `B11` | Beginning Stocks |
| `B12` | Production |
| `B13` | Imports |
| `B14` | Supply Total |
| `B15` | Feed and Residual |
| `B16` | Food, Seed & Industrial |
| `B17` | Ethanol & By-products |
| `B18` | Domestic Total |
| `B19` | Exports |
| `B20` | Use Total |
| `B21` | Ending Stocks |
| `B22` | Average Farm Price |

### Soya fasulyesi

| Hücre | Etiket |
|---|---|
| `B5` | Area Planted |
| `B6` | Area Harvested |
| `B8` | Yield per Harvested Acre |
| `B10` | Beginning Stocks |
| `B11` | Production |
| `B12` | Imports |
| `B13` | Supply Total |
| `B14` | Crushings |
| `B15` | Exports |
| `B16` | Seed |
| `B17` | Residual |
| `B18` | Use Total |
| `B19` | Ending Stocks |
| `B20` | Average Farm Price |

### Soya yağı

| Hücre | Etiket |
|---|---|
| `B6` | Beginning Stocks |
| `B7` | Production |
| `B8` | Imports |
| `B9` | Supply Total |
| `B10` | Domestic Disappearance |
| `B11` | Biofuel |
| `B12` | Industrial |
| `B13` | Exports |
| `B14` | Use Total |
| `B15` | Ending Stocks |
| `B16` | Average Price |

### Şeker

| Hücre | Etiket |
|---|---|
| `B6` | Beginning Stocks |
| `B7` | Production |
| `B8` | Beet Sugar |
| `B9` | Cane Sugar |
| `B10` | Florida |
| `B11` | Louisiana |
| `B12` | Texas |
| `B13` | Imports |
| `B14` | TRQ |
| `B15` | Other Program |
| `B16` | Nonprogram |
| `B17` | Mexico |
| `B18` | High-tier Tariff / Other |
| `B19` | Total Supply |
| `B20` | Exports |
| `B21` | Deliveries |
| `B22` | Food |
| `B23` | Other |
| `B24` | Miscellaneous |
| `B25` | Total Use |
| `B26` | Ending Stocks |
| `B27` | Stocks-to-Use Ratio |
