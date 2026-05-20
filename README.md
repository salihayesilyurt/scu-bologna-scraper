# Bologna Ders -> Excel Aracı — Kullanım Kılavuzu

Bu araç, verdiğiniz Cumhuriyet Üniversitesi Bologna ders linklerini (veya
`curCourse` ID'lerini) tek tek çekip, belirlediğimiz kurallara göre
`Dersler.xlsx` şablonunu doldurur. 40 link de olsa 400 link de olsa aynı
şekilde çalışır — her seferinde tek komut.

## 1. Kurulum (tek seferlik)

Bilgisayarınızda Python 3 kurulu olmalı. Sonra:

    pip install requests beautifulsoup4 openpyxl

## 2. Dosyaları aynı klasöre koyun

    bologna_scraper.py            <- uygulama
    links.txt                     <- işlenecek linkler/ID'ler
    Dersler.xlsx                  <- boş şablon (YÖKSİS'ten indirdiğiniz orijinal)
    2025-Guz-Mevcut-Ders.xlsx     <- öğrenci sayıları (Güz)
    2025-Bahar-Mevcut-Ders.xlsx   <- öğrenci sayıları (Bahar)

Not: "2025-Güz-..." dosya adındaki Türkçe karakter sorun çıkarırsa adını
`2025-Guz-Mevcut-Ders.xlsx` yapın; uygulama her iki yazımı da arar.

## 3. Linkleri yazın

`links.txt` içine her satıra bir link **veya** sadece ID yazın.

## 4. Çalıştırın

    python bologna_scraper.py

İsteğe bağlı:

    python bologna_scraper.py --links links.txt --template Dersler.xlsx --out Dersler_dolu.xlsx
    python bologna_scraper.py --ogrenci 2025-Guz-Mevcut-Ders.xlsx 2025-Bahar-Mevcut-Ders.xlsx

Sonuç `Dersler_dolu.xlsx` olarak kaydedilir.


## Linkleri OTOMATIK toplama (links.txt'i kendin oluşturmak)

Artık linkleri elle toplamana gerek yok. `link_toplayici.py`, bir program
(müfredat) sayfasındaki tüm derslerin detay linklerini otomatik çıkarıp
`links.txt` dosyasını kendisi oluşturur.

    pip install requests beautifulsoup4

    python link_toplayici.py "https://obs.cumhuriyet.edu.tr/oibs/bologna/index.aspx?lang=tr&curOp=showPac&curUnit=33&curSunit=123"

veya sadece curSunit numarasıyla:

    python link_toplayici.py 123

Nasıl çalışır: ders listesi sayfasındaki "i" (detay) butonu ASP.NET
`__doPostBack` ile çalışır; ders ID'si sayfada açıkça yazmaz. Araç,
tarayıcının yaptığını yapar — her "i" butonu için sunucuya bir postback
gönderir, dönen cevaptaki `progCourseDetails.aspx?curCourse=...` adresini
yakalar ve `links.txt`e yazar.

### Tam akış (iki komut)

    python link_toplayici.py 123      # 1) links.txt otomatik oluşur
    python bologna_scraper.py         # 2) Dersler_dolu.xlsx oluşur


## Uygulanan kurallar

| Sütun | Nasıl dolduruluyor |
|---|---|
| YÖKSİS Birim ID |  |
| Dersin Kodu / Adı | Sayfadan; ad Türkçe BÜYÜK harfe çevrilir |
| Okutulduğu Sınıf | Yarıyıldan: 1-2->1. SINIF, 3-4->2. SINIF, 5-6->3. SINIF ... |
| Okutulduğu Dönem | Tek yarıyıl->1. Dönem, çift->2. Dönem |
| Türü | Sayfadaki "Dersin Türü" (Zorunlu/Seçmeli) |
| Haftalık Ders Saati | T+U+L toplanır, tek sayı (3+1+0 -> 4) |
| AKTS | Sayfadan |
| **Dersi Alan Öğrenci Sayısı** | **2025 Güz/Bahar Excel'lerinden ders koduna göre** |
| Ortalama Derse Devam Yüzdesi | Sabit 75 |
| Lab Uygulaması Var Mı | Ortadaki sayı (U) > 0 ise Evet, değilse Hayır |
| Ders İşletmede Mi | "Okulda" (İşletmede Mesleki Eğitim/Staj ise "İşletmede") |
| **Dersin İlk Açıldığı Yıl** | **1. sınıf->2012, 2->2013, 3->2014, 4->2015** |
| Son Güncellenme Tarihi | Sayfadaki tarih, İngiliz biçimi (07/02/2025) |
| Dersin Amacı ve İçeriği | Amaç + boş satır + İçerik |
| Ders Bilgi Paketi Adresi | Dersin kendi linki |

### Öğrenci sayısı hakkında

Ders kodu, 2025 Güz ve Bahar dosyalarında aranır. Güz dosyasında bir dersin
birden fazla şubesi varsa şube öğrencileri ("Ara Toplam") toplanır; Bahar
dosyasında doğrudan "Toplam" değeri alınır.

Bir ders bu dosyalarda **bulunamazsa** (örn. İngilizce, Türk Dili gibi ortak
servis dersleri) varsayılan **80** yazılır ve çalışma sonunda hangi derslerin
bulunamadığı listelenir — bunları elle düzeltebilirsiniz.

### İlk açılış yılı hakkında

Sınıf bilgisinden otomatik üretilir. Son yıllarda açılmış istisna derslerini
çıktı Excel'inde elle düzeltmeniz gerekir.

## Gözden geçirilmesi gereken sütunlar

**Ders Teması** ve **Kullanılan Eğitim Platformları** ders adına bakan otomatik
tahminle dolar (Bilgisayar Mühendisliği'ne ayarlı). Çoğu derste doğrudur ama
kontrol etmeniz önerilir. Kuralları `bologna_scraper.py` başındaki
`TEMA_KURALLARI` listesinden değiştirebilirsiniz.

Sabitler (devam yüzdesi, varsayılan öğrenci sayısı, ilk yıl tabanı) yine
dosyanın başındaki AYARLAR bölümünden değiştirilebilir.
