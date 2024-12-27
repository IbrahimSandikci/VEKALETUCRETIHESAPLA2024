# VEKALETUCRETIHESAPLA2024 Formülü

<img src="/images/Kapak Vekalet Ücreti.png" alt="Vekalet Ücreti Tablosu" width="100%">

Bu VBA fonksiyonu, verilen asıl alacak ve mahkeme türüne göre vekalet ücretini hesaplar. Hesaplama, resmi ücret tarifesine göre yapılır. Şuan için sadece Asliye ve Tüketici Mahkemeleri için kullanılabilir.

## Fonksiyon Söz Dizimi
```excel
=VEKALETUCRETHESAPLA2024(AsilAlacak, MahkemeTuru)
```

## Parametreler
- **AsilAlacak** (Double): Vekalet ücretinin hesaplanacağı asıl alacak tutarı. Örnek: `900000`.
- **MahkemeTuru** (String): Mahkeme türü. "Asliye" veya "Tüketici" değerlerini alabilir (büyük/küçük harf duyarlı değildir).

## Özellikler
- Resmi ücret tarifesine göre kademeli oranlarla vekalet ücreti hesaplar.
- Hesaplanan vekalet ücreti hiçbir zaman asıl alacak tutarını aşamaz.
- "Asliye" mahkemeleri için minimum vekalet ücreti 30.000 TL olarak uygulanır.
- "Tüketici" mahkemeleri için minimum vekalet ücreti 15.000 TL olarak uygulanır.
- Geçersiz bir mahkeme türü girildiğinde hata mesajı döner.

<img src="/images/Icerik.png" alt="Vekalet Ücreti Tablosu">

## Örnekler

### Örnek 1: Geçerli Girdi
```excel
=VEKALETUCRETHESAPLA2024(900000, "Tüketici")
```
Sonuç: `138000`

 Örnek 2: Geçersiz Mahkeme Türü
```excel
=VEKALETUCRETHESAPLA2024(900000, "GeçersizTur")
```
Sonuç: `"Geçersiz mahkeme türü. Lütfen 'Asliye' veya 'Tüketici' olarak girin."`

<img src="/images/ExcelGoruntu2.png" alt="Vekalet Ücreti Tablosu">

## Nasıl Kullanılır?

### Aşağıdaki adımlardan herhangi biri ile belirtilen formülü excelde kullanmaya başlayabilirsiniz.

1. Excelde Dosya > Seçenekler > Yönet kısmında Excel Eklentileri seçili iken Git butonuna tıklayarak, açılan pencerede "Gözat" butonuna tıkladıktan sonra "Dava Vekalet Ücret Örnek Excel.xlam" dosyasını seçerek eklenti olarak excele dahil edebilir, (Bu adım ile tüm excel dosyalarınızda formülü kullanabilmenizi sağlayacaktır)

2. "Dava Vekalet Ücret Örnek Excel.xlsm" dosyasını indirerek "Makro devre dışı bırıkıldı uyarısında yer alan İçeriği Etkinleştirerek bu dosya içerisinde kullanmaya başlayabilir,

3. Depoda bulunan "DavaVekaletUcretHesaplama.bas" dosyasını Geliştirici module olarak Geliştirici > Visual Basic alanına ekleyebilir,

4. "DavaVekaletUcretFormul.txt" dosyasında yer alan kodu oluşturduğunuz module ekleyebilirsiniz.


## Daha fazla formül için takipte kalabilirsiniz 

## Yazar
**AV. İBRAHİM SANDIKCI**

