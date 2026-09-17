# 📊 Excel & CSV Veri Temizleme, Özet Çıkarma ve İndirme Arayüzü

Python, Streamlit, pandas ve openpyxl kullanılarak geliştirilmiş modern ve kapsamlı veri temizleme ve özetleme aracı.

---

## 🚀 Başlatma (Nasıl Çalıştırılır?)

1. Terminal / Komut Satırında proje klasörüne gidin:
   ```bash
   cd "C:\Users\GAMENZİA\.gemini\antigravity\scratch\data-cleaner-app"
   ```

2. Bağımlılıkları yükleyin (zaten yüklendiyse bu adımı atlayabilirsiniz):
   ```bash
   pip install -r requirements.txt
   ```

3. Streamlit uygulamasını başlatın:
   ```bash
   streamlit run app.py
   ```
   Uygulama otomatik olarak varsayılan web tarayıcınızda (`http://localhost:8501`) açılacaktır.

---

## ✨ Özellikler

### 1. 📂 Esnek Dosya Yükleme
- **Excel Desteği (.xlsx, .xls)**: Çok sayfalı (multi-sheet) dosyalarda sayfa seçimi.
- **CSV Desteği (.csv)**: Ayraç (virgül, noktalı virgül, tab, dikey çizgi) ve karakter kodlaması (utf-8, latin5, cp1254 vb.) seçenekleri.

### 2. 📈 Veri Özeti & Profilleme (Summary)
- Toplam Satır, Sütun, Eksik Hücre Sayısı/Oranı, Yinelenen Satır Sayısı ve Bellek Kullanımı metrik kartları.
- İlk 10, son 10 veya rastgele satır önizleme.
- Her sütun için veri tipi, eksiklik oranı, benzersiz değer sayısı tablosu.
- Eksik değerlerin sütun bazlı çubuk grafik görselleştirmesi.
- Sayısal ve tüm sütunlar için ayrıntılı istatistiksel özet (`describe`).

### 3. 🧹 Veri Temizleme Araçları (Cleaning)
- **Yinelenen Kayıtlar**: Tüm sütunlara veya seçilen kriter sütunlarına göre tekrarlanan satırları silme (ilk veya son kaydı koruma).
- **Eksik Değerler (NaN/Null)**:
  - Satır silme (tamamı boşsa, herhangi biri boşsa veya belirli sütunlar boşsa).
  - Doldurma (Ortalama, Medyan, Mod, Sabit Değer, İleri/Geri Doldurma).
- **Sütun Yönetimi**:
  - İstenmeyen sütunları kaldırma.
  - Sütun isimlerini otomatik standartlaştırma (küçük harf, boşlukları `_` yapma, Türkçe karakter uyarlama).
  - Tek tek sütun yeniden adlandırma.
- **Veri Tipi Dönüştürme**: Sayısal, Tam Sayı, Metin, Tarih/Saat (Datetime) ve Kategori tipleri arasında hızlı dönüşüm.
- **Metin Temizleme**: Baş/son boşlukları kırpma (`strip`), küçük harf, BÜYÜK HARF ve Baş Harfleri Büyütme dönüşümleri.
- **Aykırı Değer Filtresi**: Sayısal sütunlarda IQR (Çeyrekler Açıklığı) veya özel Min/Max eşikleri ile aykırı satırları temizleme.
- **İşlem Güvenliği**: İstenildiği an "Orijinal Veriye Sıfırla" butonu ile ilk yüklenen veriye geri dönebilme.

### 4. 💾 Dışa Aktarma & İndirme (Export)
- **Excel (.xlsx)**: `openpyxl` motoru ile stillendirilmiş mavi başlık satırı ve otomatik sütun genişliği optimizasyonu ile indirme.
- **CSV (.csv)**: Excel ile açıldığında Türkçe karakter bozulmasını engelleyen `utf-8-sig` formatı ve özel ayraç seçeneği ile indirme.
- **İşlem Günlüğü (Audit Log)**: Uygulanan tüm temizleme adımlarının zaman damgalı dökümü.

---

## 🧪 Test Verileri
Proje klasöründe denemeler yapabilmeniz için hazır test dosyaları yer almaktadır:
- `sample_data.csv`
- `sample_data.xlsx`
