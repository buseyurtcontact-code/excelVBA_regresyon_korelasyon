# 📊 Excel VBA Regression and Correlation Analysis Tool (English)

Course Project / Interactive Data Analysis Utility  
Student: Buse Yurt · Ege University, Faculty of Science – Department of Statistics  
Tool: Microsoft Excel VBA (Visual Basic for Applications)  

---

## 📌 Project Overview
This project is an automated statistical tool developed using **Excel VBA** to compute simple linear regression and Pearson correlation analysis via an interactive UserForm.

By simply inputting comma-separated **X and Y values**, users can automatically output key statistical coefficients, interpretational insights, and dynamic Scatter Plots with a single click.

---

## ⚡ Main Features

### 📉 Linear Regression Analysis
- Calculates regression slope ($b$) and intercept ($a$) dynamically ($y = a + bx$).
- Automatically generates an `xlXYScatter` Plot featuring an integrated linear trendline.
- Displays calculated formulas, $R^2$ values, and real-time directional trend interpretations.

### 🔗 Correlation Analysis
- Computes Pearson's correlation coefficient ($r$) using mean-deviation covariance logic.
- Generates dedicated visual scatter plots for variable pair inspection.
- Provides immediate contextual interpretations (positive, negative, or no correlation).

### 🧹 Automated Workspace Cleaning & Data Handling
- Clears pre-existing worksheet ranges (`A1:B100`) and removes stale chart objects upon recalculation.
- Implements custom chart object cleanup functions to prevent workspace clutter.
- Includes input length validation logic to ensure matched $X$ and $Y$ dimensions.

---

## 🚀 How It Works
1. Open the Excel workbook and launch the macro-enabled UserForm.
2. Enter comma-separated values into the $X$ and $Y$ input fields (e.g., `10, 20, 30`).
3. Click the desired action button:
   - **Calculate Regression** $\rightarrow$ Outputs model parameters, scatter plot with trendline, and trend analysis.
   - **Calculate Correlation** $\rightarrow$ Outputs Pearson's $r$, scatter plot, and correlation breakdown.
   - **Clear Form** $\rightarrow$ Resets input boxes, clears worksheet cells, and deletes generated charts.

---

## 💡 Key Advantages
- **Time Efficiency:** Instant computation and dynamic visual plotting eliminate repetitive manual tasks.
- **Error Reduction:** Automated array parsing and cell mapping minimize human calculation errors.
- **User-Friendly Interface:** Clean VBA UserForm design allows effortless statistical evaluation.
- **Academic & Professional Utility:** Ideal for rapid exploratory data checks, statistics coursework, and practical data analysis.

---

## 🛠 Tech Stack
- **Excel VBA** (UserForms, Modules, Dynamic Range Allocation)
- **Excel Charting Engine** (Programmatic `ChartObject` Creation & Trendline Customization)

# Excel VBA ile Regresyon ve Korelasyon Analizi (Türkçe)

## Proje Özeti
Bu proje, **Excel VBA** kullanarak regresyon ve korelasyon analizlerini otomatikleştiren bir araçtır.  
Kullanıcı, Excel üzerinde **X ve Y değerlerini** girdikten sonra tek tıkla analiz sonuçlarını, grafiklerini ve kısa yorumlarını elde edebilir.  

## Özellikler
- **Regresyon Analizi**
  - X ve Y değerleri üzerinden regresyon katsayıları hesaplanır.
  - Regresyon grafiği otomatik çizilir.
  - Grafik altında kısa yorum eklenir.

- **Korelasyon Analizi**
  - X ve Y değerleri arasındaki korelasyon katsayısı hesaplanır.
  - Korelasyon grafiği otomatik çizilir.
  - Grafik altında kısa yorum eklenir.

## Kullanım Adımları
1. Excel dosyasını açın.  
2. X ve Y değerlerini ilgili hücrelere girin.  
3. İlgili butona tıklayın:  
   - **Regresyon Hesapla** → Regresyon grafiği + kısa yorum  
   - **Korelasyon Hesapla** → Korelasyon grafiği + kısa yorum  
4. Sonuçlar otomatik olarak ekrana yansır.  

## Avantajlar
- **Zaman Tasarrufu:** Grafikler ve yorumlar otomatik oluşturulur.  
- **Hata Azaltma:** Manuel hesaplama ve grafik çizim hataları ortadan kalkar.  
- **Kullanıcı Dostu:** Tek tıkla analiz ve görselleştirme.  
- **Akademik ve Pratik Kullanım:** Öğrenciler, araştırmacılar ve veri analistleri için ideal.  

## Örnek Çalışma Akışı
- X ve Y değerleri girilir.  
- "Regresyon Hesapla" butonuna basılır.  
- Çıktı: Regresyon grafiği + kısa yorum.  
- "Korelasyon Hesapla" butonuna basılır.  
- Çıktı: Korelasyon grafiği + kısa yorum.  

## Teknoloji
- **Excel VBA (Visual Basic for Applications)**  
- Excel grafik motoru  

---

 **Not:** Kodlar tamamen VBA ile yazılmıştır ve Excel üzerinde çalıştırılabilir.  
 Bu proje, istatistiksel analizleri daha hızlı ve görsel hale getirmek için geliştirilmiştir.
