' Modül içinde kullanacağımız grafik silme işlemi
Private Sub DeleteChartByName(ws As Worksheet, chartName As String)
    Dim ch As ChartObject
    On Error Resume Next
    For Each ch In ws.ChartObjects
        If ch.Name = chartName Then
            ch.Delete
            Exit For
        End If
    Next ch
    On Error GoTo 0
End Sub

' Temizle Butonu
Private Sub btnTemizle_Click()
    txtX.Text = ""
    txtY.Text = ""
    txtKorelasyonSonuc.Text = ""
    txtRegresyonSonuc.Text = ""
    txtcıktı.Text = ""
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sayfa1")
    ws.Range("A1:B100").ClearContents
    
    DeleteChartByName ws, "RegresyonGrafik"
    DeleteChartByName ws, "KorelasyonGrafik"
End Sub

' Regresyon Hesaplama Butonu
Private Sub btnRegresyon_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sayfa1")
    
    Dim xArr() As String, yArr() As String
    Dim i As Integer, n As Integer
    Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
    Dim a As Double, b As Double
    
    ' Temizleme işlemleri
    ws.Range("A1:B100").ClearContents
    DeleteChartByName ws, "RegresyonGrafik"
    
    ' Girdi verilerini al
    xArr = Split(Me.txtX.Value, ",")
    yArr = Split(Me.txtY.Value, ",")
    
    ' Hata kontrolü
    If UBound(xArr) <> UBound(yArr) Then
        MsgBox "X ve Y değerlerinin sayısı eşit olmalı!", vbCritical
        Exit Sub
    End If
    
    n = UBound(xArr) + 1
    
    ' Sayfaya veri aktarımı ve toplamların hesaplanması
    For i = 0 To UBound(xArr)
        Dim xi As Double, yi As Double
        xi = Val(Trim(xArr(i)))
        yi = Val(Trim(yArr(i)))
        
        ws.Cells(i + 1, 1).Value = xi
        ws.Cells(i + 1, 2).Value = yi
        
        sumX = sumX + xi
        sumY = sumY + yi
        sumXY = sumXY + xi * yi
        sumX2 = sumX2 + xi ^ 2
    Next i
    
    ' Regresyon katsayıları
    b = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX ^ 2)
    a = (sumY - b * sumX) / n
    
    Me.txtRegresyonSonuc.Value = "y = " & Format(a, "0.00") & " + " & Format(b, "0.00") & "x"
    
    ' Yorumlama
    Select Case True
        Case b > 0
            Me.txtcıktı.Value = "X arttıkça Y artıyor. Pozitif yönlü ilişki."
        Case b < 0
            Me.txtcıktı.Value = "X arttıkça Y azalıyor. Negatif yönlü ilişki."
        Case Else
            Me.txtcıktı.Value = "X değişse bile Y sabit. İlişki yok."
    End Select
    
    ' Grafik çizimi (grafiği 200 sağa kaydırdım - Left:=350)
    Dim grafik As ChartObject
    Set grafik = ws.ChartObjects.Add(Left:=350, Width:=400, Top:=10, Height:=300)
    grafik.Name = "RegresyonGrafik"
    
    With grafik.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Regresyon Grafiği"
        
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "X Değerleri"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Y Değerleri"
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Veri Noktaları"
            .xValues = ws.Range("A1:A" & n)
            .Values = ws.Range("B1:B" & n)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6
            .MarkerForegroundColor = RGB(0, 0, 255)
            .MarkerBackgroundColor = RGB(0, 0, 255)
        End With
        
        ' Trendline (doğrusal regresyon çizgisi)
        With .SeriesCollection(1).Trendlines.Add(Type:=xlLinear)
            .DisplayEquation = True
            .DisplayRSquared = True
        End With
    End With
End Sub

' Korelasyon Hesaplama Butonu
Private Sub btnKorelasyon_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sayfa1")
    
    Dim xValues() As String, yValues() As String
    Dim xArr() As Double, yArr() As Double
    Dim avgX As Double, avgY As Double
    Dim sumXY As Double, sumX2 As Double, sumY2 As Double
    Dim korelasyon As Double
    Dim i As Integer, n As Integer
    
    ' Temizleme
    ws.Range("A1:B100").ClearContents
    DeleteChartByName ws, "KorelasyonGrafik"
    
    ' Verileri al
    xValues = Split(txtX.Text, ",")
    yValues = Split(txtY.Text, ",")
    
    If UBound(xValues) <> UBound(yValues) Then
        MsgBox "X ve Y değerleri eşit sayıda olmalıdır.", vbExclamation
        Exit Sub
    End If
    
    n = UBound(xValues) + 1
    ReDim xArr(1 To n)
    ReDim yArr(1 To n)
    
    For i = 1 To n
        xArr(i) = CDbl(Trim(xValues(i - 1)))
        yArr(i) = CDbl(Trim(yValues(i - 1)))
        
        avgX = avgX + xArr(i)
        avgY = avgY + yArr(i)
        
        ws.Cells(i, 1).Value = xArr(i)
        ws.Cells(i, 2).Value = yArr(i)
    Next i
    
    avgX = avgX / n
    avgY = avgY / n
    
    For i = 1 To n
        sumXY = sumXY + (xArr(i) - avgX) * (yArr(i) - avgY)
        sumX2 = sumX2 + (xArr(i) - avgX) ^ 2
        sumY2 = sumY2 + (yArr(i) - avgY) ^ 2
    Next i
    
    korelasyon = sumXY / (Sqr(sumX2) * Sqr(sumY2))
    Me.txtKorelasyonSonuc.Text = Format(korelasyon, "0.000")
    
    ' Yorum
    Select Case True
        Case korelasyon > 0
            Me.txtcıktı.Text = "Pozitif korelasyon: " & Format(korelasyon, "0.000")
        Case korelasyon < 0
            Me.txtcıktı.Text = "Negatif korelasyon: " & Format(korelasyon, "0.000")
        Case Else
            Me.txtcıktı.Text = "Korelasyon yok (0)."
    End Select
    
    ' Grafik çizimi (grafiği 200 sağa kaydırdım - Left:=350)
    Dim grafik As ChartObject
    Set grafik = ws.ChartObjects.Add(Left:=350, Width:=400, Top:=10, Height:=300)
    grafik.Name = "KorelasyonGrafik"
    
    With grafik.Chart
        .ChartType = xlXYScatter
        .HasTitle = True
        .ChartTitle.Text = "Korelasyon Grafiği"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "X Değerleri"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Y Değerleri"
        
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .Name = "Veri Noktaları"
            .xValues = ws.Range("A1:A" & n)
            .Values = ws.Range("B1:B" & n)
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6
            .MarkerForegroundColor = RGB(255, 0, 0)
            .MarkerBackgroundColor = RGB(255, 0, 0)
        End With
    End With
End Sub

Private Sub UserForm_Click()

End Sub

