Attribute VB_Name = "Module116"
Sub KunciFormatChartPendapatan()
    Dim ws As Worksheet
    Dim namaChart As Variant
    Dim i As Long
    Dim chtObj As ChartObject
    Dim ser As Series
    
    ' Target langsung ke sheet dashboard kas dan bank kamu
    Set ws = Sheets("PENDAPATAN_DASHBOARD")
    
    ' Buka proteksi sheet terlebih dahulu (jika di-protect, silakan sesuaikan password-nya)
    ws.Unprotect Password:="IMAS"
    
    ' Array untuk menembak dua nama chart spesifik kamu
    namaChart = Array("LinePivotPendapatan", "BarPivotPendapatan", "ChartBarPendapatan")
    
    On Error Resume Next
    For i = LBound(namaChart) To UBound(namaChart)
        Set chtObj = ws.ChartObjects(namaChart(i))
        
        If Not chtObj Is Nothing Then
            With chtObj.Chart
                ' Loop untuk mengatur semua series data/batang grafik di dalam chart
                For Each ser In .SeriesCollection
                    ' Memastikan Data Labels aktif dan muncul
                    ser.HasDataLabels = True
                    
                    With ser.DataLabels
                        ' 1. Font: Century Gothic, Ukuran 8, Bold
                        .Font.Name = "Century Gothic"
                        .Font.Size = 8
                        .Font.Bold = True
                        
                        ' 2. Fill Solid: Hitam Keabuan (RGB 64, 64, 64)
                        .Format.Fill.Solid
                        .Format.Fill.ForeColor.RGB = RGB(64, 64, 64)
                        
                        ' Tambahan otomatis: Mengubah warna teks menjadi putih agar kontras dengan background gelap
                        .Font.Color = RGB(255, 255, 255)
                    End With
                Next ser
                
                ' --- LOGIKA KHUSUS UNTUK LEGEND ---
                If namaChart(i) = "BarPivotPendapatan" Then
                    ' Pastikan legend-nya memang aktif/muncul terlebih dahulu
                    .HasLegend = True
                    
                    With .Legend
                        .Font.Name = "Century Gothic"
                        .Font.Size = 8
                        .Font.Bold = True
                        .Font.Color = RGB(0, 0, 0) ' Warna Hitam murni
                    End With
                End If
                
                ' Kunci format grafik agar tidak bisa digeser atau diubah manual oleh user
                .ProtectFormatting = True
            End With
        End If
        Set chtObj = Nothing
    Next i
    On Error GoTo 0
    
    ' Kunci kembali sheet dengan mengizinkan objek grafis/chart tetap bekerja di latar belakang
    ws.Protect Password:="IMAS", DrawingObjects:=False, Contents:=True, Scenarios:=True
End Sub

