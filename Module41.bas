Attribute VB_Name = "Module41"
Sub InputPengeluaranUsaha_NamaPelanggan()
    ' ==========================================================
    ' MODUL AUTO-WRITE KATA "Umum" - VERSI PENGELUARAN USAHA
    ' Target Tembak: Sheet "PENGELUARAN USAHA_INPUT DATA" -> G16 (Merged)
    ' Target Lompat: Sel D20 (Merged)
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT DATA")
    
    ' Supaya pergerakan layar dan kursor berjalan super smooth anti kedip di laptop ASUS
    Application.ScreenUpdating = False
    
    ' 1. EKSEKUSI: Aktifkan sheet agar proses Select kursor berjalan mulus
    LembarForm.Activate
    
    ' 2. SUNTIK TEKS: Bersihkan area merged G16 lalu ketik "Umum"
    With LembarForm.Range("G16").MergeArea
        .ClearContents
        .Value = "Umum"
    End With
    
    ' 3. UX FLOW: Lemparkan kursor secara otomatis ke D20 setelah ngetik
    ' Menggunakan .Cells(1, 1) agar Excel mendarat tepat di hulu merged cell D20
    LembarForm.Range("D20").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End Sub
