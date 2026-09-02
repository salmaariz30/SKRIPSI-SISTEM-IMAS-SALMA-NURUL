Attribute VB_Name = "Module46"
Sub InputProduksi_NamaVendorUmum()
    ' ==========================================================
    ' MODUL AUTO-WRITE KATA "Umum" - VERSI INPUT PRODUKSI
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN_INPUT PRODUKSI")
    
    ' Supaya pergerakan layar dan kursor berjalan super smooth anti kedip di laptop ASUS
    Application.ScreenUpdating = False
    
    ' 1. EKSEKUSI: Aktifkan sheet agar proses Select kursor berjalan mulus
    LembarForm.Activate
    
    ' 2. SUNTIK TEKS: Bersihkan area merged G19 lalu ketik "Umum"
    With LembarForm.Range("G19").MergeArea
        .ClearContents
        .Value = "Umum"
    End With
    
    ' 3. UX FLOW: Lemparkan kursor secara otomatis ke D23 setelah ngetik
    ' Menggunakan .Cells(1, 1) agar Excel mendarat tepat di hulu merged cell D23
    LembarForm.Range("D23").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End Sub
