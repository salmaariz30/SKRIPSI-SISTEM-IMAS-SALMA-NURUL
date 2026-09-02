Attribute VB_Name = "Module45"
Sub IsiTanggalHariIni_InputProduksi()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI INPUT PRODUKSI
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN_INPUT PRODUKSI")
    
    ' Kunci visual biar pergerakan kursornya super smooth di laptop ASUS
    Application.ScreenUpdating = False
    
    ' 1. AKTIFKAN SHEET TERLEBIH DAHULU (Wajib karena ada perintah .Select di akhir)
    LembarForm.Activate
    
    ' 2. EKSEKUSI: Bersihkan area merged D15 dan isi tanggal hari ini
    With LembarForm.Range("D15").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Pindahkan kursor secara otomatis ke cell G15 setelah isi tanggal
    LembarForm.Range("G15").Select
    
    Application.ScreenUpdating = True
End Sub
