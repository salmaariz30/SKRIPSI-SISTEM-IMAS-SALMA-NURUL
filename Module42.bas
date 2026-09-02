Attribute VB_Name = "Module42"
Sub IsiTanggalHariIni_PengeluaranUsaha()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI PENGELUARAN USAHA
    ' Target Tembak: Sheet "PENGELUARAN USAHA_INPUT DATA" -> D16 (Merged)
    ' Target Lompat: Sel G16
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT DATA")
    
    ' Kunci visual biar pergerakan kursornya super smooth di laptop
    Application.ScreenUpdating = False
    
    ' 1. AKTIFKAN SHEET TERLEBIH DAHULU (Wajib karena ada perintah .Select di akhir)
    LembarForm.Activate
    
    ' 2. EKSEKUSI: Bersihkan area merged D16 dan isi tanggal hari ini
    With LembarForm.Range("D16").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Pindahkan kursor secara otomatis ke cell G16 setelah isi tanggal
    ' Menggunakan .MergeArea.Cells(1, 1) agar aman jika G16 ternyata juga berupa merged cell
    LembarForm.Range("G16").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End Sub
