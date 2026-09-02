Attribute VB_Name = "Module66"
Sub IsiTanggalHariIni_InputDataPiutang()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI INPUT DATA PIUTANG
    ' Target Tembak: Sheet "PIUTANG_INPUT DATA PIUTANG" -> E17 (Merged)
    ' Target Lompat: Sel G17
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PIUTANG_INPUT DATA PIUTANG")
    
    ' Kunci visual biar pergerakan kursornya super smooth di laptop
    Application.ScreenUpdating = False
    
    ' 1. AKTIFKAN SHEET TERLEBIH DAHULU (Wajib karena ada perintah .Select di akhir)
    LembarForm.Activate
    
    ' 2. EKSEKUSI: Bersihkan area merged E17 dan isi tanggal hari ini
    With LembarForm.Range("E17").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Pindahkan kursor secara otomatis ke cell G17 setelah isi tanggal
    ' Menggunakan .MergeArea.Cells(1, 1) agar aman jika G17 ternyata juga berupa merged cell
    LembarForm.Range("G17").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End Sub
