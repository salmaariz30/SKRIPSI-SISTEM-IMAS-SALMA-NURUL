Attribute VB_Name = "Module35"
Sub IsiTanggalHariIni_InputRetur()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI INPUT RETUR
    ' Langsung eksekusi tanpa peduli kursor sedang di mana
    ' Target Sel: E11 (Merged) -> Lompat otomatis ke E13 (Merged)
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT RETUR")
    
    ' Kunci visual biar pergerakan kursornya super smooth anti kedip
    Application.ScreenUpdating = False
    
    ' 1. AKTIFKAN SHEET TERLEBIH DAHULU (Wajib karena ada perintah .Select di akhir)
    LembarForm.Activate
    
    ' 2. EKSEKUSI: Bersihkan area merged E11 dan isi tanggal hari ini
    With LembarForm.Range("E11").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Pindahkan kursor secara otomatis ke cell E13 setelah isi tanggal
    ' Menggunakan .Cells(1, 1) agar mendarat tepat di hulu merged cell E13 Excel USA
    LembarForm.Range("E13").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End Sub
