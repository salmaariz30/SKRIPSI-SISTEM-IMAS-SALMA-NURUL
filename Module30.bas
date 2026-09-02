Attribute VB_Name = "Module30"
Sub IsiTanggalHariIni_PenghasilanLain()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI PENGHASILAN LAIN
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_PENGHASILAN LAIN")
    
    ' Kunci visual biar pergerakan kursornya super smooth
    Application.ScreenUpdating = False
    
    ' 1. AKTIFKAN SHEET TERLEBIH DAHULU (Wajib karena ada perintah .Select di akhir)
    LembarForm.Activate
    
    ' 2. EKSEKUSI: Bersihkan area merged D11 dan isi tanggal hari ini sesuai laptop
    With LembarForm.Range("D11").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Pindahkan kursor secara otomatis ke cell F11 setelah isi tanggal
    LembarForm.Range("F11").Select
    
    Application.ScreenUpdating = True
End Sub
