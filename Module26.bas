Attribute VB_Name = "Module26"
Sub IsiTanggalHariIni_InputTanggalPengerjaan()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI INPUT PENJUALAN (PART 2)
    ' Mengisi tanggal di I19 dan otomatis melempar kursor ke C19
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    
    ' 1. EKSEKUSI: Aktifkan sheet agar proses Select kursor aman jaya
    LembarForm.Activate
    
    ' 2. SUNTIK TANGGAL: Bersihkan area merged I19 dan isi tanggal hari ini
    With LembarForm.Range("I19").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Lemparkan kursor secara otomatis ke C19 setelah isi tanggal
    LembarForm.Range("C19").Select
    
End Sub
