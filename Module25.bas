Attribute VB_Name = "Module25"
Sub IsiTanggalHariIni_InputPenjualan()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI INPUT PENJUALAN
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    
    ' 1. EKSEKUSI: Aktifkan sheet dulu agar perintah .Select di akhir tidak eror
    LembarForm.Activate
    
    ' 2. SUNTIK TANGGAL: Bersihkan area merged C11 dan isi tanggal hari ini
    With LembarForm.Range("C11").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 3. UX FLOW: Pindahkan kursor secara otomatis ke I11 setelah isi tanggal
    LembarForm.Range("I11").Select
    
End Sub
