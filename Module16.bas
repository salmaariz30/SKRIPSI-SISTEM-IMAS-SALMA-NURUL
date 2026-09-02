Attribute VB_Name = "Module16"
Sub IsiTanggalHariIni()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI EFEKTIF
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("KAS&BANK_INPUT TRANSAKSI UMUM")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("E13").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti (H13)
    ' Agar pengguna bisa langsung melanjutkan pengetikan dokumen
    LembarForm.Range("H13").Select
    
End Sub
