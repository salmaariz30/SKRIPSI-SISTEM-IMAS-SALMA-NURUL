Attribute VB_Name = "Module74"
Sub IsiTanggalPelunasanUtang()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI EFEKTIF
    ' Langsung eksekusi tanpa peduli kursor sedang di mana
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT PELUNASAN")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("D15").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti (H13)
    ' Agar pengguna bisa langsung melanjutkan pengetikan dokumen
    LembarForm.Range("G15").Select
    
End Sub
