Attribute VB_Name = "Module77"
Sub IsiTanggalPembelianPersediaan()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI EFEKTIF
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PERSEDIAAN_INPUT BARANG")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("H12").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti (H13)
    ' Agar pengguna bisa langsung melanjutkan pengetikan dokumen
    LembarForm.Range("D16").Select
    
End Sub
