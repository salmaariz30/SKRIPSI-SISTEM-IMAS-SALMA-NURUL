Attribute VB_Name = "Module86"
Sub IsiTanggalDisposal()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI EFEKTIF
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PERSEDIAAN_INPUT DISPOSAL")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("D12").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti (H13)
    ' Agar pengguna bisa langsung melanjutkan pengetikan dokumen
    LembarForm.Range("F12").Select
    
End Sub
