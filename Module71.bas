Attribute VB_Name = "Module71"
Sub IsiTanggalUtangUsaha()
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT USAHA")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("D12").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti (H13)
    ' Agar pengguna bisa langsung melanjutkan pengetikan dokumen
    LembarForm.Range("H12").Select
    
End Sub
