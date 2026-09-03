Attribute VB_Name = "Module99"
Sub IsiTanggalPelepasanAset()
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("ASET TETAP_INPUT PELEPASAN")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("D10").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti
    LembarForm.Range("F10").Select
    
End Sub
