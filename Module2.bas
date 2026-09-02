Attribute VB_Name = "Module2"
Sub IsiTanggalUtangBank()
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT BANK")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged dan isi tanggal hari ini
    With LembarForm.Range("D12").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis ke No. Bukti
    LembarForm.Range("H12").Select
    
End Sub

