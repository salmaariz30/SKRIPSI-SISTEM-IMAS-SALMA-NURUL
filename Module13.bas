Attribute VB_Name = "Module13"
Sub HapusFormInputAsetTetap()
    Dim LembarForm As Worksheet
    
    ' 1. Set Target Sheet
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("INPUT DATA NEW")
    On Error GoTo 0
    
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'INPUT DATA NEW' tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. Konfirmasi User
    If MsgBox("Apakah Anda yakin ingin mengosongkan data formulir aset tetap ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset") = vbNo Then
        Exit Sub
    End If
    
    ' 3. Proses Pembersihan Kilat (Mendukung Cell Biasa maupun Merger)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
   With LembarForm
        On Error Resume Next
        
        ' Tembak langsung ke sel pojok kiri atas masing-masing merger area
        .Range("C12").Cells(1, 1).MergeArea.ClearContents
        .Range("C17").Cells(1, 1).MergeArea.ClearContents
        .Range("C22").Cells(1, 1).MergeArea.ClearContents
        .Range("C27").Cells(1, 1).MergeArea.ClearContents
        
        .Range("F12").Cells(1, 1).MergeArea.ClearContents
        .Range("F22").Cells(1, 1).MergeArea.ClearContents
        .Range("F27").Cells(1, 1).MergeArea.ClearContents
        
        .Range("I12").Cells(1, 1).MergeArea.ClearContents
        .Range("I22").Cells(1, 1).MergeArea.ClearContents
        
        On Error GoTo 0
    End With
    
    ' 4. Kembalikan Kursor ke Cell Pertama (C12)
    LembarForm.Activate
    LembarForm.Range("C12").Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Formulir Aset Tetap Berhasil Dihapus!", vbInformation, "Sukses"
End Sub

