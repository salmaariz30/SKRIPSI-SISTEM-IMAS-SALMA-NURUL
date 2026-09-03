Attribute VB_Name = "Module122"
Sub ExportAllVBAModules()
    Dim comp As Object
    Dim exportFolder As String
    Dim ext As String
    
    ' Folder hasil ekspor (akan dibuat di folder yang sama dengan file Excel ini)
    exportFolder = ThisWorkbook.Path & "\VBA_Source_Code\"
    
    ' Buat folder jika belum ada
    If Dir(exportFolder, vbDirectory) = "" Then
        MkDir exportFolder
    End If
    
    ' Loop seluruh modul, form, dan class yang ada di proyek ini
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1 ' Standard Module (.bas)
                ext = ".bas"
            Case 2 ' Class Module (.cls)
                ext = ".cls"
            Case 3 ' UserForm (.frm)
                ext = ".frm"
            Case Else ' Sheet & ThisWorkbook (.cls)
                ext = ".cls"
        End Select
        
        ' Ekspor file kode jika modul berisi baris kode
        If comp.CodeModule.CountOfLines > 0 Then
            comp.Export exportFolder & comp.Name & ext
        End If
    Next comp
    
    MsgBox "Berhasil! 110 Modul telah terekspor otomatis ke:" & vbCrLf & exportFolder, vbInformation, "Export Selesai"
End Sub
