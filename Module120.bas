Attribute VB_Name = "Module120"
Sub HapusBarisTabelAktif()
    Dim TargetCell As Range
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim ans As VbMsgBoxResult
    Dim barisKe As Long
    Const PWD As String = "IMAS" ' <-- Sandi pelindung sheet kamu
    
    Set TargetCell = ActiveCell
    Set ws = ActiveSheet
    
    On Error Resume Next
    Set tbl = TargetCell.ListObject
    On Error GoTo 0
    
    If tbl Is Nothing Then Exit Sub
    
    ' Hitung indeks baris di dalam tabel
    barisKe = TargetCell.Row - tbl.DataBodyRange.Row + 1
    
    ' Konfirmasi pengaman
    ans = MsgBox("Apakah Anda yakin ingin menghapus baris ke-" & barisKe & " pada '" & tbl.Name & "'?" & vbCrLf & _
                 "Tindakan ini tidak dapat dibatalkan!", _
                 vbQuestion + vbYesNo, "Konfirmasi Hapus Baris")
                 
    If ans = vbYes Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        ' Eksekusi gembok & hapus
        ws.Unprotect Password:=PWD
        tbl.ListRows(barisKe).Delete
        ws.Protect Password:=PWD, AllowFiltering:=True
        
        ThisWorkbook.RefreshAll
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox "Baris tabel berhasil dihapus!", vbInformation, "Sukses"
    End If
End Sub

