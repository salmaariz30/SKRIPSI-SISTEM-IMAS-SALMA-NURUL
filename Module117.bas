Attribute VB_Name = "Module117"
Sub ProsesLogout()
    Dim ws As Worksheet
    Dim sheetLogin As String
    
    ' --- PENGATURAN ---
    sheetLogin = "LOGIN"
    ' ------------------
    
    ' Konfirmasi ke user dulu biar gak sengaja keklik
    If MsgBox("Apakah Anda yakin ingin keluar dan mengunci sistem?", vbQuestion + vbYesNo, "Konfirmasi Keluar") = vbNo Then
        Exit Sub
    End If
    
    ' == JURUS BIAR SMOOTH & TIDAK BERKEDIP KASAR ==
    Application.ScreenUpdating = False
    
    ' 1. Munculkan sheet Login terlebih dahulu sebagai jangkar fokus
    ThisWorkbook.Sheets(sheetLogin).Visible = xlSheetVisible
    ThisWorkbook.Sheets(sheetLogin).Activate
    
    ' 2. Sembunyikan semua sheet lainnya secara total (VeryHidden)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> sheetLogin Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' 3. Pastikan TextBox di halaman login langsung bersih dan siap menerima ketikan baru
    On Error Resume Next
    ThisWorkbook.Sheets(sheetLogin).OLEObjects("TextBoxLogin").Object.Text = ""
    ThisWorkbook.Sheets(sheetLogin).OLEObjects("TextBoxLogin").Select
    On Error GoTo 0
    
    ' == NYALAKAN KEMBALI LAYAR ==
    Application.ScreenUpdating = True
    
End Sub
