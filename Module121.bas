Attribute VB_Name = "Module121"
Sub CariSheetPivotTerkunci()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pesan As String
    
    pesan = "Hasil Pelacakan Sheet Berpivottable & Terkunci:" & vbCrLf
    
    For Each ws In ActiveWorkbook.Worksheets
        ' Cek jika sheet memiliki PivotTable dan statusnya terproteksi
        If ws.PivotTables.Count > 0 And ws.ProtectContents = True Then
            pesan = pesan & "- Sheet: [" & ws.Name & "] ?? TERKUNCI!" & vbCrLf
        End If
    Next ws
    
    If pesan = "Hasil Pelacakan Sheet Berpivottable & Terkunci:" & vbCrLf Then
        MsgBox "Tidak ditemukan sheet ber-Pivot yang terkunci. Cek sheet tersembunyi (Hidden).", vbInformation
    Else
        MsgBox pesan, vbWarning, "Laporan Sistem IMAS"
    End If
End Sub
