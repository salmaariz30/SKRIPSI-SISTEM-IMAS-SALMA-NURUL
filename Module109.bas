Attribute VB_Name = "Module109"
Sub BuatDropdownUtang()
    Dim wsUtangBank As Worksheet, wsUtangUsaha As Worksheet, wsTarget As Worksheet
    Dim cell As Range, listBukti As String
    
    ' Pelindung Ekstra: Jika ada error sistem tak terduga, langsung CLOSE/HENTIKAN KODE
    On Error GoTo TanganiError
    
    ' 1. ATUR NAMA SHEET-NYA DI SINI
    Set wsTarget = Sheets("UTANG_INPUT PELUNASAN")
    
    listBukti = ""
    
    ' 2. Ambil data dari TabelUtangBank
    On Error Resume Next
    For Each cell In Range("TabelUtangBank[NO. BUKTI TRANSAKSI]")
        If Trim(cell.Value) <> "" Then
            listBukti = listBukti & Trim(cell.Value) & ","
        End If
    Next cell
    
    ' 3. Ambil data dari TabelUtangUsaha
    For Each cell In Range("TabelUtangUsaha[NO. BUKTI TRANSAKSI]")
        If Trim(cell.Value) <> "" Then
            listBukti = listBukti & Trim(cell.Value) & ","
        End If
    Next cell
    On Error GoTo TanganiError
    
    ' Hilangkan koma terakhir jika ada data
    If Len(listBukti) > 0 Then
        listBukti = Left(listBukti, Len(listBukti) - 1)
    Else
        ' Jika listBukti KOSONG, LANGSUNG CLOSE KODE (tidak perlu diproses ke Validation)
        Exit Sub
    End If
    
    ' PENGAMAN UTAMA: Cek batas maksimal 255 karakter Excel Data Validation String
    If Len(listBukti) > 255 Then
        MsgBox "Gagal membuat dropdown! Jumlah karakter daftar Bukti Transaksi melebihi batas maksimum Excel (255 Karakter).", _
               vbCritical, "Karakter Overload"
        Exit Sub
    End If
    
    ' 4. TEMBAK LANGSUNG JADI DROPDOWN DI SEL TARGET
    With wsTarget.Range("D11").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:=listBukti
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    Exit Sub

TanganiError:
    ' Jika terjadi error di titik manapun, LANGSUNG CLOSE / HENTIKAN KODE
    Exit Sub
End Sub

