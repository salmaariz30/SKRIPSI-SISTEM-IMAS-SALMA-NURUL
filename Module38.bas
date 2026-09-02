Attribute VB_Name = "Module38"
Sub HapusFormPenjualanTotal()
    ' ====================================================================
    ' MODUL UTAMA: RESET TOTAL FORMULIR INPUT PENJUALAN (ANTI-MERGER BUG)
    ' Terkunci Khusus untuk Sheet: PENDAPATAN_INPUT PENJUALAN
    ' Menghapus Bersamaan: Form Per Transaksi & Form Rekap Harian
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim i As Long
    
    ' Mengunci target pembersihan hanya pada sheet input penjualan
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah gaib
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'PENDAPATAN_INPUT PENJUALAN' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum hangus)
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Penjualan ini?" & vbNewLine & _
              "(Data Per Transaksi & Rekap Harian akan dihapus permanen!)", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Penjualan") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES PEMBERSIHAN DATA DENGAN TAKTIK LOOPING AMAN
    With LembarForm
        On Error Resume Next
        
        ' --- A. BUMI HANGUS FORM PER TRANSAKSI (Baris 24 s.d 38) ---
        For i = 24 To 38
            .Range("D" & i).MergeArea.ClearContents ' Kolom D (No Nota Merger)
            .Range("F" & i).ClearContents           ' Kolom F (Detail/Nama)
            .Range("J" & i).MergeArea.ClearContents ' Kolom J (Nominal Merger)
        Next i
        .Range("L41").ClearContents                 ' Grand Total L41
        
        ' --- B. BUMI HANGUS FORM REKAP HARIAN (Baris 54 s.d 83) ---
        For i = 54 To 83
            .Range("D" & i).MergeArea.ClearContents ' Kolom D (No Nota Merger)
            .Range("F" & i).ClearContents           ' Kolom F (Detail/Nama)
            .Range("J" & i).MergeArea.ClearContents ' Kolom J (Nominal Merger)
        Next i
        .Range("L86").ClearContents                 ' Grand Total L86
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke posisi awal dropdown utama F15)
    LembarForm.Activate
    LembarForm.Range("F15").Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Penjualan Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

