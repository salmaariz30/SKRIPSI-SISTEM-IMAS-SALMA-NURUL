Attribute VB_Name = "Module68"
Sub HapusFormInputPiutangTotal()
    ' ====================================================================
    ' MODUL UTAMA: RESET TOTAL FORMULIR INPUT DATA PIUTANG (ANTI-MERGER BUG)
    ' Terkunci Khusus untuk Sheet: PIUTANG_INPUT DATA PIUTANG
    ' Menghapus Bersamaan Seluruh Field Inputan Utama
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet input data piutang
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("PIUTANG_INPUT DATA PIUTANG")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah secara gaib
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'PIUTANG_INPUT DATA PIUTANG' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum hangus permanen)
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Input Piutang ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Piutang") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat tanpa kedip
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES PEMBERSIHAN DATA SESUAI TARGET KOORDINAT
    With LembarForm
        On Error Resume Next
        
        ' Menyapu bersih sel-sel inputan kolom E
        .Range("E13").MergeArea.ClearContents ' Target 1: E13
        .Range("E17").MergeArea.ClearContents ' Target 2: E17
        .Range("E21").MergeArea.ClearContents ' Target 3: E21
        .Range("E25").MergeArea.ClearContents ' Target 4: E25
        .Range("E29").MergeArea.ClearContents ' Target 5: E29
        
        ' Menyapu bersih sel-sel inputan kolom G
        .Range("G17").MergeArea.ClearContents ' Target 6: G17
        .Range("G21").MergeArea.ClearContents ' Target 7: G21
        .Range("G25").MergeArea.ClearContents ' Target 8: G25
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke posisi awal form E13)
    LembarForm.Activate
    LembarForm.Range("E13").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Input Piutang Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

