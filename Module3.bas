Attribute VB_Name = "Module3"
Sub HapusFormInputUtangBankTotal()
    ' ====================================================================
    ' MODUL PEMBERSIH: RESET FORMULIR INPUT UTANG BANK SECARA AMAN & BERSIH
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet utang bank
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT BANK")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah atau salah ketik
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'UTANG_INPUT BANK' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Input Utang Bank ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Utang Bank") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat tanpa screen flicker
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES PEMBERSIHAN DATA DENGAN ANTI-MERGER BUG PROTECTION
    With LembarForm
        On Error Resume Next
        
        ' Menggunakan .MergeArea agar aman jika ada sel input yang tidak sengaja ter-merge di desain form
        .Range("D12").MergeArea.ClearContents   ' Kolom Input D12
        .Range("H12").MergeArea.ClearContents   ' Kolom Input H12
        .Range("D16").MergeArea.ClearContents   ' Kolom Input D16
        .Range("H16").MergeArea.ClearContents   ' Kolom Input H16
        .Range("D20").MergeArea.ClearContents   ' Kolom Input D20
        .Range("H20").MergeArea.ClearContents   ' Kolom Input H20
        .Range("D24").MergeArea.ClearContents   ' Kolom Input D24
        .Range("H24").MergeArea.ClearContents   ' Kolom Input H24
        .Range("H28").MergeArea.ClearContents   ' Kolom Input H28
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke posisi awal input utama D12)
    LembarForm.Activate
    LembarForm.Range("D12").Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Utang Bank Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

