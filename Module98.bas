Attribute VB_Name = "Module98"
Sub HapusFormPelepasanAsetTotal()
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet pelepasan aset
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("ASET TETAP_INPUT PELEPASAN")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'ASET TETAP_INPUT PELEPASAN' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Pelepasan Aset ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Pelepasan Aset") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES PEMBERSIHAN DATA DENGAN ANTI-MERGER BUG PROTECTION
    With LembarForm
        On Error Resume Next
        
        ' Menggunakan .MergeArea agar aman jika ada sel input yang tidak sengaja ter-merge
        .Range("D10").MergeArea.ClearContents   ' Kolom Input D10
        .Range("D14").MergeArea.ClearContents   ' Kolom Input D14
        .Range("D18").MergeArea.ClearContents   ' Kolom Input D18
        .Range("F10").MergeArea.ClearContents   ' Kolom Input F10
        .Range("H10").MergeArea.ClearContents   ' Kolom Input H10
        .Range("H18").MergeArea.ClearContents   ' Kolom Input H18
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke posisi awal input utama D10)
    LembarForm.Activate
    LembarForm.Range("D10").Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Pelepasan Aset Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

