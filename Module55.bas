Attribute VB_Name = "Module55"
Sub TombolSimpanPelepasanAset_Click()
    Dim wsInput As Worksheet
    Dim tblJurnal As ListObject
    Dim BarisSebelum As Long, BarisSesudah As Long
    
    Set wsInput = Sheets("ASET TETAP_INPUT PELEPASAN")
    
    ' 1. Hitung jumlah baris di Tabel Jurnal sebelum makro dijalankan
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    If Not tblJurnal Is Nothing Then
        BarisSebelum = tblJurnal.ListRows.Count
    Else
        BarisSebelum = 0
    End If
    On Error GoTo 0
    
    ' 2. JALANKAN MODUL PELEPASAN BAWAAN
    Call JurnalUmumInputPelepasanAset
    Call SimpanPelepasanKeTabel
    
    
    ' 3. Hitung jumlah baris setelah kedua makro selesai berjalan
    If Not tblJurnal Is Nothing Then
        BarisSesudah = tblJurnal.ListRows.Count
    Else
        BarisSesudah = 0
    End If
    
    ' ==========================================================
    ' EVALUASI: JIKA BARIS TIDAK BERTAMBAH = BERARTI KENA VALIDASI GAGAL!
    ' ==========================================================
    ' (Langsung ngerem di sini, amankan form agar tidak kehapus)
    If BarisSesudah <= BarisSebelum Then Exit Sub
    
    ' ==========================================================
    ' JIKA LOLOS (BARIS BERTAMBAH) -> BARU SAPU BERSIH FORM NOTA PELEPASAN
    ' ==========================================================
    On Error Resume Next
    Application.EnableEvents = False
    
    ' Pembersihan khusus sel Merged satu per satu di Form Pelepasan
    wsInput.Range("D10").MergeArea.ClearContents
    wsInput.Range("F10").MergeArea.ClearContents
    wsInput.Range("D14").MergeArea.ClearContents
    wsInput.Range("D18").MergeArea.ClearContents
    wsInput.Range("H18").MergeArea.ClearContents
    
    Application.EnableEvents = True
    On Error GoTo 0
    
    ' NOTIFIKASI JUJUR KHUSUS PELEPASAN ASET
    MsgBox "Berhasil Menyimpan Data Pelepasan Aset Tetap!", vbInformation, "Sukses"
End Sub
