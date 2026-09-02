Attribute VB_Name = "Module15"
Sub HapusDataFormKas()
    ' ==========================================================
    ' Terkunci Khusus untuk Sheet: KAS&BANK_INPUT TRANSAKSI UMUM
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet input
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("KAS&BANK_INPUT TRANSAKSI UMUM")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'KAS&BANK_INPUT TRANSAKSI UMUM' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum menghapus)
    If MsgBox("Apakah anda yakin ingin mengosongkan semua data pada form ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir") = vbNo Then
        Exit Sub
    End If
    
    ' 2. PROSES PEMBERSIHAN DATA
    With LembarForm
        .Range("E13").MergeArea.ClearContents   ' Tanggal Transaksi
        .Range("H13").MergeArea.ClearContents   ' No. Bukti
        .Range("E17").MergeArea.ClearContents   ' Jumlah Dana Masuk
        .Range("H17").MergeArea.ClearContents   ' Sumber Dana
        .Range("E21").MergeArea.ClearContents   ' Deskripsi Transaksi
        .Range("H25").MergeArea.ClearContents   ' Nama Akun
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke posisi awal)
    LembarForm.Activate
    LembarForm.Range("E13").Select
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Formulir input berhasil dibersihkan!", _
           vbInformation, "Sistem Sukses"
           
End Sub

