Attribute VB_Name = "Module23"
Sub HapusDataFormPindahSaldo()
    ' ==========================================================
    ' Terkunci Khusus untuk Sheet: KAS&BANK_INPUT PINDAH SALDO
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet input Pindah Saldo
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("KAS&BANK_INPUT PINDAH SALDO")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah atau typo
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'KAS&BANK_INPUT PINDAH SALDO' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum menghapus)
    If MsgBox("Apakah anda yakin ingin mengosongkan semua data pada form ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Pindah Saldo") = vbNo Then
        Exit Sub
    End If
    
    ' 2. PROSES PEMBERSIHAN DATA
    With LembarForm
        .Range("D11").MergeArea.ClearContents   ' Tanggal Transfer
        .Range("G11").MergeArea.ClearContents   ' No Bukti Transaksi
        .Range("D15").MergeArea.ClearContents   ' Akun Tujuan
        .Range("G15").MergeArea.ClearContents   ' Akun Asal
        .Range("D20").MergeArea.ClearContents   ' Nominal Transfer
        .Range("G20").MergeArea.ClearContents   ' Biaya Admin
        .Range("D24").MergeArea.ClearContents   ' Deskripsi Transaksi
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor otomatis ke posisi awal)
    LembarForm.Activate
    LembarForm.Range("D11").Select
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Formulir Pindah Saldo berhasil dibersihkan! Silakan input transaksi baru.", _
           vbInformation, "Sistem Sukses"
           
End Sub

