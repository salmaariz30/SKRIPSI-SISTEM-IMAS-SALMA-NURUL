Attribute VB_Name = "Module36"
Sub JurnalUmumInputKas()
    ' 1. Set Sheet dan Tabel
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("KAS&BANK_INPUT TRANSAKSI UMUM")
    Dim SheetJurnal As Worksheet: Set SheetJurnal = tblJurnal.Parent
    
    ' 2. Tarik Data dari Cell Input
    Dim tanggal As Variant: tanggal = wsInput.Range("E13").Value
    Dim noBukti As Variant: noBukti = wsInput.Range("H13").Value
    Dim nominal As Variant: nominal = wsInput.Range("E17").Value
    Dim keterangan As Variant: keterangan = wsInput.Range("E21").Value
    Dim akunKasBank As String: akunKasBank = wsInput.Range("H25").Value
    
    ' 3. Tentukan Nama Akun Kredit dari Dropdown Sumber Dana (KOREKSI LOGIKA BARU)
    Dim namaAkunKredit As String
    Select Case wsInput.Range("H17").Value
        Case "Saldo Awal":                  namaAkunKredit = "Ekuitas - Saldo Awal"
        Case "Tambahan Modal Pemilik":      namaAkunKredit = "Modal Tambahan Disetor"
        Case "Uang Pinjaman Bank":          namaAkunKredit = "Utang Bank Jangka Panjang"
        Case "Uang Pinjaman Pihak Lain":    namaAkunKredit = "Utang Lain-Lain"
        Case "Sisa Keuntungan Tahun Lalu":  namaAkunKredit = "Saldo Laba (Laba Ditahan)"
    End Select
    
    ' Validasi jika dropdown sumber dana tidak cocok dengan case di atas
    If namaAkunKredit = "" Then
        MsgBox "Peringatan: Sumber Dana tidak dikenali atau belum dipilih!", vbExclamation, "Input Tidak Valid"
        Exit Sub
    End If
    
    ' 4. Ambil Kode Akun (Kolom 1) Berdasarkan Nama Akun (Kolom 2) di COA
    Dim barisDebit As Variant, barisKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    barisDebit = Application.Match(akunKasBank, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    ' Validasi Keamanan: Cek apakah akun terdaftar di COA agar sistem tidak break
    If IsError(barisDebit) Or IsError(barisKredit) Then
        MsgBox "Error: Akun '" & akunKasBank & "' atau '" & namaAkunKredit & "' tidak ditemukan di Tabel COA!", vbCritical, "Gagal Posting"
        Exit Sub
    End If
    
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
    
    ' 5. Ekspor Dua Baris ke Jurnal Umum
    Application.ScreenUpdating = False
    
    ' BUKA PROTEKSI
    SheetJurnal.Unprotect Password:="IMAS"
    
    ' --- BARIS DEBIT ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = tanggal: .Range(2) = noBukti: .Range(3) = keterangan: .Range(4) = kodeAkunDebit
        .Range(5) = akunKasBank: .Range(6) = nominal: .Range(7) = 0: .Range(8) = "Kas dan Bank": .Range(9) = "Tidak"
    End With
    
    ' --- BARIS KREDIT ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = tanggal: .Range(2) = noBukti: .Range(3) = keterangan: .Range(4) = kodeAkunKredit
        .Range(5) = namaAkunKredit: .Range(6) = 0: .Range(7) = nominal: .Range(8) = "Kas dan Bank": .Range(9) = "Tidak"
    End With
    
    ' TUTUP PROTEKSI
    SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
    
    ' 6. Pembersihan Form Input Otomatis
    With wsInput
        .Range("E13").MergeArea.ClearContents
        .Range("H13").MergeArea.ClearContents
        .Range("E17").MergeArea.ClearContents
        .Range("H17").MergeArea.ClearContents
        .Range("E21").MergeArea.ClearContents
        .Range("H25").MergeArea.ClearContents
        
        ' Kembalikan kursor ke posisi awal input tanggal
        .Activate
        .Range("E13").Select
    End With
    
    Application.ScreenUpdating = True
    
    ' Notifikasi Profesional
    MsgBox "Transaksi telah berhasil dicatat!", _
            vbInformation, "Penyimpanan Sukses"
End Sub

