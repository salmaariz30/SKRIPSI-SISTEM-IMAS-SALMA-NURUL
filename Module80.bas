Attribute VB_Name = "Module80"
Sub JurnalUmumInputReturPenjualan()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT RETUR PENJUALAN KE JURNAL UMUM
    ' Terkunci Khusus untuk Sheet: PENDAPATAN_INPUT RETUR
    ' Otomatis Deteksi Kondisi Cell K19 (Debit) & K9/K11 (Kredit) Berdasarkan COA
    ' SPECIAL EDITION: COMPATIBLE WITH USA REGIONAL SETTINGS (EXCEL USA)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject, tblCOA As ListObject
    Dim wsInput As Worksheet
    
    Set wsInput = Sheets("PENDAPATAN_INPUT RETUR")
    
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Utama
    If tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: 'TabelJurnalUmum' atau 'TabelDataCOA' tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. Tarik Data dari Cell Input Retur   (Mendukung Merged Cells)
    Dim V_Tanggal As Date: V_Tanggal = wsInput.Range("E11").Value
    Dim V_NoBuktiRetur As String: V_NoBuktiRetur = wsInput.Range("E9").Value
    Dim V_NamaPelangan As String: V_NamaPelangan = wsInput.Range("E17").Value
    Dim V_NoReferensi As String: V_NoReferensi = wsInput.Range("E13").Value
    Dim V_Kategori As String: V_Kategori = wsInput.Range("K19").Value
    Dim V_Alasan As String: V_Alasan = wsInput.Range("K21").Value
    Dim V_MetodePenyelesaian As String: V_MetodePenyelesaian = wsInput.Range("K9").Value
    Dim V_NamaBank As String: V_NamaBank = wsInput.Range("K11").Value
    Dim V_NominalRetur As Double: V_NominalRetur = Val(wsInput.Range("K13").Value)
    Dim V_ProdukReject As String: V_ProdukReject = wsInput.Range("K23").Value
    
    ' Tambahkan deskripsi otomatis yang informatif untuk Jurnal Umum
    Dim V_Deskripsi_Jurnal As String
    V_DeskripsiJurnal = "Retur/Potongan Penjualan an. " & V_NamaPelangan & " (" & V_Alasan & ")"
    
    ' 3. LOGIKA DINAMIS TENTUKAN NAMA AKUN DEBIT BERDASARKAN K19
    Dim namaAkunDebit As String
    Select Case V_Kategori
        Case "Potongan_Penjualan", "Potongan Penjualan"
            namaAkunDebit = "Potongan Penjualan"
            
        Case "Retur_Penjualan", "Retur Penjualan"
            namaAkunDebit = "Retur Penjualan"
            
        Case Else
            MsgBox "Kategori K19 '" & V_Kategori & "' tidak dikenal oleh logika sistem jurnal!", vbCritical, "Error Logika Debit"
            Exit Sub
    End Select
    
    ' 4. LOGIKA DINAMIS TENTUKAN NAMA AKUN KREDIT BERDASARKAN K9 / K11
    Dim namaAkunKredit As String
    Select Case V_MetodePenyelesaian
        Case "Potong Piutang"
            namaAkunKredit = "Piutang Usaha"
            
        Case "Pengembalian Dana (Refund)"
            namaAkunKredit = V_NamaBank ' Menembak langsung nama akun Kas/Bank di sel K11
            
        Case Else
            MsgBox "Metode Penyelesaian K9 '" & V_MetodePenyelesaian & "' tidak dikenal oleh logika sistem jurnal!", vbCritical, "Error Logika Kredit"
            Exit Sub
    End Select
    
    ' 5. Ambil Kode Akun (Kolom 1) Berdasarkan Nama Akun (Kolom 2) di COA secara Otomatis
    Dim barisDebit As Variant, barisKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisDebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
    On Error GoTo 0
    
    ' Proteksi jika nama akun tidak terdaftar di master COA
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        MsgBox "Gagal menjurnal! Nama akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' tidak ditemukan di TabelDataCOA!", _
               vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' 6. EKSEKUSI EXPOR DUA BARIS KE JURNAL UMUM (ANTI-KEDIP)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' --- BARIS DEBIT (Retur / Potongan Bertambah di Sisi Kiri) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBuktiRetur              ' Kolom 2: No Bukti
        .Range(3) = V_DeskripsiJurnal           ' Kolom 3: Keterangan / Deskripsi
        .Range(4) = kodeAkunDebit               ' Kolom 4: Kode Akun
        .Range(5) = namaAkunDebit               ' Kolom 5: Nama Akun
        .Range(6) = V_NominalRetur              ' Kolom 6: Debit
        .Range(7) = 0                           ' Kolom 7: Kredit
        .Range(8) = "Pendapatan"                ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' --- BARIS KREDIT (Piutang Berkurang / Kas Keluar di Sisi Kanan) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBuktiRetur              ' Kolom 2: No Bukti
        .Range(3) = V_DeskripsiJurnal           ' Kolom 3: Keterangan / Deskripsi
        .Range(4) = kodeAkunKredit              ' Kolom 4: Kode Akun
        .Range(5) = namaAkunKredit              ' Kolom 5: Nama Akun
        .Range(6) = 0                           ' Kolom 6: Debit
        .Range(7) = V_NominalRetur              ' Kolom 7: Kredit
        .Range(8) = IIf(V_MetodePenyelesaian = "Potong Piutang", "Aset", "Kas & Bank") ' Kolom 8: Kelompok Dinamis
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' 7. SILENT CLEAN FORM RETUR (Otomatis Bersih Kilat & Kembalikan Kursor)
    With wsInput
        .Range("E9").MergeArea.ClearContents   ' No Bukti Retur
        .Range("E11").MergeArea.ClearContents  ' Tanggal
        .Range("E13").MergeArea.ClearContents  ' No Referensi
        .Range("K9").MergeArea.ClearContents   ' Metode Penyelesaian
        .Range("K11").MergeArea.ClearContents  ' Nama Bank
        .Range("K13").MergeArea.ClearContents  ' Nominal Retur
        .Range("K19").MergeArea.ClearContents  ' Kategori
        .Range("K21").MergeArea.ClearContents  ' Alasan
        .Range("K23").MergeArea.ClearContents  ' Produk Reject
        .Range("E21").MergeArea.ClearContents  ' Status Awal Transaksi
        
        ' UX Flow: Kembalikan kursor manis   ke hulu cell awal (E9)
        .Activate
        .Range("E9").MergeArea.Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 8. NOTIFIKASI SUKSES
    MsgBox "Data Retur Berhasil Disimpan!", _
           vbInformation, "Sistem Sukses"
End Sub

