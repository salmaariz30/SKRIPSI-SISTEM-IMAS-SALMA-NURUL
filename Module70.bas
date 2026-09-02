Attribute VB_Name = "Module70"
Sub SimpanTransaksiPelunasanPiutang()
    
    Dim LembarForm As Worksheet
    Dim LembarLunas As Worksheet
    Dim LembarHarian As Worksheet
    Dim LembarBuku As Worksheet
    
    Dim TabelLunas As ListObject
    Dim TabelHarian As ListObject
    Dim TabelBuku As ListObject
    
    Dim BarisLunas As ListRow
    Dim BarisHarian As ListRow
    Dim BarisBuku As Range
    
    Dim NoUrutLunas As Long
    Dim NoUrutHarian As Long
    Dim IsLunasPertamaKosong As Boolean
    Dim IsHarianPertamaKosong As Boolean
    Dim InvoiceDitemukan As Boolean
    Dim KolomInvoiceBuku As Range
    Dim KolomTerbayarIdx As Long
    
    ' VARIABEL DATA FORMULIR INPUT PIUTANG
    Dim V_Invoice As String
    Dim V_NamaPelanggan As String
    Dim V_Tanggal As Date
    Dim V_NoBukti As String
    Dim V_JumlahBayar As Double
    Dim V_BayarDari As String
    Dim V_KasBank As String
    Dim V_Status As String
    Dim V_Deskripsi As String
    
    Const PWD As String = "IMAS" '
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PIUTANG_INPUT DATA PIUTANG")
    Set LembarLunas = ThisWorkbook.Sheets("PIUTANG_PELUNASAN PIUTANG")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    Set LembarBuku = ThisWorkbook.Sheets("PIUTANG_BUKU PIUTANG")
    
    On Error Resume Next
    Set TabelLunas = LembarLunas.ListObjects("TabelPelunasanPiutang")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    Set TabelBuku = LembarBuku.ListObjects("TabelBukuPiutang")
    On Error GoTo 0
    
    ' Validasi Keberadaan Seluruh Sistem Tabel
    If TabelLunas Is Nothing Or TabelHarian Is Nothing Or TabelBuku Is Nothing Then
        MsgBox "Error: Salah satu tabel ('TabelPelunasanPiutang', 'TabelLaporanHarianKas', atau 'TabelBukuPiutang') tidak ditemukan!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT PIUTANG   (SUPPORT MERGED CELLS)
    With LembarForm
        V_Invoice = .Range("E13").MergeArea.Cells(1, 1).Value
        V_Tanggal = .Range("E17").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("G17").MergeArea.Cells(1, 1).Value
        V_JumlahBayar = Val(.Range("E21").MergeArea.Cells(1, 1).Value)
        V_BayarDari = .Range("G21").MergeArea.Cells(1, 1).Value
        V_KasBank = .Range("E25").MergeArea.Cells(1, 1).Value
        V_Status = .Range("G25").MergeArea.Cells(1, 1).Value
        V_Deskripsi = .Range("E29").MergeArea.Cells(1, 1).Value
        V_NamaPelanggan = .Range("G13").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL
    If V_Invoice = "" Or V_Tanggal = 0 Or V_NoBukti = "" Or V_JumlahBayar <= 0 Or V_NamaPelanggan = "" Or V_KasBank = "" Then
        MsgBox "Form Pelunasan Piutang Belum Lengkap!" & vbCrLf & _
               "Pastikan No Invoice, Tanggal, No Bukti, Nama Pelanggan, Kas/Bank, dan Jumlah Bayar terisi dengan benar.", _
               vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' ====================================================================
    ' ?? OPERASI OPEN LOCK: Jebol semua gembok sheet tujuan ekspor sebelum menulis data
    ' ====================================================================
    LembarLunas.Unprotect Password:=PWD
    LembarHarian.Unprotect Password:=PWD
    LembarBuku.Unprotect Password:=PWD
    
    ' Kunci visual layar biar gerakan pengisian datanya super smooth tanpa kedip
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' ====================================================================
    ' 3. EKSEKUSI DATA KE TABEL PELUNASAN PIUTANG (10 KOLOM MUTLAK)
    ' ====================================================================
    IsLunasPertamaKosong = False
    On Error Resume Next
    If TabelLunas.ListRows.Count = 0 Then
        IsLunasPertamaKosong = True
        NoUrutLunas = 0
    ElseIf TabelLunas.ListRows.Count = 1 And (TabelLunas.DataBodyRange.Cells(1, 1).Value = "" Or TabelLunas.DataBodyRange.Cells(1, 2).Value = "") Then
        IsLunasPertamaKosong = True
        NoUrutLunas = 0
    Else
        NoUrutLunas = Application.WorksheetFunction.Max(TabelLunas.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsLunasPertamaKosong Then
        If TabelLunas.ListRows.Count = 0 Then TabelLunas.ListRows.Add
        Set BarisLunas = TabelLunas.ListRows(1)
    Else
        Set BarisLunas = TabelLunas.ListRows.Add
    End If
    
    With BarisLunas
        .Range(1) = V_Tanggal             ' KOLOM 2: TANGGAL
        .Range(2) = V_NoBukti             ' KOLOM 3: NO BUKTI
        .Range(3) = V_Invoice             ' KOLOM 4: NO INVOICE
        .Range(4) = V_Deskripsi           ' KOLOM 5: DESKRIPSI
        .Range(5) = V_NamaPelanggan       ' KOLOM 6: NAMA PELANGGAN
        .Range(6) = V_JumlahBayar         ' KOLOM 7: JUMLAH BAYAR
        .Range(7) = V_BayarDari           ' KOLOM 8: BAYAR DARI
        .Range(8) = V_Status              ' KOLOM 9: STATUS PEMBAYARAN
        .Range(9) = V_KasBank            ' KOLOM 10: KAS DAN BANK
        
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(6).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' ====================================================================
    ' 4. SINKRONISASI KE TABEL LAPORAN HARIAN KAS (DEBIT = UANG MASUK)
    ' ====================================================================
    IsHarianPertamaKosong = False
    On Error Resume Next
    If TabelHarian.ListRows.Count = 0 Then
        IsHarianPertamaKosong = True
        NoUrutHarian = 0
    ElseIf TabelHarian.ListRows.Count = 1 And (TabelHarian.DataBodyRange.Cells(1, 1).Value = "" Or TabelHarian.DataBodyRange.Cells(1, 2).Value = "") Then
        IsHarianPertamaKosong = True
        NoUrutHarian = 0
    Else
        NoUrutHarian = Application.WorksheetFunction.Max(TabelHarian.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsHarianPertamaKosong Then
        If TabelHarian.ListRows.Count = 0 Then TabelHarian.ListRows.Add
        Set BarisHarian = TabelHarian.ListRows(1)
    Else
        Set BarisHarian = TabelHarian.ListRows.Add
    End If
    
    With BarisHarian
        .Range(1) = NoUrutHarian + 1                                       ' KOLOM 1: NO.
        .Range(2) = V_Tanggal                                              ' KOLOM 2: TANGGAL
        .Range(3) = V_KasBank                                              ' KOLOM 3: AKUN KAS / BANK
        .Range(4) = V_JumlahBayar                                          ' KOLOM 4: DEBIT
        .Range(5) = 0                                                      ' KOLOM 5: KREDIT
        .Range(6) = "Pelunasan Piutang - " & V_NamaPelanggan & " (" & V_Deskripsi & ")"
        .Range(7) = V_NoBukti                                              ' KOLOM 7: NO. BUKTI
        .Range(8) = "Operasional"                                          ' KOLOM 8: JENIS AKTIVITAS
        
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' ====================================================================
    ' 5. SISIR TABEL BUKU PIUTANG & AKUMULASIKAN KE KOLOM "JUMLAH TERBAYAR" (FIX KOLOM 1)
    ' ====================================================================
    InvoiceDitemukan = False
    Dim BarisKeBerapa As Long
    
    On Error Resume Next
    KolomTerbayarIdx = TabelBuku.ListColumns("JUMLAH TERBAYAR").Index
    On Error GoTo 0
    
    ' Cadangan jika penamaan tabel goib terdeteksi, default ke kolom 7
    If KolomTerbayarIdx = 0 Then KolomTerbayarIdx = 7
    
    If TabelBuku.ListRows.Count > 0 Then
        ' ?? REVISI SAKRAL: Set ke Kolom 1 karena No. Invoice ada di kolom pertama!
        Set KolomInvoiceBuku = TabelBuku.ListColumns(1).DataBodyRange
        BarisKeBerapa = 0
        
        For Each BarisBuku In KolomInvoiceBuku
            BarisKeBerapa = BarisKeBerapa + 1 ' Hitung nomor baris di dalam tabel
            
            ' Validasi kecocokan No. Invoice (Teks & Angka disamakan tipenya)
            If UCase(Trim(CStr(BarisBuku.Value))) = UCase(Trim(CStr(V_Invoice))) Then
                
                ' Tembak langsung koordinat cell target tanpa rumus .Offset yang membingungkan
                With TabelBuku.DataBodyRange.Cells(BarisKeBerapa, KolomTerbayarIdx)
                    .Value = Val(CStr(.Value)) + V_JumlahBayar
                End With
                
                InvoiceDitemukan = True
                Exit For
            End If
        Next BarisBuku
    End If

    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Pasang kembali satpam proteksi ketiga sheet target
    ' ====================================================================
LembarPemulihan:
    LembarLunas.Protect Password:=PWD, AllowFiltering:=True
    LembarHarian.Protect Password:=PWD, AllowFiltering:=True
    LembarBuku.Protect Password:=PWD, AllowFiltering:=True

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    
End Sub

