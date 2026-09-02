Attribute VB_Name = "Module51"
Sub SimpanTransaksiPenarikanModal()
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarData As Worksheet
    Dim LembarHarian As Worksheet
    Dim TabelData As ListObject
    Dim TabelHarian As ListObject
    Dim BarisBaru As ListRow
    Dim BarisHarian As ListRow
    
    Dim NoUrutTerakhir As Long
    Dim NoUrutHarianTerakhir As Long
    Dim IsBarisPertamaKosong As Boolean
    Dim IsHarianPertamaKosong As Boolean
    Const PWD As String = "IMAS"
    
    ' VARIABEL DATA FORMULIR INPUT TARIK
    Dim V_JenisPenarikan As String
    Dim V_NoBukti As String
    Dim V_Tanggal As Date
    Dim V_AkunKasBank As String
    Dim V_Nominal As Double
    Dim V_NamaPenarik As String
    Dim V_RekeningPenarik As String
    Dim V_Persentase As Double
    Dim V_Deskripsi As String
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT TARIK")
    Set LembarData = ThisWorkbook.Sheets("PENGELUARAN USAHA_UTANG BIAYA")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelPenarikanModal")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel
    If TabelData Is Nothing Or TabelHarian Is Nothing Then
        MsgBox "Error: 'TabelPenarikanModal' atau 'TabelLaporanHarianKas' tidak ditemukan!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT TARIK (SUPPORT MERGED CELLS)
    With LembarForm
        V_JenisPenarikan = .Range("D11").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("G11").MergeArea.Cells(1, 1).Value
        V_Tanggal = .Range("D15").MergeArea.Cells(1, 1).Value
        V_AkunKasBank = .Range("G15").MergeArea.Cells(1, 1).Value
        V_Nominal = Val(.Range("D19").MergeArea.Cells(1, 1).Value)
        V_NamaPenarik = .Range("G19").MergeArea.Cells(1, 1).Value
        V_RekeningPenarik = .Range("D23").MergeArea.Cells(1, 1).Value
        V_Persentase = Val(.Range("G23").MergeArea.Cells(1, 1).Value)
        V_Deskripsi = .Range("D27").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL (Cek jika input kritikal kosong melompong)
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_JenisPenarikan = "" Or V_Nominal = 0 Or V_NamaPenarik = "" Or V_AkunKasBank = "" Then
        MsgBox "Form Input Tarik Belum Lengkap!" & vbCrLf & _
               "Pastikan Tanggal, No Bukti, Jenis Penarikan, Akun Kas/Bank, Nama Penarik, dan Nominal sudah terisi.", _
               vbExclamation, "Validasi Input Gagal"
        Exit Sub
    End If
    
    ' Kunci visual layar biar gerakan pengisian datanya super smooth di laptop ASUS
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Buka Gembok Lembar Target sebelum diisi data baru
    LembarData.Unprotect Password:=PWD
    LembarHarian.Unprotect Password:=PWD
    
    ' ====================================================================
    ' 3. LOGIKA DAN EKSEKUSI TABEL DATA PENARIKAN MODAL
    ' ====================================================================
    IsBarisPertamaKosong = False
    On Error Resume Next
    If TabelData.ListRows.Count = 0 Then
        IsBarisPertamaKosong = True
        NoUrutTerakhir = 0
    ElseIf TabelData.ListRows.Count = 1 And (TabelData.DataBodyRange.Cells(1, 1).Value = "" Or TabelData.DataBodyRange.Cells(1, 2).Value = "") Then
        IsBarisPertamaKosong = True
        NoUrutTerakhir = 0
    Else
        NoUrutTerakhir = Application.WorksheetFunction.Max(TabelData.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsBarisPertamaKosong Then
        If TabelData.ListRows.Count = 0 Then TabelData.ListRows.Add
        Set BarisBaru = TabelData.ListRows(1)
    Else
        Set BarisBaru = TabelData.ListRows.Add
    End If
    
    ' Inject Data ke Tabel Penarikan Modal
    With BarisBaru
        .Range(1) = NoUrutTerakhir + 1     ' KOLOM 1: NO.
        .Range(2) = V_Tanggal              ' KOLOM 2: TANGGAL PENARIKAN
        .Range(3) = V_NoBukti              ' KOLOM 3: NO. BUKTI TRANSAKSI
        .Range(4) = V_Deskripsi            ' KOLOM 4: DESKRIPSI
        .Range(5) = V_JenisPenarikan       ' KOLOM 5: JENIS PENARIKAN
        .Range(6) = V_NamaPenarik          ' KOLOM 6: NAMA PENARIK
        .Range(7) = V_Persentase           ' KOLOM 7: PERSENTASE KEPEMILIKAN
        .Range(8) = V_Nominal              ' KOLOM 8: NOMINAL
        .Range(9) = V_AkunKasBank          ' KOLOM 9: AKUN KAS / BANK
        .Range(10) = V_RekeningPenarik     ' KOLOM 10: NO. REKENING PENARIK
        
        ' Format Estetik Ruangan Baris Baru
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(7).NumberFormat = "0.0%"
        .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' ====================================================================
    ' 4. LOGIKA DAN SINKRONISASI KE TABEL LAPORAN HARIAN KAS (image_5b7b3d.png)
    ' ====================================================================
    IsHarianPertamaKosong = False
    On Error Resume Next
    If TabelHarian.ListRows.Count = 0 Then
        IsHarianPertamaKosong = True
        NoUrutHarianTerakhir = 0
    ElseIf TabelHarian.ListRows.Count = 1 And (TabelHarian.DataBodyRange.Cells(1, 1).Value = "" Or TabelHarian.DataBodyRange.Cells(1, 2).Value = "") Then
        IsHarianPertamaKosong = True
        NoUrutHarianTerakhir = 0
    Else
        NoUrutHarianTerakhir = Application.WorksheetFunction.Max(TabelHarian.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsHarianPertamaKosong Then
        If TabelHarian.ListRows.Count = 0 Then TabelHarian.ListRows.Add
        Set BarisHarian = TabelHarian.ListRows(1)
    Else
        Set BarisHarian = TabelHarian.ListRows.Add
    End If
    
    ' Inject Data ke Tabel Laporan Harian Kas Sesuai image_5b7b3d.png
    With BarisHarian
        .Range(1) = NoUrutHarianTerakhir + 1 ' KOLOM 1: NO.
        .Range(2) = V_Tanggal                ' KOLOM 2: TANGGAL
        .Range(3) = V_AkunKasBank            ' KOLOM 3: AKUN KAS / BANK
        .Range(4) = 0                        ' KOLOM 4: DEBIT (KOSONG / 0)
        .Range(5) = V_Nominal                ' KOLOM 5: KREDIT
        .Range(6) = V_Deskripsi              ' KOLOM 6: DESKRIPSI TRANSAKSI
        .Range(7) = V_NoBukti                ' KOLOM 7: NO. BUKTI TRANSAKSI
        .Range(8) = "Distribusi Laba Pemilik"          ' KOLOM 8: JENIS AKTIVITAS
        
        ' Format Estetik Ruangan Baris Laporan Harian Kas
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' Kunci Kembali Lembar Target setelah selesai proses pembaruan
    LembarData.Protect Password:=PWD, AllowFiltering:=True
    LembarHarian.Protect Password:=PWD, AllowFiltering:=True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

