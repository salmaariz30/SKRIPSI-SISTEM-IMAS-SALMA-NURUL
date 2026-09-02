Attribute VB_Name = "Module54"
Sub SimpanTransaksiBiayaProduksi()
    ' ====================================================================
    ' MODUL UTAMA: SIMPAN DATA BIAYA PRODUKSI & SINKRONISASI LAPORAN KAS
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
    
    ' VARIABEL DATA FORMULIR INPUT PRODUKSI
    Dim V_Kategori As String
    Dim V_NoBukti As String
    Dim V_TanggalBayar As Date
    Dim V_AkunKasBank As String
    Dim V_Nominal As Double
    Dim V_NamaVendor As String
    Dim V_Deskripsi As String
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN_INPUT PRODUKSI")
    Set LembarData = ThisWorkbook.Sheets("PENGELUARAN USAHA_PRODUKSI")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelBiayaProduksi")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' Validasi Emergency System
    If TabelData Is Nothing Or TabelHarian Is Nothing Then
        MsgBox "Error: 'TabelBiayaProduksi' atau 'TabelLaporanHarianKas' tidak ditemukan!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT PRODUKSI (SUPPORT MERGED CELLS)
    With LembarForm
        V_Kategori = .Range("D11").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("G11").MergeArea.Cells(1, 1).Value
        V_TanggalBayar = .Range("D15").MergeArea.Cells(1, 1).Value
        V_AkunKasBank = .Range("G15").MergeArea.Cells(1, 1).Value
        V_Nominal = Val(.Range("D19").MergeArea.Cells(1, 1).Value)
        V_NamaVendor = .Range("G19").MergeArea.Cells(1, 1).Value
        V_Deskripsi = .Range("D23").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL (Cek kelengkapan input)
    If V_TanggalBayar = 0 Or V_NoBukti = "" Or V_Kategori = "" Or V_Nominal = 0 Or V_AkunKasBank = "" Then
        MsgBox "Form Input Produksi Belum Lengkap,  !" & vbCrLf & _
               "Pastikan Tanggal, No Bukti, Kategori Biaya, Akun Kas/Bank, dan Nominal sudah terisi.", _
               vbExclamation, "Validasi Input Gagal"
        Exit Sub
    End If
    
    ' Kunci visual layar biar gerakan pengisian datanya super smooth anti-kedip
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' 3. LOGIKA DAN EKSEKUSI TABEL BIAYA PRODUKSI (image_502d01.png)
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
    
    ' Inject Data ke Tabel Biaya Produksi
    With BarisBaru
        .Range(1) = NoUrutTerakhir + 1     ' KOLOM 1: NO
        .Range(2) = V_TanggalBayar         ' KOLOM 2: TANGGAL BAYAR
        .Range(3) = V_NoBukti              ' KOLOM 3: NO. BUKTI TRANSAKSI
        .Range(4) = V_Deskripsi            ' KOLOM 4: DESKRIPSI TRANSAKSI
        .Range(5) = V_NamaVendor           ' KOLOM 5: NAMA VENDOR / PENYEDIA
        .Range(6) = V_Kategori             ' KOLOM 6: KATEGORI BIAYA PRODUKSI
        .Range(7) = V_Nominal              ' KOLOM 7: NOMINAL
        .Range(8) = V_AkunKasBank          ' KOLOM 8: AKUN KAS / BANK
        
        ' Format Estetik Khas     (Segoe UI 9 Black)
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' ====================================================================
    ' 4. SINKRONISASI KE TABEL LAPORAN HARIAN KAS
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
        .Range(2) = V_TanggalBayar           ' KOLOM 2: TANGGAL
        .Range(3) = V_AkunKasBank            ' KOLOM 3: AKUN KAS / BANK
        .Range(4) = 0                        ' KOLOM 4: DEBIT (KOSONG / 0)
        .Range(5) = V_Nominal                ' KOLOM 5: KREDIT (Uang Keluar)
        .Range(6) = V_Deskripsi              ' KOLOM 6: DESKRIPSI TRANSAKSI
        .Range(7) = V_NoBukti                ' KOLOM 7: NO. BUKTI TRANSAKSI
        .Range(8) = "Operasional"            ' KOLOM 8: JENIS AKTIVITAS (MUTLAK)
        
        ' Format Estetik Ruangan Baris Laporan Harian Kas
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
