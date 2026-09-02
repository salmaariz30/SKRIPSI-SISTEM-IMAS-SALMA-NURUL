Attribute VB_Name = "Module79"
Sub SimpanTransaksiRetur()
    
    Dim LembarForm As Worksheet
    Dim LembarData As Worksheet
    Dim LembarHarian As Worksheet
    Dim LembarPiutang As Worksheet
    
    Dim TabelData As ListObject
    Dim TabelHarian As ListObject
    Dim TabelPiutang As ListObject
    
    Dim BarisBaru As ListRow
    Dim BarisHarian As ListRow
    Dim BarisPiutang As Range
    
    Dim NoUrutTerakhir As Long
    Dim NoUrutHarianTerakhir As Long
    Dim IsBarisPertamaKosong As Boolean
    Dim IsHarianPertamaKosong As Boolean
    Dim KetemuPiutang As Boolean
    
    ' VARIABEL DATA FORMULIR INPUT RETUR
    Dim V_Tanggal As Date
    Dim V_NoBuktiRetur As String
    Dim V_NamaPelangan As String
    Dim V_NoReferensi As String
    Dim V_Kategori As String
    Dim V_Alasan As String
    Dim V_MetodePenyelesaian As String
    Dim V_ProdukReject As String
    Dim V_NominalRetur As Double
    Dim V_NamaBank As String
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT RETUR")
    Set LembarData = ThisWorkbook.Sheets("PENDAPATAN_LAPORAN RETUR")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    ' Cari sheet tempat TabelBukuPiutang berada (Sesuaikan namanya jika berbeda)
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelRetur")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    
    ' Mencari tabel piutang di seluruh sheet agar aman
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.ListObjects.Count > 0 Then
            Set TabelPiutang = Sh.ListObjects("TabelBukuPiutang")
            If Not TabelPiutang Is Nothing Then Exit For
        End If
    Next Sh
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Kritikal
    If TabelData Is Nothing Then
        MsgBox "Error: 'TabelRetur' tidak ditemukan di sheet terkait!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT RETUR (SUPPORT MERGED CELLS)
    With LembarForm
        V_Tanggal = .Range("E11").MergeArea.Cells(1, 1).Value
        V_NoBuktiRetur = .Range("E9").MergeArea.Cells(1, 1).Value
        V_NamaPelangan = .Range("E17").MergeArea.Cells(1, 1).Value
        V_NoReferensi = .Range("E13").MergeArea.Cells(1, 1).Value
        V_Kategori = .Range("K19").MergeArea.Cells(1, 1).Value
        V_Alasan = .Range("K21").MergeArea.Cells(1, 1).Value
        V_MetodePenyelesaian = .Range("K9").MergeArea.Cells(1, 1).Value
        V_NamaBank = .Range("K11").MergeArea.Cells(1, 1).Value
        V_NominalRetur = Val(.Range("K13").MergeArea.Cells(1, 1).Value)
        V_ProdukReject = .Range("K23").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL (Cek input kosong)
    If V_Tanggal = 0 Or V_NoBuktiRetur = "" Or V_NoReferensi = "" Or V_NominalRetur = 0 Or V_MetodePenyelesaian = "" Then
        MsgBox "Form Input Retur Belum Lengkap!" & vbCrLf & _
               "Pastikan Tanggal, No Bukti, No Referensi, Metode, dan Nominal Retur sudah terisi.", _
               vbExclamation, "Validasi Input Gagal"
        Exit Sub
    End If
    
    ' Kunci visual layar biar eksekusi makro super mulus di laptop ASUS
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' 3. LOGIKA DAN EKSEKUSI INDUK INDEKS: TABEL RETUR (PENDAPATAN_LAPORAN RETUR)
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
    
    ' Inject Data Ke Tabel Retur Utama
    With BarisBaru
        .Range(1) = NoUrutTerakhir + 1         ' KOLOM 1: NO
        .Range(2) = V_Tanggal                 ' KOLOM 2: TANGGAL
        .Range(3) = V_NoBuktiRetur            ' KOLOM 3: NO. BUKTI RETUR
        .Range(4) = V_NamaPelangan            ' KOLOM 4: NAMA PELANGGAN
        .Range(5) = V_NoReferensi             ' KOLOM 5: NO. REFERENSI
        .Range(6) = V_Kategori                ' KOLOM 6: KATEGORI
        .Range(7) = V_Alasan                  ' KOLOM 7: ALASAN
        .Range(8) = V_MetodePenyelesaian      ' KOLOM 8: METODE PENYELESAIAN
        .Range(9) = V_ProdukReject            ' KOLOM 9: PRODUK REJECT / CACAT
        .Range(10) = V_NominalRetur           ' KOLOM 10: NOMINAL PENGEMBALIAN
        
        ' Format Estetik Font Segoe UI
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(10).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' ====================================================================
    ' 4A. JALUR PERBATASAN KONDISIONAL: JIKA POTONG PIUTANG (VERSI REVISI  )
    ' ====================================================================
    If V_MetodePenyelesaian = "Potong Piutang" Then
        If Not TabelPiutang Is Nothing Then
            KetemuPiutang = False
            If TabelPiutang.ListRows.Count > 0 Then
                ' Sisir kolom ke-2 (No. Referensi / No. Faktur) di TabelBukuPiutang
                For Each BarisPiutang In TabelPiutang.ListColumns(2).DataBodyRange
                    If CStr(BarisPiutang.Value) = V_NoReferensi Then
                        ' TARGET MUTLAK  : KOLOM KEENAM (Kolom 6) TABEL BUKU PIUTANG
                        ' Menggunakan indeks .Cells(1, 6) secara absolut dari awal kolom tabel
                        BarisPiutang.EntireRow.Cells(1, TabelPiutang.Range.Column + 5).Value = _
                            Val(BarisPiutang.EntireRow.Cells(1, TabelPiutang.Range.Column + 5).Value) + V_NominalRetur
                        KetemuPiutang = True
                        Exit For
                    End If
                Next BarisPiutang
            End If
            
            If Not KetemuPiutang Then
                MsgBox "Peringatan: No. Referensi " & V_NoReferensi & " tidak ditemukan di TabelBukuPiutang!" & vbCrLf & _
                       "Data retur tetap disimpan, namun kartu piutang gagal dipotong otomatis.", vbExclamation, "Referensi Piutang Luput"
            End If
        Else
            MsgBox "Error: 'TabelBukuPiutang' tidak ditemukan! Gagal potong piutang.", vbCritical, "Sistem Macet"
        End If
        
    ' ====================================================================
    ' 4B. JALUR PERBATASAN KONDISIONAL: JIKA REFUND DANA (ARUS KAS)
    ' ====================================================================
    ElseIf V_MetodePenyelesaian = "Pengembalian Dana (Refund)" Then
        If Not TabelHarian Is Nothing Then
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
            
            ' Tembak Data ke Laporan Harian Kas (Metode Langsung Operasional)
            With BarisHarian
                .Range(1) = NoUrutHarianTerakhir + 1        ' KOLOM 1: NO.
                .Range(2) = V_Tanggal                        ' KOLOM 2: TANGGAL
                .Range(3) = V_NamaBank                       ' KOLOM 3: AKUN KAS / BANK (Cell K11)
                .Range(4) = 0                                ' KOLOM 4: DEBIT (0)
                .Range(5) = V_NominalRetur                   ' KOLOM 5: KREDIT (REFUND DANA KULAR)
                .Range(6) = "Retur Penjualan - " & V_NamaPelangan & " (" & V_Alasan & ")" ' KOLOM 6: DESKRIPSI
                .Range(7) = V_NoBuktiRetur                   ' KOLOM 7: NO. BUKTI TRANSAKSI
                .Range(8) = "Penjualan Produk"                    ' KOLOM 8: JENIS AKTIVITAS
                
                ' Format Estetik Kamar Kas
                With .Range.Font
                    .Name = "Segoe UI"
                    .Size = 10
                    .Color = vbBlack
                End With
                .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
                .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            End With
        Else
            MsgBox "Error: 'TabelLaporanHarianKas' hilang! Gagal menyinkronkan arus kas refund.", vbCritical, "Sistem Bocor"
        End If
    End If
    
           
End Sub

