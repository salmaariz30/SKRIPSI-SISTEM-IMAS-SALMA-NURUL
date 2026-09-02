Attribute VB_Name = "Module60"
Sub SimpanTransaksiPendapatanPenjualan()
    ' ====================================================================
    '
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
    
    ' VARIABEL DATA FORMULIR INPUT PENJUALAN
    Dim V_NoBukti As String
    Dim V_Tanggal As Date
    Dim V_NamaPelangan As String
    Dim V_Deskripsi As String
    Dim V_MetodePencatatan As String
    Dim V_MetodePembayaran As String
    
    ' VARIABEL NOMINAL (LOGIKA CONDITIONAL JALUR CELL)
    Dim V_PendapatanKotor As Double
    Dim V_Pajak As Double
    Dim V_BiayaPengiriman As Double
    Dim V_TotalPeroleh As Double
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    Set LembarData = ThisWorkbook.Sheets("PENDAPATAN_DATA PENDAPATAN")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelRincianPenjualan")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel
    If TabelData Is Nothing Or TabelHarian Is Nothing Then
        MsgBox "Error: 'TabelRincianPenjualan' atau 'TabelLaporanHarianKas' tidak ditemukan!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT PENJUALAN (SUPPORT MERGED CELLS)
    With LembarForm
        V_NoBukti = .Range("I15").MergeArea.Cells(1, 1).Value
        V_Tanggal = .Range("C11").MergeArea.Cells(1, 1).Value
        V_NamaPelangan = .Range("C15").MergeArea.Cells(1, 1).Value
        V_Deskripsi = .Range("C19").MergeArea.Cells(1, 1).Value
        V_MetodePencatatan = .Range("F15").MergeArea.Cells(1, 1).Value
        V_MetodePembayaran = .Range("I11").MergeArea.Cells(1, 1).Value
        
        ' --- LOGIKA PERCABANGAN: PILIHAN METODE PENCATATAN ---
        If Trim(V_MetodePencatatan) = "Per Transaksi" Then
            V_PendapatanKotor = Val(.Range("L39").MergeArea.Cells(1, 1).Value)
            V_Pajak = Val(.Range("L40").MergeArea.Cells(1, 1).Value)
            V_BiayaPengiriman = Val(.Range("L41").MergeArea.Cells(1, 1).Value)
            V_TotalPeroleh = Val(.Range("L42").MergeArea.Cells(1, 1).Value)
        ElseIf Trim(V_MetodePencatatan) = "Rekap Harian" Then
            V_PendapatanKotor = Val(.Range("L84").Value) ' Tidak diasumsikan merger, tapi aman jika merger
            V_Pajak = Val(.Range("L85").Value)
            V_BiayaPengiriman = Val(.Range("L86").Value)
            V_TotalPeroleh = Val(.Range("L87").Value)
        Else
            ' Antisipasi jika user salah input atau kosong
            MsgBox "Metode Pencatatan pada Cell F15 Tidak Valid!" & vbCrLf & _
                   "Gunakan kata 'Per Transaksi' atau 'Rekap Harian'.", vbCritical, "Metode Invalid"
            Exit Sub
        End If
    End With
    
    ' VALIDASI INTERNAL CONTROL (Cek jika input kritikal kosong melompong)
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_MetodePencatatan = "" Or V_TotalPeroleh = 0 Or V_MetodePembayaran = "" Then
        MsgBox "Form Input Penjualan Belum Lengkap!" & vbCrLf & _
               "Pastikan Tanggal, No Bukti, Metode Pencatatan, Metode Pembayaran, dan Total Perolehan sudah terisi.", _
               vbExclamation, "Validasi Input Gagal"
        Exit Sub
    End If
    
    ' Kunci visual layar biar gerakan pengisian datanya super smooth di laptop ASUS
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' BUKA PROTEKSI SHEET TARGET BEFORE INJECT DATA
    LembarData.Unprotect Password:=PWD
    LembarHarian.Unprotect Password:=PWD
    
    ' ====================================================================
    ' 3. LOGIKA DAN EKSEKUSI TABEL DATA PENDAPATAN PENJUALAN
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
    
    ' Inject Data ke Tabel Rincian Penjualan
    With BarisBaru
        .Range(1) = NoUrutTerakhir + 1         ' KOLOM 1: NO.
        .Range(2) = V_NoBukti                  ' KOLOM 2: NO. BUKTI TRANSAKSI
        .Range(3) = V_Tanggal                  ' KOLOM 3: TANGGAL TRANSAKSI
        .Range(4) = V_NamaPelangan             ' KOLOM 4: NAMA PELANGGAN
        .Range(5) = V_Deskripsi                ' KOLOM 5: DESKRIPSI PENJUALAN
        .Range(6) = V_MetodePencatatan         ' KOLOM 6: METODE PENCATATAN
        .Range(7) = V_PendapatanKotor          ' KOLOM 7: PENDAPATAN KOTOR
        .Range(8) = V_Pajak                    ' KOLOM 8: PAJAK
        .Range(9) = V_BiayaPengiriman          ' KOLOM 9: BIAYA PENGIRIMAN
        .Range(10) = V_TotalPeroleh            ' KOLOM 10: TOTAL PEROLEHAN
        .Range(11) = V_MetodePembayaran        ' KOLOM 11: METODE PEMBAYARAN
        
        ' Format Estetik Ruangan Baris Baru
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(9).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(10).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With

   ' ====================================================================
    ' 4. LOGIKA DAN SINKRONISASI KE TABEL LAPORAN HARIAN KAS (MODIFIKASI JALUR)
    ' ====================================================================
    ' Cek kriteria: Jika Metode Pembayaran adalah Kredit, jangan ekspor data ke laporan kas
    If Trim(V_MetodePembayaran) <> "Kredit" Then
        
        Dim V_AkunKasBank As String
        
        ' Identifikasi nama akun kas / bank berdasarkan pilihan di cell I11
        If Trim(V_MetodePembayaran) = "Tunai" Then
            V_AkunKasBank = "Kas"
        ElseIf Trim(V_MetodePembayaran) = "Transfer Bank" Then
            V_AkunKasBank = LembarForm.Range("L11").MergeArea.Cells(1, 1).Value
        Else
            V_AkunKasBank = V_MetodePembayaran
        End If
        
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
        
        ' Inject Data ke Tabel Laporan Harian Kas
        With BarisHarian
            .Range(1) = NoUrutHarianTerakhir + 1   ' KOLOM 1: NO.
            .Range(2) = V_Tanggal                  ' KOLOM 2: TANGGAL
            .Range(3) = V_AkunKasBank              ' KOLOM 3: ACUAN AKUN KAS / BANK (HASIL IDENTIFIKASI)
            .Range(4) = V_TotalPeroleh             ' KOLOM 4: DEBIT (Uang Masuk dari Penjualan)
            .Range(5) = 0                          ' KOLOM 5: KREDIT (0)
            .Range(6) = V_Deskripsi                ' KOLOM 6: DESKRIPSI TRANSAKSI
            .Range(7) = V_NoBukti                  ' KOLOM 7: NO. BUKTI TRANSAKSI
            .Range(8) = "Penjualan Produk"         ' KOLOM 8: JENIS AKTIVITAS
                    
            ' Format Estetik Ruangan Baris Laporan Harian Kas
            With .Range.Font
                .Name = "Segoe UI"
                .Size = 10
                .Color = vbBlack
            End With
            .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        End With
        
    End If
    
    ' KUNCI KEMBALI SHEET TARGET RESMI
    LembarData.Protect Password:=PWD, AllowFiltering:=True
    LembarHarian.Protect Password:=PWD, AllowFiltering:=True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

