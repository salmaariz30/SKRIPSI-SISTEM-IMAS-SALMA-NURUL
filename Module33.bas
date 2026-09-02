Attribute VB_Name = "Module33"
Sub SimpanTransaksiPenghasilanLain()
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarData As Worksheet, LembarHarian As Worksheet
    Dim TabelData As ListObject, TabelHarian As ListObject
    Dim BarisBaru As ListRow, BarisHarian As ListRow
    Dim NoUrutTerakhir As Long, NoUrutHarianTerakhir As Long
    Dim IsBarisPertamaKosong As Boolean, IsHarianPertamaKosong As Boolean
    
    ' VARIABEL DATA FORMULIR
    Dim V_Tanggal As Date
    Dim V_Kategori As String
    Dim V_NoBukti As String
    Dim V_Metode As String
    Dim V_AkunPenerima As String
    Dim V_Jumlah As Double
    Dim V_Deskripsi As String
    Dim V_DiterimaDari As String
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_PENGHASILAN LAIN")
    Set LembarData = ThisWorkbook.Sheets("PENDAPATAN_DATA PENGH. LAIN")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelPenghasilanLain")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Kembar
    If TabelData Is Nothing Or TabelHarian Is Nothing Then
        MsgBox "Error: 'TabelPenghasilanLain' atau 'TabelLaporanHarianKas' tidak ditemukan di database!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR EXTREME   (SUPPORT MERGED CELLS)
    With LembarForm
        V_Tanggal = .Range("D11").MergeArea.Cells(1, 1).Value
        V_Kategori = .Range("F11").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("H11").MergeArea.Cells(1, 1).Value
        V_Metode = .Range("D15").MergeArea.Cells(1, 1).Value
        V_AkunPenerima = .Range("F15").MergeArea.Cells(1, 1).Value
        V_Jumlah = Val(.Range("H15").MergeArea.Cells(1, 1).Value)
        V_Deskripsi = .Range("D19").MergeArea.Cells(1, 1).Value
        V_DiterimaDari = .Range("H19").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL (Cek jika input kritikal kosong melompong)
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_AkunPenerima = "" Or V_Jumlah = 0 Then
        MsgBox "Form Input Belum Lengkap!" & vbCrLf & _
               "Pastikan Tanggal, No Bukti, Akun Penerima, dan Jumlah Pendapatan sudah terisi.", _
               vbExclamation, "Validasi Input Gagal"
        Exit Sub
    End If
    
    ' Kunci visual layar biar gerakan pengisian datanya super smooth secepat kilat
    Application.ScreenUpdating = False
    
    ' BUKA PROTEKSI
    LembarData.Unprotect Password:="IMAS"
    LembarHarian.Unprotect Password:="IMAS"
    
    ' --------------------------------------------------------------------
    ' 3. PROSES 1: EKSEKUSI INJEKSI DATA KE TABEL DATA PENGHASILAN LAIN
    ' --------------------------------------------------------------------
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
    
    With BarisBaru
        .Range(1) = NoUrutTerakhir + 1  ' KOLOM 1: NO
        .Range(2) = V_Tanggal           ' KOLOM 2: TANGGAL TERIMA
        .Range(3) = V_NoBukti           ' KOLOM 3: NO. BUKTI
        .Range(4) = V_Kategori          ' KOLOM 4: KATEGORI PENDAPATAN
        .Range(5) = V_DiterimaDari      ' KOLOM 5: DITERIMA DARI
        .Range(6) = V_Deskripsi         ' KOLOM 6: DESKRIPSI
        .Range(7) = V_Metode            ' KOLOM 7: METODE PEMBAYARAN
        .Range(8) = V_AkunPenerima      ' KOLOM 8: AKUN KAS / BANK
        .Range(9) = V_Jumlah            ' KOLOM 9: JUMLAH PENDAPATAN BERSIH
        
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(9).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
 ' --------------------------------------------------------------------
    ' 4. PROSES 2: EKSEKUSI INJEKSI DATA KE TABEL LAPORAN HARIAN KAS (DEBIT)
    ' --------------------------------------------------------------------
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
    
    With BarisHarian
        .Range(1) = NoUrutHarianTerakhir + 1 ' KOLOM 1: NO.
        .Range(2) = V_Tanggal                ' KOLOM 2: TANGGAL
        .Range(3) = V_AkunPenerima           ' KOLOM 3: AKUN KAS / BANK
        .Range(4) = V_Jumlah                 ' KOLOM 4: DEBIT (Uang Masuk)
        .Range(5) = 0                        ' KOLOM 5: KREDIT (Kosong / 0)
        .Range(6) = V_Deskripsi              ' KOLOM 6: DESKRIPSI TRANSAKSI
        .Range(7) = V_NoBukti                ' KOLOM 7: NO. BUKTI TRANSAKSI
        
        ' Pengkondisian Jenis Aktivitas berdasarkan Kategori Penghasilan di cell F11
        If Trim(V_Kategori) = "Pendapatan Bunga Bank" Then
            .Range(8) = "Bunga Diterima"
        Else
            .Range(8) = "Penghasilan Lain"
        End If
        
        ' Format Estetik Ruangan Baris Laporan Harian Kas
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' TUTUP PROTEKSI
    LembarData.Protect Password:="IMAS", AllowFiltering:=True
    LembarHarian.Protect Password:="IMAS", AllowFiltering:=True
    
    Application.ScreenUpdating = True
End Sub
