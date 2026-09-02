Attribute VB_Name = "Module43"
Sub SimpanTransaksiPengeluaranUsaha()
    ' ====================================================================
    ' MODUL UTAMA: SIMPAN DATA PENGELUARAN USAHA & KONTROL BIAYA DIMUKA
    ' FIX TOTAL: Urutan Kolom & Nama Header 100% Sesuai Titah   (Kolom 9-17)
    ' TIM KOMA (,): Menggunakan .FormulaLocal dengan pemisah KOMA sesuai sistem
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarData As Worksheet
    Dim TabelData As ListObject
    Dim BarisBaru As ListRow
    Dim NoUrutTerakhir As Long
    Dim IsBarisPertamaKosong As Boolean
    Const PWD As String = "IMAS"
    
    ' VARIABEL DATA FORMULIR
    Dim V_Kategori As String
    Dim V_NoBukti As String
    Dim V_TanggalBayar As Date
    Dim V_DibayarKepada As String
    Dim V_Nominal As Double
    Dim V_BayarDariAkun As String
    Dim V_Status As String
    Dim V_TglMulai As Variant
    Dim V_Deskripsi As String
    Dim V_TglSelesai As Variant
    Dim V_AktivitasSAK As String
    
    ' 1. SETTING SHEET & TABEL TARGET
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT DATA")
    Set LembarData = ThisWorkbook.Sheets("PENGELUARAN USAHA_DAFTAR")
    
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelBebanUsaha")
    On Error GoTo 0
    
    If TabelData Is Nothing Then
        MsgBox "Error: Tabel 'TabelBebanUsaha' tidak ditemukan! " & vbCrLf & _
               "Pastikan nama tabel di sheet PENGELUARAN USAHA_DAFTAR sudah benar.", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DATA FORMULIR (SUPPORT MERGED CELLS)
    With LembarForm
        V_Kategori = .Range("D12").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("G12").MergeArea.Cells(1, 1).Value
        V_TanggalBayar = .Range("D16").MergeArea.Cells(1, 1).Value
        V_DibayarKepada = .Range("G16").MergeArea.Cells(1, 1).Value
        V_Nominal = Val(.Range("D20").MergeArea.Cells(1, 1).Value)
        V_BayarDariAkun = .Range("G20").MergeArea.Cells(1, 1).Value
        V_Status = .Range("D24").MergeArea.Cells(1, 1).Value
        V_TglMulai = .Range("G24").MergeArea.Cells(1, 1).Value
        V_Deskripsi = .Range("D28").MergeArea.Cells(1, 1).Value
        V_TglSelesai = .Range("G28").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL
    If V_TanggalBayar = 0 Or V_NoBukti = "" Or V_Kategori = "" Or V_Nominal = 0 Then
        MsgBox "Form Input Belum Lengkap,  !" & vbCrLf & _
               "Pastikan Tanggal Bayar, No Bukti, Kategori Beban, dan Nominal sudah terisi.", _
               vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 3. LOGIKA DETEKSI BARIS PERTAMA KOSONG & HITUNG NOMOR URUT
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
    
    ' Kunci visual layar
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Buka gembok sheet daftar beban usaha
    LembarData.Unprotect Password:=PWD
    
    ' 4. EKSEKUSI PENAMBAHAN BARIS DATA
    If IsBarisPertamaKosong Then
        If TabelData.ListRows.Count = 0 Then TabelData.ListRows.Add
        Set BarisBaru = TabelData.ListRows(1)
    Else
        Set BarisBaru = TabelData.ListRows.Add
    End If
    
    ' ====================================================================
    ' PEMETAAN URUTAN KOLOM REVISI TOTAL (100% AMAN & SESUAI REKOMENDASI  )
    ' ====================================================================
    Dim cNo As Long: cNo = 1
    Dim cNoBukti As Long: cNoBukti = 2
    Dim cTglBayar As Long: cTglBayar = 3
    Dim cDeskripsi As Long: cDeskripsi = 4
    Dim cKategori As Long: cKategori = 5
    Dim cStatus As Long: cStatus = 6
    Dim cBayarDari As Long: cBayarDari = 7
    Dim cDibayarKe As Long: cDibayarKe = 8
    
    ' --- REVISI STRUKTUR KOLOM 9 SAMPAI 17 ---
    Dim cNominal As Long: cNominal = 9
    Dim cBebanDiakui As Long: cBebanDiakui = 10
    Dim cTglMulai As Long: cTglMulai = 11
    Dim cTglSelesai As Long: cTglSelesai = 12
    Dim cDurasi As Long: cDurasi = 13
    Dim cMasaHabis As Long: cMasaHabis = 14
    Dim cSisaMasa As Long: cSisaMasa = 15
    Dim cBebanPerBulan As Long: cBebanPerBulan = 16
    Dim cSisaSaldo As Long: cSisaSaldo = 17
    
    ' 5. PROSES INJECT DATA BERDASARKAN REVISI STRUKTUR BARU
    With BarisBaru
        ' Kolom 1 s.d 8
        .Range(cNo) = NoUrutTerakhir + 1
        .Range(cNoBukti) = V_NoBukti
        .Range(cTglBayar) = V_TanggalBayar
        .Range(cDeskripsi) = V_Deskripsi
        .Range(cKategori) = V_Kategori
        .Range(cStatus) = V_Status
        .Range(cBayarDari) = V_BayarDariAkun
        .Range(cDibayarKe) = V_DibayarKepada
        
        ' Kolom 9: NOMINAL BAYAR
        .Range(cNominal) = V_Nominal
        
        ' Kolom 11 & 12: PENGKONDISIAN TANGGAL MULAI & SELESAI PENGGUNAAN BEBAN
        If Trim(UCase(V_Status)) = "LANGSUNG HABIS" Then
            .Range(cTglMulai) = "-"
            .Range(cTglSelesai) = "-"
        Else
            If IsDate(V_TglMulai) Then .Range(cTglMulai) = V_TglMulai Else .Range(cTglMulai) = "-"
            If IsDate(V_TglSelesai) Then .Range(cTglSelesai) = V_TglSelesai Else .Range(cTglSelesai) = "-"
        End If
        
        ' Kolom 14: MASA BIAYA YANG HABIS (Total Masa dikurangi Sisa Masa)
        .Range(cMasaHabis).FormulaLocal = "=[@[TOTAL MASA BIAYA (BULAN)]]-[@[SISA MASA BIAYA DIBAYAR DIMUKA]]"
        
        ' ====================================================================
        ' EMPIRE FORMULA TIM KOMA (,) - INJECT FORMULA BERDASARKAN STRUKTUR
        ' ====================================================================
        
        ' Kolom 10: BEBAN DIAKUI
        .Range(cBebanDiakui).FormulaLocal = "=IF([@[STATUS PENGGUNAAN]]=""LANGSUNG HABIS"", [@[NOMINAL BAYAR]], [@[NOMINAL BAYAR]]-[@[SISA SALDO (ASET)]])"
        
        ' Kolom 13: TOTAL MASA BIAYA (BULAN)
        .Range(cDurasi).FormulaLocal = "=IF(OR([@[TANGGAL MULAI PENGGUNAAN BEBAN]]=""-"",[@[TANGGAL SELESAI PENGGUNAAN BEBAN]]=""-"" ), 0, DATEDIF([@[TANGGAL MULAI PENGGUNAAN BEBAN]], [@[TANGGAL SELESAI PENGGUNAAN BEBAN]], ""M""))"
        
        ' Kolom 15: SISA MASA BIAYA DIBAYAR DIMUKA
        .Range(cSisaMasa).FormulaLocal = "=IF([@[TOTAL MASA BIAYA (BULAN)]]=0, 0, MAX(0, DATEDIF(TODAY(), [@[TANGGAL SELESAI PENGGUNAAN BEBAN]], ""M"") + 1))"
        
        ' Kolom 16: BEBAN PER-BULAN
        .Range(cBebanPerBulan).FormulaLocal = "=IF([@[TOTAL MASA BIAYA (BULAN)]]=0, 0, [@[NOMINAL BAYAR]]/[@[TOTAL MASA BIAYA (BULAN)]])"
        
        ' Kolom 17: SISA SALDO (ASET)
        .Range(cSisaSaldo).FormulaLocal = "=IFERROR([@[BEBAN PER-BULAN]]*[@[SISA MASA BIAYA DIBAYAR DIMUKA]], 0)"
        
    End With
    
    ' Kunci kembali sheet daftar beban usaha
    LembarData.Protect Password:=PWD, AllowFiltering:=True
            
    ' ====================================================================
    ' AMBIL ALIH DAN KONDISIKAN TARGET SHEET LAPORAN HARIAN KAS
    ' ====================================================================
    Set LembarData = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    Set TabelData = LembarData.ListObjects("TabelLaporanHarianKas")
    
    ' Buka gembok sheet laporan harian kas
    LembarData.Unprotect Password:=PWD
    
    Select Case Trim(V_Kategori)
        Case "Biaya Bunga Bank"
            V_AktivitasSAK = "Pembayaran Bunga Pinjaman"
        Case "Biaya Peralatan Kantor"
            V_AktivitasSAK = "Pembelian Aset"
        Case "Beban Gaji & Upah Kantor"
            V_AktivitasSAK = "Pembayaran Pegawai"
        Case Else
            V_AktivitasSAK = "Pembayaran ke Vendor"
    End Select
    
    IsBarisPertamaKosong = False
    If TabelData.ListRows.Count = 0 Then
        IsBarisPertamaKosong = True
        NoUrutTerakhir = 0
    ElseIf TabelData.ListRows.Count = 1 And (TabelData.DataBodyRange.Cells(1, 1).Value = "" Or TabelData.DataBodyRange.Cells(1, 2).Value = "") Then
        IsBarisPertamaKosong = True
        NoUrutTerakhir = 0
    Else
        NoUrutTerakhir = Application.WorksheetFunction.Max(TabelData.ListColumns(1).DataBodyRange)
    End If
    
    If IsBarisPertamaKosong Then
        If TabelData.ListRows.Count = 0 Then TabelData.ListRows.Add
        Set BarisBaru = TabelData.ListRows(1)
    Else
        Set BarisBaru = TabelData.ListRows.Add
    End If
    
    With BarisBaru
        .Range(1) = NoUrutTerakhir + 1
        .Range(2) = V_TanggalBayar
        .Range(3) = V_BayarDariAkun
        .Range(4) = 0
        .Range(5) = V_Nominal
        .Range(6) = V_Deskripsi
        .Range(7) = V_NoBukti
        .Range(8) = V_AktivitasSAK
    End With
    
    ' Kunci kembali sheet laporan harian kas
    LembarData.Protect Password:=PWD, AllowFiltering:=True
    
    ' Mengembalikan stabilitas sistem grafis Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

