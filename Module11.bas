Attribute VB_Name = "Module11"
' ====================================================================
' 1. FUNGSI UTAMA: GENERATE SINGKATAN JENIS ASET (GetSingkatan)
' ====================================================================
Function GetSingkatan(jenis As String) As String
    ' Memetakan Jenis Aset dari drop-down menjadi 3 digit kode unik (Anti-Bentrok)
    Select Case Trim(jenis)
        Case "Tanah": GetSingkatan = "TNH"
        Case "Gedung Kantor, Pabrik, dan Gudang Produksi": GetSingkatan = "GDG"
        Case "Alat Pengangkut": GetSingkatan = "ANG"
        Case "Kendaraan Roda Empat atau Lebih": GetSingkatan = "K4"
        Case "Kendaraan Roda Dua": GetSingkatan = "K2"
        Case "Furnitur Besi": GetSingkatan = "FTB"
        Case "Furnitur Kayu": GetSingkatan = "FTK"
        Case "Mesin Produksi": GetSingkatan = "MSN"
        Case "Alat Elektronik Kantor & Gadget": GetSingkatan = "ELK"
        Case "Alat Elektronik Besar (AC, TV, Kulkas, dll)": GetSingkatan = "ELB"
        Case "Peralatan Kecil": GetSingkatan = "PTN"
        Case Else: GetSingkatan = "AST" ' Kode cadangan jika tidak terdaftar
    End Select
End Function


' ====================================================================
' 2. PROSEDUR UTAMA: SIMPAN DATA KE TABEL DEPRESIASI (ANTI-BLOCK LOCK)
' ====================================================================
Sub SimpanKeDepresiasi()
    Dim wsInput As Worksheet: Set wsInput = Sheets("INPUT DATA NEW")
    Dim wsDepr As Worksheet: Set wsDepr = Sheets("DATA DEPRESIASI")
    Dim tblDepr As ListObject: Set tblDepr = wsDepr.ListObjects("TabelPenyusutan")
    Dim newRow As ListRow
    Dim TargetRange As Range
    Dim namaAset As String, jenisAset As String, singkatanJenis As String
    Dim passwordSheet As String
    ' --- VARIABEL TAMBAHAN UNTUK LAPORAN KAS ---
    Dim wsKas As Worksheet, tblKas As ListObject
    Dim TotalUangKeluar As Double, EmptyCheck As Boolean
    Dim V_Metode As String, V_Tanggal As Variant, V_NoBukti As String
    Dim V_AkunKasBank As String, V_NamaAset As String
    
    ' --- SET PASSWORD SHEET   DI SINI ---
    passwordSheet = "IMAS"
    
    ' ====================================================================
    ' SENSOR PENGAMAN: VALIDASI DATA KOSONG
    ' ====================================================================
    If Trim(wsInput.Range("C12").Value) = "" Then
        MsgBox "Data belum lengkap! [Nama Aset Tetap] Harap Dilengkapi ya!", vbExclamation, "IMAS Pengaman Sistem"
        wsInput.Activate: wsInput.Range("C12").Select
        Exit Sub
    End If
    
    If Trim(wsInput.Range("C22").Value) = "" Or Not IsNumeric(wsInput.Range("C22").Value) Then
        MsgBox "[Harga Perolehan Aset] belum dilengkapi", vbExclamation, "IMAS Pengaman Sistem"
        wsInput.Activate: wsInput.Range("C22").Select
        Exit Sub
    End If
    
    If Trim(wsInput.Range("I12").Value) = "" Then
        MsgBox "Tanggal Pembelian Aset] Harap Dilengkapi!", vbExclamation, "IMAS Pengaman Sistem"
        wsInput.Activate: wsInput.Range("I12").Select
        Exit Sub
    End If
    
    ' ====================================================================
    ' ?? OPERASI PENYELAMATAN: BUKA LOCK SHEET SEBELUM MACRO BERAKSI
    ' ====================================================================
    wsDepr.Unprotect Password:=passwordSheet
    ' Buka gembok Laporan Kas Harian
    On Error Resume Next
    Set wsKas = Sheets("KAS&BANK_LAPORAN HARIAN")
    Set tblKas = wsKas.ListObjects("TabelLaporanHarianKas")
    wsKas.Unprotect Password:=passwordSheet
    On Error GoTo 0
    
    ' AKTIVASI MODE SENYAP (ANTI-KEDIP)
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' Ambil data jenis aset untuk singkatan
    jenisAset = wsInput.Range("C17").Value
    singkatanJenis = GetSingkatan(jenisAset)
    
    ' --- CARI BARIS KOSONG ATAU TAMBAH BARIS BARU (SEKARANG AMAN) ---
    If tblDepr.DataBodyRange Is Nothing Then
        Set newRow = tblDepr.ListRows.Add
        Set TargetRange = newRow.Range
    ElseIf tblDepr.ListRows.Count = 1 And tblDepr.DataBodyRange.Cells(1, 1).Value = "" Then
        Set TargetRange = tblDepr.ListRows(1).Range
    Else
        Dim i As Long
        Dim foundEmpty As Boolean: foundEmpty = False
        
        For i = 1 To tblDepr.ListRows.Count
            If tblDepr.DataBodyRange.Cells(i, 1).Value = "" Then
                Set TargetRange = tblDepr.ListRows(i).Range
                foundEmpty = True
                Exit For
            End If
        Next i
        
        If Not foundEmpty Then
            Set newRow = tblDepr.ListRows.Add
            Set TargetRange = newRow.Range
        End If
    End If

    ' --- TEMPEL DATA EXPOR & FORMULA ---
    With TargetRange
        .Cells(1, 1).Value = wsInput.Range("I27").Value   ' Column 1: KODE
        .Cells(1, 2).Value = wsInput.Range("C12").Value   ' Column 2: NAMA ASET
        .Cells(1, 3).Value = wsInput.Range("F27").Value   ' Column 3: LOKASI ASET
        .Cells(1, 4).Value = wsInput.Range("C22").Value   ' Column 4: HARGA
        .Cells(1, 5).Value = wsInput.Range("I12").Value   ' Column 5: TANGGAL BELI
        .Cells(1, 7).Value = wsInput.Range("C17").Value   ' Column 7: JENIS ASET
        .Cells(1, 8).Value = wsInput.Range("F17").Value   ' Column 8: KATEGORI ASET
        .Cells(1, 9).Value = wsInput.Range("I17").Value   ' Column 9: UMUR EKONOMIS
        .Cells(1, 11).Value = wsInput.Range("C27").Value  ' Column 11: NILAI RESIDU
        
        ' --- SUNTIKAN RUMUS OTOMATIS ---
        .Cells(1, 6).Formula2Local = "=IF(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", ""-"", IFERROR(EDATE([@[Tanggal Pembelian]], VALUE(LEFT(TRIM([@[Umur Ekonomis]]), SEARCH("" "", TRIM([@[Umur Ekonomis]]))-1))*12), """"))"
        .Cells(1, 10).Formula2Local = "=IFERROR(IFS([@[Umur Ekonomis]]=""4 Tahun"",25%,[@[Umur Ekonomis]]=""8 Tahun"",12.5%,[@[Umur Ekonomis]]=""16 Tahun"",6.25%,[@[Umur Ekonomis]]=""20 Tahun"",5%,[@[Umur Ekonomis]]=""Non-Depresiasi"",""-""),"""")"
        .Cells(1, 12).Formula2Local = "=IFERROR(IF(OR([@[Umur Ekonomis]]="""", ISBLANK([@[Umur Ekonomis]])), """", IF(OR(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", TRIM([@[Umur Ekonomis]])=""Tak Terhingga""), ""Tak Terhingga"", LET(total_bulan_awal, LEFT([@[Umur Ekonomis]], FIND("" "", [@[Umur Ekonomis]])-1) * 12, bulan_terpakai, ROUND((EOMONTH(TODAY(),0) - EOMONTH([@[Tanggal Pembelian]],0)) / 30.4375, 0) + IF(DAY([@[Tanggal Pembelian]])<=15, 1, 0), sisa_total_bulan, MAX(0, total_bulan_awal - bulan_terpakai), sisa_tahun, QUOTIENT(sisa_total_bulan, 12), sisa_bulan, MOD(sisa_total_bulan, 12), sisa_tahun & "" tahun "" & sisa_bulan & "" bulan""))), """")"
        .Cells(1, 13).Formula2Local = "=IF(OR([@[Umur Ekonomis]]="""", ISBLANK([@[Umur Ekonomis]])), """", IF(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", 0, IFERROR(([@[Harga Perolehan]] - [@[Nilai Residu]]) / (VALUE(LEFT(TRIM([@[Umur Ekonomis]]), SEARCH("" "", TRIM([@[Umur Ekonomis]]))-1)) * 12), 0)))"
        .Cells(1, 14).Formula2Local = "=IF(OR([@[Umur Ekonomis]]="""", ISBLANK([@[Umur Ekonomis]])), """", IF(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", 0, IFERROR(LET(harga, [@[Harga Perolehan]], residu, [@[Nilai Residu]], tgl_beli, [@[Tanggal Pembelian]], biaya_bln, [@[Penyusutan Per-Bulan]], tgl_mulai, IF(DAY(tgl_beli) <= 15, DATE(YEAR(tgl_beli), MONTH(tgl_beli), 1), DATE(YEAR(tgl_beli), MONTH(tgl_beli) + 1, 1)), tgl_cut_off, DATE(2025, 12, 31), jml_bln, IF(tgl_cut_off < tgl_mulai, 0, DATEDIF(tgl_mulai, tgl_cut_off, ""m"") + 1), ROUND(MIN(harga - residu, jml_bln * biaya_bln), 0)), 0)))"
        .Cells(1, 15).Formula2Local = "=IF(OR([@[Umur Ekonomis]]="""",ISBLANK([@[Umur Ekonomis]])), """", IF(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", 0, IFERROR(LET(biaya_bln, [@[Penyusutan Per-Bulan]], harga, [@[Harga Perolehan]], residu, [@[Nilai Residu]], akum_lalu, [@[Akumulasi s.d Tahun Lalu]], tgl_mulai_aset, IF(DAY([@[Tanggal Pembelian]]) <= 15, DATE(YEAR([@[Tanggal Pembelian]]), MONTH([@[Tanggal Pembelian]]), 1), DATE(YEAR([@[Tanggal Pembelian]]), MONTH([@[Tanggal Pembelian]]) + 1, 1)), tgl_awal_2026, DATE(2026, 1, 1), tgl_sekarang, TODAY(), titik_awal, MAX(tgl_mulai_aset, tgl_awal_2026), jml_bln_2026, ROUND((EOMONTH(tgl_sekarang, 0) - EOMONTH(titik_awal, 0)) / 30.4375, 0) + 1, plafon_sisa, MAX(0, harga - residu - akum_lalu), ROUND(MIN(plafon_sisa, jml_bln_2026 * biaya_bln), 0)), 0)))"
        .Cells(1, 16).Formula2Local = "=IF(OR([@[Umur Ekonomis]]="""", ISBLANK([@[Umur Ekonomis]])), """", IF(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", 0, IFERROR([@[Akumulasi s.d Tahun Lalu]] + [@[Beban Depresiasi Tahun Ini]], 0)))"
        .Cells(1, 17).Formula2Local = "=IFERROR(IF(TRIM([@[Umur Ekonomis]])=""Non-Depresiasi"", [@[Harga Perolehan]], [@[Harga Perolehan]] - [@[Total Akumulasi Penyusutan]]), """")"
    End With
    
    ' ====================================================================
    ' ?? TARGET 3: TABEL LAPORAN HARIAN KAS (MODIFIKASI KAS MASUK/KELUAR ASET)
    ' ====================================================================
    If Not tblKas Is Nothing Then
        ' Ambil Parameter dari Form Input Aset
        V_Metode = wsInput.Range("F12").Value        ' Misal: Pembelian Tunai / Pembelian Kredit
        V_Tanggal = wsInput.Range("I12").Value       ' Tanggal Perolehan
        V_NoBukti = wsInput.Range("I27").Value       ' Kode Aset dijadikan No Bukti Kas
        V_AkunKasBank = wsInput.Range("F22").Value   ' Nama Akun Kas/Bank
        V_NamaAset = wsInput.Range("C12").Value      ' Nama Aset Tetap
        
        ' Tentukan nominal uang keluar (Tunai = Full Harga, Kredit = Nominal DP)
        TotalUangKeluar = 0
        If V_Metode = "Pembelian Tunai" Then
            TotalUangKeluar = wsInput.Range("C22").Value
        ElseIf V_Metode = "Pembelian Kredit" Then
            TotalUangKeluar = wsInput.Range("I22").Value ' <-- GANTI I22 JIKA CELL DP DI TEMPAT LAIN
        End If
        
        ' Eksekusi pengisian ke Kas jika ada aliran kas keluar nyata
        If TotalUangKeluar > 0 Then
            Dim NoUrutKas As Long, BarisKas As ListRow
            
            On Error Resume Next
            EmptyCheck = (tblKas.ListRows.Count = 0 Or (tblKas.ListRows.Count = 1 And tblKas.DataBodyRange.Cells(1, 1).Value = ""))
            
            If EmptyCheck Then
                NoUrutKas = 0
                If tblKas.ListRows.Count = 0 Then tblKas.ListRows.Add
                Set BarisKas = tblKas.ListRows(1)
            Else
                NoUrutKas = Application.WorksheetFunction.Max(tblKas.ListColumns(1).DataBodyRange)
                Set BarisKas = tblKas.ListRows.Add(AlwaysInsert:=False)
            End If
            On Error GoTo 0
            
            With BarisKas
                .Range(1) = NoUrutKas + 1
                .Range(2) = V_Tanggal
                .Range(3) = V_AkunKasBank
                .Range(4) = 0                                       ' Debit = 0
                .Range(5) = TotalUangKeluar                         ' Kredit = Uang Keluar
                .Range(6) = "Pembelian Aset Tetap: " & V_NamaAset   ' Keterangan Laporan Kas
                .Range(7) = V_NoBukti
                .Range(8) = "Pembelian Aset Tetap"                  ' Kelompok Aktivitas
                
                With .Range.Font: .Name = "Segoe UI": .Size = 10: .Color = vbBlack: End With
                .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
                .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            End With
        End If
    End If
    ' ====================================================================
    ' ?? TARGET 4: AUTO REFRESH DASHBOARD / PIVOT (VERSI BANGUNKAN OTOK EXCEL)
    ' ====================================================================
    Dim wsPivot As Worksheet
    On Error Resume Next
    Set wsPivot = Sheets("ASET TETAP")
    On Error GoTo 0
    
    If Not wsPivot Is Nothing Then
        ' ?? Buka proteksi sheet Pivot
        wsPivot.Unprotect Password:=passwordSheet
        
        ' ?? CRITICAL STEP: Paksa Excel hidupkan kalkulasi otomatis dulu!
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        
        ' ?? Jalankan refresh total se-workbook secara paksa!
        ThisWorkbook.RefreshAll
        
        ' ?? Kunci kembali sheet Pivot-nya demi keamanan negara
        wsPivot.Protect Password:=passwordSheet, AllowFiltering:=True
    End If
    ' ====================================================================
    ' ?? OPERASI PENGUNCIAN KEMBALI: SHEET DI-LOCK LAGI SECARA OTOMATIS
    ' ====================================================================
    wsDepr.Protect Password:=passwordSheet, AllowFiltering:=True
    
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With

    
End Sub
