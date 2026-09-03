Attribute VB_Name = "Module101"
Sub SimpanTransaksiUtangBank()
    ' ====================================================================
    ' STATUS: AUTO UNPROTECT / PROTECT TARGETS (PASSWORD: IMAS)
    ' TARGET MULTI-TABLE: TabelUtangBank & TabelLaporanHarianKas
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarData As Worksheet
    Dim LembarKas As Worksheet
    
    Dim TabelData As ListObject
    Dim TabelKas As ListObject
    
    Dim BarisBaru As ListRow
    Dim BarisKas As ListRow
    
    Dim NoUrutTerakhir As Long
    Dim NoUrutKas As Long
    
    Dim IsBarisPertamaKosong As Boolean
    Dim IsKasKosong As Boolean
    Const PWD As String = "IMAS"
    
    ' VARIABEL DATA FORMULIR INPUT UTANG BANK
    Dim V_Tanggal As Date
    Dim V_NoBukti As String
    Dim V_NamaBank As String
    Dim V_JenisKredit As String
    Dim V_Plafon As Double
    Dim V_Tenor As Long
    Dim V_SukuBunga As Double
    Dim V_TglJatuhTempo As Date
    Dim V_Angsuran As Double
    Dim V_NominalTerbayar As Double
    Dim V_StatusUtang As String
    Dim V_PilihanKasBank As String
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT BANK")
    Set LembarData = ThisWorkbook.Sheets("UTANG_DAFTAR UTANG BANK")
    Set LembarKas = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelData = LembarData.ListObjects("TabelUtangBank")
    Set TabelKas = LembarKas.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel
    If TabelData Is Nothing Or TabelKas Is Nothing Then
        MsgBox "Error: 'TabelUtangBank' atau 'TabelLaporanHarianKas' tidak ditemukan!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT UTANG BANK (SUPPORT MERGED CELLS)
    With LembarForm
        V_Tanggal = .Range("D12").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("H12").MergeArea.Cells(1, 1).Value
        V_NamaBank = .Range("D16").MergeArea.Cells(1, 1).Value
        V_JenisKredit = .Range("H16").MergeArea.Cells(1, 1).Value
        V_Plafon = Val(.Range("D20").MergeArea.Cells(1, 1).Value)
        V_Tenor = Val(.Range("H20").MergeArea.Cells(1, 1).Value)
        V_SukuBunga = Val(.Range("D24").MergeArea.Cells(1, 1).Value)
        V_Angsuran = Val(.Range("H24").MergeArea.Cells(1, 1).Value)
        V_TglJatuhTempo = .Range("D28").MergeArea.Cells(1, 1).Value
        V_PilihanKasBank = .Range("H28").MergeArea.Cells(1, 1).Value
        V_NominalTerbayar = Val(.Range("H32 ").MergeArea.Cells(1, 1).Value)
        V_StatusUtang = "Belum Lunas" ' Default status mutlak sesuai rincian
    End With
    
    ' VALIDASI INTERNAL CONTROL (Cek jika ada input kritikal yang masih kosong)
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_NamaBank = "" Or V_Plafon = 0 Or V_Tenor = 0 Then
        MsgBox "Form Input Utang Bank Belum Lengkap!" & vbCrLf & _
               "Pastikan Tanggal, No Bukti, Nama Bank, Plafon, dan Tenor sudah terisi.", _
               vbExclamation, "Validasi Input Gagal"
        Exit Sub
    End If
    
    ' Kunci visual layar biar gerakan pengisian datanya super smooth
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' ?? OPERASI BUKA GEMBOK: Daftar Utang Bank & Laporan Harian Kas
    ' ====================================================================
    LembarData.Unprotect Password:=PWD
    LembarKas.Unprotect Password:=PWD
    
    ' ====================================================================
    ' 3. LOGIKA PROTEKSI BARIS PERTAMA TABEL UTANG BANK
    ' ====================================================================
    IsBarisPertamaKosong = False
    On Error Resume Next
    If TabelData.ListRows.Count = 0 Then
        IsBarisPertamaKosong = True
    ElseIf TabelData.ListRows.Count = 1 And (TabelData.DataBodyRange.Cells(1, 1).Value = "" Or TabelData.DataBodyRange.Cells(1, 2).Value = "") Then
        IsBarisPertamaKosong = True
    End If
    On Error GoTo 0
    
    If IsBarisPertamaKosong Then
        If TabelData.ListRows.Count = 0 Then TabelData.ListRows.Add
        Set BarisBaru = TabelData.ListRows(1)
    Else
        Set BarisBaru = TabelData.ListRows.Add
    End If
    
    ' ====================================================================
    ' 4. INJECT DATA KE TABEL DAFTAR UTANG BANK
    ' ====================================================================
    With BarisBaru
        .Range(1) = V_Tanggal           ' KOLOM 1: TANGGAL
        .Range(2) = V_NoBukti           ' KOLOM 2: NO. BUKTI
        .Range(3) = V_NamaBank          ' KOLOM 3: NAMA BANK
        .Range(4) = V_JenisKredit       ' KOLOM 4: JENIS KREDIT
        .Range(5) = V_Plafon            ' KOLOM 5: PLAFON
        .Range(6) = V_Tenor             ' KOLOM 6: TENOR
        .Range(7) = V_SukuBunga         ' KOLOM 7: SUKU BUNGA
        .Range(8) = V_TglJatuhTempo     ' KOLOM 8: TANGGAL JATUH TEMPO
        .Range(9) = V_Angsuran          ' KOLOM 9: ANGSURAN PER BULAN
        .Range(10) = V_NominalTerbayar  ' KOLOM 10: NOMINAL TERBAYAR
        .Range(11) = V_StatusUtang      ' KOLOM 11: STATUS UTANG
        
        ' Format Estetik Ruangan Baris Baru (Segoe UI 10 Black)
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
            .Bold = False
        End With
        
        ' Format Khusus Kolom Angka, Persen, dan Tenor Bulan
        .Range(1).NumberFormat = "dd/mm/yyyy"
        .Range(8).NumberFormat = "dd/mm/yyyy"
        .Range(6).NumberFormat = "#,##0"" Bulan"""
        .Range(7).NumberFormat = "0.0%"
        .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(9).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(10).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' ====================================================================
    ' 5. LOGIKA NO URUT & INJECT DATA KE LAPORAN HARIAN KAS
    ' ====================================================================
    IsKasKosong = False
    On Error Resume Next
    If TabelKas.ListRows.Count = 0 Then
        IsKasKosong = True
        NoUrutKas = 0
    ElseIf TabelKas.ListRows.Count = 1 And (TabelKas.DataBodyRange.Cells(1, 1).Value = "" Or TabelKas.DataBodyRange.Cells(1, 2).Value = "") Then
        IsKasKosong = True
        NoUrutKas = 0
    Else
        NoUrutKas = Application.WorksheetFunction.Max(TabelKas.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsKasKosong Then
        If TabelKas.ListRows.Count = 0 Then TabelKas.ListRows.Add
        Set BarisKas = TabelKas.ListRows(1)
    Else
        Set BarisKas = TabelKas.ListRows.Add
    End If
    
    ' Eksekusi Pengisian Data Kas (Hanya jika bukan migrasi saldo awal)
    If UCase(Trim(V_PilihanKasBank)) <> "SALDO AWAL" Then
        With BarisKas
            .Range(1) = NoUrutKas + 1                                       ' KOLOM 1: NO
            .Range(2) = V_Tanggal                                           ' KOLOM 2: TANGGAL
            .Range(3) = V_PilihanKasBank                                    ' KOLOM 3: AKUN KAS DAN BANK (H28)
            .Range(4) = V_Plafon                                            ' KOLOM 4: DEBIT (Plafon Masuk Kas)
            .Range(5) = 0                                                   ' KOLOM 5: KREDIT
            .Range(6) = "Penerimaan Pencairan Kredit Utang Bank " & V_NamaBank ' KOLOM 6: DESKRIPSI
            .Range(7) = V_NoBukti                                           ' KOLOM 7: NO. BUKTI
            .Range(8) = "Pendanaan Utang Bank"                                        ' KOLOM 8: JENIS AKTIVITAS
            
            ' Penyelarasan Font Kas
            With .Range.Font
                .Name = "Segoe UI"
                .Size = 10
                .Color = vbBlack
                .Bold = False
            End With
            
            ' Format Angka Kas
            .Range(2).NumberFormat = "dd/mm/yyyy"
            .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        End With
    End If
    
    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Kunci kembali seluruh lembar target
    ' ====================================================================
    LembarData.Protect Password:=PWD, AllowFiltering:=True
    LembarKas.Protect Password:=PWD, AllowFiltering:=True
    
    ' Nyalakan kembali sistem visual Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

