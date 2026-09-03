Attribute VB_Name = "Module100"
Sub SimpanPelepasanKeTabel()
    Dim wsInput As Worksheet
    Dim wsDaftar As Worksheet
    Dim wsDepr As Worksheet
    Dim tbl As ListObject
    Dim tblPenyusutan As ListObject
    Dim i As Long, j As Long
    Dim TargetRow As Long
    Dim namaAsetTarget As String
    Const PWD As String = "IMAS" ' <-- Sandi resmi dari Ratu
    
    ' Variabel Tambahan untuk Laporan Kas Harian Sesuai Contoh Ratu
    Dim tblHarian As ListObject
    Dim wsKas As Worksheet
    Dim BarisKas As ListRow
    Dim NoUrutKas As Long
    Dim EmptyCheck As Boolean
    Dim V_HargaJual As Double
    Dim V_Tanggal As Variant
    Dim V_NoBukti As Variant
    
    ' 1. Set Sheet dan Tabel Target
    Set wsInput = ThisWorkbook.Sheets("ASET TETAP_INPUT PELEPASAN")
    Set wsDaftar = ThisWorkbook.Sheets("ASET TETAP_DAFTAR PELEPASAN")
    Set tbl = wsDaftar.ListObjects("TabelPelepasanAset")
    
    ' Set Tabel Penyusutan & Sheet-nya secara aman
    On Error Resume Next
    Set tblPenyusutan = Range("TabelPenyusutan").ListObject
    Set wsDepr = tblPenyusutan.Parent
    
    ' ?? MENYISIR SHEET UNTUK MENEMUKAN TABEL LAPORAN KAS HARIAN
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.ListObjects.Count > 0 Then
            Set tblHarian = Nothing
            Set tblHarian = Sh.ListObjects("TabelLaporanHarianKas")
            If Not tblHarian Is Nothing Then
                Set wsKas = Sh
                Exit For
            End If
        End If
    Next Sh
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Kas Kritikal
    If tblHarian Is Nothing Then
        MsgBox "Error: 'TabelLaporanHarianKas' tidak ditemukan di workbook ini!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' Ambil data dasar dari form input
    V_Tanggal = wsInput.Range("D10").Value
    V_NoBukti = wsInput.Range("F10").Value
    namaAsetTarget = wsInput.Range("D14").Value
    V_HargaJual = Val(wsInput.Range("D18").Value)

    ' Validation: Pastikan input nama aset tidak kosong
    If namaAsetTarget = "" Then
        MsgBox "Error: Nama Aset di cell D14 kosong! Transaksi dibatalkan.", vbCritical, "Input Kosong"
        Exit Sub
    End If

    ' ====================================================================
    ' ?? OPERASI OPEN LOCK: Jebol pengaman sheet sebelum makro beraksi
    ' ====================================================================
    wsDaftar.Unprotect Password:=PWD
    If Not wsDepr Is Nothing Then wsDepr.Unprotect Password:=PWD
    If Not wsKas Is Nothing Then wsKas.Unprotect Password:=PWD

    ' AKTIVASI MODE SENYAP (ANTI-KEDIP + AKTIFKAN ROTASI PENUH)
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

    ' 2. STRATEGI AMAN: CARI BARIS YANG NAMA ASETNYA KOSONG (Kolom 4 di TabelPelepasanAset)
    TargetRow = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 4).Value = "" Then
            TargetRow = i
            Exit For
        End If
    Next i

    If TargetRow = 0 Then
        tbl.ListRows.Add
        TargetRow = tbl.ListRows.Count
    End If

    ' 3. PENGISIAN DATA SESUAI STRUKTUR FORM INPUT PELEPASAN
    With tbl.ListRows(TargetRow)
        .Range(1) = V_Tanggal                      ' KOLOM 1: TANGGAL
        .Range(2) = V_NoBukti                      ' KOLOM 2: NO. BUKTI
        .Range(3) = wsInput.Range("H10").Value      ' KOLOM 3: JENIS PELEPASAN
        .Range(4) = namaAsetTarget                  ' KOLOM 4: NAMA ASET TETAP
        .Range(5) = wsInput.Range("F14").Value      ' KOLOM 5: HARGA PEROLEHAN AWAL
        .Range(6) = wsInput.Range("H14").Value      ' KOLOM 6: NILAI BUKU
        .Range(7) = V_HargaJual                     ' KOLOM 7: HARGA JUAL
        .Range(8) = wsInput.Range("F18").Value      ' KOLOM 8: KEUNTUNGAN (KERUGIAN ASET)
        .Range(9) = wsInput.Range("H18").Value      ' KOLOM 9: AKUN KAS / BANK
    End With

    ' ====================================================================
    ' ?? POTONGAN KODE BARU: INPUT TRANSAKSI KE TABEL LAPORAN KAS HARIAN (FIX)
    ' ====================================================================
    ' Hanya mencatat jika ada nominal uang masuk (Harga Jual > 0)
    If V_HargaJual > 0 Then
        On Error Resume Next
        EmptyCheck = (tblHarian.ListRows.Count = 0 Or (tblHarian.ListRows.Count = 1 And tblHarian.DataBodyRange.Cells(1, 1).Value = ""))
        
        If EmptyCheck Then
            NoUrutKas = 0
            If tblHarian.ListRows.Count = 0 Then tblHarian.ListRows.Add
            Set BarisKas = tblHarian.ListRows(1)
        Else
            NoUrutKas = Application.WorksheetFunction.Max(tblHarian.ListColumns(1).DataBodyRange)
            Set BarisKas = tblHarian.ListRows.Add(AlwaysInsert:=False)
        End If
        On Error GoTo 0
        
        ' Pengisian data kas masuk mengikuti urutan contoh Ratu 100% presisi
        With BarisKas
            .Range(1) = NoUrutKas + 1                                       ' KOLOM 1: NO URUT AUTO
            .Range(2) = V_Tanggal                                           ' KOLOM 2: TANGGAL
            .Range(3) = wsInput.Range("H18").Value                          ' KOLOM 3: AKUN KAS/BANK
            .Range(4) = V_HargaJual                                         ' KOLOM 4: DEBIT / MASUK
            .Range(5) = 0                                                   ' KOLOM 5: KREDIT / KELUAR
            .Range(6) = "Pelepasan Aset: " & namaAsetTarget                 ' KOLOM 6: KETERANGAN
            .Range(7) = V_NoBukti                                           ' KOLOM 7: NO BUKTI
            .Range(8) = "Untung Rugi Pelepasan Aset"                                         ' KOLOM 8: JENIS AKTIVITAS
            
            ' Kosmetik & Font standarisasi Ratu
            With .Range.Font: .Name = "Segoe UI": .Size = 10: .Color = vbBlack: End With
            .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        End With
    End If

    ' ====================================================================
    ' 4. MENYISIR & MENGHAPUS BARIS DI TABEL PENYUSUTAN (DARI BAWAH)
    ' ====================================================================
    For j = tblPenyusutan.ListRows.Count To 1 Step -1
        ' Cek kolom ke-2 (Nama Aset) pada TabelPenyusutan
        If tblPenyusutan.DataBodyRange(j, 2).Value = namaAsetTarget Then
            tblPenyusutan.ListRows(j).Delete
            Exit For
        End If
    Next j

    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Pasang kembali satpam proteksi sheet Ratu
    ' ====================================================================
    wsDaftar.Protect Password:=PWD, AllowFiltering:=True
    If Not wsDepr Is Nothing Then wsDepr.Protect Password:=PWD, AllowFiltering:=True
    If Not wsKas Is Nothing Then wsKas.Protect Password:=PWD, AllowFiltering:=True

    ' ====================================================================
    ' ?? TARGET TAMBAHAN: AUTO REFRESH DASHBOARD (BANGUNKAN OTODIDAK EXCEL)
    ' ====================================================================
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    
    ' Paksa seluruh Pivot Table berputar memperbarui diri
    ThisWorkbook.RefreshAll

End Sub
