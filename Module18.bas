Attribute VB_Name = "Module18"
Sub SimpanTransaksiKasMasuk()
    ' ==========================================================
    
    Dim SheetInput As Worksheet
    Dim SheetLaporan As Worksheet
    Dim TabelLaporan As ListObject
    Dim BarisBaru As ListRow
    Dim NoUrut As Long
    
    ' 1. PENGATURAN RUJUKAN SHEET & TABEL EXCEL
    Set SheetInput = ThisWorkbook.Sheets("KAS&BANK_INPUT TRANSAKSI UMUM")
    Set SheetLaporan = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelLaporan = SheetLaporan.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' BUKA PROTEKSI
    SheetLaporan.Unprotect Password:="IMAS"
    
    ' 2. LOGIKA OTOMATISASI NOMOR URUT
    On Error Resume Next
    If TabelLaporan.ListRows.Count = 0 Then
        NoUrut = 1
    Else
        NoUrut = Application.WorksheetFunction.Max(TabelLaporan.ListColumns(1).DataBodyRange) + 1
    End If
    If NoUrut = 0 Then NoUrut = 1
    On Error GoTo 0
    
    ' 3. EKSEKUSI: TAMBAH BARIS BARU DI PALING BAWAH
    Set BarisBaru = TabelLaporan.ListRows.Add
    
    ' 4. PROSES INJECT DATA KE BARIS BARU
    With BarisBaru
        .Range(1) = NoUrut                                  ' KOLOM 1: NO
        .Range(2) = SheetInput.Range("E13").Value           ' KOLOM 2: TANGGAL
        .Range(3) = SheetInput.Range("H25").Value           ' KOLOM 3: AKUN KAS DAN BANK
        .Range(4) = SheetInput.Range("E17").Value           ' KOLOM 4: DEBIT
        .Range(5) = 0                                       ' KOLOM 5: KREDIT
        .Range(6) = SheetInput.Range("E21").Value           ' KOLOM 6: DESKRIPSI
        .Range(7) = SheetInput.Range("H13").Value           ' KOLOM 7: NO. BUKTI
        .Range(8) = "Modal Disetor"                             ' KOLOM 8: JENIS AKTIVITAS
    End With
    
    ' TUTUP PROTEKSI
    SheetLaporan.Protect Password:="IMAS", AllowFiltering:=True
    
End Sub
