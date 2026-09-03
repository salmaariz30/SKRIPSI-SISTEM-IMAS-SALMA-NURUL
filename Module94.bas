Attribute VB_Name = "Module94"
Function SimpanSaldoAwalPiutang() As Boolean
    ' ====================================================================
    ' MODUL SAVING DATA SALDO AWAL PIUTANG KE BUKU PIUTANG (VERSI KORPORAT)
    ' Menggunakan Pola Injeksi Dinamis Berbasis Formula R1C1 & Anti-Lag
    ' ====================================================================
    
    ' Default status di awal adalah Gagal (False)
    SimpanSaldoAwalPiutang = False
    
    ' 1. MATIKAN OPTIMASI SISTEM DI PALING ATAS (Anti-Lemot & Kedip)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim LembarForm As Worksheet, LembarPiutang As Worksheet
    Dim TabelPiutang As ListObject
    Dim BarisPiutang As ListRow
    Dim IsBarisPiutangKosong As Boolean
    Dim NoUrutPiutang As Long
    Const PWD As String = "IMAS"
    
    ' Set Lembar Kerja sesuai Request
    Set LembarForm = Sheets("PIUTANG_INPUT SALDO AWAL")
    
    On Error Resume Next
    Set LembarPiutang = Sheets("PIUTANG_BUKU PIUTANG")
    Set TabelPiutang = LembarPiutang.ListObjects("TabelBukuPiutang")
    On Error GoTo 0
    
    ' Pengaman Utama: Jika tabel target tidak ditemukan
    If LembarPiutang Is Nothing Or TabelPiutang Is Nothing Then
        MsgBox "ALARM PIUTANG! 'TabelBukuPiutang' di sheet 'PIUTANG_BUKU PIUTANG' gagal dimuat.", vbCritical, "Sistem Patah Kaki"
        GoTo ResetSistem
    End If
    
    ' 2. AMBIL DATA DARI KOORDINAT FORM INPUT   (Lengkap dengan Merge Area Protection)
    Dim V_NoBukti As Variant: V_NoBukti = Trim(LembarForm.Range("D11").MergeArea.Cells(1, 1).Value)
    Dim V_Tanggal As Variant: V_Tanggal = LembarForm.Range("F11").MergeArea.Cells(1, 1).Value
    Dim V_NamaPelanggan As Variant: V_NamaPelanggan = Trim(LembarForm.Range("D15").MergeArea.Cells(1, 1).Value)
    Dim V_TglJatuhTempo As Variant: V_TglJatuhTempo = LembarForm.Range("F15").MergeArea.Cells(1, 1).Value
    Dim V_TotalPenjualan As Variant: V_TotalPenjualan = LembarForm.Range("D19").MergeArea.Cells(1, 1).Value
    Dim V_PotonganRetur As Variant: V_PotonganRetur = LembarForm.Range("F19").MergeArea.Cells(1, 1).Value
    Dim V_TotalTerbayar As Variant: V_TotalTerbayar = LembarForm.Range("D23").MergeArea.Cells(1, 1).Value
    
    ' 3. VALIDASI WAJIB ISI (Mencegah Staf Menyimpan Data Kosong/Cacat)
    If V_NoBukti = "" Or V_Tanggal = 0 Or V_NamaPelanggan = "" Or V_TglJatuhTempo = 0 Then
        MsgBox "Mohon lengkapi No. Invoice, Tanggal, Nama Pelanggan, dan Tanggal Jatuh Tempo!", vbExclamation, "Data Form Belum Lengkap"
        GoTo ResetSistem
    End If
    
    ' ====================================================================
    ' ?? BUKA GEMBOK TARGET SEBELUM INJEKSI DATA
    ' ====================================================================
    LembarPiutang.Unprotect Password:=PWD
    
    ' 4. LOGIKA PERHITUNGAN NO URUT OTOMATIS (Sama Persis dengan Jalur Rekayasa  )
    IsBarisPiutangKosong = False
    On Error Resume Next
    If TabelPiutang.ListRows.Count = 0 Then
        IsBarisPiutangKosong = True
        NoUrutPiutang = 0
    ElseIf TabelPiutang.ListRows.Count = 1 And (TabelPiutang.DataBodyRange.Cells(1, 1).Value = "" Or TabelPiutang.DataBodyRange.Cells(1, 2).Value = "") Then
        IsBarisPiutangKosong = True
        NoUrutPiutang = 0
    Else
        NoUrutPiutang = Application.WorksheetFunction.Max(TabelPiutang.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    ' Penentuan pembuatan baris baru di tabel database
    If IsBarisPiutangKosong Then
        If TabelPiutang.ListRows.Count = 0 Then TabelPiutang.ListRows.Add
        Set BarisPiutang = TabelPiutang.ListRows(1)
    Else
        Set BarisPiutang = TabelPiutang.ListRows.Add
    End If
    
    ' 5. PROSES INJEKSI DATA KE TABEL BUKU PIUTANG (100% Mengikuti Urutan Kolom  )
    With BarisPiutang
        .Range(1) = V_NoBukti                                               ' KOLOM 2: NO. INVOICE
        .Range(2) = CDate(V_Tanggal)                                        ' KOLOM 3: TANGGAL
        .Range(3) = V_NamaPelanggan                                         ' KOLOM 4: NAMA PELANGGAN
        .Range(4) = Val(V_TotalPenjualan)                                   ' KOLOM 5: TOTAL PENJUALAN
        .Range(5) = Val(V_PotonganRetur)                                    ' KOLOM 6: POTONGAN & RETUR
        .Range(6).Formula2R1C1 = "=[@[TOTAL PENJUALAN]]-[@[POTONGAN PENJUALAN & RETUR]]" ' KOLOM 7: TOTAL PIUTANG (NETTO)
        .Range(7) = Val(V_TotalTerbayar)                                    ' KOLOM 8: JUMLAH TERBAYAR
        .Range(8).Formula2R1C1 = "=[@[TOTAL PIUTANG (NETTO)]]-[@[JUMLAH TERBAYAR]]" ' KOLOM 9: SISA PIUTANG
        .Range(9) = CDate(V_TglJatuhTempo)                                  ' KOLOM 10: TANGGAL JATUH TEMPO
        
        ' KOLOM 11: RUMUS UMUR PIUTANG SMART (Hentikan jika sisa piutang Rp 0)
        .Range(10).Formula2R1C1 = "=IF([@[SISA PIUTANG]]=0, ""-"", TODAY()-[@[TANGGAL JATUH TEMPO]])"
        
        ' KOLOM 12: RUMUS AGING BUCKET LOKAL PT MMD
        .Range(11).Formula2R1C1 = "=IF([@[UMUR PIUTANG]]=""-"", ""-"", IF([@[UMUR PIUTANG]]<=0, ""Belum Jatuh Tempo"", IF([@[UMUR PIUTANG]]<=30, ""1-30 Hari"", IF([@[UMUR PIUTANG]]<=60, ""31-60 Hari"", "">60 Hari""))))"
        
        ' KOLOM 13: RUMUS STATUS PELUNASAN
        .Range(12).Formula2R1C1 = "=IF([@[SISA PIUTANG]]=0, ""Lunas"", ""Belum Lunas"")"
        
        ' ----------------------------------------------------------------
        ' FORMATTING STYLE (Segoe UI 9, Anti-Bold & Format Rupiah Akuntansi)
        ' ----------------------------------------------------------------
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
            .Bold = False
        End With
        
        .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(6).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        
        .Range(11).NumberFormat = "#,##0"         ' Format Hari
        .Range(13).HorizontalAlignment = xlCenter ' Center Status
    End With
    
    ' Jika seluruh langkah berhasil dilewati tanpa terlempar ke ResetSistem, ubah status jadi Sukses (True)
    SimpanSaldoAwalPiutang = True
    
ResetSistem:
    ' ====================================================================
    ' ?? KUNCI KEMBALI SHEET TARGET
    ' ====================================================================
    If Not LembarPiutang Is Nothing Then
        LembarPiutang.Protect Password:=PWD, AllowFiltering:=True
    End If

    ' KEMBALIKAN KONDISI EXCEL KE NORMAL
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function
