Attribute VB_Name = "Module76"
Sub SimpanTransaksiPelunasanUtang()
    ' ====================================================================
    ' MODUL UTAMA: SIMPAN PELUNASAN UTANG, UPDATE INDUK, & SINKRONISASI KAS
    ' SPECIAL EDITION: COMPATIBLE WITH USA REGIONAL SETTINGS (EXCEL USA)
    ' Font Segoe UI 9 Black + Custom Accounting Rp + Dynamic Induk Finder
    ' TRIPLE TARGET: "TabelPelunasanUtang", "TabelUtangUsaha/Bank", & "TabelLaporanHarianKas"
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarLunas As Worksheet
    Dim LembarHarian As Worksheet
    Dim LembarInduk As Worksheet
    
    Dim TabelLunas As ListObject
    Dim TabelHarian As ListObject
    Dim TabelInduk As ListObject
    
    Dim BarisLunas As ListRow
    Dim BarisHarian As ListRow
    
    Dim NoUrutLunas As Long
    Dim NoUrutHarian As Long
    Dim IsLunasKosong As Boolean
    Dim IsHarianKosong As Boolean
    Const PWD As String = "IMAS"
    
    ' VARIABEL DATA FORMULIR INPUT PELUNASAN
    Dim V_JenisUtang As String ' Isi G11 (Usaha / Bank)
    Dim V_NoBuktiLunas As String ' Isi G15 (BPUT-xxxx)
    Dim V_Tanggal As Date ' Isi D15
    Dim V_NoRefUtang As String ' Isi D11 (No Bukti Utang Induk yang mau dibayar)
    Dim V_NominalPokok As Double ' Isi D19
    Dim V_NominalBunga As Double ' Isi G19
    Dim V_AkunBayar As String ' Isi D23
    Dim V_StatusPelunasan As String ' Isi G23
    Dim V_Deskripsi As String ' Isi D27
    Dim V_NamaKreditur As String
    
    ' Variabel bantu untuk penyisiran tabel induk
    Dim CellCari As Range
    Dim BarisIndukKe As Long
    Dim KolomTerbayarKe As Long
    Dim KolomBungaKe As Long
    Dim NilaiLamaTerbayar As Double
    Dim NilaiLamaBunga As Double
    
    ' 1. SETTING SHEET FORMULIR & TABEL UTAMA
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT PELUNASAN")
    Set LembarLunas = ThisWorkbook.Sheets("UTANG_PELUNASAN UTANG")
    Set LembarHarian = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    
    On Error Resume Next
    Set TabelLunas = LembarLunas.ListObjects("TabelPelunasanUtang")
    Set TabelHarian = LembarHarian.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' Validasi Awal Keberadaan Tabel Utama
    If TabelLunas Is Nothing Or TabelHarian Is Nothing Then
        MsgBox "Error: 'TabelPelunasanUtang' atau 'TabelLaporanHarianKas' tidak ditemukan!", vbCritical, "Eksekusi Ditolak"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT PELUNASAN (SUPPORT MERGED CELLS)
    With LembarForm
        V_JenisUtang = Trim(.Range("G11").MergeArea.Cells(1, 1).Value)
        V_NoRefUtang = Trim(.Range("D11").MergeArea.Cells(1, 1).Value)
        V_Tanggal = .Range("D15").MergeArea.Cells(1, 1).Value
        V_NoBuktiLunas = Trim(.Range("G15").MergeArea.Cells(1, 1).Value)
        V_NominalPokok = Val(.Range("D19").MergeArea.Cells(1, 1).Value)
        V_NominalBunga = Val(.Range("G19").MergeArea.Cells(1, 1).Value)
        V_AkunBayar = .Range("D23").MergeArea.Cells(1, 1).Value
        V_StatusPelunasan = .Range("G23").MergeArea.Cells(1, 1).Value
        V_Deskripsi = .Range("D27").MergeArea.Cells(1, 1).Value
    End With
    
    ' VALIDASI INTERNAL CONTROL KETAT
    If V_Tanggal = 0 Or V_NoBuktiLunas = "" Or V_JenisUtang = "" Or V_NoRefUtang = "" Or (V_NominalPokok + V_NominalBunga) = 0 Then
        MsgBox "VALIDASI GAGAL! Input Data Pelunasan Belum Lengkap." & vbCrLf & _
               "Pastikan Jenis Utang, No Referensi, Tanggal, No Bukti, dan Nominal sudah terisi!", vbCritical, "Sistem Menolak"
        Exit Sub
    End If
    
    ' Kunci visual layar biar loading kilat laptop ASUS gak kedip
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' OPERASI BUKA GEMBOK UTAMA (Pelunasan & Laporan Kas)
    ' ====================================================================
    LembarLunas.Unprotect Password:=PWD
    LembarHarian.Unprotect Password:=PWD
    
    ' ====================================================================
    ' 3. LOGIKA DETEKTIF: MENYISIR TABEL INDUK (USAHA / BANK) BERDASARKAN SEL G11
    ' ====================================================================
    V_NamaKreditur = "Tidak Diketahui" ' Default jika tidak ketemu
    
    ' PERBAIKAN: Menggunakan huruf kecil semua ("biaya usaha") agar cocok dengan fungsi LCase
    If LCase(V_JenisUtang) = "biaya usaha" Then
        ' Target: TabelUtangUsaha di sheet DAFTAR UTANG USAHA
        Set LembarInduk = ThisWorkbook.Sheets("UTANG_DAFTAR UTANG USAHA")
        Set TabelInduk = LembarInduk.ListObjects("TabelUtangUsaha")
        
        If Not TabelInduk Is Nothing Then
            ' Buka Gembok Lembar Induk Usaha sebelum diperbarui
            LembarInduk.Unprotect Password:=PWD
            
            ' Cari No Ref Utang di Kolom ke-3 (NO BUKTI TRANSAKSI)
            Set CellCari = TabelInduk.ListColumns(3).DataBodyRange.Find(What:=V_NoRefUtang, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not CellCari Is Nothing Then
                BarisIndukKe = CellCari.Row - TabelInduk.HeaderRowRange.Row
                ' Ambil nama kreditur dari kolom ke-4
                V_NamaKreditur = TabelInduk.DataBodyRange.Cells(BarisIndukKe, 4).Value
                
                KolomTerbayarKe = 11
                
                ' Tambahkan nilai lama dengan nominal pokok yang baru dibayar (D19)
                NilaiLamaTerbayar = Val(TabelInduk.DataBodyRange.Cells(BarisIndukKe, KolomTerbayarKe).Value)
                TabelInduk.DataBodyRange.Cells(BarisIndukKe, KolomTerbayarKe).Value = NilaiLamaTerbayar + V_NominalPokok
            End If
            
            ' Kunci Kembali Lembar Induk Usaha
            LembarInduk.Protect Password:=PWD, AllowFiltering:=True
        End If
        
    ' PERBAIKAN: Menggunakan huruf kecil semua ("bank") agar cocok dengan fungsi LCase
    ElseIf LCase(V_JenisUtang) = "bank" Then
        ' Target: TabelUtangBank di sheet DAFTAR UTANG BANK
        Set LembarInduk = ThisWorkbook.Sheets("UTANG_DAFTAR UTANG BANK")
        Set TabelInduk = LembarInduk.ListObjects("TabelUtangBank")
        
        If Not TabelInduk Is Nothing Then
            ' Buka Gembok Lembar Induk Bank sebelum diperbarui
            LembarInduk.Unprotect Password:=PWD
            
            ' Cari No Ref Utang di Kolom ke-2
            Set CellCari = TabelInduk.ListColumns(2).DataBodyRange.Find(What:=V_NoRefUtang, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not CellCari Is Nothing Then
                BarisIndukKe = CellCari.Row - TabelInduk.HeaderRowRange.Row
                ' Ambil nama kreditur/bank dari kolom ke-3 (Pastikan kolom ini benar)
                V_NamaKreditur = TabelInduk.DataBodyRange.Cells(BarisIndukKe, 3).Value
                
                KolomTerbayarKe = 10
                
                ' UPDATE: Nominal pokok ditambah nominal bunga saat dimasukkan ke kolom utang terbayar
                NilaiLamaTerbayar = Val(TabelInduk.DataBodyRange.Cells(BarisIndukKe, KolomTerbayarKe).Value)
                TabelInduk.DataBodyRange.Cells(BarisIndukKe, KolomTerbayarKe).Value = NilaiLamaTerbayar + V_NominalPokok + V_NominalBunga
                
                ' LOGIKA UPDATE BANK 2: Cari kolom "NOMINAL BAYAR BUNGA & ADMIN" (atau default kolom 13)
                On Error Resume Next
                KolomBungaKe = TabelInduk.ListColumns("NOMINAL BAYAR BUNGA & ADMIN").Index
                On Error GoTo 0
                If KolomBungaKe = 0 Then KolomBungaKe = 13
                
                NilaiLamaBunga = Val(TabelInduk.DataBodyRange.Cells(BarisIndukKe, KolomBungaKe).Value)
                TabelInduk.DataBodyRange.Cells(BarisIndukKe, KolomBungaKe).Value = NilaiLamaBunga + V_NominalBunga
            End If
            
            ' Kunci Kembali Lembar Induk Bank
            LembarInduk.Protect Password:=PWD, AllowFiltering:=True
        End If
    End If
    
    ' ====================================================================
    ' 4. EKSEKUSI TARGET 1: TABEL PELUNASAN UTANG GLOBAL
    ' ====================================================================
    IsLunasKosong = False
    On Error Resume Next
    If TabelLunas.ListRows.Count = 0 Then
        IsLunasKosong = True
        NoUrutLunas = 0
    ElseIf TabelLunas.ListRows.Count = 1 And (TabelLunas.DataBodyRange.Cells(1, 1).Value = "" Or TabelLunas.DataBodyRange.Cells(1, 2).Value = "") Then
        IsLunasKosong = True
        NoUrutLunas = 0
    Else
        NoUrutLunas = Application.WorksheetFunction.Max(TabelLunas.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsLunasKosong Then
        If TabelLunas.ListRows.Count = 0 Then TabelLunas.ListRows.Add
        Set BarisLunas = TabelLunas.ListRows(1)
    Else
        Set BarisLunas = TabelLunas.ListRows.Add
    End If
    
    ' Suntik Data Sesuai Rincian Mutlak Paduka
    With BarisLunas
        .Range(1) = V_Tanggal                       ' KOLOM 2: TANGGAL (D15)
        .Range(2) = V_NoBuktiLunas                  ' KOLOM 3: NO BUKTI (G15)
        .Range(3) = V_Deskripsi                     ' KOLOM 4: DESKRIPSI (D27)
        .Range(4) = V_NoRefUtang                    ' KOLOM 5: NO. REFERENSI (D11)
        .Range(5) = V_NamaKreditur                  ' KOLOM 6: KREDITUR (Hasil Sisir Otomatis Kolom 4 Induk)
        .Range(6) = V_NominalPokok                  ' KOLOM 7: NOMINAL POKOK (D19)
        .Range(7) = V_NominalBunga                  ' KOLOM 8: NOMINAL BUNGA (G19)
        .Range(8) = V_AkunBayar                     ' KOLOM 9: AKUN BAYAR (D23)
        .Range(9) = V_StatusPelunasan               ' KOLOM 10: STATUS PELUNASAN (G23)
        
        ' Estetika Segoe UI 9 Black   PT MMD
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
        End With
        .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(10).HorizontalAlignment = xlCenter
    End With
    
   ' ====================================================================
    ' 5. EKSEKUSI TARGET 2: SINKRONISASI KE TABEL LAPORAN HARIAN KAS (DIREVISI 2 BARIS)
    ' ====================================================================
    
    ' --------------------------------------------------------------------
    ' BARIS PERTAMA: PEMBAYARAN POKOK UTANG (Hanya jika Nominal Pokok > 0)
    ' --------------------------------------------------------------------
    If V_NominalPokok > 0 Then
        IsHarianKosong = False
        On Error Resume Next
        If TabelHarian.ListRows.Count = 0 Then
            IsHarianKosong = True
            NoUrutHarian = 0
        ElseIf TabelHarian.ListRows.Count = 1 And (TabelHarian.DataBodyRange.Cells(1, 1).Value = "" Or TabelHarian.DataBodyRange.Cells(1, 2).Value = "") Then
            IsHarianKosong = True
            NoUrutHarian = 0
        Else
            NoUrutHarian = Application.WorksheetFunction.Max(TabelHarian.ListColumns(1).DataBodyRange)
        End If
        On Error GoTo 0
        
        If IsHarianKosong Then
            If TabelHarian.ListRows.Count = 0 Then TabelHarian.ListRows.Add
            Set BarisHarian = TabelHarian.ListRows(1)
        Else
            Set BarisHarian = TabelHarian.ListRows.Add
        End If
        
        ' Suntik Data Baris Pokok
        With BarisHarian
            .Range(1) = NoUrutHarian + 1                                    ' KOLOM 1: NO.
            .Range(2) = V_Tanggal                                           ' KOLOM 2: TANGGAL
            .Range(3) = V_AkunBayar                                         ' KOLOM 3: AKUN KAS / BANK
            .Range(4) = 0                                                   ' KOLOM 4: DEBIT
            .Range(5) = V_NominalPokok                                      ' KOLOM 5: KREDIT (Nominal Pokok)
            .Range(6) = "Pelunasan Pokok Utang " & V_JenisUtang & ": " & V_NamaKreditur & " (" & V_Deskripsi & ")" ' KOLOM 6: DESKRIPSI
            .Range(7) = V_NoBuktiLunas                                       ' KOLOM 7: NO. BUKTI TRANSAKSI
            .Range(8) = "Pembayaran Pokok Pinjaman"                         ' KOLOM 8: JENIS AKTIVITAS
            
            With .Range.Font
                .Name = "Segoe UI"
                .Size = 10
                .Color = vbBlack
            End With
            .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        End With
    End If
    
    ' --------------------------------------------------------------------
    ' BARIS KEDUA: PEMBAYARAN BUNGA UTANG (Hanya jika Nominal Bunga > 0)
    ' --------------------------------------------------------------------
    If V_NominalBunga > 0 Then
        IsHarianKosong = False
        On Error Resume Next
        If TabelHarian.ListRows.Count = 0 Then
            IsHarianKosong = True
            NoUrutHarian = 0
        ElseIf TabelHarian.ListRows.Count = 1 And (TabelHarian.DataBodyRange.Cells(1, 1).Value = "" Or TabelHarian.DataBodyRange.Cells(1, 2).Value = "") Then
            IsHarianKosong = True
            NoUrutHarian = 0
        Else
            NoUrutHarian = Application.WorksheetFunction.Max(TabelHarian.ListColumns(1).DataBodyRange)
        End If
        On Error GoTo 0
        
        If IsHarianKosong Then
            If TabelHarian.ListRows.Count = 0 Then TabelHarian.ListRows.Add
            Set BarisHarian = TabelHarian.ListRows(1)
        Else
            Set BarisHarian = TabelHarian.ListRows.Add
        End If
        
        ' Suntik Data Baris Bunga
        With BarisHarian
            .Range(1) = NoUrutHarian + 1                                    ' KOLOM 1: NO.
            .Range(2) = V_Tanggal                                           ' KOLOM 2: TANGGAL
            .Range(3) = V_AkunBayar                                         ' KOLOM 3: AKUN KAS / BANK
            .Range(4) = 0                                                   ' KOLOM 4: DEBIT
            .Range(5) = V_NominalBunga                                      ' KOLOM 5: KREDIT (Nominal Bunga)
            .Range(6) = "Pelunasan Bunga Utang " & V_JenisUtang & ": " & V_NamaKreditur & " (" & V_Deskripsi & ")" ' KOLOM 6: DESKRIPSI
            .Range(7) = V_NoBuktiLunas                                       ' KOLOM 7: NO. BUKTI TRANSAKSI
            .Range(8) = "Pembayaran Bunga Pinjaman"                         ' KOLOM 8: JENIS AKTIVITAS
            
            With .Range.Font
                .Name = "Segoe UI"
                .Size = 10
                .Color = vbBlack
            End With
            .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        End With
    End If
    
    ' ====================================================================
    ' OPERASI RE-LOCK SYSTEM (Kunci Kembali Pelunasan & Laporan Kas)
    ' ====================================================================
    LembarLunas.Protect Password:=PWD, AllowFiltering:=True
    LembarHarian.Protect Password:=PWD, AllowFiltering:=True
    
    ' Mengembalikan stabilitas sistem grafis Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
