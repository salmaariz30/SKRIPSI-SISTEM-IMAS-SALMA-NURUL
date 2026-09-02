Attribute VB_Name = "Module73"
Sub SimpanTransaksiUtangUsaha()
    ' ====================================================================
    ' MODUL UTAMA: SINKRONISASI UTANG USAHA & BIAYA PRODUKSI (REVISI  )
    ' SPECIAL EDITION: COMPATIBLE WITH USA REGIONAL SETTINGS (EXCEL USA)
    ' Proteksi Baris Pertama + Font Segoe UI 9 Black + Custom Accounting Rp
    ' DUAL TARGET MUTLAK: "TabelUtangUsaha" & "TabelBiayaProduksi" (No Buku Kas)
    ' DESIGN EDITION: Anti-Formula Injection & Structured Reference Safety
    ' FIX FIX FIX: Meluruskan Posisi Tanggal & No. Bukti di Tabel Biaya
    ' STATUS: AUTO UNPROTECT / PROTECT SHEET TARGET (PASSWORD: IMAS)
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarUtang As Worksheet
    Dim LembarBiaya As Worksheet
    
    Dim TabelUtang As ListObject
    Dim TabelBiaya As ListObject
    
    Dim barisUtang As ListRow
    Dim BarisBiaya As ListRow
    
    Dim NoUrutUtang As Long
    Dim NoUrutBiaya As Long
    
    Dim IsUtangKosong As Boolean
    Dim IsBiayaKosong As Boolean
    Const PWD As String = "IMAS"
    
    ' VARIABEL DATA FORMULIR INPUT UTANG
    Dim V_Tanggal As Date
    Dim V_NoBukti As String
    Dim V_NamaKreditur As String
    Dim V_Deskripsi As String
    Dim V_PosPengeluaran As String
    Dim V_TOP As Long
    Dim V_TglJatuhTempo As Date
    Dim V_NominalPokok As Double
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT USAHA")
    Set LembarUtang = ThisWorkbook.Sheets("UTANG_DAFTAR UTANG USAHA")
    
    On Error Resume Next
    Set TabelUtang = LembarUtang.ListObjects("TabelUtangUsaha")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel
    If TabelUtang Is Nothing Then
        MsgBox "Error: 'TabelUtangUsaha' tidak ditemukan!", vbCritical, "Eksekusi Gagal"
        Exit Sub
    End If
    
    ' 2. AMBIL DATA DARI FORMULIR INPUT UTANG (SUPPORT MERGED CELLS DENGAN AMAN)
    With LembarForm
        V_Tanggal = .Range("D12").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("H12").MergeArea.Cells(1, 1).Value
        V_NamaKreditur = Trim(.Range("D16").MergeArea.Cells(1, 1).Value)
        
        ' REVISI UTAMA  : Deskripsi mutlak diambil murni dari D28 (Menggantikan D18 lama)
        V_Deskripsi = .Range("H24").MergeArea.Cells(1, 1).Value
        
        V_PosPengeluaran = .Range("H16").MergeArea.Cells(1, 1).Value
        V_TOP = Val(.Range("D20").MergeArea.Cells(1, 1).Value)
        V_TglJatuhTempo = .Range("H20").MergeArea.Cells(1, 1).Value
        V_NominalPokok = Val(.Range("D24").MergeArea.Cells(1, 1).Value)
    End With
    
    ' VALIDASI KETAT INTERNAL CONTROL
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_NamaKreditur = "" Or V_NominalPokok = 0 Then
        MsgBox "VALIDASI GAGAL! [Input Cacat Data]" & vbCrLf & _
               "Tanggal, No Bukti, Nama Kreditur, dan Nominal Pokok Hutang Wajib Diisi!", vbCritical, "Sistem Menolak"
        Exit Sub
    End If
    
    ' Kunci visual layar biar loading nge-looping kilatnya tidak bikin kedip
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' ?? OPERASI BUKA GEMBOK: Jebol proteksi sheet target utang sebelum menulis
    ' ====================================================================
    LembarUtang.Unprotect Password:=PWD
    
    ' ====================================================================
    ' 3. EKSEKUSI DATABASE TARGET 1: TABEL UTANG USAHA
    ' ====================================================================
    IsUtangKosong = False
    On Error Resume Next
    If TabelUtang.ListRows.Count = 0 Then
        IsUtangKosong = True
        NoUrutUtang = 0
    ElseIf TabelUtang.ListRows.Count = 1 And (TabelUtang.DataBodyRange.Cells(1, 1).Value = "" Or TabelUtang.DataBodyRange.Cells(1, 2).Value = "") Then
        IsUtangKosong = True
        NoUrutUtang = 0
    Else
        NoUrutUtang = Application.WorksheetFunction.Max(TabelUtang.ListColumns(1).DataBodyRange)
    End If
    On Error GoTo 0
    
    If IsUtangKosong Then
        If TabelUtang.ListRows.Count = 0 Then TabelUtang.ListRows.Add
        Set barisUtang = TabelUtang.ListRows(1)
    Else
        Set barisUtang = TabelUtang.ListRows.Add
    End If
    
    With barisUtang
        .Range(1) = NoUrutUtang + 1                  ' KOLOM 1: NOMOR
        .Range(2) = V_Tanggal                       ' KOLOM 2: TANGGAL
        .Range(3) = V_NoBukti                       ' KOLOM 3: NO BUKTI TRANSAKSI
        .Range(4) = V_NamaKreditur                  ' KOLOM 4: NAMA KREDITUR
        .Range(5) = V_Deskripsi                     ' KOLOM 5: DESKRIPSI TRANSAKSI (Murni Sel D28)
        .Range(6) = V_PosPengeluaran                ' KOLOM 6: POS PENGELUARAN
        .Range(7) = V_TOP                           ' KOLOM 7: TOP
        .Range(8) = V_TglJatuhTempo                 ' KOLOM 8: TANGGAL JATUH TEMPO
        
        ' KOLOM 9: RUMUS UMUR UTANG SMART CAPSLOCK
        .Range(9).Formula2R1C1 = "=IF([@[SISA UTANG]]=0, 0, TODAY()-[@[TANGGAL JATUH TEMPO]])"
        
        .Range(10) = V_NominalPokok                 ' KOLOM 10: NOMINAL POKOK HUTANG
        .Range(11) = "0"
        
        ' KOLOM 12: RUMUS SISA UTANG
        .Range(12).Formula2R1C1 = "=[@[NOMINAL HUTANG]]-[@[TERBAYAR]]"
        
        ' KOLOM 13: RUMUS STATUS PELUNASAN
        .Range(13).Formula2R1C1 = "=IF([@[SISA UTANG]]=0, ""Lunas"", ""Belum Lunas"")"
        
        ' Estetika Khas     (Segoe UI 9 Black)
        With .Range.Font
            .Name = "Segoe UI"
            .Size = 10
            .Color = vbBlack
            .Bold = False
        End With
        
        ' Format Nomor & Keuangan Rp
        .Range(7).NumberFormat = "#,##0"
        .Range(9).NumberFormat = "#,##0"
        .Range(10).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(11).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(12).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(13).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(14).HorizontalAlignment = xlCenter
    End With
    
    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Kunci kembali lembar target setelah selesai
    ' ====================================================================
    LembarUtang.Protect Password:=PWD, AllowFiltering:=True
    
    ' Nyalakan kembali sistem grafis Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

