Attribute VB_Name = "Module113"
Sub JurnalUmumInputPelunasanPiutang()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT PELUNASAN PIUTANG KE JURNAL UMUM
    ' Status: Sistem Buka-Tutup Gembok Otomatis Hanya untuk JURNAL UMUM (Password: IMAS)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject, tblCOA As ListObject
    Dim wsInput As Worksheet, wsJurnal As Worksheet
    
    Set wsInput = Sheets("PIUTANG_INPUT DATA PIUTANG")
    Set wsJurnal = ThisWorkbook.Sheets("JURNAL UMUM")
    
    Const PWD As String = "IMAS" ' <-- Sandi sakral dari Ratu
    
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Utama
    If tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: 'TabelJurnalUmum' atau 'TabelDataCOA' tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. Tarik Data dari Form Input Pelunasan (Mendukung Merged Cells Sesuai Detail Ratu)
    Dim V_NoBukti As String: V_NoBukti = wsInput.Range("G17").MergeArea.Cells(1, 1).Value
    Dim V_Tanggal As Date: V_Tanggal = wsInput.Range("E17").MergeArea.Cells(1, 1).Value
    Dim V_Nominal As Double: V_Nominal = Val(wsInput.Range("E21").MergeArea.Cells(1, 1).Value)
    Dim V_DibayarDari As String: V_DibayarDari = Trim(wsInput.Range("E25").MergeArea.Cells(1, 1).Value) ' Akun Kas/Bank Penerima
    Dim V_DeskripsiUser As String: V_DeskripsiUser = Trim(wsInput.Range("E29").MergeArea.Cells(1, 1).Value)
    
    ' Validasi Input Dasar
    If V_NoBukti = "" Or V_Tanggal = 0 Or V_Nominal <= 0 Or V_DibayarDari = "" Then
        MsgBox "Gagal! Pastikan No. Bukti, Tanggal, Nominal, dan kolom Dibayar Dari sudah terisi.", vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 3. LOGIKA PENETAPAN NAMA AKUN KREDIT & DEBIT
    Dim namaAkunDebit As String: namaAkunDebit = V_DibayarDari     ' Kas/Bank yang menerima uang
    Dim namaAkunKredit As String: namaAkunKredit = "Piutang Usaha" ' Akun piutang yang dilunasi
    
    ' 4. Ambil Semua Kode Akun dari Master COA
    Dim barisCOADebit As Variant, barisCOAKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisCOADebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisCOAKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisCOADebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisCOAKredit)
    On Error GoTo 0
    
    ' Proteksi validasi COA
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        MsgBox "Gagal Menjurnal! Pastikan nama akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' sudah terdaftar di COA!", vbCritical, "Error COA"
        Exit Sub
    End If
    
    ' ====================================================================
    ' ?? OPERASI JEBOL GEMBOK: Buka proteksi sheet JURNAL UMUM sebelum menulis
    ' ====================================================================
    wsJurnal.Unprotect Password:=PWD
    
    ' 5. EKSEKUSI EXPORT KE JURNAL UMUM (SISTEM ANTI-KEDIP TURBO)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' Menyusun Deskripsi Jurnal Gabungan Sesuai Request Ratu
    Dim V_DeskripsiJurnal As String
    V_DeskripsiJurnal = "Pelunasan Piutang Usaha" & IIf(V_DeskripsiUser <> "", " - " & V_DeskripsiUser, "")
    
    ' --- BARIS 1: DEBIT (Kas/Bank Bertambah di Sisi Kiri) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiJurnal
        .Range(4) = kodeAkunDebit
        .Range(5) = namaAkunDebit
        .Range(6) = V_Nominal
        .Range(7) = 0
        .Range(8) = "Kas & Bank"
        .Range(9) = "Tidak"
    End With
    
    ' --- BARIS 2: KREDIT (Piutang Usaha Berkurang di Sisi Kanan) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiJurnal
        .Range(4) = kodeAkunKredit
        .Range(5) = namaAkunKredit
        .Range(6) = 0
        .Range(7) = V_Nominal
        .Range(8) = "Piutang Usaha"
        .Range(9) = "Tidak"
    End With
    
    ' 6. FORM CLEANER (Otomatis Bersih Kilat)
    With wsInput
        On Error Resume Next
        .Range("G17").MergeArea.ClearContents   ' No. Bukti
        .Range("E17").MergeArea.ClearContents   ' Tanggal Transaksi
        .Range("E21").MergeArea.ClearContents   ' Nominal
        .Range("G21").MergeArea.ClearContents   ' Dibayar Dari
        .Range("E29").MergeArea.ClearContents   ' Deskripsi
        .Range("E13").MergeArea.ClearContents
        .Range("E25").MergeArea.ClearContents
        .Range("G25").MergeArea.ClearContents
        On Error GoTo 0
        
        ' Kembalikan Kursor ke Hulu Form (G17)
        .Activate
        .Range("G17").MergeArea.Cells(1, 1).Select
    End With

NyalakanSistem:
    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Kunci kembali lembar kerja JURNAL UMUM
    ' ====================================================================
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
    ' 7. NOTIFIKASI BERHASIL
    MsgBox "Data Pelunasan Piutang Berhasil Disimpan!", vbInformation, "Sistem Sukses"
End Sub
