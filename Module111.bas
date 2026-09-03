Attribute VB_Name = "Module111"
Sub JurnalUmumInputSaldoAwalPiutang()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT SALDO AWAL PIUTANG KE JURNAL UMUM
    ' Status: Sistem Buka-Tutup Gembok Otomatis Hanya untuk JURNAL UMUM (Password: IMAS)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("PIUTANG_INPUT SALDO AWAL")
    Dim wsJurnal As Worksheet: Set wsJurnal = ThisWorkbook.Sheets("JURNAL UMUM")
    
    Const PWD As String = "IMAS" ' <-- Sandi sakral dari Ratu
    
    ' 2. Tarik Data dari Cell Input Saldo Awal Piutang (Mendukung Merged Cells)
    Dim V_Tanggal As Variant: V_Tanggal = Date ' Otomatis Hari Ini Sesuai Request Ratu
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("D11").Value
    Dim V_NamaPelangan As String: V_NamaPelangan = Trim(wsInput.Range("D15").Value)
    
    ' Tarik Nominal Uang untuk Perhitungan Netto Piutang
    Dim V_NominalPiutang As Double: V_NominalPiutang = Val(wsInput.Range("D19").Value)
    Dim V_PotonganRetur As Double: V_PotonganRetur = Val(wsInput.Range("F19").Value)
    Dim V_TotalTerbayar As Double: V_TotalTerbayar = Val(wsInput.Range("D23").Value)
    
    ' RUMUS AKUNTANSI: Piutang Bersih yang Belum Tertagih
    Dim V_PiutangNetto As Double
    V_PiutangNetto = V_NominalPiutang - V_PotonganRetur - V_TotalTerbayar
    
    ' Membuat Deskripsi Otomatis Sesuai Request
    Dim V_Deskripsi As String
    V_Deskripsi = "Saldo Awal Piutang: " & V_NamaPelangan
    
    ' 3. Validasi Pengaman Awal (Anti-Kosong)
    If V_NoBukti = "" Or V_NamaPelangan = "" Or V_NominalPiutang = 0 Then
        MsgBox "VALIDASI INPUT GAGAL!" & vbCrLf & _
               "Mohon pastikan No Bukti, Nama Pelanggan, dan Nominal Piutang telah terisi.", _
               vbExclamation, "Data Belum Lengkap"
        Exit Sub
    End If
    
    ' Validasi jika hasil perhitungan piutang netto minus atau nol
    If V_PiutangNetto <= 0 Then
        MsgBox "VALIDASI ERROR!" & vbCrLf & _
               "Total Piutang Netto bernilai 0 atau minus setelah dikurangi Potongan/Terbayar!", _
               vbCritical, "Nilai Piutang Tidak Valid"
        Exit Sub
    End If
    
    ' 4. LOGIKA penetapan nama akun tetap
    Dim namaAkunDebit As String: namaAkunDebit = "Piutang Usaha"
    Dim namaAkunKredit As String: namaAkunKredit = "Ekuitas - Saldo Awal"
    
    ' 5. Ambil Kode Akun (Kolom 1) Berdasarkan Nama Akun (Kolom 2) di COA secara Otomatis
    Dim barisDebit As Variant, barisKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisDebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
    On Error GoTo 0
    
    ' Proteksi jika master COA Ratu belum mendaftarkan akun di atas
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        MsgBox "Gagal menjurnal! Akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' tidak ditemukan di TabelDataCOA!", _
               vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' ====================================================================
    ' ?? OPERASI JEBOL GEMBOK: Buka proteksi sheet JURNAL UMUM
    ' ====================================================================
    wsJurnal.Unprotect Password:=PWD
    
    ' 6. EKSEKUSI EXPOR DATA (ANTI-KEDIP TURBO)
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' --- A. EKSEKUSI JURNAL UMUM (BARIS DEBIT & KREDIT) ---
    ' Baris Debit (Piutang Usaha)
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_Deskripsi: .Range(4) = kodeAkunDebit
        .Range(5) = namaAkunDebit: .Range(6) = V_PiutangNetto: .Range(7) = 0: .Range(8) = "Piutang Usaha": .Range(9) = "Tidak"
    End With
    
    ' Baris Kredit (Ekuitas Saldo Awal)
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_Deskripsi: .Range(4) = kodeAkunKredit
        .Range(5) = namaAkunKredit: .Range(6) = 0: .Range(7) = V_PiutangNetto: .Range(8) = "Ekuitas": .Range(9) = "Tidak"
    End With
    
    ' 7. SILENT CLEAN FORM (Otomatis Bersih Kilat Menggunakan wsInput)
    With wsInput
        On Error Resume Next
        .Range("D11").MergeArea.ClearContents
        .Range("F11").MergeArea.ClearContents
        .Range("D15").MergeArea.ClearContents
        .Range("F15").MergeArea.ClearContents
        .Range("D19").MergeArea.ClearContents
        .Range("F19").MergeArea.ClearContents
        .Range("D23").MergeArea.ClearContents
        On Error GoTo 0
        
        ' Kembalikan kursor ke posisi awal biar nyaman input lagi
        .Activate
        .Range("D11").MergeArea.Cells(1, 1).Select
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
    
    ' 8. NOTIFIKASI SUKSES
    MsgBox "Saldo Awal Piutang atas nama '" & V_NamaPelangan & "' Berhasil Disimpan!", vbInformation, "Penyimpanan Sukses"
End Sub

