Attribute VB_Name = "Module1"
Sub JurnalUmumInputPelepasanAset()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS PELEPASAN ASET TETAP (ANTI LAG & ANTI GEMBOK)
    ' VERSI FINAL: LANGSUNG TEMBAK SHEET "JURNAL UMUM" & UNPROTECT OTOMATIS
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim tblPenyusutan As ListObject: Set tblPenyusutan = Range("TabelPenyusutan").ListObject
    
    Dim wsInput As Worksheet: Set wsInput = Sheets("ASET TETAP_INPUT PELEPASAN")
    Dim wsJurnal As Worksheet: Set wsJurnal = ThisWorkbook.Sheets("JURNAL UMUM")
    
    Const PWD As String = "IMAS" ' <-- Kunci gembok keramat
    
    ' Validasi Keberadaan Tabel Utama
    If tblJurnal Is Nothing Or tblCOA Is Nothing Or tblPenyusutan Is Nothing Then
        MsgBox "Error: Tabel Jurnal, COA, atau Tabel Penyusutan tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. Ambil Data dari Form Input Pelepasan Aset
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("D10").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("F10").MergeArea.Cells(1, 1).Value
    Dim V_JenisPelepasan As String: V_JenisPelepasan = wsInput.Range("H10").MergeArea.Cells(1, 1).Value
    Dim V_NamaAset As String: V_NamaAset = Trim(wsInput.Range("D14").MergeArea.Cells(1, 1).Value)
    
    Dim V_HargaPerolehan As Double: V_HargaPerolehan = Val(wsInput.Range("F14").MergeArea.Cells(1, 1).Value)
    Dim V_NilaiBuku As Double: V_NilaiBuku = Val(wsInput.Range("H14").MergeArea.Cells(1, 1).Value)
    Dim V_HargaJual As Double: V_HargaJual = Val(wsInput.Range("D18").MergeArea.Cells(1, 1).Value)
    Dim V_LabaRugi As Double: V_LabaRugi = Val(wsInput.Range("F18").MergeArea.Cells(1, 1).Value)
    Dim V_AkunKasBank As String: V_AkunKasBank = wsInput.Range("H18").MergeArea.Cells(1, 1).Value
    
    Dim V_KeteranganSeragam As String: V_KeteranganSeragam = "Pelepasan Aset: " & V_NamaAset
    Dim V_AkumulasiPenyusutan As Double: V_AkumulasiPenyusutan = V_HargaPerolehan - V_NilaiBuku
    
    ' VALIDASI INTERNAL CONTROL
    Dim Tx_NamaAset As String: Tx_NamaAset = Trim(V_NamaAset)
    If V_Tanggal = 0 Or V_NoBukti = "" Or Tx_NamaAset = "" Then
        MsgBox "Form Input Pelepasan Belum Lengkap! Periksa Tanggal, No Bukti, dan Nama Aset.", vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' ====================================================================
    ' ?? OPERASI JEBOL GEMBOK: Buka proteksi sheet Jurnal Umum sebelum diisi
    ' ====================================================================
    wsJurnal.Unprotect Password:=PWD

    ' KUNCI SISTEM UTAMA BIAR SUPER CEPAT & ANTI KUNING
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' ====================================================================
    ' 3. JURUS TURBO MATCH: TEMBAK MURNI NOMINAL TEXT DROPDOWN KE KOLOM 2
    ' ====================================================================
    Dim barisAset As Variant
    Dim kelompokAset As String: kelompokAset = ""
    
    ' Reset error bypass
    On Error Resume Next
    
    ' Mencari V_NamaAset murni di Kolom 2 (Nama Aset) TabelPenyusutan
    barisAset = Application.Match(V_NamaAset, tblPenyusutan.ListColumns(2).DataBodyRange, 0)
    
    If Not IsError(barisAset) Then
        ' Kalau ketemu, ambil Kelompok Aset dari Kolom 8
        kelompokAset = Trim(Application.Index(tblPenyusutan.ListColumns(8).DataBodyRange, barisAset))
    End If
    On Error GoTo 0
    
    ' VALIDASI ANTI GAIB: Jika tidak ketemu di Kolom 2, langsung stop
    If kelompokAset = "" Then
        MsgBox "Error: Aset '" & V_NamaAset & "' tidak ditemukan di Kolom 2 TabelPenyusutan!" & vbCrLf & _
               "Pastikan data Nama Aset di master data sama persis dengan dropdown.", vbCritical, "Aset Tidak Ada"
        GoTo NyalakanSistem
    End If
    
    ' 4. LOGIKA PENETAPAN NAMA AKUN PASANGAN
    Dim namaAkunAset As String: namaAkunAset = kelompokAset
    Dim namaAkunAkumulasi As String
    
    Select Case Trim(kelompokAset)
        Case "Bangunan / Pabrik":           namaAkunAkumulasi = "Akum. Penyusutan Bangunan"
        Case "Kendaraan":                  namaAkunAkumulasi = "Akum. Penyusutan Kendaraan"
        Case "Peralatan & Mesin":           namaAkunAkumulasi = "Akum. Penyusutan Peralatan"
        Case Else:                          namaAkunAkumulasi = "Akum. Penyusutan " & kelompokAset
    End Select
    
    Dim namaAkunLabaRugi As String
    If V_LabaRugi >= 0 Then
        namaAkunLabaRugi = "Keuntungan Penjualan Aset"
    Else
        namaAkunLabaRugi = "Rugi Pelepasan Aset Tetap"
    End If
    
    ' 5. AMBIL KODE AKUN COA
    Dim k_KasBank As Variant, k_Akumulasi As Variant, k_Aset As Variant, k_LabaRugi As Variant
    On Error Resume Next
    k_KasBank = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match(V_AkunKasBank, tblCOA.ListColumns(2).DataBodyRange, 0))
    k_Akumulasi = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match(namaAkunAkumulasi, tblCOA.ListColumns(2).DataBodyRange, 0))
    k_Aset = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match(namaAkunAset, tblCOA.ListColumns(2).DataBodyRange, 0))
    k_LabaRugi = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match(namaAkunLabaRugi, tblCOA.ListColumns(2).DataBodyRange, 0))
    On Error GoTo 0
    
    If IsError(k_Akumulasi) Or IsError(k_Aset) Or IsError(k_LabaRugi) Then
        MsgBox "Gagal Posting! Periksa kembali Master Data COA Anda!", vbCritical, "Error Master COA"
        GoTo NyalakanSistem
    End If
    
    ' ====================================================================
    ' 6. EKSEKUSI PENJURNALAN (SISTEM APPEND SECEPAT KILAT)
    ' ====================================================================
    Dim newRow As ListRow
    
    ' --- BARIS 1: AKUMULASI ---
    If V_AkumulasiPenyusutan > 0 Then
        Set newRow = tblJurnal.ListRows.Add
        With newRow
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_KeteranganSeragam: .Range(4) = k_Akumulasi
            .Range(5) = namaAkunAkumulasi: .Range(6) = V_AkumulasiPenyusutan: .Range(7) = 0: .Range(8) = "Aset Tetap": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- BARIS 2: KAS/BANK ---
    If V_HargaJual > 0 And V_AkunKasBank <> "" Then
        Set newRow = tblJurnal.ListRows.Add
        With newRow
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_KeteranganSeragam: .Range(4) = k_KasBank
            .Range(5) = V_AkunKasBank: .Range(6) = V_HargaJual: .Range(7) = 0: .Range(8) = "Aset Tetap": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- BARIS 3: LABA RUGI ---
    If V_LabaRugi > 0 Then
        Set newRow = tblJurnal.ListRows.Add
        With newRow
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_KeteranganSeragam: .Range(4) = k_LabaRugi
            .Range(5) = namaAkunLabaRugi: .Range(6) = 0: .Range(7) = V_LabaRugi: .Range(8) = "Aset Tetap": .Range(9) = "Tidak"
        End With
    ElseIf V_LabaRugi < 0 Then
        Set newRow = tblJurnal.ListRows.Add
        With newRow
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_KeteranganSeragam: .Range(4) = k_LabaRugi
            .Range(5) = namaAkunLabaRugi: .Range(6) = Abs(V_LabaRugi): .Range(7) = 0: .Range(8) = "Aset Tetap": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- BARIS 4: ASAL ASET ---
    Set newRow = tblJurnal.ListRows.Add
    With newRow
        .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_KeteranganSeragam: .Range(4) = k_Aset
        .Range(5) = namaAkunAset: .Range(6) = 0: .Range(7) = V_HargaPerolehan: .Range(8) = "Aset Tetap": .Range(9) = "Tidak"
    End With
    

' Label pemulihan sistem Excel
NyalakanSistem:
    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Kunci kembali satpam proteksi sheet Jurnal Umum
    ' ====================================================================
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

