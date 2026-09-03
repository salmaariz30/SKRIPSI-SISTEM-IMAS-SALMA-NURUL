Attribute VB_Name = "Module102"
Sub JurnalUmumInputDisposalPersediaanDinamis()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL DISPOSAL PERSEDIAAN DENGAN KALKULASI HPP DINAMIS
    ' METODE PERPETUAL - COMPATIBLE WITH USA REGIONAL SETTINGS
    ' VERSION 6.0: VERSI TOBAT (DATA INPUT TETAP UTUH & AMAN)
    ' ====================================================================
    
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("PERSEDIAAN_INPUT DISPOSAL")
    Dim SheetJurnal As Worksheet: Set SheetJurnal = tblJurnal.Parent
    
    Dim tblInputDisposal As ListObject
    Dim tblMasterStok As ListObject
    Dim namaTabelMaster As String
    
    ' 1. AMBIL DATA HEADER FORM (SUPPORT MERGED CELLS)
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("F12").MergeArea.Cells(1, 1).Value
    Dim V_Kategori As String: V_Kategori = Trim(wsInput.Range("D16").MergeArea.Cells(1, 1).Value)
    
    ' Validasi Set Tabel Input Disposal di Sheet Tersebut
    On Error Resume Next
    Set tblInputDisposal = wsInput.ListObjects("TabelInputDisposal")
    On Error GoTo 0
    
    If tblInputDisposal Is Nothing Then
        MsgBox "Error: 'TabelInputDisposal' tidak ditemukan di sheet input!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. VALIDASI DAN PENENTUAN TABEL MASTER STOK BERDASARKAN KATEGORI (D16)
    Select Case V_Kategori
        Case "Bahan Baku":    namaTabelMaster = "TabelStokBahan"
        Case "Barang Jadi":   namaTabelMaster = "TabelStockJadi"
        Case "Barang Dagang": namaTabelMaster = "TabelStokDagang"
        Case Else
            MsgBox "Kategori '" & V_Kategori & "' di D16 tidak valid!" & vbCrLf & _
                   "Pilihan harus: 'Bahan Baku', 'Barang Jadi', atau 'Barang Dagang'.", vbExclamation, "Validasi Gagal"
            Exit Sub
    End Select
    
    On Error Resume Next
    Set tblMasterStok = Range(namaTabelMaster).ListObject
    On Error GoTo 0
    
    If tblMasterStok Is Nothing Then
        MsgBox "Error: Master " & namaTabelMaster & " tidak ditemukan di workbook ini!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' Validasi Pengaman Header
    If V_Tanggal = 0 Or V_NoBukti = "" Then
        MsgBox "Mohon pastikan Tanggal dan No Bukti sudah terisi.", vbExclamation, "Data Belum Lengkap"
        Exit Sub
    End If
    
    ' Check Apakah Tabel Input Kosong Melompong
    If tblInputDisposal.ListRows.Count = 0 Then
        MsgBox "Tabel Input Disposal masih kosong!", vbExclamation, "Data Belum Lengkap"
        Exit Sub
    End If
    
    If tblInputDisposal.DataBodyRange Is Nothing Then
        MsgBox "Tabel Input Disposal tidak memiliki baris data untuk diproses!", vbExclamation, "Data Kosong"
        Exit Sub
    End If
    
    ' ====================================================================
    ' PENGAMAN UTAMA SAKTI + PROSES LOOPING NYARI NILAI BARANG
    ' ====================================================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' BUKA PROTEKSI SHEET TARGET JURNAL UMUM
    SheetJurnal.Unprotect Password:="IMAS"
    
    Dim BarisInput As ListRow
    Dim NamaBarang As String
    Dim Kuantitas As Double
    Dim BarisMaster As Variant
    Dim HargaSatuan As Double
    Dim TotalNilaiDisposal As Double: TotalNilaiDisposal = 0
    Dim IsKetemu As Boolean
    
    ' TELUSURI DATA DI TABEL INPUT DISPOSAL SATU PER SATU
    For Each BarisInput In tblInputDisposal.ListRows
        NamaBarang = Trim(BarisInput.Range(1).Value) ' Kolom 1 di Form Input: Nama Barang
        Kuantitas = Val(BarisInput.Range(3).Value)   ' Kolom 3 di Form Input: Kuantitas
        
        If NamaBarang <> "" And Kuantitas > 0 Then
            ' Cari posisi nama barang di KOLOM KETIGA (ListColumns(3)) dari Tabel Master Stok
            On Error Resume Next
            BarisMaster = Application.Match(NamaBarang, tblMasterStok.ListColumns(3).DataBodyRange, 0)
            On Error GoTo 0
            
            ' Logika Pengaman Anti-Kuning / Type Mismatch
            IsKetemu = False
            If Not IsError(BarisMaster) Then
                If BarisMaster > 0 Then IsKetemu = True
            End If
            
            If IsKetemu Then
                ' Ambil nilai harga modal dari KOLOM KELIMA (ListColumns(5)) dari Tabel Master Stok
                HargaSatuan = Val(Application.Index(tblMasterStok.ListColumns(5).DataBodyRange, BarisMaster))
                
                ' Kalikan harga modal dengan kuantitas, lalu akumulasikan ke total
                TotalNilaiDisposal = TotalNilaiDisposal + (HargaSatuan * Kuantitas)
            Else
                ' KUNCI KEMBALI SHEET SEBELUM EXIT AGAR TETAP AMAN
                SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
                
                ' Jika barang di form tidak ketemu di master, matikan proteksi layar lalu teriak!
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Application.EnableEvents = True
                MsgBox "Barang '" & NamaBarang & "' tidak ditemukan di Kolom 3 master " & namaTabelMaster & "!" & vbCrLf & _
                       "Silakan cek kembali apakah nama barang di form input sudah sama persis dengan master.", vbCritical, "Nama Barang Tidak Terdaftar"
                Exit Sub
            End If
        End If
    Next BarisInput
    
    ' Validasi jika setelah dihitung ternyata total nilainya zonk/nol
    If TotalNilaiDisposal <= 0 Then
        ' KUNCI KEMBALI SHEET SEBELUM EXIT AGAR TETAP AMAN
        SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "Total nilai disposal Rp0 atau kuantitas belum diisi dengan benar!", vbExclamation, "Gagal Posting"
        Exit Sub
    End If
    
    ' 3. PENENTUAN NAMA AKUN & KELOMPOK UNTUK JURNAL
    Dim namaAkunDebit As String
    Dim kelompokDebit As String: kelompokDebit = "Beban / HPP"
    
    If V_Kategori = "Barang Dagang" Then
        namaAkunDebit = "Beban Kerugian Disposal Persediaan Dagang"
    Else
        namaAkunDebit = "Beban Kerugian Disposal Persediaan Produksi"
    End If
    
    Dim namaAkunKredit As String: namaAkunKredit = "Persediaan " & V_Kategori
    Dim kelompokKredit As String: kelompokKredit = "Aset Lancar"
    
    Dim V_DeskripsiCustom As String: V_DeskripsiCustom = "Disposal/Penghapusan " & V_Kategori & " (No. Bukti: " & V_NoBukti & ")"
    
    ' 4. AMBIL KODE AKUN DARI MASTER COA
    Dim barisDebitCOA As Variant, barisKreditCOA As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisDebitCOA = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKreditCOA = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebitCOA)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKreditCOA)
    On Error GoTo 0
    
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        ' KUNCI KEMBALI SHEET SEBELUM EXIT AGAR TETAP AMAN
        SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        MsgBox "Nama akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' belum terdaftar di COA !", vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' 5. APPEND DATA INSTAN KE JURNAL UMUM
    Dim LastRowIdx As Long
    Dim IsTableEmpty As Boolean
    
    On Error Resume Next
    IsTableEmpty = (tblJurnal.ListRows.Count = 0 Or (tblJurnal.ListRows.Count = 1 And tblJurnal.DataBodyRange.Cells(1, 1).Value = ""))
    On Error GoTo 0
    
    If IsTableEmpty Then
        If tblJurnal.ListRows.Count = 0 Then tblJurnal.ListRows.Add
        LastRowIdx = 1
    Else
        tblJurnal.ListRows.Add AlwaysInsert:=False
        tblJurnal.ListRows.Add AlwaysInsert:=False
        LastRowIdx = tblJurnal.ListRows.Count - 1
    End If
    
    ' --- INJEKSI BARIS DEBIT (BEBAN DISPOSAL) ---
    With tblJurnal.ListRows(LastRowIdx)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiCustom
        .Range(4) = kodeAkunDebit
        .Range(5) = namaAkunDebit
        .Range(6) = TotalNilaiDisposal
        .Range(7) = 0
        .Range(8) = kelompokDebit
        .Range(9) = "Tidak"
        With .Range.Font: .Name = "Segoe UI": .Size = 10: .ColorIndex = 1: End With
        .Range(6).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' --- INJEKSI BARIS KREDIT (PERSEDIAAN KELUAR) ---
    With tblJurnal.ListRows(LastRowIdx + 1)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiCustom
        .Range(4) = kodeAkunKredit
        .Range(5) = namaAkunKredit
        .Range(6) = 0
        .Range(7) = TotalNilaiDisposal
        .Range(8) = kelompokKredit
        .Range(9) = "Tidak"
        With .Range.Font: .Name = "Segoe UI": .Size = 10: .ColorIndex = 1: End With
        .Range(6).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
    End With
    
    ' TUTUP PROTEKSI SHEET JURNAL UMUM KEMBALI
    SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
    
    ' ====================================================================
    ' 6. PASCA POSTING (DATA INPUT TETAP UTUH TIDAK DIHAPUS)
    ' ====================================================================
    ' Bagian hapus tabel otomatis sudah dibuang total demi keamanan
    
    wsInput.Activate
    wsInput.Range("D12").Select
    
    ' Pulihkan sistem Excel
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Sukses! Jurnal Disposal senilai Rp " & Format(TotalNilaiDisposal, "#,##0") & " Berhasil Diposting!", vbInformation, "Sukses"
    Exit Sub

NyalakanSistem:
    ' KUNCI KEMBALI JIKA TERJADI EMERGENCY RESET SUTDOWN KODE
    SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

