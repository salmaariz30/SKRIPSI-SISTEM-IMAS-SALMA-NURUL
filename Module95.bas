Attribute VB_Name = "Module95"
Sub JurnalUmumInputAsetTetap()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT ASET TETAP (MATCH TOTAL 100% TABEL)
    ' FITUR: KONDISIONAL "ASET LAMA" (NILAI BUKU) VS PEMBELIAN BARU (HARGA PEROLEHAN)
    ' ====================================================================
    Dim tblJurnal As ListObject, tblCOA As ListObject
    Dim wsInput As Worksheet
    
    ' Daftarkan objek sheet input ke variabel
    Set wsInput = Sheets("INPUT DATA NEW")
    
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    On Error GoTo 0
    
    If tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: 'TabelJurnalUmum' atau 'TabelDataCOA' tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' Pull Data Utama dari Form
    Dim V_Tanggal As Date: V_Tanggal = Date
    Dim V_NamaAset As String: V_NamaAset = wsInput.Range("C12").Value
    Dim V_Metode As String: V_Metode = wsInput.Range("F12").Value
    Dim V_AkunKasBank As String: V_AkunKasBank = wsInput.Range("F22").Value
    Dim V_NominalAset As Double: V_NominalAset = Val(wsInput.Range("C22").Value) ' Harga Perolehan Sesuai Form
    Dim V_NominalDP As Double: V_NominalDP = Val(wsInput.Range("I22").Value)   ' Nominal DP
    Dim V_NoBukti As String: V_NoBukti = wsInput.Range("I27").Value
    
    ' Pull Parameter Depresiasi Sesuai Form Input
    Dim V_TglBeli As Date: V_TglBeli = wsInput.Range("I12").Value
    Dim V_UmurTahun As Double: V_UmurTahun = Val(wsInput.Range("I17").Value)
    Dim V_NilaiResidu As Double: V_NilaiResidu = Val(wsInput.Range("C27").Value)
    
    Dim V_NilaiEksekusiJurnal As Double
    
    ' ====================================================================
    ' ?? LOGIKA KONDISIONAL METODE: "ASET LAMA" VS PEMBELIAN BARU
    ' ====================================================================
    If V_Metode = "Aset Lama" Or V_Metode = "Setoran Pemilik" Then
        ' --- JIKA ASET LAMA: HITUNG NILAI BUKU NETT (MATCH TABEL) ---
        If V_TglBeli > 0 And V_UmurTahun > 0 And V_Tanggal >= V_TglBeli Then
            Dim BiayaBulan As Double
            BiayaBulan = (V_NominalAset - V_NilaiResidu) / (V_UmurTahun * 12)
            
            ' 1. Hitung Batas Tgl Mulai Aset (Logika Day <= 15)
            Dim TglMulaiAset As Date
            If Day(V_TglBeli) <= 15 Then
                TglMulaiAset = DateSerial(Year(V_TglBeli), Month(V_TglBeli), 1)
            Else
                TglMulaiAset = DateSerial(Year(V_TglBeli), Month(V_TglBeli) + 1, 1)
            End If
            
            ' 2. Hitung Akumulasi s.d Tahun Lalu (Cut off 31 Des 2025)
            Dim TglCutOff2025 As Date: TglCutOff2025 = DateSerial(2025, 12, 31)
            Dim JmlBlnLalu As Long: JmlBlnLalu = 0
            Dim AkumLalu As Double: AkumLalu = 0
            
            If TglCutOff2025 >= TglMulaiAset Then
                JmlBlnLalu = DateDiff("m", TglMulaiAset, TglCutOff2025) + 1
                AkumLalu = JmlBlnLalu * BiayaBulan
                If AkumLalu > (V_NominalAset - V_NilaiResidu) Then AkumLalu = V_NominalAset - V_NilaiResidu
            End If
            
            ' 3. Hitung Beban Depresiasi Tahun Ini (2026 Berjalan)
            Dim TglAwal2026 As Date: TglAwal2026 = DateSerial(2026, 1, 1)
            Dim TitikAwal2026 As Date
            If TglMulaiAset > TglAwal2026 Then TitikAwal2026 = TglMulaiAset Else TitikAwal2026 = TglAwal2026
            
            Dim TglEoMonthSekarang As Date
            TglEoMonthSekarang = DateSerial(Year(V_Tanggal), Month(V_Tanggal) + 1, 0)
            
            Dim JmlBln2026 As Long
            JmlBln2026 = DateDiff("m", TitikAwal2026, TglEoMonthSekarang) + 1
            
            Dim PlafonSisa As Double
            PlafonSisa = V_NominalAset - V_NilaiResidu - AkumLalu
            If PlafonSisa < 0 Then PlafonSisa = 0
            
            Dim Beban2026 As Double
            Beban2026 = JmlBln2026 * BiayaBulan
            If Beban2026 > PlafonSisa Then Beban2026 = PlafonSisa
            
            ' 4. Ambil Sisa Nilai Buku Bersih
            Dim TotalAkumulasiPenyusutan As Double
            TotalAkumulasiPenyusutan = Round(AkumLalu + Beban2026, 0)
            
            V_NilaiEksekusiJurnal = V_NominalAset - TotalAkumulasiPenyusutan
        Else
            V_NilaiEksekusiJurnal = V_NominalAset
        End If
    Else
        ' --- JIKA PEMBELIAN BARU (TUNAI/KREDIT/MODAL): AMBIL HARGA PEROLEHAN UTUH ---
        V_NilaiEksekusiJurnal = V_NominalAset
    End If
    ' ====================================================================
    
    ' Sesuaikan porsi DP dan Sisa Utang secara proporsional terhadap nilai eksekusi jika metode Kredit
    Dim V_NilaiJurnalDP As Double
    Dim V_NilaiJurnalSisaUtang As Double
    
    If V_NominalAset > 0 Then
        V_NilaiJurnalDP = (V_NominalDP / V_NominalAset) * V_NilaiEksekusiJurnal
        V_NilaiJurnalSisaUtang = V_NilaiEksekusiJurnal - V_NilaiJurnalDP
    Else
        V_NilaiJurnalDP = 0
        V_NilaiJurnalSisaUtang = 0
    End If
    
    Dim V_DeskripsiJurnal As String
    If V_Metode = "Aset Lama" Then
        V_DeskripsiJurnal = "Perolehan Aset Tetap (Nilai Buku): " & V_NamaAset
    Else
        V_DeskripsiJurnal = "Pembelian Aset Tetap Baru: " & V_NamaAset & " (" & V_Metode & ")"
    End If
    
    ' 1. SET NAMA AKUN DEBIT BERDASARKAN INPUT CELL F17
    Dim namaAkunDebit As String: namaAkunDebit = wsInput.Range("F17").Value
    
    ' 2. LOGIKA COA UTK MENCARI KODE AKUN DEBIT
    Dim barisDebit As Variant, kodeAkunDebit As Variant
    On Error Resume Next
    barisDebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    On Error GoTo 0
    
    If IsError(kodeAkunDebit) Then
        MsgBox "Gagal! Akun Debit '" & namaAkunDebit & "' tidak terdaftar di COA!", vbCritical, "Error COA"
        Exit Sub
    End If

    ' 3. EKSEKUSI EXPORT KE JURNAL (ANTI-KEDIP)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' ?? OPERASI OPEN LOCK: Buka gembok Jurnal Umum sebelum menyisipkan baris baru
    ' ====================================================================
    Dim wsJurnal As Worksheet
    Set wsJurnal = tblJurnal.Parent
    wsJurnal.Unprotect Password:="IMAS"
    
    ' --- BARIS 1: DEBIT (Aset Bertambah Sebesar Nilai Eksekusi) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_DeskripsiJurnal
        .Range(4) = kodeAkunDebit: .Range(5) = namaAkunDebit
        .Range(6) = V_NilaiEksekusiJurnal: .Range(7) = 0
        .Range(8) = "Aset Tetap": .Range(9) = "Tidak"
    End With
    
    ' --- BARIS KREDIT (DIREKREASI BERDASARKAN KONDISI DENGAN ACUAN NILAI EKSEKUSI) ---
    Dim barisKredit As Variant, kodeAkunKredit As Variant
    
    If V_Metode = "Pembelian Kredit" Then
        ' === SPESIAL KONDISI KREDIT + DP ===
        
        ' A. Jurnal Baris Kredit 1: Kas/Bank untuk Uang Muka (DP)
        On Error Resume Next
        barisKredit = Application.Match(V_AkunKasBank, tblCOA.ListColumns(2).DataBodyRange, 0)
        kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
        On Error GoTo 0
        
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_DeskripsiJurnal & " - DP"
            .Range(4) = kodeAkunKredit: .Range(5) = V_AkunKasBank
            .Range(6) = 0: .Range(7) = V_NilaiJurnalDP
            .Range(8) = "Kas & Bank": .Range(9) = "Tidak"
        End With
        
        ' B. Jurnal Baris Kredit 2: Sisa ke Utang Lain-Lain
        On Error Resume Next
        barisKredit = Application.Match("Utang Lain-Lain", tblCOA.ListColumns(2).DataBodyRange, 0)
        kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
        On Error GoTo 0
        
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_DeskripsiJurnal & " - Sisa Utang"
            .Range(4) = kodeAkunKredit: .Range(5) = "Utang Lain-Lain"
            .Range(6) = 0: .Range(7) = V_NilaiJurnalSisaUtang
            .Range(8) = "Kewajiban": .Range(9) = "Tidak"
         End With

    Else
        ' === KONDISI NON-KREDIT ===
        Dim namaAkunKredit As String, kelompokKredit As String
        Select Case V_Metode
            Case "Aset Lama":         namaAkunKredit = "Ekuitas - Saldo Awal": kelompokKredit = "Ekuitas"
            Case "Pembelian Tunai":   namaAkunKredit = V_AkunKasBank:          kelompokKredit = "Kas & Bank"
            Case "Setoran Pemilik":   namaAkunKredit = "Modal Tambahan Disetor": kelompokKredit = "Ekuitas"
        End Select
        
        On Error Resume Next
        barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
        kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
        On Error GoTo 0
        
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_DeskripsiJurnal
            .Range(4) = kodeAkunKredit: .Range(5) = namaAkunKredit
            .Range(6) = 0: .Range(7) = V_NilaiEksekusiJurnal
            .Range(8) = kelompokKredit: .Range(9) = "Tidak"
        End With
    End If
    
    ' ====================================================================
    ' ?? RE-LOCK SYSTEM: Kunci kembali sheet Jurnal Umum & Blokir Shape agar Aman
    ' ====================================================================
    wsJurnal.Protect Password:="IMAS", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsJurnal.EnableSelection = xlNoSelection
    
    ' 4. CLEAR FORM & UX FOCUS
    With wsInput
        .Range("C12").Cells(1, 1).MergeArea.ClearContents
        .Range("C17").Cells(1, 1).MergeArea.ClearContents
        .Range("C22").Cells(1, 1).MergeArea.ClearContents
        .Range("C27").Cells(1, 1).MergeArea.ClearContents
        .Range("F12").Cells(1, 1).MergeArea.ClearContents
        .Range("F22").Cells(1, 1).MergeArea.ClearContents
        .Range("F27").Cells(1, 1).MergeArea.ClearContents
        .Range("I12").Cells(1, 1).MergeArea.ClearContents
        .Range("I22").Cells(1, 1).MergeArea.ClearContents
        
        .Activate: .Range("F12").Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Data Aset Tetap Berhasil Disimpan!", vbInformation, "Sistem Sukses"
End Sub
