Attribute VB_Name = "Module105"
Sub JurnalUmumInputUtangBank()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT UTANG BANK KE JURNAL UMUM
    ' Terkunci Khusus untuk Sheet: UTANG_INPUT BANK
    ' DUAL MODE: OTOMATIS DETEKSI UTANG BARU (D20) ATAU SALDO AWAL (D32)
    ' COMPATIBLE WITH USA/INDO REGIONAL SETTINGS & MERGED CELLS SUPPORT
    ' STATUS: AUTO UNPROTECT / PROTECT SHEET TARGET JURNAL (PASSWORD: IMAS)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject, tblCOA As ListObject
    Dim wsInput As Worksheet
    Dim wsJurnal As Worksheet
    Const PWD As String = "IMAS"
    
    Set wsInput = Sheets("UTANG_INPUT BANK")
    
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Utama
    If tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: 'TabelJurnalUmum' atau 'TabelDataCOA' tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' Set Sheet Jurnal Umum berdasarkan tabelnya
    Set wsJurnal = tblJurnal.Parent
    
    ' 2. Tarik Data dari Form Input Utang Bank (Mendukung Merged Cells)
    Dim V_Tanggal As Date: V_Tanggal = wsInput.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As String: V_NoBukti = wsInput.Range("H12").MergeArea.Cells(1, 1).Value
    Dim V_PlafonKredit As Double: V_PlafonKredit = Val(wsInput.Range("D20").MergeArea.Cells(1, 1).Value)
    Dim V_PilihanKasBank As String: V_PilihanKasBank = wsInput.Range("H28").MergeArea.Cells(1, 1).Value ' Akun Kas/Bank atau teks "SALDO AWAL"
    Dim V_SisaUtang As Double: V_SisaUtang = Val(wsInput.Range("D32").MergeArea.Cells(1, 1).Value)
    
    ' 3. LOGIKA PENENTUAN NOMINAL & NAMA AKUN LAWAN (DEBIT)
    Dim namaAkunDebit As String
    Dim namaAkunKredit As String
    Dim V_NominalFinal As Double
    Dim V_KelompokDebit As String
    
    namaAkunKredit = "Utang Bank Jangka Panjang" ' Akun Kredit selalu Tetap
    
    ' Kondisi Percabangan Cerdas Sesuai Permintaan
    If UCase(Trim(V_PilihanKasBank)) = "SALDO AWAL" Then
        ' Mode Saldo Awal (Utang Tahun Lalu)
        namaAkunDebit = "Ekuitas - Saldo Awal"
        V_NominalFinal = V_SisaUtang
        V_KelompokDebit = "Ekuitas"
    Else
        ' Mode Pencairan Utang Baru Berjalan
        namaAkunDebit = V_PilihanKasBank
        V_NominalFinal = V_PlafonKredit
        V_KelompokDebit = "Kas & Bank"
    End If
    
    ' Validasi Pengaman Angka & Input Kosong
    If V_NoBukti = "" Or V_NominalFinal <= 0 Or V_PilihanKasBank = "" Then
        MsgBox "Gagal! Pastikan No. Bukti, Pilihan di H28, dan Nominal terkait telah terisi dengan benar.", vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' Penyusunan Deskripsi Jurnal Otomatis
    Dim V_DeskripsiJurnal As String
    If namaAkunDebit = "Ekuitas - Saldo Awal" Then
        V_DeskripsiJurnal = "Pencatatan Migrasi Saldo Awal Utang Bank Jangka Panjang"
    Else
        V_DeskripsiJurnal = "Penerimaan Pencairan Kredit Utang Bank via " & namaAkunDebit
    End If
    
    ' 4. Ambil Kode Akun dari Master COA Berdasarkan Nama Akun Hasil Olahan Logika
    Dim barisDebit As Variant, barisKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisDebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
    On Error GoTo 0
    
    ' Proteksi jika nama akun kas/bank atau akun saldo awal belum terdaftar di COA
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        MsgBox "Gagal Menjurnal! Nama akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' tidak ditemukan di TabelDataCOA Utama!", _
               vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' 5. EKSEKUSI EXPORT DUA BARIS KE JURNAL UMUM (SISTEM ANTI-KEDIP)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' ?? BUKA GEMBOK SHEET JURNAL UMUM SEBELUM INSERSI BARIS
    ' ====================================================================
    wsJurnal.Unprotect Password:=PWD
    
    ' --- BARIS DEBIT (Kas Masuk / Penyesuaian Ekuitas Saldo Awal) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_DeskripsiJurnal           ' Kolom 3: Keterangan
        .Range(4) = kodeAkunDebit               ' Kolom 4: Kode Akun
        .Range(5) = namaAkunDebit               ' Kolom 5: Nama Akun
        .Range(6) = V_NominalFinal              ' Kolom 6: Debit
        .Range(7) = 0                           ' Kolom 7: Kredit
        .Range(8) = V_KelompokDebit             ' Kolom 8: Kelompok Dinamis
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' --- BARIS KREDIT (Kewajiban Utang Bank Jangka Panjang Bertambah) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_DeskripsiJurnal           ' Kolom 3: Keterangan
        .Range(4) = kodeAkunKredit              ' Kolom 4: Kode Akun
        .Range(5) = namaAkunKredit              ' Kolom 5: Nama Akun
        .Range(6) = 0                           ' Kolom 6: Debit
        .Range(7) = V_NominalFinal              ' Kolom 7: Kredit
        .Range(8) = "Liabilitas"                ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
        
    ' ====================================================================
    ' ?? KUNCI KEMBALI SHEET JURNAL UMUM SETELAH EKSPOR DATA SELESAI
    ' ====================================================================
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
        
    ' --- POTONGAN KODE SAPU BERSIH SEL NORMAL ---
    Dim sel As Variant
    With wsInput
        ' Looping satu per satu biar Excel gak pusing borongan
        For Each sel In Array("D12", "D16", "D20", "D24", "D28", "D32", _
                              "H12", "H16", "H20", "H24", "H28", "H32")
            .Range(sel).MergeArea.ClearContents
        Next sel
        
        ' Kembalikan kursor dengan tenang ke D12
        .Range("D12").MergeArea.Cells(1, 1).Select
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 7. NOTIFIKASI BERHASIL
    MsgBox "Data Pencatatan Utang Bank Berhasil Disimpan!", _
           vbInformation, "Sistem Sukses"
End Sub

