Attribute VB_Name = "Module56"
Sub JurnalUmumInputUtangBiaya()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT UTANG BIAYA (AKRUAL) KE JURNAL UMUM
    ' Terkunci Khusus untuk Sheet: UTANG_INPUT USAHA
    ' INPUT UTAMA DI H16 ("Utang Biaya..."), OTOMATIS DIUBAH JADI "Beban..." UNTUK DEBIT
    ' COMPATIBLE WITH USA/INDO REGIONAL SETTINGS & MERGED CELLS SUPPORT
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject, tblCOA As ListObject
    Dim wsInput As Worksheet
    Dim wsJurnal As Worksheet
    Const PWD As String = "IMAS"
    
    Set wsInput = Sheets("UTANG_INPUT USAHA")
    
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
    
    ' 2. Tarik Data dari Form Input Utang Biaya (Mendukung Merged Cells)
    Dim V_Tanggal As Date: V_Tanggal = wsInput.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As String: V_NoBukti = wsInput.Range("H12").MergeArea.Cells(1, 1).Value
    Dim V_KategoriUtang As String: V_KategoriUtang = wsInput.Range("H16").MergeArea.Cells(1, 1).Value ' Pilihan utama di H16 (Misal: "Utang Biaya Gaji & Upah Kantor")
    Dim V_KeteranganUser As String: V_KeteranganUser = wsInput.Range("D20").MergeArea.Cells(1, 1).Value
    Dim V_Nominal As Double: V_Nominal = Val(wsInput.Range("D24").MergeArea.Cells(1, 1).Value)
    
    ' Validasi Input Dasar agar tidak ada Jurnal Kosong
    If V_NoBukti = "" Or V_Nominal <= 0 Or V_KategoriUtang = "" Then
        MsgBox "Gagal! Pastikan No. Bukti, Pilihan Akun di H16, dan Nominal sudah terisi.", vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 3. LOGIKA PREMAN: UBAH "Utang Biaya..." MENJADI "Beban..." UNTUK LAWAN AKUN (DEBIT)
    Dim namaAkunDebit As String
    Dim namaAkunKredit As String
    
    namaAkunKredit = V_KategoriUtang ' Sisi Kredit murni mengambil teks dari H16
    
    ' Memotong kata "Utang Biaya" di depan, lalu diganti menjadi kata "Beban"
    If Left(namaAkunKredit, 11) = "Utang Biaya" Then
        namaAkunDebit = "Beban" & Mid(namaAkunKredit, 12)
    Else
        ' Antisipasi jika ada variasi penulisan teks lain di master COA
        namaAkunDebit = Replace(namaAkunKredit, "Utang Biaya", "Beban")
    End If
    
    ' Penyusunan Deskripsi Jurnal
    Dim V_DeskripsiJurnal As String
    If V_KeteranganUser <> "" Then
        V_DeskripsiJurnal = "Pencatatan Akrual " & namaAkunDebit & " (" & V_KeteranganUser & ")"
    Else
        V_DeskripsiJurnal = "Pencatatan Penyesuaian Akrual " & namaAkunDebit
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
    
    ' Proteksi jika nama akun rawan typo / belum terdaftar di COA Utama
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        MsgBox "Gagal Menjurnal! Akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' belum terdaftar di TabelDataCOA  !", _
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
    
    ' --- BARIS DEBIT (Beban Operasional Bertambah dari Hasil Konversi) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_DeskripsiJurnal           ' Kolom 3: Keterangan
        .Range(4) = kodeAkunDebit               ' Kolom 4: Kode Akun
        .Range(5) = namaAkunDebit               ' Kolom 5: Nama Akun
        .Range(6) = V_Nominal                   ' Kolom 6: Debit
        .Range(7) = 0                           ' Kolom 7: Kredit
        .Range(8) = "Beban Operasional"         ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' --- BARIS KREDIT (Kewajiban Utang Biaya Asli dari Sel H16) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_DeskripsiJurnal           ' Kolom 3: Keterangan
        .Range(4) = kodeAkunKredit              ' Kolom 4: Kode Akun
        .Range(5) = namaAkunKredit              ' Kolom 5: Nama Akun
        .Range(6) = 0                           ' Kolom 6: Debit
        .Range(7) = V_Nominal                   ' Kolom 7: Kredit
        .Range(8) = "Liabilitas"                ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' ====================================================================
    ' ?? KUNCI KEMBALI SHEET JURNAL UMUM SETELAH SELESAI
    ' ====================================================================
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 7. NOTIFIKASI BERHASIL
    MsgBox "Data Utang Biaya Berhasil Disimpan!", _
           vbInformation, "Sistem Sukses"
End Sub

