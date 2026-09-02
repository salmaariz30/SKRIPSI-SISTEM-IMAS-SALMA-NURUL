Attribute VB_Name = "Module52"
Sub JurnalUmumInputTarikModal()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT TARIK MODAL KE JURNAL UMUM
    ' Terkunci Khusus untuk Sheet: PENGELUARAN USAHA_INPUT TARIK
    ' Otomatis Deteksi Kode COA & Inject 2 Baris (Debit & Kredit)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("PENGELUARAN USAHA_INPUT TARIK")
    Dim wsJurnal As Worksheet: Set wsJurnal = tblJurnal.Parent
    Const PWD As String = "IMAS"
    
    ' 2. Tarik Data dari Cell Input Tarik   (Mendukung Merged Cells)
    Dim V_JenisPenarikan As String: V_JenisPenarikan = wsInput.Range("D11").Value
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("G11").Value
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("D15").Value
    Dim V_AkunKasBank As String: V_AkunKasBank = wsInput.Range("G15").Value
    Dim V_Nominal As Variant: V_Nominal = wsInput.Range("D19").Value
    Dim V_NamaPenarik As String: V_NamaPenarik = wsInput.Range("G19").Value
    Dim V_RekeningPenarik As String: V_RekeningPenarik = wsInput.Range("D23").Value
    Dim V_Persentase As String: V_Persentase = wsInput.Range("G23").Value
    Dim V_Deskripsi As Variant: V_Deskripsi = wsInput.Range("D27").Value
    
    
    ' 3. LOGIKA DINAMIS TENTUKAN NAMA AKUN DEBIT BERDASARKAN COA RESMI
    Dim namaAkunDebit As String
    Select Case V_JenisPenarikan
        Case "Penarikan Pribadi Owner", "Prive (Penarikan Pribadi Owner)", "Prive"
            namaAkunDebit = "Penarikan Pribadi Pemilik (Prive)" ' Menembak Kode 3-3100
            
        Case "Pembagian Laba Usaha", "Dividen (Pembagian Laba Usaha)", "Dividen"
            namaAkunDebit = "Pembagian Laba Usaha"              ' Menembak Kode 3-3200
            
        Case "Reduksi Ekuitas", "Pengembalian Modal (Reduksi Ekuitas)", "Pengembalian Modal"
            namaAkunDebit = "Modal Pemilik"             ' Menembak Kode 3-1100 (Atau sesuaikan nama COA)
            
        Case Else
            MsgBox "Jenis penarikan tidak dikenal oleh sistem jurnal!", vbCritical, "Error Logika"
            Exit Sub
    End Select
    
    ' Akun Kredit Kunci Mati: Akun Kas/Bank pilihan   di G15
    Dim namaAkunKredit As String: namaAkunKredit = V_AkunKasBank
    
    ' 4. Ambil Kode Akun (Kolom 1) Berdasarkan Nama Akun (Kolom 2) di COA secara Otomatis
    Dim barisDebit As Variant, barisKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisDebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
    On Error GoTo 0
    
    ' Proteksi jika nama akun di form tidak terdaftar di master COA
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        MsgBox "Gagal menjurnal! Nama akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' tidak ditemukan di TabelDataCOA!", _
               vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' 5. EKSEKUSI EXPOR DUA BARIS KE JURNAL UMUM (ANTI-KEDIP)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Buka Gembok Lembar Jurnal Umum sebelum diperbarui
    wsJurnal.Unprotect Password:=PWD
    
    ' --- BARIS DEBIT (Akun Distribusi Ekuitas / Modal Berkurang) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_Deskripsi                 ' Kolom 3: Keterangan / Deskripsi
        .Range(4) = kodeAkunDebit               ' Kolom 4: Kode Akun
        .Range(5) = namaAkunDebit               ' Kolom 5: Nama Akun
        .Range(6) = V_Nominal                   ' Kolom 6: Debit
        .Range(7) = 0                           ' Kolom 7: Kredit
        .Range(8) = "Modal & Ekuitas"           ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' --- BARIS KREDIT (Akun Kas / Bank Berkurang) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_Deskripsi                 ' Kolom 3: Keterangan / Deskripsi
        .Range(4) = kodeAkunKredit              ' Kolom 4: Kode Akun
        .Range(5) = namaAkunKredit              ' Kolom 5: Nama Akun
        .Range(6) = 0                           ' Kolom 6: Debit
        .Range(7) = V_Nominal                   ' Kolom 7: Kredit
        .Range(8) = "Modal & Ekuitas"              ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' Kunci Kembali Lembar Jurnal Umum setelah diperbarui
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
    
    ' 6. SILENT CLEAN FORM (Otomatis Bersih Kilat & Kembalikan Kursor)
    With wsInput
        .Range("D11").MergeArea.ClearContents
        .Range("G11").MergeArea.ClearContents
        .Range("D15").MergeArea.ClearContents
        .Range("G15").MergeArea.ClearContents
        .Range("D19").MergeArea.ClearContents
        .Range("G19").MergeArea.ClearContents
        .Range("D23").MergeArea.ClearContents
        .Range("D27").MergeArea.ClearContents
        
        ' UX Flow: Kembalikan kursor manis   ke hulu cell awal (D11)
        .Activate
        .Range("D11").MergeArea.Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 7. NOTIFIKASI SUKSES
    MsgBox "Transaksi Penarikan Modal telah berhasil disimpan!", _
           vbInformation, "Penyimpanan Sukses"
End Sub

