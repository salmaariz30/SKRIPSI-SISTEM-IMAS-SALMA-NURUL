Attribute VB_Name = "Module64"
Sub JurnalUmumInputPenghasilanLain()
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("PENDAPATAN_PENGHASILAN LAIN")
    Dim wsJurnal As Worksheet
    Const PWD As String = "IMAS"
    
    ' 2. Tarik Data dari Cell Input Penghasilan Lain   (Mendukung Merged Cells)
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("D11").Value
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("H11").Value
    Dim V_NamaPendapatan As String: V_NamaPendapatan = Trim(wsInput.Range("F11").Value)
    Dim V_Nominal As Variant: V_Nominal = wsInput.Range("H15").Value
    Dim V_AkunKasBank As String: V_AkunKasBank = Trim(wsInput.Range("F15").Value)
    Dim V_Deskripsi As Variant: V_Deskripsi = wsInput.Range("D19").Value
    
    ' 3. Validasi Pengaman Awal (Anti-Kosong)
    If V_Tanggal = "" Or V_NoBukti = "" Or V_NamaPendapatan = "" Or V_Nominal = "" Or V_AkunKasBank = "" Then
        MsgBox "VALIDASI INPUT GAGAL!" & vbCrLf & _
               "Mohon pastikan seluruh data input formulir Penghasilan Lain telah terisi lengkap.", _
               vbExclamation, "Data Belum Lengkap"
        Exit Sub
    End If
    
    ' 4. LOGIKA JURNAL PENERIMAAN KAS/BANK
    ' Debit = Akun Kas/Bank Penerima
    ' Kredit = Akun Pendapatan Non-Operasional yang dipilih
    Dim namaAkunDebit As String: namaAkunDebit = V_AkunKasBank
    Dim namaAkunKredit As String: namaAkunKredit = V_NamaPendapatan
    
    ' 5. Ambil Kode Akun (Kolom 1) Berdasarkan Nama Akun (Kolom 2) di COA secara Otomatis
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
    
    ' 6. EKSEKUSI EXPOR DUA BARIS KE JURNAL UMUM (ANTI-KEDIP)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' DETEKSI & BUKA PROTEKSI SHEET JURNAL UMUM SECARA DINAMIS
    Set wsJurnal = tblJurnal.Parent
    wsJurnal.Unprotect Password:=PWD
    
    ' --- BARIS DEBIT (Akun Kas & Bank Pilihan Bertambah) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(4).NumberFormat = "@"
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_Deskripsi                 ' Kolom 3: Keterangan / Deskripsi
        .Range(4) = kodeAkunDebit               ' Kolom 4: Kode Akun
        .Range(5) = namaAkunDebit               ' Kolom 5: Nama Akun
        .Range(6) = V_Nominal                   ' Kolom 6: Debit
        .Range(7) = 0                           ' Kolom 7: Kredit
        .Range(8) = "Kas & Bank"                ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' --- BARIS KREDIT (Akun Penghasilan Lain Bertambah) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(4).NumberFormat = "@"
        .Range(1) = V_Tanggal                   ' Kolom 1: Tanggal
        .Range(2) = V_NoBukti                   ' Kolom 2: No Bukti
        .Range(3) = V_Deskripsi                 ' Kolom 3: Keterangan / Deskripsi
        .Range(4) = kodeAkunKredit              ' Kolom 4: Kode Akun
        .Range(5) = namaAkunKredit              ' Kolom 5: Nama Akun
        .Range(6) = 0                           ' Kolom 6: Debit
        .Range(7) = V_Nominal                   ' Kolom 7: Kredit
        .Range(8) = "Pendapatan"                ' Kolom 8: Kelompok
        .Range(9) = "Tidak"                     ' Kolom 9: Penyesuaian
    End With
    
    ' KUNCI KEMBALI SHEET JURNAL UMUM
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
    
    ' 7. SILENT CLEAN FORM (Otomatis Bersih Kilat Anti-Merger Bug)
    With wsInput
        On Error Resume Next
        .Range("D11").MergeArea.ClearContents   ' Tanggal
        .Range("H11").MergeArea.ClearContents   ' No Bukti
        .Range("F11").MergeArea.ClearContents   ' Nama Pendapatan
        .Range("D15").MergeArea.ClearContents   ' Akun Kas/Bank (Revisi: Ini yang bener!)
        .Range("H15").MergeArea.ClearContents   ' Jumlah Nominal
        .Range("D19").MergeArea.ClearContents   ' Deskripsi
        .Range("H19").MergeArea.ClearContents   ' Catatan/Pendukung Tambahan (Revisi: Ikut disapu bersih!)
        On Error GoTo 0
        
        ' UX Flow: Kembalikan kursor manis   ke hulu cell awal (D11)
        .Activate
        .Range("D11").MergeArea.Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 8. NOTIFIKASI SUKSES
    MsgBox "Transaksi Penghasilan Lain telah berhasil dicatat ke Jurnal Umum!", _
           vbInformation, "Penyimpanan Sukses"
End Sub

