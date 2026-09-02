Attribute VB_Name = "Module57"
Sub JurnalUmumInputBeban()
    ' ====================================================================
    ' MODUL UTAMA: JURNAL OTOMATIS INPUT BEBAN KE JURNAL UMUM (DUA MUARA)
    ' FIXED KORDINAT: No Bukti (G12), Tanggal (D16), Nominal (D20)
    ' LOGIKA REVISI  : Status Penyesuaian Kolom 9 SEPAKET (Kredit Ikut Debit)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("PENGELUARAN USAHA_INPUT DATA")
    Dim wsJurnal As Worksheet: Set wsJurnal = tblJurnal.Parent
    Const PWD As String = "IMAS"
    
    ' KUNCI UTAMA: Paksa Excel fokus ke sheet input dari awal detik pertama klik!
    wsInput.Activate
    
    ' 2. Tarik Data Sesuai Kordinat Presisi   (Mendukung Merged Cells)
    Dim V_NamaBeban As String: V_NamaBeban = Trim(wsInput.Range("D12").Value)
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("G12").Value
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("D16").Value
    Dim V_Nominal As Variant: V_Nominal = wsInput.Range("D20").Value
    Dim V_AkunKasBank As String: V_AkunKasBank = Trim(wsInput.Range("G20").Value)
    Dim V_StatusBayar As String: V_StatusBayar = Trim(wsInput.Range("D24").Value)
    Dim V_Deskripsi As Variant: V_Deskripsi = wsInput.Range("D28").Value
    
    ' Validasi dasar agar tidak menjurnal data kosong
    If V_NamaBeban = "" Or V_AkunKasBank = "" Or V_StatusBayar = "" Then
        MsgBox "Data belum lengkap! Harap periksa kembali inputan Kategori Beban, Kas/Bank, dan Status Pembayaran.", vbExclamation, "Input Kosong"
        Exit Sub
    End If
    
    ' 3. LOGIKA DUA MUARA: Tentukan Nama Akun Debit & Status Penyesuaian Secara Dinamis
    Dim namaAkunDebit As String
    Dim kelompokJurnal As String
    Dim statusPenyesuaian As String
    
    Select Case V_StatusBayar
        Case "Lunas"
            namaAkunDebit = V_NamaBeban
            kelompokJurnal = "Beban Operasional"
            statusPenyesuaian = "Tidak" ' Sepaket terisi "Tidak" jika Lunas
            
        Case "Dibayar di Muka", "Dibayar di Muka"
            namaAkunDebit = "Beban Dibayar di Muka"
            kelompokJurnal = "Aset Lancar"
            statusPenyesuaian = "Iya"   ' Sepaket terisi "Iya" jika Dibayar Dimuka
            
        Case Else
            MsgBox "Status Pembayaran '" & V_StatusBayar & "' tidak dikenal!", vbCritical, "Error Logika"
            Exit Sub
    End Select
    
    Dim namaAkunKredit As String: namaAkunKredit = V_AkunKasBank
    
    ' 4. SISIR KODE COA (Mencari Kode Akun di Kolom 1 Berdasarkan Nama di Kolom 2)
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
        MsgBox "Gagal menjurnal! Akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' tidak ditemukan di master COA!", vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' 5. EKSEKUSI EXPORT DUA BARIS KE JURNAL UMUM (ANTI-KEDIP, ANTI-KLIK DUA KALI & ANTI-ERROR 1004)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Buka Gembok Lembar Jurnal Umum sebelum diperbarui
    wsJurnal.Unprotect Password:=PWD
    
    ' --- BARIS DEBIT (Sisi Anggaran Pengeluaran) ---
    With tblJurnal.ListRows.Add
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_Deskripsi
        .Range(4) = kodeAkunDebit
        .Range(5) = namaAkunDebit
        .Range(6) = V_Nominal
        .Range(7) = 0
        .Range(8) = kelompokJurnal
        .Range(9) = statusPenyesuaian   ' <-- Mengikuti variabel penyesuaian dinamis
    End With
    
    ' --- BARIS KREDIT (Sisi Kas / Bank Berkurang) ---
    With tblJurnal.ListRows.Add
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_Deskripsi
        .Range(4) = kodeAkunKredit
        .Range(5) = namaAkunKredit
        .Range(6) = 0
        .Range(7) = V_Nominal
        .Range(8) = "Aset Lancar"
        .Range(9) = statusPenyesuaian   ' <-- SEPAKET! Baris kedua sekarang mutlak menjiplak baris pertama  !
    End With
    
    ' Kunci Kembali Lembar Jurnal Umum setelah diperbarui
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
    
    ' 6. SILENT CLEAN FORM (Otomatis Sapu Bersih Tanpa Bekas)
    With wsInput
        .Range("D12").MergeArea.ClearContents   ' Bersihkan Kategori Beban
        .Range("G12").MergeArea.ClearContents   ' Bersihkan No Bukti
        .Range("D16").MergeArea.ClearContents   ' Bersihkan Tanggal
        .Range("D20").MergeArea.ClearContents   ' Bersihkan Nominal
        .Range("G20").MergeArea.ClearContents   ' Bersihkan Kas/Bank
        .Range("D24").MergeArea.ClearContents   ' Bersihkan Status Pembayaran
        .Range("D28").MergeArea.ClearContents   ' Bersihkan Deskripsi
        .Range("G28").MergeArea.ClearContents
        .Range("G24").MergeArea.ClearContents
        .Range("G16").MergeArea.ClearContents
        
        ' UX Flow: Kembalikan kursor manis   ke sel awal D12
        .Range("D12").MergeArea.Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 7. NOTIFIKASI SUKSES   PT MANIFESTASI MASA DEPAN
    MsgBox "Transaksi Pengeluaran Beban (" & V_StatusBayar & ") sebesar Rp" & Format(V_Nominal, "#,##0") & " telah berhasil dicatat!", _
            vbInformation, "Penyimpanan Sukses"
End Sub

