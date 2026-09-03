Attribute VB_Name = "Module106"
Sub JurnalUmumInputPelunasanUtang()
    ' ====================================================================
    ' STATUS: AUTO UNPROTECT / PROTECT JURNAL SHEET (PASSWORD: IMAS)
    ' ====================================================================
    
    ' 1. Set Sheet dan Tabel Target Resmi
    Dim tblJurnal As ListObject, tblCOA As ListObject, tblUtangUsaha As ListObject
    Dim wsInput As Worksheet
    Dim wsJurnal As Worksheet
    Const PWD As String = "IMAS"
    
    Set wsInput = Sheets("UTANG_INPUT PELUNASAN")
    
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    Set tblUtangUsaha = Range("TabelUtangUsaha").ListObject
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Utama
    If tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: 'TabelJurnalUmum' atau 'TabelDataCOA' tidak ditemukan!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' Ambil worksheet induk tempat Tabel Jurnal Umum berada
    Set wsJurnal = tblJurnal.Parent
    
    ' 2. Tarik Data dari Form Input Pelunasan (Mendukung Merged Cells)
    Dim V_Status As String: V_Status = wsInput.Range("G11").MergeArea.Cells(1, 1).Value ' "Biaya Usaha" atau "Bank"
    Dim V_D11 As String: V_D11 = wsInput.Range("D11").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As String: V_NoBukti = wsInput.Range("G15").MergeArea.Cells(1, 1).Value
    Dim V_Tanggal As Date: V_Tanggal = wsInput.Range("D15").MergeArea.Cells(1, 1).Value
    Dim V_NominalPokok As Double: V_NominalPokok = Val(wsInput.Range("D19").MergeArea.Cells(1, 1).Value)
    Dim V_NominalBunga As Double: V_NominalBunga = Val(wsInput.Range("G19").MergeArea.Cells(1, 1).Value)
    Dim V_AkunKasBank As String: V_AkunKasBank = wsInput.Range("D23").MergeArea.Cells(1, 1).Value
    Dim V_DeskripsiUser As String: V_DeskripsiUser = wsInput.Range("D27").MergeArea.Cells(1, 1).Value
    
    ' Validasi Input Dasar
    If V_NoBukti = "" Or V_AkunKasBank = "" Or (V_NominalPokok <= 0 And V_NominalBunga <= 0) Then
        MsgBox "Gagal! Pastikan No. Bukti, Akun Kas/Bank, dan Nominal Pokok/Bunga sudah terisi.", vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 3. LOGIKA DETEKSI NAMA AKUN UTANG (DEBIT) BERDASARKAN STATUS G11
    Dim namaAkunUtang As String
    Dim barisUtang As Variant
    
    If UCase(Trim(V_Status)) = "BIAYA USAHA" Then
        If tblUtangUsaha Is Nothing Then
            MsgBox "Error: Status 'Biaya Usaha' dipilih, tapi 'TabelUtangUsaha' tidak ditemukan!", vbCritical, "Sistem Gagal"
            Exit Sub
        End If
        
        On Error Resume Next
        barisUtang = Application.Match(V_D11, tblUtangUsaha.ListColumns(2).DataBodyRange, 0)
        On Error GoTo 0
        
        If IsError(barisUtang) Then
            MsgBox "Gagal! Data '" & V_D11 & "' tidak ditemukan di kolom kedua TabelUtangUsaha.", vbCritical, "Data Tidak Cocok"
            Exit Sub
        End If
        
        namaAkunUtang = tblUtangUsaha.ListColumns(6).DataBodyRange.Cells(barisUtang, 1).Value
        
    ElseIf UCase(Trim(V_Status)) = "BANK" Then
        namaAkunUtang = "Utang Bank Jangka Panjang"
    Else
        MsgBox "Gagal! Status di G11 '" & V_Status & "' harus berisi 'Biaya Usaha' atau 'Bank'.", vbCritical, "Error Status"
        Exit Sub
    End If
    
    ' REVISI  : Nama akun bunga dikunci mati ke "Beban Bunga Bank"
    Dim namaAkunBunga As String: namaAkunBunga = "Beban Bunga Bank"
    
    ' 4. Ambil Semua Kode Akun dari Master COA
    Dim barisCOAUtang As Variant, barisCOAKas As Variant, barisCOABunga As Variant
    Dim kodeAkunUtang As Variant, kodeAkunKas As Variant, kodeAkunBunga As Variant
    
    On Error Resume Next
    barisCOAUtang = Application.Match(namaAkunUtang, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisCOAKas = Application.Match(V_AkunKasBank, tblCOA.ListColumns(2).DataBodyRange, 0)
    
    kodeAkunUtang = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisCOAUtang)
    kodeAkunKas = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisCOAKas)
    
    If V_NominalBunga > 0 Then
        barisCOABunga = Application.Match(namaAkunBunga, tblCOA.ListColumns(2).DataBodyRange, 0)
        kodeAkunBunga = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisCOABunga)
    End If
    On Error GoTo 0
    
    ' Proteksi validasi COA
    If IsError(kodeAkunUtang) Or IsError(kodeAkunKas) Or (V_NominalBunga > 0 And IsError(kodeAkunBunga)) Then
        MsgBox "Gagal Menjurnal! Pastikan nama akun '" & namaAkunUtang & "', '" & V_AkunKasBank & "', atau '" & namaAkunBunga & "' sudah terdaftar di COA!", vbCritical, "Error COA"
        Exit Sub
    End If
    
    ' 5. EKSEKUSI EXPORT KE JURNAL UMUM (SISTEM ANTI-KEDIP)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' ?? OPERASI BUKA GEMBOK: Buka proteksi sheet Jurnal Umum sebelum suntik data
    ' ====================================================================
    wsJurnal.Unprotect Password:=PWD
    
    Dim V_DeskripsiJurnal As String
    V_DeskripsiJurnal = "Pelunasan " & namaAkunUtang & IIf(V_DeskripsiUser <> "", " - " & V_DeskripsiUser, "")
    
    ' --- BARIS 1: DEBIT (Pokok Utang Bank/Usaha Berkurang di Kiri) ---
    If V_NominalPokok > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal
            .Range(2) = V_NoBukti
            .Range(3) = V_DeskripsiJurnal
            .Range(4) = kodeAkunUtang
            .Range(5) = namaAkunUtang
            .Range(6) = V_NominalPokok
            .Range(7) = 0
            .Range(8) = "Liabilitas"
            .Range(9) = "Tidak"
        End With
    End If
    
    ' --- BARIS 2: DEBIT TAMBAHAN (Beban Bunga Bank di Kiri) ---
    If V_NominalBunga > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal
            .Range(2) = V_NoBukti
            .Range(3) = "Pembayaran Beban Bunga atas " & V_DeskripsiJurnal
            .Range(4) = kodeAkunBunga
            .Range(5) = namaAkunBunga
            .Range(6) = V_NominalBunga
            .Range(7) = 0
            .Range(8) = "Beban Operasional"
            .Range(9) = "Tidak"
        End With
    End If
    
    ' --- BARIS 3: KREDIT (Kas Keluar Sebesar Total Pokok + Bunga di Kanan) ---
    Dim V_TotalBayar As Double: V_TotalBayar = V_NominalPokok + V_NominalBunga
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiJurnal
        .Range(4) = kodeAkunKas
        .Range(5) = V_AkunKasBank
        .Range(6) = 0
        .Range(7) = V_TotalBayar
        .Range(8) = "Kas & Bank"
        .Range(9) = "Tidak"
    End With
    
    ' ====================================================================
    ' ?? OPERASI RE-LOCK SYSTEM: Kunci kembali sheet Jurnal setelah eksekusi
    ' ====================================================================
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
    
    ' 6. FORM CLEANER (Otomatis Bersih Kilat)
    With wsInput
        .Range("G11").MergeArea.ClearContents
        .Range("D11").MergeArea.ClearContents
        .Range("G15").MergeArea.ClearContents
        .Range("D15").MergeArea.ClearContents
        .Range("D19").MergeArea.ClearContents
        .Range("G19").MergeArea.ClearContents
        .Range("D23").MergeArea.ClearContents
        .Range("G23").MergeArea.ClearContents
        .Range("D27").MergeArea.ClearContents
        
        ' Kembalikan Kursor ke Hulu Form (G15)
        .Activate
        .Range("G15").MergeArea.Cells(1, 1).Select
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 7. NOTIFIKASI BERHASIL
    MsgBox "Data Pelunasan Utang Berhasil Disimpan !", vbInformation, "Sistem Sukses"
End Sub
