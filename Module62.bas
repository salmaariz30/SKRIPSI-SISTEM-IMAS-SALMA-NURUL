Attribute VB_Name = "Module62"
Sub JurnalUmumInputPenjualan()
    ' ====================================================================
    Dim tblJurnal As ListObject, tblCOA As ListObject, tblProduk As ListObject
    Dim tblStokDagang As ListObject, tblStockJadi As ListObject
    Dim wsInput As Worksheet, Sh As Worksheet
    Dim wsJurnal As Worksheet
    Const PWD As String = "IMAS"
    
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    Set tblStokDagang = Range("TabelStokDagang").ListObject
    Set tblStockJadi = Range("TabelStockJadi").ListObject
    Set wsInput = Sheets("PENDAPATAN_INPUT PENJUALAN")
    
    ' Cari TabelProduk mandiri lintas sheet
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.ListObjects.Count > 0 Then
            Set tblProduk = Sh.ListObjects("TabelProduk")
            If Not tblProduk Is Nothing Then Exit For
        End If
    Next Sh
    On Error GoTo 0
    
    ' Cek fisik keutuhan tabel sistem
    If wsInput Is Nothing Then
        MsgBox "VALIDASI GAGAL! [Sheet Error]" & vbCrLf & _
               "Sheet 'PENDAPATAN_INPUT PENJUALAN' tidak ditemukan!", vbCritical, "Error Kontrol Sistem"
        Exit Sub
    End If
    If tblJurnal Is Nothing Or tblCOA Is Nothing Or tblProduk Is Nothing Then
        MsgBox "VALIDASI GAGAL! [Table Master Error]" & vbCrLf & _
               "Pastikan TabelJurnalUmum, TabelDataCOA, dan TabelProduk terdaftar!", vbCritical, "Error Kontrol Sistem"
        Exit Sub
    End If
    If tblStokDagang Is Nothing Or tblStockJadi Is Nothing Then
        MsgBox "VALIDASI GAGAL! [Perpetual Table Error]" & vbCrLf & _
               "TabelStokDagang atau TabelStockJadi tidak ditemukan!", vbCritical, "Error Sistem Perpetual"
        Exit Sub
    End If

    ' --------------------------------------------------------------------
    ' VALIDASI PILAR 2: KELENGKAPAN HEADER NOTA UTAMA (CELL KUNCI)
    ' --------------------------------------------------------------------
    Dim V_NoBukti As String: V_NoBukti = Trim(wsInput.Range("I15").MergeArea.Cells(1, 1).Value)
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("C11").MergeArea.Cells(1, 1).Value
    Dim V_NamaPelangan As String: V_NamaPelangan = Trim(wsInput.Range("C15").MergeArea.Cells(1, 1).Value)
    Dim V_Deskripsi As String: V_Deskripsi = Trim(wsInput.Range("C19").MergeArea.Cells(1, 1).Value)
    Dim V_MetodePencatatan As String: V_MetodePencatatan = Trim(wsInput.Range("F15").MergeArea.Cells(1, 1).Value)
    Dim V_MetodePembayaran As String: V_MetodePembayaran = Trim(wsInput.Range("I11").MergeArea.Cells(1, 1).Value)
    Dim V_NamaBankPilihan As String: V_NamaBankPilihan = Trim(wsInput.Range("L11").MergeArea.Cells(1, 1).Value)
    
    If V_Tanggal = "" Or Not IsDate(V_Tanggal) Then
        MsgBox "VALIDASI GAGAL! Tanggal transaksi tidak valid!", vbExclamation, "Data Tidak Lengkap"
        Exit Sub
    End If
    If V_NoBukti = "" Then
        MsgBox "VALIDASI GAGAL! Nomor Bukti Nota kosong!", vbExclamation, "Data Tidak Lengkap"
        Exit Sub
    End If

    ' --------------------------------------------------------------------
    ' VALIDASI PILAR 3: PENENTUAN RANGE KOORDINAT KERANJANG BELANJA
    ' --------------------------------------------------------------------
    Dim V_Pajak As Double, V_BiayaPengiriman As Double, V_TotalPeroleh As Double
    Dim BarisMulai As Long, BarisAkhir As Long
    
    With wsInput
        If V_MetodePencatatan = "Per Transaksi" Then
            BarisMulai = 24: BarisAkhir = 38
            V_Pajak = Val(.Range("L40").MergeArea.Cells(1, 1).Value)
            V_BiayaPengiriman = Val(.Range("L41").MergeArea.Cells(1, 1).Value)
            V_TotalPeroleh = Val(.Range("L42").MergeArea.Cells(1, 1).Value)
        ElseIf V_MetodePencatatan = "Rekap Harian" Then
            BarisMulai = 54: BarisAkhir = 83
            V_Pajak = Val(.Range("L85").Value)
            V_BiayaPengiriman = Val(.Range("L86").Value)
            V_TotalPeroleh = Val(.Range("L87").Value)
        Else
            MsgBox "VALIDASI GAGAL! Metode Pencatatan salah!", vbCritical, "Logika Patah"
            Exit Sub
        End If
    End With
    
    If V_TotalPeroleh <= 0 Then
        MsgBox "VALIDASI GAGAL! Total Perolehan Transaksi bernilai Rp0 atau Minus.", vbExclamation, "Aborting Process"
        Exit Sub
    End If

    ' --------------------------------------------------------------------
    ' VALIDASI PILAR 4 & REKAP HPP: PERKALIAN MODAL SATUAN * QTY NOTA
    ' --------------------------------------------------------------------
    Dim TotalProduksi As Double: TotalProduksi = 0
    Dim TotalDagang As Double: TotalDagang = 0
    Dim TotalJasa As Double: TotalJasa = 0
    
    ' VARIABEL UNTUK MENAMPUNG TOTAL AKHIR NOMINAL HPP YANG SUDAH DIKALIKAN QTY
    Dim TotalHPP_Dagang As Double: TotalHPP_Dagang = 0
    Dim TotalHPP_Produksi As Double: TotalHPP_Produksi = 0
    
    Dim ProductRange As Range: Set ProductRange = tblProduk.ListColumns("Nama Produk").DataBodyRange
    Dim ColKategori As Long: ColKategori = tblProduk.ListColumns("Kategori Produk").Index
    
    Dim idx As Long, MatchIdx As Variant, MatchStok As Variant
    Dim ItemNama As String, ItemTotalUang As Double, ItemQTY As Double
    Dim ItemKategori As String, HPP_Satuan As Double, TotalHPP_Baris As Double
    Dim ValidatedItemsCount As Long: ValidatedItemsCount = 0
    
    For idx = BarisMulai To BarisAkhir
        ItemNama = Trim(wsInput.Cells(idx, "D").MergeArea.Cells(1, 1).Value)
        ItemTotalUang = Val(wsInput.Cells(idx, "L").MergeArea.Cells(1, 1).Value)
        ItemQTY = Val(wsInput.Cells(idx, "F").MergeArea.Cells(1, 1).Value)
        
        If ItemNama <> "" Then
            ValidatedItemsCount = ValidatedItemsCount + 1
            
            If ItemTotalUang <= 0 Then
                MsgBox "VALIDASI BARIS GAGAL! Nilai Uang di Kolom L kosong/0!", vbCritical, "Data Baris Cacat"
                Exit Sub
            End If
            If ItemQTY <= 0 Then
                MsgBox "VALIDASI BARIS GAGAL! QTY untuk '" & ItemNama & "' di Kolom H tidak boleh 0 atau kosong!", vbCritical, "Data QTY Cacat"
                Exit Sub
            End If
            
            ' Scan silang ke master produk
            MatchIdx = Application.Match(ItemNama, ProductRange, 0)
            
            If Not IsError(MatchIdx) Then
                ItemKategori = tblProduk.DataBodyRange.Cells(MatchIdx, ColKategori).Value
                HPP_Satuan = 0
                TotalHPP_Baris = 0
                
                ' A. JIKA KATEGORI PRODUKSI
                If InStr(1, ItemKategori, "Produksi", vbTextCompare) > 0 Then
                    TotalProduksi = TotalProduksi + ItemTotalUang
                    
                    ' Lacak modal per pcs di TabelStockJadi (Kolom 3 = Nama, Kolom 5 = Harga Modal)
                    MatchStok = Application.Match(ItemNama, tblStockJadi.ListColumns(3).DataBodyRange, 0)
                    If IsError(MatchStok) Then
                        MsgBox "SISTEM GAGAL! Produk '" & ItemNama & "' tidak terdaftar di TabelStockJadi!", vbCritical, "Stok Error"
                        Exit Sub
                    End If
                    
                    HPP_Satuan = Val(tblStockJadi.DataBodyRange.Cells(MatchStok, 5).Value)
                    TotalHPP_Baris = HPP_Satuan * ItemQTY ' <-- LOGIKA BARU: MODAL DIKALI QTY
                    TotalHPP_Produksi = TotalHPP_Produksi + TotalHPP_Baris ' Jumlahkan akumulasi totalnya
                    
                ' B. JIKA KATEGORI DAGANG
                ElseIf InStr(1, ItemKategori, "Dagang", vbTextCompare) > 0 Then
                    TotalDagang = TotalDagang + ItemTotalUang
                    
                    ' Lacak modal per pcs di TabelStokDagang (Kolom 3 = Nama, Kolom 5 = Harga Modal)
                    MatchStok = Application.Match(ItemNama, tblStokDagang.ListColumns(3).DataBodyRange, 0)
                    If IsError(MatchStok) Then
                        MsgBox "SISTEM GAGAL! Produk '" & ItemNama & "' tidak terdaftar di TabelStokDagang!", vbCritical, "Stok Error"
                        Exit Sub
                    End If
                    
                    HPP_Satuan = Val(tblStokDagang.DataBodyRange.Cells(MatchStok, 5).Value)
                    TotalHPP_Baris = HPP_Satuan * ItemQTY ' <-- LOGIKA BARU: MODAL DIKALI QTY
                    TotalHPP_Dagang = TotalHPP_Dagang + TotalHPP_Baris ' Jumlahkan akumulasi totalnya
                    
                ' C. JIKA KATEGORI JASA
                ElseIf InStr(1, ItemKategori, "Jasa", vbTextCompare) > 0 Then
                    TotalJasa = TotalJasa + ItemTotalUang ' Jasa tidak punya stok fisik, abaikan HPP
                Else
                    MsgBox "PROSES BERHENTI! Kategori master data rusak.", vbCritical, "Master Data Rusak"
                    Exit Sub
                End If
            End If
        End If
    Next idx
    
    If ValidatedItemsCount = 0 Then
        MsgBox "VALIDASI GAGAL! Nota Kosong.", vbExclamation, "Nota Kosong"
        Exit Sub
    End If

    ' --------------------------------------------------------------------
    ' VALIDASI PILAR 5: METODE PEMBAYARAN
    ' --------------------------------------------------------------------
    Dim namaAkunDebit1 As String, namaAkunKreditOngkir As String
    Select Case V_MetodePembayaran
        Case "Tunai": namaAkunDebit1 = "Kas": namaAkunKreditOngkir = "Kas"
        Case "Kredit": namaAkunDebit1 = "Piutang Usaha": namaAkunKreditOngkir = "Utang Lain-Lain"
        Case "Transfer Bank"
            If V_NamaBankPilihan = "" Then
                MsgBox "VALIDASI GAGAL! Nama Akun Bank di Cell L11 kosong!", vbCritical, "Form Input Bolong"
                Exit Sub
            End If
            namaAkunDebit1 = V_NamaBankPilihan: namaAkunKreditOngkir = V_NamaBankPilihan
        Case Else: MsgBox "VALIDASI GAGAL! Teks Metode Pembayaran salah!", vbCritical, "Patah Logika Form": Exit Sub
    End Select

    ' --------------------------------------------------------------------
    ' VALIDASI PILAR 6: VERIFIKASI SEKALIGUS PENARIKAN KODE COA
    ' --------------------------------------------------------------------
    Dim barisDr1 As Variant, barisKrPajak As Variant, barisKrOngkir As Variant
    Dim kodeDr1 As Variant, kodeKrPajak As Variant, kodeKrOngkir As Variant
    Dim barisProd As Variant, barisDagang As Variant, barisJasa As Variant
    Dim kodeProd As Variant, kodeDagang As Variant, kodeJasa As Variant
    
    Dim kodeHPP_Dagang As Variant, kodePersediaan_Dagang As Variant
    Dim kodeHPP_Produksi As Variant, kodePersediaan_Produksi As Variant
    
    On Error Resume Next
    barisDr1 = Application.Match(namaAkunDebit1, tblCOA.ListColumns(2).DataBodyRange, 0)
    kodeDr1 = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDr1)
    
    If TotalProduksi > 0 Then
        kodeProd = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Pendapatan Jual Barang Produksi", tblCOA.ListColumns(2).DataBodyRange, 0))
        kodeHPP_Produksi = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Harga Pokok Produksi", tblCOA.ListColumns(2).DataBodyRange, 0))
        kodePersediaan_Produksi = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Persediaan Barang Jadi", tblCOA.ListColumns(2).DataBodyRange, 0))
    End If
    If TotalDagang > 0 Then
        kodeDagang = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Pendapatan Jual Barang Dagang", tblCOA.ListColumns(2).DataBodyRange, 0))
        kodeHPP_Dagang = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Harga Pokok Barang Dagang", tblCOA.ListColumns(2).DataBodyRange, 0))
        kodePersediaan_Dagang = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Persediaan Barang Dagang", tblCOA.ListColumns(2).DataBodyRange, 0))
    End If
    If TotalJasa > 0 Then
        kodeJasa = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Pendapatan Jasa", tblCOA.ListColumns(2).DataBodyRange, 0))
    End If
    If V_Pajak > 0 Then
        kodeKrPajak = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match("Utang Pajak", tblCOA.ListColumns(2).DataBodyRange, 0))
    End If
    If V_BiayaPengiriman > 0 Then
        kodeKrOngkir = Application.Index(tblCOA.ListColumns(1).DataBodyRange, Application.Match(namaAkunKreditOngkir, tblCOA.ListColumns(2).DataBodyRange, 0))
    End If
    On Error GoTo 0
    
    If IsError(kodeDr1) Then MsgBox "POSTING GAGAL! Akun '" & namaAkunDebit1 & "' tidak ada di COA!", vbCritical, "Meleset": Exit Sub

    ' --------------------------------------------------------------------
    ' PILAR 7: EKSEKUSI JURNAL UMUM (SUNTIKAN DATA AMAN)
    ' --------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' DETEKSI & BUKA PROTEKSI SHEET JURNAL UMUM SECARA DINAMIS
    Set wsJurnal = tblJurnal.Parent
    wsJurnal.Unprotect Password:=PWD
    
    ' --- 1. DEBIT UTAMA (KAS/BANK/PIUTANG) ---
    With tblJurnal.ListRows.Add(AlwaysInsert:=True)
        .Range(4).NumberFormat = "@"
        .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = V_Deskripsi & " (" & V_NamaPelangan & ")"
        .Range(4) = CStr(kodeDr1): .Range(5) = namaAkunDebit1: .Range(6) = V_TotalPeroleh: .Range(7) = 0
        .Range(8) = "Pendapatan Penjualan": .Range(9) = "Tidak"
    End With
    
    ' --- 2. KREDIT: PENDAPATAN PRODUKSI ---
    If TotalProduksi > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Pendapatan Produk Hasil Produksi"
            .Range(4) = CStr(kodeProd): .Range(5) = "Pendapatan Jual Barang Produksi": .Range(6) = 0: .Range(7) = TotalProduksi
            .Range(8) = "Pendapatan Penjualan": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- 3. KREDIT: PENDAPATAN DAGANG ---
    If TotalDagang > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Pendapatan Penjualan Barang Dagangan"
            .Range(4) = CStr(kodeDagang): .Range(5) = "Pendapatan Jual Barang Dagang": .Range(6) = 0: .Range(7) = TotalDagang
            .Range(8) = "Pendapatan Penjualan": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- 4. KREDIT: PENDAPATAN JASA ---
    If TotalJasa > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Pendapatan Penyerahan Jasa Layanan"
            .Range(4) = CStr(kodeJasa): .Range(5) = "Pendapatan Jasa": .Range(6) = 0: .Range(7) = TotalJasa
            .Range(8) = "Pendapatan Penjualan": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- 5. KREDIT: UTANG PAJAK ---
    If V_Pajak > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Pajak Penjualan / PPN " & V_NoBukti
            .Range(4) = CStr(kodeKrPajak): .Range(5) = "Utang Pajak": .Range(6) = 0: .Range(7) = V_Pajak
            .Range(8) = "Pendapatan Penjualan": .Range(9) = "Tidak"
        End With
    End If
    
    ' --- 6. KREDIT: TRANSIT ONGKIR ---
    If V_BiayaPengiriman > 0 Then
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Transit Biaya Kurir Pihak Ketiga (" & V_MetodePembayaran & ")"
            .Range(4) = CStr(kodeKrOngkir): .Range(5) = namaAkunKreditOngkir: .Range(6) = 0: .Range(7) = V_BiayaPengiriman
            .Range(8) = "Pendapatan Penjualan": .Range(9) = "Tidak"
        End With
    End If

    ' ====================================================================
    ' AMBLES JURNAL PERPETUAL REVISI TOTAL (HPP YANG SUDAH DIKALIKAN QTY)
    ' ====================================================================
    
    ' A. JURNAL PERPETUAL BARANG DAGANG
    If TotalDagang > 0 And TotalHPP_Dagang > 0 Then
        ' DEBIT: Harga Pokok Barang Dagang
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "HPP Barang Dagang (Total " & TotalHPP_Dagang & ") - Nota " & V_NoBukti
            .Range(4) = CStr(kodeHPP_Dagang): .Range(5) = "Harga Pokok Barang Dagang": .Range(6) = TotalHPP_Dagang: .Range(7) = 0
            .Range(8) = "Harga Pokok Penjualan": .Range(9) = "Tidak"
        End With
        ' KREDIT: Persediaan Barang Dagang
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Pengurangan Stok Barang Dagang - Nota " & V_NoBukti
            .Range(4) = CStr(kodePersediaan_Dagang): .Range(5) = "Persediaan Barang Dagang": .Range(6) = 0: .Range(7) = TotalHPP_Dagang
            .Range(8) = "Harga Pokok Penjualan": .Range(9) = "Tidak"
        End With
    End If
    
    ' B. JURNAL PERPETUAL BARANG PRODUKSI
    If TotalProduksi > 0 And TotalHPP_Produksi > 0 Then
        ' DEBIT: Harga Pokok Produksi
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "HPP Produk Jadi (Total " & TotalHPP_Produksi & ") - Nota " & V_NoBukti
            .Range(4) = CStr(kodeHPP_Produksi): .Range(5) = "Harga Pokok Produksi": .Range(6) = TotalHPP_Produksi: .Range(7) = 0
            .Range(8) = "Harga Pokok Penjualan": .Range(9) = "Tidak"
        End With
        ' KREDIT: Persediaan Barang Jadi
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(4).NumberFormat = "@"
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Pengurangan Stok Produk Jadi - Nota " & V_NoBukti
            .Range(4) = CStr(kodePersediaan_Produksi): .Range(5) = "Persediaan Barang Jadi": .Range(6) = 0: .Range(7) = TotalHPP_Produksi
            .Range(8) = "Harga Pokok Penjualan": .Range(9) = "Tidak"
        End With
    End If
    
    ' --------------------------------------------------------------------
    ' KUNCI KEMBALI SHEET JURNAL UMUM
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    MsgBox "Data Penjualan Berhasil Disimpan!", vbInformation, "Penyimpanan Berhasil"
End Sub

