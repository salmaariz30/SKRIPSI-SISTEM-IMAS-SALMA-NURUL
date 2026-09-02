Attribute VB_Name = "Module61"
Sub SimpanRincianProdukTerjual()
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarTerjual As Worksheet
    Dim LembarMaster As Worksheet
    Dim LembarPiutang As Worksheet
    
    Dim TabelTerjual As ListObject
    Dim TabelMaster As ListObject
    Dim TabelPiutang As ListObject
    Dim BarisBaru As ListRow
    Dim BarisPiutang As ListRow
    
    Dim NoUrutTerakhir As Long
    Dim NoUrutPiutang As Long
    Dim IsBarisPertamaKosong As Boolean
    Dim IsBarisPiutangKosong As Boolean
    Const PWD As String = "IMAS"
    
    ' VARIABEL AMBIL DATA UTAMA FORM
    Dim V_Tanggal As Date
    Dim V_NoBukti As String
    Dim V_MetodePencatatan As String
    Dim V_MetodePembayaran As String
    Dim V_NamaPelangan As String
    Dim V_GrandTotalNota As Double
    Dim V_JumlahTerbayarAwal As Double
    Dim V_TglJatuhTempo As Variant
    
    ' VARIABEL LOOPING DATA
    Dim BarisMulai As Long
    Dim BarisAkhir As Long
    Dim i As Long
    
    ' VARIABEL PURE DATA ITEM
    Dim V_NamaProduk As String
    Dim V_Qty As Double
    Dim V_Kategori As String
    Dim V_Satuan As String
    Dim V_HargaJual As Double
    Dim V_Total As Double
    
    ' VARIABEL PENCARIAN MASTER (VLOOKUP INTERNAL VBA)
    Dim RentangCari As Range
    Dim KolomKategori As Long
    Dim KolomSatuan As Long
    Dim KolomHarga As Long
    Dim MatchIndex As Variant
    
    ' 1. SETTING SHEET & TABEL TARGET RESMI
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    Set LembarTerjual = ThisWorkbook.Sheets("PENDAPATAN_DATA TERJUAL")
    
    On Error Resume Next
    Set TabelTerjual = LembarTerjual.ListObjects("TabelPenjualanProduk")
    
    ' Mencari tabel master produk di seluruh sheets
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.ListObjects.Count > 0 Then
            Set TabelMaster = Sh.ListObjects("TabelProduk")
            If Not TabelMaster Is Nothing Then Exit For
        End If
    Next Sh
    On Error GoTo 0
    
    ' Validasi Proteksi Sistem Awal
    If TabelTerjual Is Nothing Then
        MsgBox "VALIDASI GAGAL! [Table Error]" & vbCrLf & _
               "'TabelPenjualanProduk' tidak ditemukan di sheet DATA TERJUAL!", vbCritical, "Eksekusi Gagal"
        Exit Sub
    End If
    If TabelMaster Is Nothing Then
        MsgBox "VALIDASI GAGAL! [Table Error]" & vbCrLf & _
               "Master data 'TabelProduk' tidak ditemukan di workbook ini!", vbCritical, "Eksekusi Gagal"
        Exit Sub
    End If
    
    ' 2. SETUP RENTANG INDEX MASTER TABEL PRODUK
    Set RentangCari = TabelMaster.ListColumns("Nama Produk").DataBodyRange
    KolomKategori = TabelMaster.ListColumns("Kategori Produk").Index
    KolomSatuan = TabelMaster.ListColumns("Satuan").Index
    KolomHarga = TabelMaster.ListColumns("Harga Jual Produk").Index
    
    ' 3. AMBIL DATA HEADER UMUM DARI FORM INPUT
    With LembarForm
        V_Tanggal = .Range("C11").MergeArea.Cells(1, 1).Value
        V_NoBukti = .Range("I15").MergeArea.Cells(1, 1).Value
        V_MetodePencatatan = .Range("F15").MergeArea.Cells(1, 1).Value
        V_MetodePembayaran = Trim(.Range("I11").MergeArea.Cells(1, 1).Value)
        V_NamaPelangan = Trim(.Range("C15").MergeArea.Cells(1, 1).Value)
        
        ' Data Khusus Tambahan Piutang
        V_JumlahTerbayarAwal = Val(.Range("L19").MergeArea.Cells(1, 1).Value)
        V_TglJatuhTempo = .Range("I19").MergeArea.Cells(1, 1).Value
        
        ' --- LOGIKA PENENTUAN RANGE BERDASARKAN METODE PENCATATAN ---
        If Trim(V_MetodePencatatan) = "Per Transaksi" Then
            BarisMulai = 24
            BarisAkhir = 38
            V_GrandTotalNota = Val(.Range("L42").MergeArea.Cells(1, 1).Value)
        ElseIf Trim(V_MetodePencatatan) = "Rekap Harian" Then
            BarisMulai = 54
            BarisAkhir = 83
            V_GrandTotalNota = Val(.Range("L87").Value)
        Else
            Exit Sub
        End If
    End With
    
    ' Validasi Khusus Metode Pembayaran Kredit
    If LCase(V_MetodePembayaran) = "kredit" Then
        If V_NamaPelangan = "" Then
            MsgBox "VALIDASI GAGAL! [Piutang Cacat Data]" & vbCrLf & _
                   "Metode Pembayaran bertuliskan 'Kredit', tetapi Nama Pelanggan di Cell C15 Kosong!", vbCritical, "Buku Piutang Menolak"
            Exit Sub
        End If
        If V_TglJatuhTempo = "" Or Not IsDate(V_TglJatuhTempo) Then
            MsgBox "VALIDASI GAGAL! [Jatuh Tempo Error]" & vbCrLf & _
                   "Metode Pembayaran Kredit wajib menyertakan Tanggal Jatuh Tempo yang valid di Cell I19!", vbCritical, "Buku Piutang Menolak"
            Exit Sub
        End If
    End If
    
    ' Matikan screen updating
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' BUKA PROTEKSI SHEET TARGET UTAMA BEFORE INJECT
    LembarTerjual.Unprotect Password:=PWD
    
    ' 4. PROSES LOOPING EKSPOR UTAMA DATA TERJUAL
    For i = BarisMulai To BarisAkhir
        V_NamaProduk = LembarForm.Cells(i, "D").MergeArea.Cells(1, 1).Value
        V_Qty = Val(LembarForm.Cells(i, "F").Value)
        
        If Trim(V_NamaProduk) <> "" And V_Qty > 0 Then
            MatchIndex = Application.Match(V_NamaProduk, RentangCari, 0)
            
            If Not IsError(MatchIndex) Then
                V_Kategori = TabelMaster.DataBodyRange.Cells(MatchIndex, KolomKategori).Value
                V_Satuan = TabelMaster.DataBodyRange.Cells(MatchIndex, KolomSatuan).Value
                V_HargaJual = Val(TabelMaster.DataBodyRange.Cells(MatchIndex, KolomHarga).Value)
            Else
                V_Kategori = "Unregistered"
                V_Satuan = "Pcs"
                V_HargaJual = 0
            End If
            
            V_Total = V_Qty * V_HargaJual
            
            IsBarisPertamaKosong = False
            On Error Resume Next
            If TabelTerjual.ListRows.Count = 0 Then
                IsBarisPertamaKosong = True
                NoUrutTerakhir = 0
            ElseIf TabelTerjual.ListRows.Count = 1 And (TabelTerjual.DataBodyRange.Cells(1, 1).Value = "" Or TabelTerjual.DataBodyRange.Cells(1, 2).Value = "") Then
                IsBarisPertamaKosong = True
                NoUrutTerakhir = 0
            Else
                NoUrutTerakhir = Application.WorksheetFunction.Max(TabelTerjual.ListColumns(1).DataBodyRange)
            End If
            On Error GoTo 0
            
            If IsBarisPertamaKosong Then
                If TabelTerjual.ListRows.Count = 0 Then TabelTerjual.ListRows.Add
                Set BarisBaru = TabelTerjual.ListRows(1)
            Else
                Set BarisBaru = TabelTerjual.ListRows.Add
            End If
            
            With BarisBaru
                .Range(1) = NoUrutTerakhir + 1
                .Range(2) = V_Tanggal
                .Range(3) = V_NoBukti
                .Range(4) = V_NamaProduk
                .Range(5) = V_Kategori
                .Range(6) = V_Qty
                .Range(7) = V_Satuan
                .Range(8) = V_HargaJual
                .Range(9) = V_Total
                
                With .Range.Font
                    .Name = "Segoe UI"
                    .Size = 10
                    .Color = vbBlack
                End With
                .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
                .Range(9).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            End With
        End If
    Next i
    
    ' ====================================================================
    ' 5. JALUR REKAYASA SISTEM BUKU PIUTANG + AGING BUCKET OTOMATIS
    ' ====================================================================
    If LCase(V_MetodePembayaran) = "kredit" Then
        
        On Error Resume Next
        Set LembarPiutang = ThisWorkbook.Sheets("PIUTANG_BUKU PIUTANG")
        Set TabelPiutang = LembarPiutang.ListObjects("TabelBukuPiutang")
        On Error GoTo 0
        
        If LembarPiutang Is Nothing Or TabelPiutang Is Nothing Then
            ' Tetap kunci kembali sheet terjual jika piutang gagal dimuat di tengah jalan
            LembarTerjual.Protect Password:=PWD, AllowFiltering:=True
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            MsgBox "ALARM PIUTANG! 'TabelBukuPiutang' gagal dimuat.", vbCritical, "Sistem Patah Kaki"
            Exit Sub
        End If
        
        ' BUKA PROTEKSI SHEET BUKU PIUTANG BEFORE INJECT
        LembarPiutang.Unprotect Password:=PWD
        
        IsBarisPiutangKosong = False
        On Error Resume Next
        If TabelPiutang.ListRows.Count = 0 Then
            IsBarisPiutangKosong = True
            NoUrutPiutang = 0
        ElseIf TabelPiutang.ListRows.Count = 1 And (TabelPiutang.DataBodyRange.Cells(1, 1).Value = "" Or TabelPiutang.DataBodyRange.Cells(1, 2).Value = "") Then
            IsBarisPiutangKosong = True
            NoUrutPiutang = 0
        Else
            NoUrutPiutang = Application.WorksheetFunction.Max(TabelPiutang.ListColumns(1).DataBodyRange)
        End If
        On Error GoTo 0
        
        If IsBarisPiutangKosong Then
            If TabelPiutang.ListRows.Count = 0 Then TabelPiutang.ListRows.Add
            Set BarisPiutang = TabelPiutang.ListRows(1)
        Else
            Set BarisPiutang = TabelPiutang.ListRows.Add
        End If
        
        ' --- INJEKSI DATABASE BUKU PIUTANG TERBARU DENGAN AGING BUCKET ---
        With BarisPiutang
            .Range(1) = V_NoBukti                                               ' KOLOM 2: NO. INVOICE
            .Range(2) = V_Tanggal                                               ' KOLOM 3: TANGGAL
            .Range(3) = V_NamaPelangan                                          ' KOLOM 4: NAMA PELANGGAN
            .Range(4) = V_GrandTotalNota                                        ' KOLOM 5: TOTAL PENJUALAN
            .Range(5) = 0                                                       ' KOLOM 6: POTONGAN & RETUR
            .Range(6).Formula2R1C1 = "=[@[TOTAL PENJUALAN]]-[@[POTONGAN PENJUALAN & RETUR]]" ' KOLOM 7: TOTAL PIUTANG (NETTO)
            .Range(7) = V_JumlahTerbayarAwal                                    ' KOLOM 8: JUMLAH TERBAYAR
            .Range(8).Formula2R1C1 = "=[@[TOTAL PIUTANG (NETTO)]]-[@[JUMLAH TERBAYAR]]" ' KOLOM 9: SISA PIUTANG
            .Range(9) = CDate(V_TglJatuhTempo)                                  ' KOLOM 10: TANGGAL JATUH TEMPO
            
            ' KOLOM 11: RUMUS UMUR PIUTANG SMART (Hentikan jika sisa piutang Rp 0)
            .Range(10).Formula2R1C1 = "=IF([@[SISA PIUTANG]]=0, 0, TODAY()-[@[TANGGAL JATUH TEMPO]])"
            
            ' ?? KOLOM 12 (BARU): RUMUS AGING BUCKET LOKAL PT MMD ??
            ' Mengunci murni relasi terdaftar ke kolom [UMUR PIUTANG] di baris berjalan
            .Range(11).Formula2R1C1 = "=IF([@[UMUR PIUTANG]]<=0, ""Belum Jatuh Tempo"", IF([@[UMUR PIUTANG]]<=30, ""1-30 Hari"", IF([@[UMUR PIUTANG]]<=60, ""31-60 Hari"", "">60 Hari"")))"
            
            ' KOLOM 13 (GESER): RUMUS STATUS PELUNASAN
            .Range(12).Formula2R1C1 = "=IF([@[SISA PIUTANG]]=0, ""Lunas"", ""Belum Lunas"")"
            
            ' Format Estetika Seragam: Segoe UI 9 Black
            With .Range.Font
                .Name = "Segoe UI"
                .Size = 10
                .Color = vbBlack
            End With
            
            ' Format Accounting Rupiah Khas USA Regional Settings
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(6).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(7).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(9).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            
            .Range(11).NumberFormat = "#,##0"         ' Format Hari
            .Range(13).HorizontalAlignment = xlCenter ' Center Status
        End With
        
        ' KUNCI KEMBALI SHEET BUKU PIUTANG
        LembarPiutang.Protect Password:=PWD, AllowFiltering:=True
    End If
    
    ' KUNCI KEMBALI SHEET DATA TERJUAL
    LembarTerjual.Protect Password:=PWD, AllowFiltering:=True
    
    ' Nyalakan grafis
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub

