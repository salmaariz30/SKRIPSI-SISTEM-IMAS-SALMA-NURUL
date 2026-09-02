Attribute VB_Name = "Module78"
Sub GenerateNoBuktiBarang()
    ' ====================================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI BARANG PERSEDIAAN (VERSI KORPORAT)
    ' Format Output: INVT-[DDMM]-[3 Digit Urutan] (Contoh: INVT-1106-001)
    ' SPECIAL EDITION: ANTI-LAG & COMPATIBLE WITH USA REGIONAL SETTINGS
    ' Sesuai Tata Letak Kolom Mutlak: TabelPembelianPersediaan (Kolom ke-3)
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarTabel As Worksheet
    Dim TabelData As ListObject
    Dim KolomBukti As Range
    Dim TglBlnStr As String
    Dim FormatPrefix As String
    Dim UrutanTerakhir As Long
    Dim Baris As Range
    Dim IsiSel As String
    Dim NomorStr As String
    Dim AngkaUrut As Long
    
    ' 1. PENGATURAN AWAL & TARGET FORM
    Set LembarForm = ThisWorkbook.Sheets("PERSEDIAAN_INPUT BARANG")
    Set LembarTabel = ThisWorkbook.Sheets("PERSEDIAAN_TABEL PEMBELIAN")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari tanggal hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "INVT-" & TglBlnStr & "-" ' Hasilnya misal hari ini: "INVT-1106-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database pembelian persediaan
    On Error Resume Next
    Set TabelData = LembarTabel.ListObjects("TabelPembelianPersediaan")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Pembelian
    If TabelData Is Nothing Then
        MsgBox "Error: 'TabelPembelianPersediaan' tidak ditemukan di sheet terkait!", vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE
    If TabelData.ListRows.Count > 0 Then
        ' Menyasar mutlak kolom ke-3 dari tabel pembelian persediaan
        Set KolomBukti = TabelData.ListColumns(2).DataBodyRange
        
        If Not KolomBukti Is Nothing Then
            For Each Baris In KolomBukti
                IsiSel = Trim(CStr(Baris.Value))
                
                ' Memeriksa apakah ada kode bukti yang prefix tanggal/bulannya sama dengan hari ini
                If Left(IsiSel, Len(FormatPrefix)) = FormatPrefix Then
                    ' Mengambil 3 digit angka urut terakhir di paling kanan
                    NomorStr = Right(IsiSel, 3)
                    If IsNumeric(NomorStr) Then
                        AngkaUrut = CLng(NomorStr)
                        If AngkaUrut > UrutanTerakhir Then
                            UrutanTerakhir = AngkaUrut
                        End If
                    End If
                End If
            Next Baris
        End If
    End If
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT BARU
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Masukkan hasil akhir secara absolut ke sel D12 yang dimerge (Anti-Error Merged Cells)
    With LembarForm.Range("D12").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat ketat ke sel input selanjutnya (H12)
    Application.GoTo LembarForm.Range("H12")
    
    End Sub

