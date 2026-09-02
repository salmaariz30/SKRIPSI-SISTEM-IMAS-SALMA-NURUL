Attribute VB_Name = "Module39"
Sub GenerateNoBuktiPenjualanOtomatis()
    ' ====================================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI PENJUALAN
    ' ====================================================================
    Dim LembarForm As Worksheet, LembarData As Worksheet
    Dim TabelData As ListObject
    Dim KolomBukti As Range
    Dim TglStrAsli As String
    Dim FormatPrefix As String, MetodeBayar As String
    Dim UrutanTerakhir As Long
    Dim Baris As Range
    Dim IsiSel As String, NomorStr As String
    Dim AngkaUrut As Long
    
    ' 1. PENGATURAN AWAL & DETEKSI METODE
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    
    ' Format tanggal hari ini: YYYYMMDD (Contoh: 20260602)
    TglStrAsli = Format(Date, "yyyymmdd")
    
    ' Ambil jenis metode pembayaran dari cell I11
    MetodeBayar = Trim(LembarForm.Range("I11").Value)
    
    ' Penentuan Format Prefix berdasarkan isi cell I11
    If LCase(MetodeBayar) = "kredit" Then
        FormatPrefix = "INV/" & TglStrAsli & "/"  ' Hasil: INV/20260602/
    Else
        FormatPrefix = "NOTA/" & TglStrAsli & "/" ' Hasil selain kredit: NOTA/20260602/
    End If
    
    UrutanTerakhir = 0
    
    ' Connect ke database penjualan target
    On Error Resume Next
    Set LembarData = ThisWorkbook.Sheets("PENDAPATAN_DATA PENDAPATAN")
    Set TabelData = LembarData.ListObjects("TabelRincianPenjualan")
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATABASE (Kolom Ke-2 dari TabelRincianPenjualan)
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' Mengunci Kolom Ke-2 (Tempat No Bukti/Invoice disimpan)
            Set KolomBukti = TabelData.ListColumns(2).DataBodyRange
            
            For Each Baris In KolomBukti
                IsiSel = Trim(CStr(Baris.Value))
                
                ' Memeriksa apakah awalan kode sama dengan prefix hari ini
                If Left(IsiSel, Len(FormatPrefix)) = FormatPrefix Then
                    ' Mengambil 3 digit angka terakhir di paling kanan
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
    
    ' 3. EKSEKUSI FORMATTING & PENYUNTIKAN KE SEL MERGER I15
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Tembak masuk ke cell merger I15 dengan aman tanpa merusak struktur merge
    With LembarForm.Range("I15").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW SUPREME: Kursor otomatis melompatke cell I19
    On Error Resume Next
    LembarForm.Select
    LembarForm.Range("I19").Select
    On Error GoTo 0
    
End Sub
