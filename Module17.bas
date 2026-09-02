Attribute VB_Name = "Module17"
Sub GenerateNoBuktiSaldoAwal()
    ' ==========================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI SALDO AWAL (VERSI KORPORAT)
    ' Format Output: SAKB-[DDMM]-[3 Digit Urutan] (Contoh: SAKB-0112-001)
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Dim TabelData As ListObject
    Dim KolomBukti As Range
    Dim TglBlnStr As String
    Dim FormatPrefix As String
    Dim UrutanTerakhir As Long
    Dim Baris As Range
    Dim IsiSel As String
    Dim NomorStr As String
    Dim AngkaUrut As Long
    
    ' 1. PENGATURAN AWAL
    Set LembarForm = ThisWorkbook.Sheets("KAS&BANK_INPUT TRANSAKSI UMUM")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "SAKB-" & TglBlnStr & "-" ' Hasilnya jika hari ini 1 Des: "SAKB-0112-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database laporan harian kas
    On Error Resume Next
    Set TabelData = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN").ListObjects("TabelLaporanHarianKas")
    
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Kolom ke-7)
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            Set KolomBukti = TabelData.ListColumns(7).DataBodyRange
            
            For Each Baris In KolomBukti
                IsiSel = CStr(Baris.Value)
                ' Memeriksa apakah ada kode yang tanggal dan bulannya sama dengan hari ini
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
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Masukkan hasil akhir secara absolut ke sel H13 yang dimerge
    With LembarForm.Range("H13").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat ke sel Jumlah Dana Masuk (E17)
    LembarForm.Range("E17").Select
    
End Sub
