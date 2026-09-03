Attribute VB_Name = "Module92"
Sub GenerateNoBuktiPemakaianBahan()
    ' ==========================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI PEMAKAIAN BAHAN (VERSI TURBO)
    ' Format Output: PMK-[DDMM]-[3 Digit Urutan] (Contoh: PMK-1306-001)
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
    
    ' 1. PENGATURAN AWAL (Form Input)
    Set LembarForm = ThisWorkbook.Sheets("PERSEDIAAN_INPUT BAHAN")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "PMK-" & TglBlnStr & "-" ' Hasilnya jika hari ini: "PMK-1306-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke database TabelPemakaianBahan
    On Error Resume Next
    Set TabelData = LembarForm.Evaluate("TabelPemakaianBahan")
    If TabelData Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelData = Sh.ListObjects("TabelPemakaianBahan")
            If Not TabelData Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Kolom ke-2 di sheet PERSEDIAAN_TABEL OPNAME)
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' Menyisir KOLOM KEDUA dari TabelPemakaianBahan sesuai request
            Set KolomBukti = TabelData.ListColumns(2).DataBodyRange
            
            If Not KolomBukti Is Nothing Then
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
    End If
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Masukkan hasil akhir secara absolut ke sel H12 yang dimerge (Sesuai Request  )
    With LembarForm.Range("H12").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat ke sel selanjutnya D16 (Sesuai Request  )
    On Error Resume Next
    LembarForm.Activate
    LembarForm.Range("D16").Select
    On Error GoTo 0
    
End Sub
