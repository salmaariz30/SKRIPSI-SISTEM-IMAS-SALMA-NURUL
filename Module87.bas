Attribute VB_Name = "Module87"
Sub GenerateNoBuktiDisposal()
    ' ==========================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI DISPOSAL (VERSI KORPORAT)
    ' Format Output: DSP-[DDMM]-[3 Digit Urutan] (Contoh: DSP-1306-001)
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
    
    ' 1. PENGATURAN AWAL (Sesuai Request  )
    Set LembarForm = ThisWorkbook.Sheets("PERSEDIAAN_INPUT DISPOSAL")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "DSP-" & TglBlnStr & "-" ' Hasilnya jika hari ini: "DSP-1306-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke database TabelDisposal
    On Error Resume Next
    Set TabelData = LembarForm.Evaluate("TabelDisposal")
    If TabelData Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelData = Sh.ListObjects("TabelDisposal")
            If Not TabelData Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Kolom ke-3 sesuai request  )
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' Menyisir kolom ke-3 dari TabelDisposal
            Set KolomBukti = TabelData.ListColumns(3).DataBodyRange
            
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
    
    ' Masukkan hasil akhir secara absolut ke sel F12 yang dimerge (Request  )
    With LembarForm.Range("F12").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat ke sel selanjutnya D16 (Request  )
    On Error Resume Next
    LembarForm.Activate
    LembarForm.Range("D16").Select
    On Error GoTo 0
    
End Sub
