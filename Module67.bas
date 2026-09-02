Attribute VB_Name = "Module67"
Sub GenerateNoBuktiPelunasanPiutang()
    ' ====================================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI PELUNASAN PIUTANG (VERSI KORPORAT)
    ' Format Output: BKM-[DDMM]-[3 Digit Urutan] (Contoh: BKM-1006-001)
    ' Target Tembak: Sheet "PIUTANG_INPUT DATA PIUTANG" -> G17 (Merged)
    ' Target Lompat: Sel E21
    ' Target Sisir: Kolom 3 dari "TabelPelunasanPiutang"
    ' ====================================================================
    
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
    
    ' 1. PENGATURAN AWAL TARGET FORM RATING
    Set LembarForm = ThisWorkbook.Sheets("PIUTANG_INPUT DATA PIUTANG")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "BKM-" & TglBlnStr & "-" ' Hasilnya jika hari ini 10 Juni: "BKM-1006-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database pelunasan piutang sesuai mandat
    On Error Resume Next
    Set TabelData = ThisWorkbook.Sheets("PIUTANG_PELUNASAN PIUTANG").ListObjects("TabelPelunasanPiutang")
    On Error GoTo 0
    
    ' Validasi Keberadaan Tabel Target
    If TabelData Is Nothing Then
        MsgBox "Error: Tabel 'TabelPelunasanPiutang' tidak ditemukan di sheet PIUTANG_PELUNASAN PIUTANG!", _
               vbCritical, "Sistem Gagal"
        Exit Sub
    End If
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Sesuai Titah: Kolom Ke-3)
    If TabelData.ListRows.Count > 0 Then
        Set KolomBukti = TabelData.ListColumns(2).DataBodyRange
        
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
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Masukkan hasil akhir secara absolut ke sel G17 yang dimerge
    Application.ScreenUpdating = False
    LembarForm.Activate
    
    With LembarForm.Range("G17").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat manis ke sel tujuan   (E21)
    LembarForm.Range("E21").MergeArea.Cells(1, 1).Select
    Application.ScreenUpdating = True
    
End Sub

