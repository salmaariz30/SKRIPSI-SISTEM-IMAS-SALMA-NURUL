Attribute VB_Name = "Module75"
Sub GenerateNoBuktiPelunasanUtang()
    
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
    
    ' 1. PENGATURAN AWAL & TARGET FORM
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT PELUNASAN")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari tanggal hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "BPUT-" & TglBlnStr & "-" ' Hasilnya misal hari ini: "BPUT-1106-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database pelunasan utang usaha global
    On Error Resume Next
    ' Sesuaikan nama sheet penampung tabel jika berbeda di laptop
    Set TabelData = ThisWorkbook.Sheets("UTANG_DAFTAR PELUNASAN UTANG").ListObjects("TabelPelunasanUtang")
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Kolom ke-3 sesuai titah  )
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' Menyasar mutlak kolom ke-3 dari tabel pelunasan utang
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
    End If
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT BARU
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Masukkan hasil akhir secara absolut ke sel G15 yang dimerge (Anti-Error Merged Cells)
    With LembarForm.Range("G15").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat ketat ke sel input selanjutnya (D19)
    Application.GoTo LembarForm.Range("D19")
    
End Sub

