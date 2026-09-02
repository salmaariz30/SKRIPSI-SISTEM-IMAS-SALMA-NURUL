Attribute VB_Name = "Module72"
Sub GenerateNoBuktiUtangUsaha()
    ' ====================================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI UTANG USAHA (VERSI KORPORAT PT MMD)
    ' Format Output: UTG-[DDMM]-[3 Digit Urutan] (Contoh: UTG-1106-001)
    ' TARGET FORM: "UTANG_INPUT USAHA" -> Cell H12 (Merge Area)
    ' TARGET DB SISIR: "TabelUtangUsaha" -> Kolom Ke-3 (No. Faktur/Invoice)
    ' FLOW UX FINAL: Kursor otomatis lompat ke Sel Input D16
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    Dim LembarDatabase As Worksheet
    Dim TabelData As ListObject
    Dim KolomBukti As Range
    Dim TglBlnStr As String
    Dim FormatPrefix As String
    Dim UrutanTerakhir As Long
    Dim Baris As Range
    Dim IsiSel As String
    Dim NomorStr As String
    Dim AngkaUrut As Long
    
    ' 1. PENGATURAN AWAL TARGET FORM
    Set LembarForm = ThisWorkbook.Sheets("UTANG_INPUT USAHA")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "UTG-" & TglBlnStr & "-" ' Hasilnya jika hari ini 11 Juni: "UTG-1106-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke sheet database dan tabel utama Utang Usaha
    On Error Resume Next
    Set LembarDatabase = ThisWorkbook.Sheets("UTANG_BUKU UTANG") ' Sesuaikan nama sheet database utang   jika berbeda
    Set TabelData = LembarDatabase.ListObjects("TabelUtangUsaha")
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' Mengunci murni target penyisiran pada Kolom Ketiga (Kolom 3)
            Set KolomBukti = TabelData.ListColumns(2).DataBodyRange
            
            ' Validasi darurat jika kolom ke-3 belum terisi data sama sekali
            If Not KolomBukti Is Nothing Then
                For Each Baris In KolomBukti
                    IsiSel = CStr(Baris.Value)
                    
                    ' Memeriksa apakah ada kode prefix yang tanggal dan bulannya sama dengan hari ini
                    If Left(IsiSel, Len(FormatPrefix)) = FormatPrefix Then
                        ' Mengambil 3 digit angka terakhir di paling kanan nomor bukti
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
    
    ' Masukkan hasil akhir secara absolut ke sel H12 yang dimerge
    With LembarForm.Range("H12").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat ketat ke sel input selanjutnya (D16)
    Application.GoTo LembarForm.Range("D16")
    
End Sub

