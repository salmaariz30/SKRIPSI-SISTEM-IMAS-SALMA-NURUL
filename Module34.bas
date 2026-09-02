Attribute VB_Name = "Module34"
Sub GenerateNoBuktiReturPenjualan()
    ' ====================================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI RETUR & POTONGAN PENJUALAN (PREMIUM)
    ' Format Output: RTR-[DDMM]-[3 Digit Urutan] (Contoh: RTR-3105-001)
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
    
    ' 1. PENGATURAN AWAL & ALAMAT BARU
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT RETUR")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "RTR-" & TglBlnStr & "-" ' Hasilnya contoh hari ini: "RTR-3105-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database utama Retur milik
    On Error Resume Next
    Set TabelData = ThisWorkbook.Sheets("PENDAPATAN_DATA RETUR").ListObjects("TabelRetur")
    
    ' Fitur Radar Keliling: Jika tabel tidak di sheet itu, cari di seluruh penjuru workbook
    If TabelData Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelData = Sh.ListObjects("TabelRetur")
            If Not TabelData Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Sesuai Titah: Kolom Kelima / Ke-5)
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' ?? RADAR DIKUNCI MATI KE KOLOM KE-5 TABEL RETUR
            Set KolomBukti = TabelData.ListColumns(5).DataBodyRange
            
            If Not KolomBukti Is Nothing Then
                For Each Baris In KolomBukti
                    IsiSel = Trim(CStr(Baris.Value))
                    ' Memeriksa apakah ada kode yang prefiksnya sama dengan hari ini
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
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT MAJU
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Kunci visual layar biar perpindahan kursor smooth tanpa kedip
    Application.ScreenUpdating = False
    LembarForm.Activate
    
    ' Masukkan hasil akhir secara absolut ke sel E9 yang dimerge
    With LembarForm.Range("E9").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 4. UX FLOW: Kursor otomatis lompat mendarat tepat di sel tujuan berikutnya E11 (Merged)
    LembarForm.Range("E11").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub
