Attribute VB_Name = "Module40"
Sub GenerateNoBuktiPengeluaranUsaha()
    ' ====================================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI PENGELUARAN USAHA (PREMIUM - FIXED)
    ' Format Output: EXP-[DDMM]-[3 Digit Urutan] (Contoh: EXP-0306-001)
    ' Target Tembak: Sheet "PENGELUARAN USAHA_INPUT DATA" -> Sel G12 (Merged)
    ' Target Database: Sheet "PENGELUARAN USAHA_DAFTAR" -> TabelBebanUsaha (Kolom 2)
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
    
    ' 1. PENGATURAN AWAL FORM INPUT
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT DATA")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "EXP-" & TglBlnStr & "-" ' Hasilnya contoh hari ini: "EXP-0306-"
    UrutanTerakhir = 0
    
    ' 2. MENGHUBUNGKAN LANGSUNG KE DATABASE PILIHAN
    On Error Resume Next
    Set TabelData = ThisWorkbook.Sheets("PENGELUARAN USAHA_DAFTAR").ListObjects("TabelBebanUsaha")
    
    ' Fitur Radar Keliling: Jaga-jaga kalau tabelnya digeser   ke sheet lain
    If TabelData Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelData = Sh.ListObjects("TabelBebanUsaha")
            If Not TabelData Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' 3. LOGIKA PENYISIRAN DATA DI DATABASE (Kolom Kedua / Ke-2)
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' RADAR DIKUNCI MATI KE KOLOM KEDUA TABEL BEBAN USAHA
            Set KolomBukti = TabelData.ListColumns(2).DataBodyRange
            
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
    
    ' 4. EKSEKUSI PENAMBAHAN NOMOR URUT MAJU
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' Kunci visual layar biar perpindahan kursor smooth tanpa kedip di laptop ASUS
    Application.ScreenUpdating = False
    LembarForm.Activate
    
    ' Masukkan hasil nomor bukti baru secara absolut ke sel G12 yang dimerge
    With LembarForm.Range("G12").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 5. UX FLOW: Kursor otomatis lompat mendarat tepat di sel input selanjutnya D16 (Merged)
    LembarForm.Range("D16").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
End Sub
