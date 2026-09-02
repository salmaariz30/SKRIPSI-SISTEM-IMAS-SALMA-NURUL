Attribute VB_Name = "Module32"
Sub GenerateNoBuktiPenghasilanLain()
    ' ==========================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI PENGHASILAN LAIN (PREMIUM)
    ' Format Output: PHL-[DDMM]-[3 Digit Urutan] (Contoh: PHL-3005-001)
    ' Menembak ke Cell H11 dan Lompat Kursor ke D15 Merged Cell
    ' Ref Database: TabelPenghasilanLain Kolom Ke-3
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
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_PENGHASILAN LAIN")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "PHL-" & TglBlnStr & "-" ' Hasilnya contoh: "PHL-3005-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database penghasilan lain
    On Error Resume Next
    ' Catatan  : Pastikan tabel ini berada di sheet database pendapatan   ya!
    Set TabelData = ThisWorkbook.Sheets("PENDAPATAN_PENGHASILAN LAIN").ListObjects("TabelPenghasilanLain")
    
    ' Jaga-jaga jika ternyata tabelnya ditaruh   di sheet khusus database/laporan lain
    If TabelData Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelData = Sh.ListObjects("TabelPenghasilanLain")
            If Not TabelData Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Sesuai Perintah: Kolom Ke-3)
    If Not TabelData Is Nothing Then
        If TabelData.ListRows.Count > 0 Then
            ' Mengunci target radar hanya pada Kolom Ke-3 tabel referensi
            Set KolomBukti = TabelData.ListColumns(3).DataBodyRange
            
            For Each Baris In KolomBukti
                IsiSel = Trim(CStr(Baris.Value))
                
                ' Memeriksa apakah ada kode yang prefix-nya sama dengan hari ini
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
    
    ' 3. EKSEKUSI PENAMBAHAN NOMOR URUT MAJU
    UrutanTerakhir = UrutanTerakhir + 1
    
    ' 4. SUNTIK DATA KE FORMULIR MASAL
    Application.ScreenUpdating = False
    LembarForm.Activate
    
    ' Masukkan hasil akhir secara absolut ke sel H11 yang dimerge
    With LembarForm.Range("H11").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 5. UX FLOW: Kursor otomatis lompat mendarat tepat di sel tujuan (D15 Merged)
    LembarForm.Range("D15").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
    ' Sound Effect Sukses di Hati
End Sub
