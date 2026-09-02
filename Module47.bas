Attribute VB_Name = "Module47"
Sub GenerateNoBuktiInputProduksi()
    ' ==========================================================
    ' MODUL AUTO-GENERATE NOMOR BUKTI INPUT PRODUKSI (PREMIUM)
    ' Format Output: PRD-[DDMM]-[3 Digit Urutan] (Contoh: PRD-0406-001)
    ' Menembak ke Cell G11 dan Lompat Kursor ke D15 Merged Cell
    ' Ref Database: TabelBiayaProduksi Kolom Ke-3 (Sheet: PENGELUARAN USAHA_PRODUKSI)
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
    
    ' 1. PENGATURAN FORM INPUT
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN_INPUT PRODUKSI")
    
    ' Mengambil 2 digit tanggal dan 2 digit bulan secara otomatis dari hari ini
    TglBlnStr = Format(Date, "ddmm")
    FormatPrefix = "PRD-" & TglBlnStr & "-" ' Hasilnya contoh harian: "PRD-0406-"
    UrutanTerakhir = 0
    
    ' Menghubungkan ke tabel database biaya produksi sesuai titah
    On Error Resume Next
    Set TabelData = ThisWorkbook.Sheets("PENGELUARAN USAHA_PRODUKSI").ListObjects("TabelBiayaProduksi")
    
    ' Jaga-jaga jika tabelnya sempat digeser   ke sheet database lain
    If TabelData Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelData = Sh.ListObjects("TabelBiayaProduksi")
            If Not TabelData Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' 2. LOGIKA PENYISIRAN DATA DI DATABASE (Kolom Ke-3)
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
    
    ' 4. SUNTIK DATA KE FORMULIR INPUT PRODUKSI
    Application.ScreenUpdating = False
    LembarForm.Activate
    
    ' Masukkan hasil akhir secara absolut ke sel G11 yang dimerge
    With LembarForm.Range("G11").MergeArea
        .ClearContents
        .Value = FormatPrefix & Format(UrutanTerakhir, "000")
    End With
    
    ' 5. UX FLOW: Kursor otomatis lompat mendarat tepat di sel tujuan (D15 Merged)
    LembarForm.Range("D15").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
    ' Sound Effect Sukses di Hati   Semakin Membahana!
End Sub
