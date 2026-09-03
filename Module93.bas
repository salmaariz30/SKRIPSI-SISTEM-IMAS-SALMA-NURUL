Attribute VB_Name = "Module93"
Sub SimpanPemakaianBahanBaku()
    ' ==========================================================
    ' MODUL SAVING & BULK INSERT DATA PEMAKAIAN BAHAN (VERSI TURBO)
    ' Menggunakan Teknik Array untuk Mencegah Excel Freeze/Lag
    ' ==========================================================
    
    ' 1. MATIKAN OPTIMASI SISTEM DI PALING ATAS (Anti-Lemot)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim LembarForm As Worksheet: Set LembarForm = Sheets("PERSEDIAAN_INPUT BAHAN")
    Dim tblSource As ListObject: Set tblSource = LembarForm.ListObjects("InputPemakaianBahan")
    Dim TabelPemakaianBahan As ListObject
    
    ' Menghubungkan ke database TabelPemakaianBahan di sheet PERSEDIAAN_TABEL OPNAME secara dinamis
    On Error Resume Next
    Set TabelPemakaianBahan = ThisWorkbook.ActiveSheet.Evaluate("TabelPemakaianBahan")
    If TabelPemakaianBahan Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelPemakaianBahan = Sh.ListObjects("TabelPemakaianBahan")
            If Not TabelPemakaianBahan Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' Pengaman jika objek tabel tidak ditemukan di workbook
    If TabelPemakaianBahan Is Nothing Or tblSource Is Nothing Then GoTo ResetSistem
    If tblSource.DataBodyRange Is Nothing Then GoTo ResetSistem
    
    ' BUKA PROTEKSI SHEET TARGET
    Dim LembarTarget As Worksheet: Set LembarTarget = TabelPemakaianBahan.Parent
    LembarTarget.Unprotect Password:="IMAS"
    
    ' 2. AMBIL DATA DARI TABEL INPUT LANGSUNG KE ARRAY
    Dim SourceData As Variant: SourceData = tblSource.DataBodyRange.Value
    Dim TotalBarisInput As Long: TotalBarisInput = UBound(SourceData, 1)
    
    ' Menyisir data yang BENAR-BENAR ISI (Kolom 2: Kode Barang di tabel input tidak boleh kosong)
    Dim DataCount As Long: DataCount = 0
    Dim i As Long
    For i = 1 To TotalBarisInput
        If Trim(SourceData(i, 2)) <> "" Then
            DataCount = DataCount + 1
        End If
    Next i
    
    ' Jika inputan ternyata kosong semua, batalkan proses
    If DataCount = 0 Then GoTo ResetSistem
    
    ' 3. AMBIL DATA HEADER FORM PEMAKAIAN (Sesuai Request Koordinat Sel  )
    Dim V_Tanggal As Variant: V_Tanggal = LembarForm.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_NoReff As Variant: V_NoReff = LembarForm.Range("H12").MergeArea.Cells(1, 1).Value
    Dim V_Staff As String: V_Staff = Trim(LembarForm.Range("D16").MergeArea.Cells(1, 1).Value)
    Dim V_Keterangan As String: V_Keterangan = Trim(LembarForm.Range("H16").MergeArea.Cells(1, 1).Value)
    
    ' Validasi wajib isi untuk header dasar
    If V_Tanggal = 0 Or V_NoReff = "" Or V_Staff = "" Or V_Keterangan = "" Then
        MsgBox "Mohon lengkapi data Tanggal, No. Reff, Nama Staff, dan Keterangan terlebih dahulu!", vbExclamation, "Data Header Kosong"
        GoTo ResetSistem
    End If
    
    ' 4. MENYIAPKAN AREA BARIS BARU DI TABEL TUJUAN (PAS 100% - ANTI BARIS GAIB)
    Dim TargetRange As Range, EmptyCheck As Boolean
    Dim k As Long, IndeksBarisAwal As Long
    
    On Error Resume Next
    EmptyCheck = (TabelPemakaianBahan.ListRows.Count = 0 Or (TabelPemakaianBahan.ListRows.Count = 1 And TabelPemakaianBahan.DataBodyRange.Cells(1, 1).Value = ""))
    On Error GoTo 0
    
    If EmptyCheck Then
        If TabelPemakaianBahan.ListRows.Count = 0 Then TabelPemakaianBahan.ListRows.Add
        Set TargetRange = TabelPemakaianBahan.ListRows(1).Range
        ' Jika data input lebih dari 1, tambahkan sisa barisnya
        If DataCount > 1 Then
            For k = 2 To DataCount
                TabelPemakaianBahan.ListRows.Add
            Next k
        End If
    Else
        ' Catat posisi baris terakhir SEBELUM ditambah data baru
        IndeksBarisAwal = TabelPemakaianBahan.ListRows.Count
        
        ' Tambahkan baris baru secara PAS sebanyak DataCount menggunakan Looping murni
        For k = 1 To DataCount
            TabelPemakaianBahan.ListRows.Add
        Next k
        
        ' Kunci koordinat TargetRange tepat di baris-baris baru yang barusan dibuat
        Set TargetRange = TabelPemakaianBahan.ListRows(IndeksBarisAwal + 1).Range
    End If
    
    ' 5. INJEKSI DATA VALID SAJA KE DATABASE (Sesuai Mapping  )
    Dim r As Long, BarisTargetKe As Long: BarisTargetKe = 1
    
    For r = 1 To TotalBarisInput
        ' Hanya proses baris yang memiliki Kode Barang (Kolom 2 pada tabel input)
        If Trim(SourceData(r, 2)) <> "" Then
            With TargetRange.Rows(BarisTargetKe)
                .Cells(1) = V_Tanggal          ' Kolom 1: Tanggal (D12)
                .Cells(2) = V_NoReff           ' Kolom 2: No. Reff (H12)
                .Cells(3) = V_Staff            ' Kolom 3: Staff (D16)
                .Cells(4) = V_Keterangan       ' Kolom 4: Keterangan (H16)
                .Cells(5) = SourceData(r, 2)   ' Kolom 5: Kode Barang (Kolom 2 Input)
                .Cells(6) = SourceData(r, 3)   ' Kolom 6: Nama Barang (Kolom 3 Input)
                .Cells(7) = SourceData(r, 4)   ' Kolom 7: Satuan (Kolom 4 Input)
                .Cells(8) = Val(SourceData(r, 5)) ' Kolom 8: Kuantitas Terpakai (Kolom 5 Input)
                .Cells(9) = SourceData(r, 6)   ' Kolom 9: Alokasi (Kolom 6 Input)
                
                ' Penataan Format Tampilan dan Pemaksaan ANTI-BOLD Gaib di Sini!
                With .Font
                    .Name = "Segoe UI"
                    .Size = 10
                    .ColorIndex = 1
                    .Bold = False
                End With
                .Cells(8).NumberFormat = "#,##0" ' Format angka untuk Kuantitas Terpakai
            End With
            BarisTargetKe = BarisTargetKe + 1
        End If
    Next r
    
    ' 6. SAPU BERSIH AKHIR (Double Protection biar bener-bener steril ga ada yang ketularan bold)
    If Not TabelPemakaianBahan.DataBodyRange Is Nothing Then
        With TabelPemakaianBahan.DataBodyRange.Font
            .Bold = False
            .Name = "Segoe UI"
            .Size = 10
            .ColorIndex = 1
        End With
    End If
    
    ' TUTUP PROTEKSI SHEET TARGET
    LembarTarget.Protect Password:="IMAS", AllowFiltering:=True
    
    ' 7. PEMBERSIHAN CELL FORM INPUT & KOLOM SPESIFIK TABEL
    With LembarForm
        .Range("D12").MergeArea.ClearContents
        .Range("H12").MergeArea.ClearContents
        .Range("D16").MergeArea.ClearContents
        .Range("H16").MergeArea.ClearContents
    End With
    
    With tblSource
        .ListColumns(3).DataBodyRange.ClearContents ' Kolom Ke-3 (Nama Barang)
        .ListColumns(5).DataBodyRange.ClearContents ' Kolom Ke-5 (Kuantitas Terpakai)
        .ListColumns(6).DataBodyRange.ClearContents ' Kolom Ke-6 (Alokasi)
    End With
    
    ' Kembalikan kursor ke posisi awal input tanggal
    LembarForm.Activate
    LembarForm.Range("D12").Select
    
    MsgBox "Data Pemakaian Bahan berhasil disimpan ke database sebanyak " & DataCount & " item!", vbInformation, "Penyimpanan Sukses"

ResetSistem:
    ' KEMBALIKAN SISTEM KE SEMULA
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
