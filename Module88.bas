Attribute VB_Name = "Module88"
Sub SimpanDisposalPersediaan()
    ' ==========================================================
    ' MODUL SAVING & BULK INSERT DATA DISPOSAL (VERSI KORPORAT)
    ' Menggunakan Teknik Array untuk Mencegah Excel Freeze/Lag
    ' ==========================================================
    
    ' 1. MATIKAN OPTIMASI SISTEM DI PALING ATAS (Anti-Lemot)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim LembarForm As Worksheet: Set LembarForm = Sheets("PERSEDIAAN_INPUT DISPOSAL")
    Dim tblSource As ListObject: Set tblSource = LembarForm.ListObjects("TabelInputDisposal")
    Dim TabelDisposal As ListObject
    
    ' Menghubungkan ke database TabelDisposal lintas sheet secara dinamis
    On Error Resume Next
    Set TabelDisposal = ThisWorkbook.ActiveSheet.Evaluate("TabelDisposal")
    If TabelDisposal Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelDisposal = Sh.ListObjects("TabelDisposal")
            If Not TabelDisposal Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    ' Pengaman jika objek tabel tidak ditemukan di workbook
    If TabelDisposal Is Nothing Or tblSource Is Nothing Then GoTo ResetSistem
    If tblSource.DataBodyRange Is Nothing Then GoTo ResetSistem
    
    ' 2. AMBIL DATA DARI TABEL INPUT LANGSUNG KE ARRAY
    Dim SourceData As Variant: SourceData = tblSource.DataBodyRange.Value
    Dim TotalBarisInput As Long: TotalBarisInput = UBound(SourceData, 1)
    
    ' Menyisir data yang BENAR-BENAR ISI (Kolom 2: Kode Barang tidak boleh kosong)
    Dim DataCount As Long: DataCount = 0
    Dim i As Long
    For i = 1 To TotalBarisInput
        If Trim(SourceData(i, 2)) <> "" Then
            DataCount = DataCount + 1
        End If
    Next i
    
    ' Jika inputan ternyata kosong semua, batalkan proses
    If DataCount = 0 Then GoTo ResetSistem
    
    ' 3. AMBIL DATA HEADER FORM DISPOSAL (Sesuai Request Koordinat Sel  )
    Dim V_Tanggal As Variant: V_Tanggal = LembarForm.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_NoDisposal As Variant: V_NoDisposal = LembarForm.Range("F12").MergeArea.Cells(1, 1).Value
    Dim V_Kategori As String: V_Kategori = Trim(LembarForm.Range("D16").MergeArea.Cells(1, 1).Value)
    Dim V_Staff As String: V_Staff = Trim(LembarForm.Range("F16").MergeArea.Cells(1, 1).Value)
    
    ' Validasi wajib isi untuk header dasar
    If V_Tanggal = 0 Or V_NoDisposal = "" Or V_Kategori = "" Or V_Staff = "" Then
        MsgBox "Mohon lengkapi data Tanggal, No. Disposal, Kategori, dan Nama Staff terlebih dahulu!", vbExclamation, "Data Header Kosong"
        GoTo ResetSistem
    End If
    
    ' 4. MENYIAPKAN AREA BARIS BARU DI TABEL TUJUAN (PAS 100% - ANTI BARIS GAIB)
    Dim TargetRange As Range, EmptyCheck As Boolean
    Dim k As Long, IndeksBarisAwal As Long
    
    On Error Resume Next
    EmptyCheck = (TabelDisposal.ListRows.Count = 0 Or (TabelDisposal.ListRows.Count = 1 And TabelDisposal.DataBodyRange.Cells(1, 1).Value = ""))
    On Error GoTo 0
    
    If EmptyCheck Then
        If TabelDisposal.ListRows.Count = 0 Then TabelDisposal.ListRows.Add
        Set TargetRange = TabelDisposal.ListRows(1).Range
        ' Jika data input lebih dari 1, tambahkan sisa barisnya
        If DataCount > 1 Then
            For k = 2 To DataCount
                TabelDisposal.ListRows.Add
            Next k
        End If
    Else
        ' Catat posisi baris terakhir SEBELUM ditambah data baru
        IndeksBarisAwal = TabelDisposal.ListRows.Count
        
        ' Tambahkan baris baru secara PAS sebanyak DataCount menggunakan Looping murni
        For k = 1 To DataCount
            TabelDisposal.ListRows.Add
        Next k
        
        ' Kunci koordinat TargetRange tepat di baris-baris baru yang barusan dibuat (tidak melar, tidak gaib!)
        Set TargetRange = TabelDisposal.ListRows(IndeksBarisAwal + 1).Range
    End If
    
    ' 5. INJEKSI DATA VALID SAJA KE DATABASE (ANTI-BARIS KOSONG)
    Dim r As Long, BarisTargetKe As Long: BarisTargetKe = 1
    
    For r = 1 To TotalBarisInput
        ' Hanya proses baris yang memiliki Kode Barang (Kolom 2)
        If Trim(SourceData(r, 2)) <> "" Then
            With TargetRange.Rows(BarisTargetKe)
                .Cells(1) = V_Tanggal        ' Kolom 1: Tanggal (D12)
                .Cells(2) = V_NoDisposal     ' Kolom 2: No. Disposal (F12)
                .Cells(3) = V_Staff          ' Kolom 3: Staff (F16)
                .Cells(4) = V_Kategori       ' Kolom 4: Kategori (D16)
                .Cells(5) = SourceData(r, 2)  ' Kolom 5: Kode Barang (Kolom 2 Input)
                .Cells(6) = SourceData(r, 1)  ' Kolom 6: Nama Barang (Kolom 1 Input)
                .Cells(7) = Val(SourceData(r, 3)) ' Kolom 7: QTY (Kolom 3 Input)
                .Cells(8) = SourceData(r, 4)  ' Kolom 8: Alasan (Kolom 4 Input)
                
                ' Penataan Format Tampilan dan Pemaksaan ANTI-BOLD Gaib di Sini!
                With .Font
                    .Name = "Segoe UI"
                    .Size = 10
                    .Color = vbBlack
                    .Bold = False
                End With
                .Cells(7).NumberFormat = "#,##0" ' Format angka untuk Qty
            End With
            BarisTargetKe = BarisTargetKe + 1
        End If
    Next r
    
    ' 6. SAPU BERSIH AKHIR (Double Protection biar bener-bener steril ga ada yang ketularan bold)
    If Not TabelDisposal.DataBodyRange Is Nothing Then
        With TabelDisposal.DataBodyRange.Font
            .Bold = False
            .Name = "Segoe UI"
            .Size = 10
        End With
    End If
    

ResetSistem:
    ' KEMBALIKAN SISTEM KE SEMULA
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
