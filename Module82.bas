Attribute VB_Name = "Module82"
Sub SimpanPembelianPersediaan()
    ' 1. MATIKAN OPTIMASI SISTEM DI PALING ATAS
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim LembarForm As Worksheet: Set LembarForm = Sheets("PERSEDIAAN_INPUT BARANG")
    Dim LembarKas As Worksheet: Set LembarKas = Sheets("KAS&BANK_LAPORAN HARIAN")
    
    Dim tblSource As ListObject: Set tblSource = LembarForm.ListObjects("TabelInputPersediaan")
    Dim TabelKas As ListObject: Set TabelKas = LembarKas.ListObjects("TabelLaporanHarianKas")
    Dim TabelPersediaan As ListObject
    
    On Error Resume Next
    Set TabelPersediaan = ThisWorkbook.ActiveSheet.Evaluate("TabelPembelianPersediaan")
    If TabelPersediaan Is Nothing Then
        Dim Sh As Worksheet
        For Each Sh In ThisWorkbook.Worksheets
            Set TabelPersediaan = Sh.ListObjects("TabelPembelianPersediaan")
            If Not TabelPersediaan Is Nothing Then Exit For
        Next Sh
    End If
    On Error GoTo 0
    
    If TabelPersediaan Is Nothing Or tblSource Is Nothing Or TabelKas Is Nothing Then GoTo ResetSistem
    
    ' Pastikan ada data body range-nya
    If tblSource.DataBodyRange Is Nothing Then GoTo ResetSistem
    
    ' BUKA PROTEKSI SHEET TARGET
    Dim LembarPersediaan As Worksheet: Set LembarPersediaan = TabelPersediaan.Parent
    LembarPersediaan.Unprotect Password:="IMAS"
    LembarKas.Unprotect Password:="IMAS"
    
    ' 2. AMBIL DATA DARI TABEL INPUT LANGSUNG KE ARRAY
    Dim SourceData As Variant: SourceData = tblSource.DataBodyRange.Value
    Dim TotalBarisInput As Long: TotalBarisInput = UBound(SourceData, 1)
    
    ' 3. AMBIL DATA HEADER
    Dim V_Tanggal As Date: V_Tanggal = LembarForm.Range("H12").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As String: V_NoBukti = LembarForm.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_Kategori As String: V_Kategori = Trim(LembarForm.Range("D16").MergeArea.Cells(1, 1).Value)
    Dim V_AkunKasBank As String: V_AkunKasBank = Trim(LembarForm.Range("H16").MergeArea.Cells(1, 1).Value)
    Dim V_NamaVendor As String: V_NamaVendor = Trim(LembarForm.Range("D20").MergeArea.Cells(1, 1).Value)
    Dim V_NominalGrandTotal As Double: V_NominalGrandTotal = Val(LembarForm.Range("I40").Value)
    
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_Kategori = "" Or V_AkunKasBank = "" Then GoTo ResetSistem
    
    Dim IsSaldoAwal As Boolean: IsSaldoAwal = (UCase(V_AkunKasBank) = "SALDO AWAL")
    Dim IsPembelianKredit As Boolean: IsPembelianKredit = (UCase(V_AkunKasBank) = "PEMBELIAN KREDIT")
    
    ' Tentukan teks untuk kolom STATUS baru
    Dim V_Status As String
    If IsSaldoAwal Then
        V_Status = "Saldo Awal"
    Else
        V_Status = "Pembelian"
    End If
    
    ' 4 & 5. INJEKSI DATA BARIS PER BARIS (ANTI-BARIS GHOIB)
    Dim r As Long
    Dim TotalUangKeluar As Double: TotalUangKeluar = 0
    Dim NewRowPersediaan As ListRow
    Dim IsFirstRowEmpty As Boolean
    
    ' Cek apakah tabel tujuan benar-benar kosong melompong
    IsFirstRowEmpty = (TabelPersediaan.ListRows.Count = 0 Or (TabelPersediaan.ListRows.Count = 1 And TabelPersediaan.DataBodyRange.Cells(1, 1).Value = ""))

    For r = 1 To TotalBarisInput
        ' Hanya eksekusi jika Kode Barang (Kolom 2) tidak kosong!
        If Trim(SourceData(r, 2)) <> "" Then
            
            ' Jika tabel masih kosong, pakai baris pertama yang sudah ada. Jika sudah ada isinya, buat baris baru di paling bawah.
            If IsFirstRowEmpty Then
                If TabelPersediaan.ListRows.Count = 0 Then TabelPersediaan.ListRows.Add
                Set NewRowPersediaan = TabelPersediaan.ListRows(1)
                IsFirstRowEmpty = False ' Reset flag agar baris berikutnya membuat baris baru
            Else
                Set NewRowPersediaan = TabelPersediaan.ListRows.Add(AlwaysInsert:=False)
            End If
            
            ' Masukkan data langsung ke dalam baris tabel tujuan
            With NewRowPersediaan
                .Range(1) = V_Tanggal
                .Range(2) = V_NoBukti
                .Range(3) = V_Status
                .Range(4) = V_Kategori
                .Range(5) = SourceData(r, 2)      ' Kode Barang
                .Range(6) = SourceData(r, 3)      ' Nama Barang
                .Range(7) = Val(SourceData(r, 4)) ' QTY
                .Range(8) = Val(SourceData(r, 5)) ' Harga Satuan
                .Range(9) = Val(SourceData(r, 6)) ' Total
                
                ' Set Kosmetik / Formatting font dan angka
                With .Range.Font: .Name = "Segoe UI": .Size = 10: .Color = vbBlack: End With
                .Range(7).NumberFormat = "#,##0"
                .Range(8).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
                .Range(9).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            End With
            
            TotalUangKeluar = TotalUangKeluar + Val(SourceData(r, 6))
        End If
    Next r
    
    ' Jika setelah looping ternyata tidak ada satu pun data valid yang masuk, batalkan proses kas
    If TotalUangKeluar = 0 Then GoTo ResetSistem
    
    ' 6. TARGET 3: TABEL LAPORAN KAS
    If Not IsSaldoAwal And Not IsPembelianKredit Then
        Dim NoUrutKas As Long, BarisKas As ListRow
        Dim IsKasEmpty As Boolean
        
        On Error Resume Next
        IsKasEmpty = (TabelKas.ListRows.Count = 0 Or (TabelKas.ListRows.Count = 1 And TabelKas.DataBodyRange.Cells(1, 1).Value = ""))
        If IsKasEmpty Then
            NoUrutKas = 0
            If TabelKas.ListRows.Count = 0 Then TabelKas.ListRows.Add
            Set BarisKas = TabelKas.ListRows(1)
        Else
            NoUrutKas = Application.WorksheetFunction.Max(TabelKas.ListColumns(1).DataBodyRange)
            Set BarisKas = TabelKas.ListRows.Add(AlwaysInsert:=False)
        End If
        On Error GoTo 0
        
        With BarisKas
            .Range(1) = NoUrutKas + 1
            .Range(2) = V_Tanggal
            .Range(3) = V_AkunKasBank
            .Range(4) = 0
            .Range(5) = TotalUangKeluar
            .Range(6) = "Pembelian Persediaan: " & V_Kategori
            .Range(7) = V_NoBukti
            .Range(8) = "Pembelian Persediaan"
            With .Range.Font: .Name = "Segoe UI": .Size = 10: .Color = vbBlack: End With
            .Range(4).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
            .Range(5).NumberFormat = "_(""Rp""* #,##0_);_(""Rp""* (#,##0);_(""Rp""* ""-""_);_(@_)"
        End With
    End If

ResetSistem:
    ' KUNCI KEMBALI SHEET TARGET DAN SHAPE AGAR AMAN
    If Not TabelPersediaan Is Nothing Then
        TabelPersediaan.Parent.Protect Password:="IMAS", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
        TabelPersediaan.Parent.EnableSelection = xlNoSelection
    End If
    
    LembarKas.Protect Password:="IMAS", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    LembarKas.EnableSelection = xlNoSelection

    ' KEMBALIKAN SISTEM KE SEMULA
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
