Attribute VB_Name = "Module83"
Sub JurnalUmumInputPembelianPersediaan()
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim tblInputBarang As ListObject
    Dim wsInput As Worksheet: Set wsInput = Sheets("PERSEDIAAN_INPUT BARANG")
    
    Set tblInputBarang = wsInput.ListObjects("TabelInputPersediaan")
    
    Dim V_Tanggal As Variant: V_Tanggal = wsInput.Range("H12").MergeArea.Cells(1, 1).Value
    Dim V_NoBukti As Variant: V_NoBukti = wsInput.Range("D12").MergeArea.Cells(1, 1).Value
    Dim V_Kategori As String: V_Kategori = Trim(wsInput.Range("D16").MergeArea.Cells(1, 1).Value)
    Dim V_MetodeBayar As String: V_MetodeBayar = Trim(wsInput.Range("H16").MergeArea.Cells(1, 1).Value)
    Dim V_DeskripsiCustom As String
    Dim IsSaldoAwal As Boolean
    
    ' --- NOMINAL DIAMBIL DARI TOTAL KOLOM TOTAL TABEL INPUT PERSEDIAAN ---
    Dim V_Nominal As Double
    Dim colIdx As Long
    colIdx = tblInputBarang.ListColumns("Total").Index
    V_Nominal = Application.WorksheetFunction.Sum(tblInputBarang.ListColumns(colIdx).DataBodyRange)
    
    ' 1. Validasi Pengaman Awal
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_Kategori = "" Or V_MetodeBayar = "" Or V_Nominal <= 0 Then
        MsgBox "Mohon pastikan Tanggal, No Bukti, Kategori, Metode, dan Nominal terisi valid.", vbExclamation, "Data Belum Lengkap"
        Exit Sub
    End If
    
    ' 2. Filter Metode Bayar
    If UCase(V_MetodeBayar) = "PEMBELIAN KREDIT" Then Exit Sub
    IsSaldoAwal = (UCase(V_MetodeBayar) = "SALDO AWAL")
    
    ' 3. Penentuan Nama Akun & Kelompok (SISTEM PERPETUAL MURNI)
    Dim namaAkunDebit As String, namaAkunKredit As String
    Dim kelompokDebit As String, kelompokKredit As String
    
    ' Kolom DEBIT untuk Perpetual selalu masuk ke Aset Lancar (Persediaan)
    kelompokDebit = "Aset Lancar"
    Select Case V_Kategori
        Case "Bahan Baku":    namaAkunDebit = "Persediaan Bahan Baku"
        Case "Barang Dagang": namaAkunDebit = "Persediaan Barang Dagang"
        Case "Barang Jadi":   namaAkunDebit = "Persediaan Barang Jadi"
        Case Else:            namaAkunDebit = "Persediaan " & V_Kategori
    End Select
    
    ' Penentuan KREDIT berdasarkan Jenis Transaksi
    If IsSaldoAwal Then
        namaAkunKredit = "Ekuitas - Saldo Awal"
        kelompokKredit = "Ekuitas"
        V_DeskripsiCustom = "Pencatatan Saldo Awal Persediaan " & V_Kategori
    Else
        namaAkunKredit = V_MetodeBayar
        kelompokKredit = "Kas & Bank"
        V_DeskripsiCustom = "Pembelian " & V_Kategori & " (No. Bukti: " & V_NoBukti & ")"
    End If
    
    ' ====================================================================
    ' PENGAMAN UTAMA SAKTI (BIAR WUSH SECEPAT KILAT)
    ' ====================================================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 4. Ambil Kode Akun dari COA
    Dim barisDebit As Variant, barisKredit As Variant
    Dim kodeAkunDebit As Variant, kodeAkunKredit As Variant
    
    On Error Resume Next
    barisDebit = Application.Match(namaAkunDebit, tblCOA.ListColumns(2).DataBodyRange, 0)
    barisKredit = Application.Match(namaAkunKredit, tblCOA.ListColumns(2).DataBodyRange, 0)
    kodeAkunDebit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisDebit)
    kodeAkunKredit = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisKredit)
    On Error GoTo 0
    
    If IsError(kodeAkunDebit) Or IsError(kodeAkunKredit) Then
        GoTo NyalakanSistem
        MsgBox "Nama akun '" & namaAkunDebit & "' atau '" & namaAkunKredit & "' tidak ditemukan di COA!", vbCritical, "Error Master COA"
        Exit Sub
    End If
    
    ' 5. TEKNIK KILAT: Ambil Posisi Baris Terbawah untuk 2 Baris Sekaligus
    Dim LastRowIdx As Long
    Dim IsTableEmpty As Boolean
    
    On Error Resume Next
    IsTableEmpty = (tblJurnal.ListRows.Count = 0 Or (tblJurnal.ListRows.Count = 1 And tblJurnal.DataBodyRange.Cells(1, 1).Value = ""))
    On Error GoTo 0
    
    If IsTableEmpty Then
        If tblJurnal.ListRows.Count = 0 Then tblJurnal.ListRows.Add
        LastRowIdx = 1
    Else
        ' Tambah 2 baris instan di akhir tabel tanpa menggeser sheet
        tblJurnal.ListRows.Add AlwaysInsert:=False
        tblJurnal.ListRows.Add AlwaysInsert:=False
        LastRowIdx = tblJurnal.ListRows.Count - 1
    End If
    
  ' --- INJEKSI BARIS DEBIT ---
    With tblJurnal.ListRows(LastRowIdx)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiCustom
        .Range(4) = kodeAkunDebit
        .Range(5) = namaAkunDebit
        .Range(6) = V_Nominal
        .Range(7) = 0
        .Range(8) = kelompokDebit
        .Range(9) = "Tidak"
    End With
    
    ' --- INJEKSI BARIS KREDIT ---
    With tblJurnal.ListRows(LastRowIdx + 1)
        .Range(1) = V_Tanggal
        .Range(2) = V_NoBukti
        .Range(3) = V_DeskripsiCustom
        .Range(4) = kodeAkunKredit
        .Range(5) = namaAkunKredit
        .Range(6) = 0
        .Range(7) = V_Nominal
        .Range(8) = kelompokKredit
        .Range(9) = "Tidak"
    End With
    
    ' 6. SAPU BERSIH MASSA MERGED CELLS (PESANAN  )
    On Error Resume Next
    wsInput.Range("D12").MergeArea.ClearContents
    wsInput.Range("D16").MergeArea.ClearContents
    wsInput.Range("D20").MergeArea.ClearContents
    wsInput.Range("H12").MergeArea.ClearContents
    wsInput.Range("H16").MergeArea.ClearContents
    
   ' Kosongkan kolom spesifik di TabelInputPersediaan tanpa merusak struktur tabel
    If tblInputBarang.ListRows.Count > 0 Then
        On Error Resume Next
        tblInputBarang.ListColumns("No.").DataBodyRange.ClearContents
        tblInputBarang.ListColumns("Nama Barang").DataBodyRange.ClearContents
        tblInputBarang.ListColumns("Qty").DataBodyRange.ClearContents
        tblInputBarang.ListColumns("Harga Satuan").DataBodyRange.ClearContents
        On Error GoTo 0
    End If
    On Error GoTo 0
    
    wsInput.Activate
    wsInput.Range("E25").Select
    
' Label pemulihan sistem Excel jika terjadi eror di tengah jalan
NyalakanSistem:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
