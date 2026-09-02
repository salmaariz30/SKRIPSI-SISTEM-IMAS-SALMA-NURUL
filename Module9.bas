Attribute VB_Name = "Module9"
Sub SimpanHargaKeTabel()
    Dim wsPricing As Worksheet
    Dim wsDaftarHarga As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim TargetRow As Long
    
    ' Matikan screen agar tidak kedip dan lebih cepat
    Application.ScreenUpdating = False
    
    ' 1. Atur Sheet
    Set wsPricing = ThisWorkbook.Sheets("PRICING STRATEGY")
    Set wsDaftarHarga = ThisWorkbook.Sheets("DAFTAR HARGA BARANG&JASA")
    
    ' 2. Atur Tabel
    On Error Resume Next
    Set tbl = wsDaftarHarga.ListObjects("TabelHargaJual")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Waduh, Tabel 'TabelHargaJual' nggak ketemu di sheet DAFTAR HARGA BARANG&JASA!", vbCritical
        Exit Sub
    End If

    ' 3. Validasi: Nama Barang tidak boleh kosong
    If wsPricing.Range("D7").Value = "" Then
        MsgBox "Nama Barang diisi dulu ya!", vbExclamation, "Data Kosong"
        Exit Sub
    End If

    ' 4. Cari baris kosong di kolom Nama Barang (kolom ke-2 di tabel)
    TargetRow = 0
    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 2).Value = "" Then
            TargetRow = i
            Exit For
        End If
    Next i

    ' Jika tidak ada baris kosong, tambah baris baru
    If TargetRow = 0 Then
        tbl.ListRows.Add
        TargetRow = tbl.ListRows.Count
    End If

    ' 5. Masukkan Data ke Tabel (Sesuaikan urutan kolom di tabelmu)
    With tbl.ListRows(TargetRow)
        .Range(1) = TargetRow                         ' Kolom 1: Nomor
        .Range(2) = wsPricing.Range("D7").Value       ' Kolom 2: Nama Barang
        .Range(3) = wsPricing.Range("D12").Value      ' Kolom 3: Harga Jual Final
        .Range(4) = wsPricing.Range("D16").Value      ' Kolom 4: Harga Sebelum Pajak/Komisi
        .Range(5) = wsPricing.Range("G34").Value      ' Kolom 5: HPP
        .Range(6) = wsPricing.Range("G35").Value      ' Kolom 6: Profit
        .Range(7) = wsPricing.Range("G37").Value      ' Kolom 7: Pajak
        .Range(8) = wsPricing.Range("G36").Value      ' Kolom 8: Komisi
    End With

    Application.ScreenUpdating = True
    
    MsgBox "Data Harga '" & wsPricing.Range("D7").Value & "' berhasil disimpan!", vbInformation, "Sukses"
End Sub
