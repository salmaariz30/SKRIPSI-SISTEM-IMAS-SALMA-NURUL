Attribute VB_Name = "Module8"
Sub SimpanResepKeTabel()
    Dim wsKalkulator As Worksheet
    Dim wsDaftar As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim TargetRow As Long
    Const PWD As String = "IMAS"
    
    Application.ScreenUpdating = False
    
    Set wsKalkulator = ThisWorkbook.Sheets("KALKULATOR USAHA")
    Set wsDaftar = ThisWorkbook.Sheets("DAFTAR RESEP")
    Set tbl = wsDaftar.ListObjects("TabelDaftarResep")

    ' BUKA PROTEKSI SHEET DAFTAR RESEP
    wsDaftar.Unprotect Password:=PWD

    ' --- STRATEGI BARU: CARI BARIS YANG NAMA BARANGNYA KOSONG ---
    TargetRow = 0
    For i = 1 To tbl.ListRows.Count
        ' Cek kolom ke-2 (Nama Barang). Jika kosong, kita pakai baris ini.
        If tbl.DataBodyRange(i, 2).Value = "" Then
            TargetRow = i
            Exit For
        End If
    Next i

    ' Jika tidak ada baris kosong di tengah/atas, baru tambah baris di paling bawah
    If TargetRow = 0 Then
        tbl.ListRows.Add
        TargetRow = tbl.ListRows.Count
    End If

    ' --- PENGISIAN DATA ---
    With tbl.ListRows(TargetRow)
        .Range(1) = wsKalkulator.Range("B6").MergeArea.Cells(1, 1).Value  ' Nama Barang (Aman Merged Cells)
        .Range(2) = wsKalkulator.Range("F26")                             ' Biaya Bahan
        .Range(3) = wsKalkulator.Range("F27")                             ' Biaya Kemasan
        .Range(4) = wsKalkulator.Range("F28")                             ' Biaya Tenaga
        .Range(5) = wsKalkulator.Range("F29")                             ' Biaya Operasional
        .Range(6) = wsKalkulator.Range("F30")                             ' Biaya Pengiriman
        .Range(7) = wsKalkulator.Range("F31")                             ' HPP
    End With

    ' KUNCI KEMBALI SHEET DAFTAR RESEP
    wsDaftar.Protect Password:=PWD, AllowFiltering:=True

    Application.ScreenUpdating = True
    MsgBox "Data '" & wsKalkulator.Range("B6").MergeArea.Cells(1, 1).Value & "' berhasil disimpan!", vbInformation
End Sub
