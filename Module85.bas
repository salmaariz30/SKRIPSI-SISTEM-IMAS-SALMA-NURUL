Attribute VB_Name = "Module85"
Sub RefreshLaporanStokDagang()
    Dim wsMaster As Worksheet, wsLaporan As Worksheet
    Dim tblProduk As ListObject, tblLaporan As ListObject
    
    ' 1. Set Sheet dan Tabel
    Set wsMaster = Sheets("PRODUK_MASTER PRODUK")
    Set tblProduk = wsMaster.ListObjects("TabelProduk")
    Set wsLaporan = Sheets("PERSEDIAAN_LAPORAN PERSEDIAAN")
    Set tblLaporan = wsLaporan.ListObjects("TabelStokDagang")
    
    ' 2. Bersihkan dulu data lama di kolom ke-3 Tabel Laporan
    If Not tblLaporan.DataBodyRange Is Nothing Then
        tblLaporan.ListColumns(3).DataBodyRange.ClearContents
    End If
    
    ' 3. Ambil data master ke Range
    Dim RngMaster As Range
    Set RngMaster = tblProduk.DataBodyRange
    
    Dim BarisMaster As Long
    Dim BarisLaporanKe As Long
    BarisLaporanKe = 1
    
    ' 4. Sisir data baris demi baris secara nyata
    For BarisMaster = 1 To RngMaster.Rows.Count
        ' Kolom 5 = Kategori Produk ("Dagang")
        If LCase(Trim(RngMaster.Cells(BarisMaster, 5).Value)) = "dagang" Then
            
            ' Jika baris di tabel laporan kurang, tambah otomatis
            If tblLaporan.ListRows.Count < BarisLaporanKe Then
                tblLaporan.ListRows.Add AlwaysInsert:=False
            End If
            
            ' !!! PERBAIKAN: Mengambil Kolom 3 dari Master, ditembak ke Kolom 3 Laporan !!!
            tblLaporan.ListColumns(3).DataBodyRange.Cells(BarisLaporanKe, 1).Value = RngMaster.Cells(BarisMaster, 3).Value
            
            ' Lanjut ke baris laporan berikutnya
            BarisLaporanKe = BarisLaporanKe + 1
        End If
    Next BarisMaster
    
End Sub
