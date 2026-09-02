Attribute VB_Name = "Module89"
Sub RefreshLaporanStokjadi()
    Dim wsMaster As Worksheet, wsLaporan As Worksheet
    Dim tblProduk As ListObject, tblLaporan As ListObject
    
    ' 1. Set Sheet dan Tabel
    Set wsMaster = Sheets("PRODUK_MASTER PRODUK")
    Set tblProduk = wsMaster.ListObjects("TabelProduk")
    Set wsLaporan = Sheets("PERSEDIAAN_LAPORAN BARANG JADI")
    Set tblLaporan = wsLaporan.ListObjects("TabelStockJadi")
    
    ' Pengaman jika tabel kosong murni
    If tblLaporan.DataBodyRange Is Nothing Or tblProduk.DataBodyRange Is Nothing Then Exit Sub
    
    ' 2. KARENA STATIS: Cukup bersihkan isi data lama di Kolom 3 saja
    ' Langkah ini dijamin 100% AMAN, tidak akan menyentuh atau merusak rumus di kolom lain!
    tblLaporan.ListColumns(3).DataBodyRange.ClearContents
    
    ' 3. Ambil data master ke Range
    Dim RngMaster As Range: Set RngMaster = tblProduk.DataBodyRange
    Dim BarisMaster As Long
    Dim BarisLaporanKe As Long: BarisLaporanKe = 1
    Dim TotalBarisTabelLaporan As Long: TotalBarisTabelLaporan = tblLaporan.ListRows.Count
    
    ' 4. Sisir data master dan masukkan ke tabel laporan secara berurutan
    For BarisMaster = 1 To RngMaster.Rows.Count
        
        If LCase(Trim(RngMaster.Cells(BarisMaster, 5).Value)) = "produksi" Then
            
            ' Pengaman agar data tidak meluber keluar jika baris tabel statis   sudah penuh
            If BarisLaporanKe <= TotalBarisTabelLaporan Then
                
                ' Tembak data nama produk ke kolom 3 tabel laporan secara berurutan dari atas
                tblLaporan.ListColumns(3).DataBodyRange.Cells(BarisLaporanKe, 1).Value = RngMaster.Cells(BarisMaster, 3).Value
                
                BarisLaporanKe = BarisLaporanKe + 1
            End If
            
        End If
    Next BarisMaster
    
End Sub
