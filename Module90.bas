Attribute VB_Name = "Module90"
Sub RefreshLaporanStokBahanStatis()
    Dim wsMaster As Worksheet, wsLaporan As Worksheet
    Dim tblProduk As ListObject, tblLaporan As ListObject
    
    ' 1. Set Sheet dan Tabel
    Set wsMaster = Sheets("PRODUK_MASTER BAHAN")
    Set tblProduk = wsMaster.ListObjects("TabelMasterBahanBaku")
    Set wsLaporan = Sheets("PERSEDIAAN_LAPORAN PER BAHAN")
    Set tblLaporan = wsLaporan.ListObjects("TabelStokBahan")
    
    ' Pengaman jika tabel kosong murni
    If tblLaporan.DataBodyRange Is Nothing Or tblProduk.DataBodyRange Is Nothing Then Exit Sub
    
    ' ?? 2. PASANG REM TANGAN (Matikan kalkulasi & layar biar wuzzz wuzzz!)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' 3. KARENA STATIS: Cukup bersihkan isi data lama di Kolom 3 saja
    tblLaporan.ListColumns(3).DataBodyRange.ClearContents
    
    ' 4. Ambil data master ke Range
    Dim RngMaster As Range: Set RngMaster = tblProduk.DataBodyRange
    Dim BarisMaster As Long
    Dim BarisLaporanKe As Long: BarisLaporanKe = 1
    Dim TotalBarisTabelLaporan As Long: TotalBarisTabelLaporan = tblLaporan.ListRows.Count
    
    ' 5. Sisir semua data dari Master Bahan Baku
    For BarisMaster = 1 To RngMaster.Rows.Count
        
        If BarisLaporanKe <= TotalBarisTabelLaporan Then
            ' Cek dulu apakah baris master ada isinya (biar gak nembak sel kosong)
            If Trim(RngMaster.Cells(BarisMaster, 3).Value) <> "" Then
                tblLaporan.ListColumns(3).DataBodyRange.Cells(BarisLaporanKe, 1).Value = RngMaster.Cells(BarisMaster, 3).Value
                BarisLaporanKe = BarisLaporanKe + 1
            End If
        End If
        
    Next BarisMaster

ResetSistem:
    ' ?? 6. LEPASKAN REM TANGAN (Kembalikan Excel ke mode normal & biarkan dia hitung barengan di akhir)
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Sub
