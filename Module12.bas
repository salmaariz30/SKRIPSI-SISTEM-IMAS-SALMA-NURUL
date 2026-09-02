Attribute VB_Name = "Module12"
Sub RefreshPivotAset()
    ' ===========================================================
    ' Nama Prosedur: RefreshPivotAset
    ' Deskripsi: Menyegarkan data pada Pivot Table tertentu
    ' Ditujukan untuk: Yang Mulia   (Skripsi Laporan Aset)
    ' ===========================================================
    
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim sheetName As String
    Dim pivotName As String
    
    ' Konfigurasi Nama Sheet dan Nama Pivot
    sheetName = "LAPORAN ASET TETAP NEW"
    pivotName = "PivotLaporanAset"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    ' Proses Refresh
    pt.RefreshTable
    
    ' Memberikan notifikasi sukses yang elegan
    MsgBox "Laporan Aset Tetap Berhasil Di-Refresh!", vbInformation, "Refresh Berhasil"
    
    On Error GoTo 0
End Sub
