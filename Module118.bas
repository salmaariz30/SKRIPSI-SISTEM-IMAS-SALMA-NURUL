Attribute VB_Name = "Module118"
Sub RefreshDanKunciPivot()
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim passwordSheet As String
    
    ' --- PENGATURAN PASWORD & SHEET ---
    passwordSheet = "IMAS"
    Set wsPivot = ThisWorkbook.Sheets("LAPORAN ASET TETAP NEW")
    ' -----------------------------------
    
    ' Amankan pendeteksi error & bekukan visual agar proses smooth
    On Error GoTo SafeExit
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. Buka proteksi sheet tujuan
    wsPivot.Unprotect Password:=passwordSheet
    
    ' 2. REFRESH MASSAL: Menembak langsung semua Pivot Table di sheet tersebut
    For Each pt In wsPivot.PivotTables
        pt.RefreshTable
    Next pt
    
    ' 3. Kunci kembali sheet Pivot (Tetap izinkan penggunaan Filter / Slicer)
    wsPivot.Protect Password:=passwordSheet, _
                    AllowFiltering:=True, _
                    DrawingObjects:=False, _
                    Contents:=True, _
                    Scenarios:=True
                    
    MsgBox "Laporan Pivot Berhasil Diperbarui!", vbInformation, "Sukses"

SafeExit:
    ' Pastikan sensor layar dan event menyala kembali dalam keadaan apa pun
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

