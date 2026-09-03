VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ws As Worksheet
    
    ' ??? PENGAMAN UTAMA: Matikan screen dan kalkulasi agar Excel tidak interupsi PivotTable
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error Resume Next ' Jika ada error kecil, tetap lanjutkan baris kode berikutnya
    
    ' 1. Pastikan sheet Login tampil terlebih dahulu
    Sheets("LOGIN").Visible = xlSheetVisible
    
    ' 2. Sembunyikan semua sheet lain secara total
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "LOGIN" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    On Error GoTo 0
    
    ' ??? KEMBALIKAN KE SEMULA: Nyalakan kembali fitur otomatis Excel
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim ctrl As CommandBarControl
    
    ' =============================================================
    ' 1. SIKAT HABIS: Looping untuk hapus menu lama agar tidak numpuk
    ' =============================================================
    On Error Resume Next
    ' Bersihkan di menu klik kanan sel biasa ("Cell")
    For Each ctrl In Application.CommandBars("Cell").Controls
        If InStr(ctrl.Caption, "Hapus Baris Tabel") > 0 Then
            ctrl.Delete
        End If
    Next ctrl
    
    ' Bersihkan di menu klik kanan khusus tabel ("List Range Popup")
    For Each ctrl In Application.CommandBars("List Range Popup").Controls
        If InStr(ctrl.Caption, "Hapus Baris Tabel") > 0 Then
            ctrl.Delete
        End If
    Next ctrl
    On Error GoTo 0
    
    ' =============================================================
    ' 2. FILTER: Hanya muncul jika diklik di dalam wilayah Tabel (ListObject)
    ' =============================================================
    If Not Target.ListObject Is Nothing Then
        
        ' Tembak ke menu klik kanan sel biasa
        On Error Resume Next
        With Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = "Hapus Baris Tabel Ini"
            .OnAction = "HapusBarisTabelAktif"
        End With
        
        ' Tembak juga ke menu klik kanan khusus Excel Table
        With Application.CommandBars("List Range Popup").Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = "Hapus Baris Tabel Ini"
            .OnAction = "HapusBarisTabelAktif"
        End With
        On Error GoTo 0
        
    End If
End Sub
