VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet63"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    ' ====================================================================
    ' ?? EVENT SINGLE PIVOT (CINEMATIC SMOOTH EDITION)
    ' Target Table: "PivotAnalisisPiutang"
    ' Status: Ultra Smooth Transition, Anti-Kedip, No-Lag System
    ' Katakunci Proteksi: IMAS
    ' ====================================================================
    Const PWD As String = "IMAS"
    Dim namaPivot As String
    
    namaPivot = "PivotAnalisisPiutang"
    
    ' ??? KUNCI SAKRAL: Bekukan semua gerakan layar di baris pertama absolut!
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual ' Tidurkan rumus lain agar Excel fokus
    End With
    
    On Error GoTo SafeExit
    
    ' 1. Jebol gembok sheet ini sendiri
    Me.Unprotect Password:=PWD
    
    ' 2. TEMBAK JANTUNG MEMORI: Hancurkan cache lama & paksa refresh data baru
    With Me.PivotTables(namaPivot)
        .PivotCache.Refresh
        .RefreshTable
    End With
    
    ' 3. Gembok kembali dengan parameter resmi Excel (Filter & Pivot tetap bisa diklik Ratu)
    Me.Protect Password:=PWD, _
               AllowFiltering:=True, _
               AllowUsingPivotTables:=True

SafeExit:
    ' ?? PEMULIHAN SISTEM: Bangunkan kembali Excel setelah proses bawah tanah selesai
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

