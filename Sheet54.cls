VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsPivot As Worksheet
    Dim passwordSheet As String
    
    ' --- PENGATURAN PASSWORD ---
    ' Sesuaikan password di bawah ini dengan password sheet Pivot
    passwordSheet = "IMAS"
    
    ' ====================================================================
    ' 1. CEK RADAR UNTUK DROPDOWN BULAN (SEL D5) - MILIK
    ' ====================================================================
    If Not Intersect(Target, Me.Range("D5")) Is Nothing Then
        ' Jika sel D5 sengaja dikosongkan atau di-delete, abaikan proses jurnal
        If Me.Range("D5").Value <> "" Then
            ' Jalankan macro posting jurnal penyesuaian utama
            Call PostingPenyesuaianBulananAuto
        End If
    End If
    
End Sub
