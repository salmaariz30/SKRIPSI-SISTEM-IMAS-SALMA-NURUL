VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Matikan deteksi perubahan sementara biar gak looping kiamat
    Application.EnableEvents = False
    
    ' Refresh semua pivot
    ThisWorkbook.RefreshAll
    
    ' Nyalakan lagi deteksinya
    Application.EnableEvents = True
End Sub
