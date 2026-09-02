VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet71"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' GANTI "B5" dengan posisi sel tempat DROPDOWN   berada!
    If Not Intersect(Target, Range("D11")) Is Nothing Then
        Call BuatDropdownUtang
    End If
End Sub
