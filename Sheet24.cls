VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    'Otomatis jalan jika user mengubah Nama Bank (G) atau No Rekening (H)
    If Not Intersect(Target, Me.Range("G7:H20")) Is Nothing Then
        Call UpdateBankCOA
    End If
End Sub
