VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TogglePassword2()
    Dim tb As Object
    Set tb = ActiveSheet.TextBoxLogin
    
    If tb.PasswordChar = "*" Then
        tb.PasswordChar = "" ' Menampilkan teks
        ActiveSheet.Shapes("MataTerbuka").Visible = False
        ActiveSheet.Shapes("MataTertutup").Visible = True
    Else
        tb.PasswordChar = "*" ' Menyembunyikan teks
        ActiveSheet.Shapes("MataTerbuka").Visible = True
        ActiveSheet.Shapes("MataTertutup").Visible = False
    End If
End Sub
