VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TogglePassword()
    Dim tb As Object
    Set tb = ActiveSheet.TextBox1
    
    If tb.PasswordChar = "*" Then
        tb.PasswordChar = "" ' Menampilkan teks
        ActiveSheet.Shapes("MataTerbuka1").Visible = False
        ActiveSheet.Shapes("MataTertutup1").Visible = True
    Else
        tb.PasswordChar = "*" ' Menyembunyikan teks
        ActiveSheet.Shapes("MataTerbuka1").Visible = True
        ActiveSheet.Shapes("MataTertutup1").Visible = False
    End If
End Sub

Sub TogglePassword1()
    Dim tb As Object
    Set tb = ActiveSheet.TextBox2
    
    If tb.PasswordChar = "*" Then
        tb.PasswordChar = "" ' Menampilkan teks
        ActiveSheet.Shapes("MataTerbuka2").Visible = False
        ActiveSheet.Shapes("MataTertutup2").Visible = True
    Else
        tb.PasswordChar = "*" ' Menyembunyikan teks
        ActiveSheet.Shapes("MataTerbuka2").Visible = True
        ActiveSheet.Shapes("MataTertutup2").Visible = False
    End If
End Sub
