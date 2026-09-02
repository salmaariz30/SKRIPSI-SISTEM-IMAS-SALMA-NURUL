VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' ====================================================================
    
    ' Jika yang berubah bukan sel D14 atau D16, langsung keluar
    If Intersect(Target, Range("D14,D16")) Is Nothing Then Exit Sub
    
    ' Pengaman jika hapus sel massal agar tidak crash
    If Target.Cells.Count > 1 Then Exit Sub
    
    On Error GoTo JalurPenyelamat
    
    ' Matikan sensor sementara (Rem otomatis biar ga loop kiamat)
    Application.EnableEvents = False
    
    ' 1. Jalankan kodingan Module lama (Tarik data)
    Call TarikDataBukuKasDinamis
    
    ' 2. Jeda napas 1 detik agar Excel selesai merapikan tabel fisik
    Application.Wait (Now + TimeValue("0:00:01"))
    
    ' 3. Jalankan kodingan Module KPI Text Box
    Call TembakKeTextBoxKPI

JalurPenyelamat:
    ' Pastikan sensor wajib menyala kembali apa pun yang terjadi!
    Application.EnableEvents = True
End Sub
