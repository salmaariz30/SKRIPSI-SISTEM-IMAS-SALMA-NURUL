VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 1. CEK RADAR UNTUK DROPDOWN BULAN (KHUSUS CELL MERGER DI E5)
    If Not Intersect(Target, Me.Range("E5")) Is Nothing Then
        ' Ambil nilai sel pertama dari cell yang di-merger biar gak eror
        Dim nilaiE5 As String
        nilaiE5 = Me.Range("E5").Cells(1, 1).Value
        
        ' Jika sel E5 dikosongkan, atau teksnya mengandung "(Telah Penyesuaian)", abaikan proses
        If nilaiE5 <> "" And InStr(nilaiE5, "(Telah Penyesuaian)") = 0 Then
            
            '--- HANYA PANGGIL MAKRO BEBAN USAHA (SESUAI REKUES  ) ---
            Call PostingPenyesuaianAmortisasiAuto   ' Macro Beban Dibayar di Muka
            '------------------------------------------------------------
            
        End If
    End If
    
    ' ====================================================================
    ' 2. REFRESH PIVOT OTOMATIS (UNTUK SEMUA PERUBAHAN DI SHEET)
    ' ====================================================================
    On Error GoTo SafeExit
    ' Cegah kiamat looping akibat refresh otomatis
    Application.EnableEvents = False
    
    ' Refresh semua Pivot Table di workbook
    ThisWorkbook.RefreshAll

SafeExit:
    ' Pastikan sensor event selalu menyala kembali dalam keadaan apa pun
    Application.EnableEvents = True
End Sub
