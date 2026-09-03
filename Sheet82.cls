VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet82"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim TabelLaporan As ListObject
    
    On Error Resume Next
    Set TabelLaporan = Me.ListObjects("TabelLaporanHarianKas")
    On Error GoTo 0
    
    ' ====================================================================
    ' ?? RADAR 1: JIKA   MENGUBAH PILIHAN DROPDOWN DI SEL E10 MERGER
    ' ====================================================================
    If Not Intersect(Target, Me.Range("E10")) Is Nothing Then
        ' Langsung jalankan makro penyaringan otomatis Jenis Aktivitas!
        Call SaringJenisAktivitasKas
        Exit Sub ' Keluar dari sub biar gak bentrok dengan proses di bawah
    End If
    
    ' ====================================================================
    ' ?? RADAR 2: JIKA DATA DI DALAM TABEL BERUBAH (INPUT/HAPUS TRANSAKSI)
    ' ====================================================================
    If Not TabelLaporan Is Nothing Then
        If Not Intersect(Target, TabelLaporan.Range) Is Nothing Then
            
            ' ? DOUBLE IMPACT: Panggil kedua kode KPI biar semuanya update serempak!
            Call TembakKeTextBoxKPILaporanKas  ' <- Ini kode lama   yang sakti
            
        End If
    End If
End Sub
