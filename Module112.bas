Attribute VB_Name = "Module112"
Sub TombolSimpanSaldoAwalPiutang_Click()
    Dim wsInput As Worksheet
    Dim tblJurnal As ListObject
    Dim SuksesSimpan As Boolean
    
    Set wsInput = Sheets("PIUTANG_INPUT SALDO AWAL")
    
    ' 1. JALANKAN MODUL PERTAMA DAN CEK STATUSNYA
    ' Pastikan di dalam Sub SimpanSaldoAwalPiutang sudah diubah/ditangkap status suksesnya
    SuksesSimpan = SimpanSaldoAwalPiutang()
    
    ' PENGAMAN: Jika modul pertama gagal (False), langsung keluar dari program
    If Not SuksesSimpan Then
        GoTo NyalakanSistem
    End If
    
    ' 2. JALANKAN KODE KEDUA HANYA JIKA MODUL PERTAMA BERHASIL
    Call JurnalUmumInputSaldoAwalPiutang

NyalakanSistem:
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
