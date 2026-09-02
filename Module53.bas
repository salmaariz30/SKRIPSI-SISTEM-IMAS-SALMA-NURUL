Attribute VB_Name = "Module53"
Sub TombolSimpanInputTarik_Click()
 
    Dim wsInput As Worksheet
    Set wsInput = Sheets("PENGELUARAN USAHA_INPUT TARIK")
    
    ' 1. SISTEM PENGAMAN (Validasi Awal Sebelum Eksekusi)
    ' Memastikan koordinat cell krusial milik   tidak kosong melompong
    If wsInput.Range("D11").Value = "" Or _
       wsInput.Range("G11").Value = "" Or _
       wsInput.Range("D15").Value = "" Or _
       wsInput.Range("G15").Value = "" Or _
       wsInput.Range("D19").Value = "" Or _
       wsInput.Range("G19").Value = "" Or _
       wsInput.Range("D23").Value = "" Or _
       wsInput.Range("G23").Value = "" Or _
       wsInput.Range("D27").Value = "" Then
        
        MsgBox "Data form penarikan belum lengkap!" & vbCrLf & _
               "Mohon pastikan Jenis Penarikan, No. Bukti, Tanggal, Akun Kas/Bank, Nominal, Nama Penarik, No Rekening, % Kepemilikan, dan Deskripsi sudah terisi semua ya.", _
               vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data mentah ke Database Utang Biaya / Penarikan Modal
    Call SimpanTransaksiPenarikanModal
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum 2 Baris (Debit-Kredit) & Sekaligus Hapus Bersih Data Form
    Call JurnalUmumInputTarikModal
    
End Sub

