Attribute VB_Name = "Module37"
Sub TombolSimpanInputKas_Click()
    ' ==========================================================
    ' EKSEKUSI ALL-IN-ONE
    ' ==========================================================
    Dim wsInput As Worksheet
    Set wsInput = Sheets("KAS&BANK_INPUT TRANSAKSI UMUM")
    
    ' 1. SISTEM PENGAMAN (Validasi Awal Sebelum Eksekusi)
    ' Cek apakah cell krusial masih kosong
    If wsInput.Range("E13").Value = "" Or _
       wsInput.Range("H13").Value = "" Or _
       wsInput.Range("E17").Value = "" Or _
       wsInput.Range("H17").Value = "" Or _
       wsInput.Range("E21").Value = "" Or _
       wsInput.Range("H25").Value = "" Then
        
        MsgBox "Data form belum lengkap!" & vbCrLf & _
               "Mohon pastikan Tanggal, No. Bukti, Nominal, Sumber Dana, Keterangan, dan Akun Kas sudah terisi semua ya.", _
               vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data ke Laporan Harian
    Call SimpanTransaksiKasMasuk
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum & Sekaligus Hapus Data Form
    Call JurnalUmumInputKas
    
End Sub

