Attribute VB_Name = "Module58"
Sub TombolSimpanInputBeban_Click()
    ' ====================================================================
    ' EKSEKUSI ALL-IN-ONE - VERSI INPUT BEBAN USAHA
    ' Terkunci Khusus untuk Sheet: PENGELUARAN USAHA_INPUT DATA
    ' PROSES: Panggil Modul Simpan Terlebih Dahulu ?? Baru Modul Jurnal Umum
    ' ====================================================================
    Dim wsInput As Worksheet
    Set wsInput = Sheets("PENGELUARAN USAHA_INPUT DATA")
    
    ' 1. SISTEM PENGAMAN (Validasi Awal Sebelum Eksekusi)
    ' Memastikan koordinat cell formulir pengeluaran milik   tidak kosong melompong
    If wsInput.Range("D12").Value = "" Or _
       wsInput.Range("G12").Value = "" Or _
       wsInput.Range("D16").Value = "" Or _
       wsInput.Range("G16").Value = "" Or _
       wsInput.Range("D20").Value = "" Or _
       wsInput.Range("G20").Value = "" Or _
       wsInput.Range("D24").Value = "" Or _
       wsInput.Range("D28").Value = "" Then
        
        MsgBox "Data formulir pengeluaran belum lengkap,  !" & vbCrLf & _
               "Mohon pastikan Kategori Beban, No. Bukti, Tanggal Bayar, Dibayar Kepada, " & vbCrLf & _
               "Nominal, Bayar Dari Akun, Status Penggunaan, dan Deskripsi sudah terisi.", _
               vbExclamation, "Validasi Gagal"
        Exit Sub
    End If
    
    ' 2. EKSEKUSI MODUL 1: Simpan data ke Database Pengeluaran Usaha (Versi Sempurna  )
    Call SimpanTransaksiPengeluaranUsaha
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum & Sekaligus Pembersihan Form Input
    Call JurnalUmumInputBeban
    
End Sub

