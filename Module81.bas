Attribute VB_Name = "Module81"
Sub TombolSimpanRetur_Click()
    Dim wsInput As Worksheet
    Set wsInput = Sheets("PENDAPATAN_INPUT RETUR")
    
    ' --------------------------------------------------------------------
    ' 2. ORKESTRASI MODUL BERUNTUN (ANTI-LONCAT)
    ' --------------------------------------------------------------------
    
    ' LANGKAH A: Simpan rekap transaksi penghasilan lain ke database
    Call SimpanTransaksiRetur
    
    ' LANGKAH B: Eksekusi penjurnalan otomatis ke Jurnal Umum
    ' (Catatan  : Sub di bawah ini otomatis menyapu bersih isi form setelah sukses)
    Call JurnalUmumInputReturPenjualan

End Sub



