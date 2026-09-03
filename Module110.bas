Attribute VB_Name = "Module110"
Sub TombolSimpanPelunasanUtang_Click()

    Dim wsInput As Worksheet
    Set wsInput = Sheets("UTANG_INPUT PELUNASAN")
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data mentah ke Database Utang Biaya / Penarikan Modal
    Call SimpanTransaksiPelunasanUtang
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum 2 Baris (Debit-Kredit) & Sekaligus Hapus Bersih Data Form
    Call JurnalUmumInputPelunasanUtang
    
End Sub

