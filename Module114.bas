Attribute VB_Name = "Module114"
Sub TombolSimpanPelunasanPiutang_Click()

    Dim wsInput As Worksheet
    Set wsInput = Sheets("PIUTANG_INPUT DATA PIUTANG")
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data mentah ke Database Utang Biaya / Penarikan Modal
    Call SimpanTransaksiPelunasanPiutang
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum 2 Baris (Debit-Kredit) & Sekaligus Hapus Bersih Data Form
    Call JurnalUmumInputPelunasanPiutang
    
End Sub
