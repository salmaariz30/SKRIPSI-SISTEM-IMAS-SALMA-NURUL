Attribute VB_Name = "Module103"
Sub TombolSimpanDisposalPersediaan_Click()
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data mentah ke Database Utang Biaya / Penarikan Modal
    Call SimpanDisposalPersediaan
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum 2 Baris (Debit-Kredit) & Sekaligus Hapus Bersih Data Form
    Call JurnalUmumInputDisposalPersediaanDinamis
    
End Sub

