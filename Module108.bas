Attribute VB_Name = "Module108"
Sub TombolSimpanInputUtangBank_Click()

    Dim wsInput As Worksheet
    Set wsInput = Sheets("UTANG_INPUT BANK")
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data mentah ke Database Utang Biaya / Penarikan Modal
    Call SimpanTransaksiUtangBank
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum 2 Baris (Debit-Kredit) & Sekaligus Hapus Bersih Data Form
    Call JurnalUmumInputUtangBank
    
End Sub
