Attribute VB_Name = "Module96"
Sub TombolSimpanInputAsetTetap_Click()

    Dim wsInput As Worksheet
    Set wsInput = Sheets("INPUT DATA NEW")
    
    ' 2. EKSEKUSI MODUL 1: Simpan Data mentah ke Database Utang Biaya / Penarikan Modal
    Call SimpanKeDepresiasi
    
    ' 3. EKSEKUSI MODUL 2: Buat Jurnal Umum 2 Baris (Debit-Kredit) & Sekaligus Hapus Bersih Data Form
    Call JurnalUmumInputAsetTetap
    
End Sub


