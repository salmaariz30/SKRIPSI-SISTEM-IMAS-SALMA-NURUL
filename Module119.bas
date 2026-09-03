Attribute VB_Name = "Module119"
Sub BukaFormHapusJurnal()
    ' Mengosongkan memori form lama (jika ada sisa data tertinggal)
    Unload FormJurnalUmum
    
    ' Memunculkan UserForm Penghapusan Jurnal Umum ke layar
    FormJurnalUmum.Show
End Sub
