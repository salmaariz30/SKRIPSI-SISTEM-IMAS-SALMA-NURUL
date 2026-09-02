Attribute VB_Name = "Module21"
Sub IsiTanggalHariIni_PindahSaldo()
    ' ==========================================================
    ' MODUL AUTO-FILL TANGGAL TRANSAKSI - VERSI PINDAH SALDO
    ' Langsung eksekusi tanpa peduli kursor sedang di mana
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("KAS&BANK_INPUT PINDAH SALDO")
    
    ' 1. EKSEKUSI: Langsung bersihkan area merged D11 dan isi tanggal hari ini
    With LembarForm.Range("D11").MergeArea
        .ClearContents
        .Value = Date
    End With
    
    ' 2. UX FLOW: Pindahkan kursor secara otomatis setelah isi tanggal
    LembarForm.Range("G11").Select
    
End Sub
