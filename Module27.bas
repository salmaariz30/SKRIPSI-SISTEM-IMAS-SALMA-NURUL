Attribute VB_Name = "Module27"
Sub InputPenjualan_NamaPelanggan()
    ' ==========================================================
    ' MODUL AUTO-WRITE KATA "Umum" - VERSI INPUT PENJUALAN
    ' Langsung mengetik kata "Umum" di C15 dan lempar kursor ke F15
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_INPUT PENJUALAN")
    
    ' 1. EKSEKUSI: Aktifkan sheet agar proses Select kursor berjalan mulus
    LembarForm.Activate
    
    ' 2. SUNTIK TEKS: Bersihkan area merged C15 lalu ketik "Umum"
    With LembarForm.Range("C15").MergeArea
        .ClearContents
        .Value = "Umum"
    End With
    
    ' 3. UX FLOW: Lemparkan kursor secara otomatis ke F15 setelah ngetik
    LembarForm.Range("F15").Select
    
End Sub
