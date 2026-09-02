Attribute VB_Name = "Module31"
Sub InputPenghasilanLain_NamaPelanggan()
    ' ==========================================================
    ' MODUL AUTO-WRITE KATA "Umum" - VERSI PENGHASILAN LAIN
    ' ==========================================================
    
    Dim LembarForm As Worksheet
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_PENGHASILAN LAIN")
    
    ' Supaya pergerakan layar dan kursor berjalan super smooth anti kedip
    Application.ScreenUpdating = False
    
    ' 1. EKSEKUSI: Aktifkan sheet agar proses Select kursor berjalan mulus
    LembarForm.Activate
    
    ' 2. SUNTIK TEKS: Bersihkan area merged H19 lalu ketik "Umum"
    With LembarForm.Range("H19").MergeArea
        .ClearContents
        .Value = "Umum"
    End With
    
    ' 3. UX FLOW: Lemparkan kursor secara otomatis ke D19 setelah ngetik
    ' Menggunakan .Cells(1, 1) agar Excel mendarat tepat di hulu merged cell D19
    LembarForm.Range("D19").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End Sub
