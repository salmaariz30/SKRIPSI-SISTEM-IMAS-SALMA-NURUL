Attribute VB_Name = "Module59"
Sub PrintCashOpname()
    Dim ws As Worksheet
    Set ws = Sheets("KAS&BANK_CASH OPNAME")
    
    Application.ScreenUpdating = False
    
    With ws.PageSetup
        .PrintArea = "$C$6:$I$33"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.25)
        
        .PrintQuality = 600
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .CenterHorizontally = True
    End With
    
    ' LANGSUNG TEMBAK KE PRINTER!
    ws.PrintOut
    
    Application.ScreenUpdating = True
    
    ' INI DIA MESSAGENYA
    MsgBox "Dokumen Cash Opname Berhasil Diprint!", vbInformation, "Sukses Cetak"
End Sub
