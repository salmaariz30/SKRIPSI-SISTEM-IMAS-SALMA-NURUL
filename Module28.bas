Attribute VB_Name = "Module28"
Sub TembakKeTextBoxKPILaporanKas()

    ' ====================================================================

    ' MODUL KPI GLOBAL: SUNTIK DATA TOTAL HARIAN KE TEXT BOX LAPORAN HARIAN

    ' Menembak ke TextBoxDebitLK & TextBoxKreditLK di Sheet Laporan Harian

    ' Kalibrasi Font Otomatis & Bebas Cut-Off Sesuai Standar Premium

    ' ====================================================================

    Dim SheetLaporan As Worksheet

    Dim TabelLaporan As ListObject

    Dim BarisTotalFisik As Long

    Dim totalDebit As Variant, totalKredit As Variant

    Dim TeksDebit As String, TeksKredit As String

    Dim UkuranFontDebit As Integer, UkuranFontKredit As Integer

    

    Set SheetLaporan = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")

    Set TabelLaporan = SheetLaporan.ListObjects("TabelLaporanHarianKas")

    

    ' ====================================================================

    ' AMBIL ANGKA TOTAL DARI DATABASE GLOBAL (ANTI BIAS)

    ' ====================================================================

    If TabelLaporan.ShowTotals = True Then

        ' Jika baris total aktif, ambil langsung dari fisik paling bawah biar sinkron visual

        BarisTotalFisik = TabelLaporan.Range.Rows(TabelLaporan.Range.Rows.Count).Row

        totalDebit = SheetLaporan.Range("F" & BarisTotalFisik).Value ' Kolom F (Debit)

        totalKredit = SheetLaporan.Range("G" & BarisTotalFisik).Value ' Kolom G (Kredit)

    Else

        ' Jika baris total dinonaktifkan, hitung manual via WorksheetFunction Sum kilat

        On Error Resume Next

        totalDebit = Application.WorksheetFunction.Sum(TabelLaporan.ListColumns(4).DataBodyRange)

        totalKredit = Application.WorksheetFunction.Sum(TabelLaporan.ListColumns(5).DataBodyRange)

        On Error GoTo 0

    End If

    

    ' Validasi angka numeric anti eror tipe data (mengamankan sel kosong hantu)

    If IsNumeric(totalDebit) Then totalDebit = CDbl(totalDebit) Else totalDebit = 0

    If IsNumeric(totalKredit) Then totalKredit = CDbl(totalKredit) Else totalKredit = 0

    

    ' 1. Ubah nominal ke bentuk format mata uang rupiah koma laptop

    TeksDebit = Format(totalDebit, "Rp #,##0")

    TeksKredit = Format(totalKredit, "Rp #,##0")

    

    ' 2. KALIBRASI FONT AUTO-SHRINK BERDASARKAN PANJANG STRING TEKS

    ' Logika Ukuran Font DEBIT GLOBAL

    Select Case Len(TeksDebit)

        Case Is > 15: UkuranFontDebit = 12 ' MILIARAN ke atas -> Font 12

        Case Is > 13: UkuranFontDebit = 14 ' JUTAAN s.d  SAN JUTA -> Font 14

        Case Else:    UkuranFontDebit = 16 '  SAN RIBU / NOL -> Font 16

    End Select

    

    ' Logika Ukuran Font KREDIT GLOBAL

    Select Case Len(TeksKredit)

        Case Is > 15: UkuranFontKredit = 12

        Case Is > 13: UkuranFontKredit = 14

        Case Else:    UkuranFontKredit = 16

    End Select

    

    ' 3. EKSEKUSI SUNTIK MASSAL KE TEXT BOX BARU   DI SHEET LAPORAN HARIAN

    With SheetLaporan.Shapes("TextBoxDebitLK").TextFrame

        .Characters.Text = TeksDebit

        .Characters.Font.Size = UkuranFontDebit

    End With

    

    With SheetLaporan.Shapes("TextBoxKreditLK").TextFrame

        .Characters.Text = TeksKredit

        .Characters.Font.Size = UkuranFontKredit

    End With

End Sub
