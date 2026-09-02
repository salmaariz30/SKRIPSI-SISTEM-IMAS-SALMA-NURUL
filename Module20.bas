Attribute VB_Name = "Module20"
Sub TembakKeTextBoxKPI()

    ' ====================================================================

    ' MODUL KPI: MULTI-FILTER DETECTOR (SUPPORT DROPDOWN AKUN & BULAN "SEMUA")

    ' Sudah Disamakan Sempurna Berdasarkan Aturan Filter Baru

    ' ====================================================================

    Dim SheetHasil As Worksheet, SheetSumber As Worksheet

    Dim TabelHasil As ListObject, TabelSumber As ListObject

    Dim FilterAkun As String, FilterBulan As String

    Dim BarisTotalFisik As Long

    Dim totalDebit As Variant, totalKredit As Variant

    Dim TeksDebit As String, TeksKredit As String

    Dim UkuranFontDebit As Integer, UkuranFontKredit As Integer

    

    Set SheetHasil = ThisWorkbook.Sheets("KAS&BANK_BUKU KAS&BANK")

    Set SheetSumber = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")

    Set TabelHasil = SheetHasil.ListObjects("TabelBukuKas")

    Set TabelSumber = SheetSumber.ListObjects("TabelLaporanHarianKas")

    

    ' Ambil kedua nilai filter dropdown merger kesayangan

    FilterAkun = Trim(SheetHasil.Range("D14").Value)

    FilterBulan = Trim(SheetHasil.Range("D16").Value)

    

    ' ====================================================================

    ' LOGIKA SINKRONISASI ANGKA: DETEKSI KONDISI "SEMUA" GLOBAL vs FILTER AKTIF

    ' ====================================================================

    ' Jika KEDUA filter disetel ke "SEMUA" (artinya bener-bener gak nyaring apa-apa)

    If (UCase(FilterAkun) = "SEMUA" Or FilterAkun = "") And (UCase(FilterBulan) = "SEMUA" Or FilterBulan = "") Then

        

        ' Ambil Sum TOTAL GLOBAL langsung dari database harian asli tanpa bias

        On Error Resume Next

        totalDebit = Application.WorksheetFunction.Sum(TabelSumber.ListColumns(4).DataBodyRange)

        totalKredit = Application.WorksheetFunction.Sum(TabelSumber.ListColumns(5).DataBodyRange)

        On Error GoTo 0

        

    Else

        ' JIKA SALAH SATU ATAU KEDUA FILTER AKTIF:

        ' Langsung tembak hasil saringan fisik yang tertera di baris TOTAL TabelBukuKas  !

        If TabelHasil.ShowTotals = True Then

            BarisTotalFisik = TabelHasil.Range.Rows(TabelHasil.Range.Rows.Count).Row

            totalDebit = SheetHasil.Range("F" & BarisTotalFisik).Value

            totalKredit = SheetHasil.Range("G" & BarisTotalFisik).Value

        Else

            ' Jaga-jaga kalau tabel kosong melompong gak ada transaksi yang lolos sensor

            totalDebit = 0

            totalKredit = 0

        End If

    End If

    

    ' Validasi angka numeric anti eror tipe data

    If IsNumeric(totalDebit) Then totalDebit = CDbl(totalDebit) Else totalDebit = 0

    If IsNumeric(totalKredit) Then totalKredit = CDbl(totalKredit) Else totalKredit = 0

    

    ' 1. Ubah nominal ke bentuk format mata uang rupiah koma laptop

    TeksDebit = Format(totalDebit, "Rp #,##0")

    TeksKredit = Format(totalKredit, "Rp #,##0")

    

    ' 2. KALIBRASI FONT AUTO-SHRINK BERDASARKAN PANJANG STRING TEKS

    ' Logika Ukuran Font DEBIT

    Select Case Len(TeksDebit)

        Case Is > 15: UkuranFontDebit = 12 ' MILIARAN ke atas -> Font 12

        Case Is > 13: UkuranFontDebit = 14 ' JUTAAN s.d  SAN JUTA -> Font 14

        Case Else:    UkuranFontDebit = 16 '  SAN RIBU / NOL -> Font 16

    End Select

    

    ' Logika Ukuran Font KREDIT

    Select Case Len(TeksKredit)

        Case Is > 15: UkuranFontKredit = 12

        Case Is > 13: UkuranFontKredit = 14

        Case Else:    UkuranFontKredit = 16

    End Select

    

    ' 3. EKSEKUSI SUNTIK MASSAL KE TEXT BOX   (VISUAL SUPREME)

    With SheetHasil.Shapes("TextBoxDebit").TextFrame

        .Characters.Text = TeksDebit

        .Characters.Font.Size = UkuranFontDebit

    End With

    

    With SheetHasil.Shapes("TextBoxKredit").TextFrame

        .Characters.Text = TeksKredit

        .Characters.Font.Size = UkuranFontKredit

    End With

End Sub


