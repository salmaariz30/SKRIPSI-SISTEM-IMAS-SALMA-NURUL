Attribute VB_Name = "Module19"
Sub TarikDataBukuKasDinamis()
    ' ====================================================================
    ' MODULE MASTER - BALIK KE SETELAN AWAL (KOLOM C SAMPAI H)
    ' ====================================================================
    Dim SheetSumber As Worksheet, SheetHasil As Worksheet
    Dim TabelSumber As ListObject, TabelHasil As ListObject
    Dim AkunDicari As String, BulanDicari As String
    Dim BarisSumber As Long, BarisHasil As Long
    Dim TotalBarisSumber As Long
    Dim SaldoBerjalan As Double
    Dim TglTrans As Variant
    Dim CocokAkun As Boolean, CocokBulan As Boolean
    Dim NilaiDebit As Double, NilaiKredit As Double
    Dim V_Debit As Variant, V_Kredit As Variant
    Dim BulanTeksIndo As String
    
    Set SheetSumber = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")
    Set SheetHasil = ThisWorkbook.Sheets("KAS&BANK_BUKU KAS&BANK")
    Set TabelSumber = SheetSumber.ListObjects("TabelLaporanHarianKas")
    Set TabelHasil = SheetHasil.ListObjects("TabelBukuKas")
    
    ' Kunci visual biar jalannya secepat kilat di laptop ASUS
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. SINKRONISASI AWAL: Matikan Total & Reset tabel hasil jadi 1 baris kosong murni
    On Error Resume Next
    TabelHasil.ShowTotals = False
    Do While TabelHasil.ListRows.Count > 1
        TabelHasil.ListRows(2).Delete
    Loop
    TabelHasil.DataBodyRange.Rows(1).ClearContents
    On Error GoTo 0
    
    ' 2. AMBIL VALUE FILTER DROPDOWN DROPDOWN MERGER
    AkunDicari = Trim(SheetHasil.Range("D14").Value)
    BulanDicari = Trim(SheetHasil.Range("D16").Value) ' Mengambil teks bulan/SEMUA dari D16
    
    TotalBarisSumber = TabelSumber.ListRows.Count
    BarisHasil = 19
    SaldoBerjalan = 0
  
    If TotalBarisSumber = 0 Then
        SheetHasil.Range("G14").Value = 0
        GoTo AkhirSaja
    End If

    ' 3. LOOPING DATABASE HARIAN (Tanggal Sumber Tetap di Kolom 2)
    For BarisSumber = 1 To TotalBarisSumber
        
        ' Kunci Pengaman: Jika kolom tanggal di harian kosong, skip!
        If TabelSumber.DataBodyRange(BarisSumber, 2).Value = "" Then GoTo LanjutLoop
        
        ' --- LOGIKA FILTER AKUN ---
        If AkunDicari = "" Or UCase(AkunDicari) = "SEMUA" Then
            CocokAkun = True
        Else
            CocokAkun = (TabelSumber.DataBodyRange(BarisSumber, 3).Value = AkunDicari)
        End If
        
        ' --- LOGIKA FILTER BULAN (Dropdown Teks vs Tanggal Lengkap) ---
        If CocokAkun = True Then
            TglTrans = TabelSumber.DataBodyRange(BarisSumber, 2).Value
            
            If BulanDicari = "" Or UCase(BulanDicari) = "SEMUA" Then
                CocokBulan = True
            Else
                ' Ubah tanggal lengkap harian jadi teks bulan Indonesia murni ([$-421])
                BulanTeksIndo = Application.Text(TglTrans, "[$-421]mmmm")
                CocokBulan = (LCase(Trim(BulanTeksIndo)) = LCase(Trim(BulanDicari)))
            End If
            
            ' --- EKSEKUSI SUNTIK DATA KE TABEL BUKU KAS VINTAGE ---
            If CocokBulan = True Then
                V_Debit = TabelSumber.DataBodyRange(BarisSumber, 4).Value
                V_Kredit = TabelSumber.DataBodyRange(BarisSumber, 5).Value
                
                If IsNumeric(V_Debit) Then NilaiDebit = CDbl(V_Debit) Else NilaiDebit = 0
                If IsNumeric(V_Kredit) Then NilaiKredit = CDbl(V_Kredit) Else NilaiKredit = 0
                
                SaldoBerjalan = SaldoBerjalan + NilaiDebit - NilaiKredit
                
                ' Jika baris pertama (19) sudah terisi, baru buat baris baru di bawahnya
                If BarisHasil > 19 Then
                    TabelHasil.ListRows.Add
                End If
                
                ' KEMBALI KE SUSUNAN AWAL (Kolom C sampai H)
                With SheetHasil
                    .Range("C" & BarisHasil).Value = TglTrans                          ' Kolom 1: Tanggal
                    .Range("D" & BarisHasil).Value = TabelSumber.DataBodyRange(BarisSumber, 7).Value ' Kolom 2: No Bukti
                    .Range("E" & BarisHasil).Value = TabelSumber.DataBodyRange(BarisSumber, 6).Value ' Kolom 3: Deskripsi
                    .Range("F" & BarisHasil).Value = NilaiDebit                        ' Kolom 4: Debit
                    .Range("G" & BarisHasil).Value = NilaiKredit                       ' Kolom 5: Kredit
                    .Range("H" & BarisHasil).Value = SaldoBerjalan                     ' Kolom 6: Saldo Akhir Berjalan
                End With
                
                BarisHasil = BarisHasil + 1
            End If
        End If

LanjutLoop:
    Next BarisSumber
    
    ' 4. AKTIFKAN BARIS TOTAL JIKA DATA BERHASIL DIISI (Susunan Rumus Awal)
    On Error Resume Next
    If BarisHasil > 19 Then
        TabelHasil.ShowTotals = True
        With TabelHasil
            .ListColumns(1).TotalsRowLabel = "Total"
            .ListColumns(4).TotalsCalculation = xlTotalsCalculationSum  ' Kolom 4 (Debit)
            .ListColumns(5).TotalsCalculation = xlTotalsCalculationSum  ' Kolom 5 (Kredit)
            .ListColumns(6).TotalsCalculation = xlTotalsCalculationNone ' Kolom 6 (Saldo Berjalan)
        End With
    End If
    On Error GoTo 0
    
AkhirSaja:
    SheetHasil.Range("G14").Value = SaldoBerjalan
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
