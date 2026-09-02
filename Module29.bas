Attribute VB_Name = "Module29"
Sub SaringJenisAktivitasKas()

    ' ====================================================================

    ' MODUL UTAMA: AUTO-FILTER TABEL BERDASARKAN DROPDOWN JENIS AKTIVITAS

    ' Menyaring Kolom Ke-8 (Jenis Aktivitas) Sesuai Pilihan   di E10

    ' ====================================================================

    Dim SheetLaporan As Worksheet

    Dim TabelLaporan As ListObject

    Dim PilihanFilter As String

    

    Set SheetLaporan = ThisWorkbook.Sheets("KAS&BANK_LAPORAN HARIAN")

    Set TabelLaporan = SheetLaporan.ListObjects("TabelLaporanHarianKas")

    

    ' Ambil nilai dari Cell E10 yang sudah di-merge oleh

    PilihanFilter = Trim(SheetLaporan.Range("E10").Value)

    

    ' Supaya pergerakan saringannya mulus anti patah-patah

    Application.ScreenUpdating = False

    

    ' Jika tabel dalam keadaan kosong melompong, lepas filter dan keluar biar ga eror

    If TabelLaporan.ListRows.Count = 0 Then

        TabelLaporan.Range.AutoFilter Field:=8

        GoTo AkhirSaja

    End If

    

    ' ====================================================================

    ' LOGIKA EKSEKUSI PENYARINGAN (KOLOM 8 = JENIS AKTIVITAS)

    ' ====================================================================

    If UCase(PilihanFilter) = "SEMUA" Or PilihanFilter = "" Then

        ' Jika pilih "SEMUA", buka semua sumbatan data (Tampilkan seluruhnya)

        TabelLaporan.Range.AutoFilter Field:=8

    Else

        ' Jika pilih aktivitas spesifik, saring kolom ke-8 sesuai teks dropdown

        TabelLaporan.Range.AutoFilter Field:=8, Criteria1:=PilihanFilter

    End If



AkhirSaja:

    ' ?? SEKALIGUS UPDATE TEXT BOX KPI BIAR ANGKA YANG MUNCUL IKUT SINKRON!

    Call TembakKeTextBoxKPILaporanKas

    

    Application.ScreenUpdating = True

End Sub


