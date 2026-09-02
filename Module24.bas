Attribute VB_Name = "Module24"
Sub SimpanTransaksiPindahSaldo()
    ' ==========================================================
    ' MODUL GABUNGAN: LAPORAN HARIAN + JURNAL UMUM BIAYA ADMIN
    ' ==========================================================
    Dim LembarForm As Worksheet: Set LembarForm = Sheets("KAS&BANK_INPUT PINDAH SALDO")
    Dim LembarData As Worksheet: Set LembarData = Sheets("KAS&BANK_LAPORAN HARIAN")
    Dim TabelData As ListObject: Set TabelData = LembarData.ListObjects("TabelLaporanHarianKas")
    Dim tblJurnal As ListObject: Set tblJurnal = Range("TabelJurnalUmum").ListObject
    Dim tblCOA As ListObject: Set tblCOA = Range("TabelDataCOA").ListObject
    Dim SheetJurnal As Worksheet: Set SheetJurnal = tblJurnal.Parent
    Dim BarisBaru As ListRow
    
    ' 1. AMBIL DATA DARI FORM MERGER
    Dim V_Tanggal As Date: V_Tanggal = LembarForm.Range("D11").Value
    Dim V_NoBukti As String: V_NoBukti = LembarForm.Range("G11").Value
    Dim V_AkunTujuan As String: V_AkunTujuan = LembarForm.Range("D15").Value
    Dim V_AkunAsal As String: V_AkunAsal = LembarForm.Range("G15").Value
    Dim V_Nominal As Double: V_Nominal = Val(LembarForm.Range("D20").Value)
    Dim V_BiayaAdmin As Double: V_BiayaAdmin = Val(LembarForm.Range("G20").Value)
    Dim V_Deskripsi As String: V_Deskripsi = LembarForm.Range("D24").Value
    
    ' 2. VALIDASI INPUT UTAMA
    If V_Tanggal = 0 Or V_NoBukti = "" Or V_AkunTujuan = "" Or V_AkunAsal = "" Or V_Nominal = 0 Then
        MsgBox "Mohon lengkapi data Tanggal, No Bukti, Akun, dan Nominal terlebih dahulu!", vbExclamation, "Input Tidak Valid"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False ' MATIKAN LAYAR (SPEED MODE ON! ?)
    
    ' BUKA PROTEKSI SHEET-SHEET TARGET
    LembarData.Unprotect Password:="IMAS"
    SheetJurnal.Unprotect Password:="IMAS"
    
    ' 3. AMBIL NOMOR URUT TERAKHIR LAPORAN HARIAN
    Dim NoUrutTerakhir As Long
    If TabelData.ListRows.Count = 0 Then NoUrutTerakhir = 0 Else NoUrutTerakhir = Application.WorksheetFunction.Max(TabelData.ListColumns(1).DataBodyRange)
    
    ' 4. PROSES POSTING KE LAPORAN HARIAN KAS
    ' --- BARIS 1: KREDIT AKUN ASAL (Uang Keluar) ---
    With TabelData.ListRows.Add
        .Range(1) = NoUrutTerakhir + 1: .Range(2) = V_Tanggal: .Range(3) = V_AkunAsal: .Range(4) = 0
        .Range(5) = V_Nominal: .Range(6) = V_Deskripsi & " (Rekening Asal)": .Range(7) = V_NoBukti: .Range(8) = "-"
    End With
    
    ' --- BARIS 2: DEBIT AKUN TUJUAN (Uang Masuk) ---
    With TabelData.ListRows.Add
        .Range(1) = NoUrutTerakhir + 2: .Range(2) = V_Tanggal: .Range(3) = V_AkunTujuan: .Range(4) = V_Nominal
        .Range(5) = 0: .Range(6) = V_Deskripsi & " (Rekening Tujuan)": .Range(7) = V_NoBukti: .Range(8) = "-"
    End With
    
    ' --- BARIS 3: BIAYA ADMIN DI LAPORAN HARIAN (Jika Ada) ---
    If V_BiayaAdmin > 0 Then
        With TabelData.ListRows.Add
            .Range(1) = NoUrutTerakhir + 3: .Range(2) = V_Tanggal: .Range(3) = V_AkunAsal: .Range(4) = 0
            .Range(5) = V_BiayaAdmin: .Range(6) = "Biaya Admin " & V_Deskripsi: .Range(7) = V_NoBukti: .Range(8) = "Operasional"
        End With
    End If
    
    ' 5. PROSES JURNAL UMUM OTOMATIS UNTUK BIAYA ADMIN (Jika Ada)
    If V_BiayaAdmin > 0 Then
        Dim namaBebanAdmin As String: namaBebanAdmin = "Beban Administrasi Bank"
        
        ' Sisir Kode Akun di COA
        Dim barisBevan As Variant: barisBevan = Application.Match(namaBebanAdmin, tblCOA.ListColumns(2).DataBodyRange, 0)
        Dim BarisKas As Variant: BarisKas = Application.Match(V_AkunAsal, tblCOA.ListColumns(2).DataBodyRange, 0)
        
        ' Ambil Kode Akun (Jika tidak ada di COA, ingatkan ??)
        If IsError(barisBevan) Or IsError(BarisKas) Then
            ' KUNCI KEMBALI SHEET SEBELUM EXIT AGAR TETAP AMAN
            LembarData.Protect Password:="IMAS", AllowFiltering:=True
            SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
            Application.ScreenUpdating = True
            MsgBox "nama akun '" & namaBebanAdmin & "' atau '" & V_AkunAsal & "' tidak ditemukan di COA! Jurnal Admin gagal.", vbCritical, "Error COA"
            Exit Sub
        End If
        
        Dim kodeBebanAdmin As Variant: kodeBebanAdmin = Application.Index(tblCOA.ListColumns(1).DataBodyRange, barisBevan)
        Dim kodeKasAsal As Variant: kodeKasAsal = Application.Index(tblCOA.ListColumns(1).DataBodyRange, BarisKas)
        
        ' --- BARIS 1 JURNAL: DEBIT BEBAN ADMIN ---
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Biaya Admin - " & V_Deskripsi: .Range(4) = kodeBebanAdmin
            .Range(5) = namaBebanAdmin: .Range(6) = V_BiayaAdmin: .Range(7) = 0: .Range(8) = "Kas dan Bank": .Range(9) = "Tidak"
        End With
        
        ' --- BARIS 2 JURNAL: KREDIT KAS BANK ASAL ---
        With tblJurnal.ListRows.Add(AlwaysInsert:=True)
            .Range(1) = V_Tanggal: .Range(2) = V_NoBukti: .Range(3) = "Biaya Admin - " & V_Deskripsi: .Range(4) = kodeKasAsal
            .Range(5) = V_AkunAsal: .Range(6) = 0: .Range(7) = V_BiayaAdmin: .Range(8) = "Kas dan Bank": .Range(9) = "Tidak"
        End With
    End If
    
    ' TUTUP PROTEKSI SHEET-SHEET TARGET
    LembarData.Protect Password:="IMAS", AllowFiltering:=True
    SheetJurnal.Protect Password:="IMAS", AllowFiltering:=True
    
    ' 6. SILENT CLEAN FORM (Otomatis Bersih Kilat)
    With LembarForm
        .Range("D11").MergeArea.ClearContents: .Range("G11").MergeArea.ClearContents
        .Range("D15").MergeArea.ClearContents: .Range("G15").MergeArea.ClearContents
        .Range("D20").MergeArea.ClearContents: .Range("G20").MergeArea.ClearContents
        .Range("D24").MergeArea.ClearContents
        .Activate: .Range("D11").Select ' Kursor balik manis ke Tanggal
    End With
    
    Application.ScreenUpdating = True
    MsgBox "Pindah Saldo Sukses Dicatat!", vbInformation, "Sistem Sukses"
End Sub
