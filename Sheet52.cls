VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cellBayar As Range: Set cellBayar = Range("I11")
    Dim cellPencatatan As Range: Set cellPencatatan = Range("F15")
    Dim cellStatusLunas As Range: Set cellStatusLunas = Range("L15") ' <-- DAFTARKAN L15 DI SINI,  !
    Dim i As Long
    Const PWD As String = "IMAS" ' <-- Deklarasi Password di sini
    
    ' KUNCI UTAMA: Ditambah cellStatusLunas (L15) agar macro mau jalan kalau sel ini diubah user
    If Intersect(Target, cellBayar) Is Nothing And _
       Intersect(Target, cellPencatatan) Is Nothing And _
       Intersect(Target, cellStatusLunas) Is Nothing Then Exit Sub

    ' Matikan screen biar gak kedap-kedip di laptop ASUS
    Application.ScreenUpdating = False
    
    ' ==========================================================
    ' BUKA KUNCI SHEET SEBELUM EKSEKUSI PERUBAHAN DATA
    ' ==========================================================
    Me.Unprotect Password:=PWD
    ' ==========================================================
    
    ' ==========================================================
    ' TAMBAHAN  : JIKA L15 = "Lunas", MAKA L19 DIHAPUS TOTAL
    ' ==========================================================
    If Not Intersect(Target, cellStatusLunas) Is Nothing Then
        If cellStatusLunas.Value = "Lunas" Then
            Application.EnableEvents = False
            Range("L19").MergeArea.ClearContents ' Pakai MergeArea biar aman kalau L19 ternyata di-merge
            Application.EnableEvents = True
        End If
    End If
    ' ==========================================================
    
    ' --- 1. LOGIKA VALIDASI ATURAN   (KREDIT vs REKAP) ---
    If cellBayar.Value = "Kredit" And cellPencatatan.Value = "Rekap Harian" Then
        MsgBox "Penjualan Kredit Tidak Diperbolehkan Rekap Harian! Wajib Per Transaksi!", _
               vbCritical, "Peringatan"
               
        Application.EnableEvents = False
        cellPencatatan.Value = "Per Transaksi" ' Paksa balik ke jalan yang benar
        Application.EnableEvents = True
    End If

    ' --- 2. LOGIKA BUMI HANGUS DATA FORM LAWAN (KOLOM D & J DI-MERGE) ---
    Application.EnableEvents = False ' Matikan events agar tidak memicu macro berulang
    On Error Resume Next
    
    Select Case cellPencatatan.Value
        Case "Rekap Harian"
            ' ?? SAPU BERSIH DATA FORM PER TRANSAKSI (Baris 24 sampai 38)
            For i = 24 To 38
                Range("D" & i).MergeArea.ClearContents ' Bersihkan Kolom D (Nomor Nota Merger)
                Range("F" & i).ClearContents           ' Bersihkan Kolom F (Detail/Nama Polosan)
                Range("J" & i).MergeArea.ClearContents ' Bersihkan Kolom J (Nominal/Jumlah Merger - SEARAH  !)
            Next i
            Range("L41").ClearContents                 ' Bersihkan Grand Total L41
            
        Case "Per Transaksi"
            ' ?? SAPU BERSIH DATA FORM REKAP HARIAN (Baris 54 sampai 83)
            For i = 54 To 83
                Range("D" & i).MergeArea.ClearContents ' Bersihkan Kolom D (Nomor Nota Merger)
                Range("F" & i).ClearContents           ' Bersihkan Kolom F (Detail/Nama Polosan)
                Range("J" & i).MergeArea.ClearContents ' Bersihkan Kolom J (Nominal/Jumlah Merger - SEARAH  !)
            Next i
            Range("L86").ClearContents                 ' Bersihkan Grand Total L86
    End Select
    
    On Error GoTo 0
    Application.EnableEvents = True

    ' --- 3. LOGIKA TAMPILAN PREMIUM SHOW/HIDDEN ROWS ---
    Select Case cellPencatatan.Value
        Case "Per Transaksi"
            Rows("21:50").EntireRow.Hidden = False
            Rows("51:91").EntireRow.Hidden = True
            
        Case "Rekap Harian"
            Rows("21:49").EntireRow.Hidden = True
            Rows("50:91").EntireRow.Hidden = False
            
        Case Else
            ' Jika kosong atau memilih yang lain, buka semua biar kelihatan
            Rows("21:50").EntireRow.Hidden = False
    End Select

    ' ==========================================================
    ' KUNCI KEMBALI SHEET AGAR USER TIDAK BISA MERUSAK FORMULA
    ' ==========================================================
    Me.Protect Password:=PWD, AllowFiltering:=True
    ' ==========================================================

    ' Pastikan Excel selalu kembali dalam keadaan sadar dan aktif
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

