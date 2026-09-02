Attribute VB_Name = "Module10"
Sub UpdateBankCOA()
    Dim wsMaster As Worksheet, wsCOA As Worksheet
    Dim tblCOA As ListObject
    Dim rStart As Range, rEnd As Range
    Dim i As Long, startRow As Long, endRow As Long, jmlBank As Long
    Dim BarisTarget As Long, BarisBaru As ListRow
    
    ' ====================================================================
    ' ALAMAT RUMAH BARU   & DEKLARASI TABEL RESMI
    ' ====================================================================
    Set wsMaster = Sheets("DATA USAHA")
    Set wsCOA = Sheets("AKUNTANSI_DATA COA")
    Set tblCOA = wsCOA.ListObjects("TabelDataCOA") ' Mengunci target langsung ke tabel
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. Cari koordinat absolut Kas Kecil (Batas Atas) dan Piutang Usaha (Batas Bawah) di Kolom B
    Set rStart = wsCOA.Columns("B").Find("1-1111", LookIn:=xlValues, LookAt:=xlWhole)
    Set rEnd = wsCOA.Columns("B").Find("1-1200", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rStart Is Nothing And Not rEnd Is Nothing Then
        startRow = rStart.Row
        endRow = rEnd.Row
        
        ' ?? PAGAR PENGAMAN: Jaga-jaga kalau batas bawahnya bocor
        If (endRow - startRow) > 50 Or (endRow <= startRow) Then
            MsgBox "?? PERINGATAN  ! Jarak baris terdeteksi tidak wajar." & vbCrLf & _
                   "Proses dihentikan demi menjaga keamanan Akun Ekuitas  !", vbCritical, "Sistem Proteksi Nyala"
            GoTo AkhirProgram
        End If
        
        ' 2. BERSIHKAN BANK LAMA (Cara Legal Excel Table)
        ' Kita hapus baris tabelnya dari bawah ke atas agar indeksnya gak bergeser berantakan
        If endRow - startRow > 1 Then
            For i = (endRow - 1) To (startRow + 1) Step -1
                ' Menghapus baris tabel secara utuh khusus di area tabel saja
                tblCOA.ListRows(i - tblCOA.HeaderRowRange.Row).Delete
            Next i
        End If
        
        ' Cari ulang koordinat startRow setelah bank lama disapu bersih
        startRow = wsCOA.Columns("B").Find("1-1111", LookIn:=xlValues, LookAt:=xlWhole).Row
        
        ' 3. LOOP AMBIL DATA BARU DARI DATA USAHA (Baris 7 sampai 20)
        jmlBank = 0
        For i = 7 To 16
            If wsMaster.Cells(i, "XFD").Value <> "" Then
                
                jmlBank = jmlBank + 1
                
                ' Hitung posisi index baris baru di dalam tabel
                BarisTarget = startRow - tblCOA.HeaderRowRange.Row + jmlBank
                
                ' 4. SISIPKAN BARIS BARU SECARA LEGAL DI DALAM EXCEL TABLE
                ' Ini cara resmi yang disukai Excel, dijamin Error 1004 langsung tobat!
                Set BarisBaru = tblCOA.ListRows.Add(BarisTarget)
                
                ' 5. PROSES PENGISIAN DATA BERDASARKAN URUTAN KOLOM TABEL (B sampai G)
                ' Range(1) = Kolom B, Range(2) = Kolom C, dst.
                BarisBaru.Range(1).Value = "1-" & (1111 + jmlBank)                ' Kolom B: Kode Akun
                BarisBaru.Range(2).Value = wsMaster.Cells(i, "XFD").Value          ' Kolom C: Nama Akun (XFD)
                BarisBaru.Range(3).Value = "Aset"                                  ' Kolom D: Kelompok
                BarisBaru.Range(4).Value = "Aset Lancar"                           ' Kolom E: Sub-Kelompok
                BarisBaru.Range(5).Value = "Kas & Bank"                            ' Kolom F: Kategori Akun
                BarisBaru.Range(6).Value = "Debit"                                 ' Kolom G: Saldo Normal
                
                ' ? Bonus  : Untuk kolom "Tipe Akun" (Kolom H / Range 7), kita isi General Ledger!
                If tblCOA.ListColumns.Count >= 7 Then
                    BarisBaru.Range(7).Value = "General Ledger"
                End If
                
            End If
        Next i
        
        MsgBox "Sinkronisasi Berhasil!", vbInformation, "Sistem Sukses"
        
    Else
        MsgBox "Error: Akun '1-1111' atau '1-1200' tidak ditemukan di Kolom B COA!", vbCritical
    End If

AkhirProgram:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

