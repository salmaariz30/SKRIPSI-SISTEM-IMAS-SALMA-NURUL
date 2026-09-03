Attribute VB_Name = "Module104"
Sub PostingPenyesuaianAmortisasiAuto()
    Dim tblPengeluaran As ListObject, tblJurnal As ListObject, tblCOA As ListObject
    Dim wsMenu As Worksheet
    Dim namaBulan As String, buktiPenyesuaian As String
    Dim i As Long, dataRow As Long
    Dim nominalAmortisasi As Double
    Dim nBeban As String, nKreditPrepaid As String
    Dim kBeban As Variant, kKreditPrepaid As Variant
    Dim checkDouble As Range
    Dim konfirmasi As VbMsgBoxResult
    
    Dim matchBeban As Variant, matchPrepaid As Variant
    Dim newRowBeban As ListRow, newRowPrepaid As ListRow
    
    ' 1. AMBIL PERIODE BULAN DARI SHEET TEMPAT   BERADA SEKARANG!
    Set wsMenu = ThisWorkbook.Worksheets("PENGELUARAN USAHA_DAFTAR")
    namaBulan = wsMenu.Range("E5").Cells(1, 1).Value ' Ambil nilai sel merger E5
    
    ' Jika tulisan di E5 ada embel-embel (Telah Penyesuaian), kita bersihkan dulu teks murninya
    If InStr(namaBulan, " (Telah Penyesuaian)") > 0 Then
        namaBulan = Replace(namaBulan, " (Telah Penyesuaian)", "")
    End If
    
    If namaBulan = "" Then
        MsgBox "Silakan pilih periode bulan terlebih dahulu di sel E5!", vbExclamation, "Periode Kosong"
        Exit Sub
    End If
    
    ' SISTEM KONFIRMASI AWAL SEBELUM PROSES
    konfirmasi = MsgBox("Apakah Anda yakin ingin memproses Jurnal Penyesuaian " & _
                        "Beban Dibayar di Muka untuk periode " & namaBulan & "?", _
                        vbQuestion + vbYesNo, "Konfirmasi Amortisasi")
                        
    If konfirmasi = vbNo Then Exit Sub
    
    ' Mengunci objek tabel database
    On Error Resume Next
    Set tblPengeluaran = wsMenu.ListObjects(1) ' Otomatis mengunci tabel pertama di sheet ini
    Set tblJurnal = Range("TabelJurnalPenyesuaian").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    On Error GoTo 0
    
    If tblPengeluaran Is Nothing Or tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: Struktur Tabel Pengeluaran, TabelJurnalPenyesuaian, atau COA tidak ditemukan!", vbCritical, "Gagal"
        Exit Sub
    End If
    
    ' Format bukti jurnal otomatis: AJP-PRE-JAN-2026
    buktiPenyesuaian = "AJP-PRE-" & UCase(Left(namaBulan, 3)) & "-" & Year(Date)
    
    ' FILTER SAFETY: Sistem Validasi Anti-Duplikasi Jurnal
    If tblJurnal.ListRows.Count > 0 Then
        Set checkDouble = tblJurnal.ListColumns(2).DataBodyRange.Find(What:=buktiPenyesuaian, LookIn:=xlValues, LookAt:=xlWhole)
        If Not checkDouble Is Nothing Then
            MsgBox "Gagal! Jurnal Penyesuaian Beban Dibayar di Muka untuk periode " & namaBulan & " sudah pernah di-posting!", vbCritical, "Sistem Keamanan"
            Exit Sub
        End If
    End If
    
    ' Sinkronisasi Sistem Latar Belakang
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    dataRow = tblPengeluaran.ListRows.Count
    If dataRow = 0 Then GoTo Selesai
    
    For i = 1 To dataRow
        nominalAmortisasi = Val(tblPengeluaran.DataBodyRange.Cells(i, 15).Value)
        
        If nominalAmortisasi > 0 Then
            nBeban = Trim(tblPengeluaran.DataBodyRange.Cells(i, 5).Value)
            nKreditPrepaid = "Beban Dibayar di Muka"
            
            matchBeban = Application.Match(nBeban, tblCOA.ListColumns(2).DataBodyRange, 0)
            matchPrepaid = Application.Match(nKreditPrepaid, tblCOA.ListColumns(2).DataBodyRange, 0)
            
            If IsError(matchBeban) Or IsError(matchPrepaid) Then
                Application.ScreenUpdating = True
                Application.EnableEvents = True
                MsgBox "ERROR KODE AKUN TIDAK KETEMU DI COA!" & vbCrLf & _
                       "Periksa ejaan nama akun berikut di tabel COA:" & vbCrLf & _
                       "- " & nBeban & vbCrLf & _
                       "- " & nKreditPrepaid, vbCritical, "Gagal Posting"
                Exit Sub
            Else
                kBeban = Application.Index(tblCOA.ListColumns(1).DataBodyRange, matchBeban)
                kKreditPrepaid = Application.Index(tblCOA.ListColumns(1).DataBodyRange, matchPrepaid)
            End If
            
            ' LINE 1: DEBIT
            Set newRowBeban = tblJurnal.ListRows.Add(AlwaysInsert:=True)
            With newRowBeban
                .Range(4).NumberFormat = "@"
                .Range(1) = Date: .Range(2) = buktiPenyesuaian
                .Range(3) = "Amortisasi " & nBeban & " - Periode " & namaBulan
                .Range(4) = CStr(kBeban): .Range(5) = nBeban: .Range(6) = nominalAmortisasi: .Range(7) = 0
                .Range(8) = "Penyusutan Beban"
            End With
            
            ' LINE 2: KREDIT
            Set newRowPrepaid = tblJurnal.ListRows.Add(AlwaysInsert:=True)
            With newRowPrepaid
                .Range(4).NumberFormat = "@"
                .Range(1) = Date: .Range(2) = buktiPenyesuaian
                .Range(3) = "Amortisasi " & nBeban & " - Periode " & namaBulan
                .Range(4) = CStr(kKreditPrepaid): .Range(5) = nKreditPrepaid: .Range(6) = 0: .Range(7) = nominalAmortisasi
                .Range(8) = "Penyusutan Beban"
            End With
        End If
    Next i
    
    ' 2. MENGUBAH STATUS CELL E5 DI SHEET PENGELUARAN USAHA MENJADI TELAH PENYESUAIAN
    wsMenu.Range("E5").Value = namaBulan & " (Telah Penyesuaian)"
    
    MsgBox "Proses Berhasil! Penyesuaian Beban Dibayar di Muka periode " & namaBulan & " telah ter-posting.", vbInformation, "Sukses"

Selesai:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

