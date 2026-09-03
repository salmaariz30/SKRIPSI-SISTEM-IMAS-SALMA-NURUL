Attribute VB_Name = "Module97"
Sub PostingPenyesuaianBulananAuto()
    Dim tblPenyusutan As ListObject, tblJurnal As ListObject, tblCOA As ListObject
    Dim wsMenu As Worksheet, wsJurnal As Worksheet
    Dim namaBulan As String, buktiPenyesuaian As String
    Dim i As Long, dataRow As Long
    Dim sisaUmur As String
    Dim nominalPenyusutan As Double
    Dim namaAkunAset As String, nBeban As String, nAkumulasi As String
    Dim kBeban As Variant, kAkumulasi As Variant
    Dim checkDouble As Range
    Dim konfirmasi As VbMsgBoxResult
    
    Dim matchBeban As Variant, matchAkumulasi As Variant
    Dim newRowBeban As ListRow, newRowAkumulasi As ListRow
    
    ' 1. PENYESUAIAN ALAMAT SHEET & CELL DROPDOWN
    Set wsMenu = ThisWorkbook.Worksheets("DATA DEPRESIASI")
    namaBulan = wsMenu.Range("D5").Value ' Ambil nama bulan murni dari sel D5
    
    ' Safety check jika kosong atau jika ternyata statusnya sudah "Telah Penyesuaian"
    If namaBulan = "" Or InStr(namaBulan, "(Telah Penyesuaian)") > 0 Then Exit Sub
    
    ' SISTEM KONFIRMASI AWAL SEBELUM PROSES JURNAL
    konfirmasi = MsgBox("Apakah Anda yakin ingin memproses dan mem-posting " & _
                        "Jurnal Penyesuaian untuk periode " & namaBulan & "?", _
                        vbQuestion + vbYesNo, "Konfirmasi Posting Jurnal")
                        
    If konfirmasi = vbNo Then
        Exit Sub
    End If
    
    ' Mengunci objek tabel database
    On Error Resume Next
    Set tblPenyusutan = Range("TabelPenyusutan").ListObject
    Set tblJurnal = Range("TabelJurnalPenyesuaian").ListObject
    Set tblCOA = Range("TabelDataCOA").ListObject
    On Error GoTo 0
    
    If tblPenyusutan Is Nothing Or tblJurnal Is Nothing Or tblCOA Is Nothing Then
        MsgBox "Error: Struktur Tabel Penyusutan, TabelJurnalPenyesuaian, atau COA tidak ditemukan!", vbCritical, "Gagal"
        Exit Sub
    End If
    
    ' Format bukti jurnal otomatis: AJP-JAN-2026
    buktiPenyesuaian = "AJP-" & UCase(Left(namaBulan, 3)) & "-" & Year(Date)
    
    ' FILTER SAFETY: Sistem Validasi Anti-Duplikasi Jurnal
    If tblJurnal.ListRows.Count > 0 Then
        Set checkDouble = tblJurnal.ListColumns(2).DataBodyRange.Find(What:=buktiPenyesuaian, LookIn:=xlValues, LookAt:=xlWhole)
        If Not checkDouble Is Nothing Then
            MsgBox "Gagal! Jurnal Penyesuaian untuk periode " & namaBulan & " sudah pernah dimasukkan sebelumnya!", vbCritical, "Sistem Keamanan"
            Exit Sub
        End If
    End If
    
    ' Sinkronisasi Sistem Latar Belakang (Mencegah Screen Flicker)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' ====================================================================
    ' ?? OPERASI OPEN LOCK: Buka gembok sheet Jurnal Penyesuaian sebelum isi data
    ' ====================================================================
    Set wsJurnal = tblJurnal.Parent
    wsJurnal.Unprotect Password:="IMAS"
    ' ====================================================================
    
    dataRow = tblPenyusutan.ListRows.Count
    If dataRow = 0 Then GoTo Selesai
    
    For i = 1 To dataRow
        ' Ambil teks sisa umur dari Kolom 12
        sisaUmur = Trim(CStr(tblPenyusutan.DataBodyRange.Cells(i, 12).Value))
        
        ' LOGIKA: Hanya proses jika sisa umur BUKAN "0 Tahun 0 Bulan" dan tidak kosong
        If sisaUmur <> "0 Tahun 0 Bulan" And sisaUmur <> "" Then
            
            ' Ambil nama kelompok aset di Kolom 8
            namaAkunAset = Trim(tblPenyusutan.DataBodyRange.Cells(i, 8).Value)
            
            ' Ambil nominal informasi penyusutan bulanan di Kolom 13
            nominalPenyusutan = Val(tblPenyusutan.DataBodyRange.Cells(i, 13).Value)
            
            ' Validasi tambahan, pastikan nilainya di atas 0 sebelum di-jurnal
            If nominalPenyusutan > 0 Then
                
                '--- ATURAN NAMA AKUN BARU SESUAI POLA   ---
                nBeban = "Beban Penyusutan"
                
                Select Case namaAkunAset
                    Case "Bangunan / Pabrik", "Bangunan"
                        nAkumulasi = "Akum. Penyusutan Bangunan"
                    Case "Kendaraan"
                        nAkumulasi = "Akum. Penyusutan Kendaraan"
                    Case "Peralatan & Mesin", "Peralatan Bar", "Peralatan"
                        nAkumulasi = "Akum. Penyusutan Peralatan"
                    Case Else
                        nAkumulasi = "Akum. Penyusutan " & namaAkunAset
                End Select
                
                '--- PROSES PENCARI KODE DI COA ---
                matchBeban = Application.Match(nBeban, tblCOA.ListColumns(2).DataBodyRange, 0)
                matchAkumulasi = Application.Match(nAkumulasi, tblCOA.ListColumns(2).DataBodyRange, 0)
                
                ' Cek keamanan COA
                If IsError(matchBeban) Or IsError(matchAkumulasi) Then
                    ' Kunci kembali sebelum keluar mendadak akibat error COA
                    wsJurnal.Protect Password:="IMAS", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
                    wsJurnal.EnableSelection = xlNoSelection
                    
                    Application.ScreenUpdating = True
                    Application.EnableEvents = True
                    MsgBox "ERROR KODE AKUN KOSONG!" & vbCrLf & _
                           "Nama akun ini gak ketemu di COA kamu:" & vbCrLf & _
                           "- " & nBeban & vbCrLf & _
                           "- " & nAkumulasi & vbCrLf & vbCrLf & _
                           "Pastikan di COA tulisan dan kodenya sudah terdaftar ya!", vbCritical, "Gagal Cari Kode Akun"
                    Exit Sub
                Else
                    kBeban = Application.Index(tblCOA.ListColumns(1).DataBodyRange, matchBeban)
                    kAkumulasi = Application.Index(tblCOA.ListColumns(1).DataBodyRange, matchAkumulasi)
                End If
                
                ' LINE 1: DEBIT BEBAN
                Set newRowBeban = tblJurnal.ListRows.Add(AlwaysInsert:=True)
                With newRowBeban
                    .Range(4).NumberFormat = "@" ' Kunci format kolom Kode Akun ke Teks
                    .Range(1) = Date: .Range(2) = buktiPenyesuaian
                    .Range(3) = "Penyesuaian Penyusutan " & namaAkunAset & " - Periode " & namaBulan
                    .Range(4) = CStr(kBeban): .Range(5) = nBeban: .Range(6) = nominalPenyusutan: .Range(7) = 0
                    .Range(8) = "Beban Operasional"
                End With
                
                ' LINE 2: KREDIT AKUMULASI
                Set newRowAkumulasi = tblJurnal.ListRows.Add(AlwaysInsert:=True)
                With newRowAkumulasi
                    .Range(4).NumberFormat = "@" ' Kunci format kolom Kode Akun ke Teks
                    .Range(1) = Date: .Range(2) = buktiPenyesuaian
                    .Range(3) = "Penyesuaian Penyusutan " & namaAkunAset & " - Periode " & namaBulan
                    .Range(4) = CStr(kAkumulasi): .Range(5) = nAkumulasi: .Range(6) = 0: .Range(7) = nominalPenyusutan
                    .Range(8) = "Aset Tetap"
                End With
            End If
            
        End If
    Next i
    
    ' 2. MENGUBAH STATUS CELL D5 SETELAH BERHASIL TERPOSTING
    wsMenu.Range("D5").Value = namaBulan & " (Telah Penyesuaian)"
    
    MsgBox "Proses Berhasil! Jurnal Penyesuaian periode " & namaBulan & " telah ter-posting otomatis.", vbInformation, "Sistem Otomatisasi Sukses"

Selesai:
    ' ====================================================================
    ' ?? RE-LOCK SYSTEM: Kunci mutlak sheet Jurnal + Bekukan Shape Tombol biar mati total!
    ' ====================================================================
    wsJurnal.Protect Password:="IMAS", DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    wsJurnal.EnableSelection = xlNoSelection
    ' ====================================================================
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

