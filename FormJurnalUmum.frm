VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormJurnalUmum 
   Caption         =   "Form Hapus Jurnal Umum"
   ClientHeight    =   5390
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8060
   OleObjectBlob   =   "FormJurnalUmum.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormJurnalUmum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ====================================================================
' MASTER CODE: USERFORM PENGHAPUSAN JURNAL UMUM (MULTI-BARIS)
' DILENGKAPI AUTO BYPASS PROTECT SHEET (PASSWORD: "IMAS")
' ====================================================================

' --- SENSOR SAAT USERFORM PERTAMA KALI DIBUKA ---
Private Sub UserForm_Initialize()
    Dim wsJurnal As Worksheet
    Dim tblJurnal As ListObject
    Dim cell As Range
    Dim dictUnique As Object
    Const PWD As String = "IMAS"
    
    ' Bersihkan detail data di awal biar rapi berbentuk strip "-"
    Call BersihkanDetail
    
    Set wsJurnal = ThisWorkbook.Worksheets("JURNAL UMUM")
    On Error Resume Next
    Set tblJurnal = wsJurnal.ListObjects("TabelJurnalUmum")
    On Error GoTo 0
    
    If tblJurnal Is Nothing Then
        MsgBox "TabelJurnalUmum tidak ditemukan di sheet JURNAL UMUM!", vbCritical, "Error"
        Exit Sub
    End If
    
    ' ?? BUKA PENGAMAN: Izinkan kode membaca data meskipun sheet terkunci
    wsJurnal.Unprotect Password:=PWD
    
    ' Gunakan Scripting Dictionary untuk menyaring No. Bukti agar tidak duplikat di Dropdown
    Set dictUnique = CreateObject("Scripting.Dictionary")
    
    ' Isi Dropdown dengan menyisir kolom ke-2 dari TabelJurnalUmum (Kolom No. Bukti Transaksi)
    If Not tblJurnal.DataBodyRange Is Nothing Then
        For Each cell In tblJurnal.ListColumns(2).DataBodyRange
            If cell.Value <> "" And Not dictUnique.Exists(cell.Value) Then
                dictUnique.Add cell.Value, Nothing
                Me.ComboBoxBukti.AddItem cell.Value
            End If
        Next cell
    End If
    
    ' ?? KUNCI KEMBALI: Kembalikan proteksi sheet
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
End Sub

' --- OTOMATIS MUNCUL DETAIL PREVIEW SAAT NO. BUKTI DIUBAH ---
Private Sub ComboBoxBukti_Change()
    Dim wsJurnal As Worksheet
    Dim tblJurnal As ListObject
    Dim noBukti As String
    Dim i As Long
    Dim jmlBaris As Long
    Dim deskripsiFirst As String
    Dim namaAkunList As String
    Dim totalDebit As Double, totalKredit As Double
    Const PWD As String = "IMAS"
    
    noBukti = Me.ComboBoxBukti.Value
    
    If noBukti = "" Then
        Call BersihkanDetail
        Exit Sub
    End If
    
    Set wsJurnal = ThisWorkbook.Worksheets("JURNAL UMUM")
    Set tblJurnal = wsJurnal.ListObjects("TabelJurnalUmum")
    
    ' Initialize nilai perhitungan
    jmlBaris = 0
    deskripsiFirst = ""
    namaAkunList = ""
    totalDebit = 0
    totalKredit = 0
    
    ' ?? BUKA PENGAMAN UNTUK MEMBACA DATA
    wsJurnal.Unprotect Password:=PWD
    
    If Not tblJurnal.DataBodyRange Is Nothing Then
        For i = 1 To tblJurnal.ListRows.Count
            ' Cocokkan No. Bukti di Kolom ke-2
            If Trim(CStr(tblJurnal.DataBodyRange.Cells(i, 2).Value)) = Trim(noBukti) Then
                jmlBaris = jmlBaris + 1
                
                ' 1. Ambil Deskripsi dari baris pertama yang ditemukan (Kolom ke-3)
                ' Silakan sesuaikan indeks kolom jika Deskripsi bukan di kolom 3
                If deskripsiFirst = "" Then
                    deskripsiFirst = tblJurnal.DataBodyRange.Cells(i, 3).Value
                End If
                
                ' 2. Gabungkan Nama Akun (Kolom ke-5)
                Dim akunSekarang As String
                akunSekarang = Trim(CStr(tblJurnal.DataBodyRange.Cells(i, 5).Value))
                If akunSekarang <> "" Then
                    If InStr(namaAkunList, akunSekarang) = 0 Then ' Biar gak dobel di list preview
                        If namaAkunList = "" Then
                            namaAkunList = akunSekarang
                        Else
                            namaAkunList = namaAkunList & ", " & akunSekarang
                        End If
                    End If
                End If
                
                ' 3. Akumulasi Total Debit (Kolom ke-6) & Kredit (Kolom ke-7)
                totalDebit = totalDebit + Val(tblJurnal.DataBodyRange.Cells(i, 6).Value)
                totalKredit = totalKredit + Val(tblJurnal.DataBodyRange.Cells(i, 7).Value)
            End If
        Next i
    End If
    
    ' Tampilkan ke Label Form jika data ditemukan
    If jmlBaris > 0 Then
        Me.LabelDeskripsiIsi.Caption = IIf(deskripsiFirst = "", "-", deskripsiFirst)
        Me.LabelJumlahBarisIsi.Caption = jmlBaris & " Baris"
        Me.LabelNamaAkunIsi.Caption = namaAkunList
        Me.LabelTotalIsi.Caption = Format(totalDebit, "Rp #,##0") & " (D) - " & Format(totalKredit, "Rp #,##0") & " (K)"
    Else
        Call BersihkanDetail
    End If
    
    ' ?? KUNCI KEMBALI
    wsJurnal.Protect Password:=PWD, AllowFiltering:=True
End Sub

' --- PENYAPU BERSIH LABEL PREVIEW ---
Sub BersihkanDetail()
    Me.LabelDeskripsiIsi.Caption = "-"
    Me.LabelJumlahBarisIsi.Caption = "-"
    Me.LabelNamaAkunIsi.Caption = "-"
    Me.LabelTotalIsi.Caption = "-"
End Sub

' --- TOMBOL UTAMA (HAPUS MASAL AYAT JURNAL BERDASARKAN NO. BUKTI) ---
Private Sub CommandHapus_Click()
    Dim wsJurnal As Worksheet
    Dim tblJurnal As ListObject
    Dim noBukti As String
    Dim ans As VbMsgBoxResult
    Dim i As Long, countTerhapus As Long
    Const PWD As String = "IMAS"
    
    ' 1. Validasi jika dropdown belum dipilih
    If Me.ComboBoxBukti.Value = "" Then
        MsgBox "Pilih No. Bukti Transaksi yang akan Dihapus!", vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    noBukti = Me.ComboBoxBukti.Value
    
    ' Konfirmasi tindakan krusial ke user
    ans = MsgBox("Apakah anda yakin ingin menghapus transaksi No. Bukti: '" & noBukti & "'?" & vbCrLf & _
                 "Semua baris ayat jurnal yang terkait dengan nomor bukti ini akan DIHAPUS TOTAL!", _
                 vbQuestion + vbYesNo, "Konfirmasi Hapus Jurnal")
                 
    If ans = vbYes Then
        ' ??? BEKUKAN VISUAL: Biar proses berjalan cepat & smooth tanpa kedip
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        Set wsJurnal = ThisWorkbook.Worksheets("JURNAL UMUM")
        Set tblJurnal = wsJurnal.ListObjects("TabelJurnalUmum")
        
        ' ?? BUKA PENGAMAN SHEET JURNAL
        wsJurnal.Unprotect Password:=PWD
        
        countTerhapus = 0
        If Not tblJurnal.DataBodyRange Is Nothing Then
            ' WAJIB LOOPING MUNDUR (Dari bawah ke atas) agar index baris tabel tidak rusak saat didelete
            For i = tblJurnal.ListRows.Count To 1 Step -1
                If Trim(CStr(tblJurnal.DataBodyRange.Cells(i, 2).Value)) = Trim(noBukti) Then
                    tblJurnal.ListRows(i).Delete
                    countTerhapus = countTerhapus + 1
                End If
            Next i
        End If
        
        ' ?? KUNCI KEMBALI SHEETNYA
        wsJurnal.Protect Password:=PWD, AllowFiltering:=True
        
        ' AUTOMATIC REFRESH ALL PIVOT + NYALAKAN LAYAR KEMBALI
        ThisWorkbook.RefreshAll
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        ' Notifikasi Sukses
        MsgBox "Sukses! " & countTerhapus & " baris ayat jurnal dengan No. Bukti '" & noBukti & "' berhasil dihapus bersih!", _
               vbInformation, "Sistem Sukses"
        
        ' Tutup UserForm
        Unload Me
    End If
End Sub

' --- TOMBOL BATALKAN ---
Private Sub CommandBatalkan_Click()
    Unload Me
End Sub

