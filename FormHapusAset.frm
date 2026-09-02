VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormHapusAset 
   Caption         =   "Formulir Koreksi Aset Tetap"
   ClientHeight    =   5460
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7770
   OleObjectBlob   =   "FormHapusAset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormHapusAset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ====================================================================
' MASTER CODE: USERFORM MANAJEMEN ASET - DROPDOWN LANGSUNG TEMBAK TABEL
' DILENGKAPI AUTO BYPASS PROTECT SHEET (PASSWORD: "IMAS")
' ====================================================================

' --- SENSOR SAAT USERFORM PERTAMA KALI DIBUKA ---
Private Sub UserForm_Initialize()
    Dim wsDepr As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Const PWD As String = "IMAS"
    
    ' Bersihkan detail data di awal biar rapi berbentuk strip "-"
    Call BersihkanDetail
    
    Set wsDepr = ThisWorkbook.Worksheets("DATA DEPRESIASI")
    Set tbl = wsDepr.ListObjects("TabelPenyusutan")
    
    ' ?? BUKA PENGAMAN: Izinkan kode membaca data meskipun sheet terkunci
    wsDepr.Unprotect Password:=PWD
    
    ' Isi Dropdown murni menyisir kolom ke-2 dari TabelPenyusutan (Kolom Nama Aset)
    If Not tbl.DataBodyRange Is Nothing Then
        For Each cell In tbl.ListColumns(2).DataBodyRange
            If cell.Value <> "" Then
                Me.cmbAset.AddItem cell.Value
            End If
        Next cell
    End If
    
    ' ?? KUNCI KEMBALI: Kembalikan satpam proteksi sheet
    wsDepr.Protect Password:=PWD, AllowFiltering:=True
End Sub

' --- OTOMATIS MUNCUL DETAIL SAAT DROPDOWN DIUBAH ---
Private Sub cmbAset_Change()
    Dim wsDepr As Worksheet
    Dim tbl As ListObject
    Dim r As Range
    Dim rowIdx As Long
    Const PWD As String = "IMAS"
    
    If Me.cmbAset.Value = "" Then
        Call BersihkanDetail
        Exit Sub
    End If
    
    Set wsDepr = ThisWorkbook.Worksheets("DATA DEPRESIASI")
    Set tbl = wsDepr.ListObjects("TabelPenyusutan")
    
    ' ?? BUKA PENGAMAN: Buka proteksi agar pencarian (Find) tidak diblokir sistem
    wsDepr.Unprotect Password:=PWD
    
    On Error Resume Next
    Set r = tbl.ListColumns(2).DataBodyRange.Find(What:=Me.cmbAset.Value, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not r Is Nothing Then
        rowIdx = r.Row - tbl.Range.Row
        
        ' Menampilkan data akurat ke label pilihan
        Me.lblKode.Caption = tbl.DataBodyRange.Cells(rowIdx, 1).Value       ' Kode Aset
        Me.lblNama.Caption = tbl.DataBodyRange.Cells(rowIdx, 2).Value       ' Nama Aset
        Me.lblHarga.Caption = Format(tbl.DataBodyRange.Cells(rowIdx, 4).Value, "Rp #,##0") ' Harga
        Me.lblTanggal.Caption = Format(tbl.DataBodyRange.Cells(rowIdx, 5).Value, "dd-mmm-yyyy") ' Tanggal
        Me.lblSisaUmur.Caption = tbl.DataBodyRange.Cells(rowIdx, 12).Value  ' Sisa Umur
    Else
        Call BersihkanDetail
    End If
    
    ' ?? KUNCI KEMBALI: Amankan sheet kembali
    wsDepr.Protect Password:=PWD, AllowFiltering:=True
End Sub

' --- REVISI PENYAPU BERSIH LABEL SESUAI GAYA DROPDOWN ---
Sub BersihkanDetail()
    Me.lblKode.Caption = "-"
    Me.lblNama.Caption = "-"
    Me.lblHarga.Caption = "-"
    Me.lblTanggal.Caption = "-"
    Me.lblSisaUmur.Caption = "-"
End Sub

Private Sub btnHapus_Click()
    Dim wsDepr As Worksheet, wsJurnal As Worksheet
    Dim tblPenyusutan As ListObject, tblJurnal As ListObject
    Dim foundRow As Range
    Dim ans As VbMsgBoxResult
    Dim namaAset As String, kodeAset As String
    Dim i As Long
    Const PWD As String = "IMAS"
    
    ' 1. Validasi jika dropdown belum dipilih di UserForm
    If Me.cmbAset.Value = "" Then
        MsgBox "Pilih Aset yang akan Dihapus!", vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    ' 2. Ambil parameter target hapus langsung dari komponen UserForm
    namaAset = Me.cmbAset.Value
    kodeAset = Trim(Me.lblKode.Caption) ' Bersihkan spasi dari label form
    
    ' Set Objek Sheet dan Tabel secara presisi
    Set wsDepr = ThisWorkbook.Worksheets("DATA DEPRESIASI")
    Set tblPenyusutan = wsDepr.ListObjects("TabelPenyusutan")
    
    Set wsJurnal = ThisWorkbook.Worksheets("JURNAL UMUM")
    On Error Resume Next
    Set tblJurnal = wsJurnal.ListObjects("TabelJurnalUmum")
    On Error GoTo 0
    
    ' Cari baris di Tabel Penyusutan berdasarkan Nama Aset
    Set foundRow = tblPenyusutan.ListColumns(2).DataBodyRange.Find(What:=namaAset, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundRow Is Nothing Then
        
        ' Konfirmasi tindakan krusial ke user
        ans = MsgBox("Apakah anda yakin ingin menghapus aset '" & namaAset & "' (Kode: " & kodeAset & ")?" & vbCrLf & _
                     "Tindakan ini akan menghapus data di Tabel Penyusutan", _
                     vbQuestion + vbYesNo, "Konfirmasi Hapus Total")
                     
        If ans = vbYes Then
            ' ??? BEKUKAN VISUAL: Biar proses cepat & smooth
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            
            ' ?? BUKA PENGAMAN KEDUA SHEET
            wsDepr.Unprotect Password:=PWD
            wsJurnal.Unprotect Password:=PWD
            
            
            ' =============================================================
            ' PROSES 2: HAPUS BARIS DI TABEL PENYUSUTAN
            ' =============================================================
            ' Menghapus baris berdasarkan posisi index data row yang presisi
            Dim targetRowIndex As Long
            targetRowIndex = foundRow.Row - tblPenyusutan.DataBodyRange.Row + 1
            tblPenyusutan.ListRows(targetRowIndex).Delete
            
            ' =============================================================
            ' ?? KUNCI KEMBALI KEDUA SHEET
            ' =============================================================
            wsDepr.Protect Password:=PWD, AllowFiltering:=True
            wsJurnal.Protect Password:=PWD, AllowFiltering:=True
            
            ' REFRESH ALL PIVOT + NYALAKAN LAYAR KEMBALI
            ThisWorkbook.RefreshAll
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            
            MsgBox "Data Aset '" & namaAset & "' berhasil dihapus bersih!", _
                   vbInformation, "Sistem Sukses"
            
            ' Tutup UserForm
            Unload Me
        End If
    Else
        MsgBox "Data aset tidak ditemukan di tabel penyusutan!", vbCritical, "Error"
    End If
End Sub

' --- TOMBOL BATAL ---
Private Sub btnBatal_Click()
    Unload Me
End Sub

