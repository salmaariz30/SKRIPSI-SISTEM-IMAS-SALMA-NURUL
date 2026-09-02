Attribute VB_Name = "Module50"
Sub HapusFormInputTarikTotal()
    ' ====================================================================
    ' MODUL UTAMA: RESET TOTAL FORMULIR INPUT TARIK MODAL (ANTI-MERGER BUG)
    ' Terkunci Khusus untuk Sheet: PENGELUARAN USAHA_INPUT TARIK
    ' Menghapus Bersih Semua Area Merged Cell Inputan Milik
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet input tarik modal
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT TARIK")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet bergeser gaib
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'PENGELUARAN USAHA_INPUT TARIK' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum data hangus permanen)
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Penarikan ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Penarikan") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat tanpa kedip di laptop ASUS
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES BUMI HANGUS DATA SECARA ABSOLUT (MENDUKUNG AREA MERGER)
    With LembarForm
        On Error Resume Next
        
        .Range("D11").MergeArea.ClearContents   ' JUDUL FORM / FIELD TERKAIT D11
        .Range("G11").MergeArea.ClearContents   ' NO. BUKTI (WDR-xxxx)
        .Range("D15").MergeArea.ClearContents   ' TANGGAL TRANSAKSI
        .Range("G15").MergeArea.ClearContents   ' JENIS PENARIKAN DROPDOWN (4 Pilihan)
        .Range("D19").MergeArea.ClearContents   ' NOMINAL PENARIKAN
        .Range("G19").MergeArea.ClearContents   ' NAMA OWNER / PENERIMA ("Umum" / Nama)
        .Range("D23").MergeArea.ClearContents   ' AMBIL DARI AKUN KAS / BANK
        .Range("D27").MergeArea.ClearContents   ' DESKRIPSI / CATATAN TRANSAKSI
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke hulu sel No Bukti G11 biar auto-ready)
    LembarForm.Activate
    LembarForm.Range("G11").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Penarikan Modal Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

