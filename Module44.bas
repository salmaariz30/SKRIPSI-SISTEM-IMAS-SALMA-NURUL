Attribute VB_Name = "Module44"
Sub HapusFormPengeluaranUsahaTotal()
    ' ====================================================================
    ' MODUL UTAMA: RESET TOTAL FORMULIR INPUT PENGELUARAN (ANTI-MERGER BUG)
    ' Terkunci Khusus untuk Sheet: PENGELUARAN USAHA_INPUT DATA
    ' Menghapus Bersih Semua Field Inputan Milik
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet input pengeluaran usaha
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("PENGELUARAN USAHA_INPUT DATA")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet bergeser gaib
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'PENGELUARAN USAHA_INPUT DATA' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum data hangus permanen)
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Pengeluaran ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir Pengeluaran") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat tanpa kedip di laptop
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES BUMI HANGUS DATA SECARA ABSOLUT (MENDUKUNG AREA MERGER)
    With LembarForm
        On Error Resume Next
        
        .Range("D12").MergeArea.ClearContents   ' KATEGORI BEBAN
        .Range("G12").MergeArea.ClearContents   ' NO. BUKTI
        .Range("D16").MergeArea.ClearContents   ' TANGGAL BAYAR
        .Range("G16").MergeArea.ClearContents   ' DIBAYAR KEPADA
        .Range("D20").MergeArea.ClearContents   ' NOMINAL PEMBAYARAN
        .Range("G20").MergeArea.ClearContents   ' BAYAR DARI AKUN
        .Range("D24").MergeArea.ClearContents   ' STATUS PENGGUNAAN
        .Range("G24").MergeArea.ClearContents   ' TANGGAL MULAI PENGGUNAAN
        .Range("D28").MergeArea.ClearContents   ' DESKRIPSI
        .Range("G28").MergeArea.ClearContents   ' TANGGAL SELESAI PENGGUNAAN
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor manis   ke hulu sel awal D12)
    LembarForm.Activate
    LembarForm.Range("D12").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Pengeluaran Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

