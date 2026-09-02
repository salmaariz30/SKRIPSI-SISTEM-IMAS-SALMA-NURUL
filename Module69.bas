Attribute VB_Name = "Module69"
Sub HapusFormPenghasilanLainTotal()
    ' ====================================================================
    ' MODUL UTAMA: RESET TOTAL FORMULIR PENGHASILAN LAIN (ANTI-MERGER BUG)
    ' Terkunci Khusus untuk Sheet: PENDAPATAN_PENGHASILAN LAIN
    ' Menghapus Bersamaan Seluruh Field Inputan Utama
    ' ====================================================================
    
    Dim LembarForm As Worksheet
    
    ' Mengunci target pembersihan hanya pada sheet pendapatan penghasilan lain
    On Error Resume Next
    Set LembarForm = ThisWorkbook.Sheets("PENDAPATAN_PENGHASILAN LAIN")
    On Error GoTo 0
    
    ' Antisipasi jika nama sheet berubah secara gaib di kemudian hari
    If LembarForm Is Nothing Then
        MsgBox "Error: Sheet 'PENDAPATAN_PENGHASILAN LAIN' tidak ditemukan!", _
               vbCritical, "Sistem Gagal Menemukan Sheet"
        Exit Sub
    End If
    
    ' 1. ASPEK INTERNAL CONTROL (Konfirmasi sebelum hangus permanen)
    If MsgBox("Apakah anda yakin ingin mengosongkan SELURUH data pada Form Penghasilan Lain ini?", _
              vbQuestion + vbYesNo, "Konfirmasi Reset Formulir") = vbNo Then
        Exit Sub
    End If
    
    ' Supaya pergerakan hapusnya mulus secepat kilat tanpa kedip-kedip
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 2. PROSES PEMBERSIHAN DATA SESUAI TARGET KOORDINAT   (SAFE MERGER)
    With LembarForm
        On Error Resume Next
        
        ' Menyapu bersih kolom D
        .Range("D11").MergeArea.ClearContents ' Tanggal
        .Range("D15").MergeArea.ClearContents ' Metode Pembayaran
        .Range("D19").MergeArea.ClearContents ' Deskripsi / Keterangan
        
        ' Menyapu bersih kolom F
        .Range("F11").MergeArea.ClearContents ' Kategori Pendapatan
        .Range("F15").MergeArea.ClearContents ' Akun Kas / Bank
        
        ' Menyapu bersih kolom H
        .Range("H11").MergeArea.ClearContents ' No. Bukti
        .Range("H15").MergeArea.ClearContents ' Jumlah Nominal
        .Range("H19").MergeArea.ClearContents ' Diterima Dari (Koreksi typo sel 519  )
        
        On Error GoTo 0
    End With
    
    ' 3. ASPEK USER EXPERIENCE (Kembalikan kursor ke posisi awal form D11)
    LembarForm.Activate
    LembarForm.Range("D11").MergeArea.Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ' 4. NOTIFIKASI SUKSES
    MsgBox "Seluruh Data Formulir Penghasilan Lain Berhasil Dihapus!", _
           vbInformation, "Sistem Sukses"
           
End Sub

