Attribute VB_Name = "Module84"
Sub TombolSimpanPembelianPersediaan_Click()
    ' ==========================================================
    ' EKSEKUSI ALL-IN-ONE - VERSI DETEKSI MULTI-TABEL (SMART DETECTION)
    ' Terkunci Khusus untuk Sheet: PERSEDIAAN_INPUT BARANG
    ' ==========================================================
    Dim wsInput As Worksheet
    Dim tblJurnal As ListObject, tblPersediaan As ListObject
    Dim JurnalSebelum As Long, JurnalSesudah As Long
    Dim PersediaanSebelum As Long, PersediaanSesudah As Long
    
    Set wsInput = Sheets("PERSEDIAAN_INPUT BARANG")
    
    ' 1. Ambil referensi tabel secara aman
    On Error Resume Next
    Set tblJurnal = Range("TabelJurnalUmum").ListObject
    
    ' Cari TabelPembelianPersediaan secara dinamis di seluruh sheet
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Worksheets
        Set tblPersediaan = Sh.ListObjects("TabelPembelianPersediaan")
        If Not tblPersediaan Is Nothing Then Exit For
    Next Sh
    On Error GoTo 0
    
    ' 2. Hitung jumlah baris awal masing-masing tabel
    JurnalSebelum = IIf(Not tblJurnal Is Nothing, tblJurnal.ListRows.Count, 0)
    PersediaanSebelum = IIf(Not tblPersediaan Is Nothing, tblPersediaan.ListRows.Count, 0)
    
    ' 3. JALANKAN PROSES SIMPAN BAWAAN
    Call SimpanPembelianPersediaan
    Call JurnalUmumInputPembelianPersediaan
    
    ' 4. Hitung jumlah baris akhir setelah makro selesai berjalan
    JurnalSesudah = IIf(Not tblJurnal Is Nothing, tblJurnal.ListRows.Count, 0)
    PersediaanSesudah = IIf(Not tblPersediaan Is Nothing, tblPersediaan.ListRows.Count, 0)
    
    ' ==========================================================
    ' EVALUASI SMART: JIKA KEDUA TABEL TIDAK BERTAMBAH = VALIDASI GAGAL!
    ' ==========================================================
    If (JurnalSesudah <= JurnalSebelum) And (PersediaanSesudah <= PersediaanSebelum) Then
        ' Jika kedua tabel tidak bertambah, artinya memang gagal lolos validasi input.
        Exit Sub
    End If
    
    ' ==========================================================
    ' JIKA LOLOS (SALAH SATU ATAU KEDUA TABEL BERTAMBAH) -> CLEAR FORM
    ' ==========================================================
    On Error Resume Next
    Application.EnableEvents = False
    
    ' Pembersihan khusus sel Merged satu per satu agar bersih total
    wsInput.Range("D12").MergeArea.ClearContents
    wsInput.Range("D16").MergeArea.ClearContents
    wsInput.Range("D20").MergeArea.ClearContents
    wsInput.Range("H12").MergeArea.ClearContents
    wsInput.Range("H16").MergeArea.ClearContents
    
    ' Bagian bawah keranjang belanja
    wsInput.Range("D25:D39,F25:H39").ClearContents
    
    ' Kembalikan kursor ke kolom input tanggal awal
    wsInput.Activate
    wsInput.Range("D12").Select
    
    Application.EnableEvents = True
    On Error GoTo 0
    
    ' NOTIFIKASI JUJUR DAN PASTI MUNCUL
    MsgBox "Berhasil Menyimpan Data Pembelian Persediaan!", vbInformation, "Sukses"
End Sub
