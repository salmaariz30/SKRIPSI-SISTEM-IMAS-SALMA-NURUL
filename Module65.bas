Attribute VB_Name = "Module65"
Sub TombolSimpanInputPenghasilanLain_Click()
    ' ====================================================================
    ' EKSEKUSI ALL-IN-ONE - VERSI INPUT PENGHASILAN LAIN
    ' Terkunci Khusus untuk Sheet: PENDAPATAN_PENGHASILAN LAIN
    ' ALUR EKSEKUSI: Validasi Mandiri -> Simpan Transaksi -> Jurnal Umum (+ Sapu Bersih Form)
    ' ====================================================================
    Dim wsInput As Worksheet
    
    Set wsInput = Sheets("PENDAPATAN_PENGHASILAN LAIN")
    
    ' --------------------------------------------------------------------
    ' 1. SISTEM PENGAMAN (Validasi Awal Sebelum Seluruh Eksekusi Ambles)
    ' --------------------------------------------------------------------
    With wsInput
        ' Cek fisik kelengkapan cell utama berdasarkan koordinat form terbaru
        If .Range("D11").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("H11").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("F11").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("D15").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("H15").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("D19").MergeArea.Cells(1, 1).Value = "" Then
            
            MsgBox "VALIDASI FORM GAGAL!" & vbCrLf & _
                   "Mohon pastikan data berikut sudah terisi penuh:" & vbCrLf & _
                   "- Tanggal (D11)" & vbCrLf & _
                   "- No. Bukti (H11)" & vbCrLf & _
                   "- Nama Pendapatan (F11)" & vbCrLf & _
                   "- Akun Kas/Bank (D15)" & vbCrLf & _
                   "- Jumlah Nominal (H15)" & vbCrLf & _
                   "- Deskripsi (D19)", _
                   vbExclamation, "Data Belum Lengkap"
            Exit Sub
        End If
        
        ' Validasi tambahan memastikan Nominal di H15 adalah angka dan di atas 0
        If Not IsNumeric(.Range("H15").Value) Or .Range("H15").Value <= 0 Then
            MsgBox "VALIDASI NOMINAL GAGAL!" & vbCrLf & _
                   "Jumlah Nominal di Cell H15 harus berupa angka dan lebih besar dari 0!", _
                   vbCritical, "Kesalahan Nominal"
            Exit Sub
        End If
    End With
    
    ' --------------------------------------------------------------------
    ' 2. ORKESTRASI MODUL BERUNTUN (ANTI-LONCAT)
    ' --------------------------------------------------------------------
    
    ' LANGKAH A: Simpan rekap transaksi penghasilan lain ke database
    Call SimpanTransaksiPenghasilanLain
    
    ' LANGKAH B: Eksekusi penjurnalan otomatis ke Jurnal Umum
    ' (Catatan  : Sub di bawah ini otomatis menyapu bersih isi form setelah sukses)
    Call JurnalUmumInputPenghasilanLain

End Sub

