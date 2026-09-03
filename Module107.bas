Attribute VB_Name = "Module107"
Sub TombolSimpanUtangBiayaUsaha_Click()
    Dim wsInput As Worksheet
    Set wsInput = Sheets("UTANG_INPUT USAHA")
    
    ' --------------------------------------------------------------------
    ' 2. ORKESTRASI MODUL BERUNTUN (PENGAMAN STRICT - ANTI LONCAT)
    ' --------------------------------------------------------------------
    
    ' LANGKAH A: Eksekusi simpan ke database dan tangkap laporannya
    Call SimpanTransaksiUtangUsaha
    
    ' LANGKAH B: Hanya jalan jika Langkah A terbukti SUKSES (True)
    Call JurnalUmumInputUtangBiaya
    
    ' --- POTONGAN KODE TAMBAHAN: SAPU BERSIH FORM UTANG BIAYA USAHA ---
    With wsInput
        .Range("D12").MergeArea.ClearContents   ' Tanggal Transaksi
        .Range("H12").MergeArea.ClearContents   ' Nomor Bukti
        .Range("D16").MergeArea.ClearContents   ' Pilihan Dropdown / Isian D16
        .Range("H16").MergeArea.ClearContents   ' Pilihan Dropdown / Isian H16
        .Range("D20").MergeArea.ClearContents   ' Keterangan / Deskripsi D20
        .Range("D24").MergeArea.ClearContents   ' Nominal D24
        .Range("H24").MergeArea.ClearContents   ' Nominal H24
        
        ' UX Flow: Kembalikan kursor dengan anggun ke sel pertama (D12)
        .Activate
        .Range("D12").MergeArea.Cells(1, 1).Select
    End With

End Sub
