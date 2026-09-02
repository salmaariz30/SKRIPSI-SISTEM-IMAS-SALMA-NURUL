Attribute VB_Name = "Module63"
Sub TombolSimpanInputPenjualan_Click()
    ' ====================================================================
    ' EKSEKUSI ALL-IN-ONE - VERSI INPUT PENJUALAN MULTI-PENDAPATAN
    ' Terkunci Khusus untuk Sheet: PENDAPATAN_INPUT PENJUALAN
    ' ALUR EKSEKUSI: Validasi -> Simpan Nota -> Simpan Rincian Menu -> Jurnal Umum -> PEMBERSIHAN TOTAL
    ' ====================================================================
    Dim wsInput As Worksheet
    Dim V_MetodePembayaran As String
    Dim V_NamaPelanggan As String
    Dim V_TglJatuhTempo As Variant
    Dim i As Long
    
    Set wsInput = Sheets("PENDAPATAN_INPUT PENJUALAN")
    
    ' --------------------------------------------------------------------
    ' 1. SISTEM PENGAMAN (Validasi Awal Sebelum Seluruh Eksekusi Ambles)
    ' --------------------------------------------------------------------
    With wsInput
        ' Ambil nilai untuk validasi kondisional piutang
        V_MetodePembayaran = Trim(.Range("I11").MergeArea.Cells(1, 1).Value)
        V_NamaPelanggan = Trim(.Range("C15").MergeArea.Cells(1, 1).Value)
        V_TglJatuhTempo = .Range("I19").MergeArea.Cells(1, 1).Value
        
        ' Cek fisik kelengkapan cell utama (Tanggal, No Bukti, Metode Pencatatan, Metode Pembayaran)
        If .Range("C11").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("I15").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("F15").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("I11").MergeArea.Cells(1, 1).Value = "" Or _
           .Range("C19").MergeArea.Cells(1, 1).Value = "" Then
            
            MsgBox "VALIDASI FORM GAGAL!" & vbCrLf & _
                   "Mohon pastikan Tanggal (C11), No. Bukti (I15), Metode Pencatatan (F15), " & vbCrLf & _
                   "Metode Pembayaran (I11), dan Deskripsi (C19) sudah terisi dengan benar.", _
                   vbExclamation, "Data Belum Lengkap"
            Exit Sub
        End If
        
        ' Validasi Pengaman Khusus Akun Bank (Jika Pembayaran via Transfer Bank)
        If LCase(V_MetodePembayaran) = "transfer bank" And .Range("L11").MergeArea.Cells(1, 1).Value = "" Then
            MsgBox "VALIDASI METODE TRANSFER GAGAL!" & vbCrLf & _
                   "Anda memilih 'Transfer Bank', mohon isi Pilihan Bank di Cell L11 terlebih dahulu!", _
                   vbCritical, "Akun Bank Kosong"
            Exit Sub
        End If
        
        ' Validasi Pengaman Khusus Piutang (Jika Pembayaran Sistem Kredit)
        If LCase(V_MetodePembayaran) = "kredit" Then
            If V_NamaPelanggan = "" Then
                MsgBox "VALIDASI SISTEM KREDIT GAGAL!" & vbCrLf & _
                       "Pembayaran 'Kredit' wajib mengisi Nama Pelanggan di Cell C15 untuk Buku Piutang!", _
                       vbCritical, "Identitas Pelanggan Kosong"
                Exit Sub
            End If
            If V_TglJatuhTempo = "" Or Not IsDate(V_TglJatuhTempo) Then
                MsgBox "VALIDASI SISTEM KREDIT GAGAL!" & vbCrLf & _
                       "Pembayaran 'Kredit' wajib menentukan Tanggal Jatuh Tempo yang valid di Cell I19!", _
                       vbCritical, "Jatuh Tempo Belum Diisi"
                Exit Sub
            End If
        End If
    End With
    
    ' --------------------------------------------------------------------
    ' 2. ORKESTRASI MODUL BERUNTUN (ANTI-LONCAT)
    ' --------------------------------------------------------------------
    
    ' LANGKAH A: Simpan rekap nota induk
    Call SimpanTransaksiPendapatanPenjualan
    
    ' LANGKAH B: Parsing & pecah detil menu produk terjual ke database item (dan Buku Piutang jika Kredit)
    Call SimpanRincianProdukTerjual
    
    ' LANGKAH C: Eksekusi penjurnalan otomatis ke Jurnal Umum
    Call JurnalUmumInputPenjualan
    
    ' --------------------------------------------------------------------
    ' 3. SAPU BERSIH TOTAL FORM INPUT KASIR (VERSI AMAN ANTI-MERGER BUG)
    ' --------------------------------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With wsInput
        On Error Resume Next ' Pelindung ekstra saat proses pembersihan
        
        ' --- Pembersihan Cell-Cell Atas (Header & Identitas) ---
        .Range("C11").MergeArea.ClearContents ' Tanggal
        .Range("C15").MergeArea.ClearContents ' Nama Pelanggan
        .Range("C19").MergeArea.ClearContents ' Deskripsi
        .Range("F15").MergeArea.ClearContents ' Metode Pencatatan
        .Range("I11").MergeArea.ClearContents ' Metode Pembayaran
        .Range("I15").MergeArea.ClearContents ' No. Invoice
        .Range("I19").MergeArea.ClearContents ' Tanggal Jatuh Tempo
        .Range("L11").MergeArea.ClearContents ' Nama Bank Pilihan
        .Range("L19").MergeArea.ClearContents ' Jumlah Terbayar Awal
        
        ' --- Pembersihan Blok Area Keranjang "Per Transaksi" Berbaris (Baris 24-38) ---
        For i = 24 To 38
            .Range("D" & i).MergeArea.ClearContents ' Kolom D (Nama Produk)
            .Range("F" & i).ClearContents           ' Kolom F (Qty)
            .Range("J" & i).MergeArea.ClearContents ' Kolom J (Catatan/Pendukung)
        Next i
        
        ' --- Pembersihan Blok Area Keranjang "Rekap Harian" Berbaris (Baris 54-83) ---
        For i = 54 To 83
            .Range("D" & i).MergeArea.ClearContents ' Kolom D (Nama Produk)
            .Range("F" & i).ClearContents           ' Kolom F (Qty)
            .Range("J" & i).MergeArea.ClearContents ' Kolom J (Catatan/Pendukung)
        Next i
        
        On Error GoTo 0
        
        ' Kembalikan kursor fokus ke Cell Tanggal agar Kasir siap ngetik nota baru
        .Activate
        .Range("C11").MergeArea.Cells(1, 1).Select
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

