Attribute VB_Name = "Module6"
Sub ValidasiDanSimpan()
    Dim pw1 As String
    Dim pw2 As String
    Dim namaSheetLogin As String
    Dim passSheet As String

    ' --- PENGATURAN ---
    namaSheetLogin = "LOGIN" ' nama sheet login
    passSheet = "777" 'password cell
    
    ' Mengambil teks dari TextBox
    pw1 = ActiveSheet.TextBox1.Text
    pw2 = ActiveSheet.TextBox2.Text

    ' 1. Cek yang kosong
    If pw1 = "" Or pw2 = "" Then
        MsgBox "Isi Terlebih Dahulu Kedua Kolom!", vbExclamation, "Data Belum Lengkap"
        Exit Sub
    End If

    ' 2. Cek panjang password (min 6 karakter)
    If Len(pw1) < 6 Then
        MsgBox "Kata Sandi Minimal 6 Karakter, biar lebih aman!", vbCritical, "Terlalu Pendek"
        Exit Sub
    End If

    ' 3. Cek kesamaan password
    If pw1 <> pw2 Then
        MsgBox "Kata Sandi Tidak Sama. Coba Cek Lagi!", vbCritical, "Konfirmasi Gagal"
        
        ' Menghapus isi kolom kedua agar user isi ulang
        ActiveSheet.TextBox2.Text = ""
        ActiveSheet.TextBox2.Activate
    Else
        ' --- PROSES PENYIMPANAN SANDI KE UJUNG DUNIA WKWKWK (XFC5) ---
        
        ' Buka kunci sheet agar VBA bisa menulis ke sel XFC5
        ActiveSheet.Unprotect Password:=passSheet
        
        ' Simpan password ke sel XFC5
        Range("XFC5").Value = pw1
        
        ' Kunci kembali sheet agar sel XFC5 tidak bisa diklik/dilihat di Formula Bar
        ActiveSheet.Protect Password:=passSheet
        
        MsgBox "Kata Sandi Berhasil Diperbarui!", vbInformation, "Berhasil"
        
        ' OTOMATIS PINDAH KE HALAMAN LOGIN
        On Error Resume Next ' Menghindari error jika nama sheet salah
        Sheets(namaSheetLogin).Activate
        If Err.Number <> 0 Then
            MsgBox "Sheet '" & namaSheetLogin & "' tidak ditemukan!", vbCritical, "Error Navigasi"
        End If
    End If
End Sub
