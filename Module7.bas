Attribute VB_Name = "Module7"
Sub ProsesLogin()
    Dim inputPW As String
    Dim passwordBenar As String
    Dim sheetSandi As Worksheet
    Dim sheetUtama As String
    Dim sheetLogin As Worksheet
    Dim ws As Worksheet
    
    ' --- PENGATURAN ---
    Set sheetLogin = ActiveSheet ' Mengunci objek sheet login yang sedang aktif saat ini
    Set sheetSandi = ThisWorkbook.Sheets("KATA SANDI") ' Nama sheet tempat XFC5 berada
    sheetUtama = "MENU UTAMA NEW" ' Nama sheet tujuan jika berhasil
    ' ------------------

    ' 1. Ambil apa yang diketik user di TextBox Login (Menggunakan OLEObjects agar tidak error)
    On Error Resume Next
    inputPW = sheetLogin.OLEObjects("TextBoxLogin").Object.Text
    On Error GoTo 0
    
    ' 2. Ambil password asli dari sel "Ujung Dunia"
    passwordBenar = sheetSandi.Range("XFC5").Value
    
    ' 3. Validasi input kosong
    If inputPW = "" Then
        MsgBox "Masukkan Password Terlebih Dahulu.", vbExclamation, "Ups!"
        Exit Sub
    End If
    
    ' 4. Cek Cocok atau Tidak
    If inputPW = passwordBenar Then
        
        ' == JURUS BIAR SMOOTH & SUPAYA TIDAK BERKEDIP KASAR ==
        Application.ScreenUpdating = False
        
        MsgBox "Akses Diterima! Selamat Datang.", vbInformation, "Login Berhasil!"
        
        ' LOKASI PERUBAHAN UTAMA:
        ' 1. Buka semua sheet data lainnya terlebih dahulu di latar belakang
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> sheetLogin.Name And ws.Name <> sheetSandi.Name Then
                ws.Visible = xlSheetVisible
            End If
        Next ws
        
        ' 2. TERAKHIR, baru kunci fokus pandangan dan mendarat di MENU UTAMA NEW
        ThisWorkbook.Sheets(sheetUtama).Activate
        
        ' 3. Sembunyikan sheet "Login" dan "KATA SANDI" secara total (VeryHidden)
        sheetLogin.Visible = xlSheetVeryHidden
        sheetSandi.Visible = xlSheetVeryHidden
        
        ' 4. Kosongkan textbox login agar bersih saat sistem kembali terkunci nanti
        sheetLogin.OLEObjects("TextBoxLogin").Object.Text = ""
        
        ' == NYALAKAN KEMBALI LAYAR SETELAH SEMUA PROSES BERES ==
        Application.ScreenUpdating = True
        
    Else
        ' Jika Salah
        MsgBox "Kata sandi salah! Coba Lagi.", vbCritical, "Akses Ditolak"
        
        ' 1. Kosongkan textbox agar user bisa coba lagi
        sheetLogin.OLEObjects("TextBoxLogin").Object.Text = ""
        
        ' 2. INI CARA AMAN BUAT AKTIFIN TEXTBOX-NYA LAGI TANPA ERROR:
        sheetLogin.OLEObjects("TextBoxLogin").Select
    End If
End Sub
