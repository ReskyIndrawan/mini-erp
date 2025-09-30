================================================================================
                          Mini ERP - 不良品データ 管理システム
================================================================================

【概要】
Mini ERP adalah aplikasi desktop berbasis Python yang dikembangkan menggunakan
Tkinter untuk mengelola data produk tidak memenuhi standar (不良品データ - Furyo Hin Data).
Aplikasi ini dirancang khusus untuk membantu proses dokumentasi dan pelacakan produk
bermasalah dalam lingkungan manufaktur.

Aplikasi ini menyediakan dua fungsi utama:
1. テンプレート生成 (Template Generation) - Membuat template Excel untuk pencatatan data
2. 不良品データ入力 (Defect Data Entry) - Memasukkan dan mengelola data produk bermasalah

【Persyaratan Sistem】
- Python 3.7 atau lebih tinggi
- Library yang dibutuhkan:
  - openpyxl>=3.1.5
  - tkinter (biasanya sudah termasuk dalam instalasi Python)

【Instalasi】
1. Pastikan Python sudah terinstal di sistem Anda
2. Install library yang dibutuhkan:
   pip install openpyxl
3. Jalankan aplikasi:
   python main.py

================================================================================
                           PANDUAN PENGGUNAAN
================================================================================

【1. Tab テンプレート生成 (Template Generation)】
Tab ini digunakan untuk membuat template Excel baru untuk pencatatan data produk bermasalah.

【Komponen Interface:】
• 作成者 (Creator) - Input field untuk nama pembuat/pengguna
• 保存先ディレクトリ (Save Directory) - Label yang menampilkan direktori penyimpanan yang dipilih
• 参照 (Browse) - Tombol untuk memilih direktori penyimpanan
• Excel作成 (Create Excel) - Tombol untuk membuat file Excel template

【Cara Penggunaan:】
1. Masukkan nama Anda di field 作成者
2. Klik tombol 参照 untuk memilih direktori penyimpanan
3. Setelah direktori dipilih, path akan ditampilkan di label 保存先ディレクトリ
4. Klik tombol Excel作成 untuk membuat template

【Yang Terjadi Saat Tombol Ditekan:】
• Excel作成:
  - Validasi input: Jika nama pembuat kosong atau direktori belum dipilih, akan muncul pesan error
  - Membuat folder dengan format YYYY-MM-不良品データ (contoh: 2024-09-不良品データ)
  - Membuat subfolder 不良発生連絡書発行 di dalamnya
  - Membuat file Excel 不具合品一覧表.xlsx dengan template yang sudah diformat
  - Menambahkan baris dummy sebagai contoh
  - Menampilkan pesan sukses dengan path file yang dibuat

================================================================================

【2. Tab 不良品データ入力 (Defect Data Entry)】
Tab ini digunakan untuk memasukkan, mengelola, dan mencari data produk bermasalah yang
sudah tercatat dalam file Excel.

【Komponen Interface:】

【Bagian File Selection:】
• Excelファイル (Excel File) - Dropdown untuk memilih file Excel dari history
• Excel検索 (Excel Search) - Tombol untuk mencari file Excel baru
• シート (Sheet) - Dropdown untuk memilih sheet dalam file Excel
• 履歴 (History) - Tombol besar untuk menampilkan/mengelola history file yang pernah dibuka

【Bagian Data Entry:】
• 発生月 (Occurrence Month) - Dropdown untuk memilih bulan terjadinya masalah
• 累計 (Total) - Field yang menampilkan total akumulasi (otomatis terisi)
• № (Number) - Field nomor urut (otomatis terisi)
• 発生日 (Occurrence Date) - Field tanggal terjadinya masalah dengan date picker
• 項目 (Item) - Field untuk kategori masalah
• 事象 (Phenomenon) - Field untuk deskripsi masalah utama
• 事象（一次）(Primary Phenomenon) - Field untuk deskripsi masalah primer
• 事象（二次）(Secondary Phenomenon) - Field untuk deskripsi masalah sekunder
• 品番 (Part Number) - Field untuk nomor part produk
• サプライヤー名 (Supplier Name) - Field untuk nama supplier
• 不良発生連絡書発行 (Defect Report Issued) - Checkbox untuk menandai apakah laporan sudah diterbitkan
• 不良発生№ (Defect Number) - Field untuk nomor laporan

【Tombol Aksi:】
• 保存（追加）(Save - Add) - Menyimpan data baru
• 更新（編集）(Update - Edit) - Memperbarui data yang dipilih
• 削除 (Delete) - Menghapus data yang dipilih
• クリア (Clear) - Membersihkan semua field input
• フィルタ (Filter) - Membuka dialog filter

【Bagian Preview:】
• Excelデータプレビュー (Excel Data Preview) - Tabel yang menampilkan data dari file Excel

【Cara Penggunaan:】

【Membuka File Excel:】
1. Klik dropdown Excelファイル untuk memilih file dari history, atau
2. Klik tombol Excel検索 untuk mencari file Excel baru
3. Pilih sheet yang diinginkan dari dropdown シート

【Memasukkan Data Baru:】
1. Isi semua field yang diperlukan:
   - Pilih 発生月 dari dropdown
   - Klik field 発生日 untuk memilih tanggal dari date picker
   - Isi 項目, 事象, 事象（一次）, 事象（二次）
   - Isi 品番 dan サプライヤー名
   - Centang 不良発生連絡書発行 jika sudah diterbitkan
   - Isi 不良発生№ jika ada
2. Klik tombol 保存（追加） untuk menyimpan

【Mengedit Data:】
1. Klik pada baris data di tabel preview
2. Data akan otomatis terisi di form input
3. Lakukan perubahan yang diperlukan
4. Klik tombol 更新（編集） untuk menyimpan perubahan

【Menghapus Data:】
1. Klik pada baris data di tabel preview
2. Data akan otomatis terisi di form input
3. Klik tombol 削除 untuk menghapus data (akan ada konfirmasi)

【Menggunakan Filter:】
1. Klik tombol フィルタ
2. Di dialog filter, Anda bisa:
   - Memilih rentang tanggal (発生日: dari ~ sampai)
   - Melakukan pencarian free text (フリーワード検索)
   - Memilih file PDF/Excel untuk filter tambahan
3. Klik フィルタ適用 untuk menerapkan filter

【Yang Terjadi Saat Tombol Ditekan:】

• Excel検索:
  - Membuka file dialog untuk memilih file Excel
  - File yang dipilih akan ditambahkan ke history
  - Memuat sheet-sheet yang tersedia dalam file

• 履歴 (History):
  - Membuka dialog yang menampilkan 10 file terakhir yang dibuka
  - Dapat menghapus item dari history
  - Dapat memilih file dari history untuk dibuka

• 保存（追加）:
  - Validasi semua field yang diperlukan
  - Menambahkan data baru ke file Excel
  - Memperbarui tampilan preview
  - Membersihkan form input

• 更新（編集）:
  - Memperbarui data yang ada di file Excel
  - Memperbarui tampilan preview
  - Membersihkan form input

• 削除:
  - Menampilkan dialog konfirmasi
  - Menghapus data dari file Excel
  - Memperbarui tampilan preview
  - Membersihkan form input

• クリア:
  - Membersihkan semua field input
  - Menghapus selection di tabel preview

• フィルタ:
  - Membuka dialog filter dengan berbagai opsi
  - Hasil filter akan ditampilkan di tabel preview

================================================================================
                              FITUR TAMBAHAN
================================================================================

【Date Picker】
Saat Anda mengklik field 発生日, akan muncul date picker yang memungkinkan Anda
memilih tanggal dengan mudah menggunakan kalender interaktif.

【File History Management】
Aplikasi secara otomatis menyimpan history file Excel yang pernah dibuka
(maksimal 10 file). History ini tersimpan di:
- Windows: %APPDATA%\defect_data_app\recent_excel_files.json
- Linux/macOS: ~/.config/defect_data_app/recent_excel_files.json

【Auto Numbering】
Field № dan 累計 akan otomatis terisi berdasarkan data yang sudah ada di file Excel.

【Data Validation】
Aplikasi melakukan validasi data sebelum menyimpan:
- Field wajib harus diisi
- Format tanggal harus valid
- File Excel harus ada dan dapat dibaca

================================================================================
                              TROUBLESHOOTING
================================================================================

【File Tidak Bisa Dibuka】
- Pastikan file Excel tidak sedang dibuka di aplikasi lain
- Periksa apakah file memiliki format yang benar (.xlsx)
- Pastikan file tidak rusak

【Error Saat Menyimpan】
- Pastikan Anda memiliki hak akses tulis di direktori tersebut
- Periksa apakah file Excel tidak dalam mode read-only
- Pastikan disk space mencukupi

【Data Tidak Muncul di Preview】
- Pastikan file Excel dan sheet sudah dipilih dengan benar
- Periksa apakah file memiliki data di sheet yang dipilih
- Coba refresh dengan membuka kembali file

================================================================================
                           STRUKTUR FILE EXCEL
================================================================================

File Excel yang dibuat memiliki struktur kolom sebagai berikut:
1. No
2. 発生月 (Bulan kejadian)
3. 累計 (Total akumulasi)
4. 発生日 (Tanggal kejadian)
5. 項目 (Kategori)
6. 事象 (Deskripsi)
7. 事象（一次）(Deskripsi primer)
8. 事象（二次）(Deskripsi sekunder)
9. 品番 (Nomor part)
10. サプライヤー名 (Nama supplier)
11. 不良発生連絡書発行 (Status laporan)
12. 不良発生№ (Nomor laporan)

================================================================================
                                 SUPPORT
================================================================================

Jika Anda mengalami masalah atau memiliki pertanyaan, silakan periksa
troubleshooting section di atas atau hubungi administrator sistem.

================================================================================