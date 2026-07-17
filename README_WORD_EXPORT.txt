EXPORT WORD SUDAH DIPERBAIKI

Perubahan:
- Word dibuat langsung dari template DOCX asli.
- Data formulir dikirim ke backend dalam ukuran kecil, bukan tiga gambar Base64 besar.
- Hasil Word terdiri dari tepat 3 halaman.
- Isi Word tetap editable.
- Mendukung data identitas, laboratorium, EKG, rontgen, diagnosis, tanggal, dokter, klinik, serta cap/tanda tangan.

Back4App/Railway:
- Pastikan build menggunakan Dockerfile terbaru.
- requirements.txt harus memuat python-docx==1.2.0.

Penggunaan lokal:
- Jangan buka index.html langsung karena endpoint Word tidak tersedia pada file://.
- Klik BUKA_APLIKASI_LOCAL.bat.
- Aplikasi terbuka melalui http://127.0.0.1:8080/index.html.
