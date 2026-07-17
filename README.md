# MCU Document Studio

Generator dokumen MCU tiga halaman dengan layout Word, batch scan PDF/gambar melalui Gemini, cetak PDF, export Word, dashboard responsive, dan credit pembuat.

## Environment variables

Set di Railway atau Back4App Containers:

```text
GEMINI_API_KEY=API_KEY_BARU
GEMINI_PRIMARY_MODEL=gemini-2.5-flash
GEMINI_FALLBACK_MODELS=gemini-2.5-flash-lite,gemini-2.0-flash
PORT=8080
```

API key tidak boleh dimasukkan ke repository.

## Back4App Containers

Repository sudah memuat `Dockerfile`. Hubungkan repository GitHub lalu deploy dari branch `main`.

## Railway

Repository sudah memuat `railway.json`. Tambahkan Variables lalu deploy dari branch `main`.

## Menjalankan lokal

```bash
pip install -r requirements.txt
python serve_railway.py
```

Buka `http://127.0.0.1:8080`.

## Credit

© Syaeful Imam Al Kusyaeri  
Blok Selasa, Desa Kertabasuki, Kecamatan Maja, Kabupaten Majalengka, Jawa Barat, Indonesia  
WhatsApp: 0853-2129-6926

## Export Word
Tombol Export Word menghasilkan DOCX editable dari template asli melalui endpoint `/api/export-word`. Untuk penggunaan lokal, jalankan `BUKA_APLIKASI_LOCAL.bat` dan jangan membuka `index.html` langsung.
