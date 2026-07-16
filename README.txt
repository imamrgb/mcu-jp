Generator MCU dengan scan batch + Gemini fallback

Penyimpanan API key:
- API key disimpan di file: gemini.config.json
- Anda bisa mengganti key di file itu kapan saja.
- Input API key di halaman web hanya untuk override sementara saat aplikasi berjalan.

Fallback model otomatis:
- Model utama dibaca dari primaryModel.
- Jika model utama gagal, kuota habis, rate limit, atau model tidak tersedia, server akan mencoba fallbackModels secara otomatis.

Urutan bawaan:
1. gemini-2.5-flash
2. gemini-2.5-flash-lite
3. gemini-2.0-flash

Cara pakai:
1. Ekstrak ZIP.
2. Klik BUKA_APLIKASI.bat.
3. Jika perlu, edit gemini.config.json.
4. Pilih file PDF/gambar dan jalankan analisis batch.
