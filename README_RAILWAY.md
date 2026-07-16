# Generator MCU Railway Lite

Versi ringan agar mudah diunduh dan diunggah ke GitHub.

## Fitur
- Layout MCU 3 halaman
- Batch scan PDF/gambar memakai Gemini API dari backend Railway
- API key hanya dari Railway Variables
- Fallback model otomatis
- Cetak PDF
- Export Word DOCX

## Railway Variables
- GEMINI_API_KEY
- GEMINI_PRIMARY_MODEL=gemini-2.5-flash
- GEMINI_FALLBACK_MODELS=gemini-2.5-flash-lite,gemini-2.0-flash

## Deploy
1. Push semua file ke GitHub.
2. Deploy repository di Railway.
3. Tambahkan Variables di atas.
4. Generate domain.
