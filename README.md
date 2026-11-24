# Gudang Kendali Dokumen — Puskesmas Ngemplak I

Dashboard kendali dokumen SOP (React + Vite) dengan backend Google Apps Script + Google Sheet.

## Struktur
- `src/App.jsx` — seluruh frontend
- `backend/Code.gs` — Apps Script (salin ke script.google.com, Deploy > New version)
- Tab Sheet: `Dokumen`, `Counter`, `Referensi`

## Jalankan
```
npm install
npm run dev      # lokal
npm run build    # hasil di dist/, upload ke Cloudflare
```
`API_URL` di `src/App.jsx` = URL Web App Apps Script (/exec).
