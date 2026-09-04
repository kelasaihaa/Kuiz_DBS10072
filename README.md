# Kuiz Interaktif DBS10072 Science

Kuiz aneka pilihan dwibahasa (English + Bahasa Melayu) untuk **DBS10072 Science** — Topik 1.0 (Kuantiti Fizikal & Pengukuran) & Topik 2.0 (Gerakan Linear). Aplikasi satu fail (`index.html`) + backend Google Apps Script (`backend/Code.gs`).

> **Nota penggunaan:** kuiz ini untuk **latihan**. Jawapan berada di sisi klien, jadi ia **tidak sesuai untuk peperiksaan bermarkah rasmi**.

---

## ✨ Ciri-ciri

| Ciri | Keterangan |
|---|---|
| Bank soalan | **30 soalan** (15 setiap topik). Setiap sesi pilih **rawak 20** (10 + 10). |
| Acak | Susunan **soalan & pilihan jawaban** diacak setiap kali. |
| Pemarkahan | 1000 mata asas + 10 mata bonus / baki saat. Timer 60s. |
| Cuba berkali-kali | Dibenarkan; **markah tertinggi** setiap no. matrik kekal di leaderboard. |
| AI Tutor | Melalui backend selamat (API key **tidak** terdedah). |
| Semak Jawapan | Papar jawapan anda vs jawapan betul + penerangan. |
| Papan kekunci | `A–D` / `1–4` untuk pilih, `Enter` untuk hantar. |
| Aksesibiliti | ARIA, fokus keyboard, hormat `prefers-reduced-motion`. |
| Bunyi | Maklum balas betul/salah (Web Audio). |
| Sambung semula | Kemajuan disimpan; boleh sambung jika terputus. |
| Dashboard pensyarah | Statistik kelas + soalan paling kerap salah (dilindungi PIN). |
| Visual 3D | Latar zarah/atom Three.js + kesan kad condong. |
| CSV selamat | UTF-8 BOM + perlindungan *formula injection*. |

---

## 🔐 Keselamatan

- **PIN pensyarah** disimpan hanya sebagai **cincang SHA-256 bergaram** di dalam `index.html` — kata laluan sebenar tidak pernah muncul dalam kod.
- **API key Gemini TIDAK** diletakkan dalam `index.html` (repo ini awam). Ia disimpan di **Script Properties** Apps Script (server), dan klien memanggil backend sebagai proksi.
- Input pelajar dineutralkan sebelum ditulis ke Sheet/CSV untuk elak *formula injection*.

---

## 🚀 Pemasangan Backend (Google Apps Script)

Lakukan **sekali** untuk mengaktifkan AI Tutor, leaderboard berpusat, dan statistik kelas.

1. Buka **Google Sheet** anda → menu **Extensions → Apps Script**.
2. Padam kod lama, **tampal seluruh** `backend/Code.gs`, kemudian **Save**.
3. **Project Settings (gear) → Script Properties → Add script property:**
   - `GEMINI_KEY` = *API key Gemini anda* (WAJIB untuk AI Tutor)
   - `GEMINI_MODEL` = `gemini-2.5-flash` (pilihan; boleh tukar)
4. **Deploy → New deployment → Web app**
   - *Execute as:* **Me**
   - *Who has access:* **Anyone**
5. Salin URL **/exec**, dan tampal ke `index.html` pada pemalar `BACKEND_URL`.
6. Setiap kali kod backend diubah: **Manage deployments → Edit → Version: New version → Deploy**.

> Tanpa langkah di atas, kuiz masih **berfungsi** untuk menyimpan keputusan (serasi dengan `doPost` ringkas sedia ada). Cuma AI Tutor, leaderboard berpusat & statistik memerlukan backend penuh ini.

---

## ⚙️ Tetapan pantas (dalam `index.html`)

```js
const BACKEND_URL = "https://script.google.com/macros/s/XXXX/exec";
const QUIZ_CONFIG = {
    perTopic: 10,        // soalan rawak setiap topik
    timePerQuestion: 60, // saat
    basePoints: 1000,
    bonusPerSecond: 10
};
```

## 📄 Struktur

```
index.html          # Aplikasi kuiz (frontend, satu fail)
backend/Code.gs     # Backend Apps Script (simpan, leaderboard, statistik, AI proxy)
README.md
```
