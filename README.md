# Almerass Premium Inventory System

Sistem manajemen inventaris premium yang dirancang khusus untuk operasional PT. Almera Sukses Sejahtera. Sistem ini mengintegrasikan manajemen stok, kalkulasi BOM (*Bill of Materials*), dan pelacakan unit dengan antarmuka modern berbasis *Glassmorphism*.

## Deskripsi Proyek
Sistem ini dikembangkan menggunakan Google Apps Script sebagai backend untuk integrasi langsung dengan Google Sheets sebagai basis datanya, dan menggunakan antarmuka frontend yang dioptimalkan dengan Tailwind CSS.

## Fitur Utama
- **Dashboard Analitik:** Visualisasi tren stok dan data barang.
- **Kalkulasi BOM (Bill of Materials):** Penghitungan otomatis kebutuhan komponen berdasarkan target produksi.
- **Manajemen Stok Real-time:** Pelacakan keluar-masuk barang secara digital.
- **Sistem Keamanan:** Otentikasi pengguna berbasis RBAC (*Role-Based Access Control*).
- **Desain Modern:** Antarmuka dengan tema *Dark Mode* dan aksentuasi *Neon Green* yang ergonomis.

## Struktur Kode
- `code.gs`: Kontroler backend untuk routing, otentikasi, dan API.
- `setup.gs`: Skrip inisialisasi sheet dan database.
- `index.html`: Struktur tampilan frontend (UI).

## Cara Penggunaan
1. Buat Google Sheet baru untuk database.
2. Buka **Extensions > Apps Script**.
3. Salin file `code.gs` dan `setup.gs` ke dalam editor Apps Script.
4. Buat file `index.html` dan tempelkan kode dari `html.html`.
5. Jalankan fungsi `setupSpreadsheet()` untuk menyiapkan struktur database otomatis.
6. Deploy sebagai **Web App** (akses: "Anyone").

## Teknis & Keamanan
Sistem ini dioptimalkan untuk perangkat monitoring industri yang digunakan di lapangan, memastikan data dapat diakses dengan cepat meski pada koneksi yang terbatas.

---
*Dikembangkan oleh Wihardodo | PT. Almera Sukses Sejahtera*
