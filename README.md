# 📘 Aplikasi Rekapitulasi TQQ – PAI UNESA

Aplikasi ini dikembangkan untuk mendukung kegiatan akademik **mata kuliah Pendidikan Agama Islam (PAI)** di Universitas Negeri Surabaya (UNESA), khususnya pada bagian **Ta’limul Qiroatil Qur’an (TQQ)**.  
Dibuat oleh **Puguh Setya Wibowo**, mahasiswa **Sains Data UNESA**, aplikasi ini membantu proses **rekap dan analisis data nilai TQQ** secara otomatis dengan tampilan interaktif berbasis web menggunakan **Streamlit**.

---

## 🎯 Tujuan
Mempermudah dosen, asisten, dan mahasiswa dalam:
- Menggabungkan nilai TQQ dari berbagai kelas atau file Excel.
- Menyusun laporan rekap per kelas dengan cepat.
- Menampilkan statistik nilai dan distribusi kelas secara visual.

---

## ⚙️ Fitur Utama
✅ Membaca file **ZIP** berisi banyak Excel atau satu file **multi-sheet**  
✅ Otomatis mendeteksi kolom seperti `NAMA`, `PRODI`, `TOTAL`, dan `ABJAD`  
✅ Menggabungkan semua data menjadi satu tabel besar dan memecah **per kelas**  
✅ Menampilkan **grafik batang Top-5 kelas per abjad (A–E)**  
✅ Menampilkan **diagram donut** persentase sebaran nilai abjad  
✅ Menyediakan **log duplikat nama** dan peringatan sheet bermasalah  
✅ Ekspor hasil ke **Excel**:
   - Gabungan semua kelas  
   - Per kelas  
   - Ringkasan ABJAD (overview + top-5 per abjad)

---

## 🧠 Pengembangan
Aplikasi ini merupakan hasil pengembangan mandiri oleh mahasiswa **Program Studi Sains Data UNESA** dalam upaya:
- Menerapkan konsep **data engineering** dan **data analytics** ke konteks pendidikan agama.  
- Membangun sistem rekap nilai otomatis berbasis **Python dan Streamlit** yang dapat digunakan tanpa memerlukan instalasi rumit.  
- Mengintegrasikan logika pemeriksaan data (deteksi duplikat, missing data, validasi format) agar hasil rekap menjadi **lebih akurat dan transparan**.  
- Menyediakan tampilan interaktif untuk dosen/asisten TQQ agar mudah memahami distribusi nilai antar kelas.  
- Menjadi proyek awal menuju sistem dashboard nilai **TQQ online terintegrasi** berbasis web kampus.

---

## 🧰 Teknologi
- **Python**  
- **Streamlit**  
- **Pandas & NumPy**  
- **Plotly**  
- **OpenPyXL**  
- **XlsxWriter**

---
By : PuguhSW
