# app.py
# Streamlit: (ZIP banyak file) ATAU (1 Excel multi-sheet) → Gabung → Per Kelas + LOG tunggal
# Termasuk LOG "NAMA_DUPLIKAT" (total duplikat, sebaran kelas, dan ringkasan ABJAD per kelas)
import streamlit as st
import pandas as pd
import zipfile
import io
import re
import unicodedata
import time
from pathlib import Path
from datetime import datetime

# Tambahan import
import numpy as np
import plotly.express as px

ALLOWED_EXT = {".xlsx", ".xlsm", ".xls"}

# =========================
# Utilities
# =========================
def norm_name(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^0-9a-zA-Z ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def make_unique_cols(cols):
    seen = {}
    new_cols = []
    for c in cols:
        c = str(c)
        if c in seen:
            seen[c] += 1
            new_cols.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            new_cols.append(c)
    return new_cols

def list_duplicated_cols(cols):
    idx = pd.Index(cols)
    return list(idx[idx.duplicated()])

@st.cache_data(show_spinner=False)
def extract_zip_to_memory(zip_bytes: bytes):
    zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    files = {}
    for info in zf.infolist():
        if info.is_dir():
            continue
        files[info.filename] = zf.read(info)
    return files

def load_master(master_bytes: bytes, master_sheet: str | None = None):
    try:
        if master_sheet:
            dfm = pd.read_excel(io.BytesIO(master_bytes), sheet_name=master_sheet, engine='openpyxl')
        else:
            xls = pd.ExcelFile(io.BytesIO(master_bytes))
            first = xls.sheet_names[0]
            dfm = pd.read_excel(io.BytesIO(master_bytes), sheet_name=first, engine='openpyxl')
    except Exception:
        dfm = pd.read_csv(io.BytesIO(master_bytes))
    dfm.columns = [c.strip().upper() for c in dfm.columns]
    name_col = None
    class_col = None
    for c in dfm.columns:
        if name_col is None and re.search(r"\bNAMA\b", c, re.I):
            name_col = c
        if class_col is None and re.search(r"\bKELAS\b", c, re.I):
            class_col = c
    if not name_col or not class_col:
        raise ValueError("Master harus punya kolom 'NAMA' dan 'KELAS'.")
    dfm['_NORM_NAMA'] = dfm[name_col].astype(str).map(norm_name)
    return dfm, name_col, class_col

def read_possible_sheet(bytes_data: bytes, filename: str, sheet_pattern: str | None, read_all_when_no_pattern: bool = False):
    frames = {}
    try:
        xls = pd.ExcelFile(io.BytesIO(bytes_data))
    except Exception as e:
        return {"__ERROR__": pd.DataFrame({"FILE":[filename], "ERROR":[str(e)]})}
    sheets = xls.sheet_names
    if sheet_pattern:
        regex = re.compile(sheet_pattern, flags=re.I)
        target_sheets = [s for s in sheets if regex.search(s)]
        if not target_sheets:
            fallback = [s for s in sheets if re.search(r"(rekap|nilai.?akhir|sheet\s*1|lembar\s*1)", s, re.I)]
            target_sheets = fallback or ([sheets[0]] if not read_all_when_no_pattern else sheets)
    else:
        target_sheets = sheets if read_all_when_no_pattern else [sheets[0]]
    for sname in target_sheets:
        try:
            df = pd.read_excel(io.BytesIO(bytes_data), sheet_name=sname, engine='openpyxl')
        except Exception as e:
            df = pd.DataFrame({"FILE":[filename], "SHEET":[sname], "ERROR":[str(e)]})
        frames[sname] = df
    return frames

def read_from_zip(files_dict: dict, sheet_pattern: str | None):
    for fname, data in files_dict.items():
        if Path(fname).suffix.lower() not in ALLOWED_EXT:
            continue
        frames = read_possible_sheet(data, fname, sheet_pattern, read_all_when_no_pattern=False)
        yield fname, data, frames

def read_from_single_excel(excel_bytes: bytes, filename: str, sheet_pattern: str | None):
    frames = read_possible_sheet(excel_bytes, filename, sheet_pattern, read_all_when_no_pattern=True)
    return [(filename, excel_bytes, frames)]

def standardize_cols(df: pd.DataFrame):
    df = df.copy()
    cols = [str(c).strip().upper() if c is not None else "" for c in df.columns]
    cols = [c if c != "" else "UNNAMED" for c in cols]
    df.columns = make_unique_cols(cols)
    aliases = {
        "NAMA": ["NAMA", "NAMA SISWA", "NAMA LENGKAP"],
        "PRODI": ["PRODI", "PROGRAM STUDI"],
        "PRESENSI": ["PRESENSI", "ABSENSI", "KEHADIRAN"],
        "BACAAN": ["BACAAN"],
        "HAFALAN": ["HAFALAN"],
        "EVALUASI": ["EVALUASI", "NILAI"],
        "TOTAL": ["TOTAL", "JUMLAH", "SKOR"],
        "ABJAD": ["ABJAD", "GRADE", "HURUF"],
        "NIM": ["NIM"]  # agar NIM ikut ke urutan atas bila ada
    }
    canonical = {}
    for key, candidates in aliases.items():
        for c in candidates:
            if c in df.columns:
                canonical[key] = c
                break
    ordered_keys = ["NAMA","PRODI","NIM","PRESENSI","BACAAN","HAFALAN","EVALUASI","TOTAL","ABJAD"]
    ordered = [canonical[k] for k in ordered_keys if k in canonical]
    rest = [c for c in df.columns if c not in ordered]
    return df[ordered + rest]

def to_excel_download(df_dict: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sname, frame in df_dict.items():
            safe = re.sub(r"[:\\/?*\[\]]", "-", sname)[:31] or "SHEET"
            frame.to_excel(writer, sheet_name=safe, index=False)
    return output.getvalue()

def peek_excel_sheets(excel_bytes: bytes) -> list[str]:
    try:
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))
        return xls.sheet_names
    except Exception as e:
        return [f"__ERROR__: {e}"]

# ====== Tambahan utilitas kecil untuk ABJAD ======
def _coerce_str(x):
    return None if pd.isna(x) else str(x).strip()

def _safe_upper(x):
    return None if x is None else str(x).strip().upper()

def compute_abjad_overview_and_top5(merged: pd.DataFrame, kelas_col: str, abjad_col: str):
    """Kembalikan:
       - df_overall: total per abjad
       - top5_map: dict {abjad -> DataFrame(kelas, abjad, jumlah) top-5}
    """
    df = merged[[kelas_col, abjad_col]].copy()
    df[kelas_col] = df[kelas_col].map(_coerce_str).fillna("TANPA_KELAS")
    df[abjad_col] = df[abjad_col].map(_safe_upper)

    grp = (
        df.groupby([kelas_col, abjad_col], dropna=False)
          .size().reset_index(name="JUMLAH")
    )
    grp = grp[grp[abjad_col].notna() & (grp[abjad_col] != "")]

    df_overall = (
        grp.groupby(abjad_col)["JUMLAH"]
           .sum().reset_index()
           .sort_values("JUMLAH", ascending=False)
           .reset_index(drop=True)
    )

    top5_map = {}
    for a in df_overall[abjad_col].tolist():
        sub = grp[grp[abjad_col] == a].sort_values("JUMLAH", ascending=False).head(5)
        sub = sub[[kelas_col, abjad_col, "JUMLAH"]].reset_index(drop=True)
        top5_map[a] = sub

    return df_overall, top5_map

# ====== IFS: konversi skor numerik → ABJAD (NaN → E) ======
def compute_abjad_from_score(df: pd.DataFrame, score_col: str) -> pd.Series:
    """
    Pemetaan persis rumus IFS:
    H>=86:'A'; H>=81:'A-'; H>=76:'B+'; H>=71:'B'; H>=66:'B-';
    H>=61:'C+'; H>=56:'C'; H>=41:'D'; H>=0:'E'; lainnya/NaN → 'E'
    """
    s = pd.to_numeric(df[score_col], errors="coerce")
    conds = [
        s >= 86, s >= 81, s >= 76, s >= 71, s >= 66,
        s >= 61, s >= 56, s >= 41, s >= 0
    ]
    choices = ["A","A-","B+","B","B-","C+","C","D","E"]
    return np.select(conds, choices, default="E")

# =========================
# UI
# =========================
st.set_page_config(page_title="Rekap Excel → Gabung & Per Kelas (ZIP opsional)", layout="wide")
st.title("📊 Rekap Excel: Gabung → Per Kelas (ZIP opsional, Excel multi-sheet didukung)")

with st.sidebar:
    mode = st.radio("Pilih sumber data", ["ZIP berisi banyak file Excel", "Satu file Excel (banyak sheet)"])
    st.markdown("### Pengaturan Baca Sheet")
    sheet_pattern = st.text_input(
        "Filter nama sheet (regex, opsional)",
        help="Contoh: rekap|nilai.?akhir  \nZIP: kosong → ambil sheet pertama\nExcel multi: kosong → ambil SEMUA sheet"
    )
    user_name = st.text_input("Nama pengguna (opsional) untuk dicatat di LOG", value="")
    st.caption("ZIP boleh berisi subfolder. Excel multi-sheet akan membaca semua sheet jika filter kosong.")

col1, col2 = st.columns(2)
zip_file = None
excel_multi_file = None
with col1:
    if mode == "ZIP berisi banyak file Excel":
        zip_file = st.file_uploader("📦 Upload ZIP berisi file Excel", type=["zip"])
with col2:
    if mode == "Satu file Excel (banyak sheet)":
        excel_multi_file = st.file_uploader("📘 Upload 1 file Excel multi-sheet", type=["xlsx","xlsm","xls"])
        if excel_multi_file is not None:
            sheets_preview = peek_excel_sheets(excel_multi_file.getvalue())
            st.info(
                f"Sheet terbaca di **{excel_multi_file.name}**: "
                f"{', '.join(map(str, sheets_preview[:20]))}" +
                (" ..." if len(sheets_preview) > 20 else "")
            )

master_file = st.file_uploader("🧭 Upload Master (Excel/CSV) berisi kolom NAMA & KELAS", type=["xlsx","xlsm","xls","csv"])
master_sheet = st.text_input("Nama sheet master (opsional)", placeholder="Kosongkan untuk pakai sheet pertama")

ok_to_run = (zip_file is not None or excel_multi_file is not None) and (master_file is not None)

if ok_to_run:
    st.success("File terunggah. Siap diproses.")
    if st.button("▶️ Proses sekarang", type="primary"):
        try:
            t0 = time.time()
            ts_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # ====== LOG containers ======
            log_baca_file_rows = []
            daftar_sheet_rows = []
            cek_status_rows   = []

            # Muat master
            df_master, master_name_col, master_class_col = load_master(master_file.getvalue(), master_sheet or None)
            st.success(f"Master dimuat. Kolom Nama: **{master_name_col}**, Kolom Kelas: **{master_class_col}**")

            # Siapkan iterator sumber data
            if mode == "ZIP berisi banyak file Excel":
                with st.spinner("Ekstraksi ZIP..."):
                    files = extract_zip_to_memory(zip_file.getvalue())
                total_items = sum(1 for k in files.keys() if Path(k).suffix.lower() in ALLOWED_EXT)
                st.write(f"Ditemukan **{total_items}** file Excel di dalam ZIP.")
                src_iter = read_from_zip(files, sheet_pattern.strip() or None)
            else:
                if excel_multi_file is None:
                    st.error("File Excel multi-sheet belum dipilih.")
                    st.stop()
                xls = pd.ExcelFile(io.BytesIO(excel_multi_file.getvalue()))
                st.write(f"📄 File **{excel_multi_file.name}** punya **{len(xls.sheet_names)}** sheet.")
                if not sheet_pattern:
                    st.write("Mode multi-sheet tanpa regex → semua sheet akan dibaca.")
                src_iter_list = read_from_single_excel(excel_multi_file.getvalue(), excel_multi_file.name, sheet_pattern.strip() or None)
                src_iter = iter(src_iter_list)
                total_items = 1

            rows = []
            problem_logs = []
            total_sheets_ok = 0
            processed = 0
            prog = st.progress(0.0, text="Memulai...")

            # Loop sumber
            for fname, data, frames in src_iter:
                processed += 1
                file_has_ok = False

                # catat sheet yang dipilih
                for sname in list(frames.keys()):
                    if sname != "__ERROR__":
                        daftar_sheet_rows.append({"FILE": fname, "SHEET": sname})

                if "__ERROR__" in frames:
                    df_err = frames["__ERROR__"]
                    problem_logs.append(df_err.assign(FILE=fname, SHEET="__LOAD__"))
                    log_baca_file_rows.append({
                        "FILE": fname, "SHEET": "__LOAD__", "STATUS": "ERROR",
                        "ERROR": df_err.iloc[0]["ERROR"], "JUMLAH_BARIS": 0
                    })
                    cek_status_rows.append({"FILE": fname, "STATUS": "ERROR"})
                else:
                    for sname, df in frames.items():
                        if df is None or df.empty:
                            problem_logs.append(pd.DataFrame({"FILE":[fname], "SHEET":[sname], "ERROR":["Sheet kosong / tidak terbaca"]}))
                            log_baca_file_rows.append({
                                "FILE": fname, "SHEET": sname, "STATUS": "KOSONG",
                                "ERROR": "Sheet kosong / tidak terbaca", "JUMLAH_BARIS": 0
                            })
                            continue

                        # Standarisasi & hardening kolom
                        df = standardize_cols(df)
                        dup_cols = list_duplicated_cols(df.columns)
                        if dup_cols:
                            log_baca_file_rows.append({
                                "FILE": fname, "SHEET": sname, "STATUS": "PERINGATAN_DUP_KOLOM",
                                "ERROR": f"Kolom duplikat: {', '.join(map(str, dup_cols))}",
                                "JUMLAH_BARIS": int(len(df))
                            })
                            df = df.loc[:, ~pd.Index(df.columns).duplicated()]

                        df["_FILE_ASAL"]  = fname
                        df["_SHEET_ASAL"] = sname

                        rows.append(df)
                        file_has_ok = True
                        total_sheets_ok += 1
                        log_baca_file_rows.append({
                            "FILE": fname, "SHEET": sname, "STATUS": "OK",
                            "ERROR": "", "JUMLAH_BARIS": int(len(df))
                        })

                    cek_status_rows.append({"FILE": fname, "STATUS": "OK" if file_has_ok else "TIDAK ADA"})

                prog.progress(min(processed / max(total_items, 1), 1.0), text=f"Membaca: {fname}")

            # Gabung
            df_all = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
            df_log_err = pd.concat(problem_logs, ignore_index=True) if problem_logs else pd.DataFrame()

            if df_all.empty:
                st.warning("Tidak ada data terbaca dari sumber dengan konfigurasi saat ini.")
                log_excel_bytes = to_excel_download({
                    "RINGKASAN": pd.DataFrame([{
                        "TIMESTAMP_START": ts_start,
                        "TIMESTAMP_END": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "USER": user_name,
                        "SUMBER": mode,
                        "TOTAL_ITEM_SUMBER": total_items,
                        "TOTAL_SHEET_OK": total_sheets_ok,
                        "TOTAL_BARIS": 0,
                        "KELAS_UNIK": 0,
                        "DURASI_DETIK": round(time.time() - t0, 3),
                        "STATUS": "GAGAL (DATA KOSONG)"
                    }]),
                    "CEK_STATUS_SHEET": pd.DataFrame(cek_status_rows),
                    "DAFTAR_SHEET": pd.DataFrame(daftar_sheet_rows),
                    "LOG_BACA_FILE": pd.DataFrame(log_baca_file_rows),
                    "LOG_MASALAH": df_log_err if not df_log_err.empty else pd.DataFrame()
                })
                st.download_button(
                    "⬇️ Unduh LOG_REKAP.xlsx",
                    data=log_excel_bytes,
                    file_name=f"LOG_REKAP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.stop()

            # =========================
            # Join ke master & pecah per kelas
            # =========================
            cols_all = df_all.columns.tolist()
            name_col_candidates = [c for c in df_all.columns if re.fullmatch(r"NAMA", c, re.I)]
            nama_col = name_col_candidates[0] if name_col_candidates else cols_all[0]
            df_all["_NORM_NAMA"] = df_all[nama_col].astype(str).map(norm_name)

            merged = df_all.merge(
                df_master[["_NORM_NAMA", master_name_col, master_class_col]].drop_duplicates("_NORM_NAMA"),
                on="_NORM_NAMA", how="left", suffixes=("", "_MASTER")
            )

            kelas_counts = merged[master_class_col].fillna("TANPA_KELAS").value_counts().sort_index()

            st.subheader("Ringkasan")
            c1, c2, c3 = st.columns(3)
            c1.metric("Total baris gabungan", f"{len(merged):,}")
            c2.metric("Jumlah item sumber", f"{(1 if mode!='ZIP berisi banyak file Excel' else total_items):,}")
            c3.metric("Kelas unik", f"{merged[master_class_col].nunique():,}")

            st.write("Distribusi per kelas:")
            st.dataframe(kelas_counts.rename("JUMLAH").to_frame())

            # =========================
            # Buat kolom ABJAD dari skor numerik (NaN → E) + Donut persentase
            # =========================
            SCORE_COL = "TOTAL"  # ganti jika kolom skor-mu bukan 'TOTAL'
            if SCORE_COL not in merged.columns:
                for alt in ["NILAI", "SKOR", "JUMLAH", "TOTAL"]:
                    if alt in merged.columns:
                        SCORE_COL = alt
                        st.warning(f"Kolom skor 'TOTAL' tidak ditemukan. Pakai '{SCORE_COL}'.")
                        break
                else:
                    st.error("Tidak menemukan kolom skor numerik (mis. TOTAL/NILAI/SKOR). Donut & grafik ABJAD dilewati.")
                    SCORE_COL = None

            if SCORE_COL is not None:
                merged["ABJAD"] = compute_abjad_from_score(merged, SCORE_COL)

                # Ringkasan total & persentase A–E (urut tetap)
                order = ["A","A-","B+","B","B-","C+","C","D","E"]
                counts = merged["ABJAD"].value_counts().reindex(order, fill_value=0)
                total_all = counts.sum()
                perc = (counts / total_all * 100).round(2) if total_all > 0 else counts.astype(float)

                summary_abjad = pd.DataFrame({
                    "ABJAD": order,
                    "TOTAL": counts.values,
                    "PERSENTASE_%": perc.values
                })

                st.subheader("Rekap Nilai Abjad (Donut)")
                st.dataframe(summary_abjad, use_container_width=True)

                fig = px.pie(
                    summary_abjad, names="ABJAD", values="TOTAL",
                    hole=0.6, title="Persentase Nilai Abjad (A–E)"
                )
                fig.update_traces(textposition="inside", texttemplate="%{label}<br>%{percent:.1%}")
                st.plotly_chart(fig, use_container_width=True)

            # =========================
            # DETEKSI NAMA DUPLIKAT → LOG "NAMA_DUPLIKAT"
            # =========================
            m = merged.copy()
            m[master_class_col] = m[master_class_col].fillna("TANPA_KELAS")

            dupe_total = (
                m.groupby("_NORM_NAMA", dropna=False)
                 .size()
                 .reset_index(name="TOTAL_DUPLIKAT")
            )
            dupe_total = dupe_total[dupe_total["TOTAL_DUPLIKAT"] > 1]

            if not dupe_total.empty:
                nama_var = (
                    m.groupby("_NORM_NAMA")["NAMA"]
                     .apply(lambda s: "; ".join(sorted({str(x).strip() for x in s if str(x).strip()})))
                     .reset_index(name="NAMA_VARIAN")
                )
                kelas_count = (
                    m.groupby(["_NORM_NAMA", master_class_col])
                     .size().reset_index(name="JML")
                )
                kelas_ringkas = (
                    kelas_count
                    .assign(PAIR=lambda df: df[master_class_col].astype(str) + "(" + df["JML"].astype(str) + ")")
                    .groupby("_NORM_NAMA")["PAIR"]
                    .apply(lambda s: "; ".join(sorted(s)))
                    .reset_index(name="KELAS_RINGKAS")
                )
                if "ABJAD" in m.columns:
                    abjad_count = (
                        m.groupby(["_NORM_NAMA", master_class_col, "ABJAD"])
                         .size().reset_index(name="JML")
                    )
                    abjad_per_kelas = (
                        abjad_count
                        .assign(PAIR=lambda df: df["ABJAD"].astype(str) + ":" + df["JML"].astype(str))
                        .groupby(["_NORM_NAMA", master_class_col])["PAIR"]
                        .apply(lambda s: "{" + ",".join(sorted(s)) + "}")
                        .reset_index()
                    )
                    abjad_ringkas = (
                        abjad_per_kelas
                        .assign(ITEM=lambda df: df[master_class_col].astype(str) + ":" + df["PAIR"])
                        .groupby("_NORM_NAMA")["ITEM"]
                        .apply(lambda s: "; ".join(sorted(s)))
                        .reset_index(name="ABJAD_PER_KELAS")
                    )
                else:
                    abjad_ringkas = pd.DataFrame({"_NORM_NAMA": dupe_total["_NORM_NAMA"], "ABJAD_PER_KELAS": ""})

                df_dupe_names = (
                    dupe_total
                    .merge(nama_var, on="_NORM_NAMA", how="left")
                    .merge(kelas_ringkas, on="_NORM_NAMA", how="left")
                    .merge(abjad_ringkas, on="_NORM_NAMA", how="left")
                    .sort_values("TOTAL_DUPLIKAT", ascending=False)
                    .reset_index(drop=True)
                )
            else:
                df_dupe_names = pd.DataFrame(columns=["_NORM_NAMA","TOTAL_DUPLIKAT","NAMA_VARIAN","KELAS_RINGKAS","ABJAD_PER_KELAS"])

            with st.expander("🔁 Nama Duplikat (berdasarkan normalisasi)"):
                if df_dupe_names.empty:
                    st.write("Tidak ada nama duplikat.")
                else:
                    st.dataframe(df_dupe_names.head(200))

            # =========================
            # 🔥 Grafik Nilai Abjad & Top-5 Kelas per Abjad
            # =========================
            st.subheader("Grafik Nilai Abjad & Top-5 Kelas per Abjad")
            if "ABJAD" not in merged.columns:
                st.info("Kolom **ABJAD** tidak ditemukan di data gabungan. Grafik abjad tidak dapat dibuat.")
                abjad_overview = pd.DataFrame(columns=["ABJAD", "JUMLAH"])
                top5_map = {}
            else:
                abjad_col = "ABJAD"
                # Hitung overview & top-5
                abjad_overview, top5_map = compute_abjad_overview_and_top5(merged, master_class_col, abjad_col)

                # Grafik overview total per abjad
                if not abjad_overview.empty:
                    st.write("**Total per Abjad (semua kelas)**")
                    st.bar_chart(abjad_overview.set_index(abjad_col)["JUMLAH"])
                    st.dataframe(abjad_overview.rename(columns={abjad_col: "ABJAD"}))

                # Top-5 per abjad (grafik batang)
                if top5_map:
                    tabs = st.tabs([f"Top-5 {a}" for a in top5_map.keys()])
                    for (a, df_top5), tab in zip(top5_map.items(), tabs):
                        with tab:
                            if df_top5.empty:
                                st.write(f"Tidak ada data untuk abjad {a}.")
                            else:
                                st.write(f"**Top-5 kelas dengan nilai abjad {a} terbanyak**")
                                chart_df = df_top5.set_index(master_class_col)["JUMLAH"]
                                st.bar_chart(chart_df)
                                st.dataframe(df_top5.rename(columns={master_class_col: "KELAS"}))

            # =========================
            # Mismatch
            # =========================
            master_only = df_master.loc[~df_master["_NORM_NAMA"].isin(merged["_NORM_NAMA"]), [master_name_col, master_class_col]].copy()
            rekap_only = merged.loc[merged[master_class_col].isna(), [nama_col, "_FILE_ASAL", "_SHEET_ASAL"]].copy()
            master_only.rename(columns={master_name_col:"NAMA", master_class_col:"KELAS"}, inplace=True)
            rekap_only.rename(columns={nama_col:"NAMA"}, inplace=True)

            # =========================
            # Unduhan gabungan & perkelas
            # =========================
            gabungan_download = to_excel_download({
                "GABUNGAN": merged,
                "LOG_MASALAH": df_log_err if not df_log_err.empty else pd.DataFrame()
            })

            # --- Per Kelas: hanya kolom yang diminta ---
            desired_cols = ["NAMA", "PRODI", "NIM", "PRESENSI", "BACAAN", "HAFALAN", "EVALUASI", "TOTAL", "ABJAD"]
            drop_meta = ["_NORM_NAMA", "_FILE_ASAL", "_SHEET_ASAL", "NAMA_MASTER", "KELAS",
                         f"{master_name_col}_MASTER", f"{master_class_col}_MASTER"]

            perkelas_dict = {}
            for kelas, sub in merged.groupby(merged[master_class_col].fillna("TANPA_KELAS")):
                sub_clean = sub.drop(columns=[c for c in drop_meta if c in sub.columns], errors="ignore").copy()
                keep = [c for c in desired_cols if c in sub_clean.columns]
                sub_final = sub_clean.reindex(columns=keep)
                perkelas_dict[str(kelas)] = sub_final

            # Sheet mismatch (disederhanakan juga - tanpa meta)
            if not master_only.empty:
                mo_keep = [c for c in ["NAMA", "KELAS"] if c in master_only.columns]
                perkelas_dict["MASTER_TDK_KETEMU"] = master_only[mo_keep] if mo_keep else master_only

            if not rekap_only.empty:
                rekap_only_clean = rekap_only.drop(columns=["_FILE_ASAL", "_SHEET_ASAL"], errors="ignore")
                ro_keep = [c for c in ["NAMA"] if c in rekap_only_clean.columns]
                perkelas_dict["REKAP_TDK_ADA_MASTER"] = (
                    rekap_only_clean[ro_keep] if ro_keep else rekap_only_clean
                )

            perkelas_download = to_excel_download(perkelas_dict)

            # ===== Unduhan ringkasan ABJAD (OVERVIEW + TOP5 per abjad) =====
            if "ABJAD" in merged.columns:
                abjad_book = {}
                abjad_book["OVERVIEW_ABJAD"] = (
                    abjad_overview.rename(columns={"JUMLAH": "TOTAL"})
                                   .rename(columns={"ABJAD":"ABJAD"} if "ABJAD" in abjad_overview.columns else {})
                )
                for a, df_top5 in top5_map.items():
                    sh = f"TOP5_{a}"[:31]
                    abjad_book[sh] = df_top5.rename(
                        columns={master_class_col: "KELAS", "JUMLAH": f"TOTAL_{a}", "ABJAD":"ABJAD"}
                    )
                abjad_download = to_excel_download(abjad_book)
            else:
                abjad_download = None

            st.subheader("Preview Gabungan (atas 200 baris)")
            st.dataframe(merged.head(200))

            st.subheader("Unduhan")
            st.download_button(
                "⬇️ Unduh Excel Gabungan (dgn LOG_MASALAH)",
                data=gabungan_download,
                file_name=f"gabungan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "⬇️ Unduh Excel Per Kelas (+ mismatch)",
                data=perkelas_download,
                file_name=f"rekap_perkelas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            if abjad_download is not None:
                st.download_button(
                    "⬇️ Unduh Ringkasan ABJAD (OVERVIEW + TOP-5 per Abjad)",
                    data=abjad_download,
                    file_name=f"TOP5_PER_ABJAD_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            # =========================
            # Susun LOG tunggal & tampilkan
            # =========================
            ts_end = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            summary_info = {
                "TIMESTAMP_START": ts_start,
                "TIMESTAMP_END": ts_end,
                "USER": user_name,
                "SUMBER": mode,
                "TOTAL_ITEM_SUMBER": (1 if mode != "ZIP berisi banyak file Excel" else total_items),
                "TOTAL_SHEET_OK": total_sheets_ok,
                "TOTAL_BARIS": int(len(merged)),
                "KELAS_UNIK": int(merged[master_class_col].nunique()),
                "DURASI_DETIK": round(time.time() - t0, 3),
                "STATUS": "SUKSES"
            }

            log_excel_bytes = to_excel_download({
                "RINGKASAN": pd.DataFrame([summary_info]),
                "CEK_STATUS_SHEET": pd.DataFrame(cek_status_rows),
                "DAFTAR_SHEET": pd.DataFrame(daftar_sheet_rows),
                "LOG_BACA_FILE": pd.DataFrame(log_baca_file_rows),
                "DISTRIBUSI_KELAS": kelas_counts.rename("JUMLAH").to_frame().reset_index(names=["KELAS"]),
                "MASTER_TDK_KETEMU": master_only if not master_only.empty else pd.DataFrame(columns=["NAMA","KELAS"]),
                "REKAP_TDK_ADA_MASTER": rekap_only if not rekap_only.empty else pd.DataFrame(columns=["NAMA","_FILE_ASAL","_SHEET_ASAL"]),
                "NAMA_DUPLIKAT": df_dupe_names,   # <= ditambahkan ke LOG
            })

            st.subheader("📘 LOG Rekap (preview)")
            with st.expander("RINGKASAN"):
                st.dataframe(pd.DataFrame([summary_info]))
            with st.expander("CEK_STATUS_SHEET"):
                st.dataframe(pd.DataFrame(cek_status_rows))
            with st.expander("DAFTAR_SHEET"):
                st.dataframe(pd.DataFrame(daftar_sheet_rows))
            with st.expander("LOG_BACA_FILE"):
                st.dataframe(pd.DataFrame(log_baca_file_rows))
            with st.expander("DISTRIBUSI_KELAS"):
                st.dataframe(kelas_counts.rename("JUMLAH").to_frame())
            if not df_log_err.empty:
                with st.expander("LOG_MASALAH (error/exception saat baca)"):
                    st.dataframe(df_log_err)

            st.download_button(
                "⬇️ Unduh LOG_REKAP.xlsx",
                data=log_excel_bytes,
                file_name=f"LOG_REKAP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error("Terjadi error tak tertangani saat proses.")
            st.exception(e)
else:
    st.info("Pilih sumber data (ZIP atau Excel multi-sheet) dan upload Master untuk memulai.")
