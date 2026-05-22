import re
import os
import pandas as pd

# Nama file database Excel yang akan dibuat/diperbarui
EXCEL_FILE = "database_absensi.xlsx"


def ekstrak_data_wa(teks_wa):
    """Fungsi untuk membedah teks WA menggunakan Regex (Regular Expressions)"""
    try:
        # Pola pencarian data (Regex)
        tanggal = re.search(r"Absensi Tgl\s*:\s*(.*)", teks_wa).group(1).strip()
        hari = re.search(r"Hari\s*:\s*(.*)", teks_wa).group(1).strip()
        line = re.search(r"Line\s*:\s*(.*)", teks_wa).group(1).strip()
        gl_tl = re.search(r"GL/TL\s*:\s*(\d+)", teks_wa).group(1).strip()
        tm = re.search(r"TM\.\s*:\s*(\d+)", teks_wa).group(1).strip()

        cuti_plan = re.search(r"1\.\s*Cuti Plan\s*:\s*(.*)", teks_wa).group(1).strip()
        cuti_unplan_skel = (
            re.search(r"2\.\s*Cuti Unplan kel\.sakit\s*=\s*(.*)", teks_wa)
            .group(1)
            .strip()
        )
        cuti_unplan_ukel = (
            re.search(r"3\.\s*Cuti Unplan kel\.urgent\s*=\s*(.*)", teks_wa)
            .group(1)
            .strip()
        )
        cuti_unplan_sakit = (
            re.search(r"4\.\s*Cuti Unplan sakit\s*=\s*(.*)", teks_wa).group(1).strip()
        )
        training = re.search(r"5\.\s*Training\s*=\s*(.*)", teks_wa).group(1).strip()
        other = re.search(r"6\.\s*Other\s*=\s*(.*)", teks_wa).group(1).strip()
        total_absen = (
            re.search(r"Total MP Absen\s*:\s*(\d+)", teks_wa).group(1).strip()
        )

        # Susun data menjadi struktur Dictionary/Baris Database
        data_baru = {
            "Tanggal": [tanggal],
            "Hari": [hari],
            "Line": [line],
            "GL_TL_MP": [int(gl_tl)],
            "TM_MP": [int(tm)],
            "Cuti_Plan": [cuti_plan],
            "Cuti_Unplan_Kel_Sakit": [cuti_unplan_skel],
            "Cuti_Unplan_Kel_Urgent": [cuti_unplan_ukel],
            "Cuti_Unplan_Sakit": [cuti_unplan_sakit],
            "Training": [training],
            "Other": [other],
            "Total_MP_Absen": [int(total_absen)],
        }
        return pd.DataFrame(data_baru)

    except AttributeError:
        print(
            "\n[Error] Format teks WA yang kamu masukkan tidak sesuai atau ada baris yang hilang!"
        )
        return None


def simpan_ke_excel(df_baru):
    """Fungsi untuk menyimpan atau menumpuk data ke dalam file Excel"""
    if os.path.exists(EXCEL_FILE):
        # Jika file excel sudah ada, baca data lama lalu gabungkan dengan data baru (Append)
        df_lama = pd.read_excel(EXCEL_FILE)
        df_total = pd.concat([df_lama, df_baru], ignore_index=True)
    else:
        # Jika belum ada, buat baru
        df_total = df_baru

    # Simpan ke Excel
    df_total.to_excel(EXCEL_FILE, index=False)
    print(f"\n[Sukses] Data berhasil disimpan ke dalam file '{EXCEL_FILE}'!")


if __name__ == "__main__":
    print("==================================================")
    print("   PROGRAM INPUT WA TO EXCEL DATABASE")
    print("==================================================")
    print("Silakan copas format laporan dari WA ke bawah ini.")
    print("Jika sudah selesai, tekan Enter, lalu tekan Ctrl+Z (Windows) ")
    print("atau Ctrl+D (Mac/Linux) di baris baru untuk memproses.\n")

    # Membaca input multi-line (banyak baris sekaligus dari copas)
    lines = []
    while True:
        try:
            line = input()
            lines.append(line)
        except EOFError:
            break

    teks_input_wa = "\n".join(lines)

    if teks_input_wa.strip():
        # Ekstrak data dari teks
        df_hasil = ekstrak_data_wa(teks_input_wa)

        # Jika ekstraksi sukses, masukkan ke excel
        if df_hasil is not None:
            simpan_ke_excel(df_hasil)
    else:
        print("[Gagal] Tidak ada teks yang dimasukkan.")