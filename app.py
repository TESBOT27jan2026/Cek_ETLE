import io
from urllib.parse import urlencode

from bs4 import BeautifulSoup
import pandas as pd
import requests
import streamlit as st

# --- KONFIGURASI HALAMAN STREAMLIT ---
st.set_page_config(
    page_title="CEK ETLE", page_icon="⚡", layout="wide"
)

# --- HEADER ASLI (Lengkap 24 Kolom) ---
FULL_HEADERS = [
    "No. Baris",
    "Plat Nomor",
    "No Rangka",
    "No Mesin",
    "Status ETLE",
    "Link Pengecekan",
    "Lokasi",
    "Tanggal",
    "Jenis Pelanggaran",
    "Status",
    "Link Detail",
    "Path Foto",
    "Detail Jenis Pelanggaran",
    "Dasar Hukum",
    "Waktu Kejadian",
    "Merk",
    "Tipe",
    "Warna",
    "Nomor Rangka Detail",
    "Model",
    "Tahun",
    "Nomor Mesin Detail",
    "STNK Berlaku Sampai",
    "Estimasi Denda",
]

# User Agent agar HTTP request menyerupai browser asli
HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
        " like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}


# --- FUNGSI HITUNG DENDA ---
def hitung_denda(jenis_pelanggaran):
  text = str(jenis_pelanggaran).lower()
  if any(x in text for x in ["handphone", "ponsel", "wajar"]):
    return 750000
  elif "helm" in text:
    return 250000
  elif any(x in text for x in ["sabuk", "safety belt"]):
    return 250000
  elif any(x in text for x in ["dua orang", "berboncengan"]):
    return 250000
  elif any(
      x in text
      for x in [
          "rambu",
          "marka",
          "lampu",
          "arus",
          "jalur",
          "ganjil",
          "genap",
          "kecepatan",
          "stnk",
          "keabsahan",
      ]
  ):
    return 500000
  else:
    return 0


# --- SCRAPE DETAIL PELANGGARAN ---
def scrape_violation_details(session, detail_url):
  detailed_data = {}
  try:
    r = session.get(detail_url, headers=HTTP_HEADERS, timeout=15)
    soup = BeautifulSoup(r.text, "html.parser")
    container = soup.select_one("div.col-10.d-flex.flex-column")

    if container:
      # Ambil Link Foto
      img_element = container.find("img")
      if img_element and img_element.get("src"):
        img_src = img_element["src"]
        full_url = (
            "https://etle-pmj.id" + img_src
            if img_src.startswith("/")
            else img_src
        )
        detailed_data["Path Foto"] = full_url
      else:
        detailed_data["Path Foto"] = "-"

      full_text = container.get_text()

      def extract_info(text, keyword):
        try:
          return text.split(keyword)[1].split("\n")[1].strip()
        except (IndexError, ValueError):
          return "-"

      detailed_data["Detail Jenis Pelanggaran"] = extract_info(
          full_text, "Data Pelanggaran"
      )
      detailed_data["Dasar Hukum"] = extract_info(
          full_text, detailed_data.get("Detail Jenis Pelanggaran", "")
      )
      detailed_data["Waktu Kejadian"] = extract_info(
          full_text, "Hari, Tanggal & Waktu"
      )
      detailed_data["Merk"] = extract_info(full_text, "Merk")
      detailed_data["Tipe"] = extract_info(full_text, "Tipe")
      detailed_data["Warna"] = extract_info(full_text, "Warna")
      detailed_data["Nomor Rangka Detail"] = extract_info(
          full_text, "Nomor Rangka"
      )
      detailed_data["Model"] = extract_info(full_text, "Model")
      detailed_data["Tahun"] = extract_info(full_text, "Tahun")
      detailed_data["Nomor Mesin Detail"] = extract_info(
          full_text, "Nomor Mesin"
      )
      detailed_data["STNK Berlaku Sampai"] = extract_info(
          full_text, "STNK Berlaku Sampai"
      )
  except Exception:
    pass

  return detailed_data


# --- CEK STATUS ETLE MAIN ---
def check_etle_status(session, url):
  try:
    r = session.get(url, headers=HTTP_HEADERS, timeout=20)
    soup = BeautifulSoup(r.text, "html.parser")

    # Cek Pop-up alert jika data tidak ditemukan
    popup = soup.select_one("div.swal2-html-container")
    if popup and "tidak ditemukan" in popup.text.lower():
      return "Aman", None

    # Parse baris tabel HTML
    rows = soup.select("tbody tr")
    violations = []
    for row in rows:
      cols = row.find_all("td")
      detail_link = ""

      # Ambil link detail dari tombol/anchor jika ada
      link_element = row.select_one("a.btn-secondary")
      if link_element and link_element.get("href"):
        href = link_element["href"]
        if href.startswith("http"):
          detail_link = href
        else:
          detail_link = (
              "https://etle-pmj.id" + href
              if href.startswith("/")
              else "https://etle-pmj.id/" + href
          )

      if len(cols) >= 4:
        violations.append({
            "Tanggal": cols[0].get_text(strip=True),
            "Lokasi": cols[1].get_text(strip=True),
            "Jenis Pelanggaran": cols[2].get_text(strip=True),
            "Status": cols[3].get_text(strip=True),
            "Link Detail": detail_link,
        })

    if len(violations) > 0:
      return "Ada ETLE", violations
    else:
      return "Aman", None

  except Exception as e:
    return f"Error: {e}", None


# --- ANTARMUKA USER (STREAMLIT UI) ---
st.title("⚡ Pengecekan ETLE Automated (Tanpa Selenium)")
st.write(
    "Unggah file Excel berisi kolom Plat Nomor, No Rangka, dan No Mesin untuk"
    " cek etle gabut."
)

uploaded_file = st.file_uploader(
    "Upload File Excel Input", type=["xlsx", "xls"]
)

if uploaded_file is not None:
  try:
    df = pd.read_excel(uploaded_file, dtype=str)

    # Identifikasi kolom dinamis
    col_plat = next((c for c in df.columns if "plat" in c.lower()), None)
    col_rangka = next((c for c in df.columns if "rangka" in c.lower()), None)
    col_mesin = next((c for c in df.columns if "mesin" in c.lower()), None)

    if not col_plat or not col_rangka or not col_mesin:
      st.error(
          "❌ Kolom Plat, No Rangka, atau No Mesin tidak terdeteksi pada file"
          " Excel!"
      )
    else:
      st.success(
          f"✅ Kolom terdeteksi: Plat (`{col_plat}`), Rangka (`{col_rangka}`),"
          f" Mesin (`{col_mesin}`)"
      )
      st.subheader("Preview Data Input")
      st.dataframe(df.head(), use_container_width=True)

      if st.button("🚀 Mulai Pengecekan ETLE", type="primary"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        log_box = st.expander("📄 Live Log Proses Scraping", expanded=True)

        final_report = []
        total_rows = len(df)

        # Gunakan Session untuk efisiensi HTTP Connection Reuse
        session = requests.Session()

        for idx, row in df.iterrows():
          plat = str(row[col_plat]).strip()
          rangka = str(row[col_rangka]).strip()
          mesin = str(row[col_mesin]).strip()
          real_row = idx + 2

          # Update Progress UI
          progress = (idx + 1) / total_rows
          progress_bar.progress(progress)
          status_text.text(f"Memproses {idx+1}/{total_rows}: {plat}")

          if not plat or plat.lower() == "nan":
            continue

          log_box.write(f"**[{idx+1}/{total_rows}]** Memeriksa Plat: `{plat}`")

          clean_plat = plat.replace(" ", "").replace("-", "")
          params = {
              "aksi": "cek",
              "nopol": clean_plat,
              "norangka": rangka,
              "nomesin": mesin,
          }
          url = "https://etle-pmj.id/?" + urlencode(params)

          status, violations = check_etle_status(session, url)

          base_row = dict.fromkeys(FULL_HEADERS, "-")
          base_row.update({
              "No. Baris": real_row,
              "Plat Nomor": plat,
              "No Rangka": rangka,
              "No Mesin": mesin,
              "Status ETLE": status,
              "Link Pengecekan": url,
          })

          if status == "Ada ETLE" and violations:
            log_box.warning(
                f"⚠️ **{plat}**: Ditemukan {len(violations)} pelanggaran!"
            )
            for v in violations:
              row_data = base_row.copy()
              row_data.update(v)

              if v.get("Link Detail"):
                log_box.write(" └ Scraping detail pelanggaran...")
                det = scrape_violation_details(session, v["Link Detail"])
                row_data.update(det)

              if "Jenis Pelanggaran" in row_data:
                row_data["Estimasi Denda"] = hitung_denda(
                    row_data["Jenis Pelanggaran"]
                )

              final_report.append(row_data)
          else:
            log_box.info(f"✅ **{plat}**: Status `{status}`")
            final_report.append(base_row)

        # Format DataFrame Output
        df_hasil = pd.DataFrame(final_report)
        df_hasil = df_hasil.reindex(columns=FULL_HEADERS)

        st.success("🎉 Pengecekan Selesai!")
        st.subheader("Hasil Pengecekan")
        st.dataframe(df_hasil, use_container_width=True)

        # Export Hasil ke Excel
        output_stream = io.BytesIO()
        with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
          df_hasil.to_excel(writer, index=False)
        excel_data = output_stream.getvalue()

        st.download_button(
            label="📥 Download Hasil Excel (.xlsx)",
            data=excel_data,
            file_name="HASIL_PENGECEKAN_ETLE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

  except Exception as e:
    st.error(f"Terjadi kesalahan saat membaca file Excel: {e}")
