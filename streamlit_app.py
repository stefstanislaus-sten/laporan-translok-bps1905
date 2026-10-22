import io
import re
from datetime import date
from xml.sax.saxutils import escape

import streamlit as st
from PIL import Image, ImageOps

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (Image as RLImage, KeepTogether, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

st.set_page_config(page_title="Generator Laporan Translok", page_icon="📝", layout="wide")

# ----------------------------------------------------------------------------
# DATA
# ----------------------------------------------------------------------------
SURVEI = [
    "Survei Kerangka Sampel Area (Survei KSA)",
    "Survei Ubinan",
    "SP Palawija",
    "SP Lahan",
    "Survei Konversi Gabah ke Beras (SKGB)",
    "Statistik Pertanian Hortikultura (SPH)",
    "VN Horti",
    "Survei Indeks Produksi Hortikultura (VIP-Hortikultura)",
    "Survei Perkebunan Bulanan (SEDAPP Online)",
    "Survei Perusahaan Perkebunan (SKB) Tahunan",
    "Updating Direktori Perusahaan Pertanian dan Usaha Tani Lainnya (UDPP DUTL)",
    "Survei Komoditas Strategis (Komstrat)",
    "Laporan Pemotongan Ternak Bulanan (LPTB)",
    "Laporan Tahunan Perusahaan Budidaya dan Penangkapan Ikan (LTB dan LTP)",
    "Survei Kehutanan",
    "Pencacahan Lengkap Tempat Pelelangan Ikan Triwulanan (PP/TPI)",
    "Survei IBS Tahunan - Pemutakhiran Direktori Perusahaan Awal (DPA)",
    "Survei IBS Tahunan - Survei Tahunan Perusahaan Industri Manufaktur - STPIM A",
    "Survei IBS Triwulan B",
    "STPU (Survei Triwulanan Pelaku Usaha) Triwulan B",
    "Survei IMK Triwulan B",
    "Survei IMK Tahunan",
    "Updating Perusahaan PE (Pertambangan Energi)",
    "Survei PE Tahunan (Survei Galian URT)",
    "Survei PE Tahunan (Survei Non Migas)",
    "Survei PE Tahunan (Survei Captive Power)",
    "Survei PE Tahunan (Survei Air Bersih)",
    "Survei PE Triwulanan (Air Bersih) Triwulan B",
    "Pengutipan Direktori Konstruksi",
    "Survei Konstruksi Tahunan",
    "Survei Kesejahteraan Petani (SKP)",
    "Survei Konstruksi Triwulanan Triwulan B",
]
KEGIATAN = ["Listing", "Pendataan"]
TAHUN = [2026, 2027, 2028, 2029, 2030]
SUBROUND = ["Subround 1 (SR 1)", "Subround 2 (SR 2)", "Subround 3 (SR 3)"]
JENIS_UBINAN = ["Palawija", "Padi"]

KECAMATAN = {  # label dropdown -> nama di database desa
    "Payung": "PAYUNG",
    "Pulau Besar": "PULAU BESAR",
    "Simpang Rimba": "SIMPANG RIMBA",
    "Toboali": "TOBOALI",
    "Tukak Sadai": "TUKAK SADAI",
    "Airgegas": "AIR GEGAS",
    "Lepar": "LEPAR PONGOK",
    "Kepulauan Pongok": "KEPULAUAN PONGOK",
}
KEC_DOC = {"Airgegas": "Air Gegas"}  # penulisan di dokumen

_DESA_RAW = """Desa GADUNG|TOBOALI
Kelurahan TELADAN|TOBOALI
Kelurahan TANJUNG KETAPANG|TOBOALI
Kelurahan TOBOALI|TOBOALI
Desa SERDANG|TOBOALI
Desa JERIJI|TOBOALI
Desa KEPOH|TOBOALI
Desa BIKANG|TOBOALI
Desa RINDIK|TOBOALI
Desa RIAS|TOBOALI
Desa KAPOSANG|TOBOALI
Desa SEBAGIN|SIMPANG RIMBA
Desa GUDANG|SIMPANG RIMBA
Desa SIMPANG RIMBA|SIMPANG RIMBA
Desa BANGKA KOTA|SIMPANG RIMBA
Desa RAJIK|SIMPANG RIMBA
Desa PERMIS|SIMPANG RIMBA
Desa JELUTUNG II|SIMPANG RIMBA
Desa MALIK|PAYUNG
Desa NADUNG|PAYUNG
Desa RANGGUNG|PAYUNG
Desa IRAT|PAYUNG
Desa PAYUNG|PAYUNG
Desa BEDENGUNG|PAYUNG
Desa PAKU|PAYUNG
Desa PANGKALBULUH|PAYUNG
Desa SENGIR|PAYUNG
Desa FAJAR INDAH|PULAU BESAR
Desa SUMBER JAYA PERMAI|PULAU BESAR
Desa PANCA TUNGGAL|PULAU BESAR
Desa SUKAJAYA|PULAU BESAR
Desa BATU BETUMPANG|PULAU BESAR
Desa NANGKA|AIR GEGAS
Desa TEPUS|AIR GEGAS
Desa DELAS|AIR GEGAS
Desa NYELANDING|AIR GEGAS
Desa SIDOHARJO|AIR GEGAS
Desa PERGAM|AIR GEGAS
Desa AIRGEGAS|AIR GEGAS
Desa BENCAH|AIR GEGAS
Desa AIR BARA|AIR GEGAS
Desa RANGGAS|AIR GEGAS
Desa TIRAM|TUKAK SADAI
Desa BUKIT TERAP|TUKAK SADAI
Desa TUKAK|TUKAK SADAI
Desa SADAI|TUKAK SADAI
Desa PASIR PUTIH|TUKAK SADAI
Desa CELAGEN|KEPULAUAN PONGOK
Desa PONGOK|KEPULAUAN PONGOK
Desa KUMBUNG|LEPAR PONGOK
Desa PENUTUK|LEPAR PONGOK
Desa TANJUNG SANGKAR|LEPAR PONGOK
Desa TANJUNG LABU|LEPAR PONGOK"""
DESA_DB = dict(line.split("|") for line in _DESA_RAW.splitlines())

_PELAKSANA_RAW = """Agung Rachmadi, SE|197111261999031003
Imam Hidayat, SST, M.Si|197807081999121001
Damri Ismail|197102212002121003
Firman, SST., M.M.|198604052009121002
KMS M Hidayat, SST|198902062013111001
Sridewi Nainggolan, S.Si, M.Si|198505132009022009
Rezky Mayda Putri, SST|199305212016022001
Ali Shaumanjaya, A.Md|197409182002121004
Zullizar, SE|198510092010031003
Raka Suryadi, S.Stat|199907012019121001
Sutri Vininda, S.Tr. Stat|199903122021042001
Risma Dillah, A.Md.Stat|200104132022012001
Adinda Hermambang, S.Tr.Stat|199812302022012002
Zahrotul Firdaus, S.Tr.Stat|199902182022012001
Ragil Chelvin Pratama, A.Md.Stat|200112212023021002
Fasyina Tsurayya Nisyah, S.Tr.Stat|200009122023102001
Prasetya Dwi Anggara Kushendrawan, SST|199410172018021001
Ruri Tri Wahyuni, S.Tr.Stat|200210092024122001
Ego, A.Md.Stat|200202072024121003
Nurusyifa Somanaya, S.Tr.Stat|200309202024122001
Maya Imas Tirany, S.Tr.Stat|200007172024122001
Lailatul Hasanah, S.Tr.Stat|200112032024122001
Umi Syafiyah, S.Tr.Stat|200107222023102002
Dwi Lestari, S.Tr. Stat|199811072019122001
Stenislaus Angga Aprianto, S.Tr.Stat.|200304072026031003
Satrya, A.Md|198601272024211001
Rosidi|198703232025211051
Samsudin|198704202025211050"""


def _clean_name(n: str) -> str:
    n = n.replace("S.Tr. Stat", "S.Tr.Stat")
    if n.endswith("S.Tr.Stat."):
        n = n[:-1]
    return n


PELAKSANA = [(_clean_name(a), b) for a, b in (l.split("|") for l in _PELAKSANA_RAW.splitlines())]

BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]

ROMAN = {"I", "II", "III", "IV", "V", "VI"}


# ----------------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------------
def pretty_desa(raw: str) -> str:
    """'Desa BUKIT TERAP' -> 'Desa Bukit Terap' (Capitalize Each Word)."""
    words = raw.split()
    return " ".join(w if w in ROMAN else w.capitalize() for w in words)


def tgl_id(d: date) -> str:
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


def survei_label(survei, tahun, triwulan, subround, jenis):
    """Return (label, tahun_suffix)."""
    if survei == "Survei Ubinan":
        return f"Survei Ubinan {jenis} {subround}", f" tahun {tahun}"
    label, suffix = survei, f" tahun {tahun}"
    if re.search(r"\bA$", label):
        label = re.sub(r"\bA$", str(tahun), label)
        suffix = ""  # tahun sudah masuk ke nama survei
    if re.search(r"\bB$", label):
        label = re.sub(r"\bB$", str(triwulan), label)
    return label, suffix


def split_names(text: str):
    return [x.strip() for x in re.split(r"[,\n]+", text) if x.strip()]


def process_photo(uploaded):
    img = Image.open(uploaded)
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((1400, 1400))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    w, h = img.size
    w_cm = 8.5
    h_cm = w_cm * h / w
    if h_cm > 9.5:  # batasi tinggi agar tetap rapi
        h_cm = 9.5
        w_cm = h_cm * w / h
    return buf.getvalue(), w_cm, h_cm


ITALIC_RE = re.compile(r"(Subround|door to door)")


# ----------------------------------------------------------------------------
# WORD
# ----------------------------------------------------------------------------
def _set_font(run, size=12, bold=False, italic=False, underline=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.underline = underline


def _add_runs(par, text, size=12):
    for part in ITALIC_RE.split(text):
        if part:
            r = par.add_run(part)
            _set_font(r, size, italic=bool(ITALIC_RE.fullmatch(part)))


def _no_space(par, after=0):
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.space_after = Pt(after)


def build_docx(d) -> bytes:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.5)
    sec.top_margin = sec.bottom_margin = Cm(2)

    st_normal = doc.styles["Normal"]
    st_normal.font.name = "Times New Roman"
    st_normal.font.size = Pt(12)

    # Judul
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_font(p.add_run("Laporan Perjalanan Translok"), 13, bold=True, underline=True)
    _no_space(p, 24)

    # Header
    rows = [("Kepada", d["kepada"]), ("Perihal", d["perihal"]),
            ("Tujuan", d["tujuan"]), ("Tanggal Pelaksanaan", d["tanggal"])]
    t = doc.add_table(rows=len(rows), cols=3)
    t.autofit = False
    widths = [Cm(4.2), Cm(0.4), Cm(11.4)]
    for i, (a, b) in enumerate(rows):
        for j, txt in enumerate([a, ":", b]):
            c = t.rows[i].cells[j]
            c.width = widths[j]
            par = c.paragraphs[0]
            _no_space(par)
            _add_runs(par, txt)
    sp = doc.add_paragraph()
    _no_space(sp, 6)

    p = doc.add_paragraph()
    _add_runs(p, "Hasil Pelaksanaan")
    _no_space(p, 2)

    def numbered(num, text):
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = par.paragraph_format
        pf.left_indent = Cm(0.75)
        pf.first_line_indent = Cm(-0.75)
        pf.tab_stops.add_tab_stop(Cm(0.75))
        _no_space(par, 3)
        _add_runs(par, f"{num}.\t{text}")
        return par

    n = 0
    for txt in d["points"]:
        n += 1
        numbered(n, txt)
    n += 1
    numbered(n, "Dokumentasi:")
    for img_bytes, w_cm, h_cm in d["photos"]:
        par = doc.add_paragraph()
        par.paragraph_format.left_indent = Cm(0.75)
        _no_space(par, 4)
        par.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(w_cm), height=Cm(h_cm))

    sp = doc.add_paragraph()
    _no_space(sp, 6)

    # Tanda tangan
    sig = doc.add_table(rows=1, cols=2)
    sig.autofit = False
    sig.alignment = WD_TABLE_ALIGNMENT.LEFT
    sig.rows[0].cells[0].width = Cm(8.5)
    sig.rows[0].cells[1].width = Cm(7.5)
    cell = sig.rows[0].cells[1]
    lines = [f"Toboali, {d['tanggal']}", "Pelaksana SPD", "", "", "", d["nama"], f"NIP. {d['nip']}"]
    first = True
    for ln in lines:
        par = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _no_space(par)
        _add_runs(par, ln)
    # jaga tanda tangan tidak terpisah halaman
    tr_pr = sig.rows[0]._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    tr_pr.append(cant)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ----------------------------------------------------------------------------
# PDF
# ----------------------------------------------------------------------------
def _rl(text):
    return ITALIC_RE.sub(r"<i>\1</i>", escape(text))


def build_pdf(d) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=2.5 * cm, rightMargin=2.5 * cm,
                            topMargin=2 * cm, bottomMargin=2 * cm,
                            title="Laporan Perjalanan Translok")
    base = ParagraphStyle("base", fontName="Times-Roman", fontSize=12, leading=15, alignment=TA_LEFT)
    title = ParagraphStyle("title", parent=base, fontName="Times-Bold", fontSize=13,
                           alignment=TA_CENTER, spaceAfter=22)
    just = ParagraphStyle("just", parent=base, alignment=TA_JUSTIFY, leftIndent=0.75 * cm,
                          firstLineIndent=-0.75 * cm, spaceAfter=3, bulletIndent=0)
    center = ParagraphStyle("center", parent=base, alignment=TA_CENTER)

    story = [Paragraph("<u>Laporan Perjalanan Translok</u>", title)]

    rows = [("Kepada", d["kepada"]), ("Perihal", d["perihal"]),
            ("Tujuan", d["tujuan"]), ("Tanggal Pelaksanaan", d["tanggal"])]
    tbl = Table([[Paragraph(a, base), Paragraph(":", base), Paragraph(_rl(b), base)] for a, b in rows],
                colWidths=[4.2 * cm, 0.4 * cm, 11.4 * cm], hAlign="LEFT")
    tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                             ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                             ("TOPPADDING", (0, 0), (-1, -1), 0),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story += [tbl, Spacer(1, 14), Paragraph("Hasil Pelaksanaan", base), Spacer(1, 2)]

    n = 0
    for txt in d["points"]:
        n += 1
        story.append(Paragraph(_rl(txt), just, bulletText=f"{n}."))
    n += 1
    story.append(Paragraph("Dokumentasi:", just, bulletText=f"{n}."))
    for img_bytes, w_cm, h_cm in d["photos"]:
        im = RLImage(io.BytesIO(img_bytes), width=w_cm * cm, height=h_cm * cm)
        t = Table([["", im]], colWidths=[0.75 * cm, w_cm * cm], hAlign="LEFT")
        t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                               ("TOPPADDING", (0, 0), (-1, -1), 2),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        story.append(t)

    story.append(Spacer(1, 14))
    right = [Paragraph(f"Toboali, {d['tanggal']}", center), Paragraph("Pelaksana SPD", center),
             Spacer(1, 45), Paragraph(escape(d["nama"]), center),
             Paragraph(f"NIP. {d['nip']}", center)]
    sig = Table([["", right]], colWidths=[8.5 * cm, 7.5 * cm], hAlign="LEFT")
    sig.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                             ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(KeepTogether(sig))
    doc.build(story)
    return buf.getvalue()


def render_preview(pdf_bytes):
    import fitz  # PyMuPDF
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    return [p.get_pixmap(dpi=110).tobytes("png") for p in pdf]


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.title("📝 Generator Laporan Perjalanan Translok")
st.caption("Isi form → cek & edit di preview → pilih format → download.")

col_form, col_prev = st.columns([1, 1.15], gap="large")

with col_form:
    st.subheader("1. Isi Data")
    c1, c2 = st.columns(2)
    kegiatan = c1.selectbox("Pilih Kegiatan", KEGIATAN)
    tahun = c2.selectbox("Pilih Tahun", TAHUN)
    survei = st.selectbox("Pilih Survei", SURVEI, format_func=lambda s: s)

    subround = jenis = triwulan = None
    if survei == "Survei Ubinan":
        c1, c2 = st.columns(2)
        subround = c1.selectbox("Pilih Subround", SUBROUND)
        jenis = c2.selectbox("Jenis Ubinan", JENIS_UBINAN)
    if re.search(r"\bB$", survei):
        triwulan = st.selectbox("Pilih Triwulan (nilai B)", [1, 2, 3, 4])

    kec = st.selectbox("Tujuan: Kecamatan", list(KECAMATAN))
    desa = st.selectbox("Pilih Desa", list(DESA_DB), format_func=pretty_desa)
    unit_label = "SLS" if kegiatan == "Listing" else "sampel"
    sls = st.text_input(f"Ketikkan nama {unit_label}",
                        placeholder="RT 007 Dusun Terap" if kegiatan == "Listing" else "BPP Air Gegas")
    pcl = st.text_input("Ketik Nama PCL (1 PCL)")
    min_resp = 3 if kegiatan == "Listing" else 1
    resp_label = ("Ketik Nama Responden (min. 3 responden)" if kegiatan == "Listing"
                  else "Ketik Nama Responden / Penanggungjawab Usaha")
    resp_txt = st.text_area(resp_label, help="Pisahkan dengan koma atau baris baru.", height=90)

    tgl_default = date.today() if date.today().year == tahun else date(tahun, 1, 1)
    tgl = st.date_input("Pilih Tanggal", value=tgl_default, min_value=date(tahun, 1, 1),
                        max_value=date(tahun, 12, 31), key=f"tgl_{tahun}", format="DD/MM/YYYY")

    pel_idx = st.selectbox("Nama Pelaksana", range(len(PELAKSANA)),
                           format_func=lambda i: PELAKSANA[i][0])
    nama_pel, nip = PELAKSANA[pel_idx]

    files = st.file_uploader("Upload Dokumentasi (JPG/JPEG, maks 5 MB per foto)",
                             type=["jpg", "jpeg"], accept_multiple_files=True)

    # ---- validasi ----
    errors = []
    if DESA_DB[desa] != KECAMATAN[kec]:
        valid = [pretty_desa(k) for k, v in DESA_DB.items() if v == KECAMATAN[kec]]
        errors.append(f"{pretty_desa(desa)} bukan bagian dari Kecamatan {KEC_DOC.get(kec, kec)}. "
                      f"Desa yang sesuai: {', '.join(valid)}.")
    if not sls.strip():
        errors.append(f"Nama {unit_label} belum diisi.")
    if not pcl.strip():
        errors.append("Nama PCL belum diisi.")
    responden = split_names(resp_txt)
    if len(responden) < min_resp:
        errors.append(f"Nama responden minimal {min_resp} orang (terisi {len(responden)}).")
    photos = []
    for f in files or []:
        if f.size > 5 * 1024 * 1024:
            errors.append(f"Foto '{f.name}' melebihi 5 MB.")
        else:
            try:
                photos.append(process_photo(f))
            except Exception:
                errors.append(f"Foto '{f.name}' tidak bisa dibaca.")

for e in errors:
    col_form.error(e)

# ----------------------------------------------------------------------------
# Teks otomatis
# ----------------------------------------------------------------------------
label, suffix = survei_label(survei, tahun, triwulan, subround, jenis)
phrase = f"{kegiatan} {label}"
kec_doc = KEC_DOC.get(kec, kec)
desa_pretty = pretty_desa(desa)
tgl_str = tgl_id(tgl)

auto_perihal = f"Laporan Perjalanan Dinas Dalam Kota dalam rangka melakukan Pengawasan {phrase}{suffix}"
auto_p1 = (f"Pada saat melakukan pengawasan {phrase}{suffix} di Kecamatan {kec_doc} "
           f"tanggal {tgl.day} {BULAN[tgl.month - 1]} tahun {tgl.year}, saya bertemu dengan PCL {pcl.strip()}.")
if kegiatan == "Listing":
    auto_p2 = (f"Saya melakukan pengawasan {phrase}{suffix} di 1 SLS yaitu {sls.strip()}, {desa_pretty} "
               f"dan bertemu dengan beberapa responden, yaitu {', '.join(responden)}, dll.")
    auto_p3 = (f"Pengawasan dilakukan dengan mendampingi dan mengawasi petugas lapangan saat melakukan "
               f"listing {label} door to door di SLS tersebut, dan melihat ketepatan wawancaranya.")
else:
    auto_p2 = (f"Saya melakukan pengawasan {phrase}{suffix} di 1 sampel, yaitu {sls.strip()} {desa_pretty} "
               f"dan bertemu dengan Penanggungjawab Usaha, yaitu {', '.join(responden)}.")
    auto_p3 = (f"Pengawasan dilakukan dengan mendampingi dan mengawasi petugas lapangan saat melakukan "
               f"{phrase} di desa tersebut, dan melihat ketepatan wawancara dan isiannya.")

sig_key = (kegiatan, survei, tahun, subround, jenis, triwulan, kec, desa, sls, pcl,
           tuple(responden), str(tgl))
if st.session_state.get("sig") != sig_key:
    st.session_state.update(sig=sig_key, f_perihal=auto_perihal,
                            f_p1=auto_p1, f_p2=auto_p2, f_p3=auto_p3)


def _reset():
    st.session_state["sig"] = None


# ----------------------------------------------------------------------------
# Preview + edit
# ----------------------------------------------------------------------------
with col_prev:
    st.subheader("2. Preview & Edit")
    if errors:
        st.info("Lengkapi/perbaiki data di sebelah kiri untuk menampilkan preview.")
        st.stop()

    with st.expander("✏️ Edit teks laporan (tambah / ubah sebelum download)", expanded=False):
        st.text_area("Perihal", key="f_perihal", height=80)
        st.text_area("Poin 1", key="f_p1", height=100)
        st.text_area("Poin 2", key="f_p2", height=110)
        st.text_area("Poin 3", key="f_p3", height=100)
        extra_txt = st.text_area("Poin tambahan (opsional, satu poin per baris — tampil sebelum Dokumentasi)",
                                 key="f_extra", height=80)
        st.button("↩️ Kembalikan ke teks otomatis", on_click=_reset)

    extras = [x.strip() for x in extra_txt.splitlines() if x.strip()]
    data = {
        "kepada": "Kepala BPS Kabupaten Bangka Selatan",
        "perihal": st.session_state["f_perihal"],
        "tujuan": f"Kecamatan {kec_doc}",
        "tanggal": tgl_str,
        "points": [st.session_state["f_p1"], st.session_state["f_p2"], st.session_state["f_p3"]] + extras,
        "photos": photos,
        "nama": nama_pel,
        "nip": nip,
    }

    pdf_bytes = build_pdf(data)
    for png in render_preview(pdf_bytes):
        st.image(png, use_container_width=True)
    st.caption("Preview dibuat dari versi PDF. Tata letak file Word bisa sedikit berbeda.")

    st.subheader("3. Download")
    fmt = st.radio("Pilih format download", ["PDF", "Word (.docx)"], horizontal=True)
    fname = f"Laporan_Translok_{tgl.strftime('%Y-%m-%d')}"
    if fmt == "PDF":
        st.download_button("⬇️ Download PDF", pdf_bytes, file_name=f"{fname}.pdf",
                           mime="application/pdf", type="primary")
    else:
        st.download_button("⬇️ Download Word", build_docx(data), file_name=f"{fname}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           type="primary")
