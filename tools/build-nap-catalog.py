"""Builds the NAP SDS catalog embedded in docs/index.html (defaultData) and the
PDF library in docs/assets/sds-library from:

  - the chemical register  "ทะเบียนสารเคมี NAP.xlsx"   (sheet 1, rows from 4)
  - the SDS PDF folder      "SDS ลงในแพลตฟอร์ม NAP"

Run:  python tools/build-nap-catalog.py [register.xlsx] [pdf-folder]

The PDF <-> register matching is the reviewed table PDF_MAP below (file names
are not consistent with the register names, several PDFs are scans, two are
swapped: "17.Monoammonium Phosphate.pdf" is the MKP SDS and "18.Monopotassium
Phosphate.pdf" the MAP SDS). PDFs not in the register become extra records.
"""
import io
import json
import os
import re
import shutil
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DL = os.path.join(os.path.expanduser("~"), "Downloads")
XLSX = sys.argv[1] if len(sys.argv) > 1 else os.path.join(DL, "ทะเบียนสารเคมี NAP.xlsx")
PDF_DIR = sys.argv[2] if len(sys.argv) > 2 else os.path.join(DL, "SDS ลงในแพลตฟอร์ม NAP", "SDS ลงในแพลตฟอร์ม NAP")
SITE = "https://watanathep8-dotcom.github.io/sds-management-nap/"
LIB = os.path.join(ROOT, "docs", "assets", "sds-library")
ID_BASE = 1000

# register sequence -> {"th": Thai SDS file, "en": English (or other) SDS file}
PDF_MAP = {
    3: {"en": "2.AP Solution.pdf"},
    5: {"en": "46.Amyl Acetate liquid (กลิ่นนมแมว).pdf"},
    7: {"en": "41.Attagel 50.pdf"},
    8: {"th": "1.ACTIBOR 20.pdf"},
    9: {"en": "69.Borax Pentahydrate.pdf"},
    10: {"th": "5.Boric Acids.pdf", "en": "34.Boric Acid.pdf"},
    13: {"en": "35.Borresperse CA.pdf"},
    18: {"en": "37.Calcium Carbonate.pdf"},
    19: {"th": "36.Calcium Chloride.pdf"},
    20: {"th": "36.Calcium Chloride.pdf"},
    21: {"th": "36.Calcium Chloride.pdf"},
    22: {"th": "36.Calcium Chloride.pdf"},
    23: {"en": "44.Calcium Nitrate 12-0-0.pdf"},
    25: {"en": "7.Citric acid anhydrous.pdf"},
    26: {"th": "74.Citric acid monohydrous.pdf", "en": "68.Citric acid Monohydrate (ซิตริก แอซิดโมโนไฮเดรท).pdf"},
    35: {"en": "43.CALCIUM PROPIONATE.pdf"},
    39: {"en": "24.Dextrose Monohyrate.pdf"},
    41: {"en": "8.EDTA 2Na.pdf"},
    44: {"en": "66.Ferric Chloride 46% (เฟอร์ริค คลอไรด์).pdf"},
    45: {"th": "9.Ferrous Sulfate Heptahydrate.pdf"},
    46: {"en": "65.Ferrous Sulphate Monohydrate.pdf"},
    49: {"en": "10.Formaldehyde.pdf"},
    55: {"th": "12.Isopropanol(IPA).pdf"},
    56: {"en": "13.Kelzan(Xanthan Gum).pdf"},
    58: {"en": "14.L-Lysine Monohydrochloride.pdf"},
    59: {"en": "15.L-Threonine.pdf"},
    62: {"en": "38.Magnesium Carbonate.pdf"},
    63: {"en": "23.Magnesium Chloride Hexahydrate.pdf"},
    64: {"th": "16.Magnesium Sulfhate heptahydrate.pdf", "en": "72.Magnesime Sulphate Heptahydrate.pdf"},
    65: {"en": "60.Magnesium Sulphate Monohydrate (แม็กนีเซียม ซัลเฟต).pdf"},
    66: {"en": "59.Manganese Sulphate Monohydrate Powder.pdf"},
    69: {"en": "18.Monopotassium Phosphate.pdf"},   # content: Monoammonium Phosphate
    70: {"th": "45.MoNo Ethanolamine.pdf"},
    71: {"en": "17.Monoammonium Phosphate.pdf"},    # content: Monopotassium Phosphate (MKP)
    73: {"en": "61.Marasperse AG Powder.pdf"},
    82: {"th": "19.Phosphoric Acid.pdf"},
    84: {"th": "70.Potassium Hydroxide.pdf"},
    85: {"en": "21.Potassium Nitrate.pdf"},
    86: {"en": "22.Potassium Sulphate.pdf"},
    90: {"en": "56.Potassium Humate.pdf"},
    91: {"en": "56.Potassium Humate.pdf"},
    92: {"en": "57.Potassium Nitrate.pdf"},
    93: {"en": "20.Pluriol.pdf"},
    101: {"en": "54.Sodium Gluconate (โซเดียม กลูโคเนท).pdf"},
    102: {"en": "53.Sodium Molybdate (โซเดียม โมลิบเดท).pdf"},
    103: {"en": "76.Solubor.pdf"},
    115: {"en": "33.Thanant DF 920 Antifoam.pdf"},
    116: {"th": "30.Teepol Multipurpose Perfume Fre.pdf"},
    118: {"en": "31.Triethanolamine.pdf"},
    119: {"en": "27.Tetrapossium Pyrophosphate.pdf"},
    120: {"en": "32.Urea..pdf"},
    121: {"en": "50.Urea.pdf"},
    124: {"en": "28.Zinc Chloride Solution.pdf"},
    126: {"en": "49.Zinc Sulphate Monohydrate.pdf"},
    127: {"en": "29.Zinc Sulfate Powder.pdf"},
    128: {"en": "48.Zinc Sulphate Monohydrate (Feed Grade).pdf"},
}

# PDFs whose chemical is not in the register: added as extra records (after the register)
EXTRA = [
    ("Product", "Cannaboost Four", {"en": "Cannaboost Four.pdf"}),
    ("Product", "Cannaboost Three (7-14-34)", {"en": "Cannaboost Three (7-14-34).pdf"}),
    ("Product", "Cannaboost Two (15-5-20)", {"en": "Cannaboost Two (15-5-20).pdf"}),
    ("Product", "Excellent Mix Calcium Boron", {"en": "Excellent Mix Calcium Boron.pdf"}),
    ("Product", "Humus Kali", {"en": "Humus Kali.pdf"}),
    ("Product", "MR.K-One", {"en": "MR.K-One.pdf"}),
    ("Product", "Power Plus ST", {"en": "Power Plus ST.pdf"}),
    ("Product", "Ready Mix 17.5%Fe", {"en": "Ready Mix 17.5%Fe.pdf"}),
    ("Product", "Ready Mix Combi One", {"en": "Ready Mix Combi One.pdf"}),
    ("Product", "Super Nitrate Balancer", {"en": "Super Nitrate Balancer.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "Bactercid B20", {"en": "3.Bactercid B20.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "BIOPOL VR 20", {"th": "4.BIOPOL VR 20.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "Glucose Syrup", {"en": "11.Glucose Syrup.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "Etifert B-20", {"th": "75.ETIFERT-B20.pdf", "en": "25.Etifert-B20.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "MAX SeaSailer", {"en": "39.MAX SeaSailer.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "Magnesium Oxide", {"en": "42.Magnesium Oxide.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "REAX 88A (รีแอกซ์)", {"en": "55.REAX 88A (รีแอกซ์).pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "Methanol (เมทิล แอลกอฮอล์)", {"th": "67.Methanol เมทิล แอลกอฮอล์.pdf"}),
    ("อื่น ๆ (ไม่มีในทะเบียน)", "Ethylene Glycol", {"th": "71.Ethylene Glycol.pdf"}),
]

# Register pictogram columns D..J -> the app's pictogram files
PICTO = ["pictogram-4.png", "pictogram-5.png", "pictogram-6.png", "pictogram-7.png",
         "pictogram-8.png", "pictogram-9.png", "pictogram-12.png"]
FIRST_AID = [("การสูดดม", "Inhalation"), ("การสัมผัสผิวหนัง", "Skin contact"),
             ("การสัมผัสดวงตา", "Eye contact"), ("การกลืนกิน", "Ingestion")]
FIRE = [("ละอองน้ำ", "Water spray"), ("โฟม", "Foam"), ("ผงเคมีแห้ง", "Dry chemical powder"),
        ("คาร์บอนไดออกไซด์", "Carbon dioxide")]
PPE = ["V", "W", "X", "Y", "Z"]  # register W..AA: suit, mask, goggles, safety shoes, gloves
CAS_RE = re.compile(r"\b\d{2,7}-\d{2}-\d\b")


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def ratio_txt(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 0 < v <= 1:
        return ("%g" % (v * 100)) + "%"  # Excel percent cell
    return txt(v)


def cas_txt(v):
    s = txt(v)
    found = CAS_RE.findall(s)
    if found:
        return ", ".join(dict.fromkeys(found))
    return "" if s in ("", "-") else s


def lib_url(rid, lang):
    return SITE + "assets/sds-library/library-%d-%s.pdf" % (rid, lang)


def real_name(f):
    """The file as named on disk (numbered files matched by their number: some
    names contain non-breaking spaces)."""
    if not f or os.path.exists(os.path.join(PDF_DIR, f)):
        return f
    m = re.match(r"(\d+)\.", f)
    hits = [x for x in os.listdir(PDF_DIR) if m and x.startswith(m.group(0))]
    if len(hits) != 1:
        raise SystemExit("PDF not found: " + f)
    return hits[0]


def attach(rec, files, used):
    th, en = real_name(files.get("th")), real_name(files.get("en"))
    langs = []
    rec.update({"hasPdf": bool(th or en), "pdfUrl": "", "pdfName": "", "thaiPdfUrl": "", "thaiPdfName": "",
                "englishPdfUrl": "", "englishPdfName": ""})
    if en:
        name = "library-%d-english.pdf" % rec["id"]
        rec.update({"pdfUrl": lib_url(rec["id"], "english"), "pdfName": name,
                    "englishPdfUrl": lib_url(rec["id"], "english"), "englishPdfName": name})
        used[name] = en
        langs.append("English")
    if th:
        name = "library-%d-thai.pdf" % rec["id"]
        rec.update({"thaiPdfUrl": lib_url(rec["id"], "thai"), "thaiPdfName": name})
        used[name] = th
        langs.append("Thai")
    rec["sdsLanguages"] = langs
    rec["thaiSds"] = bool(th)
    rec["language"] = "English" if en or not th else "Thai"


def base_record(rid, seq, group, name):
    return {
        "id": rid, "chemical": name, "chemicalEnglish": name, "chemicalThai": name,
        "cas": "", "composition": "", "ratio": "", "components": [], "supplier": "",
        "responsibleParty": "", "recorder": "", "dateOfUse": "", "registerSds": "",
        "revision": "", "status": "Current", "signalWord": "", "hazards": [], "pictograms": [],
        "flashPoint": "", "sequence": seq, "group": group,
        "firstAidMeasures": [], "firefightingMeasures": [], "personalProtectiveEquipment": [],
    }


def read_register():
    ws = openpyxl.load_workbook(XLSX, data_only=True).worksheets[0]
    out, cur = [], None
    for r in ws.iter_rows(min_row=4, values_only=True):
        r = list(r) + [None] * 40
        if r[0] is not None and txt(r[2]):
            cur = {"row": r, "components": []}
            out.append(cur)
        if cur is None:
            continue
        comp = txt(r[10])
        if comp or txt(r[11]) or txt(r[12]):
            cur["components"].append({"composition": comp, "cas": cas_txt(r[11]), "ratio": ratio_txt(r[12])})
    return out


def main():
    used = {}
    data = []
    for item in read_register():
        r = item["row"]
        seq = int(r[0])
        rec = base_record(ID_BASE + seq, seq, txt(r[1]) or "Fertilizer", txt(r[2]))
        comps = item["components"]
        rec["components"] = comps
        rec["composition"] = "\n".join(c["composition"] for c in comps)
        rec["cas"] = "\n".join(c["cas"] or "-" for c in comps) if len(comps) > 1 else (comps[0]["cas"] if comps else "")
        rec["ratio"] = "\n".join(c["ratio"] for c in comps)
        rec["hazards"] = [PICTO[i] for i in range(7) if txt(r[3 + i])]
        rec["pictograms"] = list(rec["hazards"])
        rec["flashPoint"] = txt(r[13])
        rec["firstAidMeasures"] = [{"thai": t, "english": e, "value": txt(r[14 + i])}
                                   for i, (t, e) in enumerate(FIRST_AID) if txt(r[14 + i])]
        rec["firefightingMeasures"] = [{"thai": t, "english": e, "detail": ""}
                                       for i, (t, e) in enumerate(FIRE) if r[18 + i] is True]
        rec["personalProtectiveEquipment"] = [code for i, code in enumerate(PPE) if r[22 + i] is True]
        rec["responsibleParty"] = txt(r[27])
        # recorder / date of use are left empty (as on the Thepwatana site)
        attach(rec, PDF_MAP.get(seq, {}), used)
        data.append(rec)
    seq = max(r["sequence"] for r in data)
    for group, name, files in EXTRA:
        seq += 1
        rec = base_record(ID_BASE + seq, seq, group, name)
        attach(rec, files, used)
        data.append(rec)

    # PDF library: exactly the files referenced above
    os.makedirs(LIB, exist_ok=True)
    for f in os.listdir(LIB):
        if f.startswith("library-") and f.endswith(".pdf") and f not in used:
            os.remove(os.path.join(LIB, f))
    for name, src in used.items():
        shutil.copyfile(os.path.join(PDF_DIR, src), os.path.join(LIB, name))

    # embed into docs/index.html
    p = os.path.join(ROOT, "docs", "index.html")
    s = io.open(p, encoding="utf-8", newline="").read()
    crlf = "\r\n" in s
    s = s.replace("\r\n", "\n")
    a = s.index("  const defaultData = [")
    b = s.index("  const REGISTER_WORKBOOK_DATA = {")
    c = s.index("  defaultData.forEach(record => {\n    const source = REGISTER_WORKBOOK_DATA")
    d = s.index("  const thaiChemicalNames = {")
    e = s.index("  defaultData.forEach(record => {\n    const englishName")
    body = json.dumps(data, ensure_ascii=False, indent=2).replace("\n", "\n  ")
    s = (s[:a] + "  const defaultData = " + body + ";\n"
         + "  const REGISTER_WORKBOOK_DATA = {};\n" + s[c:d]
         + "  const thaiChemicalNames = {};\n" + s[e:])
    if crlf:
        s = s.replace("\n", "\r\n")
    io.open(p, "w", encoding="utf-8", newline="").write(s)

    all_pdfs = sorted(f for f in os.listdir(PDF_DIR) if f.lower().endswith(".pdf"))
    unused = [f for f in all_pdfs if f not in set(used.values())]
    print("records:", len(data), "(register %d + extra %d)" % (len(data) - len(EXTRA), len(EXTRA)))
    print("with PDF:", sum(1 for r in data if r["hasPdf"]), "library files:", len(used))
    print("PDFs not used (duplicates of a used SDS):", len(unused))
    for f in unused:
        print("  -", f)


if __name__ == "__main__":
    main()
