"""
DGD Auto-Generator - Web App
-----------------------------
Upload an Annexure Excel + DGD Template, download the generated DGD(s).
Run locally with:  streamlit run app.py
Or deploy free at:  https://share.streamlit.io
"""

import re
import io
import zipfile
from pathlib import Path
from collections import defaultdict

import streamlit as st
import openpyxl

# ---------------------------------------------------------------------------
# EDIT THIS: shipment-level header info, same for every DGD in a shipment/PO
# but not present in the Annexure file.
# ---------------------------------------------------------------------------
DEFAULT_HEADER = {
    "shipper_lines": [
        "AXALTA COATING SYSTEMS INDIA PVT LTD",
        "C/O TVS SUPPLY CHAIN SOLUTIONS LTD",
        "SURVEY NO.258, VILLAGE - ALINDRA",
        "VADODARA 391775, INDIA",
    ],
    "emergency_contact": "MR. HIREN BHATT - +91 8511896338",
    "consignee_lines": [
        "AXALTA COATING SYSTEMS AUSTRALIA PTY LTD",
        "16 DARLING STREET",
        "MARSDEN PARK NSW 2765",
        "AUSTRALIA",
    ],
    "carrier": "ANL INDIA",
    "packer_signatory_company": "AXALTA COATING SYSTEMS INDIA PVT LTD",
    "packer_place_date": "VADODARA-10-08-2026",
    "packer_signatory_name": "Mr. Pranav Dave",
    "port_of_loading": "NHAVA SHEVA",
    "port_of_unloading": "SYDNEY , AUSTRALIA",
    "final_place_of_delivery": "SYDNEY , AUSTRALIA",
    "shipper_decl_company": "AXALTA COATING SYSTEMS INDIA PVT LTD",
    "shipper_decl_place_date": "VADODARA-10-08-2026",
    "shipper_decl_signatory_name": "Mr. Pranav Dave",
}

COL = {
    "material_desc": 3, "no_in_tins": 6, "net_wt_kg": 9, "no_of_boxes_drums": 10,
    "gross_wt_with_pellet_kg": 14, "un_number": 15, "psn": 16, "haz_class": 17,
    "pkg_group": 18, "marine_pollutant": 19, "flash_point": 20, "ems": 21,
    "un_cert_no": 24, "un_cert_no_1": 25, "technical_name": 30,
}

CELLS = {
    "un_no": "D36", "gr_wt": "H36", "psn": "D38", "nt_wt": "H38",
    "technical_name": "D40", "outer_packages": "A43", "class": "C42",
    "pkg_group": "D44", "inner_packing": "A45", "packing_code": "D46",
    "ems": "D48", "flash_point": "D50", "marine_pollutant": "D52",
}

HEADER_CELLS = {
    "shipper_lines": ["A9", "A10", "A11", "A12"],
    "emergency_contact": "D15",
    "consignee_lines": ["A17", "A18", "A19", "A20"],
    "carrier": "D17",
    "packer_signatory_company": "D22",
    "packer_place_date": "D25",
    "packer_signatory_name": "D28",
    "port_of_loading": "B31",
    "port_of_unloading": "A33",
    "final_place_of_delivery": "B33",
    "shipper_decl_company": "D55",
    "shipper_decl_place_date": "D57",
    "shipper_decl_signatory_name": "D60",
}


def parse_flash_point(text):
    if text is None:
        return None
    match = re.search(r"[-+]?\d*\.?\d+", str(text))
    return float(match.group()) if match else None


def packing_prefix(cert_code_str):
    if not cert_code_str:
        return ""
    return str(cert_code_str).strip().split("/")[0].strip()


def packing_unit_word(prefix):
    if prefix.startswith("1A"):
        return "DRUMS"
    if prefix.startswith("4G"):
        return "BOXES"
    return "PACKAGES"


def psn_category(psn_text):
    if psn_text is None:
        return ""
    text = str(psn_text).strip().upper()
    return "PAINT" if text == "PAINT" else "PAINT RELATED"


def load_annexure_rows(file_bytes, sheet_name="Filtered Data"):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[COL["un_number"]] is None:
            continue
        rows.append(row)
    return rows


def build_group_key(row):
    flash_val = parse_flash_point(row[COL["flash_point"]])
    flash_bucket = "LT23" if (flash_val is not None and flash_val < 23) else "GT23"
    psn_cat = psn_category(row[COL["psn"]])
    pkg_prefix = packing_prefix(row[COL["un_cert_no_1"]])
    marine = str(row[COL["marine_pollutant"]]).strip().upper()
    un_no = row[COL["un_number"]]
    return (flash_bucket, psn_cat, pkg_prefix, marine, un_no)


def group_rows(rows):
    groups = defaultdict(list)
    for row in rows:
        key = build_group_key(row)
        groups[key].append(row)
    return groups


def aggregate_group(rows):
    un_no = rows[0][COL["un_number"]]
    psn = rows[0][COL["psn"]]
    haz_class = rows[0][COL["haz_class"]]
    pkg_group = rows[0][COL["pkg_group"]]
    marine = rows[0][COL["marine_pollutant"]]
    ems = rows[0][COL["ems"]]

    gr_wt = sum(r[COL["gross_wt_with_pellet_kg"]] or 0 for r in rows)
    nt_wt = sum(r[COL["net_wt_kg"]] or 0 for r in rows)
    outer_count = sum(r[COL["no_of_boxes_drums"]] or 0 for r in rows)
    inner_count = sum(r[COL["no_in_tins"]] or 0 for r in rows)

    tech_names = []
    for r in rows:
        name = r[COL["technical_name"]]
        if name and name not in tech_names:
            tech_names.append(str(name))
    technical_name = " , ".join(tech_names)

    pkg_codes = []
    for r in rows:
        code = r[COL["un_cert_no_1"]]
        if code and code not in pkg_codes:
            pkg_codes.append(str(code))
    packing_code = " , ".join(pkg_codes)

    pkg_prefix = packing_prefix(rows[0][COL["un_cert_no_1"]])
    unit_word = packing_unit_word(pkg_prefix)

    min_row = min(rows, key=lambda r: parse_flash_point(r[COL["flash_point"]]) or float("inf"))
    flash_point_display = min_row[COL["flash_point"]]

    return {
        "un_no": un_no, "psn": psn, "technical_name": technical_name,
        "class": haz_class, "pkg_group": pkg_group,
        "outer_packages": f"{outer_count} {unit_word}",
        "inner_packing": f"{inner_count} TINS",
        "packing_code": packing_code, "ems": ems,
        "flash_point": flash_point_display, "marine_pollutant": marine,
        "gr_wt": f"{gr_wt:.2f} KGS", "nt_wt": f"{nt_wt:.2f} KGS",
    }, [r[COL["material_desc"]] for r in rows]


def fill_header(ws, header):
    for i, val in enumerate(header["shipper_lines"]):
        ws[HEADER_CELLS["shipper_lines"][i]] = val
    ws[HEADER_CELLS["emergency_contact"]] = header["emergency_contact"]
    for i, val in enumerate(header["consignee_lines"]):
        ws[HEADER_CELLS["consignee_lines"][i]] = val
    ws[HEADER_CELLS["carrier"]] = header["carrier"]
    ws[HEADER_CELLS["packer_signatory_company"]] = header["packer_signatory_company"]
    ws[HEADER_CELLS["packer_place_date"]] = header["packer_place_date"]
    ws[HEADER_CELLS["packer_signatory_name"]] = header["packer_signatory_name"]
    ws[HEADER_CELLS["port_of_loading"]] = header["port_of_loading"]
    ws[HEADER_CELLS["port_of_unloading"]] = header["port_of_unloading"]
    ws[HEADER_CELLS["final_place_of_delivery"]] = header["final_place_of_delivery"]
    ws[HEADER_CELLS["shipper_decl_company"]] = header["shipper_decl_company"]
    ws[HEADER_CELLS["shipper_decl_place_date"]] = header["shipper_decl_place_date"]
    ws[HEADER_CELLS["shipper_decl_signatory_name"]] = header["shipper_decl_signatory_name"]


def fill_group_fields(ws, data):
    for key, cell in CELLS.items():
        ws[cell] = data[key]


# --------------------------- STREAMLIT UI ---------------------------------

st.set_page_config(page_title="DGD Auto-Generator", page_icon="📦")
st.title("📦 DGD Auto-Generator")
st.write("Upload your Annexure Excel and DGD Template — get generated DGD file(s) back.")

annexure_file = st.file_uploader("1. Annexure Excel file", type=["xlsx"])
template_file = st.file_uploader("2. DGD Template file", type=["xlsx"])

with st.expander("Shipment header details (edit if this shipment differs from the default)"):
    header = dict(DEFAULT_HEADER)
    header["shipper_lines"][0] = st.text_input("Shipper - line 1", header["shipper_lines"][0])
    header["consignee_lines"][0] = st.text_input("Consignee - line 1", header["consignee_lines"][0])
    header["carrier"] = st.text_input("Carrier", header["carrier"])
    header["port_of_loading"] = st.text_input("Port of Loading", header["port_of_loading"])
    header["port_of_unloading"] = st.text_input("Port of Unloading", header["port_of_unloading"])
    header["packer_place_date"] = st.text_input("Place and Date", header["packer_place_date"])
    header["packer_signatory_name"] = st.text_input("Signatory Name", header["packer_signatory_name"])

if annexure_file and template_file:
    if st.button("Generate DGD(s)", type="primary"):
        annexure_bytes = annexure_file.read()
        template_bytes = template_file.read()

        rows = load_annexure_rows(annexure_bytes)
        groups = group_rows(rows)

        st.success(f"Loaded {len(rows)} line item(s), grouped into {len(groups)} DGD(s).")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for idx, (key, group_rows_list) in enumerate(groups.items(), start=1):
                flash_bucket, psn_cat, pkg_prefix, marine, un_no = key
                data, materials = aggregate_group(group_rows_list)

                with st.expander(f"DGD #{idx}: UN{un_no} | {psn_cat} | {pkg_prefix} | Marine={marine}"):
                    for m in materials:
                        st.write(f"- {m}")
                    st.json(data)

                wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
                ws = wb.active
                fill_header(ws, header)
                fill_group_fields(ws, data)

                out_buffer = io.BytesIO()
                wb.save(out_buffer)
                fname = f"DGD_{idx}_UN{un_no}_{pkg_prefix.replace('/', '-')}.xlsx"
                zf.writestr(fname, out_buffer.getvalue())

        st.download_button(
            "⬇️ Download all generated DGD(s) (.zip)",
            data=zip_buffer.getvalue(),
            file_name="Generated_DGDs.zip",
            mime="application/zip",
        )
else:
    st.info("Upload both files above to get started.")
