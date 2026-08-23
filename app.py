# DGD (Dangerous Goods Declaration) Auto-Generator
# --------------------------------------------------
# Reads consignment line items from an Annexure Excel file, groups them
# according to the 5-point checkpoint logic, and produces one DGD per
# group using the company's DGD template layout.
#
# USAGE:
#     python generate_dgd.py <Annexure.xlsx> <DGD_Template.xlsx> <output_folder>
#
# Annexure sheet expected: "Filtered Data" (the shipment's actual line items).
# Template expected: single sheet with the same cell layout as MANUAL_DGD1.xlsx.


import sys
import re
import shutil
from pathlib import Path
from collections import defaultdict
import openpyxl

# ---------------------------------------------------------------------------
# EDIT THIS: shipment-level header info that is the SAME for every DGD in a
# given shipment/PO, but is NOT present in the Annexure file. Fill these in
# per shipment before running, or wire them up to another data source later.
# ---------------------------------------------------------------------------
SHIPMENT_HEADER = {
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
    "carrier_booking_number": "",
    "packer_signatory_company": "AXALTA COATING SYSTEMS INDIA PVT LTD",
    "packer_place_date": "VADODARA-10-08-2026",
    "packer_signatory_name": "Mr. Pranav Dave",
    "port_of_loading": "NHAVA SHEVA",
    "port_of_unloading": "SYDNEY , AUSTRALIA",
    "final_place_of_delivery": "SYDNEY , AUSTRALIA",
    "ship_name_voyage": "",
    "container_numbers": "",
    "shipper_decl_company": "AXALTA COATING SYSTEMS INDIA PVT LTD",
    "shipper_decl_place_date": "VADODARA-10-08-2026",
    "shipper_decl_signatory_name": "Mr. Pranav Dave",
}

# ---------------------------------------------------------------------------
# Column mapping (0-indexed) for the Annexure "Filtered Data" / "Master Data"
# sheet layout, based on the header row.
# ---------------------------------------------------------------------------
COL = {
    "material_desc": 3,
    "no_in_tins": 6,
    "net_wt_kg": 9,
    "no_of_boxes_drums": 10,
    "gross_wt_with_pellet_kg": 14,
    "un_number": 15,
    "psn": 16,
    "haz_class": 17,
    "pkg_group": 18,
    "marine_pollutant": 19,
    "flash_point": 20,
    "ems": 21,
    "un_cert_no": 24,
    "un_cert_no_1": 25,  # packing code string, e.g. "4G/Z 13/S/** IND/A/8101088"
    "technical_name": 30,
}

# Cell coordinates in the DGD template that get filled per-group
CELLS = {
    "un_no": "D36",
    "gr_wt": "H36",
    "psn": "D38",
    "nt_wt": "H38",
    "technical_name": "D40",
    "outer_packages": "A43",
    "class": "C42",
    "pkg_group": "D44",
    "inner_packing": "A45",
    "packing_code": "D46",
    "ems": "D48",
    "flash_point": "D50",
    "marine_pollutant": "D52",
}

HEADER_CELLS = {
    "shipper_lines": [
        "A9",
        "A10",
        "A11",
        "A12"
    ],
    "emergency_contact": "D15",
    "consignee_lines": [
        "A17",
        "A18",
        "A19",
        "A20"
    ],
    "carrier": "D17",
    "carrier_booking_number": "D18",  # note: template row for value may need adjusting
    "packer_signatory_company": "D22",
    "packer_place_date": "D25",
    "packer_signatory_name": "D28",
    "port_of_loading": "B31",
    "port_of_unloading": "A33",
    "final_place_of_delivery": "B33",
    "ship_name_voyage": "A30",  # label cell; actual value cell may need its own row -- verify against template
    "container_numbers": "D30",  # label cell; verify actual value placement
    "shipper_decl_company": "D55",
    "shipper_decl_place_date": "D57",
    "shipper_decl_signatory_name": "D60",
}

def parse_flash_point(text):
    """Extract numeric °C value from strings like '4.47 DEG CEL'."""
    if text is None:
        return None
    match = re.search(r"[-+]?\d*\.?\d+", str(text))
    return float(match.group()) if match else None


def packing_prefix(cert_code_str):
    """First segment of a packing-code string, e.g. '4G' or '1A1'."""
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
    """Bucket PSN into PAINT vs PAINT RELATED per checkpoint #2."""
    if psn_text is None:
        return ""
    text = str(psn_text).strip().upper()
    return "PAINT" if text == "PAINT" else "PAINT RELATED"


def load_annexure_rows(path, sheet_name="Filtered Data"):
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # skip blank rows and the TOTAL summary row
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

    # unique technical names, in order of first appearance
    tech_names = []
    for r in rows:
        name = r[COL["technical_name"]] 
        if name and name not in tech_names:
            tech_names.append(str(name))
    technical_name = " , ".join(tech_names)

    # unique packing codes, in order of first appearance
    pkg_codes = []
    for r in rows:
        code = r[COL["un_cert_no_1"]] 
        if code and code not in pkg_codes:
            pkg_codes.append(str(code))
    packing_code = " , ".join(pkg_codes)

    pkg_prefix = packing_prefix(rows[0][COL["un_cert_no_1"]])
    unit_word = packing_unit_word(pkg_prefix)

    # flash point: minimum numeric value, but keep its ORIGINAL text
    min_row = min(rows, key=lambda r: parse_flash_point(r[COL["flash_point"]]) or float("inf"))
    flash_point_display = min_row[COL["flash_point"]]

    return {
        "un_no": un_no,
        "psn": psn,
        "technical_name": technical_name,
        "class": haz_class,
        "pkg_group": pkg_group,
        "outer_packages": f"{outer_count} {unit_word}",
        "inner_packing": f"{inner_count} TINS",
        "packing_code": packing_code,
        "ems": ems,
        "flash_point": flash_point_display,
        "marine_pollutant": marine,
        "gr_wt": f"{gr_wt:.2f} KGS",
        "nt_wt": f"{nt_wt:.2f} KGS",
    }


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
    ws[HEADER_CELLS["carrier_booking_number"]] = header["carrier_booking_number"]
    ws[HEADER_CELLS["ship_name_voyage"]] = header["ship_name_voyage"]
    ws[HEADER_CELLS["container_numbers"]] = header["container_numbers"]


def fill_group_fields(ws, data):
    ws[CELLS["un_no"]] = data["un_no"]
    ws[CELLS["gr_wt"]] = data["gr_wt"]
    ws[CELLS["psn"]] = data["psn"]
    ws[CELLS["nt_wt"]] = data["nt_wt"]
    ws[CELLS["technical_name"]] = data["technical_name"]
    ws[CELLS["outer_packages"]] = data["outer_packages"]
    ws[CELLS["class"]] = data["class"]
    ws[CELLS["pkg_group"]] = data["pkg_group"]
    ws[CELLS["inner_packing"]] = data["inner_packing"]
    ws[CELLS["packing_code"]] = data["packing_code"]
    ws[CELLS["ems"]] = data["ems"]
    ws[CELLS["flash_point"]] = data["flash_point"]
    ws[CELLS["marine_pollutant"]] = data["marine_pollutant"]


def generate(annexure_path, template_path, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    rows = load_annexure_rows(annexure_path)
    groups = group_rows(rows)

    print(f"Loaded {len(rows)} line item(s) from Annexure.")
    print(f"Grouped into {len(groups)} DGD(s):\n")

    generated_files = []
    for idx, (key, group_rows_list) in enumerate(groups.items(), start=1):
        flash_bucket, psn_cat, pkg_prefix, marine, un_no = key
        data = aggregate_group(group_rows_list)

        materials = [r[COL["material_desc"]] for r in group_rows_list]
        print(f"  DGD #{idx}: UN{un_no} | {psn_cat} | {pkg_prefix} | Marine={marine} | FlashPt={flash_bucket}")
        for m in materials:
            print(f"      - {m}")

        out_path = output_dir / f"DGD_{idx}_UN{un_no}_{pkg_prefix.replace('/', '-')}.xlsx"
        shutil.copy(template_path, out_path)

        wb = openpyxl.load_workbook(out_path)
        ws = wb.active
        fill_header(ws, SHIPMENT_HEADER)
        fill_group_fields(ws, data)
        wb.save(out_path)
        generated_files.append(out_path)

    print(f"\nDone. {len(generated_files)} DGD file(s) written to: {output_dir}")
    return generated_files

# Removed the __main__ block to prevent SystemExit when running in Colab.
# You can now call the generate function directly with appropriate arguments.
# Example: generate("path/to/Annexure.xlsx", "path/to/DGD_Template.xlsx", "path/to/output_folder")
