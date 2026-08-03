#!/usr/bin/env python3
"""
build_bao_gia.py

Dien du lieu vao file mau BBG_mau.xlsm de tao bao gia noi that Hoa Phat / The One.
Dong goi toan bo logic da kiem chung (chen dong khi >3 san pham, dich merged cells
dung cach, cong thuc tong ket, resize/anchor anh san pham) de goi tu node
"Execute Command" trong n8n hoac tu API (main.py).

Cach dung:
    python3 build_bao_gia.py --config config.json

Ghi chu:
- Neu san pham co "image_url" thay vi "image_path", script se tu tai anh ve.
- Neu KHONG co anh (ca image_path lan image_url deu thieu, hoac tai anh loi),
  script se BO QUA anh cho dong san pham do va van tiep tuc tao bao gia binh
  thuong (KHONG dung sys.exit/crash toan bo tien trinh nua - fix quan trong,
  vi truoc day 1 san pham thieu anh se lam sap toan bo API server).
- KHONG duoc sua doi cong thuc tren cac o khac ngoai nhung o duoc liet ke trong
  script nay.
"""
import argparse
import json
import os
import sys
import zipfile
from copy import copy

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import range_boundaries, get_column_letter

SHEET_NAME = "BG mau"
FIRST_PRODUCT_ROW = 18
DEFAULT_TEMPLATE_ROWS = 3
DEFAULT_TOTAL_START_ROW = 21
IMG_WIDTH = 140
IMG_HEIGHT = 146
ROW_HEIGHT_WITH_IMAGE = 170


def die(msg):
    print(f"LOI: {msg}", file=sys.stderr)
    sys.exit(1)


def maybe_download_image(product, tmp_dir):
    if product.get("image_path"):
        return product["image_path"]
    url = product.get("image_url")
    if not url:
        print(
            f"CANH BAO: san pham {product.get('code')} khong co image_path lan "
            f"image_url - bo qua anh cho dong nay, van tiep tuc tao bao gia.",
            file=sys.stderr,
        )
        return None
    import requests
    os.makedirs(tmp_dir, exist_ok=True)
    fname = os.path.join(tmp_dir, f"{product['code'].replace(' ', '_').replace('/', '_')}.png")
    try:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        with open(fname, "wb") as f:
            f.write(resp.content)
        return fname
    except Exception as e:
        print(
            f"CANH BAO: khong tai duoc anh cho san pham {product.get('code')} "
            f"tu {url} ({e}) - bo qua anh cho dong nay, van tiep tuc tao bao gia.",
            file=sys.stderr,
        )
        return None


def unmerge_all(ws):
    ranges = [str(mc) for mc in ws.merged_cells.ranges]
    for r in ranges:
        ws.unmerge_cells(r)
    return ranges


def remerge_shifted(ws, old_ranges, insert_at, n_insert):
    for r in old_ranges:
        min_col, min_row, max_col, max_row = range_boundaries(r)
        if min_row >= insert_at:
            min_row += n_insert
            max_row += n_insert
        new_range = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        ws.merge_cells(new_range)


def build(config, tmp_dir="/tmp/bao_gia_images"):
    template_path = config["template_path"]
    output_path = config["output_path"]
    products = config["products"]
    n_products = len(products)

    if n_products == 0:
        die("Danh sach san pham rong - khong the tao bao gia")

    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    ws = wb[SHEET_NAME]

    phone = config.get("phone", "")
    name_line = f"Kính gửi : {config.get('customer_name', '')}"
    if phone:
        name_line += f" - {phone}"
    ws["A10"] = name_line
    if config.get("company"):
        ws["A11"] = f"Công Ty: {config['company']}"
    if config.get("tax_address"):
        ws["A12"] = f"Địa chỉ thuế: {config['tax_address']}"
    ws["A13"] = f"Địa chỉ giao hàng: {config.get('delivery_address', '')}"
    if config.get("mst"):
        ws["A14"] = f"MST: {config['mst']}"
    email = config.get("email", "")
    ws["A15"] = f"Điện thoại: {phone}" + " " * 60 + f"Email: {email}"

    insert_at = FIRST_PRODUCT_ROW + DEFAULT_TEMPLATE_ROWS

    if n_products > DEFAULT_TEMPLATE_ROWS:
        n_extra = n_products - DEFAULT_TEMPLATE_ROWS
        old_ranges = unmerge_all(ws)
        ws.insert_rows(insert_at, n_extra)
        remerge_shifted(ws, old_ranges, insert_at, n_extra)
        last_existing_row = FIRST_PRODUCT_ROW + DEFAULT_TEMPLATE_ROWS - 1
        for new_row in range(insert_at, insert_at + n_extra):
            ws.row_dimensions[new_row].height = ws.row_dimensions[last_existing_row].height
            for c in range(1, 10):
                src = ws.cell(row=last_existing_row, column=c)
                dst = ws.cell(row=new_row, column=c)
                dst._style = copy(src._style)

    for idx, product in enumerate(products, start=1):
        row = FIRST_PRODUCT_ROW + idx - 1
        ws.row_dimensions[row].height = ROW_HEIGHT_WITH_IMAGE
        ws[f"A{row}"] = idx
        ws[f"B{row}"] = product["code"]
        ws[f"D{row}"] = product.get("description", "")
        ws[f"E{row}"] = product.get("color", "")
        ws[f"F{row}"] = "cái"
        ws[f"G{row}"] = product["qty"]
        ws[f"H{row}"] = product["unit_price"]
        ws[f"I{row}"] = f"=H{row}*G{row}"

        img_path = maybe_download_image(product, tmp_dir)
        if img_path:
            img = XLImage(img_path)
            img.width = IMG_WIDTH
            img.height = IMG_HEIGHT
            ws.add_image(img, f"C{row}")

    if n_products < DEFAULT_TEMPLATE_ROWS:
        for row in range(FIRST_PRODUCT_ROW + n_products, FIRST_PRODUCT_ROW + DEFAULT_TEMPLATE_ROWS):
            for col in ("A", "B", "D", "E", "F", "G", "H"):
                ws[f"{col}{row}"] = None

    shipping_fee = config.get("shipping_fee", 0)
    if n_products <= DEFAULT_TEMPLATE_ROWS:
        ws["I24"] = shipping_fee
    else:
        total_start = FIRST_PRODUCT_ROW + n_products
        last_product_row = FIRST_PRODUCT_ROW + n_products - 1
        r_tong_cong = total_start
        r_chiet_khau = total_start + 1
        r_gia_tri_con_lai = total_start + 2
        r_phi_vc = total_start + 3
        r_tong_cong_vc = total_start + 4
        r_vat = total_start + 5
        r_tong_tt_cuoi = total_start + 6
        r_tam_ung = total_start + 7
        r_con_lai = total_start + 8

        ws[f"I{r_tong_cong}"] = f"=SUM(I{FIRST_PRODUCT_ROW}:I{last_product_row})"
        ws[f"I{r_gia_tri_con_lai}"] = f"=I{r_tong_cong}-I{r_chiet_khau}"
        ws[f"I{r_phi_vc}"] = shipping_fee
        ws[f"I{r_tong_cong_vc}"] = f"=I{r_gia_tri_con_lai}+I{r_phi_vc}"
        ws[f"I{r_vat}"] = f"=I{r_tong_cong_vc}*8%"
        ws[f"I{r_tong_tt_cuoi}"] = f"=I{r_tong_cong_vc}+I{r_vat}"
        ws[f"I{r_tam_ung}"] = f"=I{r_tong_tt_cuoi}*50%"
        ws[f"I{r_con_lai}"] = f"=I{r_tong_tt_cuoi}-I{r_tam_ung}"

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)

    n_media = len([n for n in zipfile.ZipFile(output_path).namelist() if "media" in n])
    print(f"Da luu: {output_path}")
    print(f"So luong anh (media) trong file: {n_media}")
    if n_media < 29:
        print("CANH BAO: so luong anh thap bat thuong!", file=sys.stderr)

    return output_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True)
    ap.add_argument("--tmp-dir", default="/tmp/bao_gia_images")
    args = ap.parse_args()

    with open(args.config, "r", encoding="utf-8") as f:
        config = json.load(f)

    build(config, tmp_dir=args.tmp_dir)


if __name__ == "__main__":
    main()
