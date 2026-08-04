#!/usr/bin/env python3
"""
build_bao_gia.py

Dien du lieu vao file mau BBG_mau.xlsm de tao bao gia noi that Hoa Phat / The One.
Dong goi toan bo logic da kiem chung (chen dong khi >3 san pham, dich merged cells
dung cach, cong thuc tong ket, resize/anchor anh san pham) de goi tu node
"Execute Command" trong n8n hoac tu API (main.py).

Cach dung:
    python3 build_bao_gia.py --config config.json

File config.json (vi du):
{
  "template_path": "/data/templates/BBG_mau.xlsm",
  "output_path": "/data/output/BG_ChiLan.xlsm",
  "customer_name": "Chị Lan",
  "phone": "0901234567",
  "company": "",
  "tax_address": "",
  "delivery_address": "123 Nguyễn Văn A, Q.1, TP.HCM",
  "mst": "",
  "email": "",
  "shipping_fee": 0,
  "products": [
    {
      "code": "X100",
      "description": "Mo ta san pham lay tu hoaphatsaigon.com",
      "color": "",
      "qty": 5,
      "unit_price": 1500000,
      "image_path": "/data/images/x100.png"
    }
  ]
}

Ghi chu:
- Neu san pham co "image_url" thay vi "image_path", script se tu tai anh ve
  (dung requests - server chay n8n thuong khong bi chan mang nhu sandbox Claude,
  nen KHONG can chup man hinh + crop nhu quy trinh thu cong).
- Neu KHONG co anh (ca image_path lan image_url deu thieu, hoac tai anh loi),
  script se BO QUA anh cho dong san pham do va van tiep tuc tao bao gia binh
  thuong (KHONG dung sys.exit/crash toan bo tien trinh nua - fix quan trong,
  vi truoc day 1 san pham thieu anh se lam sap toan bo API server).
- KHONG duoc sua doi cong thuc tren cac o khac ngoai nhung o duoc liet ke trong
  script nay - day la file .xlsm co san cong thuc VAT/tam ung, sua sai se lam
  bao gia tinh sai tien.
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
DEFAULT_TEMPLATE_ROWS = 3          # so dong san pham co san trong template (18,19,20)
DEFAULT_TOTAL_START_ROW = 21       # dong "Tong cong" mac dinh khi khong chen them dong
IMG_WIDTH = 140
IMG_HEIGHT = 146
ROW_HEIGHT_WITH_IMAGE = 170


def die(msg):
    """Chi dung cho loi nghiem trong KHONG THE tiep tuc (vi du: thieu template,
    danh sach san pham rong). KHONG dung cho loi anh thieu/tai anh loi nua."""
    print(f"LOI: {msg}", file=sys.stderr)
    sys.exit(1)


def maybe_download_image(product, tmp_dir):
    """Neu san pham co image_url thi tai ve, tra ve duong dan file cuc bo.
    Neu khong co anh (thieu du lieu tu web) hoac tai loi, tra ve None thay vi
    crash - de van tao duoc bao gia, chi thieu anh o dong do."""
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


def new_row_after_splice(old_row, threshold_row, shift):
    """Tinh vi tri dong moi sau khi chen (shift > 0) hoac xoa (shift < 0) dong
    tai threshold_row. Voi truong hop xoa, vung bi xoa la
    [threshold_row + shift, threshold_row - 1] (shift la so am) - tra ve None
    neu dong do nam trong vung bi xoa (khong con ton tai nua)."""
    if shift >= 0:
        return old_row + shift if old_row >= threshold_row else old_row
    delete_start = threshold_row + shift
    if old_row < delete_start:
        return old_row
    if old_row < threshold_row:
        return None
    return old_row + shift


def remerge_shifted(ws, old_ranges, threshold_row, shift):
    """shift duong: chen them dong. shift am: xoa bot dong."""
    for r in old_ranges:
        min_col, min_row, max_col, max_row = range_boundaries(r)
        new_min_row = new_row_after_splice(min_row, threshold_row, shift)
        new_max_row = new_row_after_splice(max_row, threshold_row, shift)
        if new_min_row is None or new_max_row is None:
            continue  # dong bi xoa, khong re-merge nua
        new_range = f"{get_column_letter(min_col)}{new_min_row}:{get_column_letter(max_col)}{new_max_row}"
        ws.merge_cells(new_range)


def shift_row_heights(ws, old_heights, threshold_row, shift):
    """openpyxl KHONG tu dong dich chuyen row_dimensions (chieu cao dong) khi
    chen/xoa dong bang insert_rows/delete_rows - chi dich chuyen gia tri o va
    KHONG dich merged cells (da xu ly rieng o remerge_shifted). Ham nay dich
    chuyen chieu cao dong tuong ung, tranh tinh trang dong "Tong cong"/"Chiet
    khau" bi ke thua nham chieu cao cao cua dong san pham (anh) sau khi xoa."""
    new_heights = {}
    for old_row, h in old_heights.items():
        new_row = new_row_after_splice(old_row, threshold_row, shift)
        if new_row is None:
            continue  # dong nay da bi xoa, bo qua
        new_heights[new_row] = h

    # Xoa het custom height hien tai (co the con sot lai gia tri cu sai vi tri)
    for row_idx in list(ws.row_dimensions.keys()):
        ws.row_dimensions[row_idx].height = None

    for row_idx, h in new_heights.items():
        ws.row_dimensions[row_idx].height = h


def build(config, tmp_dir="/tmp/bao_gia_images"):
    template_path = config["template_path"]
    output_path = config["output_path"]
    products = config["products"]
    n_products = len(products)

    if n_products == 0:
        die("Danh sach san pham rong - khong the tao bao gia")

    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    ws = wb[SHEET_NAME]

    # --- Thong tin khach hang ---
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

    boundary_row = FIRST_PRODUCT_ROW + DEFAULT_TEMPLATE_ROWS  # = 21, ranh gioi goc

    # --- Neu so san pham khac 3 (mac dinh cua template), chen hoac XOA HAN dong
    # de khong con dong san pham rong hien thi tren PDF ---
    if n_products != DEFAULT_TEMPLATE_ROWS:
        old_ranges = unmerge_all(ws)
        old_heights = {
            r: dim.height for r, dim in ws.row_dimensions.items() if dim.height is not None
        }
        if n_products > DEFAULT_TEMPLATE_ROWS:
            n_diff = n_products - DEFAULT_TEMPLATE_ROWS
            ws.insert_rows(boundary_row, n_diff)
            shift = n_diff
            # copy dinh dang tu dong san pham cuoi cung co san sang cac dong moi
            last_existing_row = FIRST_PRODUCT_ROW + DEFAULT_TEMPLATE_ROWS - 1  # = 20
            for new_row in range(boundary_row, boundary_row + n_diff):
                ws.row_dimensions[new_row].height = ws.row_dimensions[last_existing_row].height
                for c in range(1, 10):
                    src = ws.cell(row=last_existing_row, column=c)
                    dst = ws.cell(row=new_row, column=c)
                    dst._style = copy(src._style)
        else:
            # It san pham hon template -> XOA HAN cac dong du (khong chi xoa gia tri),
            # de PDF khong con hien o vien trong.
            n_diff = DEFAULT_TEMPLATE_ROWS - n_products
            delete_at = FIRST_PRODUCT_ROW + n_products
            ws.delete_rows(delete_at, n_diff)
            shift = -n_diff
        remerge_shifted(ws, old_ranges, boundary_row, shift)
        shift_row_heights(ws, old_heights, boundary_row, shift)

    # --- Dien du lieu san pham ---
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

    # --- Cong thuc khu vuc tong ket - LUON tinh dong theo vi tri thuc te sau khi
    # chen/xoa dong, khong con nhanh dac biet nua (vi gio luon xoa hoac chen dong
    # cho khop dung n_products, khong con truong hop "dong rong con lai" nua) ---
    shipping_fee = config.get("shipping_fee", 0)
    total_start = FIRST_PRODUCT_ROW + n_products  # dong "Tong cong"
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

    # --- Kiem tra khong lam mat anh goc cua template ---
    n_media = len([n for n in zipfile.ZipFile(output_path).namelist() if "media" in n])
    print(f"Da luu: {output_path}")
    print(f"So luong anh (media) trong file: {n_media} (phai >= 29 + so anh san pham tai duoc, neu template goc co 29 anh trang tri)")
    if n_media < 29:
        print("CANH BAO: so luong anh thap bat thuong - co the da mat anh goc cua template!", file=sys.stderr)

    return output_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True, help="Duong dan file JSON config")
    ap.add_argument("--tmp-dir", default="/tmp/bao_gia_images", help="Thu muc tam de tai anh san pham")
    args = ap.parse_args()

    with open(args.config, "r", encoding="utf-8") as f:
        config = json.load(f)

    build(config, tmp_dir=args.tmp_dir)


if __name__ == "__main__":
    main()
