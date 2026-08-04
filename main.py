#!/usr/bin/env python3
"""
main.py - FastAPI service boc quanh build_bao_gia.py

Dung cho n8n Cloud: vi n8n Cloud khong cho chay Execute Command / khong co
quyen truy cap file he thong, service nay chay tren mot server rieng (Render,
Railway, Fly.io...), nhan JSON qua HTTP, tu tai anh san pham (image_url), dien
file .xlsm, convert sang PDF bang LibreOffice, roi tra PDF ve thang cho n8n.

Endpoint:
    POST /build-quote
    Body: JSON giong config cua build_bao_gia.py, nhung moi san pham dung
          "image_url" thay vi "image_path" (server se tu tai ve):
    {
      "customer_name": "...",
      "phone": "...",
      "company": "",
      "tax_address": "",
      "delivery_address": "...",
      "mst": "",
      "email": "",
      "shipping_fee": 0,
      "format": "pdf",   // hoac "xlsm" neu muon lay file Excel goc, hoac "both"
                          // de lay CA HAI (tra ve JSON co pdf_base64/xlsm_base64)
      "products": [
        {"code": "X100", "description": "...", "color": "", "qty": 5,
         "unit_price": 1500000, "image_url": "https://..."}
      ]
    }
    Response: file nhi phan (PDF hoac xlsm) tra thang trong body, hoac JSON
    (neu format="both") gom pdf_base64 + xlsm_base64.

    GET /health -> {"status": "ok"}
"""
import base64
import os
import re
import subprocess
import unicodedata
import uuid

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, Response

from build_bao_gia import build

app = FastAPI(title="Bao Gia Hoa Phat API")

TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", "/app/templates/BBG_mau.xlsm")


def ascii_filename(name: str) -> str:
    """Bo dau tieng Viet va ky tu khong phai ASCII de dung an toan trong
    HTTP header Content-Disposition (header chi cho phep latin-1)."""
    normalized = unicodedata.normalize("NFKD", name)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_only = ascii_only.replace("đ", "d").replace("Đ", "D")
    ascii_only = re.sub(r"\s+", "_", ascii_only.strip())
    ascii_only = re.sub(r"[^A-Za-z0-9_\-]", "", ascii_only)
    return ascii_only or "KhachHang"


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/build-quote")
async def build_quote(request: Request):
    payload = await request.json()

    products = payload.get("products") or []
    if not products:
        raise HTTPException(status_code=400, detail="products rong - can it nhat 1 san pham")
    if not payload.get("customer_name"):
        raise HTTPException(status_code=400, detail="thieu customer_name")
    if not payload.get("delivery_address"):
        raise HTTPException(status_code=400, detail="thieu delivery_address")

    request_id = uuid.uuid4().hex[:10]
    work_dir = f"/tmp/bao_gia_{request_id}"
    image_dir = os.path.join(work_dir, "images")
    os.makedirs(image_dir, exist_ok=True)

    config = {
        "template_path": TEMPLATE_PATH,
        "output_path": os.path.join(work_dir, "bao_gia.xlsm"),
        "customer_name": payload.get("customer_name", ""),
        "phone": payload.get("phone", ""),
        "company": payload.get("company", ""),
        "tax_address": payload.get("tax_address", ""),
        "delivery_address": payload.get("delivery_address", ""),
        "mst": payload.get("mst", ""),
        "email": payload.get("email", ""),
        "shipping_fee": payload.get("shipping_fee", 0),
        "discount_percent": payload.get("discount_percent", 0),
        "products": products,
    }

    try:
        xlsm_path = build(config, tmp_dir=image_dir)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Loi khi tao file bao gia: {e}")

    fmt = (payload.get("format") or "pdf").lower()
    safe_name = ascii_filename(payload["customer_name"])

    if fmt == "xlsm":
        with open(xlsm_path, "rb") as f:
            data = f.read()
        return Response(
            content=data,
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            headers={"Content-Disposition": f'attachment; filename="BG_{safe_name}.xlsm"'},
        )

    # fmt == "pdf" hoac fmt == "both" deu can convert PDF (giu nguyen hanh vi cu
    # cho "pdf"; "both" tra ve CA HAI file trong 1 response JSON de n8n khong
    # can goi API 2 lan hay dung them node Merge).
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", work_dir, xlsm_path],
        capture_output=True,
        text=True,
        timeout=120,
    )
    pdf_path = xlsm_path.rsplit(".", 1)[0] + ".pdf"
    if not os.path.exists(pdf_path):
        raise HTTPException(
            status_code=500,
            detail=f"Loi convert PDF (LibreOffice): {result.stderr[:500]}",
        )

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    if fmt == "both":
        with open(xlsm_path, "rb") as f:
            xlsm_bytes = f.read()
        return JSONResponse(content={
            "customer_name": payload.get("customer_name", ""),
            "chat_id": payload.get("chat_id"),
            "pdf_filename": f"BG_{safe_name}.pdf",
            "pdf_base64": base64.b64encode(pdf_bytes).decode("ascii"),
            "xlsm_filename": f"BG_{safe_name}.xlsm",
            "xlsm_base64": base64.b64encode(xlsm_bytes).decode("ascii"),
        })

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="BG_{safe_name}.pdf"'},
    )
