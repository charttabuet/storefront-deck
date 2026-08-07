# -*- coding: utf-8 -*-
"""
service.py  ---  โหมด service (รันบน cloud ให้ n8n เรียกผ่าน HTTP)
POST /generate   body JSON:
{
  "campaign": "แคมเปญลดราคากลางปี 2569",
  "month": "กรกฎาคม 2569",
  "branches": [
     {"name": "01 สาขาเซ็นทรัลลาดพร้าว", "note": "", "photos": ["https://...jpg", ...]},
     ...
  ]
}
คืนไฟล์ .pptx กลับไป (binary)  — n8n เอาไปแนบอีเมล/อัปโหลด Drive ต่อได้เลย

ความปลอดภัย: ถ้าตั้ง env API_KEY ไว้ ต้องส่ง header  X-API-KEY  ให้ตรงกัน
"""
import os
import time
import tempfile
import requests
from flask import Flask, request, jsonify, send_file
from deck_core import build_deck

app = Flask(__name__)
HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "template.pptx")
API_KEY = os.environ.get("API_KEY", "")          # ตั้งใน cloud เพื่อกันคนอื่นยิง
MAX_PHOTOS = int(os.environ.get("MAX_PHOTOS", "400"))
DL_TRIES = int(os.environ.get("DL_TRIES", "4"))  # ลองโหลดรูปซ้ำกี่ครั้งถ้าพลาด
# บาง endpoint (เช่น googleusercontent) เมิน request ที่ไม่มี User-Agent
_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")


def _download(url, dest):
    """โหลดรูป พร้อม retry + backoff กันรูปหายเวลา Drive ตอบช้า/ติด rate-limit"""
    last = None
    for attempt in range(DL_TRIES):
        try:
            r = requests.get(url, timeout=45, stream=True,
                             headers={"User-Agent": _UA})
            # 429/5xx = ชั่วคราว -> ลองใหม่;  4xx อื่น ๆ = ถาวร -> เลิก
            if r.status_code in (429, 500, 502, 503, 504):
                raise requests.HTTPError(f"transient {r.status_code}")
            r.raise_for_status()
            with open(dest, "wb") as f:
                for chunk in r.iter_content(8192):
                    f.write(chunk)
            if os.path.getsize(dest) > 0:
                return
            raise IOError("empty file")
        except Exception as e:  # noqa: BLE001
            last = e
            if attempt < DL_TRIES - 1:
                time.sleep(1.5 * (attempt + 1))   # 1.5s, 3s, 4.5s ...
    raise last


@app.get("/health")
def health():
    return jsonify(status="ok")


@app.post("/generate")
def generate():
    if API_KEY and request.headers.get("X-API-KEY", "") != API_KEY:
        return jsonify(error="unauthorized"), 401

    data = request.get_json(force=True, silent=True) or {}
    campaign = data.get("campaign", "แคมเปญหน้าร้าน")
    month = data.get("month", "")
    kicker = data.get("kicker", "")          # ข้อความบรรทัดเล็กบนหน้าปก
    branches_in = data.get("branches", [])
    if not branches_in:
        return jsonify(error="no branches"), 400

    with tempfile.TemporaryDirectory() as tmp:
        branches = []
        count = 0
        for bi, b in enumerate(branches_in):
            local = []
            for pi, url in enumerate(b.get("photos", [])):
                if count >= MAX_PHOTOS:
                    break
                ext = ".png" if ".png" in url.lower() else ".jpg"
                dest = os.path.join(tmp, f"b{bi}_{pi}{ext}")
                try:
                    _download(url, dest)
                    local.append(dest)
                    count += 1
                except Exception as e:
                    app.logger.warning("download failed %s: %s", url, e)
            branches.append({"name": b.get("name", f"สาขา {bi+1}"),
                             "note": b.get("note", ""),
                             "photos": local})

        safe = "".join(ch if ch.isalnum() or ch in " _-" else "_"
                       for ch in f"{campaign} {month}").strip() or "deck"
        out = os.path.join(tmp, safe + ".pptx")
        total, missing = build_deck(branches, campaign, month, TEMPLATE, out,
                                    kicker=kicker)
        app.logger.info("built deck: %d photos, missing=%s", total, missing)
        return send_file(
            out, as_attachment=True,
            download_name=safe + ".pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8080")))
