from fastapi import FastAPI,UploadFile,File,HTTPException
from fastapi.responses import FileResponse
import subprocess
from datetime import datetime
from pathlib import Path
import zipfile
import requests
import msal

app = FastAPI()

from pathlib import Path

# Poppler（後で対応するので一旦そのままでOK）
POPPLER = "pdftocairo"

# ベースディレクトリ（Azure内の作業場所）
BASE = Path("/tmp/pdf2png_api_runtime")

# フォルダ作成
IN_DIR = BASE / "in"
OUT_DIR = BASE / "out"
ZIP_DIR = BASE / "zip"

for d in (IN_DIR, OUT_DIR, ZIP_DIR):
    d.mkdir(parents=True, exist_ok=True)

@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    if not (file.filename or "").lower().endswith(".pdf"):
        raise HTTPException(status_code=400,detail="PDFを選んでください")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 入力PDF保存先
    input_pdf = IN_DIR / f"input_{stamp}.pdf"
    with open(input_pdf,"wb") as f:
        f.write(await file.read())

    # 出力フォルダ（この変換専用）
    out_sub = OUT_DIR / f"output_{stamp}"
    out_sub.mkdir(parents=True,exist_ok=True)

    # pdftocairoは「prefix-1.png」「prefix-2.png…」の連番を作る
    # https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?
    prefix = out_sub / "page"
    cmd = [POPPLER,"-png","-r","300",str(input_pdf),str(prefix)]
    r = subprocess.run(cmd,capture_output=True,text=True)
    if r.returncode != 0:
        raise HTTPException(status_code=500,detail=r.stderr)

    # ZIP化
    zip_path = ZIP_DIR / f"png_{stamp}.zip"
    with zipfile.ZipFile(zip_path,"w",compression=zipfile.ZIP_DEFLATED) as z:
        for p in sorted(out_sub.glob("*.png")):
            z.write(p,arcname=p.name)

    print("★ upload 呼ぶ直前")
    upload_to_sharepoint(zip_path, zip_path.name)
    print("★ upload 呼んだ後")

    # ZIPを返す（FileResponseの例はFastAPI解説にあり）[1](https://qiita.com/phyblas/items/d94bf93606806027315a)[2](https://self-methods.com/fastapi-file/)
    return FileResponse(path=str(zip_path),filename=zip_path.name,media_type="application/zip")

import os

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

if CLIENT_SECRET is None:
    raise Exception("CLIENT_SECRETが設定されていません")

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    res = requests.post(url,data=data)
    return res.json()["access_token"]

def upload_to_sharepoint(file_path, file_name):
    token = get_access_token()

    headers_bin = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }
    headers_json = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    # ① drive_id 取得（ここが今回のキー）
    drive_meta_url = "https://graph.microsoft.com/v1.0/sites/aillz.sharepoint.com:/sites/Aillz:/drive"
    r1 = requests.get(drive_meta_url, headers=headers_json)

    print("drive取得:", r1.status_code, r1.text)

    if r1.status_code != 200:
        raise Exception("drive取得失敗")

    drive_id = r1.json()["id"]
    print("drive_id:", drive_id)

    # ② アップロード
    upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/AI新規事業/pdf_png/{file_name}:/content"

    with open(file_path, "rb") as f:
        r2 = requests.put(upload_url, headers=headers_bin, data=f)

    print("upload status:", r2.status_code)
    print("upload body:", r2.text)

    if r2.status_code not in (200, 201):
        raise Exception("upload失敗")