from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import fitz
import pdfplumber
import pandas as pd
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches
import subprocess
import os
import uuid

app = FastAPI(title="DocuSwift Pro Conversion Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "temp_workspace"
os.makedirs(TEMP_DIR, exist_ok=True)

def create_temp_file(ext: str) -> str:
    return os.path.join(TEMP_DIR, f"{uuid.uuid4()}.{ext}")

def cleanup_files(*paths):
    for path in paths:
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except Exception:
            pass

@app.get("/")
def health_check():
    return {"status": "online", "engine": "DocuSwift Pro Python API"}

@app.post("/api/compress")
async def compress_pdf(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    level: str = Form("medium")
):
    in_path = create_temp_file("pdf")
    out_path = create_temp_file("pdf")
    
    try:
        with open(in_path, "wb") as f:
            f.write(await file.read())
            
        doc = fitz.open(in_path)
        quality = 45 if level == "high" else (68 if level == "medium" else 85)
        
        for page in doc:
            for img in page.get_images(full=True):
                xref = img[0]
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n - pix.alpha > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    img_bytes = pix.tobytes("jpeg", jpg_quality=quality)
                    page.replace_image(xref, stream=img_bytes)
                except Exception:
                    continue
                    
        doc.save(out_path, garbage=4, deflate=True, clean=True)
        doc.close()
        
        background_tasks.add_task(cleanup_files, in_path, out_path)
        return FileResponse(out_path, filename=f"compressed_{file.filename}", media_type="application/pdf")
    except Exception as e:
        cleanup_files(in_path, out_path)
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/pdf-to-excel")
async def pdf_to_excel(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    in_path = create_temp_file("pdf")
    out_path = create_temp_file("xlsx")
    
    try:
        with open(in_path, "wb") as f:
            f.write(await file.read())
            
        with pdfplumber.open(in_path) as pdf:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                has_tables = False
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        if table and len(table) > 0:
                            headers = table[0] if len(table) > 1 else None
                            rows = table[1:] if len(table) > 1 else table
                            df = pd.DataFrame(rows, columns=headers)
                            df.to_excel(writer, sheet_name=f"P{i+1}_Table{j+1}", index=False)
                            has_tables = True
                if not has_tables:
                    lines = []
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text() or ""
                        for line in text.split("\n"):
                            if line.strip():
                                lines.append({"Page": i + 1, "Content": line.strip()})
                    if not lines:
                        lines = [{"Page": 1, "Content": "No extractable text found"}]
                    pd.DataFrame(lines).to_excel(writer, sheet_name="Extracted_Data", index=False)
                    
        base_name = os.path.splitext(file.filename)[0]
        background_tasks.add_task(cleanup_files, in_path, out_path)
        return FileResponse(out_path, filename=f"{base_name}.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        cleanup_files(in_path, out_path)
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/pdf-to-word")
async def pdf_to_word(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    in_path = create_temp_file("pdf")
    out_path = create_temp_file("docx")
    
    try:
        with open(in_path, "wb") as f:
            f.write(await file.read())
            
        cv = Converter(in_path)
        cv.convert(out_path)
        cv.close()
        
        base_name = os.path.splitext(file.filename)[0]
        background_tasks.add_task(cleanup_files, in_path, out_path)
        return FileResponse(out_path, filename=f"{base_name}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        cleanup_files(in_path, out_path)
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/pdf-to-ppt")
async def pdf_to_ppt(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    in_path = create_temp_file("pdf")
    out_path = create_temp_file("pptx")
    img_paths = []
    
    try:
        with open(in_path, "wb") as f:
            f.write(await file.read())
            
        doc = fitz.open(in_path)
        prs = Presentation()
        blank_layout = prs.slide_layouts[6]
        
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img_p = create_temp_file("png")
            pix.save(img_p)
            img_paths.append(img_p)
            
            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(img_p, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)
            
        doc.close()
        prs.save(out_path)
        
        base_name = os.path.splitext(file.filename)[0]
        background_tasks.add_task(cleanup_files, in_path, out_path, *img_paths)
        return FileResponse(out_path, filename=f"{base_name}.pptx", media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        cleanup_files(in_path, out_path, *img_paths)
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/office-to-pdf")
async def office_to_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    ext = file.filename.split(".")[-1].lower()
    in_path = create_temp_file(ext)
    out_path = in_path.rsplit(".", 1)[0] + ".pdf"
    
    try:
        with open(in_path, "wb") as f:
            f.write(await file.read())
            
        subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", in_path, "--outdir", TEMP_DIR],
            check=True
        )
        
        base_name = os.path.splitext(file.filename)[0]
        background_tasks.add_task(cleanup_files, in_path, out_path)
        return FileResponse(out_path, filename=f"{base_name}.pdf", media_type="application/pdf")
    except Exception as e:
        cleanup_files(in_path, out_path)
        raise HTTPException(status_code=500, detail=str(e))
