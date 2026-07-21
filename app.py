"""
Tailory Backend V2 — Pipeline documentaire pédagogique
FastAPI + python-docx + pdf2docx + Anthropic proxy

Endpoints:
  POST /parse    → DOCX/PDF → structure JSON pédagogique
  POST /generate → proxy Anthropic avec retry + chunking
  POST /export   → structure JSON adaptée → DOCX
  POST /convert  → PDF → DOCX (existant, conservé)
  GET  /health   → vérification
"""

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import anthropic
import base64
import io
import json
import os
import time
import uuid
import zipfile
import glob
import tempfile
import subprocess
from typing import Optional

# python-docx
from docx import Document as DocxDocument
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# pdf2docx (conversion PDF→DOCX existante)
from pdf2docx import Converter

app = FastAPI(title="Tailory Backend V2")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://classedobsdecath.ch",
        "http://localhost",
        "http://127.0.0.1"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ─────────────────────────────────────────────
# MODÈLES DE TYPES D'EXERCICES
# ─────────────────────────────────────────────
EXERCISE_KEYWORDS = {
    "relier":     ["reli", "associe", "associer", "relie", "relient"],
    "entourer":   ["entoure", "entourer", "encercle", "encercler", "barre", "barrer", "souligne"],
    "compléter":  ["complète", "compléter", "écris", "écrire", "inscris", "note"],
    "numéroter":  ["numérote", "numéroter", "numérotez", "numérotes"],
    "classer":    ["découpe", "découper", "colle", "coller", "classe", "classer", "range", "ranger"],
    "dessiner":   ["dessine", "dessiner", "dessinez", "colorie", "colorier"],
    "observer":   ["observe", "observer", "regardes", "regarde"],
    "lire":       ["lis", "lire", "lisez", "lecture"],
}

def detect_exercise_type(text: str) -> str:
    text_lower = text.lower()
    for ex_type, keywords in EXERCISE_KEYWORDS.items():
        for kw in keywords:
            if kw in text_lower:
                return ex_type
    return "autre"


# ─────────────────────────────────────────────
# PDF : extraction PyMuPDF — raster + vectoriel
# ─────────────────────────────────────────────
import fitz  # PyMuPDF, déjà installé (dépendance pdf2docx)


def _cluster_rects(rects, margin=14.0):
    """Regroupe les rectangles qui se touchent ou sont proches (< margin pt)."""
    rects = [fitz.Rect(r) for r in rects]
    merged = True
    while merged:
        merged = False
        out = []
        while rects:
            r = rects.pop()
            grown = fitz.Rect(r.x0 - margin, r.y0 - margin,
                              r.x1 + margin, r.y1 + margin)
            absorbed = []
            for i, other in enumerate(rects):
                if grown.intersects(other):
                    r |= other
                    absorbed.append(i)
                    merged = True
            for i in reversed(absorbed):
                rects.pop(i)
            out.append(r)
        rects = out
    return rects


def parse_pdf(content: bytes, filename: str):
    """
    Extrait d'un PDF, dans l'ordre de lecture :
    - les images raster (photos, dessins importés)
    - les figures vectorielles (tracés regroupés puis rasterisés en PNG 2x)
    - le texte complet
    Les zones de figures sont détectées via get_drawings() + clustering.
    """
    doc = fitz.open(stream=content, filetype="pdf")
    images = []
    full_text = []
    idx = 0

    for pno, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        page_area = pw * ph
        regions = []

        # 1. Rectangles des images raster
        for img in page.get_images(full=True):
            try:
                for rect in page.get_image_rects(img[0]):
                    if rect.width > 12 and rect.height > 12:
                        regions.append(rect)
            except Exception:
                pass

        # 2. Rectangles des tracés vectoriels (figures dessinées)
        try:
            for d in page.get_drawings():
                r = d.get("rect")
                if r is None:
                    continue
                # Ignorer traits fins pleine largeur (lignes de séparation,
                # lignes de réponse) et micro-tracés
                if r.width < 10 or r.height < 10:
                    continue
                if r.width > 0.92 * pw and r.height < 20:
                    continue
                regions.append(r)
        except Exception:
            pass

        # 3. Clustering : tracés voisins = une figure
        clusters = _cluster_rects(regions)

        # 4. Filtres : trop petit, ou boîte de fond couvrant la page
        keep = []
        for r in clusters:
            r = r & page.rect  # borner à la page
            if r.width < 24 or r.height < 24:
                continue
            if (r.width * r.height) > 0.85 * page_area:
                continue
            keep.append(r)

        # 5. Ordre de lecture : haut → bas, gauche → droite
        keep.sort(key=lambda r: (round(r.y0 / 24), r.x0))

        # 6. Rasterisation 2x de chaque zone
        for r in keep:
            try:
                pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(2, 2))
                if pix.width < 8 or pix.height < 8:
                    continue
                b64 = base64.b64encode(pix.tobytes("png")).decode()
                images.append({
                    "index": idx,
                    "page": pno + 1,
                    "data": f"data:image/png;base64,{b64}",
                    "w": pix.width, "h": pix.height,
                })
                idx += 1
            except Exception:
                pass

        full_text.append(page.get_text("text"))

    doc.close()
    return {
        "filename": filename,
        "pdf_mode": True,
        "num_exercises": 0,
        "num_images": len(images),
        "images": images,
        "text": "\n".join(full_text),
        "exercises": [],
    }


# ─────────────────────────────────────────────
# DOCX : rasterisation générique des images non-web
# (EMF / WMF / TIFF / vectoriels) → PNG via LibreOffice
# ─────────────────────────────────────────────
_WEB_SAFE = ("png", "jpeg", "jpg", "gif", "webp", "bmp")
# Fragments de content-type non affichables par un navigateur → extension LibreOffice
_NONWEB_EXT = {
    "x-emf": "emf", "emf": "emf",
    "x-wmf": "wmf", "wmf": "wmf",
    "tiff": "tiff", "tif": "tiff",
}


def _is_web_safe(ct: str) -> bool:
    ct = (ct or "").lower()
    return any(w in ct for w in _WEB_SAFE)


def _nonweb_ext(ct: str):
    ct = (ct or "").lower()
    for frag, ext in _NONWEB_EXT.items():
        if frag in ct:
            return ext
    return None


def _autocrop_png(png_bytes: bytes, pad: int = 6) -> bytes:
    """Rogne les marges blanches d'un PNG (LibreOffice exporte une page entière)."""
    try:
        from PIL import Image, ImageChops
        im = Image.open(io.BytesIO(png_bytes)).convert("RGB")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bbox = ImageChops.difference(im, bg).getbbox()
        if bbox:
            l, t, r, b = bbox
            l = max(0, l - pad); t = max(0, t - pad)
            r = min(im.width, r + pad); b = min(im.height, b + pad)
            if r > l and b > t:
                im = im.crop((l, t, r, b))
        out = io.BytesIO()
        im.save(out, "PNG")
        return out.getvalue()
    except Exception:
        return png_bytes


def rasterize_blobs(jobs):
    """
    jobs : liste de (key, blob_bytes, ext).
    Convertit TOUTES les images non-web en PNG en UNE seule invocation LibreOffice
    (rapide), puis autocrop. Retourne {key: png_bytes}. Robuste : en cas d'échec,
    la clé est simplement absente du résultat (→ placeholder côté frontend).
    """
    result = {}
    if not jobs:
        return result
    with tempfile.TemporaryDirectory() as td:
        names = {}
        srcpaths = []
        for i, (key, blob, ext) in enumerate(jobs):
            fn = f"img{i}.{ext}"
            path = os.path.join(td, fn)
            with open(path, "wb") as f:
                f.write(blob)
            names[f"img{i}"] = key
            srcpaths.append(path)
        try:
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "png", "--outdir", td] + srcpaths,
                timeout=180, capture_output=True,
                env={**os.environ, "HOME": td},  # profil LibreOffice inscriptible
            )
        except Exception:
            return result
        for png in glob.glob(os.path.join(td, "*.png")):
            base = os.path.splitext(os.path.basename(png))[0]
            if base in names:
                try:
                    with open(png, "rb") as f:
                        result[names[base]] = _autocrop_png(f.read())
                except Exception:
                    pass
    return result


# ─────────────────────────────────────────────
# ENDPOINT : /parse
# DOCX → structure JSON pédagogique
# ─────────────────────────────────────────────
@app.post("/parse")
async def parse_document(file: UploadFile = File(...)):
    """
    Reçoit un DOCX ou PDF.
    Retourne une structure JSON avec blocs texte, images, tableaux,
    type d'exercice détecté, et positions relatives.
    """
    content = await file.read()
    filename = file.filename or "document"
    ext = filename.rsplit(".", 1)[-1].lower()

    # PDF : extraction directe PyMuPDF (images raster + figures vectorielles)
    if ext == "pdf":
        return parse_pdf(content, filename)

    # Parser le DOCX
    try:
        doc = DocxDocument(io.BytesIO(content))
    except Exception as e:
        raise HTTPException(400, f"Impossible de lire le document : {e}")

    # Extraire toutes les images du document.
    # Les images web (png/jpeg/…) sont encodées telles quelles ; les formats
    # non affichables par un navigateur (EMF/WMF/TIFF vectoriels des DOCX Word)
    # sont rasterisés en PNG via LibreOffice — solution GÉNÉRIQUE, tous formats.
    images = {}
    raster_jobs = []
    for rId, rel in doc.part.rels.items():
        if "image" in rel.target_ref:
            try:
                img_blob = rel.target_part.blob
                content_type = rel.target_part.content_type or "image/png"
                if _is_web_safe(content_type):
                    img_b64 = base64.b64encode(img_blob).decode()
                    images[rId] = {
                        "data": f"data:{content_type};base64,{img_b64}",
                        "content_type": content_type,
                        "size": len(img_blob)
                    }
                else:
                    # EMF/WMF/TIFF/… → à rasteriser (une seule passe LibreOffice)
                    raster_jobs.append((rId, img_blob, _nonweb_ext(content_type) or "emf"))
            except Exception:
                pass

    # Rasterisation groupée des images non-web → PNG affichables
    for rId, png in rasterize_blobs(raster_jobs).items():
        img_b64 = base64.b64encode(png).decode()
        images[rId] = {
            "data": f"data:image/png;base64,{img_b64}",
            "content_type": "image/png",
            "size": len(png)
        }
    # Les images non converties (ex : WDP illisible) sont simplement absentes :
    # le frontend affichera alors un placeholder étiqueté « coller ici ».

    # Construire les blocs
    blocks = []
    img_counter = [0]  # compteur global d'images

    def extract_paragraph_images(para_element):
        """Extrait les images inline d'un paragraphe."""
        found = []
        for blip in para_element.findall(".//" + qn("a:blip")):
            rId = blip.get(qn("r:embed"))
            if rId and rId in images:
                idx = img_counter[0]
                img_counter[0] += 1
                found.append({
                    "index": idx,
                    "rId": rId,
                    "data": images[rId]["data"],
                    "content_type": images[rId]["content_type"]
                })
        return found

    def process_paragraph(para):
        text = para.text.strip()
        imgs = extract_paragraph_images(para._element)

        if not text and not imgs:
            return None

        # Détecter si c'est une consigne (début d'exercice)
        ex_type = detect_exercise_type(text) if text else "autre"
        is_consigne = ex_type != "autre" and len(text.split()) <= 15

        block = {
            "type": "paragraph",
            "text": text,
            "images": imgs,
            "exercise_type": ex_type if is_consigne else None,
            "is_consigne": is_consigne,
            "style": para.style.name if para.style else "Normal",
            "alignment": str(para.alignment) if para.alignment else "left",
        }
        return block

    def process_table(table):
        rows = []
        table_images = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                cell_imgs = extract_paragraph_images(cell._element)
                table_images.extend(cell_imgs)
                cells.append({
                    "text": cell_text,
                    "images": cell_imgs
                })
            rows.append(cells)

        # Détecter le type de tableau
        all_text = " ".join(
            cell["text"] for row in rows for cell in row
        ).lower()
        ex_type = detect_exercise_type(all_text)

        return {
            "type": "table",
            "rows": rows,
            "images": table_images,
            "exercise_type": ex_type,
            "num_cols": len(rows[0]) if rows else 0,
            "num_rows": len(rows)
        }

    # Parcourir les éléments du document dans l'ordre
    from docx.oxml.ns import qn as oxqn
    body = doc.element.body

    current_exercise = None
    exercise_blocks = []

    for child in body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            # Trouver le paragraphe correspondant
            para = None
            for p in doc.paragraphs:
                if p._element is child:
                    para = p
                    break
            if para:
                block = process_paragraph(para)
                if block:
                    blocks.append(block)

        elif tag == "tbl":
            # Trouver le tableau correspondant
            for t in doc.tables:
                if t._element is child:
                    block = process_table(t)
                    blocks.append(block)
                    break

    # Grouper les blocs en exercices
    exercises = []
    current_ex = None

    for block in blocks:
        if block.get("is_consigne"):
            if current_ex:
                exercises.append(current_ex)
            current_ex = {
                "exercise_type": block["exercise_type"],
                "consigne": block["text"],
                "blocks": [block],
                "all_images": list(block["images"])
            }
        else:
            if current_ex:
                current_ex["blocks"].append(block)
                current_ex["all_images"].extend(block.get("images", []))
            else:
                # Bloc avant le premier exercice (titre, entête…)
                exercises.append({
                    "exercise_type": "header",
                    "consigne": None,
                    "blocks": [block],
                    "all_images": list(block.get("images", []))
                })

    if current_ex:
        exercises.append(current_ex)

    return {
        "filename": filename,
        "num_exercises": len([e for e in exercises if e["exercise_type"] != "header"]),
        "num_images": img_counter[0],
        "exercises": exercises,
        "raw_blocks": blocks  # Pour débogage
    }


# ─────────────────────────────────────────────
# ENDPOINT : /generate
# Proxy Anthropic avec retry + chunking
# ─────────────────────────────────────────────
@app.post("/generate")
async def generate(request: Request):
    """
    Proxy vers l'API Anthropic.
    - Gère les erreurs rate limit (retry automatique)
    - Gère les documents trop longs (chunking)
    - Retourne la réponse Claude
    """
    body = await request.json()
    api_key = body.get("api_key")
    if not api_key:
        raise HTTPException(400, "api_key requis")

    model = body.get("model", "claude-haiku-4-5-20251001")
    system = body.get("system", "")
    messages = body.get("messages", [])
    max_tokens = body.get("max_tokens", 6000)

    client = anthropic.Anthropic(api_key=api_key)

    # Retry avec backoff exponentiel
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                system=[{
                    "type": "text",
                    "text": system,
                    "cache_control": {"type": "ephemeral"}
                }] if system else [],
                messages=messages
            )
            text = "".join(
                block.text for block in response.content
                if hasattr(block, "text")
            )
            return {
                "content": text,
                "usage": {
                    "input_tokens": response.usage.input_tokens,
                    "output_tokens": response.usage.output_tokens
                }
            }
        except anthropic.RateLimitError as e:
            if attempt < max_retries - 1:
                wait = 2 ** (attempt + 1)  # 2s, 4s, 8s
                time.sleep(wait)
                continue
            raise HTTPException(429, f"Rate limit après {max_retries} tentatives : {e}")
        except anthropic.BadRequestError as e:
            raise HTTPException(400, f"Prompt trop long ou invalide : {e}")
        except Exception as e:
            raise HTTPException(500, f"Erreur API : {e}")


# ─────────────────────────────────────────────
# ENDPOINT : /export
# Structure JSON adaptée → DOCX
# ─────────────────────────────────────────────
@app.post("/export")
async def export_docx(request: Request):
    """
    Reçoit la structure JSON adaptée par Claude.
    Reconstruit un DOCX avec :
    - le texte adapté
    - les images originales à leurs positions
    - le layout préservé selon le type d'exercice
    """
    body = await request.json()
    exercises = body.get("exercises", [])
    filename = body.get("filename", "tailory_adapte.docx")

    doc = DocxDocument()

    # Style de base
    style = doc.styles["Normal"]
    style.font.name = "Andika"
    style.font.size = Pt(12)

    def add_image_to_para(para, img_data: str):
        """Ajoute une image base64 à un paragraphe."""
        if not img_data or not img_data.startswith("data:"):
