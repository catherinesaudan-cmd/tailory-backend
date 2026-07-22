"""
Tailory Backend V2.8 — Pipeline documentaire pédagogique
FastAPI + python-docx + pdf2docx + Anthropic proxy

Endpoints:
  POST /parse    → DOCX/PDF/ODT → structure JSON pédagogique
                   (ODT : converti en PDF via LibreOffice puis pipeline PDF —
                    indispensable car les ODT contiennent souvent des formes
                    vectorielles natives invisibles pour l'extracteur DOCX)
  POST /generate → proxy Anthropic avec retry + chunking
  POST /export   → structure JSON adaptée → DOCX
  POST /convert  → PDF → DOCX (existant, conservé)
  GET  /health   → vérification

V2.8 (retour essai 39, validé localement sur Pacomefantome.odt) :
  · RASTERS AUTONOMES : une image raster (scan, photo, dessin importé) est
    déjà une unité sémantique complète — elle n'est PLUS clusterisée, ni avec
    les autres rasters, ni avec les vecteurs. Cas prouvé (essai 39) : trois
    dessins scannés aux marges blanches se chevauchant fusionnaient en un
    bloc de 180×134 mm dont la bbox avalait les lignes basses de la banque
    de mots — le fragment de texte tronqué « er, hululer… » partait comme
    figure et était placé tel quel sur la fiche. Seuls les rasters quasi
    identiques (recouvrement ≥ 80 %) sont dédupliqués. Le clustering ne
    s'applique désormais qu'aux tracés VECTORIELS (une figure = plusieurs
    primitives), sa raison d'être.

V2.7 (retours essais 36-37, validé localement sur AireDefinitionetMesure.odt) :
  · COMPOSANTES CONNEXES : le clustering par union de bounding-box était
    structurellement vorace — chaque fusion créait une boîte plus grande qui
    absorbait tout ce qu'elle SURVOLAIT sans en être réellement proche.
    Cas prouvé (essais 36-37) : le trait séparateur pointillé passait à
    10,5 pt du tissu E → fusion → la bbox résultante (x45-185 mm)
    intersectait D, puis C, puis B → composite unique de 169×109 mm,
    impossible à placer figure par figure côté modèle (placeholders
    contradictoires, « B et E » alors que C et D y étaient aussi).
    Désormais : composantes connexes calculées sur les primitives D'ORIGINE
    (arête si écart < marge), bbox par composante calculée À LA FIN.
    Résultat validé : les 5 tissus A-E sortent individuellement.
  · MARGE 14 → 10 pt : l'écart réel entre les tissus B et C est de 12,2 pt ;
    les sous-éléments d'une même figure (cellules de grille, hachures) se
    touchent ou restent < 10 pt. Aucune figure légitime connue du corpus
    n'a d'écart interne de 10-14 pt.
  · SÉPARATEURS PARTIELS : une RANGÉE de barres fines (h < 16 pt) alignées
    couvrant ensemble > 50 % de la largeur de page, sans voisin plein,
    est un trait de section coupé par du texte — exclue avant clustering.
    Le filtre pleine largeur (0.92·pw) ne voyait pas ces morceaux courts.
  · BANDES DE TEXTE : une bande plate (h < 32 pt ≈ 11 mm) contenant du texte
    (≥ 2 mots de ≥ 3 lettres, ou 1 mot couvrant > 30 % de sa largeur) est
    une cellule d'en-tête de tableau ou un titre décoré, PAS une figure —
    exclue avant clustering. Validé : « Figures | Calcul de l'aire »,
    « Exercice N », « L'aire du triangle », « Définition N » — précisément
    les bandeaux gris qui fuyaient dans les fiches des essais 36-37.
    Les figures plates légitimes survivent : tissus sans texte (lettres
    vectorielles), lignes graduées (chiffres seulement).

V2.3 (retours essais 25-29) :
  · FONDS DE PAGE : les rectangles couvrant > 85 % de la page (LibreOffice
    exporte un fond blanc pleine page selon le modèle de document) sont exclus
    AVANT clustering — en v2.2 un tel fond absorbait toutes les régions en un
    cluster pleine page, jeté ensuite → 0 figure transmise (cas essai 27 :
    les 15 images d'animaux et les schémas existaient dans la source).
  · RE-SPLIT : un cluster qui couvre malgré tout > 85 % de la page est
    re-clusterisé plus finement (marge 5) au lieu d'être jeté ; en dernier
    recours, ses primitives pleines sont conservées individuellement.
  · LIGNES FINES : les segments à mesurer et lignes graduées (une dimension
    quasi nulle, l'autre ≥ 40 pt) sont désormais capturés comme figures —
    en v2.2 le filtre « dim < 10 » les jetait (cas essai 29 : segments a-f
    « non disponibles » alors qu'ils existaient). Ils partent avec leur
    taille physique réelle (w_mm/h_mm du clip rasterisé) pour l'impression
    à l'échelle (class="img-echelle") — mesure à la règle valide.
    Anti-bruit, une ligne fine n'est une figure QUE si :
      - elle est isolée (rien à moins de 3 pt — élimine les bordures de
        tableaux, qui se croisent aux coins),
      - elle n'est pas un souligné de texte (texte juste au-dessus couvrant
        ≥ 60 % de sa longueur),
      - elle n'appartient pas à une famille de ≥ 3 lignes parallèles de même
        longueur (lignes d'écriture, grilles) — les segments à mesurer ont
        des longueurs toutes différentes, c'est le principe de l'exercice.
"""

from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import anthropic
import base64
import io
import json
import os
import re
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
        "http://127.0.0.1",
        "null"  # v2.6 — fichier tailory ouvert en local (file://) : Origin: null
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
# PDF : extraction PyMuPDF — raster + vectoriel + lignes fines
# Clustering v2.7 : composantes connexes (fini l'absorption par bbox)
# ─────────────────────────────────────────────
import fitz  # PyMuPDF, déjà installé (dépendance pdf2docx)


def _connected_components(rects, margin=10.0):
    """v2.7 — Regroupe les primitives par COMPOSANTES CONNEXES : arête entre
    deux primitives d'ORIGINE si leur écart < margin ; la bbox de chaque
    composante est calculée à la fin. Remplace l'union itérative de bbox,
    qui absorbait toute primitive SURVOLÉE par une boîte déjà fusionnée
    sans qu'elle soit réellement proche (cas essais 36-37 : le séparateur
    collé au tissu E entraînait D, C et B dans un composite unique)."""
    rects = [fitz.Rect(r) for r in rects]
    n = len(rects)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for i in range(n):
        gi = fitz.Rect(rects[i].x0 - margin, rects[i].y0 - margin,
                       rects[i].x1 + margin, rects[i].y1 + margin)
        for j in range(i + 1, n):
            if gi.intersects(rects[j]):
                union(i, j)

    comps = {}
    for i in range(n):
        comps.setdefault(find(i), []).append(rects[i])

    out = []
    for members in comps.values():
        bb = fitz.Rect(members[0])
        for m in members[1:]:
            bb |= m
        out.append(bb)
    return out


def _cluster_rasters(rasters, page, margin=16.0):
    """v2.9 — Les rasters se regroupent ENTRE EUX (jamais avec les vecteurs,
    acquis v2.8 conservé) : les scènes composées de nombreux petits rasters
    (pièces de monnaie, billets, tas de tomates, vignettes de situations —
    128 rasters sur la seule série 5P) partaient en figures individuelles,
    saturaient le plafond de miniatures du frontend et sortaient en
    « coller ici » (essai 43). Trois règles :
    1. Composantes connexes, marge 16 pt (écarts intra-scène mesurés : 0-14,6 pt ;
       figures distinctes du corpus : >= 18 pt).
    2. GARDE PHOTOS : deux grands rasters (min-dim > 120 pt ~ 42 mm) sans
       recouvrement ne fusionnent JAMAIS — les photos côte à côte restent
       autonomes (essai 42 insectes : 15/15 photos, à ne pas casser).
    3. SCISSION TABLEAU : un composite multi-rasters est re-scindé à chaque
       bordure horizontale de tableau qui le traverse sans toucher aucun de
       ses membres (collections en lignes de tableau, p5 : écarts
       intra/inter-collections indistinguables géométriquement)."""
    n = len(rasters)
    if n <= 1:
        return list(rasters)

    def _gap(a, b):
        dx = max(0.0, max(a.x0, b.x0) - min(a.x1, b.x1))
        dy = max(0.0, max(a.y0, b.y0) - min(a.y1, b.y1))
        return (dx * dx + dy * dy) ** 0.5

    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for i in range(n):
        for j in range(i + 1, n):
            a, b = rasters[i], rasters[j]
            if _gap(a, b) >= margin:
                continue
            big_a = min(a.width, a.height) > 120
            big_b = min(b.width, b.height) > 120
            if big_a and big_b and (a & b).is_empty:
                continue  # garde photos : pas de fusion sans recouvrement
            union(i, j)

    comps = {}
    for i in range(n):
        comps.setdefault(find(i), []).append(rasters[i])

    # Bordures horizontales candidates pour la scission (triplets de traits
    # LibreOffice regroupés à 3 pt près)
    hlines = []
    try:
        for d in page.get_drawings():
            r = d.get("rect")
            if r is not None and r.height < 4 and r.width > 100:
                hlines.append(fitz.Rect(r))
    except Exception:
        pass

    out = []
    for members in comps.values():
        if len(members) == 1:
            out.append(members[0])
            continue
        bb = fitz.Rect(members[0])
        for m in members[1:]:
            bb |= m
        # Coupes : bordures traversant le composite entre les membres
        cuts = []
        for t in hlines:
            cy = (t.y0 + t.y1) / 2
            if not (bb.y0 + 4 < cy < bb.y1 - 4):
                continue
            if t.x0 > bb.x0 + 0.2 * bb.width or t.x1 < bb.x1 - 0.2 * bb.width:
                continue  # ne traverse pas le composite
            if any(m.y0 < cy < m.y1 for m in members):
                continue  # couperait un raster : pas une frontière de lignes
            if not any(abs(c - cy) <= 3 for c in cuts):
                cuts.append(cy)
        if not cuts:
            out.append(bb)
            continue
        cuts.sort()
        bands = {}
        for m in members:
            k = sum(1 for c in cuts if c < (m.y0 + m.y1) / 2)
            bands.setdefault(k, []).append(m)
        for grp in bands.values():
            gb = fitz.Rect(grp[0])
            for m in grp[1:]:
                gb |= m
            out.append(gb)
    return out


def _drop_separator_rows(solids, pw):
    """v2.7 — RANGÉE de barres fines (h < 16 pt) alignées couvrant ensemble
    > 50 % de la largeur de page, sans voisin plein = trait de section coupé
    par du texte (« ---- Exercice 2 ---- »). Le filtre pleine largeur (0.92)
    ne les voyait pas puisque chaque morceau est court. C'est ce trait qui,
    passant à 10,5 pt du tissu E, déclenchait la fusion en chaîne."""
    bars = [r for r in solids if r.height < 16]
    others = [r for r in solids if r.height >= 16]
    drop = set()
    used = set()
    for i, b in enumerate(bars):
        if i in used:
            continue
        row = [i]
        cy = (b.y0 + b.y1) / 2
        for j, o in enumerate(bars):
            if j != i and abs((o.y0 + o.y1) / 2 - cy) < 6:
                row.append(j)
        span = sum(bars[k].width for k in row)
        if span > 0.5 * pw:
            rowrects = [bars[k] for k in row]
            near_full = any(
                fitz.Rect(rr.x0 - 3, rr.y0 - 3, rr.x1 + 3, rr.y1 + 3).intersects(o)
                for rr in rowrects for o in others)
            if not near_full:
                drop.update(row)
                used.update(row)
    return others + [b for i, b in enumerate(bars) if i not in drop]


def _is_text_band(r, words):
    """v2.7 — bande plate (h < 32 pt ≈ 11 mm) contenant du texte : ≥ 2 mots
    de ≥ 3 lettres, ou 1 seul mot couvrant > 30 % de sa largeur = cellule
    d'en-tête de tableau ou titre décoré, PAS une figure. Validé sur la
    source aires : filtre exactement « Figures | Calcul de l'aire »,
    « Exercice N », « L'aire du triangle », « Définition N » — les bandeaux
    gris qui fuyaient dans les fiches. Les figures plates légitimes passent :
    tissus (lettres vectorielles, pas de texte), lignes graduées (chiffres)."""
    if r.height >= 32:
        return False
    n = 0
    cover = 0.0
    for w in words:
        wr = fitz.Rect(w[:4])
        inter = wr & r
        if not inter.is_empty and wr.get_area() > 0 \
                and inter.get_area() > 0.5 * wr.get_area():
            if sum(1 for ch in w[4] if ch.isalpha()) >= 3:
                n += 1
                cover += wr.width
    return n >= 2 or (n >= 1 and cover > 0.3 * r.width)


def _is_underline_of_text(words, t):
    """Ligne horizontale = souligné si du texte immédiatement au-dessus
    couvre ≥ 60 % de sa longueur.
    v2.4 — tolérance élargie : les jambages (descenders) du texte souligné
    descendent SOUS le trait ; mesuré sur les sources réelles, le bas des mots
    est 1,7 à 1,9 pt sous le haut du trait, et l'ancienne condition
    « y1 <= t.y0 + 2 » ne tenait qu'à 0,1 pt près — d'où le soulignement du
    titre « Remplir un tableau… » transmis comme segment par le serveur
    (métriques de police différentes). Fenêtre désormais : bas du mot entre
    10 pt au-dessus et 6 pt au-dessous du trait. Les vrais segments à mesurer
    ont leur texte le plus proche à ≥ 14 pt : aucun faux positif."""
    if t.height > t.width:
        return False
    cover = 0.0
    for w in words:
        x0, y0, x1, y1 = w[:4]
        if y1 <= t.y0 + 6 and t.y0 - y1 < 10:
            ov = min(x1, t.x1) - max(x0, t.x0)
            if ov > 0:
                cover += ov
    return cover >= 0.6 * t.width


def _collect_page_regions(page):
    """
    Régions candidates d'une page, en deux familles :
    - solids : images raster + tracés vectoriels « pleins »
    - thins  : lignes fines légitimes (segments à mesurer, lignes graduées)
    Filtres : fonds de page (> 85 % de l'aire), séparateurs pleine largeur,
    micro-tracés, soulignés de texte, bordures de tableaux (non isolées),
    familles de lignes identiques (lignes d'écriture, grilles).
    v2.7 : + rangées de séparateurs partiels, + bandes de texte (en-têtes).
    """
    pw, ph = page.rect.width, page.rect.height
    page_area = pw * ph
    solids, thins = [], []
    rasters = []  # v2.8 — les rasters ne passent PAS par le clustering

    # 1. Images raster — v2.8 : chaque raster est une figure autonome.
    for img in page.get_images(full=True):
        try:
            for rect in page.get_image_rects(img[0]):
                if rect.width > 12 and rect.height > 12:
                    if rect.width * rect.height > 0.85 * page_area:
                        continue  # image de fond pleine page
                    rasters.append(fitz.Rect(rect))
        except Exception:
            pass
    # Déduplication des rasters quasi identiques (recouvrement ≥ 80 %)
    dedup = []
    for r in rasters:
        dup = False
        for o in dedup:
            inter = r & o
            if not inter.is_empty and inter.get_area() >= 0.8 * min(r.get_area(), o.get_area()):
                o |= r
                dup = True
                break
        if not dup:
            dedup.append(fitz.Rect(r))
    rasters = dedup
    # v2.9 — composition des scènes multi-rasters (pièces, billets, tas…)
    rasters = _cluster_rasters(rasters, page)

    # 2. Tracés vectoriels
    try:
        for d in page.get_drawings():
            r = d.get("rect")
            if r is None:
                continue
            if r.width > 0.92 * pw and r.height < 20:
                continue  # ligne de séparation pleine largeur
            if r.width * r.height > 0.85 * page_area:
                continue  # FOND DE PAGE (v2.3) — absorbait tout en v2.2
            if r.width < 10 and r.height < 10:
                continue  # micro-tracé
            if r.width < 10 or r.height < 10:
                # Ligne fine : candidate seulement si assez longue pour être
                # une figure (segment, ligne graduée), pas un simple tiret
                if max(r.width, r.height) >= 40:
                    thins.append(fitz.Rect(r))
                continue
            solids.append(fitz.Rect(r))
    except Exception:
        pass

    # 2b. v2.7 — séparateurs partiels (rangées de barres) et bandes de texte
    # (en-têtes de tableaux, titres décorés) exclus AVANT clustering : ce sont
    # eux qui provoquaient les fusions en chaîne et les fuites de bandeaux gris.
    solids = _drop_separator_rows(solids, pw)
    words_v27 = page.get_text("words")
    solids = [r for r in solids if not _is_text_band(r, words_v27)]
    # v2.8 — un cluster vectoriel entièrement contenu dans un raster serait un
    # doublon (annotation déjà visible dans le clip du raster) : on le retire
    # en amont pour ne pas générer deux figures du même contenu.
    solids = [r for r in solids
              if not any(ra.contains(r) for ra in rasters)]

    # 3. Anti-bruit lignes fines
    if thins:
        words = words_v27

        def _isolated(t):
            probe = fitz.Rect(t.x0 - 3, t.y0 - 3, t.x1 + 3, t.y1 + 3)
            if any(probe.intersects(o) for o in solids):
                return False
            return not any((o is not t) and probe.intersects(o) for o in thins)

        thins = [t for t in thins if _isolated(t)]
        thins = [t for t in thins if not _is_underline_of_text(words, t)]

        kept = []
        for t in thins:
            horiz = t.width >= t.height
            L = t.width if horiz else t.height
            same = sum(1 for o in thins
                       if (o.width >= o.height) == horiz
                       and abs((o.width if horiz else o.height) - L) <= 0.05 * L)
            if same < 3:
                kept.append(t)
        thins = kept

    return solids, thins, rasters


def parse_pdf(content: bytes, filename: str):
    """
    Extrait d'un PDF, dans l'ordre de lecture :
    - les images raster (photos, dessins importés)
    - les figures vectorielles (tracés regroupés puis rasterisés en PNG 2x)
    - les lignes fines légitimes (segments à mesurer, lignes graduées) — v2.3
    - le texte complet
    Chaque figure part avec sa taille physique (w_mm/h_mm du clip rasterisé)
    pour l'impression à l'échelle réelle (class="img-echelle").
    v2.7 : clustering par composantes connexes, marge 10 pt.
    """
    doc = fitz.open(stream=content, filetype="pdf")
    images = []
    full_text = []
    idx = 0

    for pno, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        page_area = pw * ph

        solids, thins, rasters = _collect_page_regions(page)

        # v2.5 — les lignes fines ne partent que si la page parle de mesure ou
        # de tracé : sur une fiche de français, les lignes de réponse aux
        # longueurs variées traversent le filtre « familles » (28 fausses
        # lignes sur la série de conjugaison 5P) et noieraient le modèle sous
        # des « segments » sans objet. Mots-clés calibrés sur les 9 sources :
        # les pages segments/périmètres/aires/quadrillages en contiennent
        # toutes au moins un, aucune page de français n'en contient.
        if thins:
            ptxt = page.get_text("text").lower()
            if not (any(k in ptxt for k in (
                    "mesur", "segment", "périmètre", "perimetre",
                    "centimètre", "centimetre", "millimètre", "millimetre",
                    "gradu", "quadrill", "trace", "construc",
                    "géométr", "geometr"))
                    or re.search(r"\baires?\b", ptxt)):  # « aire » en mot entier — pas « faire »
                thins = []

        # v2.4 — les lignes fines ne sont JAMAIS clusterisées : une ligne qui a
        # survécu aux filtres anti-bruit est un segment autonome à mesurer.
        regions = solids

        # Clustering v2.7 : composantes connexes sur les primitives d'origine,
        # marge 10 pt (l'écart réel entre deux figures distinctes du corpus
        # est ≥ 12 pt ; les sous-éléments d'une même figure restent < 10 pt).
        clusters = _connected_components(regions, margin=10.0)

        # Re-split (v2.3, adapté v2.7) : un cluster quasi pleine page n'est pas
        # jeté, il est re-clusterisé plus finement ; en dernier recours, ses
        # primitives pleines sont gardées individuellement.
        final = []
        for c in clusters:
            c = c & page.rect
            if c.width * c.height <= 0.85 * page_area:
                final.append(c)
                continue
            subs = _connected_components(
                [r for r in regions if r.intersects(c)], margin=5.0)
            for s in subs:
                s = s & page.rect
                if s.width * s.height <= 0.85 * page_area:
                    final.append(s)
                else:
                    final.extend(r for r in solids
                                 if r.intersects(s) and r.width >= 24 and r.height >= 24)

        # v2.4 — chaque ligne fine part individuellement
        final.extend(thins)
        # v2.8/v2.9 — les rasters ne se mélangent jamais aux vecteurs ; depuis
        # v2.9 les scènes multi-rasters arrivent déjà composées (et re-scindées
        # aux bordures de tableau) par _cluster_rasters.
        final.extend(rasters)

        # Filtres finaux : figures pleines OU lignes fines assez longues
        keep, seen = [], set()
        for r in final:
            r = r & page.rect
            key = (round(r.x0), round(r.y0), round(r.x1), round(r.y1))
            if key in seen:
                continue
            seen.add(key)
            full = r.width >= 24 and r.height >= 24
            slim = min(r.width, r.height) < 24 and max(r.width, r.height) >= 40
            if not (full or slim):
                continue
            if (r.width * r.height) > 0.85 * page_area:
                continue
            keep.append(r)

        # Ordre de lecture : haut → bas, gauche → droite
        keep.sort(key=lambda r: (round(r.y0 / 24), r.x0))

        # Rasterisation 2x de chaque zone
        page_words = page.get_text("words")
        for r in keep:
            try:
                # Lignes fines : dilater le clip pour que le trait soit
                # visible ; w_mm/h_mm décrivent le CLIP rasterisé, donc
                # l'échelle réelle reste exacte à l'impression.
                clip = fitz.Rect(r)
                is_thin = min(r.width, r.height) < 24
                if r.height < 24 and r.width >= r.height:
                    clip = fitz.Rect(r.x0 - 3, r.y0 - 6, r.x1 + 3, r.y1 + 6)
                elif r.width < 24:
                    clip = fitz.Rect(r.x0 - 6, r.y0 - 3, r.x1 + 6, r.y1 + 3)
                clip = clip & page.rect
                # v2.4 — GARDE FINALE : un clip de ligne fine qui contient du
                # texte n'est PAS un segment à mesurer (soulignement de titre,
                # ligne collée à un libellé) → on ne le transmet pas, quel que
                # soit le verdict des filtres géométriques en amont. Seules les
                # LETTRES comptent : les chiffres d'une ligne graduée restent
                # légitimes.
                if is_thin:
                    letters = 0
                    for w in page_words:
                        wr = fitz.Rect(w[:4])
                        inter = wr & clip
                        if not inter.is_empty and wr.get_area() > 0 \
                                and inter.get_area() > 0.3 * wr.get_area():
                            letters += sum(1 for ch in w[4] if ch.isalpha())
                    if letters >= 3:
                        continue
                pix = page.get_pixmap(clip=clip, matrix=fitz.Matrix(2, 2))
                if pix.width < 8 or pix.height < 8:
                    continue
                b64 = base64.b64encode(pix.tobytes("png")).decode()
                images.append({
                    "index": idx,
                    "page": pno + 1,
                    "data": f"data:image/png;base64,{b64}",
                    "w": pix.width, "h": pix.height,
                    # Taille PHYSIQUE de la zone rasterisée dans le document
                    # source (points PDF → mm : 1 pt = 25.4/72 mm). Permet au
                    # frontend d'imprimer la figure à taille réelle
                    # (class="img-echelle") pour la mesure à la règle.
                    "w_mm": round(clip.width * 25.4 / 72, 1),
                    "h_mm": round(clip.height * 25.4 / 72, 1),
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
# ODT (et formats bureautiques) : conversion → PDF via LibreOffice
# La route ODT→PDF est volontaire : les ODT contiennent souvent des formes
# vectorielles dessinées nativement (quadrillages, figures géométriques…)
# qui seraient PERDUES en ODT→DOCX (DrawingML non extrait), alors que le
# pipeline PDF/PyMuPDF les récupère comme figures via get_drawings().
# ─────────────────────────────────────────────
def convert_office_to_pdf(content: bytes, ext: str):
    """
    Convertit un document bureautique (odt, doc, rtf…) en PDF via LibreOffice.
    Retourne (pdf_bytes, "") en cas de succès, (None, message_erreur) sinon —
    le message contient la sortie de soffice pour diagnostiquer sans deviner.
    """
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, f"doc.{ext}")
        with open(src, "wb") as f:
            f.write(content)
        try:
            proc = subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", td, src],
                timeout=180, capture_output=True,
                env={**os.environ, "HOME": td},  # profil LibreOffice inscriptible
            )
        except subprocess.TimeoutExpired:
            return None, "timeout soffice (180 s)"
        except Exception as e:
            return None, f"soffice injoignable : {e}"
        pdf_path = os.path.join(td, "doc.pdf")
        if not os.path.exists(pdf_path):
            out = (proc.stdout or b"").decode(errors="replace")[-300:]
            err = (proc.stderr or b"").decode(errors="replace")[-300:]
            return None, f"soffice n'a pas produit de PDF. stdout: {out} | stderr: {err}"
        with open(pdf_path, "rb") as f:
            return f.read(), ""


# ─────────────────────────────────────────────
# ENDPOINT : /parse
# DOCX / PDF / ODT → structure JSON pédagogique
# ─────────────────────────────────────────────
@app.post("/parse")
async def parse_document(file: UploadFile = File(...)):
    """
    Reçoit un DOCX, un PDF ou un ODT.
    Retourne une structure JSON avec blocs texte, images, tableaux,
    type d'exercice détecté, et positions relatives.
    ODT : converti en PDF (LibreOffice) puis traité par le pipeline PDF.
    """
    content = await file.read()
    filename = file.filename or "document"
    ext = filename.rsplit(".", 1)[-1].lower()

    # ODT (LibreOffice) : conversion en PDF puis pipeline PDF existant.
    # (le même mécanisme fonctionnerait pour doc/rtf si besoin un jour)
    if ext == "odt":
        pdf_bytes, conv_err = convert_office_to_pdf(content, ext)
        if pdf_bytes is None:
            raise HTTPException(
                422, f"Conversion ODT→PDF échouée (LibreOffice) : {conv_err}. "
                     "Exportez le document en PDF depuis LibreOffice et réessayez.")
        result = parse_pdf(pdf_bytes, filename)
        result["source_format"] = "odt"
        # Joindre le PDF converti si assez petit pour être transmis à l'IA
        # comme document natif (limite frontend : 150 000 caractères base64).
        if len(pdf_bytes) * 4 / 3 < 150_000:
            result["pdf_b64"] = base64.b64encode(pdf_bytes).decode()
        return result

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
            return
        try:
            _, b64 = img_data.split(",", 1)
            img_bytes = base64.b64decode(b64)
            img_stream = io.BytesIO(img_bytes)
            run = para.add_run()
            run.add_picture(img_stream, width=Inches(1.5))
        except Exception:
            pass

    def add_exercise_header(title: str, num: int):
        para = doc.add_paragraph()
        run = para.add_run(f"Exercice {num} — {title.upper()}")
        run.bold = True
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x6e, 0xb7, 0x9e)

    def add_consigne(text: str):
        para = doc.add_paragraph()
        run = para.add_run(text)
        run.bold = True
        run.font.size = Pt(13)

    def add_response_line(label: str = ""):
        para = doc.add_paragraph()
        if label:
            para.add_run(f"{label} ").bold = True
        para.add_run("_" * 20)

    def build_relier_table(images, words):
        """Template fixe pour exercice relier : image | • | mot"""
        if not words:
            return
        table = doc.add_table(rows=len(words), cols=3)
        table.style = "Table Grid"
        for i, word in enumerate(words):
            row = table.rows[i]
            # Colonne image
            cell_img = row.cells[0]
            if i < len(images):
                para = cell_img.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                add_image_to_para(para, images[i]["data"])
            # Colonne point
            row.cells[1].text = "•"
            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Colonne mot
            row.cells[2].text = word
            row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

    ex_num = 0
    for exercise in exercises:
        ex_type = exercise.get("exercise_type", "autre")
        consigne = exercise.get("adapted_consigne") or exercise.get("consigne", "")
        blocks = exercise.get("adapted_blocks") or exercise.get("blocks", [])
        images = exercise.get("all_images", [])

        if ex_type == "header":
            for block in blocks:
                if block.get("text"):
                    doc.add_paragraph(block["text"])
            continue

        ex_num += 1

        # En-tête exercice
        add_exercise_header(ex_type, ex_num)

        # Consigne
        if consigne:
            add_consigne(consigne)

        # Corps selon le type d'exercice
        if ex_type == "relier":
            # Extraire les mots de la colonne droite depuis les blocs
            words = [
                b["text"] for b in blocks
                if b.get("text") and not b.get("is_consigne")
                and len(b["text"].split()) <= 4
            ]
            build_relier_table(images, words)

        elif ex_type in ("compléter", "numéroter"):
            for block in blocks:
                if block.get("is_consigne"):
                    continue
                if block["type"] == "table":
                    # Reconstruire le tableau
                    rows = block.get("rows", [])
                    if rows:
                        t = doc.add_table(rows=len(rows), cols=len(rows[0]))
                        t.style = "Table Grid"
                        for i, row in enumerate(rows):
                            for j, cell in enumerate(row):
                                t.rows[i].cells[j].text = cell.get("text", "")
                else:
                    text = block.get("text", "")
                    if text:
                        doc.add_paragraph(text)
                    if block.get("images"):
                        para = doc.add_paragraph()
                        for img in block["images"][:2]:
                            add_image_to_para(para, img["data"])

        elif ex_type == "classer":
            # Tableau découpe compact
            all_items = [
                b["text"] for b in blocks
                if b.get("text") and not b.get("is_consigne")
            ]
            if all_items:
                t = doc.add_table(rows=len(all_items), cols=2)
                t.style = "Table Grid"
                for i, item in enumerate(all_items):
                    t.rows[i].cells[0].text = str(i + 1)
                    t.rows[i].cells[1].text = ""  # case réponse

        else:
            # Fallback : texte + images
            for block in blocks:
                if block.get("is_consigne"):
                    continue
                if block.get("text"):
                    doc.add_paragraph(block["text"])
                for img in block.get("images", [])[:2]:
                    para = doc.add_paragraph()
                    add_image_to_para(para, img["data"])

        # Séparateur
        doc.add_paragraph("─" * 30)

    # Sauvegarder et retourner
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ─────────────────────────────────────────────
# ENDPOINT : /convert (existant — conservé)
# PDF → DOCX
# ─────────────────────────────────────────────
@app.post("/convert")
async def convert_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "Format PDF requis")

    uid = str(uuid.uuid4())[:8]
    pdf_path = f"/tmp/{uid}.pdf"
    docx_path = f"/tmp/{uid}.docx"

    try:
        content = await file.read()
        with open(pdf_path, "wb") as f:
            f.write(content)
        cv = Converter(pdf_path)
        cv.convert(docx_path)
        cv.close()
        with open(docx_path, "rb") as f:
            docx_bytes = f.read()
        return StreamingResponse(
            io.BytesIO(docx_bytes),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f'attachment; filename="{file.filename.replace(".pdf", ".docx")}"'}
        )
    finally:
        for p in [pdf_path, docx_path]:
            if os.path.exists(p):
                os.remove(p)


# ─────────────────────────────────────────────
# ENDPOINT : /health
# ─────────────────────────────────────────────
@app.get("/health")
def health():
    return {"status": "ok", "version": "2.9"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
