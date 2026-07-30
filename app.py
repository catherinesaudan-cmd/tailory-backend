"""
Tailory Backend V2.14 — Pipeline documentaire pédagogique
FastAPI + python-docx + pdf2docx + Anthropic proxy + pictogrammes ARASAAC

Endpoints:
  POST /pictos   → résolution de mots-clés en pictogrammes ARASAAC (base64)
                   (chantier prioritaire validé : supports visuels du mode
                    Participation — non-lecteur, non-verbal/CAA, allophone)
  POST /parse    → DOCX/PDF/ODT → structure JSON pédagogique
                   (ODT : converti en PDF via LibreOffice puis pipeline PDF —
                    indispensable car les ODT contiennent souvent des formes
                    vectorielles natives invisibles pour l'extracteur DOCX)
  POST /generate → proxy Anthropic avec retry + chunking
  POST /export   → structure JSON adaptée → DOCX
  POST /convert  → PDF → DOCX (existant, conservé)
  GET  /health   → vérification

V2.11 (chantier 6.8 — segments courts étiquetés, 31 juillet 2026) :
  · Un trait fin de 7 à 14 mm devient une figure candidate s'il porte une
    ÉTIQUETTE DE SÉRIE juste à sa gauche — une lettre ou un chiffre seul
    suivi d'une parenthèse ou d'un point : a) b) e) 1) 2. — ET qu'un autre
    trait étiqueté de longueur différente (> 5 %) existe sur la même page.
    Mesuré sur essai_10117 : le segment e) fait 8,0 mm dans la source,
    sous le plancher de 14,1 mm (40 pt) qui écarte tirets, puces et
    soulignements ; dans un exercice de MESURE, un trait court qui porte
    sa lettre est une vraie figure, pas un parasite.
  · Ce qui reste dehors, et pourquoi : les tirets et puces (jamais
    étiquetés) ; les lignes de réponse d'une liste numérotée « 1) ___ »
    (étiquetées mais TOUTES de la même longueur — les segments d'un
    exercice de mesure ont des longueurs toutes différentes, c'est le
    principe de l'exercice) ; un trait court étiqueté isolé sans compagnon
    (le plancher de 14,1 mm garde le dernier mot) ; l'étiquette au-dessus
    du trait (trop proche d'un mot souligné ou d'un numérateur — porte
    fermée pour cette version). Les promus subissent ensuite les mêmes
    filtres anti-bruit que les longs (isolement, souligné, familles) et le
    filtre v2.5 des pages de mesure.

V2.10 (grilles et formes composites, 27 juillet 2026) :
  · Les quadrillages et les tableaux d'une source ne partent plus en miettes.
    Trois symptômes réparés, tous mesurés sur un cas d'essai fabriqué :
      - un quadrillage vide disparaissait entièrement (ses traits réguliers
        étaient éliminés comme du bruit) : le modèle recevait des taches
        coloriées sans repère pour compter, d'où les aires inventées ;
      - un tableau ressortait en file de mots, sans plus dire quelle valeur
        allait avec quelle entrée (« 36 / 24 / 1 2 3 … 12 ») ;
      - ce même tableau repartait AUSSI en figure découpée : vu deux fois,
        une fois juste et une fois faux.
    Le serveur ne résout rien : il mesure et il le dit. Une grille produit sa
    forme (« quadrillage 8 colonnes x 6 lignes, case 10,0 mm »), son contenu
    si elle porte du texte, et le nombre de cases pleines de chaque figure —
    ce dernier UNIQUEMENT si les figures épousent la grille. Une forme
    oblique, un triangle, un décalage : rien n'est annoncé, et la raison du
    silence est transmise. Un comptage douteux est plus dangereux qu'absent.
    Module `grilles.py`, batterie `test_grilles_v210.py` (21 cas, à vide).
    Si le module est absent au déploiement, le pipeline retombe à l'identique
    sur le comportement 2.9.3 — aucune panne, seulement l'ancien défaut.

V2.9.3 (expérience du mode dégradé, 26 juillet 2026) :
  · PDF_B64_MAX : le plafond d'attachement du PDF converti (ODT) passe de
    150 000 à 4 000 000 caractères base64. Il doublait en silence le plafond
    du frontend : remonter celui du frontend seul ne changeait rien, le PDF
    ne quittait jamais le backend. Quand le plafond est dépassé, la réponse
    porte désormais « pdf_b64_skipped_bytes » — plus de rejet muet.

V2.9.2 (chantier ARASAAC seul — périmètre validé par Catherine) :
  · POST /pictos : {"mots": ["araignée", …], "lang": "fr"} → pour chaque mot,
    le meilleur pictogramme ARASAAC en data-URL PNG 300 px, prêt à injecter
    dans le HTML (l'export reste autonome, aucune dépendance réseau à
    l'impression). Recherche via bestsearch avec repli search ; normalisation
    des mots (minuscules, déterminants élidés « l'araignée » → « araignée ») ;
    langues fr/de/en… (code ARASAAC).
  · CACHE deux niveaux, zéro dépendance nouvelle (urllib stdlib) :
    mémoire (mot→id, id→png) + disque (ARASAAC_CACHE_DIR, défaut
    /tmp/arasaac_cache — les pictos sont immuables). Un mot introuvable est
    aussi mis en cache (négatif, TTL session) pour ne pas re-frapper l'API.
  · Limites de sécurité : 24 mots max par requête, timeout 8 s par appel
    ARASAAC, mots introuvables → null (le frontend applique ses replis
    émoji/sans-image). Meta d'attribution CC BY-NC-SA incluse dans chaque
    réponse (obligation de licence — ligne à imprimer en pied de fiche).
  · /health expose arasaac:"ok"/"unreachable" (sonde légère, 1 requête test
    mise en cache) pour vérifier l'accès réseau depuis Render au déploiement.

V2.9.1 (retour essai 44 — régression Pacôme corrigée) :
  · COMPOSITES = SCÈNES SEULEMENT : la garde par paires de la V2.9 laissait
    fusionner les grands rasters qui se CHEVAUCHENT (scans à marges
    blanches) — les dessins Pacôme/Momo ont refusionné (161×153 mm) en
    avalant le fragment « er, hululer… », le bug même de l'essai 39.
    Un composite exige désormais >= 3 membres tous petits (min-dim
    <= 120 pt) ; toute paire et tout composant contenant un grand raster
    sont dissous en rasters individuels (comportement V2.8 restauré pour
    les scans/photos/cliparts, composites conservés pour les scènes :
    pièces, billets, tas de tomates, vignettes).

V2.9 (audit essai 43, validé localement sur EvalSerie1maths5P.odt — 168 → 93 figures) :
  · CLUSTERING RASTER-RASTER : la règle V2.8 « rasters autonomes » pulvérisait
    les scènes composées de nombreux petits rasters (série 5P : 16 pièces de
    monnaie, 15 tas de tomates, ~19 billets, 17 vignettes… = 128 rasters) ;
    le plafond de miniatures du frontend saturait et le modèle déclarait
    « non disponibles » des images existantes (essai 43 : 154 figures
    inutilisées, 12 placeholders « coller ici »). Les rasters se regroupent
    désormais ENTRE EUX (_cluster_rasters, marge 16 pt calibrée sur les
    écarts mesurés : intra-scène 0-14,6 pt, figures distinctes ≥ 18 pt),
    jamais avec les vecteurs — l'acquis V2.8 est conservé.
  · GARDE PHOTOS : deux grands rasters (min-dim > 120 pt ≈ 42 mm) sans
    recouvrement ne fusionnent jamais — les photos côte à côte restent
    autonomes (essai 42 insectes : 15/15 photos, à ne pas casser).
  · SCISSION TABLEAU : un composite multi-rasters est re-scindé à chaque
    bordure horizontale de tableau qui le traverse sans couper aucun membre
    (p5 de la 5P : écarts intra/inter-collections indistinguables — seules
    les bordures séparent les 4 collections, retrouvées une à une : 1, 8,
    4 et 3 pièces).

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

# V2.9.3 — plafond d'attachement du PDF converti, en caractères base64.
# Doit rester égal à PDF_B64_MAX du frontend (tailoryv10_62.html) : un plafond
# plus bas ici filtrerait en amont, invisiblement, quoi que fasse le frontend.
PDF_B64_MAX = 4_000_000  # ~3 Mo de PDF, une trentaine de pages illustrées

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
    2. SCÈNES SEULEMENT (v2.9.1, retour essai 44) : un composite exige
       >= 3 membres TOUS petits (min-dim <= 120 pt ~ 42 mm). Tout composant
       contenant un grand raster (scan, photo, clipart) ou réduit à une paire
       est dissous en rasters individuels : les scans à marges blanches qui se
       chevauchent (Pacôme/Momo, essai 39 et 44) et les photos côte à côte
       (essai 42 insectes) ne fusionnent jamais.
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
            if _gap(rasters[i], rasters[j]) < margin:
                union(i, j)

    comps = {}
    for i in range(n):
        comps.setdefault(find(i), []).append(rasters[i])

    # v2.9.1 — FILTRE AU NIVEAU DU COMPOSANT (retour essai 44 : régression sur
    # Pacôme). La garde v2.9 par paires (« deux grands rasters SANS recouvrement
    # ne fusionnent pas ») laissait passer les scans à marges blanches qui se
    # CHEVAUCHENT — exactement le bug de l'essai 39 que la V2.8 avait éliminé :
    # les dessins de Pacôme et Momo ont refusionné en un composite de
    # 161×153 mm avalant le fragment « er, hululer… » de la banque de mots.
    # Nouvelle règle : un composite n'est légitime que pour une SCÈNE —
    # au moins 3 membres, TOUS petits (min-dim <= 120 pt ~ 42 mm : pièces,
    # billets, tas, vignettes). Tout composant contenant un grand raster, ou
    # réduit à une paire, est dissous en rasters individuels (comportement
    # V2.8). Les scans, photos et cliparts ne fusionnent donc plus jamais,
    # avec ou sans recouvrement.
    def _scene(members):
        return len(members) >= 3 and all(
            min(m.width, m.height) <= 120 for m in members)

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
        if not _scene(members):
            out.extend(members)  # v2.9.1 — pas une scène : chacun reste autonome
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


# V2.11 (chantier 6.8) — plancher des segments courts étiquetés : 7,0 mm.
# En dessous, rien ne passe, étiquette ou pas ; entre 7 et 14 mm, seulement
# sous les deux conditions du bloc 2c ; à partir de 14,1 mm (40 pt), régime
# inchangé.
SEUIL_COURT_ETIQUETE = 19.8  # 7,0 mm en points PDF

# V2.12 (chantier 4 — tirage essai_10127) — DISTANCE LETTRE → TRAIT.
# La V2.11 exigeait l'étiquette à moins de 12 pt (4 mm) du départ du trait.
# Mesuré le 29.07.2026 sur la page de mesure réellement utilisée (évaluation 3,
# « Partie 2 Géométrie », six segments a) à f) alignés sur des tabulations) :
#
#     d)  85,0 mm   →   7,6 pt      (seul à passer la porte de 12 pt)
#     f) 114,0 mm   →  16,6 pt
#     b) 110,0 mm   →  26,1 pt
#     c)  30,0 mm   →  27,1 pt
#     e)   8,0 mm   →  45,4 pt      ← le seul qui AVAIT BESOIN de la porte
#
# Les quatre longs entraient sans elle (≥ 40 pt). Le court, lui, était rejeté
# pour 33 pt de tabulation. La porte était calibrée sur une hypothèse, pas sur
# une mesure. Elle passe à 60 pt (21 mm), ce qui couvre une tabulation avec de
# la marge.
#
# CE QUI REMPLACE LA DISTANCE COMME GARDE-FOU. Élargir seul rouvrirait la porte
# aux soulignements. Deux conditions la referment :
#   — RIEN ENTRE LES DEUX (ci-dessous) : aucun mot ne s'intercale entre
#     l'étiquette et le départ du trait. Un mot souligné a son texte AU-DESSUS
#     du trait, pas une lettre isolée à sa gauche avec du vide entre les deux ;
#   — la condition de série du bloc 2c, inchangée : il faut un autre trait
#     étiqueté de longueur différente sur la page.
TOL_ETIQUETTE_PT = 60.0

_ETIQUETTE_SERIE = re.compile(r"^([A-Za-z]|\d{1,2})\s*[).]$")
_LETTRE_SEULE = re.compile(r"^([A-Za-z]|\d{1,2})$")


def _etiquette_de_serie(words, t):
    """V2.11 (chantier 6.8) — vrai si un mot-étiquette de série (lettre ou
    chiffre seul suivi d'une parenthèse ou d'un point : a) b) e) 1) 2.) se
    tient JUSTE À GAUCHE du trait. C'est la signature d'un segment d'exercice
    de mesure ; un tiret, une puce ou un soulignement n'est jamais étiqueté.
    À gauche seulement — une étiquette au-dessus ressemble trop à un mot
    souligné ou à un numérateur de fraction : cette porte reste fermée."""
    # V2.12 — une étiquette peut arriver en DEUX mots (« a )» écrit « a » puis
    # « ) », relevé sur la page de mesure réelle). On recompose avant de juger.
    etiquettes = []
    for i, w in enumerate(words):
        txt = w[4].strip()
        if _ETIQUETTE_SERIE.match(txt):
            etiquettes.append(w)
        elif _LETTRE_SEULE.match(txt) and i + 1 < len(words):
            nxt = words[i + 1]
            if nxt[4].strip() in (")", ".") and abs(nxt[1] - w[1]) < 3 \
               and 0 <= nxt[0] - w[2] <= 6:
                etiquettes.append((w[0], w[1], nxt[2], nxt[3], txt + nxt[4].strip()))

    for w in etiquettes:
        wx0, wy0, wx1, wy1 = w[:4]
        # à gauche : le mot finit avant le trait (chevauchement de 2 pt toléré),
        # à moins de TOL_ETIQUETTE_PT de son départ (V2.12 : 60 pt, mesuré)
        if not (wx1 <= t.x0 + 2 and t.x0 - wx1 <= TOL_ETIQUETTE_PT):
            continue
        # V2.12 — RIEN ENTRE LES DEUX. Ce qui sépare une étiquette de série d'un
        # mot souligné, ce n'est pas la distance : c'est le vide. Si un mot
        # s'intercale entre la lettre et le départ du trait, sur la même bande,
        # ce n'est pas une étiquette de segment.
        bande0, bande1 = min(wy0, t.y0) - 3, max(wy1, t.y1) + 3
        intercale = False
        for autre in words:
            if autre is w or not autre[4].strip():
                continue
            ax0, ay0, ax1, ay1 = autre[:4]
            if ax0 >= wx1 - 0.5 and ax1 <= t.x0 + 0.5 \
               and ay1 > bande0 and ay0 < bande1:
                intercale = True
                break
        if intercale:
            continue
        if t.width >= t.height:      # trait horizontal : son axe traverse le mot
            cy = (t.y0 + t.y1) / 2
            if wy0 - 3 <= cy <= wy1 + 3:
                return True
        else:                        # trait vertical : le mot à hauteur du trait
            cy = (wy0 + wy1) / 2
            if t.y0 - 3 <= cy <= t.y1 + 3:
                return True
    return False


def _collect_page_regions(page):
    """
    Régions candidates d'une page, en deux familles :
    - solids : images raster + tracés vectoriels « pleins »
    - thins  : lignes fines légitimes (segments à mesurer, lignes graduées)
      V2.11 : + segments courts (7–14 mm) étiquetés en série, bloc 2c
    Filtres : fonds de page (> 85 % de l'aire), séparateurs pleine largeur,
    micro-tracés, soulignés de texte, bordures de tableaux (non isolées),
    familles de lignes identiques (lignes d'écriture, grilles).
    v2.7 : + rangées de séparateurs partiels, + bandes de texte (en-têtes).
    """
    pw, ph = page.rect.width, page.rect.height
    page_area = pw * ph
    solids, thins = [], []
    courts = []  # V2.11 — candidats 7-14 mm, promus au bloc 2c ou jetés
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
                # Ligne fine : candidate d'office si assez longue pour être
                # une figure (segment, ligne graduée), pas un simple tiret.
                # V2.11 (chantier 6.8) — entre 7 et 14 mm, candidate SOUS
                # CONDITIONS, vérifiées au bloc 2c : un segment court qui
                # porte sa lettre dans un exercice de mesure est une vraie
                # figure (segment e de 8 mm, essai_10117).
                if max(r.width, r.height) >= 40:
                    thins.append(fitz.Rect(r))
                elif max(r.width, r.height) >= SEUIL_COURT_ETIQUETE:
                    courts.append(fitz.Rect(r))
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

    # 2c. V2.11 (chantier 6.8) — promotion des segments courts étiquetés.
    # Un trait de 7 à 14 mm n'entre qu'à deux conditions cumulées :
    #   1. une étiquette de série juste à sa gauche (a), e), 1), 2. …) ;
    #   2. un AUTRE trait étiqueté de longueur différente (> 5 %) sur la
    #      page — les segments d'un exercice de mesure ont des longueurs
    #      toutes différentes, c'est le principe de l'exercice ; les lignes
    #      de réponse d'une liste numérotée font toutes la même : dehors.
    # Les promus subissent ensuite les mêmes filtres anti-bruit que les
    # longs (isolement, souligné, familles) puis le filtre v2.5 des pages
    # de mesure — la porte ne s'ouvre que d'un cran, pas en grand.
    if courts:
        candidats = [t for t in courts
                     if _etiquette_de_serie(words_v27, t)]
        if candidats:
            longueurs = [max(t.width, t.height)
                         for t in list(thins) + candidats
                         if _etiquette_de_serie(words_v27, t)]
            for t in candidats:
                L = max(t.width, t.height)
                if any(abs(L - L2) > 0.05 * max(L, L2)
                       for L2 in longueurs):
                    thins.append(t)

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


# ─────────────────────────────────────────────
# V2.10 — grilles et formes composites
# Import tolérant : si le module n'a pas été déployé à côté de app.py, le
# pipeline garde exactement son comportement 2.9.3. Pas de panne, pas de
# demi-mesure silencieuse : l'état est visible sur /health.
# ─────────────────────────────────────────────
try:
    import grilles as _grilles
    GRILLES_ACTIVES = True
except Exception:  # module absent ou illisible
    _grilles = None
    GRILLES_ACTIVES = False

try:
    import formes as _formes
    FORMES_ACTIVES = True
except Exception:
    _formes = None
    FORMES_ACTIVES = False


# ══ v2.14 — LE GRAS DE LA SOURCE SURVIT À LA LECTURE DU PDF ═══════════════════════════
# Chantier ouvert le 30.07.2026. La règle de mise en évidence (frontend v10.151) dit que
# le gras de la source est reproduit par défaut ; mais quand la source arrive en PDF, le
# serveur lisait `page.get_text("text")`, qui rend une chaîne PLATE : le gras était perdu
# avant même que le modèle le voie. Aucune règle de conservation ne pouvait donc être
# tenue sur ce chemin — la perte était silencieuse, le pire cas de la doctrine.
#
# Ce que fait la fonction : elle relit la page span par span et entoure les passages en
# gras de deux marqueurs — ⟦gras⟧ … ⟦/gras⟧ — que le frontend sait reconvertir en <strong>
# (filet v10.158 b). Le choix de marqueurs à crochets blancs plutôt que d'étoiles ou de
# balises est délibéré : ils n'existent dans aucune source scolaire, ne se confondent avec
# aucune syntaxe, et s'ils survivent par accident jusqu'à la feuille élève le filet du
# frontend les rattrape et le signale.
#
# BORNES CONNUES, écrites ici plutôt qu'oubliées :
#  - le chemin « grilles » (texte_avec_grilles) construit son texte lui-même et n'est pas
#    couvert : sur une page à grilles, le gras reste perdu. À traiter en v2.15.
#  - PyMuPDF signale le gras par le bit 4 des drapeaux de span (valeur 16) ; certaines
#    polices déclarent le gras dans leur NOM seulement (« …-Bold »), d'où le second test.
_GRAS_OUV = "\u27e6gras\u27e7"
_GRAS_FER = "\u27e6/gras\u27e7"


def _span_est_gras(span):
    if (span.get("flags", 0) & 16):
        return True
    nom = (span.get("font", "") or "").lower()
    return ("bold" in nom) or ("black" in nom) or ("heavy" in nom)


def _texte_gras(page):
    """Texte de la page, gras marqué. Retombe sur get_text('text') au moindre doute."""
    try:
        d = page.get_text("dict")
    except Exception:
        return page.get_text("text")
    blocs = []
    for b in d.get("blocks", []):
        if b.get("type", 0) != 0:
            continue
        lignes = []
        for l in b.get("lines", []):
            morceaux, ouvert = [], False
            for s in l.get("spans", []):
                txt = s.get("text", "")
                if not txt:
                    continue
                g = _span_est_gras(s)
                if g and not ouvert:
                    morceaux.append(_GRAS_OUV); ouvert = True
                elif not g and ouvert:
                    morceaux.append(_GRAS_FER); ouvert = False
                morceaux.append(txt)
            if ouvert:
                morceaux.append(_GRAS_FER)
            ligne = "".join(morceaux)
            # un marqueur qui n'entoure que des espaces n'apporte rien et gêne la relecture
            ligne = ligne.replace(_GRAS_OUV + " " + _GRAS_FER, " ")
            ligne = ligne.replace(_GRAS_OUV + _GRAS_FER, "")
            lignes.append(ligne)
        if lignes:
            blocs.append("\n".join(lignes))
    if not blocs:
        return page.get_text("text")
    return "\n".join(blocs) + "\n"


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
    grilles_inventaire = []
    formes_inventaire = []
    idx = 0

    for pno, page in enumerate(doc):
        pw, ph = page.rect.width, page.rect.height
        page_area = pw * ph

        solids, thins, rasters = _collect_page_regions(page)

        # V2.10 — grilles de la page : forme, contenu, comptage des cases.
        # `zones_grilles` liste les tableaux de texte, qui ne doivent plus
        # repartir en figure découpée : leur contenu est déjà transmis en
        # clair, et l'image du même tableau faisait doublon.
        # V2.10 — formes composites : une figure faite de plusieurs
        # rectangles ne doit plus partir comme un bloc plein. Sa
        # décomposition est décrite, son aire réelle mesurée.
        formes_page = []
        if FORMES_ACTIVES:
            try:
                formes_page = _formes.analyser_formes(page)
            except Exception:
                formes_page = []
        blocs_formes = [(f["rect"].y0, f["rect"].x0, _formes.decrire(f))
                        for f in formes_page]

        grilles_page, zones_grilles = [], []
        if GRILLES_ACTIVES:
            try:
                analyse = _grilles.analyser_page(page)
                grilles_page = analyse["grilles"]
                zones_grilles = analyse["zones_grilles"]
                texte_page = _grilles.texte_avec_grilles(
                    page, grilles_page, blocs_formes)
            except Exception:
                texte_page = _texte_gras(page)      # v2.14
        else:
            texte_page = _texte_gras(page)          # v2.14
            if blocs_formes:
                texte_page += "\n" + "\n".join(b[2] for b in sorted(blocs_formes))

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
        # V2.11 — les segments courts étiquetés, promus au bloc 2c, passent
        # au même titre que les longs : la porte ne s'ouvre que pour EUX,
        # le plancher des 14 mm reste la règle pour tout le reste.
        cles_courts = set()
        for t in thins:
            if max(t.width, t.height) < 40:
                tt = t & page.rect
                cles_courts.add((round(tt.x0), round(tt.y0),
                                 round(tt.x1), round(tt.y1)))
        keep, seen = [], set()
        for r in final:
            r = r & page.rect
            key = (round(r.x0), round(r.y0), round(r.x1), round(r.y1))
            if key in seen:
                continue
            seen.add(key)
            full = r.width >= 24 and r.height >= 24
            slim = min(r.width, r.height) < 24 and (
                max(r.width, r.height) >= 40 or key in cles_courts)
            if not (full or slim):
                continue
            if (r.width * r.height) > 0.85 * page_area:
                continue
            # V2.10 — un tableau de texte n'est pas une figure : son contenu
            # part déjà en clair, l'image ferait doublon. MAIS beaucoup de
            # sources se servent d'un tableau comme d'une mise en page et y
            # posent leurs dessins : une figure logée dans une cellule reste
            # une figure, et perdre un dessin est une faute grave.
            dans_tableau = any(not (r & z).is_empty
                               and (r & z).get_area() >= 0.6 * r.get_area()
                               for z in zones_grilles)
            if dans_tableau and not any(
                    not (r & ra).is_empty
                    and (r & ra).get_area() >= 0.5 * min(r.get_area(),
                                                         ra.get_area())
                    for ra in rasters):
                continue
            keep.append(r)

        # V2.10 — le quadrillage part d'un seul tenant. Jusqu'ici ses traits
        # réguliers étaient éliminés comme du bruit et seules les zones
        # coloriées survivaient : le modèle recevait des taches sans repère,
        # et l'élève une figure sans quadrillage à compter. On ajoute donc le
        # cadre complet, et on retire les morceaux qu'il contient — ils y sont
        # déjà, en place.
        for g in grilles_page:
            if g["nature"] != "quadrillage":
                continue
            cadre = fitz.Rect(g["rect"]) & page.rect
            if cadre.is_empty or cadre.width < 24 or cadre.height < 24:
                continue
            keep = [r for r in keep
                    if not (cadre.contains(r)
                            or (not (r & cadre).is_empty
                                and (r & cadre).get_area() >= 0.6 * r.get_area()))]
            keep.append(cadre)

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

        for g in grilles_page:
            grilles_inventaire.append({
                "page": pno + 1,
                "nature": g["nature"],
                "lignes": g["lignes"],
                "colonnes": g["colonnes"],
                "case_l_mm": g["case_l_mm"],
                "case_h_mm": g["case_h_mm"],
                "zones": [{"etiquette": z["etiquette"], "cases": z["cases"],
                           "composite": z["composite"]} for z in g["zones"]],
                "comptage_refuse": g["comptage_refuse"],
            })

        for f in formes_page:
            formes_inventaire.append({
                "page": pno + 1,
                "etiquette": f["etiquette"],
                "membres": f["membres"],
                "encombrement_mm2": f["encombrement_mm2"],
                "aire_mm2": f["aire_mm2"],
                "aire_refusee": f["aire_refusee"],
            })

        full_text.append(texte_page)

    doc.close()
    return {
        "filename": filename,
        "pdf_mode": True,
        "num_exercises": 0,
        "num_images": len(images),
        "images": images,
        "text": "\n".join(full_text),
        "exercises": [],
        # V2.10 — inventaire des grilles rencontrées, pour le journal et pour
        # le vérificateur : nature, dimensions, cases comptées ou refus motivé.
        "grilles": grilles_inventaire,
        "grilles_actives": GRILLES_ACTIVES,
        # Idem pour les formes composites : membres, encombrement, aire
        # réelle ou refus motivé.
        "formes": formes_inventaire,
        "formes_actives": FORMES_ACTIVES,
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
        # Joindre le PDF converti pour qu'il soit transmis à l'IA comme document
        # natif. V2.9.3 : le plafond passe de 150 000 à 4 000 000 caractères
        # base64 (~3 Mo de PDF), aligné sur PDF_B64_MAX du frontend v10.62.
        # L'ancienne valeur écartait SILENCIEUSEMENT les documents illustrés
        # (EvalSerie1maths6P : 1 474 ko, 1 341 % du plafond) : le modèle ne
        # recevait alors que le texte et les miniatures, sans jamais voir la page.
        # Les deux plafonds doivent rester égaux : c'est le frontend qui décide,
        # le backend ne doit plus filtrer en amont sans le dire.
        if len(pdf_bytes) * 4 / 3 < PDF_B64_MAX:
            result["pdf_b64"] = base64.b64encode(pdf_bytes).decode()
        else:
            result["pdf_b64_skipped_bytes"] = len(pdf_bytes)
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
# ─────────────────────────────────────────────
# ENDPOINT : /pictos  (V2.9.2 — ARASAAC)
# Mots-clés → pictogrammes ARASAAC en data-URL
# ─────────────────────────────────────────────
import urllib.request
import urllib.parse
import unicodedata

ARASAAC_API = "https://api.arasaac.org/v1/pictograms"
ARASAAC_STATIC = "https://static.arasaac.org/pictograms/{id}/{id}_300.png"
ARASAAC_CACHE_DIR = os.environ.get("ARASAAC_CACHE_DIR", os.path.join(tempfile.gettempdir(), "arasaac_cache"))
ARASAAC_TIMEOUT = 8
ARASAAC_MAX_MOTS = 24
ARASAAC_ATTRIBUTION = ("Pictogrammes : ARASAAC (arasaac.org) — auteur Sergio Palao, "
                       "propriété du Gouvernement d'Aragon, licence CC BY-NC-SA")

_arasaac_ids: dict = {}      # "fr:araignée" -> id ARASAAC (ou -1 = introuvable, cache négatif)
_arasaac_png: dict = {}      # id -> bytes PNG

_ARASAAC_DETS = ("l'", "d'", "le ", "la ", "les ", "un ", "une ", "des ", "du ", "de ")

def _arasaac_norm(mot: str) -> str:
    """Normalise un mot-clé : minuscules, espaces réduits, déterminant élidé."""
    m = (mot or "").strip().lower()
    m = re.sub(r"\s+", " ", m)
    for det in _ARASAAC_DETS:
        if m.startswith(det):
            m = m[len(det):]
            break
    return m.strip()

def _arasaac_http_json(url: str):
    req = urllib.request.Request(url, headers={"User-Agent": "Tailory/2.9.2"})
    with urllib.request.urlopen(req, timeout=ARASAAC_TIMEOUT) as r:
        return json.loads(r.read().decode("utf-8"))

def _arasaac_http_bytes(url: str) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": "Tailory/2.9.2"})
    with urllib.request.urlopen(req, timeout=ARASAAC_TIMEOUT) as r:
        return r.read()

def _arasaac_search_id(mot: str, lang: str):
    """bestsearch puis repli search ; retourne l'id du 1er pictogramme ou None."""
    q = urllib.parse.quote(mot)
    for route in ("bestsearch", "search"):
        try:
            data = _arasaac_http_json(f"{ARASAAC_API}/{lang}/{route}/{q}")
            if isinstance(data, list) and data and isinstance(data[0], dict) and "_id" in data[0]:
                return int(data[0]["_id"])
        except Exception:
            continue
    return None

def _arasaac_png_by_id(pid: int):
    """PNG 300 px d'un picto — cache mémoire puis disque puis réseau."""
    if pid in _arasaac_png:
        return _arasaac_png[pid]
    os.makedirs(ARASAAC_CACHE_DIR, exist_ok=True)
    fp = os.path.join(ARASAAC_CACHE_DIR, f"{pid}_300.png")
    if os.path.exists(fp):
        with open(fp, "rb") as f:
            b = f.read()
        _arasaac_png[pid] = b
        return b
    try:
        b = _arasaac_http_bytes(ARASAAC_STATIC.format(id=pid))
    except Exception:
        return None
    if not b or len(b) < 100:
        return None
    with open(fp, "wb") as f:
        f.write(b)
    _arasaac_png[pid] = b
    return b

def _arasaac_resolve(mot: str, lang: str):
    """Mot → {id, dataurl} ou None. Caches négatifs inclus."""
    key = f"{lang}:{mot}"
    pid = _arasaac_ids.get(key)
    if pid == -1:
        return None
    if pid is None:
        pid = _arasaac_search_id(mot, lang)
        _arasaac_ids[key] = pid if pid is not None else -1
        if pid is None:
            return None
    b = _arasaac_png_by_id(pid)
    if b is None:
        return None
    return {"id": pid, "dataurl": "data:image/png;base64," + base64.b64encode(b).decode("ascii")}

@app.post("/pictos")
async def pictos(request: Request):
    """
    {"mots": ["araignée", "l'insecte", …], "lang": "fr"}
    → {"pictos": {"araignée": {"id":…, "dataurl":"data:image/png;base64,…"} | null, …},
       "attribution": "…CC BY-NC-SA…", "lang": "fr"}
    Un mot introuvable vaut null : le frontend applique ses replis (émoji, sans image).
    """
    body = await request.json()
    mots = body.get("mots") or []
    lang = (body.get("lang") or "fr").strip().lower()[:2]
    if not isinstance(mots, list) or not mots:
        raise HTTPException(400, "mots (liste non vide) requis")
    if len(mots) > ARASAAC_MAX_MOTS:
        raise HTTPException(400, f"maximum {ARASAAC_MAX_MOTS} mots par requête")
    out = {}
    for mot_brut in mots:
        mot = _arasaac_norm(str(mot_brut))
        if not mot:
            out[str(mot_brut)] = None
            continue
        out[str(mot_brut)] = _arasaac_resolve(mot, lang)
    return {"pictos": out, "attribution": ARASAAC_ATTRIBUTION, "lang": lang}


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

    # ══ v2.13 — LE CACHE SURVIT À LA DURÉE D'UNE PARTIE ═══════════════════════════════════
    # Constat des tirages (pied de page essai_10138) : « 10 009 lus + 0 relus +
    # 160 006 mis en cache » — on PAYE la mise en cache à chaque partie et on ne relit
    # JAMAIS. Deux causes dans le code v2.12 :
    #   1. le cache éphémère vit 5 minutes, or l'écriture d'une partie par le modèle en
    #      prend davantage : à la partie suivante, le cache est déjà mort ;
    #   2. seul le prompt système portait un repère de cache — le DOCUMENT (le gros du
    #      volume), envoyé dans le premier message, n'en portait aucun.
    # Remèdes : durée longue (1 h) demandée sur chaque repère, avec REPLI AUTOMATIQUE en
    # durée courte si le serveur refuse la syntaxe (aucune génération ne casse, la réponse
    # dit la durée réellement servie) ; et un repère posé sur le dernier bloc du premier
    # message utilisateur, là où vit le document. La réponse renvoie désormais TOUS les
    # compteurs (lus / relus / mis en cache) : la facturation exige des coûts prouvables.
    for m in messages:
        if m.get("role") == "user":
            c = m.get("content")
            if isinstance(c, str):
                m["content"] = [{"type": "text", "text": c}]
            if isinstance(m["content"], list) and m["content"]:
                dernier = m["content"][-1]
                if isinstance(dernier, dict) and dernier.get("type") in ("text", "document", "image"):
                    dernier.setdefault("cache_control", {"type": "ephemeral"})
            break

    def _pose_ttl(duree):
        sys_blocs = [{"type": "text", "text": system,
                      "cache_control": {"type": "ephemeral"}}] if system else []
        for blocs in ([b for b in sys_blocs],):
            pass
        tous = (sys_blocs or []) + [b for m in messages if isinstance(m.get("content"), list)
                                    for b in m["content"] if isinstance(b, dict) and "cache_control" in b]
        for b in tous:
            if duree == "1h":
                b["cache_control"] = {"type": "ephemeral", "ttl": "1h"}
            else:
                b["cache_control"] = {"type": "ephemeral"}
        return sys_blocs

    # v2.13 — le client peut IMPOSER la durée ("cache_ttl": "5m" ou "1h") : le
    # frontend sait mieux que le serveur si plusieurs parties sont probables
    # (taille du document). Valeur absente ou inconnue : durée longue par défaut.
    duree_cache = body.get("cache_ttl") if body.get("cache_ttl") in ("5m", "1h") else "1h"

    # Retry avec backoff exponentiel
    max_retries = 3
    for attempt in range(max_retries):
        try:
            sys_blocs = _pose_ttl(duree_cache)
            extra = {"anthropic-beta": "extended-cache-ttl-2025-04-11"} if duree_cache == "1h" else {}
            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                system=sys_blocs,
                messages=messages,
                extra_headers=extra
            )
            text = "".join(
                block.text for block in response.content
                if hasattr(block, "text")
            )
            u = response.usage
            return {
                "content": text,
                "usage": {
                    "input_tokens": u.input_tokens,
                    "output_tokens": u.output_tokens,
                    "cache_creation_input_tokens": getattr(u, "cache_creation_input_tokens", 0) or 0,
                    "cache_read_input_tokens": getattr(u, "cache_read_input_tokens", 0) or 0
                },
                "cache_ttl": duree_cache
            }
        except anthropic.RateLimitError as e:
            if attempt < max_retries - 1:
                wait = 2 ** (attempt + 1)  # 2s, 4s, 8s
                time.sleep(wait)
                continue
            raise HTTPException(429, f"Rate limit après {max_retries} tentatives : {e}")
        except anthropic.BadRequestError as e:
            # v2.13 — repli : si le refus vise la durée du cache (syntaxe ttl ou en-tête
            # inconnus du serveur), on repart en durée courte, une seule fois, sans
            # consommer de tentative — la génération ne casse jamais pour du cache.
            msg = str(e).lower()
            if duree_cache == "1h" and ("ttl" in msg or "cache" in msg or "beta" in msg):
                duree_cache = "5m"
                continue
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
_arasaac_probe_cache = {"t": 0.0, "state": "unknown"}

@app.get("/health")
def health():
    # V2.9.2 — sonde ARASAAC (mise en cache 10 min : /health est appelé à chaque
    # chargement du frontend, on ne frappe pas l'API à chaque fois)
    now = time.time()
    if now - _arasaac_probe_cache["t"] > 600:
        try:
            data = _arasaac_http_json(f"{ARASAAC_API}/fr/bestsearch/maison")
            _arasaac_probe_cache["state"] = "ok" if isinstance(data, list) and data else "empty"
        except Exception:
            _arasaac_probe_cache["state"] = "unreachable"
        _arasaac_probe_cache["t"] = now
    # V2.10 — la version du module grilles est remontée telle qu'elle est
    # DANS le fichier déployé, jamais recopiée à la main : c'est ce qui
    # permet de savoir, en ligne, laquelle tourne vraiment.
    return {"status": "ok", "version": "2.14",
            "grilles": (getattr(_grilles, "VERSION", "inconnue")
                        if GRILLES_ACTIVES else "absent"),
            "formes": (getattr(_formes, "VERSION", "inconnue")
                       if FORMES_ACTIVES else "absent"),
            "arasaac": _arasaac_probe_cache["state"]}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
