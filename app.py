"""
Tailory Backend — Pipeline documentaire pédagogique
FastAPI + python-docx + Anthropic

Le numéro de version : UNE seule déclaration, la constante VERSION ci-dessous ; /health la lit.
L'histoire des versions, le pourquoi et les mesures : JOURNAL BACKEND (un seul journal, une entrée par version).

Endpoints:
  POST /parse    → PDF seul (tout autre fichier est refusé, DEC-1249) → structure JSON pédagogique
  (v2.83 — /parse et /figures_vues exigent le code personnel de l'enseignante (en-tête X-Tailory-Code) et tiennent ses limites : voir LA PORTE)
  GET  /health   → vérification
"""

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import JSONResponse  # v2.83 — LA PORTE
from fastapi.middleware.cors import CORSMiddleware
import anthropic
import base64
import io
import json
import os
import re
import uuid

# V2.9.3 — plafond d'attachement d'un PDF, en caractères base64 : le même que celui de la page (PDF_B64_MAX). Gardé (DEC-3262, une sûreté qui ne coûte rien) ;
# aucune route du serveur ne l'emploie plus depuis le retrait de la route Word (8.6, DEC-1249).
PDF_B64_MAX = 4_000_000  # ~3 Mo de PDF, une trentaine de pages illustrées

# Seule déclaration du numéro de version du serveur : /health la lit ; elle n'est recopiée nulle part (voir JOURNAL BACKEND v2.23).
# Ce fichier ne porte que ce qu'il fait ; le pourquoi et les mesures de chaque version sont au JOURNAL BACKEND.
VERSION = "2.84"

# ═══ v2.34 — C11 : UN CADRE SANS DESSIN N'EST PAS UNE FIGURE ═══
# Condition : dans le cadre découpé, aucun objet dessiné. Effet : le cadre n'est pas transmis au modèle.
# Ce que la fonction lit : page.get_drawings() (les tracés), page.get_text("dict") (blocs de texte et d'image), rasters (les images collées,
# déjà constituées) — deux proportions géométriques, aucune déclaration d'une autre couche. (voir JOURNAL BACKEND v2.34)
C11_AIRE_MIN_PT2 = 20.0    # sous cette aire, un tracé est un filet, pas un objet
C11_PART_TEXTE = 0.10      # la part de texte au-delà de laquelle un tracé est le contenant d'un texte (voir JOURNAL BACKEND v2.34)
C11_PLAFOND_CADRE = 0.60   # un tracé qui couvre plus de 60 % du cadre est le cadre — seulement si le cadre porte du texte


def _v234_porte_un_dessin(page, clip, rasters):
    """Vrai si le cadre porte au moins un objet dessiné (image collée ou
    tracé qui n'est pas le contenant d'un texte)."""
    dic = page.get_text("dict")
    spans = [sp for b in dic["blocks"] if b["type"] == 0
             for l in b["lines"] for sp in l["spans"]]
    dedans = [sp for sp in spans if clip.intersects(fitz.Rect(sp["bbox"]))]
    # une image collée est toujours un dessin
    for ra in rasters:
        if not (ra & clip).is_empty and (ra & clip).get_area() >= 0.5 * ra.get_area():
            return True
    aire_clip = clip.get_area()
    for d in page.get_drawings():
        r = fitz.Rect(d["rect"])
        if not clip.contains(r) or r.width < 3 or r.height < 3:
            continue
        if r.get_area() < C11_AIRE_MIN_PT2:
            continue
        # un cadre qui n'est QU'une forme, c'est la forme elle-même : le
        # plafond ne s'applique que si le cadre porte du texte.
        if dedans and r.get_area() > C11_PLAFOND_CADRE * aire_clip:
            continue
        aire_txt = sum((fitz.Rect(sp["bbox"]) & r).get_area()
                       for sp in spans if r.intersects(fitz.Rect(sp["bbox"])))
        if aire_txt / r.get_area() < C11_PART_TEXTE:
            return True
    return False

# ═══ v2.46 — LA FRONTIÈRE : « PORTE UN DESSIN OU UNE COULEUR » ═══
# Deux lectures, déclarées, qui ne comptent rien :
#   · PORTE UN DESSIN — au moins une image raster dont le CENTRE est dans le rectangle de la grille ;
#   · PORTE UNE COULEUR — un remplissage qui n'est ni noir (toutes composantes ≤ 0,03) ni blanc (toutes ≥ 0,97), ET plus large qu'un
#     trait (1 mm).
# Ce qu'elle ne fait pas : elle ne lit pas les tracés au trait ; un raster à cheval sur le bord, centre dehors, n'est pas vu.
# (voir JOURNAL BACKEND v2.46)
_V246_SEUIL_TRAIT_MM = 1.0   # (2) — plus large qu'un trait : petit côté > 1 mm
_V246_NOIR = 0.03            # (1) — noir : toutes composantes ≤ 0,03
_V246_BLANC = 0.97           # (1) — blanc : toutes composantes ≥ 0,97


def _v246_porte_dessin_ou_couleur(page, rect):
    """Vrai si la grille `rect` porte un dessin (une image posée dont le centre est dedans — les images de la page
    telles quelles, pas les régions regroupées) ou une couleur (remplissage ni noir ni blanc, plus large qu'un trait)."""
    for img in page.get_images(full=True):
        for ra in page.get_image_rects(img[0]):
            if rect.contains(fitz.Point((ra.x0 + ra.x1) / 2, (ra.y0 + ra.y1) / 2)):
                return True
    aire = rect.get_area()
    for d in page.get_drawings():
        fill = d.get("fill")
        if not fill:
            continue
        r = d.get("rect")
        if r is None:
            continue
        r = fitz.Rect(r)
        if r.is_empty or r.get_area() >= aire * 1.01:
            continue                       # le fond de la grille ou de la page
        if not rect.contains(fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)):
            continue
        comp = [float(v) for v in fill]
        if all(v <= _V246_NOIR for v in comp) or all(v >= _V246_BLANC for v in comp):
            continue                       # (1) noir ou blanc : pas une couleur
        if min(r.width, r.height) * 25.4 / 72 <= _V246_SEUIL_TRAIT_MM:
            continue                       # (2) un trait, même coloré, reste un trait
        return True
    return False


# v2.28 — LE SERVEUR SIGNE SON CONTENU. Condition : le module se charge. Effet : l'empreinte du fichier lui-même est calculée
# UNE fois et servie par /health. (voir JOURNAL BACKEND v2.28)
import hashlib as _hl
try:
    with open(__file__, "rb") as _f:
        EMPREINTE = _hl.sha256(_f.read()).hexdigest()[:12]
except Exception:
    EMPREINTE = "illisible"
# ═══ v2.44 — LE SERVEUR DIT S'IL EST TOUJOURS LE MÊME ═══
# Une valeur fabriquée AU DÉMARRAGE DU PROCESSUS, servie par /health : elle change chaque fois que le processus meurt et repart.
# Ce n'est pas RENDER_INSTANCE_ID (qui change au remplacement de l'instance) : celui-là est servi en plus, comme information.
# Limite : avec plusieurs processus, chacun aurait la sienne (aujourd'hui WEB_CONCURRENCY=1). (voir JOURNAL BACKEND v2.44)
DEMARRAGE = uuid.uuid4().hex[:12]
INSTANCE_HEBERGEUR = os.environ.get("RENDER_INSTANCE_ID") or "non fourni"

# v2.29 — LE SERVEUR ANNONCE AVEC QUOI IL LIT. Condition : le module se charge. Effet : la version de la bibliothèque
# d'extraction est lue UNE fois et servie par /health. (voir JOURNAL BACKEND v2.29)
try:
    import pymupdf as _pm
    LECTEUR_PDF = "pymupdf " + getattr(_pm, "__version__", "?")
except Exception:
    try:
        LECTEUR_PDF = (fitz.__doc__ or "pymupdf ?").split(":")[0].strip().lower()
    except Exception:
        LECTEUR_PDF = "illisible"

app = FastAPI(title="Tailory Backend V2")

# ═══ v2.83 — LA PORTE : LE CODE PERSONNEL ET LES LIMITES (élément 8 de l'ouverture, étapes 2 et 3 de l'audit ; DEC-3490 à DEC-3493) ═══
# Sur /parse et /figures_vues, AVANT que le fichier soit lu : l'adresse bloquée après trop d'essais faux ; le code (en-tête X-Tailory-Code) absent,
# inconnu, retiré ou expiré ; la taille annoncée ; le crédit du mois du code ; le plafond du jour pour tous ; les analyses du jour du code ; une analyse
# du même code déjà en cours. La liste des codes : la variable secrète TAILORY_CODES (JSON : nom, empreinte = SHA-256 du code normalisé, fin, actif),
# relue à chaque appel ; absente ou illisible, tout est refusé. Le compte des dépenses : TAILORY_COMPTE_DIR (le disque Render) ; absent, tout est refusé.
# Posée AVANT la couche CORS dans le code, donc SOUS elle à l'exécution : la page lit les refus. Le pourquoi et les épreuves : JOURNAL BACKEND v2.83.
import threading as _thp, time as _tp, datetime as _dtp
from zoneinfo import ZoneInfo as _ZIp
_P_ROUTES = ("/parse", "/figures_vues")
_P_TAILLE_MAX = 20 * 1024 * 1024          # DEC-3492 : 20 Mo
_P_PAGES_MAX = 30                         # DEC-3492 : 30 pages
_P_MOIS_MAX_USD = 15.0                    # DEC-3492 : 15 $ par mois par code
_P_JOUR_TOUS_MAX_USD = 10.0               # DEC-3492 : 10 $ par jour pour tous
_P_ANALYSES_JOUR_MAX = 20                 # DEC-3492 : 20 feuilles par jour par code (compté en analyses)
_P_ESSAIS_FAUX_MAX, _P_ESSAIS_FENETRE_S = 10, 600   # au-delà de 10 codes faux en 10 min, l'adresse attend 10 min
_P_PRIX = {"claude-sonnet-5-5": (2.0, 10.0), "claude-opus-5-5": (4.0, 20.0)}   # $ par million de jetons (entrée, sortie) : les prix du projet (garde_pf1.py)
_P_VERROU = _thp.Lock()
_P_EN_COURS = set()
_P_FAUX = {}

def _p_norm(code):
    return "".join(ch for ch in (code or "").upper() if ch.isalnum())

def _p_codes():
    try:
        L = json.loads(os.environ.get("TAILORY_CODES", ""))
        return L if isinstance(L, list) else None
    except Exception:
        return None

def _p_maintenant():
    return _dtp.datetime.now(_ZIp("Europe/Zurich"))

def _p_compte_chemin():
    d = os.environ.get("TAILORY_COMPTE_DIR", "/var/data/tailory-compte")
    return os.path.join(d, "compte.json") if os.path.isdir(d) and os.access(d, os.W_OK) else None

def _p_compte_lire(ch):
    try:
        with open(ch, encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def _p_compte_ecrire(ch, c):
    tmp = ch + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False)
    os.replace(tmp, ch)

def _p_refus(statut, raison, phrase):
    return JSONResponse(status_code=statut, content={"detail": phrase, "raison": raison})

def _p_juge(request):
    """Rend (nom du code, None) si la porte s'ouvre, sinon (None, la réponse de refus). Ne lit pas le corps de la requête."""
    ip = (request.client.host if request.client else "?")
    t = _tp.time()
    with _P_VERROU:
        L = [x for x in _P_FAUX.get(ip, []) if t - x < _P_ESSAIS_FENETRE_S]
        _P_FAUX[ip] = L
        if len(L) > _P_ESSAIS_FAUX_MAX:
            return None, _p_refus(429, "essais", "Trop de codes faux depuis cet ordinateur : réessayez dans quelques minutes.")
    codes = _p_codes()
    if codes is None:
        return None, _p_refus(503, "acces_non_configure", "L'accès à Tailory n'est pas encore ouvert sur ce serveur.")
    code = _p_norm(request.headers.get("x-tailory-code"))
    if not code:
        return None, _p_refus(401, "absent", "Entrez votre code Tailory pour continuer.")
    emp = _hl.sha256(code.encode("utf-8")).hexdigest()
    fiche = next((c for c in codes if isinstance(c, dict) and c.get("empreinte") == emp), None)
    if fiche is None:
        with _P_VERROU:
            _P_FAUX.setdefault(ip, []).append(t)
        return None, _p_refus(401, "inconnu", "Ce code n'est pas reconnu : vérifiez-le, ou demandez-en un.")
    nom = str(fiche.get("nom") or "?")
    if not fiche.get("actif", False):
        return None, _p_refus(403, "retire", "Ce code n'est plus valable : demandez-en un nouveau.")
    fin = str(fiche.get("fin") or "")
    auj = _p_maintenant().date().isoformat()
    if not fin or fin < auj:
        return None, _p_refus(403, "expire", "Votre code a expiré le %s : demandez sa prolongation." % (fin or "?"))
    try:
        taille = int(request.headers.get("content-length", ""))
    except ValueError:
        return None, _p_refus(411, "taille_inconnue", "La taille du document n'est pas annoncée : réessayez.")
    if taille > _P_TAILLE_MAX:
        return None, _p_refus(413, "taille", "Le document dépasse 20 Mo : divisez-le, ou enregistrez-le en PDF plus léger.")
    ch = _p_compte_chemin()
    if ch is None:
        return None, _p_refus(503, "compte_indisponible", "Le compte des dépenses est indisponible : l'analyse ne part pas. Réessayez plus tard.")
    mois, jour = auj[:7], auj
    with _P_VERROU:
        c = _p_compte_lire(ch)
        if c.get("mois", {}).get(mois, {}).get(nom, 0.0) >= _P_MOIS_MAX_USD:
            return None, _p_refus(402, "credit_mois", "Votre crédit du mois est épuisé.")
        if c.get("jour", {}).get(jour, 0.0) >= _P_JOUR_TOUS_MAX_USD:
            return None, _p_refus(503, "plafond_jour", "Tailory a atteint sa limite du jour : réessayez demain.")
        if request.url.path == "/parse" and c.get("analyses", {}).get(jour, {}).get(nom, 0) >= _P_ANALYSES_JOUR_MAX:
            return None, _p_refus(429, "analyses_jour", "Vous avez atteint les 20 documents du jour : réessayez demain.")
        if nom in _P_EN_COURS:
            return None, _p_refus(429, "en_cours", "Une analyse est déjà en cours : attendez qu'elle finisse.")
        _P_EN_COURS.add(nom)
        if request.url.path == "/parse":
            c.setdefault("analyses", {}).setdefault(jour, {})
            c["analyses"][jour][nom] = c["analyses"][jour].get(nom, 0) + 1
            _p_compte_ecrire(ch, c)
    return nom, None

def _p_depense(nom, result):
    """Après une analyse : le prix recalculé de ses appels au modèle (les jetons rendus dans `result`), ajouté au mois du code et au jour de tous."""
    usd = 0.0
    for cle, modele in (("corrige_lecture", None), ("exercices_lecture", None), ("symbiose", _S243_MODELE)):
        b = (result or {}).get(cle) or {}
        m = modele or b.get("lecteur")
        if m in _P_PRIX:
            pe, ps = _P_PRIX[m]
            usd += (float(b.get("jetons_entree") or 0) * pe + float(b.get("jetons_sortie") or 0) * ps) / 1e6
    ch = _p_compte_chemin()
    if not ch or usd <= 0:
        return usd
    auj = _p_maintenant().date().isoformat()
    with _P_VERROU:
        c = _p_compte_lire(ch)
        c.setdefault("mois", {}).setdefault(auj[:7], {})
        c["mois"][auj[:7]][nom] = round(c["mois"][auj[:7]].get(nom, 0.0) + usd, 6)
        c.setdefault("jour", {})
        c["jour"][auj] = round(c["jour"].get(auj, 0.0) + usd, 6)
        _p_compte_ecrire(ch, c)
    return usd

@app.middleware("http")
async def _p_porte(request: Request, call_next):
    if request.method == "OPTIONS" or request.url.path not in _P_ROUTES:
        return await call_next(request)
    nom, refus = _p_juge(request)
    if refus is not None:
        return refus
    request.state.tailory_code = nom
    try:
        return await call_next(request)
    finally:
        with _P_VERROU:
            _P_EN_COURS.discard(nom)

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
# PDF : extraction PyMuPDF — raster + vectoriel + lignes fines
# Clustering v2.7 : composantes connexes (fini l'absorption par bbox)
# ─────────────────────────────────────────────
import fitz  # PyMuPDF, déjà installé (dépendance pdf2docx)


def _connected_components(rects, margin=10.0):
    """v2.7 — Regroupe les primitives par COMPOSANTES CONNEXES : arête entre deux primitives d'ORIGINE si leur écart
    < margin ; la bbox de chaque composante est calculée à la fin (et non par union itérative de boîtes, qui absorbait
    toute primitive survolée). (voir JOURNAL BACKEND v2.7)"""
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
    """v2.9 — Les rasters se regroupent ENTRE EUX, jamais avec les vecteurs. Trois règles :
    1. composantes connexes, marge 16 pt ;
    2. SCÈNES SEULEMENT (v2.9.1) : un composite exige ≥ 3 membres TOUS petits (min-dim ≤ 120 pt) ; tout composant
       contenant un grand raster, ou réduit à une paire, est dissous en rasters individuels ;
    3. SCISSION TABLEAU : un composite multi-rasters est re-scindé à chaque bordure horizontale de tableau qui le
       traverse sans toucher aucun de ses membres.
    (voir JOURNAL BACKEND v2.9)"""
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

    # v2.9.1 — filtre au niveau du composant : un composite n'est légitime que pour une SCÈNE (au moins 3 membres, tous
    # petits : min-dim ≤ 120 pt) ; tout composant contenant un grand raster, ou réduit à une paire, est dissous en rasters
    # individuels. (voir JOURNAL BACKEND v2.9)
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
    """v2.7 — RANGÉE de barres fines (h < 16 pt) alignées couvrant ensemble > 50 % de la largeur de page, sans voisin
    plein : un trait de section coupé par du texte, que le filtre pleine largeur (0.92) ne voit pas puisque chaque
    morceau est court."""
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
    """v2.7 — bande plate (h < 32 pt) contenant du texte (≥ 2 mots de ≥ 3 lettres, ou 1 seul mot couvrant > 30 % de sa
    largeur) : une cellule d'en-tête de tableau ou un titre décoré, PAS une figure. Les figures plates sans texte
    (tissus à lettres vectorielles, lignes graduées à chiffres) passent."""
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
    """Ligne horizontale = souligné si du texte immédiatement au-dessus couvre ≥ 60 % de sa longueur.
    v2.4 — la fenêtre tolère les jambages : le bas du mot entre 10 pt au-dessus et 6 pt au-dessous du trait.
    (voir JOURNAL BACKEND v2.4)"""
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

# V2.12 — DISTANCE LETTRE → TRAIT : l'étiquette peut être jusqu'à 60 pt (21 mm) du départ du trait (une tabulation
# avec de la marge). Deux conditions referment la porte aux soulignements : RIEN ENTRE LES DEUX (aucun mot ne
# s'intercale entre l'étiquette et le départ du trait, sur la même bande), et la condition de série du bloc 2c
# (un autre trait étiqueté de longueur différente sur la page). (voir JOURNAL BACKEND v2.12)
TOL_ETIQUETTE_PT = 60.0

_ETIQUETTE_SERIE = re.compile(r"^([A-Za-z]|\d{1,2})\s*[).]$")
_LETTRE_SEULE = re.compile(r"^([A-Za-z]|\d{1,2})$")


def _etiquette_de_serie(words, t):
    """V2.11 — vrai si un mot-étiquette de série (lettre ou chiffre seul suivi d'une parenthèse ou d'un point :
    a) b) e) 1) 2.) se tient JUSTE À GAUCHE du trait : la signature d'un segment d'exercice de mesure ; un tiret,
    une puce ou un soulignement n'est jamais étiqueté. À gauche seulement : une étiquette au-dessus ressemble
    trop à un mot souligné ou à un numérateur de fraction."""
    # V2.12 — une étiquette peut arriver en DEUX mots (« a » puis « ) ») : on recompose avant de juger.
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
        # à gauche : le mot finit avant le trait (chevauchement de 2 pt toléré), à moins de TOL_ETIQUETTE_PT de son départ
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


# ═══ V2.21 — LES CADRES, LES ZONES VIDES, ET LA FRONTIÈRE DU NATIF ═══
# Le natif ne remplace que ce que le code sait redessiner sans rien perdre : cases, grilles, bandes de nombres ; tout
# support figuratif (dessin, photo, carte, schéma) reste image. Rien ici ne touche les rasters. Trois signatures :
#   · cadre d'exercice = trait sans remplissage, ≤ 10 segments, lignes sur le pourtour et petites courbes DE COIN (chaque
#     courbe ≤ 15 % du petit côté) ; une ellipse a des courbes en quarts entiers : jamais prise ;
#   · CADRE ENGLOBANT : un cadre qui contient ≥ 3 autres primitives ou rasters — il ne lie plus et ne part plus en figure ;
#   · CADRE NU VIDE (rien dedans, ≤ 2 mots) : une ZONE, déclarée (zones_libres), jamais une image.
# (voir JOURNAL BACKEND v2.21 et v2.22)
def _v21_rect_au_trait(d):
    r = d.get("rect")
    if r is None:
        return False
    f = d.get("fill")
    # un fond BLANC n'empêche pas d'être un cadre (les cartes sont des rectangles pointillés à fond blanc)
    if f is not None and not all(v >= 0.95 for v in f):
        return False
    items = d.get("items", [])
    # jusqu'à 24 segments : un rectangle arrondi pointillé se rend en 7-8 lignes + 12 courbes de coin
    if not items or len(items) > 24:
        return False
    petit = max(1.0, min(r.width, r.height))
    for it in items:
        t = it[0]
        if t == "l":
            for p in (it[1], it[2]):
                if min(abs(p.x - r.x0), abs(p.x - r.x1)) > 3 and \
                   min(abs(p.y - r.y0), abs(p.y - r.y1)) > 3:
                    return False
        elif t == "c":
            bb = fitz.Rect(min(p.x for p in it[1:]), min(p.y for p in it[1:]),
                           max(p.x for p in it[1:]), max(p.y for p in it[1:]))
            # V2.24 — le coin se juge AUSSI en absolu : un rayon de coin est un trait d'auteur (~20-23 pt), il ne
            # grandit pas avec le cadre. (voir JOURNAL BACKEND v2.24)
            if max(bb.width, bb.height) > max(0.15 * petit, 25.0):
                return False
        elif t == "re":
            continue
        else:
            return False
    return True

def _v21_classe_grands_traces(page):
    """Renvoie (indices_englobants, zones_libres) sur les tracés de la page."""
    dessins = page.get_drawings()
    rects = [d.get("rect") for d in dessins]
    rasters = []
    for img in page.get_images(full=True):
        try:
            rasters.extend(fitz.Rect(rr) for rr in page.get_image_rects(img[0]))
        except Exception:
            pass
    englobants, zones = set(), []
    for i, d in enumerate(dessins):
        r = rects[i]
        if r is None or not _v21_rect_au_trait(d):
            continue
        contenus = sum(1 for j, o in enumerate(rects)
                       if j != i and o is not None and r.contains(o))
        contenus += sum(1 for o in rasters if r.contains(o))
        # le verrou des tissus : un contour de tissu est PLEIN là où les cadres d'exercice sont POINTILLÉS —
        # sans tirets, pas d'englobant. (voir JOURNAL BACKEND v2.21)
        tirets = bool(d.get("dashes")) and str(d.get("dashes")).strip() not in ("", "[] 0")
        if r.get_area() >= 8000 and contenus >= 3 and tirets:
            englobants.add(i)
        # V2.25 — une zone s'annonce en pointillés : un rectangle PLEIN et vide est une forme d'exercice, pas une zone.
        # (voir JOURNAL BACKEND v2.25)
        elif contenus == 0 and r.get_area() >= 2000 and tirets:
            import re as _re
            mots = len(_re.findall(r"[A-Za-zàâçéèêëîïôöûüù]{2,}",
                                   page.get_text(clip=r)))
            if mots <= 2:
                englobants.add(i)          # il ne part pas non plus en figure
                zones.append({"page": None, "w_mm": round(r.width * 25.4 / 72, 1),
                              "h_mm": round(r.height * 25.4 / 72, 1)})
    return englobants, zones

def _v21_est_receptacle(t):
    """« … + … = … » : un réceptacle de réponse, jamais une image (B7)."""
    import re as _re
    t2 = (t or "").strip()
    if not t2:
        return False
    reste = _re.sub(r"\bégal\b", "", t2, flags=_re.I)
    if _re.search(r"[A-Za-zàâçéèêëîïôöûüù]{3,}", reste):
        return False
    return t2.count("…") >= 2 or ("=" in t2 and "…" in t2) \
        or bool(_re.fullmatch(r"[\s.…_+=\-0-9]+", t2))


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
    v21_exclus, v21_zones = _v21_classe_grands_traces(page)   # V2.21
    solids, thins = [], []
    micros = []          # V2.26 — recueillis pour étendre, jamais pour fonder
    courts = []  # V2.11 — candidats 7-14 mm, promus au bloc 2c ou jetés
    rasters = []  # v2.8 — les rasters ne passent PAS par le clustering

    # 1. Images raster — v2.8 : chaque raster est une figure autonome.
    for img in page.get_images(full=True):
        try:
            for rect in page.get_image_rects(img[0]):
                # v2.33 — une image collée a été posée EXPRÈS : son plancher
                # est celui du résidu invisible (8 pt), jamais celui des traits.
                # (voir JOURNAL BACKEND v2.33)
                if rect.width > 8 and rect.height > 8:
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
        for _v21_i, d in enumerate(page.get_drawings()):
            if _v21_i in v21_exclus:      # V2.21 — cadre englobant ou zone vide
                continue
            r = d.get("rect")
            if r is None:
                continue
            if r.width > 0.92 * pw and r.height < 20:
                continue  # ligne de séparation pleine largeur
            if r.width * r.height > 0.85 * page_area:
                continue  # FOND DE PAGE (v2.3) — absorbait tout en v2.2
            if r.width < 10 and r.height < 10:
                # V2.26 — un micro-tracé ne FONDE jamais une figure, mais il ÉTEND le cadre de celle qu'il touche.
                # (voir JOURNAL BACKEND v2.26)
                micros.append(fitz.Rect(r))
                continue  # micro-tracé
            if r.width < 10 or r.height < 10:
                # Ligne fine : candidate d'office si assez longue pour être une figure (segment, ligne graduée), pas un
                # simple tiret. V2.11 — entre 7 et 14 mm, candidate SOUS CONDITIONS, vérifiées au bloc 2c.
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

    # 2c. V2.11 — promotion des segments courts étiquetés. Un trait de 7 à 14 mm n'entre qu'à deux conditions cumulées :
    #   1. une étiquette de série juste à sa gauche (a), e), 1), 2. …) ;
    #   2. un AUTRE trait étiqueté de longueur différente (> 5 %) sur la page (les segments d'un exercice de mesure ont
    #      des longueurs toutes différentes ; les lignes de réponse d'une liste numérotée font toutes la même : dehors).
    # Les promus subissent ensuite les mêmes filtres anti-bruit que les longs (isolement, souligné, familles) puis le
    # filtre v2.5 des pages de mesure. (voir JOURNAL BACKEND v2.11)
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

    return solids, thins, rasters, v21_zones, micros


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


# ══ v2.14 — LE GRAS DE LA SOURCE SURVIT À LA LECTURE DU PDF ═══
# La fonction relit la page span par span et entoure les passages en gras de deux marqueurs — ⟦gras⟧ … ⟦/gras⟧ — que la
# page sait reconvertir. Le gras : le bit 4 des drapeaux de span (16), ou « Bold » dans le nom de la police.
# Borne connue : le chemin « grilles » (texte_avec_grilles) construit son texte lui-même et n'est pas couvert.
# (voir JOURNAL BACKEND v2.14)
_GRAS_OUV = "\u27e6gras\u27e7"
_GRAS_FER = "\u27e6/gras\u27e7"


def _span_est_gras(span):
    if (span.get("flags", 0) & 16):
        return True
    nom = (span.get("font", "") or "").lower()
    return ("bold" in nom) or ("black" in nom) or ("heavy" in nom)


# ═══ v2.36 — R1 : L'ANCRE DE LECTURE ═══
# CONDITION : une page dont les blocs de texte sont rendus avec leur position.
# EFFET : chaque figure de la page reçoit, dans le texte transmis au modèle, le marqueur [TAILORY_IMG_N] posé APRÈS le
# bloc de texte qui la précède dans l'ordre de lecture. Le texte n'est pas reconstruit : les marqueurs y sont INSÉRÉS.
# (voir JOURNAL BACKEND v2.36)
# v2.64 — sauf une figure d'exercice dont la consigne est trouvée sur la page : son marqueur est inséré juste après la
# dernière ligne de sa consigne, dans le bloc (voir JOURNAL BACKEND v2.64)
BANDE_ENJAMBEMENT = 8.0


def bloc_enjambe(bloc, y_figure, bande=BANDE_ENJAMBEMENT):
    """v2.37 — LA FIGURE EST-ELLE ENJAMBÉE PAR SON BLOC D'ANCRAGE ?
    Un bloc de TEXTE à plusieurs lignes qui commence au-dessus de la figure et se termine en dessous ne l'annonce
    plus : cette figure ne reçoit pas de marqueur, elle est comptée à part et déclarée. Trois conditions, toutes
    nécessaires : le bloc est LU SUR LA PAGE (pas fabriqué par le serveur) ; il porte PLUSIEURS LIGNES ; son bas
    dépasse le haut de la figure de plus d'une bande de lecture. (voir JOURNAL BACKEND v2.37)
    """
    if len(bloc) < 5 or bloc[4] is not False:
        return False                      # fabriqué, ou origine inconnue
    if bloc[3] is None:
        return False
    if (bloc[2] or "").count("\n") < 1:
        return False                      # une seule ligne
    return bloc[3] > y_figure + bande   # v2.37 — une bande de lecture ; au-delà, le bloc
                          # enjambe vraiment la figure et ne l'annonce plus.


def _v239_distance_sous(page, figures, _tol=0.5):
    """Pose `ecart_mm` sur chaque figure qui a quelque chose en dessous d'elle (v2.39). Rend le nombre de figures
    servies. NE POSE RIEN quand rien n'est trouvé : l'absence du champ est le silence, et la page n'ajoute rien alors.
    « En dessous » : la première chose sous la figure, quelle qu'elle soit (bloc de texte ou autre figure), avec un
    recouvrement horizontal exigé — sinon la colonne voisine serait prise pour le dessous.
    (voir JOURNAL BACKEND v2.39)
    """
    if not figures:
        return 0
    objets = []
    try:
        for b in page.get_text("blocks"):
            if len(b) >= 5 and str(b[4]).strip():
                objets.append((b[1], b[0], b[2]))          # (haut, gauche, droite)
    except Exception:
        return 0                                            # on s'abstient, on ne devine pas
    for f in figures:
        c = f.get("cadre")
        if c:
            objets.append((c[1], c[0], c[2]))
    n = 0
    for f in figures:
        c = f.get("cadre")
        if not c:
            continue
        x0, y0, x1, y1 = c[0], c[1], c[2], c[3]
        haut_du_dessous = None
        for (oy0, ox0, ox1) in objets:
            if oy0 < y1 - _tol:                             # pas en dessous
                continue
            if min(x1, ox1) - max(x0, ox0) <= 0:            # pas sous elle
                continue
            if haut_du_dessous is None or oy0 < haut_du_dessous:
                haut_du_dessous = oy0
        if haut_du_dessous is not None:
            f["ecart_mm"] = round((haut_du_dessous - y1) * 25.4 / 72, 1)
            n += 1
    return n


def poser_ancres(texte, blocs, figures, cle=None, apres=None):
    """Rend (texte marqué, nombre posé, nombre non posé).

    Une figure dont l'ancre ne peut pas être calculée n'est PAS posée au hasard :
    elle est comptée à part et son marqueur n'est pas écrit. Un contrôle qui ne
    peut pas s'exécuter ressemble à un contrôle qui passe (famille A8).

    v2.64 — `apres` : {index : position dans le texte}. Une figure qui y figure reçoit son marqueur à cette position,
    juste après la dernière ligne de sa consigne ; toute autre, à la fin de son bloc (voir JOURNAL BACKEND v2.64)."""
    if not figures:
        return texte, 0, 0
    if not blocs:
        return texte, 0, len(figures)

    # Fins de bloc dans le texte. Le séparateur du join n'est PAS le même selon la fabrique qui a produit le texte :
    # on ne suppose rien, on RETROUVE chaque bloc dans le texte, de proche en proche. (voir JOURNAL BACKEND v2.36)
    fins, curseur = [], 0
    for b in blocs:
        t = b[2]
        pos = texte.find(t, curseur)
        if pos < 0:
            return texte, 0, len(figures)  # le texte n'est pas celui des blocs
        curseur = pos + len(t)
        fins.append(curseur)

    # L'ancre : le DERNIER bloc qui PRÉCÈDE la figure DANS L'ORDRE DU TEXTE.
    # v2.37 — elle emploie LA MÊME CLÉ que l'assemblage (`cle_lecture` du module des grilles, une seule au projet) et
    # parcourt les blocs DANS L'ORDRE OÙ ILS SONT ASSEMBLÉS. (voir JOURNAL BACKEND v2.37)
    if cle is None:
        return texte, 0, len(figures)     # sans la clé du projet : abstention
    poses = []
    for f in figures:
        cf = cle(f["cadre"][1], f["cadre"][0])
        k = None
        for i, b in enumerate(blocs):
            if cle(b[0], b[1]) < cf:
                k = i
        # v2.37 — une figure ENJAMBÉE par son bloc d'ancrage ne reçoit PAS de marqueur : elle est comptée à part et
        # déclarée. Seuil : le bas du bloc dépasse le haut de la figure de plus d'une bande de lecture. (voir JOURNAL BACKEND v2.37)
        poses.append((k, f["index"], (apres or {}).get(f["index"])))

    # On insère de la fin vers le début pour que les positions restent justes.
    # v2.64 — rangées par leur place dans le texte : la position après la consigne quand elle est donnée, sinon la fin du
    # bloc ; sans position donnée, l'ordre est celui d'avant (les fins de bloc ne décroissent pas). (voir JOURNAL BACKEND v2.64)
    poses.sort(key=lambda p: ((p[2] if p[2] is not None else (fins[p[0]] if p[0] is not None else -1)),
                              (p[0] if p[0] is not None else -1), p[1]))
    n_pose = 0
    for k, idx, apres_consigne in reversed(poses):
        marque = "[TAILORY_IMG_%d]" % idx
        if apres_consigne is not None:
            texte = texte[:apres_consigne] + "\n" + marque + texte[apres_consigne:]   # v2.64 — juste après la consigne
        elif k is None:
            texte = marque + "\n" + texte          # au-dessus de tout texte
        elif fins[k] < len(texte) and texte[fins[k]] == "\n":
            texte = texte[:fins[k] + 1] + marque + "\n" + texte[fins[k] + 1:]
        elif fins[k] > 0 and texte[fins[k] - 1] == "\n":
            texte = texte[:fins[k]] + marque + "\n" + texte[fins[k]:]
        elif fins[k] == len(texte):
            texte = texte + "\n" + marque          # dernier bloc, pas de saut
        else:
            continue                                # position douteuse : rien
        n_pose += 1
    return texte, n_pose, len(figures) - n_pose


def _texte_gras(page, rendre_blocs=False):
    """Texte de la page, gras marqué. Retombe sur get_text('text') au moindre doute.

    v2.36 — rendre_blocs=True rend AUSSI la liste des blocs dans l'ordre où ils
    ont été assemblés, chacun avec sa position : (y0, x0, texte). Elle sert à
    poser les ancres de lecture. Le texte rendu est le MÊME, caractère pour
    caractère ; rien n'est reconstruit. (voir JOURNAL BACKEND v2.36)"""
    try:
        d = page.get_text("dict")
    except Exception:
        t = page.get_text("text")
        return (t, None) if rendre_blocs else t
    blocs = []
    pos_y, pos_x, pos_y1 = [], [], []   # v2.36/37 — haut, gauche ET bas
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
            _bb = b.get("bbox") or (0, 0, 0, 0)
            pos_x.append(_bb[0]); pos_y.append(_bb[1]); pos_y1.append(_bb[3])
    if not blocs:
        t = page.get_text("text")
        return (t, None) if rendre_blocs else t
    texte = "\n".join(blocs) + "\n"
    if rendre_blocs:
        return texte, [(y, x, t, y1, False)
                       for y, x, t, y1 in zip(pos_y, pos_x, blocs, pos_y1)]
    return texte


# ─── v2.52 — LE JUGEMENT CORRIGÉ : SA RÈGLE, LUE PAR SONNET 5.5 ───
# Le modèle recopie les lignes de la page ; le code juge. (voir JOURNAL BACKEND v2.52, v2.57)
_V252_MODELE = "claude-sonnet-5-5"    # le lecteur
_V252_PPP = 150                        # la définition de la page rendue pour le lecteur
_V252_MAX_SORTIE = 2000                # la borne de la recopie, en jetons
_V252_PARALLELE = 4                    # appels en parallèle
_V252_MOTS_TITRE = 2                   # T1 : le mot seul, au plus 2 mots
_V252_CONSIGNE = "Recopie chaque ligne de texte visible sur cette image de page, une ligne de texte par ligne, exactement telle qu'elle est écrite (y compris les titres, les cachets, les mots dessinés ou manuscrits et les mots en cursive), dans l'ordre de lecture. Ne commente pas, ne corrige pas, n'ajoute rien, ne résume pas. Si une ligne est illisible, écris [illisible]."   # la consigne du compte 191, à l'octet (batterie_252_juge.py la vérifie)
_V252_PONCT = "\"'«»“”‘’()[]{}:;,.!?…-–—_*/\\|<>•·"


def _v252_sans_accents(s):
    import unicodedata as _ud
    return "".join(c for c in _ud.normalize("NFD", s) if _ud.category(c) != "Mn")


def _v252_jeton(t):
    import unicodedata as _ud
    t = _ud.normalize("NFC", t).strip(_V252_PONCT)
    return t if any(ch.isalnum() for ch in t) else None


def _v252_est_le_mot(t):
    """Le mot large : « corrigé / corrigés », l'accent et la casse ignorés. Une déclaration : un seul mot, ses deux
    nombres — pas une liste de formes."""
    m = _v252_jeton(t)
    return bool(m) and _v252_sans_accents(m).lower() in ("corrige", "corriges")


def _v252_ligne_en_titre(ligne):
    """T1 : la ligne porte le mot et n'a pas plus de 2 mots (la ponctuation seule ne compte pas)."""
    jetons = [j for j in (_v252_jeton(t) for t in ligne.split()) if j is not None]
    return len(jetons) <= _V252_MOTS_TITRE and any(_v252_est_le_mot(t) for t in ligne.split())


def _v252_page_est_corrige(lignes):
    return any(_v252_ligne_en_titre(l) for l in lignes)


def _v252_compte_vide():
    """La marque machine du lecteur : rendue dans la réponse, jamais silencieuse, jamais sur le papier (la page ne lit pas ce champ)."""
    return {"lecteur": _V252_MODELE, "ppp": _V252_PPP, "pages": 0, "pages_lues": 0, "pages_corrige": [], "pages_en_panne": [],
            "pages_coupees": [], "sans_cle": False, "jetons_entree": 0, "jetons_sortie": 0, "derniere_erreur": ""}


def _v252_lire_page(api_key, png):
    """UNE page recopiée par le lecteur. Rend (lignes | None, faute | None, coupée, jetons_in, jetons_out).
    La reprise est celle de la bibliothèque (ses essais, ses attentes), comme pour la gomme et la fiche ; une panne
    est rendue, jamais levée (v2.53)."""
    try:
        client = anthropic.Anthropic(api_key=api_key, timeout=120.0)
        rep = client.messages.create(
            model=_V252_MODELE, max_tokens=_V252_MAX_SORTIE,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                             "data": base64.b64encode(png).decode()}},
                {"type": "text", "text": _V252_CONSIGNE}]}])
    except Exception as e:
        return None, "panne d'appel : %s" % type(e).__name__, False, 0, 0
    texte = "".join(b.text for b in rep.content if getattr(b, "type", "") == "text")
    lignes = [l.strip() for l in texte.split("\n") if l.strip()]
    u = rep.usage
    return (lignes, None, rep.stop_reason == "max_tokens",
            getattr(u, "input_tokens", 0) or 0, getattr(u, "output_tokens", 0) or 0)
    return None, faute, False, 0, 0


def _v252_pages_de_corrige(doc, api_key, compte):
    """L'ensemble (1-indexé) des pages jugées corrigé par sa règle sur la recopie de chaque page.
    Sans clé : aucune lecture, aucune page jugée, et c'est dit. Une page en panne reste dans la fiche, et c'est dit.
    Le rendu des pages se fait avant les appels (le document ne se partage pas entre fils) ; les appels, à 4 en parallèle."""
    compte["pages"] = len(doc)
    if not api_key:
        compte["sans_cle"] = True
        return set()
    pngs = []
    for pno, page in enumerate(doc):
        try:
            pngs.append((pno + 1, page.get_pixmap(matrix=fitz.Matrix(_V252_PPP / 72.0, _V252_PPP / 72.0),
                                                  alpha=False).tobytes("png")))
        except Exception as e:
            compte["pages_en_panne"].append([pno + 1, "rendu impossible : %s" % type(e).__name__])
    from concurrent.futures import ThreadPoolExecutor as _TPE252
    flag = set()
    with _TPE252(max_workers=_V252_PARALLELE) as ex:
        lectures = list(ex.map(lambda t: _v252_lire_page(api_key, t[1]), pngs))
    for (pno, _png), (lignes, faute, coupee, j_in, j_out) in zip(pngs, lectures):
        compte["jetons_entree"] += j_in
        compte["jetons_sortie"] += j_out
        if lignes is None:
            compte["pages_en_panne"].append([pno, faute])
            compte["derniere_erreur"] = faute or ""
            continue
        compte["pages_lues"] += 1
        if coupee:
            compte["pages_coupees"].append(pno)
        if _v252_page_est_corrige(lignes):
            flag.add(pno)
    compte["pages_corrige"] = sorted(flag)
    return flag


# ═══ v2.38, v2.54 — LA FEUILLE ET LE MODÈLE NE REÇOIVENT PAS LA MÊME IMAGE ═══
# `data` est ce que le modèle reçoit (144 ppp) ; `data_feuille`, la même zone rendue à 300 ppp, est ce que la feuille reçoit.
# (voir JOURNAL BACKEND v2.38, v2.54)

_V238_PPP_FEUILLE = 300.0    # la définition du rendu de feuille


def _v238_champs(page, clip):
    """L'image de feuille d'une figure : la zone rendue à 300 ppp. Rend (dict, 'feuille'|'aucun').
    ⛔ NE TOUCHE JAMAIS `data`. Si le rendu échoue : aucun champ, la figure garde `data`."""
    import base64 as _b64238
    try:
        m = _V238_PPP_FEUILLE / 72.0
        pf = _lis2_page(page).get_pixmap(clip=clip, matrix=fitz.Matrix(m, m))
        return ({"data_feuille": "data:image/png;base64,"
                                 + _b64238.b64encode(pf.tobytes("png")).decode()},
                "feuille")
    except Exception:
        return ({}, "aucun")


# ══════════════════════════════════════════════════════════════════════════════
# LIS2 — LE LISERÉ RÉPARÉ SUR LES SEULES PHOTOGRAPHIES DE LA FEUILLE (ligne 2.11 du plan ; DEC-2223 : « La réparation du liseré lit juste
# toutes les images à palette, comme elle lit les autres : une correction générale (comme F321), pas une exclusion. » ; F334). Ce que ce
# bloc fait, et rien d'autre :
#   · beaucoup d'images détourées (poissons, gâteaux) sont rangées dans le PDF avec leurs couleurs mélangées au noir (le « Matte ») :
#     rendues telles quelles, elles ont un contour sombre, le liseré ;
#   · la découpe lit le PDF d'origine, et l'image que voit le modèle (`data`, 144 ppp) reste celle d'origine ; seule la photographie de
#     la feuille (`data_feuille`, 300 ppp, rendue par `_v238_champs`) est rendue sur une COPIE du PDF, propre à ce fil, où chaque image à
#     masque retrouve ses vraies couleurs : le mélange est défait, la couleur du pixel visible le plus proche est étendue sous les pixels
#     tout à fait transparents, le Matte est retiré ; le masque ne change pas ;
#   · seules les images en couleurs ou en gris à 8 bits sont réparées ; les autres sont comptées et laissées ;
#   · gris ou couleur se lit sur l'image DÉCODÉE : une image à palette (Indexed) n'a qu'une composante d'index, mais le moteur la rend en
#     couleur (F334 : sinon elle passait en gris) ; un Matte plus court que les couleurs décodées n'est pas touché.
# Recopié du banc sans changement de sens (outils/fabrique_variantes.py : LIS2_NEUF, FEUILLE_NEUF, _banc_photo_prepare, _banc_page_photo,
# _banc_lisere_repare ; comptes 212, 213, 215). QUATRE ÉCARTS DÉCLARÉS : scipy, absent du serveur, est remplacé par numpy (la distance
# exacte au pixel visible le plus proche, mêmes indices que scipy.ndimage.distance_transform_edt, égalité prouvée) ; la copie réparée est
# propre au fil (comme l'état `_V254`) et non plus un dictionnaire commun ; les noms `_banc_*` deviennent `_lis2_*` ; les fonctions sont
# placées ici. Ce que la gomme (2.70) fait de cette réparation : `_v270_photographie_apres` refait les pixels de `data_feuille` que
# l'effacement change d'après le PDF d'origine et le PDF effacé, tous deux non réparés.
# (voir JOURNAL BACKEND, LIS2)
# ══════════════════════════════════════════════════════════════════════════════
import threading as _thl2
_LIS2 = _thl2.local()

def _lis2_prepare(content):
    """La copie réparée du PDF, propre à ce fil : elle ne sert qu'aux rendus à 300 ppp de la feuille (`_v238_champs`) ; la découpe lit l'original."""
    d = fitz.open(stream=content, filetype="pdf"); _lis2_repare(d); _LIS2.doc = d

def _lis2_page(page):
    d = getattr(_LIS2, "doc", None)
    try:
        return d[page.number] if d is not None and page.parent is not None and len(d) == len(page.parent) else page
    except Exception:
        return page

def _lis2_repare(doc):
    """Répare, dans la copie `doc` du PDF, chaque image à masque. Masque avec « Matte » (couleurs stockées pré-mélangées au noir) :
    dé-mélange c = m + (c' - m) x 255 / a pour a > 0, puis le Matte est retiré ; sous les pixels tout à fait transparents (a = 0), la
    couleur du pixel visible le plus proche. Le masque ne change pas. Seuls le RGB et le gris 8 bits sont touchés ; le reste est compté,
    pas touché. Rend (réparées, sautées)."""
    from PIL import Image as _Il
    faits = sautes = 0
    for x in range(1, doc.xref_length()):
        try:
            if doc.xref_get_key(x, "Subtype")[1] != "/Image":
                continue
            sm = doc.xref_get_key(x, "SMask")
            if sm[0] != "xref":
                continue
            smx = int(sm[1].split()[0]); _e = doc.extract_image(x)
            # (réglé sur son témoin à 17 h 05, avant le corpus : un espace ICC — une référence — était sauté ; on lit le nombre de composantes)
            if _e.get("colorspace") not in (1, 3) or doc.xref_get_key(x, "BitsPerComponent")[1] != "8" or doc.xref_get_key(x, "Decode")[0] != "null":
                sautes += 1; continue
            # § 215 (DEC-2223, écrit à 20 h 24, avant la mesure) — gris ou couleur se lit sur l'image DÉCODÉE : une image à palette
            # (Indexed) n'a qu'une composante d'index, mais le moteur la rend en couleur
            _brut = _Il.open(io.BytesIO(_e["image"])); _gris = _brut.mode in ("L", "LA", "1")
            base = _brut.convert("L" if _gris else "RGB")
            a = _np247.asarray(_Il.open(io.BytesIO(doc.extract_image(smx)["image"])).convert("L").resize(base.size)).astype(_np247.float32)
            c = _np247.asarray(base).astype(_np247.float32)
            if c.ndim == 2:
                c = c[:, :, None]
            matte = doc.xref_get_key(smx, "Matte")
            if matte[0] == "array":
                mv = _np247.array([float(v) * 255.0 for v in matte[1].strip("[]").split()][:c.shape[2]], _np247.float32)
                # § 215 — un Matte plus court que les couleurs décodées (palette à Matte d'index) n'est pas touché : compté
                if len(mv) < c.shape[2]:
                    sautes += 1; continue
                al = a[:, :, None]
                c = _np247.where(al > 0, mv + (c - mv) * 255.0 / _np247.maximum(al, 1.0), c)
                c = _np247.clip(c, 0, 255)
            zero = a == 0
            if zero.any() and (~zero).any():
                iy, ix = _lis2_plus_proche(zero)
                c = c[iy, ix]
            out = _np247.rint(c).astype(_np247.uint8)
            if out.shape[2] == 1:
                out = out[:, :, 0]
            doc.update_stream(x, out.tobytes())
            doc.xref_set_key(x, "Filter", "/FlateDecode"); doc.xref_set_key(x, "DecodeParms", "null")
            if _gris:
                doc.xref_set_key(x, "ColorSpace", "/DeviceGray")
            elif doc.xref_get_key(x, "ColorSpace")[0] != "name":
                doc.xref_set_key(x, "ColorSpace", "/DeviceRGB")   # l'octet écrit est du RGB 8 bits, décodé par PIL
            if matte[0] == "array":
                doc.xref_set_key(smx, "Matte", "null")
            faits += 1
        except Exception:
            sautes += 1
    return faits, sautes

def _lis2_plus_proche(zero):
    """(iy, ix) : pour chaque pixel, la ligne et la colonne du site (pixel où `zero` est faux) le plus proche, comme
    scipy.ndimage.distance_transform_edt(zero, return_indices=True). Suppose qu'il y a au moins un site."""
    H, W = zero.shape
    i32 = _np247.int32
    ys = _np247.arange(H, dtype=i32)
    gros = i32(4 * (H + W) + 8)
    site = ~zero
    # passe 1 : par colonne, le site le plus proche ; à égalité, le plus haut (le plus petit indice de ligne)
    haut = _np247.maximum.accumulate(_np247.where(site, ys[:, None], i32(-1)), axis=0)
    bas = _np247.minimum.accumulate(_np247.where(site, ys[:, None], gros)[::-1], axis=0)[::-1]
    d_haut = _np247.where(haut >= 0, ys[:, None] - haut, gros)
    d_bas = _np247.where(bas < gros, bas - ys[:, None], gros)
    F = _np247.where(d_haut <= d_bas, haut, bas)
    F = _np247.where((haut < 0) & (bas >= gros), i32(-1), F).astype(i32)
    del haut, bas, d_haut, d_bas
    colonne = (F >= 0).any(axis=0)
    # distance verticale au carré de chaque colonne candidate (valable là où la colonne a un site)
    D2 = ((_np247.where(F >= 0, F, i32(0)) - ys[:, None]).astype(_np247.int64) ** 2).astype(i32)
    # passe 2 : par ligne, enveloppe inférieure des paraboles (pile g, hauteur l), toutes les lignes ensemble
    g = _np247.zeros((H, W), dtype=i32)
    l = _np247.full(H, -1, dtype=_np247.int64)
    lignes = _np247.arange(H)
    for ii in range(W):
        if not colonne[ii]:
            continue
        wR = D2[:, ii].astype(_np247.int64)
        act = l >= 1
        while act.any():
            r = lignes[act]
            lr = l[r]
            idx1 = g[r, lr].astype(_np247.int64)
            idx2 = g[r, lr - 1].astype(_np247.int64)
            a = idx1 - idx2
            b = ii - idx1
            c = a + b
            dv = D2[r, idx1].astype(_np247.int64)
            du = D2[r, idx2].astype(_np247.int64)
            garde = (c * dv - b * du - a * wR[r] - a * b * c) <= 0
            retire = r[~garde]
            l[retire] -= 1
            act = _np247.zeros(H, dtype=bool)
            act[retire] = True
            act &= (l >= 1)
        l += 1
        g[lignes, l] = ii
    maxl = l
    # marche sur la pile : pour chaque colonne ii, la première parabole de la pile qui n'est pas dépassée par la suivante
    ptr = _np247.zeros(H, dtype=_np247.int64)
    IX = _np247.zeros((H, W), dtype=i32)
    for ii in range(W):
        cur = g[lignes, ptr].astype(_np247.int64)
        delta1 = D2[lignes, cur].astype(_np247.int64) + (cur - ii) ** 2
        act = ptr < maxl
        while act.any():
            r = lignes[act]
            nxt = g[r, ptr[r] + 1].astype(_np247.int64)
            delta2 = D2[r, nxt].astype(_np247.int64) + (nxt - ii) ** 2
            avance = delta1[r] > delta2
            ra = r[avance]
            ptr[ra] += 1
            delta1[ra] = delta2[avance]
            act = _np247.zeros(H, dtype=bool)
            act[ra] = True
            act &= (ptr < maxl)
        IX[:, ii] = g[lignes, ptr]
    IY = F[lignes[:, None], IX]
    return IY, IX


# ══════════════════════════════════════════════════════════════════════════════
# v2.47 — LA DÉCOUPE EN MORCEAUX. Ce que ce bloc fait, et rien d'autre :
#   · on coupe sans plancher : un rectangle ne prend jamais autre chose que sa figure ; sinon plusieurs rectangles qui
#     épousent le bord de la figure ;
#   · les morceaux partent avec leur position (`cadre`) et leur marque (`figure`, `morceau`, `n_morceaux`) ; la page les
#     remet à leur place ;
#   · l'intérieur d'une figure n'est jamais effacé (une fenêtre qui en contient une autre n'est pas « une autre figure ») ;
#   · les lettres et les chiffres à 6 mm ou moins du dessin partent avec la figure, à leur place ; un mot part avec une
#     seule figure ; jamais pris : une ligne de texte, un mot portant une ponctuation de fin (. ? ! : ; …) ;
#   · un tableau, une grille, une rangée de bandeaux est une figure : ce qui est dedans lui appartient ;
#   · le dessin, c'est la flèche qui touche la figure, la bande de cases couchée ou debout, et le pâle (tout pixel qui
#     s'écarte du blanc) ;
#   · pour toute figure de plus d'un morceau, `entier` (le PNG du rectangle qui contient tous ses morceaux).
# L'étiquetage est un union-find par segments (sans scipy) ; la distance d'un mot au dessin est exacte (boîte → pixel
# d'encre le plus proche). (voir JOURNAL BACKEND v2.47)
# ══════════════════════════════════════════════════════════════════════════════
import numpy as _np247
from PIL import Image as _Im247, ImageFilter as _If247
from PIL import ImageDraw as _Id249   # v2.49 : le tracé des contours de décor

_V247_K = 2.0                       # 144 ppp, comme la photographie du modèle
_V247_MM = 25.4 / 72.0
# v2.50 — l'écart au blanc qui fait l'encre : 25 (voir JOURNAL BACKEND v2.50)
_V247_ECART = 25                    # encre = écart au blanc ≥ 25 (le pâle compte)
_V247_BRUIT = 9                     # px² : poussières écartées
_V247_MARGE = 20.0 / _V247_MM       # pt : la région autour d'une fenêtre (flèches, mots à 6 mm)
_V247_DIST_MM = 6.0                 # les lettres et les chiffres à 6 mm ou moins
_V247_PONCT = '.?!:;…'
_V247_MOT = re.compile(r'[A-Za-zÀ-ÿ]{3,}')


def _v247_lignes(page):
    """Les mots de la page, et pour chacun la taille de sa ligne et le texte de sa ligne
    (même hauteur à 50 %, écart ≤ 1,5 × hauteur)."""
    mots = [w for w in page.get_text('words') if w[4].strip()]
    # v2.49 — un mot écrit plusieurs fois au même endroit (faux gras) est UN mot : un seul exemplaire par texte quand les boîtes se recouvrent à 80 %, avant de former les lignes
    _uniques = []
    for w in mots:
        _b = fitz.Rect(w[:4]); _double = False
        for u in _uniques:
            if u[4] != w[4]:
                continue
            _q = fitz.Rect(u[:4]) & _b
            if not _q.is_empty and _q.get_area() >= 0.8 * min(_b.get_area(), fitz.Rect(u[:4]).get_area()):
                _double = True
                break
        if not _double:
            _uniques.append(w)
    mots = _uniques
    par = list(range(len(mots)))
    _ga_rot = [fitz.Rect(l["bbox"]) for b in page.get_text("dict")["blocks"] if b.get("type") == 0 for l in b.get("lines", []) if abs((l.get("dir") or (1, 0))[1]) > 0.5]   # v2.72 G-a (DEC-3312) : un mot d'une ligne tournée ne se groupe qu'avec les mots de même sens
    _tv = [any(r.contains(fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2)) for r in _ga_rot) for w in mots]   # v2.72 G-a (DEC-3312) : un mot d'une ligne tournée ne se groupe qu'avec les mots de même sens

    def tr(a):
        while par[a] != a:
            par[a] = par[par[a]]
            a = par[a]
        return a
    for i in range(len(mots)):
        xi0, yi0, xi1, yi1 = mots[i][:4]
        hi = yi1 - yi0
        for j in range(i + 1, len(mots)):
            xj0, yj0, xj1, yj1 = mots[j][:4]
            hj = yj1 - yj0
            if (_tv[i] == _tv[j] and min(yi1, yj1) - max(yi0, yj0) >= 0.5 * min(hi, hj)
                    and max(xj0 - xi1, xi0 - xj1) <= 1.5 * max(hi, hj)):
                par[tr(i)] = tr(j)
    grp = {}
    for i in range(len(mots)):
        grp.setdefault(tr(i), []).append(i)
    L, texte = {}, {}
    for g, idx in grp.items():
        t = ' '.join(mots[j][4] for j in sorted(idx, key=lambda j: mots[j][0]))
        for i in idx:
            L[i] = len(idx)
            texte[i] = t
    # v2.60 — un mot d'adresse (http, www, .fr, .com, .org, .net) compte pour une ligne : son groupe vaut 3 au moins
    for i, w in enumerate(mots):
        if _V260_ADR.search(w[4] or ""):
            L[i] = max(L[i], 3)
    return mots, L, texte


def _v247_est_ligne(i, L, texte):
    return L[i] >= 3 and bool(_V247_MOT.search(texte[i]))


_C_SIGNE = re.compile(r'[+\-−–×÷=<>]')   # v2.72 C
_C_NOMBRE = re.compile(r'[\d.,]+')   # v2.72 C
_C_CALCUL = re.compile(r'(?:^|\s)[+\-−–×÷=<>](?:\s|$)|\d\s*[x×+\-−–÷=<>]\s*\d')   # v2.72 C
def _v247_est_glyphe(i, mots, L, texte):
    """Un mot court, sans ponctuation de fin : candidat à partir avec une figure."""
    t = mots[i][4]
    if _v247_est_ligne(i, L, texte):
        return False
    # v2.72 C (DEC-2961, DEC-3208, DEC-3312) : un signe de calcul, ou un nombre sur une ligne de calcul, est du texte, jamais une étiquette
    _ct = (t or '').strip()
    if _C_SIGNE.fullmatch(_ct) or (_ct in ('x', '×') and _C_CALCUL.search(texte[i] or '')) or (_C_NOMBRE.fullmatch(_ct) and _C_CALCUL.search(texte[i] or '')):
        return False
    # v2.72 E (DEC-2977, DEC-3312) : un mot de trois lettres et plus n'est jamais une étiquette de dessin
    if _V247_MOT.search(t or ''):
        return False
    return not (t and t[-1] in _V247_PONCT)


def _v247_cases_et_decor(page):
    cases, decor = [], []
    for d in page.get_drawings():
        r = d.get('rect')
        if r is None or r.width < 8 or r.height < 8:
            continue
        tirets = bool(d.get('dashes')) and str(d.get('dashes')).strip() not in ('', '[] 0')
        if tirets:
            # v2.49 — le décor garde son chemin (segments, rectangles, courbes) : l'anneau effacé suivra le contour, pas la boîte
            decor.append((fitz.Rect(r), list(d.get('items') or [])))
            continue
        if r.width * r.height > 40000:
            continue
        t = (page.get_text(clip=r) or '').strip()
        if (not t) or _v21_est_receptacle(t):
            # v2.48 — une case où l'élève écrit est un rectangle : traits et rectangles seulement ; un œil est rond, ce n'est pas une case
            rectangle = all(it[0] in ('l', 're') for it in d.get('items', []))
            if rectangle and d.get('fill') in (None, (1.0, 1.0, 1.0)):
                cases.append(fitz.Rect(r))
    # v2.59 — un trait fin (moins de 8 pt dans un sens) est du décor s'il est pointillé, fait de segments droits, et ne touche ni un
    # autre tracé ni une image (voir JOURNAL BACKEND v2.59)
    vus = {tuple(round(v, 2) for v in rd) for rd, it in decor}
    for d in page.get_drawings():
        r = d.get('rect')
        if r is not None and (r.width < 8 or r.height < 8) and _v259_est_decor(page, d) and tuple(round(v, 2) for v in r) not in vus:
            decor.append((fitz.Rect(r), list(d.get('items') or [])))
    return cases, decor


def _v247_en_bande(c, cases):
    """La bande de cases, couchée ou debout : une case n'est un intrus que seule sur sa ligne."""
    petit = min(c.width, c.height)
    for c2 in cases:
        if c2 is c:
            continue
        if (abs(c2.height - c.height) <= 0.15 * c.height
                and min(c.y1, c2.y1) - max(c.y0, c2.y0) >= 0.5 * min(c.height, c2.height)
                and max(c2.x0 - c.x1, c.x0 - c2.x1) <= 1.5 * petit):
            return True
        if (abs(c2.width - c.width) <= 0.15 * c.width
                and min(c.x1, c2.x1) - max(c.x0, c2.x0) >= 0.5 * min(c.width, c2.width)
                and max(c2.y0 - c.y1, c.y0 - c2.y1) <= 1.5 * petit):
            return True
    return False


def _v247_tableaux(page, cases):
    """Un tableau, une grille, une rangée de bandeaux : (a) find_tables s'il existe ;
    (b) les grilles de traits fins (bordures Word : rectangles remplis d'épaisseur nulle) ;
    (c) les rangées de cellules fermées. Rend (liste de Rect, tableaux_non_cherches)."""
    tabs, non_cherches = [], 0
    if hasattr(page, 'find_tables'):
        try:
            tabs += [fitz.Rect(t.bbox) for t in page.find_tables().tables]
        except Exception:
            non_cherches = 1
    else:
        non_cherches = 1
    tf = []
    for d in page.get_drawings():
        r = d.get('rect')
        if r is None:
            continue
        if r.height < 1.0 and r.width >= 8:
            tf.append(('h', round(r.x0, 0), round(r.x1, 0), (r.y0 + r.y1) / 2))
        elif r.width < 1.0 and r.height >= 8:
            tf.append(('v', (r.x0 + r.x1) / 2, r.y0, r.y1))
    grp = {}
    for t in tf:
        if t[0] == 'h':
            grp.setdefault((t[1], t[2]), []).append(t[3])
    for (x0, x1), ys in grp.items():
        if len(ys) < 3:
            continue
        y0, y1 = min(ys), max(ys)
        nv = sum(1 for t in tf if t[0] == 'v' and x0 - 2 <= t[1] <= x1 + 2
                 and t[2] >= y0 - 2 and t[3] <= y1 + 2)
        if nv >= 2:
            tabs.append(fitz.Rect(x0, y0, x1, y1))
    vus = set()
    for i, c in enumerate(cases):
        if i in vus:
            continue
        grp2 = [i]
        for j, c2 in enumerate(cases):
            if j == i or j in vus:
                continue
            if ((abs(c2.height - c.height) <= 0.15 * c.height
                 and min(c.y1, c2.y1) - max(c.y0, c2.y0) >= 0.5 * min(c.height, c2.height))
                    or (abs(c2.width - c.width) <= 0.15 * c.width
                        and min(c.x1, c2.x1) - max(c.x0, c2.x0) >= 0.5 * min(c.width, c2.width))):
                grp2.append(j)
        if len(grp2) >= 3:
            vus.update(grp2)
            tabs.append(fitz.Rect(min(cases[j].x0 for j in grp2), min(cases[j].y0 for j in grp2),
                                  max(cases[j].x1 for j in grp2), max(cases[j].y1 for j in grp2)))
    # un tableau contenu dans un autre n'est pas un second tableau
    tabs = [t for t in tabs if not any(o is not t and o.contains(t) and o != t for o in tabs)]
    return tabs, non_cherches


def _v247_porte_un_glyphe(page, r):
    """Vrai si le rectangle porte un mot court sans ponctuation (un chiffre, une étiquette) :
    un tableau qui en porte est une figure ; un tableau qui ne porte que des lignes de texte est du texte."""
    try:
        mots, L, texte = _v247_lignes(page)
    except Exception:
        return False
    for i, w in enumerate(mots):
        if not _v247_est_glyphe(i, mots, L, texte):
            continue
        b = fitz.Rect(w[:4])
        if r.contains(fitz.Point((b.x0 + b.x1) / 2, (b.y0 + b.y1) / 2)):
            return True
    return False


def _v247_fusionner_tableaux(tabs):
    """Deux tableaux qui se recouvrent sont un seul tableau (les cellules d'un quadrillage, les rangées et les colonnes d'une même grille)."""
    tabs = [fitz.Rect(t) for t in tabs]
    change = True
    while change:
        change = False
        out = []
        for t in tabs:
            for i, o in enumerate(out):
                if not (o & t).is_empty:
                    out[i] = o | t
                    change = True
                    break
            else:
                out.append(t)
        tabs = out
    return tabs


def _v247_etiqueter(mask, huit=True):
    """Composantes d'un masque : union-find par segments (numpy). 8-connexes par défaut (la découpe, 2.47) ;
    huit=False : 4-connexes (E1, v2.52 — celles de scipy.ndimage.label)."""
    h, w = mask.shape
    lab = _np247.zeros((h, w), dtype=_np247.int32)
    par = [0]
    nxt = 1
    prev = []
    for y in range(h):
        row = mask[y]
        if not row.any():
            prev = []
            continue
        d = _np247.diff(_np247.concatenate(([0], row.astype(_np247.int8), [0])))
        starts = _np247.where(d == 1)[0]
        ends = _np247.where(d == -1)[0]
        cur = []
        for s, e in zip(starts, ends):
            lbl = 0
            for (ps, pe, pl) in prev:
                if (ps <= e and pe >= s) if huit else (ps < e and pe > s):
                    if lbl == 0:
                        lbl = pl
                    else:
                        a, b = lbl, pl
                        while par[a] != a:
                            a = par[a]
                        while par[b] != b:
                            b = par[b]
                        if a != b:
                            par[max(a, b)] = min(a, b)
            if lbl == 0:
                lbl = nxt
                par.append(nxt)
                nxt += 1
            lab[y, s:e] = lbl
            cur.append((s, e, lbl))
        prev = cur
    racine = _np247.arange(nxt)
    for i in range(1, nxt):
        a = i
        while par[a] != a:
            a = par[a]
        racine[i] = a
    return racine[lab]


# ─── v2.52 — E1 : L'ENCRE DES IMAGES EN LIGNE. Sur une page qui porte au moins une image en ligne : l'encre de TOUS ses
# blocs image, sur un bitmap à 144 ppp, réunie à 1 mm (dilatation en croix), composantes 4-connexes ≥ 4 mm de côté
# (l'étiquetage de la 2.47). (voir JOURNAL BACKEND v2.52) ───
_V252_PPP_ENCRE = 144
_V252_REUNION_MM = 1.0
_V252_PLANCHER_MM = 4.0


def _v252_dilate_croix(B, iterations):
    for _ in range(iterations):
        D = B.copy()
        D[1:, :] |= B[:-1, :]
        D[:-1, :] |= B[1:, :]
        D[:, 1:] |= B[:, :-1]
        D[:, :-1] |= B[:, 1:]
        B = D
    return B


def _v252_encre_du_bloc(bl):
    """La carte d'encre d'un bloc image (True = pixel non blanc), ou None si l'image ne se décode pas."""
    try:
        pm = fitz.Pixmap(bl["image"])
        if pm.n - pm.alpha >= 3:
            pm = fitz.Pixmap(fitz.csGRAY, pm)
        a = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, 0]
        return a < 200
    except Exception:
        return None


def _v252_fenetres_encre(page):
    """Les fenêtres d'encre (fitz.Rect, pt) d'une page qui porte au moins une image en ligne ; sinon rien."""
    try:
        infos = page.get_image_info(xrefs=True)
    except Exception:
        return []
    if not any(not it.get("xref") for it in infos):
        return []
    K = _V252_PPP_ENCRE / 72.0
    W, H = int(page.rect.width * K) + 1, int(page.rect.height * K) + 1
    B = _np247.zeros((H, W), dtype=bool)
    for bl in page.get_text("dict")["blocks"]:
        if bl.get("type") != 1:
            continue
        r = fitz.Rect(bl["bbox"])
        x0, y0 = int(r.x0 * K), int(r.y0 * K)
        x1, y1 = max(int(r.x1 * K) + 1, x0 + 1), max(int(r.y1 * K) + 1, y0 + 1)
        x0, y0, x1, y1 = max(0, x0), max(0, y0), min(W, x1), min(H, y1)
        if x1 <= x0 or y1 <= y0:
            continue
        m = _v252_encre_du_bloc(bl)
        if m is None:
            B[y0:y1, x0:x1] = True
            continue
        hh, ww = y1 - y0, x1 - x0
        if m.shape != (hh, ww):
            im = _Im247.fromarray(m.astype(_np247.uint8) * 255).resize((max(ww, 1), max(hh, 1)), _Im247.NEAREST)
            m = _np247.array(im) > 127
        B[y0:y1, x0:x1] |= m[:hh, :ww]
    rad = max(1, int(round((_V252_REUNION_MM * 72.0 / 25.4) * K / 2)))
    B = _v252_dilate_croix(B, rad)
    lab = _v247_etiqueter(B, huit=False)
    plancher = int(_V252_PLANCHER_MM * 72.0 / 25.4 * K)
    ys, xs = _np247.nonzero(lab)
    if len(ys) == 0:
        return []
    labs = lab[ys, xs]
    n = int(labs.max()) + 1
    miny = _np247.full(n, H); maxy = _np247.full(n, -1); minx = _np247.full(n, W); maxx = _np247.full(n, -1)
    _np247.minimum.at(miny, labs, ys); _np247.maximum.at(maxy, labs, ys)
    _np247.minimum.at(minx, labs, xs); _np247.maximum.at(maxx, labs, xs)
    out = []
    for l in _np247.unique(labs):
        y0, y1 = int(miny[l]) + rad, int(maxy[l]) + 1 - rad
        x0, x1 = int(minx[l]) + rad, int(maxx[l]) + 1 - rad
        if x1 - x0 >= plancher and y1 - y0 >= plancher:
            out.append(fitz.Rect(x0 / K, y0 / K, x1 / K, y1 / K))
    return out


def _v252_couverte(rect, fenetres, K=2.0):
    """La part (0-1) de `rect` couverte par l'union des `fenetres` (rasterisée à 2 px/pt, comme le juge du banc)."""
    w, h = int(rect.width * K) + 1, int(rect.height * K) + 1
    if w <= 1 or h <= 1:
        return 0.0
    M = _np247.zeros((h, w), dtype=bool)
    for f in fenetres:
        q = rect & fitz.Rect(f)
        if q.is_empty:
            continue
        M[int((q.y0 - rect.y0) * K):int((q.y1 - rect.y0) * K) + 1, int((q.x0 - rect.x0) * K):int((q.x1 - rect.x0) * K) + 1] = True
    return float(M.mean())


def _v247_absorber(rects):
    rects = sorted(rects, key=lambda r: -(r[2] - r[0]) * (r[3] - r[1]))
    gardes = []
    for r in rects:
        if any(r[0] >= q[0] and r[1] >= q[1] and r[2] <= q[2] and r[3] <= q[3] for q in gardes):
            continue
        gardes.append(r)
    return gardes


def _v247_dilate(m, r, E):
    return (_np247.array(_Im247.fromarray((m * 255).astype(_np247.uint8)).filter(_If247.MaxFilter(2 * r + 1))) > 0) & E


def _v249_masque_contour(chemins, R, K, H_, W_, E):
    """v2.49 : le masque des contours de décor — chaque segment, côté de rectangle ou courbe tracé à ± 2 px, en pixels de la région R à l'échelle K ; & E."""
    im = _Im247.new('L', (W_, H_), 0); dr = _Id249.Draw(im)
    def P(p): return ((p.x - R.x0) * K, (p.y - R.y0) * K)
    def bez(p0, p1, p2, p3, n=16):
        pts = []
        for k in range(n + 1):
            t = k / n; u = 1 - t
            pts.append((u*u*u*p0[0] + 3*u*u*t*p1[0] + 3*u*t*t*p2[0] + t*t*t*p3[0], u*u*u*p0[1] + 3*u*u*t*p1[1] + 3*u*t*t*p2[1] + t*t*t*p3[1]))
        return pts
    for items in chemins:
        for it in items:
            try:
                if it[0] == 'l': dr.line([P(it[1]), P(it[2])], fill=255, width=5)
                elif it[0] == 're':
                    q = it[1]; a, b, c, d = P(fitz.Point(q.x0, q.y0)), P(fitz.Point(q.x1, q.y0)), P(fitz.Point(q.x1, q.y1)), P(fitz.Point(q.x0, q.y1)); dr.line([a, b, c, d, a], fill=255, width=5)
                elif it[0] == 'qu':
                    q = it[1]; dr.line([P(q.ul), P(q.ur), P(q.lr), P(q.ll), P(q.ul)], fill=255, width=5)
                elif it[0] == 'c': dr.line(bez(P(it[1]), P(it[2]), P(it[3]), P(it[4])), fill=255, width=5)
            except Exception:
                pass
    return (_np247.array(im) > 0) & E


def _v247_decouper_page(page, fenetres, tableaux, mots, L, texte, cases, decor):
    """Les morceaux de chaque fenêtre de la page.
    fenetres : liste de dict {clip: Rect, tableau: bool, ...} ; rend, pour chaque fenêtre,
    {'pieces': [Rect pt], 'mots': [indices de mots], 'entier': Rect, 'sans_dessin': bool}
    et le compte des mots attachés."""
    # v2.60 — la boîte d'un mot de ligne ou ponctué s'étend à l'encre de ses lettres
    mots = _v260_mots_etendus(page, mots, L, texte)
    K = _V247_K
    n = len(fenetres)
    glyphes = [i for i in range(len(mots)) if _v247_est_glyphe(i, mots, L, texte)]
    boites = {i: fitz.Rect(mots[i][:4]) for i in glyphes}
    # 1. l'appartenance à un tableau : sans distance
    attribution = {}          # mot -> fenêtre
    for k, f in enumerate(fenetres):
        if not f.get('tableau'):
            continue
        T = f['clip']
        for i in glyphes:
            b = boites[i]
            if i in attribution:
                continue
            if T.contains(fitz.Point((b.x0 + b.x1) / 2, (b.y0 + b.y1) / 2)):
                attribution[i] = k
    # 2. pour chaque fenêtre : la région, l'encre, les intrus, les éléments, le dessin
    etats = []
    for k, f in enumerate(fenetres):
        W = f['clip']
        R = fitz.Rect(W.x0 - _V247_MARGE, W.y0 - _V247_MARGE, W.x1 + _V247_MARGE, W.y1 + _V247_MARGE) & page.rect
        pm = page.get_pixmap(matrix=fitz.Matrix(K, K), clip=R, alpha=False)
        A = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
        E = (255 - A.min(axis=2)) >= _V247_ECART
        H_, W_ = E.shape

        def px(bx):
            q = bx & R
            return (int((q.x0 - R.x0) * K), int((q.y0 - R.y0) * K),
                    int(_np247.ceil((q.x1 - R.x0) * K)), int(_np247.ceil((q.y1 - R.y0) * K)))

        def masque_de(rects, anneau=0, dil=0):
            m = _np247.zeros_like(E)
            for r in rects:
                q = fitz.Rect(r) & R
                if q.is_empty:
                    continue
                x0, y0 = int((q.x0 - R.x0) * K) - dil, int((q.y0 - R.y0) * K) - dil
                x1, y1 = int(_np247.ceil((q.x1 - R.x0) * K)) + dil, int(_np247.ceil((q.y1 - R.y0) * K)) + dil
                if anneau:
                    m[max(0, y0 - anneau):min(H_, y0 + anneau), max(0, x0 - anneau):min(W_, x1 + anneau)] = True
                    m[max(0, y1 - anneau):min(H_, y1 + anneau), max(0, x0 - anneau):min(W_, x1 + anneau)] = True
                    m[max(0, y0 - anneau):min(H_, y1 + anneau), max(0, x0 - anneau):min(W_, x0 + anneau)] = True
                    m[max(0, y0 - anneau):min(H_, y1 + anneau), max(0, x1 - anneau):min(W_, x1 + anneau)] = True
                else:
                    m[max(0, y0):min(H_, y1), max(0, x0):min(W_, x1)] = True
            return m & E
        def _cp_texte(i):   # v2.72 C′ + E′ (DEC-3208, DEC-2977) : un mot de calcul, ou de trois lettres et plus, est du texte : ses lettres sont ôtées de l'encre comme celles d'une ligne de texte
            _ct = (mots[i][4] or '').strip()
            if _V247_MOT.search(mots[i][4] or ''):   # v2.72 E′
                return True
            return bool(_C_SIGNE.fullmatch(_ct) or (_ct in ('x', '×') and _C_CALCUL.search(texte[i] or '')) or (_C_NOMBRE.fullmatch(_ct) and _C_CALCUL.search(texte[i] or '')))
        lignes_txt = [fitz.Rect(mots[i][:4]) for i in range(len(mots))
                      if (_v247_est_ligne(i, L, texte) or _cp_texte(i)) and not (fitz.Rect(mots[i][:4]) & R).is_empty]
        ponct = [fitz.Rect(mots[i][:4]) for i in range(len(mots))
                 if (not _v247_est_ligne(i, L, texte)) and mots[i][4] and mots[i][4][-1] in _V247_PONCT
                 and not (fitz.Rect(mots[i][:4]) & R).is_empty]
        # v2.54 — les lettres des lignes de texte et des mots ponctués, pas leurs boîtes
        T = (masque_de(lignes_txt) | masque_de(ponct)) & _v254_lettres_region(page, R, E.shape)
        # v2.49 — l'anneau du décor suit son contour, pas le rectangle qui l'entoure : les coins arrondis d'un cadre pointillé n'échappent plus
        D = _v247_dilate(_v249_masque_contour([it for (rd, it) in decor if not (rd & R).is_empty], R, K, H_, W_, E), 2, E)
        # v2.54 — le décor : son trait (la bande ci-dessus reste pour une page tournée)
        D = (_v254_decor_region(page, R, E.shape, D) & E) if decor else D
        # l'intérieur n'est jamais effacé : une fenêtre qui contient W n'est pas un intrus ; une fenêtre contenue dans une
        # fenêtre-tableau non plus (elle lui appartient) ; un tableau qui tient la moitié de W n'est pas un intrus pour W ;
        # v2.48 — deux dessins qui se chevauchent ne sont jamais des intrus l'un pour l'autre : une fenêtre qui croise W
        # (aire non nulle) n'est pas retirée de W
        autres = [g['clip'] for j, g in enumerate(fenetres) if j != k and not g['clip'].contains(W)
                  and not (f.get('tableau') and W.contains(g['clip']))
                  and not (g.get('tableau') and (W & g['clip']).get_area() >= 0.5 * W.get_area())
                  and (g['clip'] & W).is_empty]
        O = masque_de(autres)
        F0 = E & ~(T | D | O)
        lab0 = _v247_etiqueter(_v247_dilate(F0, 1, _np247.ones_like(E, dtype=bool)))
        cases_isolees = []
        for c in cases:
            q = c & R
            if q.is_empty or c.contains(W) or _v247_en_bande(c, cases):
                continue
            x0, y0, x1, y1 = px(c)
            labs = _np247.unique(lab0[max(0, y0 - 2):min(H_, y1 + 2), max(0, x0 - 2):min(W_, x1 + 2)])
            labs = labs[labs > 0]
            if not len(labs):
                continue
            dehors = 0
            for lb in labs:
                ys, xs = _np247.where(lab0 == lb)
                if xs.min() < x0 - 4 or ys.min() < y0 - 4 or xs.max() > x1 + 4 or ys.max() > y1 + 4:
                    dehors += 1
            if dehors == 0:
                cases_isolees.append(c)
        # v2.48 — on n'efface jamais ce qu'il y a dans une case : une case isolée n'efface que son bord, l'anneau de ± 2 px autour de sa boîte
        C = masque_de(cases_isolees, anneau=2)
        I = T | C | D | O
        F = E & ~I
        lab = _v247_etiqueter(_v247_dilate(F, 1, _np247.ones_like(E, dtype=bool)))
        wx0, wy0, wx1, wy1 = px(W)
        labs = _np247.unique(lab[wy0:wy1, wx0:wx1])
        labs = labs[labs > 0]
        corps = _np247.isin(lab, labs) & F if len(labs) else _np247.zeros_like(F)
        G = masque_de([boites[i] for i in glyphes if not (boites[i] & R).is_empty], dil=1)
        dessin = corps & ~G
        etats.append(dict(R=R, E=E, F=F, I=I, corps=corps, dessin=dessin, a_dessin=bool(dessin.any()), lab=lab, labs=labs, px=px))
    # 3. les mots à 6 mm ou moins : la fenêtre la plus proche, une seule
    for i in glyphes:
        if i in attribution:
            continue
        b = boites[i]
        meilleur = None
        portee = int(_np247.ceil(_V247_DIST_MM / _V247_MM * K)) + 1   # 6 mm en pixels : seule l'encre à cette portée compte
        for k, st in enumerate(etats):
            R = st['R']
            if (b & R).is_empty or not st['a_dessin']:
                continue
            x0, y0, x1, y1 = st['px'](b)
            des = st['dessin']
            sy0, sy1 = max(0, y0 - portee), min(des.shape[0], y1 + portee)
            sx0, sx1 = max(0, x0 - portee), min(des.shape[1], x1 + portee)
            tranche = des[sy0:sy1, sx0:sx1]
            if not tranche.any():
                continue
            ys, xs = _np247.where(tranche)
            ys = ys + sy0
            xs = xs + sx0
            dx = _np247.maximum(0, _np247.maximum(x0 - xs, xs - (x1 - 1)))
            dy = _np247.maximum(0, _np247.maximum(y0 - ys, ys - (y1 - 1)))
            d = float(_np247.sqrt(dx * dx + dy * dy).min()) / K * _V247_MM
            if d <= _V247_DIST_MM and (meilleur is None or d < meilleur[0]):
                meilleur = (d, k)
        if meilleur is not None:
            attribution[i] = meilleur[1]
    # 4. les morceaux
    resultats = []
    for k, f in enumerate(fenetres):
        st = etats[k]
        R, corps = st['R'], st['corps']
        mots_k = [i for i, kk in attribution.items() if kk == k]
        # v2.54 — les rectangles des dessins dont cette fenêtre est la première
        rects = _v254_rectangles(page, k, st, fenetres, etats, mots, L, texte, attribution, decor)
        K = _V247_K
        pieces = [fitz.Rect(R.x0 + x0 / K, R.y0 + y0 / K, R.x0 + x1 / K, R.y0 + y1 / K) for (x0, y0, x1, y1) in rects]
        pieces.sort(key=lambda r: (round(r.y0), r.x0))
        entier = None
        if pieces:
            entier = fitz.Rect(min(p.x0 for p in pieces), min(p.y0 for p in pieces),
                               max(p.x1 for p in pieces), max(p.y1 for p in pieces))
        resultats.append(dict(pieces=pieces, mots=mots_k, entier=entier, sans_dessin=not bool(corps.any())))
    return resultats, len(attribution)


def _v247_absorber_rects(rects):
    """Un rectangle (pt) contenu dans un autre de la même figure disparaît."""
    rects = sorted(rects, key=lambda r: -(r.width * r.height))
    gardes = []
    for r in rects:
        if any(q.contains(r) for q in gardes):
            continue
        gardes.append(r)
    return gardes


# ══════════════════════════════════════════════════════════════════════════════
# v2.54, v2.55 — LA DÉCOUPE PAR DESSIN. Ce que ce bloc fait, et rien d'autre :
#   · un dessin : les taches des fenêtres qui partagent un pixel ou sont à 3 px ou moins, d'une fenêtre à l'autre ; l'encre
#     libre qui le touche ; ses mots rattachés et tout mot court à 6 mm ou moins ; un tableau : les dessins et l'encre libre
#     dont la moitié au moins est dans son rectangle ; l'encre libre qui ne touche rien (9 px au moins) est un dessin ;
#   · chaque dessin est coupé en rectangles qui ne prennent pas l'encre d'un autre, posés à leur place exacte ;
#   · une figure par dessin ; puis, par la lecture des exercices de la page, une figure par exercice (une page dont la lecture
#     est en panne garde une figure par dessin) ; les morceaux d'un tableau qui ne se touchent pas sont des figures séparées ;
#   · chaque image ne montre que son dessin : l'encre d'un dessin d'une autre figure y est blanchie, dans l'image du modèle
#     et dans celle de la feuille ;
#   · l'état du document en cours est propre au fil qui le lit.
# (voir JOURNAL BACKEND v2.54, v2.55, v2.56, v2.59)
# ══════════════════════════════════════════════════════════════════════════════
import threading as _th254
_V254 = _th254.local()
_V254_LECTURE_MODELE = "claude-sonnet-5-5" # le lecteur des exercices
_V254_LECTURE_BORNE = 32000                # la borne de la réponse, en jetons
_V254_LECTURE_ATTENTE = 900.0              # l'attente d'une réponse, en secondes
# la demande ; la liste des dessins numérotés (P1 · P2 · …) la suit
_V254_DEMANDE = 'Voici une page de fiche d\'école. Des MORCEAUX de dessin y sont entourés en orange et numérotés P1, P2… 1) Liste les EXERCICES de la page, dans l\'ordre : pour chacun, son numéro tel qu\'il est écrit sur la page (ou, s\'il n\'a pas de numéro écrit, "sans numéro A", "sans numéro B"… : une lettre différente pour chaque exercice sans numéro, dans l\'ordre de la page) et les premiers mots de sa consigne. 2) Pour CHAQUE morceau, dis à quel exercice il appartient (son numéro), ou "aucun" s\'il n\'appartient à aucun exercice (en-tête, titre, décor). Réponds UNIQUEMENT par un bloc JSON : {"exercices":[{"numero":"1","consigne":"…"}],"morceaux":[{"n":1,"exercice":"1"}]} — chaque numéro de morceau EXACTEMENT une fois. Les morceaux : '
# reconnaissance : les formes écrites d'un exercice sans numéro
_V254_SANS_NUMERO = ("sans numéro", "sans numero", "sans num", "sans numéro.", "non numéroté", "non numerote")


def _v254_etat_neuf():
    """Un état vide pour le prochain document de ce fil."""
    _V254.etat = {"r1": {"doc": None}, "lettres": {"doc": None}, "decor": {"doc": None}, "lecture": {"doc": None}}
    return _V254.etat


def _v254_etat():
    """L'état du document en cours, propre à ce fil."""
    e = getattr(_V254, "etat", None)
    return e if e is not None else _v254_etat_neuf()


def _v254_aire(c):
    return max(0.0, c[2] - c[0]) * max(0.0, c[3] - c[1])


def _v254_renumerote(page, ims, idx_fin, base=None):
    """Les images de la page renumérotées (index continus depuis `base`, le premier index de la page) ; morceau, n_morceaux
    et l'image entière refaits pour chaque figure."""
    if not ims:
        return ims, (base if base is not None else idx_fin)
    base = ims[0]["index"] if base is None else base
    for k, im in enumerate(ims):
        im["index"] = base + k
    figs = {}
    for im in ims:
        if im.get("figure"):
            figs.setdefault(im["figure"], []).append(im)
    for f, L in figs.items():
        for m, im in enumerate(L, 1):
            im["morceau"] = m; im["n_morceaux"] = len(L)
            im.pop("entier", None); im.pop("entier_cadre", None)
        if len(L) > 1:
            e = fitz.Rect(min(i["cadre"][0] for i in L), min(i["cadre"][1] for i in L),
                          max(i["cadre"][2] for i in L), max(i["cadre"][3] for i in L)) & page.rect
            pe = page.get_pixmap(clip=e, matrix=fitz.Matrix(2, 2))
            L[0]["entier"] = "data:image/png;base64," + base64.b64encode(pe.tobytes("png")).decode()
            L[0]["entier_cadre"] = [e.x0, e.y0, e.x1, e.y1]
    return ims, base + len(ims)


# ══ v2.60 — À LA SORTIE D'UNE PAGE, LA MINIATURE, LES ANCRES, LA PAGE SCANNÉE. Ce que ce bloc fait, et rien d'autre :
#   · F-c à l'image, après la séparation : une image de 36 pt au plus qui contient le centre d'un numéro écrit est retirée ; la zone
#     de ce que F-c retire, élargie de 6 pt : toute image à 80 % dedans est retirée ; la pastille prise dans une grande image y est
#     effacée (l'encre qui touche le numéro, si elle tient en 36 pt) ;
#   · une image d'exercice à 80 % dans un surlignage est retirée ;
#   · un glyphe (1 à 3 chiffres ou signes) dont l'encre n'est dans aucune image, à 6 mm d'un dessin d'exercice, devient un morceau de
#     la figure la plus proche — jamais le chiffre d'un numéro, jamais un glyphe d'une ligne de mots, jamais sur une autre figure ;
#   · P2 lit la ligne de son premier à son dernier caractère visible ;
#   · le décor : le trait gardé là où il se voit, le soulignement d'adresse, le mot d'adresse, l'encre des lettres jusqu'au bout ;
#   · l'image entière d'une figure est faite de ses morceaux posés ;
#   · le marqueur d'une figure d'exercice suit la consigne de son exercice quand elle est trouvée dans le texte de la page ;
#   · une page qu'une image couvre à 85 % et qui porte moins de 20 mots est déclarée scannée.
# (voir JOURNAL BACKEND v2.60)
_V260_ADR = re.compile(r"(https?:|www\.|\.fr\b|\.com\b|\.org\b|\.net\b)", re.I)
_V260_GLYPHE = re.compile(r"[\d+\-−=×x:<>÷]+")
_V260_DEUX_LETTRES = re.compile(r"[A-Za-zÀ-ÿ]{2,}")
_V260_ZONE = 6.0          # pt : la zone d'un numéro autour de ce que F-c retire
_V260_DANS = 0.8          # une image à 80 % dans une zone ou un surlignage
_V260_GRIS = 200          # l'encre d'un glyphe : gris < 200 à 144 ppp
_V260_PRES = 17.0         # pt (6 mm) : un glyphe près d'un dessin d'exercice
_V260_PAREIL = 30         # le décor redessiné et la page, pareils à 30 près (somme des trois couleurs)
_V260_SCAN_PART = 0.85    # une image qui couvre 85 % de la page
_V260_SCAN_MOTS = 20      # moins de 20 mots
_V261_PALE = 231          # un aplat pâle : les trois couleurs du remplissage à 231 ou plus
_V261_BORD = 15.0         # pt : une grille de notation est à plus de 15 pt des bords gauche et droit de la page


def _v260_cache(page, nom):
    """Un cache par document et par page, dans l'état du document en cours."""
    E = _v254_etat().setdefault("v260", {})
    if E.get("doc") is not page.parent:
        E.clear(); E["doc"] = page.parent
    return E.setdefault(nom, {})


def _v260_png(arr):
    b = io.BytesIO(); _Im247.fromarray(arr).save(b, format="PNG")
    return "data:image/png;base64," + base64.b64encode(b.getvalue()).decode()


def _v260_rgb(data):
    return _np247.asarray(_Im247.open(io.BytesIO(base64.b64decode(data.split(",", 1)[1]))).convert("RGB"))


def _v260_dans(c, z):
    a = _v254_aire(c)
    return a > 0 and _v254_aire([max(c[0], z[0]), max(c[1], z[1]), min(c[2], z[2]), min(c[3], z[3])]) / a >= _V260_DANS


def _v260_reperes(page, L, exercices, pno):
    """Les numéros écrits (_v259_reperes) lus sur la boîte visible de chaque ligne (du premier au dernier caractère visible, même
    texte, même bloc, même haut et bas à 3 pt, dans la largeur de la ligne à 0,5 pt) ; les boîtes rendues restent celles des lignes."""
    V = []
    for bi, b in enumerate(page.get_text("rawdict")["blocks"]):
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            ch = [c for sp in l["spans"] for c in sp.get("chars", [])]
            t = "".join(c["c"] for c in ch).strip()
            vis = [c for c in ch if c["c"].strip()]
            if t and vis:
                V.append((t, (min(c["bbox"][0] for c in vis), min(c["bbox"][1] for c in vis),
                              max(c["bbox"][2] for c in vis), max(c["bbox"][3] for c in vis)), bi))
    vu = {}
    for t, bb, bi in L:
        c_ = [vb for vt, vb, vbi in V if vbi == bi and vt == t.strip() and abs(vb[1] - bb[1]) < 3 and abs(vb[3] - bb[3]) < 3
              and vb[0] >= bb[0] - 0.5 and vb[2] <= bb[2] + 0.5]
        if c_:
            vu[tuple(round(v, 2) for v in bb)] = c_[0]
    L2 = [(t, vu.get(tuple(round(v, 2) for v in bb), bb), bi) for t, bb, bi in L]
    orig = {tuple(round(v, 2) for v in vu.get(tuple(round(v, 2) for v in bb), bb)): bb for t, bb, bi in L}
    R_ = _v259_reperes(L2, exercices, pno)
    return {e: (fitz.Rect(orig.get(tuple(round(v, 2) for v in a), a)), fitz.Rect(orig.get(tuple(round(v, 2) for v in b), b)))
            for e, (a, b) in R_.items()}


def _v260_adresses(page):
    C = _v260_cache(page, "adresses")
    if page.number not in C:
        C[page.number] = [w[:4] for w in page.get_text("words") if _V260_ADR.search(w[4])]
    return C[page.number]


def _v260_soulignement_adresse(page, d):
    """Un rectangle rempli de 2 pt au plus d'un côté et 20 pt au moins de l'autre, à 3 pt ou moins d'un mot d'adresse."""
    r = d.get("rect")
    if r is None or not d.get("fill") or min(r.width, r.height) > 2 or max(r.width, r.height) < 20:
        return False
    Z = fitz.Rect(r.x0 - 3, r.y0 - 3, r.x1 + 3, r.y1 + 3)
    return any(Z.intersects(fitz.Rect(w)) for w in _v260_adresses(page))


def _v260_decor_couleur(page):
    """Les tracés du décor redessinés en couleur sur une page vierge : couleur, remplissage, épaisseur, tirets, fermeture, bout ; et,
    dans l'ordre de la page, les aplats pâles non décor (remplissage seul)."""
    C = _v260_cache(page, "decor_couleur")
    if page.number not in C:
        tmp = fitz.open(); p2 = tmp.new_page(width=page.rect.width, height=page.rect.height); sh = p2.new_shape()
        for d in page.get_drawings():
            dec = _v259_est_decor(page, d); fond = d.get("fill")
            pale = (not dec) and fond is not None and min(fond) * 255 >= _V261_PALE
            if not dec and not pale:
                continue
            for it in d.get("items") or []:
                if it[0] == "l":
                    sh.draw_line(it[1], it[2])
                elif it[0] == "re":
                    sh.draw_rect(it[1])
                elif it[0] == "qu":
                    sh.draw_quad(it[1])
                elif it[0] == "c":
                    sh.draw_bezier(it[1], it[2], it[3], it[4])
            if pale:
                sh.finish(color=None, fill=fond, width=0, closePath=bool(d.get("closePath")))
                continue
            lc = d.get("lineCap"); lc = max(lc) if isinstance(lc, (tuple, list)) and lc else (lc or 0)
            trait = d.get("color")
            sh.finish(color=trait, fill=fond if trait is None else None, width=(d.get("width") or 1.0) if trait is not None else 0,
                      dashes=d.get("dashes") if trait is not None else None, closePath=bool(d.get("closePath")), lineCap=int(lc))
        sh.commit(); C[page.number] = (tmp, p2)
    return C[page.number][1]


def _v260_la_ou_il_se_voit(page, R, Mk):
    """Le trait du décor gardé là où la page et le décor redessiné en couleur sont pareils à _V260_PAREIL près."""
    K = _V247_K; p2 = _v260_decor_couleur(page)
    a = page.get_pixmap(matrix=fitz.Matrix(K, K), clip=R, alpha=False); c = p2.get_pixmap(matrix=fitz.Matrix(K, K), clip=R, alpha=False)
    A = _np247.frombuffer(a.samples, dtype=_np247.uint8).reshape(a.height, a.width, a.n)[:, :, :3].astype(int)
    B = _np247.frombuffer(c.samples, dtype=_np247.uint8).reshape(c.height, c.width, c.n)[:, :, :3].astype(int)
    h, w = min(Mk.shape[0], A.shape[0], B.shape[0]), min(Mk.shape[1], A.shape[1], B.shape[1])
    out = Mk.copy(); out[:h, :w] &= _np247.abs(A[:h, :w] - B[:h, :w]).sum(axis=2) <= _V260_PAREIL
    return out


def _v260_mots_etendus(page, mots, L, texte):
    """La boîte d'un mot de ligne ou ponctué étendue aux composantes 8-connexes des lettres de la page qui ont un pixel dedans."""
    C = _v260_cache(page, "lettres")
    if page.number not in C:
        lab = _v247_etiqueter(_v254_lettres_page(page), huit=True)
        ys, xs = _np247.nonzero(lab); l = lab[ys, xs]
        u, inv = _np247.unique(l, return_inverse=True)
        b = _np247.zeros((len(u), 4), dtype=_np247.int64)
        b[:, 0] = lab.shape[1]; b[:, 1] = lab.shape[0]
        _np247.minimum.at(b[:, 0], inv, xs); _np247.minimum.at(b[:, 1], inv, ys)
        _np247.maximum.at(b[:, 2], inv, xs + 1); _np247.maximum.at(b[:, 3], inv, ys + 1)
        C[page.number] = (lab, {int(k): b[i] for i, k in enumerate(u)})
    lab, boites = C[page.number]; K = _V247_K
    out = []
    for i, w in enumerate(mots):
        if _v247_est_ligne(i, L, texte) or (w[4] and w[4][-1] in _V247_PONCT):
            x0, y0, x1, y1 = int(w[0] * K), int(w[1] * K), int(_np247.ceil(w[2] * K)), int(_np247.ceil(w[3] * K))
            sub = lab[max(0, y0):max(0, y1), max(0, x0):max(0, x1)]
            bx = [w[0], w[1], w[2], w[3]]
            for j in _np247.unique(sub[sub > 0]):
                q = boites[int(j)]
                bx = [min(bx[0], q[0] / K), min(bx[1], q[1] / K), max(bx[2], q[2] / K), max(bx[3], q[3] / K)]
            out.append(tuple(bx) + tuple(w[4:]))
        else:
            out.append(w)
    return out


def _v260_numeros(page):
    """Les boîtes des numéros d'exercice écrits repérés sur la page (la lecture du document en cours) ; aucune sans lecture."""
    S = _v254_etat()["lecture"]; pno = page.number + 1
    if S.get("doc") is not page.parent or pno in S.get("en_panne", ()):
        return [], S, pno
    ex = (S.get("exercices") or {}).get(pno) or []
    if not ex:
        return [], S, pno
    return [fitz.Rect(bb) for bb, b2 in _v260_reperes(page, _v259_lignes(page), ex, pno).values()], S, pno


def _v260_efface_pastille(im, n):
    """Dans l'image (modèle et feuille, chacune à sa définition) : l'encre qui touche la boîte du numéro, et ce qu'elle enferme,
    mise au blanc — si elle tient en _V259_PETIT pt dans les deux sens. Rend vrai si l'image a été touchée."""
    c = im["cadre"]; fait = False
    for cle_d in ("data", "data_feuille"):
        if not im.get(cle_d) or "," not in im[cle_d]:
            continue
        A = _v260_rgb(im[cle_d]).copy(); H, W = A.shape[:2]
        sx, sy = W / max(1e-6, c[2] - c[0]), H / max(1e-6, c[3] - c[1])
        bx0, by0 = max(0, int((n.x0 - c[0]) * sx)), max(0, int((n.y0 - c[1]) * sy))
        bx1, by1 = min(W, int(_np247.ceil((n.x1 - c[0]) * sx))), min(H, int(_np247.ceil((n.y1 - c[1]) * sy)))
        if bx1 <= bx0 or by1 <= by0:
            continue
        E = (255 - A.min(axis=2)) >= _V247_ECART
        lab = _v247_etiqueter(E, huit=True)
        ids = _np247.unique(lab[by0:by1, bx0:bx1]); ids = ids[ids > 0]
        if not len(ids):
            continue
        O = _np247.isin(lab, ids); ys, xs = _np247.nonzero(O)
        x0, y0, x1, y1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
        if (x1 - x0) / sx > _V259_PETIT or (y1 - y0) / sy > _V259_PETIT:
            continue
        sub = O[y0:y1, x0:x1]; dehors = _v247_etiqueter(~sub, huit=False)
        bord = set(_np247.unique(_np247.concatenate([dehors[0], dehors[-1], dehors[:, 0], dehors[:, -1]]))) - {0}
        enferme = (~sub) & ~_np247.isin(dehors, list(bord))
        T = _np247.zeros_like(O); T[y0:y1, x0:x1] = sub | enferme
        A[T] = 255
        im[cle_d] = _v260_png(A); fait = True
    return fait


def _v260_pastilles(page, ims, nums, S, pno):
    """F-c à l'image (36 pt au plus, le centre d'un numéro) ; la zone de ce que F-c retire, élargie de 6 pt (80 %)."""
    if not nums:
        return ims
    centres = [((n.x0 + n.x1) / 2, (n.y0 + n.y1) / 2) for n in nums]
    zones = [list(z) for z in ((S.get("v260_fc") or {}).get(pno) or [])]
    garde = []
    for im in ims:
        c = im["cadre"]
        if (c[2] - c[0]) <= _V259_PETIT and (c[3] - c[1]) <= _V259_PETIT and any(c[0] <= x <= c[2] and c[1] <= y <= c[3] for x, y in centres):
            zones.append(list(c))
            continue
        garde.append(im)
    Z = [[z[0] - _V260_ZONE, z[1] - _V260_ZONE, z[2] + _V260_ZONE, z[3] + _V260_ZONE] for z in zones]
    garde = [im for im in garde if not any(_v260_dans(im["cadre"], z) for z in Z)]
    return garde


def _v261_pastilles_grandes(page, ims):
    """À la fin de la page : dans chaque image de la page de plus de _V259_PETIT pt d'un côté qui contient le centre d'un numéro
    repéré, la pastille effacée (_v260_efface_pastille)."""
    nums, S, pno = _v260_numeros(page)
    for n in nums:
        x, y = (n.x0 + n.x1) / 2, (n.y0 + n.y1) / 2
        for im in ims:
            c = im["cadre"]
            if im.get("page") == pno and c[0] <= x <= c[2] and c[1] <= y <= c[3] and ((c[2] - c[0]) > _V259_PETIT or (c[3] - c[1]) > _V259_PETIT):
                _v260_efface_pastille(im, n)


def _v260_surlignages(page, ims):
    """Une image d'exercice à 80 % dans un surlignage (un rectangle rempli non blanc, sans trait, de 30 pt de haut au plus, sous un
    mot de ligne de même hauteur à 4 pt et pas plus large que lui de 8 pt), élargi de 1 pt, est retirée."""
    mots, L, texte = _v247_lignes(page)
    lig = [fitz.Rect(mots[i][:4]) for i in range(len(mots)) if _v247_est_ligne(i, L, texte)]
    Q = []
    for d in page.get_drawings():
        f = d.get("fill"); r = fitz.Rect(d["rect"])
        if not f or min(f) > 0.97 or r.height > 30 or r.is_empty or d.get("color") is not None:
            continue
        if any((r & b).get_area() >= 0.5 * max(1e-6, b.get_area()) and abs(r.height - b.height) <= 4 and r.width <= b.width + 8 for b in lig):
            Q.append(fitz.Rect(r.x0 - 1, r.y0 - 1, r.x1 + 1, r.y1 + 1))
    if not Q:
        return ims
    out = []
    for im in ims:
        if re.match(r"p\d+-e", im.get("figure") or ""):
            c = fitz.Rect(im["cadre"]); A = max(1e-6, c.get_area())
            if any((c & q).get_area() / A >= _V260_DANS for q in Q):
                continue
        out.append(im)
    return out


def _v260_glyphes(page, ims, nums):
    """Les glyphes orphelins à 6 mm d'un dessin d'exercice : un morceau de plus de la figure la plus proche ; puis les mots de la
    rangée orpheline d'une figure (_v262_rangee_orpheline, v2.62) : un morceau de plus de cette figure. Rend (images, voies)."""
    ex = [im for im in ims if re.match(r"p\d+-e", im.get("figure") or "")]
    if not ex:
        return ims, []
    avant = ims
    pno = page.number + 1
    pm = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
    G = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].min(axis=2)
    LIG = [(l["bbox"], "".join(sp["text"] for sp in l["spans"])) for bl in page.get_text("dict")["blocks"] if bl.get("type") == 0
           for l in bl["lines"]]
    def inter(a, b):
        return max(0, min(a[2], b[2]) - max(a[0], b[0])) * max(0, min(a[3], b[3]) - max(a[1], b[1]))
    def dist(a, b):
        dx = max(b[0] - a[2], a[0] - b[2], 0); dy = max(b[1] - a[3], a[1] - b[3], 0); return (dx * dx + dy * dy) ** 0.5
    pris = []
    for w in page.get_text("words"):
        t = w[4].strip()
        if not (1 <= len(t) <= 3) or t[-1] in ".,;:!?" or not _V260_GLYPHE.fullmatch(t):
            continue
        x0, y0, x1, y1 = int(w[0] * 2), int(w[1] * 2), int(_np247.ceil(w[2] * 2)), int(_np247.ceil(w[3] * 2))
        sub = G[max(0, y0):y1, max(0, x0):x1]
        ys, xs = _np247.nonzero(sub < _V260_GRIS)
        if not len(xs):
            continue
        e = [(max(0, x0) + xs.min()) / 2, (max(0, y0) + ys.min()) / 2, (max(0, x0) + xs.max() + 1) / 2, (max(0, y0) + ys.max() + 1) / 2]
        A = max(1e-6, (e[2] - e[0]) * (e[3] - e[1]))
        if max([inter(e, im["cadre"]) / A for im in ims] or [0]) >= 0.5:
            continue
        pres = sorted((dist(e, im["cadre"]), im.get("figure")) for im in ex)
        if not pres or pres[0][0] > _V260_PRES:
            continue
        fg = pres[0][1]
        if any(im.get("figure") != fg and inter(e, im["cadre"]) > 0 for im in ims):
            continue
        cx, cy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
        if any(q.x0 <= cx <= q.x1 and q.y0 <= cy <= q.y1 for q in nums):
            continue
        if any(bb[0] <= cx <= bb[2] and bb[1] <= cy <= bb[3] and _V260_DEUX_LETTRES.search(tx) for bb, tx in LIG):
            continue
        pris.append((fg, fitz.Rect(e) & page.rect))
    voies = []
    for fg, r in pris:
        pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(2, 2))
        if pix.width < 1 or pix.height < 1:
            continue
        png = pix.tobytes("png"); extra, voie = _v238_champs(page, r)
        neuf = {"cadre": [r.x0, r.y0, r.x1, r.y1], "sceau": _hl.sha256(png).hexdigest()[:12], "index": 0, "page": pno,
                "data": "data:image/png;base64," + base64.b64encode(png).decode(), "w": pix.width, "h": pix.height,
                "w_mm": round(r.width * 25.4 / 72, 1), "h_mm": round(r.height * 25.4 / 72, 1),
                "figure": fg, "morceau": 0, "n_morceaux": 0, **extra}
        k = max(i for i, im in enumerate(ims) if im.get("figure") == fg)
        ims = ims[:k + 1] + [neuf] + ims[k + 1:]
        voies.append(voie)
    mots = []
    for w in page.get_text("words"):
        t = w[4].strip()
        if 1 <= len(t) <= 3 and t[-1] not in ".,;:!?" and _V260_GLYPHE.fullmatch(t):
            mots.append(dict(t=t, box=list(w[:4]), ink=_v262_encre(G, w)))
    for w, fg in _v262_rangee_orpheline(mots, avant, ims, [[q.x0, q.y0, q.x1, q.y1] for q in nums], LIG):
        r = fitz.Rect(w["ink"]) & page.rect
        pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(2, 2))
        if pix.width < 1 or pix.height < 1:
            continue
        png = pix.tobytes("png"); extra, voie = _v238_champs(page, r)
        neuf = {"cadre": [r.x0, r.y0, r.x1, r.y1], "sceau": _hl.sha256(png).hexdigest()[:12], "index": 0, "page": pno,
                "data": "data:image/png;base64," + base64.b64encode(png).decode(), "w": pix.width, "h": pix.height,
                "w_mm": round(r.width * 25.4 / 72, 1), "h_mm": round(r.height * 25.4 / 72, 1),
                "figure": fg, "morceau": 0, "n_morceaux": 0, **extra}
        k = max(i for i, im in enumerate(ims) if im.get("figure") == fg)
        ims = ims[:k + 1] + [neuf] + ims[k + 1:]
        voies.append(voie)
    return ims, voies


def _v262_encre(G, w):
    """La boîte de l'encre d'un mot (gris < _V260_GRIS à 144 ppp, en pt) ; None s'il n'a pas d'encre."""
    x0, y0, x1, y1 = int(w[0] * 2), int(w[1] * 2), int(_np247.ceil(w[2] * 2)), int(_np247.ceil(w[3] * 2))
    sub = G[max(0, y0):y1, max(0, x0):x1]
    ys, xs = _np247.nonzero(sub < _V260_GRIS)
    if not len(xs):
        return None
    return [(max(0, x0) + xs.min()) / 2, (max(0, y0) + ys.min()) / 2, (max(0, x0) + xs.max() + 1) / 2, (max(0, y0) + ys.max() + 1) / 2]


def _v262_rangee_orpheline(mots, avant, apres, nums, lignes):
    """Les mots de calcul (1 à 3 chiffres ou signes, sans ponctuation finale) à poser comme morceaux d'une figure d'exercice F : dans
    aucune image (encre à moins de 50 %), ni au centre d'un numéro, ni dans une ligne de mots, centre dans la boîte réunie des
    images de F, recoupés par aucune image d'une autre figure, avec une sœur de leur rangée (recouvrement vertical à 50 % de la plus
    petite hauteur, hors de leur ligne : écart de plus de 1,5 fois leur hauteur) dont l'encre est à 50 % dans une image de F d'avant
    les glyphes, et dont toute la ligne (mots de calcul chaînés par recouvrement vertical à 50 % et écart d'au plus 1,5 fois la plus
    grande hauteur) n'a aucun mot à 50 % dans une image. Rend [(mot, figure)]."""
    def inter(a, b):
        return max(0, min(a[2], b[2]) - max(a[0], b[0])) * max(0, min(a[3], b[3]) - max(a[1], b[1]))
    def part(e_):
        A_ = max(1e-6, (e_[2] - e_[0]) * (e_[3] - e_[1]))
        return max([inter(e_, im["cadre"]) / A_ for im in apres] or [0])
    U = {}
    for im in apres:
        if re.match(r"p\d+-e", im.get("figure") or ""):
            c = im["cadre"]; u = U.get(im["figure"], list(c))
            U[im["figure"]] = [min(u[0], c[0]), min(u[1], c[1]), max(u[2], c[2]), max(u[3], c[3])]
    out = []
    for w in mots:
        e = w.get("ink")
        if e is None or part(e) >= 0.5:
            continue
        cx, cy = (w["box"][0] + w["box"][2]) / 2, (w["box"][1] + w["box"][3]) / 2
        if any(q[0] <= cx <= q[2] and q[1] <= cy <= q[3] for q in nums):
            continue
        if any(bb[0] <= cx <= bb[2] and bb[1] <= cy <= bb[3] and _V260_DEUX_LETTRES.search(tx) for bb, tx in lignes):
            continue
        F = next((f for f, u in U.items() if u[0] <= cx <= u[2] and u[1] <= cy <= u[3]), None)
        if F is None or any(im.get("figure") != F and inter(e, im["cadre"]) > 0 for im in apres):
            continue
        hh = e[3] - e[1]; soeur = False
        for w2 in mots:
            e2 = w2.get("ink")
            if e2 is None or w2 is w:
                continue
            if min(e[3], e2[3]) - max(e[1], e2[1]) < 0.5 * min(hh, e2[3] - e2[1]):
                continue
            if not (e2[2] < e[0] - 1.5 * hh or e2[0] > e[2] + 1.5 * hh):
                continue
            A2 = max(1e-6, (e2[2] - e2[0]) * (e2[3] - e2[1]))
            if any(im.get("figure") == F and inter(e2, im["cadre"]) / A2 >= 0.5 for im in avant):
                soeur = True
                break
        if not soeur:
            continue
        ligne = {id(w)}; frontiere = [w]
        while frontiere:
            a = frontiere.pop(); ea = a["ink"]
            for b in mots:
                eb = b.get("ink")
                if eb is None or id(b) in ligne:
                    continue
                if (min(ea[3], eb[3]) - max(ea[1], eb[1]) >= 0.5 * min(ea[3] - ea[1], eb[3] - eb[1])
                        and max(eb[0] - ea[2], ea[0] - eb[2]) <= 1.5 * max(ea[3] - ea[1], eb[3] - eb[1])):
                    ligne.add(id(b)); frontiere.append(b)
        if any(id(b) in ligne and b.get("ink") is not None and part(b["ink"]) >= 0.5 for b in mots):
            continue
        out.append((w, F))
    return out


def _v260_sortie_page(page, ims, idx_fin):
    """À la sortie d'une page, dans cet ordre : les restes de pastille, les surlignages, les glyphes orphelins ; si une image a été
    retirée, ajoutée ou effacée, les images de la page sont renumérotées. Rend (images, index suivant, voies des images ajoutées)."""
    if not ims:
        return ims, idx_fin, []
    base = ims[0]["index"]; avant = [id(im) for im in ims]
    sceaux = {id(im): (im.get("data"), im.get("data_feuille")) for im in ims}
    nums, S, pno = _v260_numeros(page)
    ims = _v260_pastilles(page, ims, nums, S, pno)
    ims = _v260_surlignages(page, ims)
    ims, voies = _v260_glyphes(page, ims, nums)
    if [id(im) for im in ims] == avant and all(sceaux[id(im)] == (im.get("data"), im.get("data_feuille")) for im in ims):
        return ims, idx_fin, voies
    if not ims:
        return ims, base, voies
    ims, idx = _v254_renumerote(page, ims, idx_fin, base=base)
    return ims, idx, voies


def _v260_entier_compose(images):
    """L'image entière d'une figure de plusieurs morceaux : ses images posées, chacune à son cadre, sur blanc, la plus sombre gardée
    pixel par pixel et couleur par couleur, à _V247_K px par pt ; `entier_cadre` et tout le reste inchangés."""
    K = _V247_K; figs = {}
    for im in images:
        if im.get("figure"):
            figs.setdefault((im.get("page"), im["figure"]), []).append(im)
    for L in figs.values():
        tete = [im for im in L if im.get("entier") and im.get("entier_cadre")]
        if len(L) < 2 or not tete:
            continue
        E = tete[0]["entier_cadre"]
        W_, H_ = max(1, int(round((E[2] - E[0]) * K))), max(1, int(round((E[3] - E[1]) * K)))
        C = _np247.full((H_, W_, 3), 255, _np247.uint8)
        for im in L:
            c = im["cadre"]; w, h = max(1, int(round((c[2] - c[0]) * K))), max(1, int(round((c[3] - c[1]) * K)))
            a = _np247.asarray(_Im247.open(io.BytesIO(base64.b64decode(im["data"].split(",", 1)[1]))).convert("RGB").resize((w, h), _Im247.BILINEAR))
            x0, y0 = int(round((c[0] - E[0]) * K)), int(round((c[1] - E[1]) * K))
            xa, ya, xb, yb = max(0, x0), max(0, y0), min(W_, x0 + w), min(H_, y0 + h)
            if xb > xa and yb > ya:
                C[ya:yb, xa:xb] = _np247.minimum(C[ya:yb, xa:xb], a[ya - y0:yb - y0, xa - x0:xb - x0])
        tete[0]["entier"] = _v260_png(C)
    return images


def _v262_entier_feuille(images):
    """L'image entière de FEUILLE d'une figure de plusieurs images dont la tête porte `entier` et `entier_cadre` : les `data_feuille`
    de ses images collés sans rééchantillonnage, chacun à son rectangle au rendu de feuille (_V238_PPP_FEUILLE), dans le rectangle de
    `entier_cadre` au même rendu, sur blanc, la plus sombre gardée pixel par pixel et couleur par couleur ; une image dont la taille
    n'est pas celle de son rectangle est redimensionnée. Portée par la tête dans `entier_feuille` ; une figure dont une image n'a pas
    de `data_feuille` n'en a pas ; aucune autre image n'en porte ; tout le reste inchangé."""
    m = _V238_PPP_FEUILLE / 72.0; M = fitz.Matrix(m, m); figs = {}
    for im in images:
        im.pop("entier_feuille", None)
        if im.get("figure"):
            figs.setdefault((im.get("page"), im["figure"]), []).append(im)
    for L in figs.values():
        tete = [im for im in L if im.get("entier") and im.get("entier_cadre")]
        if len(L) < 2 or not tete or not all(im.get("data_feuille") for im in L):
            continue
        Eir = (fitz.Rect(tete[0]["entier_cadre"]) * M).irect
        W_, H_ = max(1, Eir.width), max(1, Eir.height)
        C = _np247.full((H_, W_, 3), 255, _np247.uint8)
        for im in L:
            a = _v260_rgb(im["data_feuille"])
            ir = (fitz.Rect(im["cadre"]) * M).irect
            if (ir.height, ir.width) != a.shape[:2]:
                a = _np247.asarray(_Im247.fromarray(a).resize((max(1, ir.width), max(1, ir.height)), _Im247.BILINEAR))
            x0, y0 = ir.x0 - Eir.x0, ir.y0 - Eir.y0
            h, w = a.shape[:2]
            xa, ya, xb, yb = max(0, x0), max(0, y0), min(W_, x0 + w), min(H_, y0 + h)
            if xb > xa and yb > ya:
                C[ya:yb, xa:xb] = _np247.minimum(C[ya:yb, xa:xb], a[ya - y0:yb - y0, xa - x0:xb - x0])
        tete[0]["entier_feuille"] = _v260_png(C)
    return images


def _v260_ancres(page, figs, avec_contexte=False):
    """Une ancre par figure, sur son premier morceau : une figure d'exercice après la consigne de son exercice, quand elle est
    trouvée dans le texte de la page (la ligne de consigne de son numéro écrit, le début de ligne, la recherche) ; toute autre au
    coin de son image entière. v2.64 — avec `avec_contexte`, rend aussi (lieux, lignes de la page, exercices lus), déjà
    calculés ici, pour poser le marqueur après la consigne (voir JOURNAL BACKEND v2.64)."""
    lieux, L_, exs = {}, None, []
    S = _v254_etat()["lecture"]; pno = page.number + 1
    cle = getattr(_grilles, "cle_lecture", None) if _grilles else None
    C = (S.get("consignes") or {}).get(pno) or [] if S.get("doc") is page.parent else []
    if C and cle is not None and pno not in S.get("en_panne", ()):
        L_ = _v259_lignes(page)
        exs = (S.get("exercices") or {}).get(pno) or []
        lieux = {e: b2 for e, (bb, b2) in _v260_reperes(page, L_, exs, pno).items()}
        for part in (_v259_par_debut(L_, C, cle), _v259_par_recherche(page, C, cle)):
            for e, r in part.items():
                lieux.setdefault(e, r)
        noms = {x for x, _ in C}
        lieux = {e: r for e, r in lieux.items() if e in noms}
    out = []
    for f in figs:
        if f.get("morceau", 1) != 1:
            continue
        m = re.match(r"p%d-e(.+)$" % pno, f.get("figure") or "")
        if m and m.group(1) in lieux:
            r = lieux[m.group(1)]
            out.append(dict(f, cadre=[r.x0 + 0.001, r.y0, r.x1, r.y1]))
        else:
            out.append(dict(f, cadre=(f.get("entier_cadre") or f["cadre"])))
    return (out, lieux, L_, exs) if avec_contexte else out


# ══ v2.64 — S3 : LE MARQUEUR D'UNE FIGURE D'EXERCICE APRÈS LA DERNIÈRE LIGNE DE SA CONSIGNE (« la ligne, pas le bloc ») ══
# Condition : une figure d'exercice dont la consigne a un lieu sur la page (`_v260_ancres`), la ligne de ce lieu retrouvée parmi
# les lignes de la page, et un seul bloc du texte assemblé qui la porte. Effet : son marqueur est inséré juste après la dernière
# ligne de la consigne — la ligne du lieu puis, tant qu'elle ne finit pas par . ? ! : …, au plus trois lignes de son bloc, jamais
# au-delà d'un numéro d'exercice lu, d'une autre consigne ou d'un barème — et non plus à la fin du bloc. Sinon, la fin du bloc.
# (voir JOURNAL BACKEND v2.64)
_V264_BAREME = re.compile(r'^[\s_.…]*/\s*\d+([.,]\d+)?\s*(pts?|points?)?\s*[\s_.…]*$|^.*_{2,}\s*/\s*\d+([.,]\d+)?\s*(pts?|points?)?\s*$'
                          r'|^\s*\d+([.,]\d+)?\s*(pts?|points?)\s*$', re.I)


def _v264_ligne_du_lieu(L, r):
    """Le rang, dans les lignes de la page, de la ligne du lieu ; un lieu qui joint deux lignes porte la boîte de la première."""
    for i, (t, bb, bi) in enumerate(L):
        if all(abs(a - b) < 0.02 for a, b in zip(bb, (r[0], r[1], r[2], r[3]))):
            return i
    for i, (t, bb, bi) in enumerate(L):
        if abs(bb[0] - r[0]) < 0.6 and abs(bb[1] - r[1]) < 0.6:
            return i
    return None


def _v264_apres_consigne(texte, blocs, ancres, lieux, L, exercices, pno):
    """Rend {index de figure : position dans le texte, juste après la dernière ligne de sa consigne} ; une figure absente
    du résultat garde son marqueur à la fin de son bloc."""
    out = {}
    if not lieux or not blocs or not L:
        return out
    fins, curseur = [], 0
    for b in blocs:
        pos = texte.find(b[2], curseur)
        if pos < 0:
            return out
        curseur = pos + len(b[2])
        fins.append(curseur)
    il_de = {e: _v264_ligne_du_lieu(L, r) for e, r in lieux.items()}
    nums = {str(x.get("numero", "")).strip().strip(" .°)-:") for x in exercices} - {""}
    for f in ancres:
        m = re.match(r"p%d-e(.+)$" % pno, f.get("figure") or "")
        if not m or m.group(1) not in lieux:
            continue
        e = m.group(1); il = il_de[e]
        if il is None:
            continue
        autres = {_v259_norme(L[j][0]) for k, j in il_de.items() if k != e and j is not None}
        ext = [il]
        if not re.search(r"[.?!:…]\s*$", L[il][0]):
            j = il + 1
            while j < len(L) and L[j][2] == L[il][2] and len(ext) < 4:
                n = _v259_norme(L[j][0])
                if n and (n in autres or n.strip(" .°)-:") in nums or _V264_BAREME.match(L[j][0] or "")):
                    break
                ext.append(j)
                if re.search(r"[.?!:…]\s*$", L[j][0]):
                    break
                j += 1
        t, bb, bi = L[il]; cy = (bb[1] + bb[3]) / 2.0; nt = _v259_norme(t)
        kb = [k for k, b in enumerate(blocs) if len(b) > 3 and b[3] is not None and b[0] - 1.0 <= cy <= b[3] + 1.0
              and b[1] <= bb[0] + 1.0 and nt and nt in _v259_norme(b[2])]
        if len(kb) != 1:
            continue
        tb = blocs[kb[0]][2]; debut = fins[kb[0]] - len(tb)
        lignes_b = tb.split("\n"); cible = _v259_norme(L[ext[-1]][0])
        rang = [i for i, l in enumerate(lignes_b) if _v259_norme(l) == cible]
        if len(rang) != 1:
            # le texte ne suffit pas (consignes identiques) : le rang de la ligne dans son bloc PDF, si les deux découpes se
            # correspondent une à une
            mem = [j for j, x in enumerate(L) if x[2] == L[ext[-1]][2]]
            nonvides = [i for i, l in enumerate(lignes_b) if l.strip()]
            if len(mem) == len(nonvides) and all(_v259_norme(L[j][0]) == _v259_norme(lignes_b[i]) for j, i in zip(mem, nonvides)):
                rang = [nonvides[mem.index(ext[-1])]]
            else:
                continue
        out[f["index"]] = debut + len("\n".join(lignes_b[:rang[0] + 1]))
    return out


def _v260_est_scannee(page):
    """Une page scannée : une image y couvre _V260_SCAN_PART au moins de la surface, et elle porte moins de _V260_SCAN_MOTS mots."""
    A = page.rect.get_area()
    if A <= 0:
        return False
    try:
        grande = any(fitz.Rect(i["bbox"]).get_area() >= _V260_SCAN_PART * A for i in page.get_image_info())
    except Exception:
        return False
    return grande and len(page.get_text("words")) < _V260_SCAN_MOTS


def _v254_separer(page, ims, idx_fin):
    """Dans une figure-tableau, les morceaux qui ne se touchent pas (marge 0,5 pt) sont des figures séparées ; jamais dans une
    figure d'exercice (v2.60)."""
    figs = {}
    for im in ims:
        if im.get("figure") and im.get("tableau") and not re.match(r"p\d+-e", im["figure"]):
            figs.setdefault(im["figure"], []).append(im)
    change = False
    for f, L in figs.items():
        if len(L) < 2:
            continue
        n = len(L); parent = list(range(n))
        def trouve(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]; i = parent[i]
            return i
        for i in range(n):
            a = L[i]["cadre"]
            for j in range(i + 1, n):
                b = L[j]["cadre"]
                if a[0] <= b[2] + 0.5 and b[0] <= a[2] + 0.5 and a[1] <= b[3] + 0.5 and b[1] <= a[3] + 0.5:
                    parent[trouve(i)] = trouve(j)
        comps = {}
        for i in range(n):
            comps.setdefault(trouve(i), []).append(L[i])
        if len(comps) > 1:
            change = True
            for k, (r, G) in enumerate(sorted(comps.items(), key=lambda kv: min(im["index"] for im in kv[1]))):
                if k == 0:
                    continue
                for im in G:
                    im["figure"] = "%s-t%d" % (f, k + 1)
    if not change:
        return ims, idx_fin
    ims.sort(key=lambda im: im["index"])
    return _v254_renumerote(page, ims, idx_fin)


def _v254_couper(A, O, prof_max=10):
    """Les rectangles de A (le dessin) qui ne prennent pas O (l'encre des autres) : coupe en deux le long de la ligne, qui peut
    traverser A, laissant le moins de O dans les deux moitiés resserrées sur A (à égalité : la plus petite surface) ; dix
    niveaux au plus ; s'arrête si rien ne s'améliore. Rend [((x0, y0, x1, y1), reste de O)] en pixels de A."""
    np_ = _np247
    IO = np_.zeros((O.shape[0] + 1, O.shape[1] + 1), dtype=np_.int64)
    IO[1:, 1:] = O.astype(np_.int64).cumsum(0).cumsum(1)
    def cnt(xa, ya, xb, yb):
        return IO[yb, xb] - IO[ya, xb] - IO[yb, xa] + IO[ya, xa]
    out = []
    def serre(x0, y0, x1, y1):
        s = A[y0:y1, x0:x1]
        r = s.any(axis=1); c = s.any(axis=0)
        if not r.any():
            return None
        ya = int(np_.argmax(r)); yb = len(r) - int(np_.argmax(r[::-1])); xa = int(np_.argmax(c)); xb = len(c) - int(np_.argmax(c[::-1]))
        return (x0 + xa, y0 + ya, x0 + xb, y0 + yb)
    def coupes(s, x0, y0, sens):
        # sens 0 : coupes horizontales (lignes) ; sens 1 : verticales (colonnes), par transposition
        t = s if sens == 0 else s.T
        h, w = t.shape
        if h < 2:
            return None
        rows = t.any(axis=1)
        big = w + 1
        rmin = np_.where(rows, t.argmax(axis=1), big)
        rmax = np_.where(rows, w - t[:, ::-1].argmax(axis=1), -1)
        idx = np_.arange(h)
        pmin = np_.minimum.accumulate(rmin); pmax = np_.maximum.accumulate(rmax)
        plast = np_.maximum.accumulate(np_.where(rows, idx, -1)); pfirst = int(np_.argmax(rows))
        smin = np_.minimum.accumulate(rmin[::-1])[::-1]; smax = np_.maximum.accumulate(rmax[::-1])[::-1]
        sfirst = np_.minimum.accumulate(np_.where(rows, idx, h)[::-1])[::-1]; slast = h - 1 - int(np_.argmax(rows[::-1]))
        k = np_.arange(1, h)
        vt = plast[k - 1] >= 0; vb = sfirst[k] < h
        # boîtes dans le repère de t : lignes [a, b), colonnes [c, d)
        ta, tb, tc, td = np_.full(k.shape, pfirst), plast[k - 1] + 1, pmin[k - 1], pmax[k - 1]
        ba, bb, bc, bd = sfirst[k], np_.full(k.shape, slast + 1), smin[k], smax[k]
        def comptes(a, b, c, d, v):
            a = np_.where(v, a, 0); b = np_.where(v, b, 0); c = np_.where(v, np_.minimum(c, w), 0); d = np_.where(v, np_.maximum(d, 0), 0)
            if sens == 0:
                n = cnt(x0 + c, y0 + a, x0 + d, y0 + b)
            else:
                n = cnt(x0 + a, y0 + c, x0 + b, y0 + d)
            return np_.where(v, n, 0), np_.where(v, (b - a) * (d - c), 0)
        nt, at = comptes(ta, tb, tc, td, vt); nb, ab = comptes(ba, bb, bc, bd, vb)
        cout = nt + nb; aire = at + ab
        j = int(np_.lexsort((aire, cout))[0])
        cut = int(k[j])
        if sens == 0:
            return int(cout[j]), int(aire[j]), (x0, y0, x0 + w, y0 + cut), (x0, y0 + cut, x0 + w, y0 + h)
        return int(cout[j]), int(aire[j]), (x0, y0, x0 + cut, y0 + w), (x0 + cut, y0, x0 + h, y0 + w)
    def rec(b, prof):
        b = serre(*b)
        if b is None:
            return
        x0, y0, x1, y1 = b
        n = int(cnt(x0, y0, x1, y1))
        if n == 0 or prof >= prof_max:
            out.append((b, n)); return
        s = A[y0:y1, x0:x1]
        cands = [c for c in (coupes(s, x0, y0, 0), coupes(s, x0, y0, 1)) if c is not None]
        if not cands:
            out.append((b, n)); return
        best = min(cands, key=lambda c: (c[0], c[1]))
        if best[0] >= n:
            out.append((b, n)); return
        rec(best[2], prof + 1); rec(best[3], prof + 1)
    rec((0, 0, A.shape[1], A.shape[0]), 0)
    return out


def _v254_lettres_page(page):
    """Les lettres de la page à _V247_K : les pixels qui changent (écart ≥ _V247_ECART) quand la page est rendue sans son texte
    (une copie du document, le texte seul retiré, faite une fois par document)."""
    np_ = _np247; K = _V247_K
    S = _v254_etat()["lettres"]
    if S.get("doc") is not page.parent:
        d2 = fitz.open(stream=page.parent.tobytes(), filetype="pdf")
        for p2 in d2:
            p2.add_redact_annot(p2.rect, fill=False)
            p2.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_REMOVE)
        S.clear(); S.update(doc=page.parent, d2=d2, pages={})
    if page.number not in S["pages"]:
        a = page.get_pixmap(matrix=fitz.Matrix(K, K), alpha=False); b = S["d2"][page.number].get_pixmap(matrix=fitz.Matrix(K, K), alpha=False)
        A = np_.frombuffer(a.samples, dtype=np_.uint8).reshape(a.height, a.width, a.n)[:, :, :3].astype(np_.int16)
        B = np_.frombuffer(b.samples, dtype=np_.uint8).reshape(b.height, b.width, b.n)[:, :, :3].astype(np_.int16)
        h, w = min(A.shape[0], B.shape[0]), min(A.shape[1], B.shape[1])
        M = np_.zeros(A.shape[:2], dtype=bool)
        M[:h, :w] = np_.abs(A[:h, :w] - B[:h, :w]).max(axis=2) >= _V247_ECART
        S["pages"][page.number] = M
    return S["pages"][page.number]


def _v254_lettres_region(page, R, shape):
    """Les lettres de la page dans la région R (pixels de la région, même origine que le rendu de R)."""
    np_ = _np247; K = _V247_K
    M = _v254_lettres_page(page); ox, oy = int(np_.floor(R.x0 * K + 1e-6)), int(np_.floor(R.y0 * K + 1e-6))
    out = np_.zeros(shape, dtype=bool)
    h = max(0, min(shape[0], M.shape[0] - oy)); w = max(0, min(shape[1], M.shape[1] - ox))
    if h > 0 and w > 0:
        out[:h, :w] = M[oy:oy + h, ox:ox + w]
    return out


# ─── v2.59 — LE DÉCOR POINTILLÉ : LE TRAIT FIN, ET LE RENDU DANS LE CADRAGE DE LA FENÊTRE (voir JOURNAL BACKEND v2.59) ───


def _v259_tirets(d):
    return bool(d.get("dashes")) and str(d.get("dashes")).strip() not in ("", "[] 0")


def _v259_etat_decor(page):
    S = _v254_etat()["decor"]
    if S.get("doc") is not page.parent:
        S.clear(); S.update(doc=page.parent, pages={}, traces={})
    return S


def _v259_touche(page, d):
    """Le trait touche (à 1,5 pt) un tracé non pointillé ou une image."""
    q = d["rect"]; R = fitz.Rect(q.x0 - 1.5, q.y0 - 1.5, q.x1 + 1.5, q.y1 + 1.5)
    for e in page.get_drawings():
        if e is d or _v259_tirets(e):
            continue
        if fitz.Rect(e["rect"]).intersects(R):
            return True
    return any(fitz.Rect(i["bbox"]).intersects(R) for i in page.get_image_info())


# ═══ v2.69 — F081, LE TRAIT : F13 (DEC-3271, E14 ; DEC-3155 « Le trait », sa forme F13 choisie le 18.10) ═══
# Portage de c287/A196_F6/formes/F13/app_v2_62.py (f8015f0e0d870813), l. 2353-2543, noms du banc gardés ; _a196_cotes_contour_B (appelée par
# aucune) n'est pas portée. Pas de plafond de 100 images (c'est la différence avec F14). Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
# ═══ A196 — LE SERVICE À PART (règle neuve, DÉCLARATION : la découpe est celle de la 2.62 ; un trait pointillé fin « à servir » est ajouté APRÈS comme un dessin à lui) ═══
# Forme F13 : (F13 = décision C (cadre simple + trait qui touche une ligne pleine), service en fin de liste.) décision C : comme B2, et un trait qui touche un tracé plein (test écrit avec des bornes) n'est pas servi : il garde la réponse de la 2.62.
def _a196_fin_pointille(d):
    # un tracé pointillé fait de segments droits dont un côté de la boîte fait moins de 8 pt (la branche fine de la règle v2.59, inchangée)
    r = d.get("rect")
    return (r is not None and _v259_tirets(d) and (r.width < 8 or r.height < 8)
            and all(it[0] == "l" for it in (d.get("items") or [])))


def _a196_cle(d):
    return tuple(round(v, 2) for v in d["rect"])


def _a196_touche_bornes(page, d):
    # le test « touche » de la 2.62 (un tracé non pointillé ou une image à 1,5 pt), écrit avec des bornes : fitz.Rect.intersects rend toujours Faux pour un
    # rectangle de largeur ou de hauteur nulle, donc pour toute ligne droite
    q = d["rect"]; R = (q.x0 - 1.5, q.y0 - 1.5, q.x1 + 1.5, q.y1 + 1.5)
    def croise(r):
        return r[0] <= R[2] and R[0] <= r[2] and r[1] <= R[3] and R[1] <= r[3]
    for e in page.get_drawings():
        if e is d or _v259_tirets(e):
            continue
        if croise(tuple(e["rect"])):
            return True
    return any(croise(tuple(i["bbox"])) for i in page.get_image_info())


def _a196_cotes_cadre_simple(page):
    # règle neuve (reconnaissance d'une structure) : un tracé pointillé fin et droit est le côté d'un cadre simple si (1) chacun de ses deux bouts est un COIN : le bout est dans la boîte
    # (1,5 pt) d'un autre tracé de ce genre, perpendiculaire, encore en jeu, ET un bout de cet autre tracé est dans la boîte (1,5 pt) du premier ; (2) l'intérieur du tracé (privé de 3 pt à chaque bout)
    # n'est traversé par aucun tracé perpendiculaire encore en jeu. On retire en boucle ceux qui ne tiennent pas, jusqu'à ce que plus rien ne change. Un quadrillage n'est pas un cadre.
    C = _v260_cache(page, "a196_cadre_simple")
    if page.number not in C:
        T = []
        for d in page.get_drawings():
            if _a196_fin_pointille(d):
                r = d["rect"]
                h = r.width >= r.height
                ym, xm = (r.y0 + r.y1) / 2, (r.x0 + r.x1) / 2
                bouts = ((r.x0, ym), (r.x1, ym)) if h else ((xm, r.y0), (xm, r.y1))
                big = (r.x0 - 1.5, r.y0 - 1.5, r.x1 + 1.5, r.y1 + 1.5)
                inter = (r.x0 + 3.0, r.y0 - 1.5, r.x1 - 3.0, r.y1 + 1.5) if h else (r.x0 - 1.5, r.y0 + 3.0, r.x1 + 1.5, r.y1 - 3.0)
                T.append((_a196_cle(d), h, bouts, big, inter))
        def dans(p, b):
            return b[0] <= p[0] <= b[2] and b[1] <= p[1] <= b[3]
        def recoupe_strict(a, b):
            return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]
        vivants = set(range(len(T)))
        change = True
        while change:
            change = False
            for i in sorted(vivants):
                ti = T[i]; ok = True
                for p in ti[2]:
                    coin = False
                    for j in vivants:
                        if j == i or T[j][1] == ti[1]:
                            continue
                        if dans(p, T[j][3]) and any(dans(q, ti[3]) for q in T[j][2]):
                            coin = True
                            break
                    if not coin:
                        ok = False
                        break
                if ok and ti[4][0] < ti[4][2] and ti[4][1] < ti[4][3]:
                    for j in vivants:
                        if j != i and T[j][1] != ti[1] and recoupe_strict(ti[4], T[j][3]):
                            ok = False
                            break
                if not ok:
                    vivants.discard(i)
                    change = True
        C[page.number] = {T[i][0] for i in vivants}
    return C[page.number]


def _a196_garde_decor(page, d):
    """décision C : comme B2, et un trait qui touche un tracé plein (test écrit avec des bornes) n'est pas servi : il garde la réponse de la 2.62. Vrai : le tracé reste décor (il n'est pas servi)."""
    return (_a196_cle(d) in _a196_cotes_cadre_simple(page)) or _a196_touche_bornes(page, d)


def _a196_traits_a_servir(page):
    # les tracés pointillés fins et droits que la 2.62 juge décor (réponse INCHANGÉE) et que la décision de la forme ne garde pas en décor ; pas de page tournée
    C = _v260_cache(page, "a196_servir")
    if page.number not in C:
        out = []
        if not page.rotation:
            for d in page.get_drawings():
                if not _a196_fin_pointille(d):
                    continue
                if _v260_soulignement_adresse(page, d):
                    continue
                if not _v259_est_decor(page, d):
                    continue
                if _a196_garde_decor(page, d):
                    continue
                out.append(d)
        C[page.number] = out
    return C[page.number]


def _a196_encre_du_trait(page, d, Ep, K):
    # l'encre d'UN tracé : redessiné seul sur une page vierge (épaisseur, tirets, bout gardés), rendu à K, encre >= _V247_ECART, 1 px de marge, restreinte à l'encre réelle de la page
    tmp = fitz.open(); p2 = tmp.new_page(width=page.rect.width, height=page.rect.height); sh = p2.new_shape()
    for it in d.get("items") or []:
        if it[0] == "l":
            sh.draw_line(it[1], it[2])
    lc = d.get("lineCap"); lc = max(lc) if isinstance(lc, (tuple, list)) and lc else (lc or 0)
    sh.finish(color=(0, 0, 0), fill=None, width=d.get("width") or 1.0, dashes=d.get("dashes"), closePath=bool(d.get("closePath")), lineCap=int(lc))
    sh.commit()
    pm = p2.get_pixmap(matrix=fitz.Matrix(K, K), alpha=False)
    a = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
    M = _v247_dilate((255 - a.min(axis=2)) >= _V247_ECART, 1, _np247.ones(a.shape[:2], dtype=bool))
    h, w = min(M.shape[0], Ep.shape[0]), min(M.shape[1], Ep.shape[1])
    out = _np247.zeros(Ep.shape, dtype=bool)
    out[:h, :w] = M[:h, :w] & Ep[:h, :w]
    tmp.close()
    return out


def _a196_ajouter_traits(doc, images):
    # A196 F12/F13 — après toute la découpe : un trait « à servir » = UNE image (le rectangle de son encre, rendu à 144 ppp et à 300 ppp), ajoutée À LA FIN de la liste ;
    # aucun autre dessin n'est touché, aucune image n'est blanchie ; seulement les pages où la découpe a tourné (au moins une fenêtre) ; pas de page tournée
    if not len(doc):
        return images
    pf = _v260_cache(doc[0], "a196_pf")
    K = _V247_K
    idx = max([int(f.get("index", -1)) for f in images], default=-1) + 1
    for pn in sorted(pf):
        page = doc[pn]
        if page.rotation:
            continue
        trs = _a196_traits_a_servir(page)
        if not trs:
            continue
        pm = page.get_pixmap(matrix=fitz.Matrix(K, K), alpha=False)
        a = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
        Ep = (255 - a.min(axis=2)) >= _V247_ECART
        for k, d in enumerate(trs):
            m = _a196_encre_du_trait(page, d, Ep, K)
            ys, xs = _np247.nonzero(m)
            if not len(ys):
                continue
            r = fitz.Rect(xs.min() / K, ys.min() / K, (xs.max() + 1) / K, (ys.max() + 1) / K) & page.rect
            if r.is_empty:
                continue
            pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(2, 2))
            if pix.width < 1 or pix.height < 1:
                continue
            png = pix.tobytes("png")
            extra, _voie = _v238_champs(page, r)
            images.append({
                "cadre": [r.x0, r.y0, r.x1, r.y1],
                "sceau": _hl.sha256(png).hexdigest()[:12],
                "index": idx, "page": pn + 1,
                "data": "data:image/png;base64," + base64.b64encode(png).decode(),
                "w": pix.width, "h": pix.height,
                "w_mm": round(r.width * 25.4 / 72, 1), "h_mm": round(r.height * 25.4 / 72, 1),
                "figure": "p%d-s%d" % (pn + 1, k), "morceau": 1, "n_morceaux": 1, "trait_servi": True,
                **extra,
            })
            idx += 1
    return images


def _v259_est_decor(page, d):
    """Un tracé pointillé est du décor s'il fait 8 pt au moins dans les deux sens ; plus fin : s'il est fait de segments droits et
    ne touche rien. Le soulignement d'une adresse est du décor (v2.60)."""
    if _v260_soulignement_adresse(page, d):
        return True
    r = d.get("rect")
    if r is None or not _v259_tirets(d):
        return False
    if not (r.width < 8 or r.height < 8):
        return True
    return all(it[0] == "l" for it in (d.get("items") or [])) and not _v259_touche(page, d)


def _v259_decor_trace(page):
    """Les tracés du décor redessinés sur une page vierge de la taille de la page, épaisseur et tirets gardés ; gardée par page."""
    S = _v259_etat_decor(page)
    if page.number not in S["traces"]:
        tmp = fitz.open(); p2 = tmp.new_page(width=page.rect.width, height=page.rect.height); sh = p2.new_shape()
        for d in page.get_drawings():
            if not _v259_est_decor(page, d):
                continue
            for it in d.get("items") or []:
                if it[0] == "l":
                    sh.draw_line(it[1], it[2])
                elif it[0] == "re":
                    sh.draw_rect(it[1])
                elif it[0] == "qu":
                    sh.draw_quad(it[1])
                elif it[0] == "c":
                    sh.draw_bezier(it[1], it[2], it[3], it[4])
            lc = d.get("lineCap")
            lc = max(lc) if isinstance(lc, (tuple, list)) and lc else (lc or 0)
            sh.finish(color=(0, 0, 0), fill=None, width=d.get("width") or 1.0, dashes=d.get("dashes"), closePath=bool(d.get("closePath")), lineCap=int(lc))
        sh.commit()
        S["traces"][page.number] = (tmp, p2)
    return S["traces"][page.number][1]


def _v259_rendu_decor(p2, clip):
    """Le trait du décor rendu à _V247_K dans le cadrage `clip` (None : la page entière), encre ≥ _V247_ECART, 1 px de marge."""
    pm = p2.get_pixmap(matrix=fitz.Matrix(_V247_K, _V247_K), clip=clip, alpha=False)
    a = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
    return _v247_dilate((255 - a.min(axis=2)) >= _V247_ECART, 1, _np247.ones(a.shape[:2], dtype=bool))


def _v254_decor_region(page, R, shape, repli):
    """Le trait du décor dans la région R, rendu dans le cadrage même de R ; `repli` pour une page tournée."""
    if page.rotation:
        return repli
    M = _v259_rendu_decor(_v259_decor_trace(page), R)
    out = _np247.zeros(shape, dtype=bool)
    h, w = min(shape[0], M.shape[0]), min(shape[1], M.shape[1])
    out[:h, :w] = M[:h, :w]
    # v2.60 — le trait du décor n'est gardé que là où la page est le décor redessiné, à 30 près
    return _v260_la_ou_il_se_voit(page, R, out)

# ═══ v2.76 — R2c (S122, DEC-3362 rang 2) : UN ENCADRÉ DE TEXTE N'EST PAS UNE FIGURE ═══
# Condition : un dessin ORPHELIN (qui n'appartient à aucun tableau) dont un rectangle a son centre dans un cadre de texte noir de la source, agrandi de 3 pt.
# Le cadre est relevé dans la source (un cadre fermé tracé, isolé, sans image, aplat, courbe, case ni trait vertical), de classe « TC » (>= 4 mots et au moins un
# blanc à compléter) ou « T » (texte seul, >= 2 lignes, >= 6 mots), tous ses caractères étant noirs (aucune composante RVB au-dessus de 60/255). Deux relevés, réunis :
# coins arrondis jusqu'à 30 pt (avec courbes de coin) et jusqu'à 10 pt (sans courbe).
# Effet : ce rectangle est ôté du dessin (avec son encre du masque du dessin), et un dessin sans rectangle est ôté. Le texte du cadre n'est plus posé en image : il reste du texte.
# C'est une déclaration de structure lue dans la source, ni liste de mots ni nom de document. Le filtre s'applique au résultat de _v254_dessins_page, donc aussi quand la gomme
# rejoue parse_pdf. Ce que la fonction lit : page.get_drawings(), page.get_text("words" et "dict"), page.get_image_info(). Code porté du banc (A217_273, i1b/i1/i7b, forme R2c).

_R2C_TOL = 0.8            # tolérance d'alignement d'un trait (pt)
_R2C_TOL_CAD = 2.5        # tolérance de fermeture d'un cadre (pt)
_R2C_MIN_L, _R2C_MIN_H = 36.0, 14.0     # taille minimale d'un cadre (pt)
_R2C_MAX_AIRE = 0.85      # un cadre qui couvre plus de 85 % de la page n'est pas un encadré
_R2C_EPAIS_TRAIT = 1.5    # un fin rectangle plein plus mince que ça est un trait
_R2C_BLANC_MIN = 8.0      # un trait horizontal fin de 8 pt ou plus, dans le cadre, est un blanc à remplir
_R2C_SEUIL_MM = 2.84      # 1 mm en points : une couleur plus large que ça est un aplat
_R2C_PLAFOND_TRACES = 20000   # au-delà, la page n'est pas jouée
_R2C_MOTS_TC, _R2C_MOTS_T, _R2C_LIGNES_T = 4, 6, 2

_R2C_RE_BLANC = re.compile(r"_{2,}|\.{3,}|…")
_R2C_RE_MOT = re.compile(r"[^\W_]", re.U)   # un caractère qui n'est ni espace, ni ponctuation, ni « _ »


def _r2c_proche(a, b, t=_R2C_TOL):
    return abs(a - b) <= t


def _r2c_couleur_blanche(c):
    return c is None or all(float(v) >= 0.97 for v in c)


def _r2c_couleur_noire(c):
    return c is not None and all(float(v) <= 0.03 for v in c)


def _r2c_segments_item(it):
    """segments axe-alignés d'un item de tracé : liste de ('h', y, x0, x1) ou ('v', x, y0, y1)."""
    out = []
    k = it[0]
    pts = []
    if k == 'l':
        pts = [(it[1].x, it[1].y, it[2].x, it[2].y)]
    elif k == 're':
        r = it[1]
        pts = [(r.x0, r.y0, r.x1, r.y0), (r.x1, r.y0, r.x1, r.y1), (r.x1, r.y1, r.x0, r.y1), (r.x0, r.y1, r.x0, r.y0)]
    elif k == 'qu':
        q = it[1]
        P = [q.ul, q.ur, q.lr, q.ll]
        pts = [(P[i].x, P[i].y, P[(i + 1) % 4].x, P[(i + 1) % 4].y) for i in range(4)]
    for (xa, ya, xb, yb) in pts:
        if _r2c_proche(ya, yb) and abs(xa - xb) >= 1.0:
            out.append(('h', (ya + yb) / 2, min(xa, xb), max(xa, xb)))
        elif _r2c_proche(xa, xb) and abs(ya - yb) >= 1.0:
            out.append(('v', (xa + xb) / 2, min(ya, yb), max(ya, yb)))
    return out


def _r2c_fusionner(segs):
    """fusionne les segments colinéaires qui se touchent (trou <= 1.5 pt)."""
    par = {}
    for s in segs:
        par.setdefault(s[0], []).append(s)
    res = []
    for sens, L in par.items():
        L = sorted(L, key=lambda s: (round(s[1] / _R2C_TOL), s[2]))
        cur = None
        for s in L:
            if cur and _r2c_proche(cur[1], s[1]) and s[2] <= cur[3] + 1.5:
                cur = (sens, cur[1], cur[2], max(cur[3], s[3]))
            else:
                if cur:
                    res.append(cur)
                cur = s
        if cur:
            res.append(cur)
    return res


def _r2c_cadres_d_un_trace(d, CT, courbes):
    """P1 — un seul tracé dont tous les items sont des traits axe-alignés posés sur les 4 bords de sa boîte (tirets compris)."""
    r = fitz.Rect(d['rect'])
    if r.width < _R2C_MIN_L or r.height < _R2C_MIN_H:
        return None
    if d.get('type') not in ('s', 'fs'):
        return None
    items = d['items']
    if len(items) < 1 or any(it[0] not in (('l', 're', 'qu', 'c') if courbes else ('l', 're', 'qu')) for it in items):
        return None
    cov = {'t': 0.0, 'b': 0.0, 'l': 0.0, 'r': 0.0}
    hors = 0
    for it in items:
        if it[0] == 'c':
            pts = [it[1], it[2], it[3], it[4]]
            cx0 = min(p.x for p in pts); cx1 = max(p.x for p in pts); cy0 = min(p.y for p in pts); cy1 = max(p.y for p in pts)
            if (cx1 - cx0) <= CT + 1 and (cy1 - cy0) <= CT + 1 and (cx0 <= r.x0 + CT + 1 or cx1 >= r.x1 - CT - 1) and (cy0 <= r.y0 + CT + 1 or cy1 >= r.y1 - CT - 1):
                continue
            return None
        for s in _r2c_segments_item(it):
            if s[0] == 'h':
                if _r2c_proche(s[1], r.y0, 1.0):
                    cov['t'] += s[3] - s[2]
                elif _r2c_proche(s[1], r.y1, 1.0):
                    cov['b'] += s[3] - s[2]
                else:
                    hors += 1
            else:
                if _r2c_proche(s[1], r.x0, 1.0):
                    cov['l'] += s[3] - s[2]
                elif _r2c_proche(s[1], r.x1, 1.0):
                    cov['r'] += s[3] - s[2]
                else:
                    hors += 1
    if hors:
        return None
    if cov['t'] >= 0.5 * r.width and cov['b'] >= 0.5 * r.width and cov['l'] >= 0.5 * r.height and cov['r'] >= 0.5 * r.height:
        return r
    return None


def _r2c_cadres_assembles(hs, vs, CT):
    """P2 — côtés assemblés : deux verticaux et deux horizontaux qui ferment un rectangle (coins arrondis tolérés)."""
    res = []
    for i, v1 in enumerate(vs):
        for v2 in vs[i + 1:]:
            a, b = (v1, v2) if v1[1] <= v2[1] else (v2, v1)
            if b[1] - a[1] < _R2C_MIN_L:
                continue
            if not (_r2c_proche(a[2], b[2], _R2C_TOL_CAD) and _r2c_proche(a[3], b[3], _R2C_TOL_CAD)):
                continue
            y0, y1 = a[2], a[3]
            if y1 - y0 < _R2C_MIN_H - 2 * CT or y1 - y0 < 4:
                continue
            xa, xb = a[1], b[1]
            top = [h for h in hs if (y0 - CT - 0.01 <= h[1] <= y0 + _R2C_TOL_CAD)
                   and h[2] <= xa + CT + _R2C_TOL_CAD and h[3] >= xb - CT - _R2C_TOL_CAD
                   and h[2] >= xa - _R2C_TOL_CAD and h[3] <= xb + _R2C_TOL_CAD]
            bot = [h for h in hs if (y1 - _R2C_TOL_CAD <= h[1] <= y1 + CT + 0.01)
                   and h[2] <= xa + CT + _R2C_TOL_CAD and h[3] >= xb - CT - _R2C_TOL_CAD
                   and h[2] >= xa - _R2C_TOL_CAD and h[3] <= xb + _R2C_TOL_CAD]
            if not top or not bot:
                continue
            ht = min(top, key=lambda h: abs(h[1] - y0))
            hb = min(bot, key=lambda h: abs(h[1] - y1))
            R = fitz.Rect(xa, ht[1], xb, hb[1])
            if R.width >= _R2C_MIN_L and R.height >= _R2C_MIN_H:
                res.append(R)
    return res


def _r2c_partagent_un_cote(A, B):
    """A et B (rects) partagent un côté : même ligne (1.5 pt) et recouvrement d'au moins la moitié du plus court."""
    def rec(a0, a1, b0, b1):
        return min(a1, b1) - max(a0, b0)
    for ya, yb in ((A.y0, B.y0), (A.y0, B.y1), (A.y1, B.y0), (A.y1, B.y1)):
        if _r2c_proche(ya, yb, 1.5) and rec(A.x0, A.x1, B.x0, B.x1) >= 0.5 * min(A.width, B.width):
            return True
    for xa, xb in ((A.x0, B.x0), (A.x0, B.x1), (A.x1, B.x0), (A.x1, B.x1)):
        if _r2c_proche(xa, xb, 1.5) and rec(A.y0, A.y1, B.y0, B.y1) >= 0.5 * min(A.height, B.height):
            return True
    return False


def _r2c_relever_page(page, doc_nom, CT, courbes):
    """rend (lignes, joue, raison_non_joue). lignes : une par cadre."""
    ds = page.get_drawings()
    if len(ds) > _R2C_PLAFOND_TRACES:
        return [], False, "trop de tracés (%d)" % len(ds)
    pr = page.rect
    aire_page = pr.width * pr.height
    segs = []
    p1 = []
    fonds = []   # (rect, fill) des remplissages non minces
    for d in ds:
        typ = d.get('type')
        r = fitz.Rect(d['rect'])
        cadre1 = _r2c_cadres_d_un_trace(d, CT, courbes)
        if cadre1 is not None:
            p1.append(cadre1)
        if typ in ('s', 'fs'):
            for it in d['items']:
                segs += _r2c_segments_item(it)
        if typ in ('f', 'fs') and d.get('fill') is not None:
            # un fin rectangle plein est un trait (bordure Word)
            if min(r.width, r.height) <= _R2C_EPAIS_TRAIT and max(r.width, r.height) >= 6:
                if r.width >= r.height:
                    segs.append(('h', (r.y0 + r.y1) / 2, r.x0, r.x1))
                else:
                    segs.append(('v', (r.x0 + r.x1) / 2, r.y0, r.y1))
            else:
                fonds.append((r, d['fill']))
    segs = _r2c_fusionner(segs)
    hs = [s for s in segs if s[0] == 'h']
    vs = [s for s in segs if s[0] == 'v']
    if len(hs) > 400 or len(vs) > 400:
        return [], False, "trop de côtés (%d h, %d v)" % (len(hs), len(vs))
    cadres = [(r, 'P1') for r in p1] + [(r, 'P2') for r in _r2c_cadres_assembles(hs, vs, CT)]
    # dédoublonner en gardant le premier
    uniques = []
    for r, src in cadres:
        if r.width * r.height > _R2C_MAX_AIRE * aire_page:
            continue
        if not any(_r2c_proche(r.x0, o.x0, 1.5) and _r2c_proche(r.y0, o.y0, 1.5) and _r2c_proche(r.x1, o.x1, 1.5) and _r2c_proche(r.y1, o.y1, 1.5) for o, _ in uniques):
            uniques.append((r, src))
    rects = [r for r, _ in uniques]
    # mots, images
    mots = page.get_text('words')
    try:
        imgs = [fitz.Rect(i['bbox']) for i in page.get_image_info()]
    except Exception:
        imgs = []
    lignes = []
    for idx, (R, src) in enumerate(uniques):
        raison = None
        # englobe un autre cadre / partage un côté (cellule de grille)
        englobe = any(j != idx and R.contains(rects[j]) and not (_r2c_proche(R.x0, rects[j].x0, 1.5) and _r2c_proche(R.x1, rects[j].x1, 1.5) and _r2c_proche(R.y0, rects[j].y0, 1.5) and _r2c_proche(R.y1, rects[j].y1, 1.5)) and rects[j].width * rects[j].height >= 0.02 * R.width * R.height for j in range(len(rects)))
        partage = any(j != idx and _r2c_partagent_un_cote(R, rects[j]) and not rects[j].contains(R) and not R.contains(rects[j]) for j in range(len(rects)))
        inter = fitz.Rect(R.x0 + 1.0, R.y0 + 1.0, R.x1 - 1.0, R.y1 - 1.0)
        m_dans = [w for w in mots if inter.contains(fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2))]
        n_mots = sum(1 for w in m_dans if _R2C_RE_MOT.search(w[4]))
        n_lignes = len(set((w[5], w[6]) for w in m_dans if _R2C_RE_MOT.search(w[4])))
        blancs = sum(1 for w in m_dans if _R2C_RE_BLANC.search(w[4]))
        trouve = {'image': 0, 'couleur': 0, 'forme': 0, 'trait': 0}
        # images
        for im in imgs:
            if inter.contains(fitz.Point((im.x0 + im.x1) / 2, (im.y0 + im.y1) / 2)):
                trouve['image'] += 1
        # fond du cadre : un remplissage non blanc qui couvre >= 80 % du cadre
        for (fr, fill) in fonds:
            if not _r2c_couleur_blanche(fill) and (fr & R).get_area() >= 0.8 * R.get_area() and fr.get_area() <= 1.1 * R.get_area() + 40:
                trouve['couleur'] += 1
        # tracés à l'intérieur (hors les côtés du cadre)
        for d in ds:
            r = fitz.Rect(d['rect'])
            if r.is_empty and not (r.width >= 0 and r.height >= 0):
                continue
            if not inter.contains(fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)):
                continue
            w_, h_ = r.width, r.height
            # un trait fin (cloison, ou trait de réponse) peut toucher les bords du cadre : seul son milieu doit être dedans
            fin_h = (h_ <= _R2C_EPAIS_TRAIT and w_ >= _R2C_BLANC_MIN and R.x0 - 2 <= r.x0 and r.x1 <= R.x1 + 2)
            fin_v = (w_ <= _R2C_EPAIS_TRAIT and h_ >= _R2C_BLANC_MIN and R.y0 - 2 <= r.y0 and r.y1 <= R.y1 + 2
                     and inter.x0 + 2 <= (r.x0 + r.x1) / 2 <= inter.x1 - 2)
            if not (fin_h or fin_v):
                if not (inter.x0 - 0.5 <= r.x0 and r.x1 <= inter.x1 + 0.5 and inter.y0 - 0.5 <= r.y0 and r.y1 <= inter.y1 + 0.5):
                    continue
            fill = d.get('fill')
            a_courbe = any(it[0] == 'c' for it in d['items'])
            if d.get('type') in ('f', 'fs') and fill is not None and min(w_, h_) > _R2C_SEUIL_MM:
                if _r2c_couleur_blanche(fill):
                    pass
                elif _r2c_couleur_noire(fill):
                    trouve['forme'] += 1
                else:
                    trouve['couleur'] += 1
                continue
            if a_courbe and w_ >= 3 and h_ >= 3:
                trouve['forme'] += 1
                continue
            if h_ <= _R2C_EPAIS_TRAIT and w_ >= _R2C_BLANC_MIN:
                if w_ >= 0.95 * R.width:
                    trouve['forme'] += 1      # une cloison qui traverse tout le cadre : deux cases, pas un blanc
                    continue
                # un soulignement (le trait est juste sous un mot, sur plus de 60 % de sa largeur) n'est pas un blanc
                rec_tot = 0.0
                for w in m_dans:
                    if -4.0 <= (r.y0 + r.y1) / 2 - w[3] <= 1.5:
                        rec_tot += max(0.0, min(r.x1, w[2]) - max(r.x0, w[0]))
                sous = rec_tot >= 0.6 * w_
                if not sous:
                    trouve['trait'] += 1
                continue
            if w_ <= _R2C_EPAIS_TRAIT and h_ >= _R2C_BLANC_MIN:
                trouve['forme'] += 1
                continue
            if w_ >= 3 and h_ >= 3:
                trouve['forme'] += 1
        blancs_tot = blancs + trouve['trait']
        if englobe:
            raison = 'englobe'
        elif partage:
            raison = 'grille'
        elif trouve['image']:
            raison = 'image'
        elif trouve['couleur']:
            raison = 'couleur'
        elif trouve['forme']:
            raison = 'forme'
        classe = None
        if raison is None:
            if n_mots >= _R2C_MOTS_TC and blancs_tot >= 1:
                classe = 'TC'
            elif n_mots >= _R2C_MOTS_T and n_lignes >= _R2C_LIGNES_T and blancs_tot == 0:
                classe = 'T'
            else:
                raison = 'trop court'
        lignes.append(dict(doc=doc_nom, page=page.number + 1, rect=[round(R.x0, 1), round(R.y0, 1), round(R.x1, 1), round(R.y1, 1)],
                           source=src, mots=n_mots, lignes=n_lignes, blancs=blancs_tot, classe=classe, exclusion=raison,
                           trouve=trouve))
    return lignes, True, None


def _r2c_spans_noirs(page, R):
    inter = fitz.Rect(R.x0 + 1, R.y0 + 1, R.x1 - 1, R.y1 - 1)
    vu = 0
    for b in page.get_text('dict')['blocks']:
        if b['type'] != 0:
            continue
        for l in b['lines']:
            for s in l['spans']:
                bb = fitz.Rect(s['bbox'])
                if not s['text'].strip() or not inter.contains(fitz.Point((bb.x0 + bb.x1) / 2, (bb.y0 + bb.y1) / 2)):
                    continue
                vu += 1
                c = s['color']
                if max((c >> 16) & 255, (c >> 8) & 255, c & 255) > 60:
                    return False
    return vu > 0


def _r2c_cadres_du_perimetre(page):
    lignes = []
    for ct, courbes in ((30.0, True), (10.0, False)):
        try:
            l_, ok, _ = _r2c_relever_page(page, 'x', ct, courbes)
        except Exception:
            l_ = []
        lignes += l_
    out = []
    for c in lignes:
        if c['classe'] in ('TC', 'T'):
            R = fitz.Rect(c['rect'])
            if _r2c_spans_noirs(page, R):
                out.append(fitz.Rect(R.x0 - 3, R.y0 - 3, R.x1 + 3, R.y1 + 3))
    return out


def _r2c_centre(r):
    K = _V247_K
    return fitz.Point((r[0] + r[2]) / 2 / K, (r[1] + r[3]) / 2 / K)


def _r2c_filtrer(page, out):
    cadres = _r2c_cadres_du_perimetre(page)
    if cadres:
        garde = []
        for dd in out:
            if not dd.get('orphelin'):
                garde.append(dd)
                continue
            nr = []
            for r in dd['rects']:
                c = _r2c_centre(r)
                if any(f.contains(c) for f in cadres):
                    if 'own' in dd:
                        bx0, by0, A = dd['own']
                        ya, yb = max(0, int(r[1]) - by0), max(0, int(r[3]) - by0)
                        xa, xb = max(0, int(r[0]) - bx0), max(0, int(r[2]) - bx0)
                        A = A.copy(); A[ya:yb, xa:xb] = False
                        dd = dict(dd); dd['own'] = (bx0, by0, A)
                else:
                    nr.append(r)
            if nr:
                dd = dict(dd); dd['rects'] = nr
                garde.append(dd)
        out = garde
    return out


def _v254_dessins_page(page, fenetres, etats, mots, L, texte, attribution, decor):
    """Les dessins de la page. Un dessin : les taches des fenêtres (9 px au moins) qui partagent un pixel ou sont à 3 px ou
    moins ; un tableau : les dessins et l'encre libre dont la moitié au moins est dans son rectangle ; l'encre libre (hors
    lettres, hors trait du décor) qui touche un dessin (3 px) est à lui, celle qui ne touche rien (9 px au moins) est un
    dessin ; un mot rattaché va au dessin le plus proche, un mot court non rattaché aussi, à 6 mm ou moins. Rend, par dessin :
    sa première fenêtre (`prem`), sa clé (`tag`), `orphelin`, ses rectangles (pixels de page à _V247_K, qui ne prennent pas
    l'encre des autres : _v254_couper) et son encre (`own`)."""
    # v2.60 — la boîte d'un mot de ligne ou ponctué s'étend à l'encre de ses lettres
    mots = _v260_mots_etendus(page, mots, L, texte)
    np_ = _np247; K = _V247_K
    pno = page.number + 1
    pm = page.get_pixmap(matrix=fitz.Matrix(K, K), alpha=False)
    A0 = np_.frombuffer(pm.samples, dtype=np_.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
    Ep = (255 - A0.min(axis=2)) >= _V247_ECART
    Hp, Wp = Ep.shape
    LET = np_.zeros((Hp, Wp), dtype=bool)
    M_ = _v254_lettres_page(page); h_, w_ = min(Hp, M_.shape[0]), min(Wp, M_.shape[1]); LET[:h_, :w_] = M_[:h_, :w_]
    parent = []
    def trouve(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]; i = parent[i]
        return i
    owner = np_.zeros((Hp, Wp), dtype=np_.int32)
    pix, prem = [], []
    for kk, stk in enumerate(etats):
        R = stk["R"]; ox, oy = int(np_.floor(R.x0 * K + 1e-6)), int(np_.floor(R.y0 * K + 1e-6))
        lab, labs, F = stk["lab"], stk["labs"], stk["F"]
        for lb in labs:
            m = (lab == lb) & F
            if int(m.sum()) < _V247_BRUIT:
                continue
            ys, xs = np_.nonzero(m); gy = ys + oy; gx = xs + ox
            ok = (gy >= 0) & (gy < Hp) & (gx >= 0) & (gx < Wp); gy, gx = gy[ok], gx[ok]
            if not len(gy):
                continue
            new = len(pix); parent.append(new); pix.append((gy, gx)); prem.append(kk)
            for e in np_.unique(owner[gy, gx]):
                if e > 0:
                    a, b = trouve(int(e) - 1), trouve(new)
                    if a != b:
                        parent[a] = b
            owner[gy, gx] = new + 1
    # deux taches dont des pixels sont à 3 px ou moins sont un seul dessin
    oy_, ox_ = np_.nonzero(owner)
    if len(oy_):
        o1_ = owner[oy_, ox_]
        for dy_ in range(0, 4):
            for dx_ in range(-3, 4):
                if dy_ == 0 and dx_ <= 0:
                    continue
                y2_, x2_ = oy_ + dy_, ox_ + dx_
                ok_ = (y2_ < Hp) & (x2_ >= 0) & (x2_ < Wp)
                a_ = o1_[ok_]; b_ = owner[y2_[ok_], x2_[ok_]]
                m_ = (b_ > 0) & (b_ != a_)
                if m_.any():
                    for u_, w_ in set(zip(a_[m_].tolist(), b_[m_].tolist())):
                        ra_, rb_ = trouve(int(u_) - 1), trouve(int(w_) - 1)
                        if ra_ != rb_:
                            parent[ra_] = rb_
    groupes = {}
    for i in range(len(pix)):
        groupes.setdefault(trouve(i), []).append(i)
    boites_t = np_.zeros((Hp, Wp), dtype=bool)
    for i in range(len(mots)):
        q = fitz.Rect(mots[i][:4]) & page.rect
        if q.is_empty:
            continue
        boites_t[int(q.y0 * K):int(np_.ceil(q.y1 * K)), int(q.x0 * K):int(np_.ceil(q.x1 * K))] = True
    lett = boites_t & LET
    Ddec = _v247_dilate(_v249_masque_contour([it for (rd, it) in decor], page.rect, K, Hp, Wp, Ep), 2, Ep) if decor else np_.zeros_like(Ep)
    Ddec = (_v254_decor_region(page, page.rect, (Hp, Wp), Ddec) & Ep) if decor else Ddec
    tabs = [g["clip"] for g in fenetres if g.get("tableau")]
    dessins = []
    for r_, ids in groupes.items():
        gy = np_.concatenate([pix[i][0] for i in ids]); gx = np_.concatenate([pix[i][1] for i in ids])
        dessins.append(dict(gy=gy, gx=gx, prem=min(prem[i] for i in ids), mots=[], tab=None))
    couv = np_.zeros((Hp, Wp), dtype=bool)
    for dd in dessins:
        couv[dd["gy"], dd["gx"]] = True
    libre = Ep & ~lett & ~Ddec & ~_v247_dilate(couv, 1, np_.ones_like(couv, dtype=bool))
    labl = _v247_etiqueter(_v247_dilate(libre, 1, np_.ones_like(libre, dtype=bool)))
    labl = np_.where(libre, labl, 0)
    tot = np_.bincount(labl.ravel()) if labl.max() > 0 else np_.zeros(1, dtype=np_.int64)
    pris = set()
    if tabs:
        par2 = list(range(len(dessins) + len(tabs)))
        def tr2(i):
            while par2[i] != i:
                par2[i] = par2[par2[i]]; i = par2[i]
            return i
        extra = [[] for _ in tabs]
        for ti, T in enumerate(tabs):
            q = T & page.rect
            tx0, ty0, tx1, ty1 = int(q.x0 * K), int(q.y0 * K), int(np_.ceil(q.x1 * K)), int(np_.ceil(q.y1 * K))
            for di, dd in enumerate(dessins):
                n = len(dd["gy"])
                if not n:
                    continue
                dedans = int(((dd["gy"] >= ty0) & (dd["gy"] < ty1) & (dd["gx"] >= tx0) & (dd["gx"] < tx1)).sum())
                if 2 * dedans >= n:
                    a, b = tr2(di), tr2(len(dessins) + ti)
                    if a != b:
                        par2[a] = b
            z = labl[ty0:ty1, tx0:tx1]
            if z.size:
                u, c = np_.unique(z[z > 0], return_counts=True)
                for v_, c_ in zip(u, c):
                    if 2 * int(c_) >= int(tot[v_]) and int(v_) not in pris:
                        pris.add(int(v_)); extra[ti].append(int(v_))
        neufs = {}
        for di in range(len(dessins)):
            neufs.setdefault(tr2(di), []).append(("d", di))
        for ti in range(len(tabs)):
            neufs.setdefault(tr2(len(dessins) + ti), []).append(("t", ti))
        fusion = []
        for racine, membres in neufs.items():
            ds = [dessins[x] for t_, x in membres if t_ == "d"]; ts = [x for t_, x in membres if t_ == "t"]
            if not ts:
                fusion += ds
                continue
            gys = [dd["gy"] for dd in ds]; gxs = [dd["gx"] for dd in ds]
            for ti in ts:
                for v_ in extra[ti]:
                    yy, xx = np_.nonzero(labl == v_); gys.append(yy); gxs.append(xx)
            if not gys:
                continue
            gy = np_.concatenate(gys); gx = np_.concatenate(gxs)
            if not len(gy):
                continue
            fusion.append(dict(gy=gy, gx=gx, prem=min([dd["prem"] for dd in ds] or [0]), mots=[], tab=min(ts)))
        dessins = fusion
    # l'encre libre qui touche un dessin (3 px) est à lui ; celle qui ne touche rien est un dessin à elle
    couv2 = np_.zeros((Hp, Wp), dtype=bool); own2 = np_.zeros((Hp, Wp), dtype=np_.int32)
    for di_, dd_ in enumerate(dessins):
        couv2[dd_["gy"], dd_["gx"]] = True; own2[dd_["gy"], dd_["gx"]] = di_ + 1
    libre2 = Ep & ~lett & ~Ddec & ~couv2
    if libre2.any():
        lab2 = _v247_etiqueter(_v247_dilate(libre2, 1, np_.ones_like(libre2, dtype=bool)))
        lab2 = np_.where(libre2, lab2, 0)
        px_ = own2.copy()
        for d_ in (1, 2, 3):
            px_[:, d_:] = np_.maximum(px_[:, d_:], own2[:, :-d_]); px_[:, :-d_] = np_.maximum(px_[:, :-d_], own2[:, d_:])
        prox = px_.copy()
        for d_ in (1, 2, 3):
            prox[d_:, :] = np_.maximum(prox[d_:, :], px_[:-d_, :]); prox[:-d_, :] = np_.maximum(prox[:-d_, :], px_[d_:, :])
        ys_, xs_ = np_.nonzero(lab2)
        if len(ys_):
            vv_ = lab2[ys_, xs_]; pp_ = prox[ys_, xs_]
            o_ = np_.argsort(vv_, kind="stable"); ys_, xs_, vv_, pp_ = ys_[o_], xs_[o_], vv_[o_], pp_[o_]
            cp_ = np_.flatnonzero(np_.diff(vv_)) + 1
            for gy_, gx_, pr_ in zip(np_.split(ys_, cp_), np_.split(xs_, cp_), np_.split(pp_, cp_)):
                ci_ = pr_[pr_ > 0]
                if len(ci_):
                    u_, c_ = np_.unique(ci_, return_counts=True); d2_ = dessins[int(u_[np_.argmax(c_)]) - 1]
                    d2_["gy"] = np_.concatenate([d2_["gy"], gy_]); d2_["gx"] = np_.concatenate([d2_["gx"], gx_])
                elif len(gy_) >= _V247_BRUIT:
                    dessins.append(dict(gy=gy_, gx=gx_, prem=0, mots=[], tab=None, orphelin=True))
    # v2.72 H-aL (DEC-3312, DEC-1898) : un dessin fait de la seule encre d'une étiquette de lettres se joint au dessin le plus proche à 6 mm ou moins
    _ha_port = _V247_DIST_MM / _V247_MM * K
    _ha_boites = []
    for _i in range(len(mots)):
        if not _v247_est_glyphe(_i, mots, L, texte):
            continue
        if not re.fullmatch(r"[^\w]*[A-Za-zÀ-ÿ]+[^\w]*", mots[_i][4] or ""):
            continue
        _q = fitz.Rect(mots[_i][:4]) & page.rect
        if _q.is_empty:
            continue
        _ha_boites.append((int(_q.x0 * K) - 1, int(_q.y0 * K) - 1, int(np_.ceil(_q.x1 * K)) + 1, int(np_.ceil(_q.y1 * K)) + 1))
    _ha_seuls = []
    for _di, _dd in enumerate(dessins):
        if not len(_dd["gx"]) or _dd.get("tab") is not None:
            continue
        _x0, _x1, _y0, _y1 = int(_dd["gx"].min()), int(_dd["gx"].max()), int(_dd["gy"].min()), int(_dd["gy"].max())
        if any(b[0] <= _x0 and _x1 < b[2] and b[1] <= _y0 and _y1 < b[3] for b in _ha_boites):
            _ha_seuls.append(_di)
    _ha_set = set(_ha_seuls)
    for _di in _ha_seuls:
        _dd = dessins[_di]; _best = None
        _x0, _x1, _y0, _y1 = int(_dd["gx"].min()), int(_dd["gx"].max()) + 1, int(_dd["gy"].min()), int(_dd["gy"].max()) + 1
        for _dj, _ee in enumerate(dessins):
            if _dj in _ha_set or not len(_ee["gx"]):
                continue
            gx_, gy_ = _ee["gx"], _ee["gy"]
            gxm = max(0, int(gx_.min()) - _x1, _x0 - int(gx_.max())); gym = max(0, int(gy_.min()) - _y1, _y0 - int(gy_.max()))
            if gxm * gxm + gym * gym > _ha_port * _ha_port:
                continue
            dx = np_.maximum(0, np_.maximum(_x0 - gx_, gx_ - (_x1 - 1))); dy = np_.maximum(0, np_.maximum(_y0 - gy_, gy_ - (_y1 - 1)))
            d2 = float((dx * dx + dy * dy).min())
            if d2 <= _ha_port * _ha_port and (_best is None or d2 < _best[0]):
                _best = (d2, _dj)
        if _best is not None:
            _ee = dessins[_best[1]]
            _ee["gx"] = np_.concatenate([_ee["gx"], _dd["gx"]]); _ee["gy"] = np_.concatenate([_ee["gy"], _dd["gy"]])
            _dd["gx"] = np_.array([], dtype=np_.int64); _dd["gy"] = np_.array([], dtype=np_.int64)
    # un mot rattaché va au dessin le plus proche
    for i, kk in attribution.items():
        q = fitz.Rect(mots[i][:4]) & page.rect
        if q.is_empty:
            continue
        x0, y0, x1, y1 = int(q.x0 * K), int(q.y0 * K), int(np_.ceil(q.x1 * K)), int(np_.ceil(q.y1 * K))
        g = Ep[y0:y1, x0:x1]
        if not g.any():
            continue
        ys, xs = np_.nonzero(g); wb = (x0 + int(xs.min()), y0 + int(ys.min()), x0 + int(xs.max()) + 1, y0 + int(ys.max()) + 1)
        best = None
        for di, dd in enumerate(dessins):
            gx_, gy_ = dd["gx"], dd["gy"]
            if not len(gx_):
                continue
            gxm = max(0, int(gx_.min()) - wb[2], wb[0] - int(gx_.max())); gym = max(0, int(gy_.min()) - wb[3], wb[1] - int(gy_.max()))
            if best is not None and gxm * gxm + gym * gym > best[0]:
                continue
            dx = np_.maximum(0, np_.maximum(wb[0] - gx_, gx_ - (wb[2] - 1))); dy = np_.maximum(0, np_.maximum(wb[1] - gy_, gy_ - (wb[3] - 1)))
            d2 = float((dx * dx + dy * dy).min())
            if best is None or d2 < best[0]:
                best = (d2, di)
        if best is not None:
            dessins[best[1]]["mots"].append(wb)
        else:
            dessins.append(dict(gy=np_.array([], dtype=np_.int64), gx=np_.array([], dtype=np_.int64), prem=kk, mots=[wb], tab=None, seul=True))
    # un mot court non rattaché va au dessin le plus proche, à 6 mm ou moins
    port_ = _V247_DIST_MM / _V247_MM * K
    for i in range(len(mots)):
        if i in attribution or not _v247_est_glyphe(i, mots, L, texte):
            continue
        q = fitz.Rect(mots[i][:4]) & page.rect
        if q.is_empty:
            continue
        x0, y0, x1, y1 = int(q.x0 * K), int(q.y0 * K), int(np_.ceil(q.x1 * K)), int(np_.ceil(q.y1 * K))
        g = Ep[y0:y1, x0:x1]
        if not g.any():
            continue
        ys, xs = np_.nonzero(g); wb = (x0 + int(xs.min()), y0 + int(ys.min()), x0 + int(xs.max()) + 1, y0 + int(ys.max()) + 1)
        best = None
        for di, dd in enumerate(dessins):
            gx_, gy_ = dd["gx"], dd["gy"]
            if not len(gx_):
                continue
            gxm = max(0, int(gx_.min()) - wb[2], wb[0] - int(gx_.max())); gym = max(0, int(gy_.min()) - wb[3], wb[1] - int(gy_.max()))
            if gxm * gxm + gym * gym > port_ * port_ or (best is not None and gxm * gxm + gym * gym > best[0]):
                continue
            dx = np_.maximum(0, np_.maximum(wb[0] - gx_, gx_ - (wb[2] - 1))); dy = np_.maximum(0, np_.maximum(wb[1] - gy_, gy_ - (wb[3] - 1)))
            d2 = float((dx * dx + dy * dy).min())
            if best is None or d2 < best[0]:
                best = (d2, di)
        if best is not None and best[0] <= port_ * port_:
            dessins[best[1]]["mots"].append(wb)
    out = []
    for di, dd in enumerate(dessins):
        xs_ = [int(dd["gx"].min()), int(dd["gx"].max()) + 1] if len(dd["gx"]) else []
        ys_ = [int(dd["gy"].min()), int(dd["gy"].max()) + 1] if len(dd["gy"]) else []
        for wb in dd["mots"]:
            xs_ += [wb[0], wb[2]]; ys_ += [wb[1], wb[3]]
        if not xs_:
            continue
        bx0, by0, bx1, by1 = min(xs_), min(ys_), max(xs_), max(ys_)
        bx0, by0, bx1, by1 = max(0, bx0), max(0, by0), min(Wp, bx1), min(Hp, by1)
        A = np_.zeros((by1 - by0, bx1 - bx0), dtype=bool)
        if len(dd["gy"]):
            A[dd["gy"] - by0, dd["gx"] - bx0] = True
        for wb in dd["mots"]:
            A[wb[1] - by0:wb[3] - by0, wb[0] - bx0:wb[2] - bx0] |= Ep[wb[1]:wb[3], wb[0]:wb[2]]
        ys, xs = np_.nonzero(A)
        if not len(ys):
            continue
        a0, b0, a1, b1 = int(xs.min()), int(ys.min()), int(xs.max()) + 1, int(ys.max()) + 1
        A = A[b0:b1, a0:a1]; bx0, by0, bx1, by1 = bx0 + a0, by0 + b0, bx0 + a1, by0 + b1
        Dd = _v247_dilate(A, 1, np_.ones_like(A, dtype=bool))
        O = Ep[by0:by1, bx0:bx1] & ~Dd
        if O.any():
            lab = _v247_etiqueter(O)
            if lab.max() > 0:
                taille = np_.bincount(lab.ravel()); petit = taille < _V247_BRUIT; petit[0] = False
                O = O & ~petit[lab]
        # v2.69 — F033 W12h : un dessin de tableau se coupe par W12h (voir _v269_w12h_couper)
        feuilles = _v269_w12h_couper(A, O, 10, dd.get("tab") is not None, page, bx0, by0)
        tag = ("t", pno, dd["tab"]) if dd.get("tab") is not None else ("d", pno, di)
        out.append(dict(prem=dd["prem"], tag=tag, orphelin=bool(dd.get("orphelin")),
                        rects=[(bx0 + a, by0 + b, bx0 + c, by0 + d) for (a, b, c, d), n in feuilles], own=(bx0, by0, A)))
    return out


def _v254_rectangles(page, k, st, fenetres, etats, mots, L, texte, attribution, decor):
    """Les rectangles des dessins dont la fenêtre k est la première, en pixels de sa région comptés depuis R.x0·K (position
    exacte), absorbés dessin par dessin ; la clé de chaque rectangle (son cadre en points) retient son dessin."""
    K = _V247_K
    R1 = _v254_etat()["r1"]
    if R1.get("doc") is not page.parent:
        R1.clear(); R1.update(doc=page.parent, cle={}, pages={})
    pno = page.number + 1
    if pno not in R1["pages"]:
        R1["pages"][pno] = _r2c_filtrer(page, _v254_dessins_page(page, fenetres, etats, mots, L, texte, attribution, decor))
    R = st["R"]; ox, oy = R.x0 * K, R.y0 * K
    cle = R1["cle"]; rects = []
    for dd in R1["pages"][pno]:
        if dd["prem"] != k:
            continue
        rd_ = []
        for (x0, y0, x1, y1) in dd["rects"]:
            rc = (x0 - ox, y0 - oy, x1 - ox, y1 - oy); rd_.append(rc)
            p = fitz.Rect(R.x0 + rc[0] / K, R.y0 + rc[1] / K, R.x0 + rc[2] / K, R.y0 + rc[3] / K) & page.rect
            if not p.is_empty:
                cle[(pno, tuple(round(v, 1) for v in (p.x0, p.y0, p.x1, p.y1)))] = dd["tag"]
        rects += _v247_absorber(rd_)
    return rects


def _v254_encre_portee(page, r, cadres):
    """La part de l'encre du rectangle r (écart ≥ _V247_ECART, à _V247_K) qui est sous les cadres donnés ; 1.0 sans encre."""
    np_ = _np247; K = _V247_K
    q = fitz.Rect(r) & page.rect
    if q.is_empty:
        return 1.0
    pm = page.get_pixmap(matrix=fitz.Matrix(K, K), clip=q, alpha=False)
    A = np_.frombuffer(pm.samples, dtype=np_.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
    E = (255 - A.min(axis=2)) >= _V247_ECART
    n = int(E.sum())
    if not n:
        return 1.0
    ox, oy = int(np_.floor(q.x0 * K + 1e-6)), int(np_.floor(q.y0 * K + 1e-6)); Z = np_.zeros_like(E)
    for c in cadres:
        c = fitz.Rect(c) & q
        if c.is_empty:
            continue
        Z[max(0, int(c.y0 * K) - oy):max(0, int(np_.ceil(c.y1 * K)) - oy), max(0, int(c.x0 * K) - ox):max(0, int(np_.ceil(c.x1 * K)) - ox)] = True
    return int((E & Z).sum()) / n


def _v254_absorber_par_dessin(page, rects):
    """_v247_absorber_rects dessin par dessin : les morceaux rangés par la clé de leur dessin ; un morceau sans clé est seul."""
    pno = page.number + 1
    cle = _v254_etat()["r1"].get("cle") or {}
    par = {}
    for j, r in enumerate(rects):
        t = cle.get((pno, tuple(round(v, 1) for v in (r.x0, r.y0, r.x1, r.y1))))
        par.setdefault(t if t is not None else ("seul", j), []).append(r)
    out = []
    for L_ in par.values():
        out += _v247_absorber_rects(L_)
    return out


def _v254_par_dessin(page, ims, idx_fin):
    """Une figure par dessin, par la clé de ses morceaux (un tableau est un seul dessin) ; un morceau sans clé garde sa
    fenêtre ; morceau, n_morceaux et l'image entière refaits par figure."""
    R1 = _v254_etat()["r1"]
    if R1.get("doc") is not page.parent:
        return ims, idx_fin
    pno = page.number + 1
    cle = R1["cle"]
    groupes = {}
    for im in ims:
        fg = im.get("figure") or ""
        if not (fg.startswith("p") and "-f" in fg and fg.split("-f", 1)[1].isdigit()):
            continue
        tag = cle.get((pno, tuple(round(v, 1) for v in im["cadre"])))
        if tag is None:
            tag = ("f", fg)
        groupes.setdefault(tag, []).append(im)
    for n, (tag, L_) in enumerate(sorted(groupes.items(), key=lambda kv: min(i["index"] for i in kv[1]))):
        L_.sort(key=lambda i: i["index"])
        fid = "p%d-%s%d" % (pno, "t" if tag[0] == "t" else "d", n)
        for j, im in enumerate(L_):
            im["figure"] = fid; im["morceau"] = j + 1; im["n_morceaux"] = len(L_)
            im.pop("entier", None); im.pop("entier_cadre", None)
        if len(L_) > 1:
            U = fitz.Rect(min(i["cadre"][0] for i in L_), min(i["cadre"][1] for i in L_), max(i["cadre"][2] for i in L_), max(i["cadre"][3] for i in L_)) & page.rect
            pe = page.get_pixmap(clip=U, matrix=fitz.Matrix(2, 2))
            L_[0]["entier"] = "data:image/png;base64," + base64.b64encode(pe.tobytes("png")).decode()
            L_[0]["entier_cadre"] = [U.x0, U.y0, U.x1, U.y1]
    return ims, idx_fin


def _v254_lecture_vide():
    """La marque machine de la lecture des exercices : rendue dans la réponse, jamais sur le papier (la page ne lit pas ce champ)."""
    return {"lecteur": _V254_LECTURE_MODELE, "pages": 0, "pages_lues": 0, "pages_sans_dessin": 0, "pages_en_panne": [],
            "sans_cle": False, "dessins_numerotes": 0, "jetons_entree": 0, "jetons_sortie": 0, "derniere_erreur": ""}


def _v254_dessins_numerotes(page, ims):
    """Les dessins à numéroter : toutes les figures de la page après la figure par dessin, dans l'ordre de leurs index, chacune
    par la boîte de ses morceaux (v2.59 : celles faites de dessins trouvés hors des fenêtres comprises)."""
    figs = {}
    for im in ims:
        figs.setdefault(im.get("figure"), []).append(im)
    F = []
    for f, L_ in sorted(figs.items(), key=lambda kv: min(i["index"] for i in kv[1])):
        bx = [min(i["cadre"][0] for i in L_), min(i["cadre"][1] for i in L_), max(i["cadre"][2] for i in L_), max(i["cadre"][3] for i in L_)]
        F.append(dict(figure=f, boite=[round(v, 2) for v in bx]))
    return F


def _v254_image_numerotee(page, F):
    """La page entière (1 500 px au plus sur son grand côté), chaque dessin de F entouré en orange et numéroté P1, P2…"""
    g = min(4.0, 1500.0 / max(page.rect.width, page.rect.height))
    pm = page.get_pixmap(matrix=fitz.Matrix(g, g))
    o = fitz.open(); q = o.new_page(width=pm.width, height=pm.height)
    q.insert_image(fitz.Rect(0, 0, pm.width, pm.height), pixmap=pm)
    for k, m in enumerate(F):
        c = m["boite"]; rr = fitz.Rect(c[0] * g - 1, c[1] * g - 1, c[2] * g + 1, c[3] * g + 1)
        q.draw_rect(rr, color=(1, 0.45, 0), width=1.0)
    places = []  # RL (DEC-3362, rang 3) : chaque numéro sur fond blanc, à la première place libre
    for k, m in enumerate(F):
        c = m["boite"]; rr = fitz.Rect(c[0] * g - 1, c[1] * g - 1, c[2] * g + 1, c[3] * g + 1)
        t = "P%d" % (k + 1); w = fitz.get_text_length(t, fontsize=10) + 3; h = 11
        for x, y in ((rr.x0, rr.y0 - h - 1), (rr.x1 - w, rr.y0 - h - 1), (rr.x0, rr.y1 + 1), (rr.x1 - w, rr.y1 + 1), (rr.x0 + 1, rr.y0 + 1)):
            R = fitz.Rect(x, y, x + w, y + h)
            if q.rect.contains(R) and not any(R.intersects(p) for p in places):
                break
        places.append(R)
        q.draw_rect(R, color=None, fill=(1, 1, 1))
        q.insert_text((R.x0 + 1.5, R.y1 - 2), t, fontsize=10, color=(0.85, 0.1, 0))
    png = q.get_pixmap(matrix=fitz.Matrix(1, 1)).tobytes("png")
    o.close()
    return png


def _v254_exercice(e, pno):
    """Le nom d'un exercice lu : « sans numéro » → « sn<page> » ; « aucun » (pas d'exercice) ; sinon le numéro, sans le mot
    « exercice » ni sa ponctuation. Reconnaissance : une liste de formes."""
    brut = str(e).strip().lower()
    # v2.81 — F1 (DEC-3417, DEC-3416) : « sans numéro A », « sans numéro B »… (la question les demande) → sn<page>_A, sn<page>_B… ;
    # chaque partie sans numéro garde son nom et ses dessins (DEC-2620). « sans numéro » nu reste sn<page>.
    _f1 = re.match(r"^(?:sans num[ée]ro|non num[ée]rot[ée])\s+([a-z])$", brut)
    if _f1:
        return "sn%d_%s" % (pno, _f1.group(1).upper())
    if brut in _V254_SANS_NUMERO:
        return "sn%d" % pno
    e = str(e).lower()
    for w in ("exercice", "partie", "parte", "n°", "ex."):
        e = e.replace(w, "")
    e = e.strip(" .°)-:")
    return "aucun" if e in ("aucun", "none", "", "sans numéro", "sans numero") else e


def _v254_lire_exercices(api_key, png, n):
    """UN appel : les exercices de la page et l'exercice de chacun des n dessins numérotés. Rend ((exercices, {n°: exercice})
    | None, faute | None, jetons_in, jetons_out). La reprise est celle de la bibliothèque ; une panne est rendue, jamais levée."""
    try:
        client = anthropic.Anthropic(api_key=api_key, timeout=_V254_LECTURE_ATTENTE)
        rep = client.messages.create(
            model=_V254_LECTURE_MODELE, max_tokens=_V254_LECTURE_BORNE,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                             "data": base64.b64encode(png).decode()}},
                {"type": "text", "text": _V254_DEMANDE + " · ".join("P%d" % (k + 1) for k in range(n))}]}])
    except Exception as e:
        return None, "panne d'appel : %s" % type(e).__name__, 0, 0
    j_in = getattr(rep.usage, "input_tokens", 0) or 0
    j_out = getattr(rep.usage, "output_tokens", 0) or 0
    if rep.stop_reason == "max_tokens":
        return None, "réponse coupée", j_in, j_out
    texte = "".join(b.text for b in rep.content if getattr(b, "type", "") == "text")
    try:
        r = json.loads(texte[texte.index("{"):texte.rindex("}") + 1])
        exercices = list(r["exercices"])
        for e in exercices:
            str(e.get("numero", "")).strip()
        mo = {int(x["n"]): str(x["exercice"]).strip() for x in r["morceaux"]}
    except Exception as e:
        return None, "réponse illisible : %s" % type(e).__name__, j_in, j_out
    return (exercices, mo), None, j_in, j_out


def _v254_lire_la_page(page, ims, api_key, compte):
    """La lecture des exercices de la page, rangée dans l'état du document : l'exercice de chaque dessin numéroté (par sa
    boîte) et les exercices de la page avec le début de leur consigne. Sans clé, en panne, coupée ou illisible : la page n'a
    pas de lecture, et c'est dit dans `compte`."""
    S = _v254_etat()["lecture"]
    if S.get("doc") is not page.parent:
        S.clear(); S.update(doc=page.parent, fig={}, consignes={}, exercices={}, dernier=None, en_panne=set())
    pno = page.number + 1
    compte["pages"] += 1
    F = _v254_dessins_numerotes(page, ims)
    if not F:
        compte["pages_sans_dessin"] += 1
        return
    if not api_key:
        compte["sans_cle"] = True
        return
    lu, faute, j_in, j_out = _v254_lire_exercices(api_key, _v254_image_numerotee(page, F), len(F))
    compte["jetons_entree"] += j_in
    compte["jetons_sortie"] += j_out
    if lu is None:
        compte["pages_en_panne"].append([pno, faute])
        compte["derniere_erreur"] = faute or ""
        S["en_panne"].add(pno)
        return
    exercices, mo = lu
    compte["pages_lues"] += 1
    compte["dessins_numerotes"] += len(F)
    for k, m in enumerate(F):
        if (k + 1) in mo:
            S["fig"][(pno, tuple(m["boite"]))] = _v254_exercice(mo[k + 1], pno)
    S["consignes"][pno] = [(_v254_exercice(e.get("numero"), pno), e.get("consigne") or "") for e in exercices]
    S["exercices"][pno] = exercices


# ─── v2.59 — LE NUMÉRO ÉCRIT D'UN EXERCICE LU ET LA PLACE DE SA CONSIGNE : le numéro écrit, le début de ligne, la recherche, les
# dessins lus (voir JOURNAL BACKEND v2.59) ───
import unicodedata as _ud259
_V259_BANDE, _V259_ECART = 9.0, 45.0     # le numéro et sa consigne : même bande (± 9 pt), la consigne à 45 pt au plus à droite
_V259_PETIT = 36.0                       # F-c : un dessin d'au plus 36 pt de côté


def _v259_mots(t):
    return re.findall(r"[\w’']+", _ud259.normalize("NFC", t or "").strip().lower())


def _v259_norme(t):
    t = _ud259.normalize("NFC", t or "").lower().replace("’", "'").replace("‘", "'")
    t = re.sub(r"\s+([.,;:!?])", r"\1", t); t = re.sub(r"([<>=≠])\s*,", r"\1,", t)
    return re.sub(r"\s+", " ", t).strip()


def _v259_lignes(page):
    """Les lignes de texte de la page : (texte, boîte, rang du bloc)."""
    out = []
    for bi, b in enumerate(page.get_text("dict")["blocks"]):
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            t = "".join(sp["text"] for sp in l["spans"]).strip()
            if t:
                out.append((t, l["bbox"], bi))
    return out


def _v259_reperes(L, exercices, pno):
    """Le numéro écrit de chaque exercice lu : une ligne égale au numéro lu, suivie à droite, sur la même bande, d'une ligne dont le
    premier mot est celui de la consigne lue (celle qui partage le plus de premiers mots) ; rend {exercice: (boîte du numéro,
    boîte de la ligne de consigne)}."""
    rep = {}
    for e in exercices:
        brut = str(e.get("numero", "")).strip(); nom = _v254_exercice(brut, pno)
        if nom == "aucun" or nom.startswith("sn"):
            continue
        chiffre = brut.strip(" .°)-:"); m1 = _v259_mots(e.get("consigne"))
        if not chiffre or not m1:
            continue
        cands = []
        for t, bb, _bi in L:
            if t.strip(" .°)-:") != chiffre:
                continue
            cy = (bb[1] + bb[3]) / 2
            for t2, b2, _bi2 in L:
                if b2 is bb:
                    continue
                if ((-2 <= b2[0] - bb[2] <= _V259_ECART) or (b2[0] <= bb[0] + 0.5 and b2[2] >= bb[2] and b2[0] < bb[2])) and abs((b2[1] + b2[3]) / 2 - cy) <= _V259_BANDE:
                    m2 = _v259_mots(t2)
                    if m2 and m2[0] == m1[0]:
                        k = 0
                        while k < min(len(m1), len(m2)) and m1[k] == m2[k]:
                            k += 1
                        cands.append((k, bb, b2))
        if cands:
            cands.sort(key=lambda c: -c[0]); rep[nom] = (fitz.Rect(cands[0][1]), fitz.Rect(cands[0][2]))
    return rep


def _v259_par_debut(L, C, cle):
    """La place par le début de ligne : la ligne (ou deux lignes consécutives d'un même bloc, jointes) qui commence par les trois
    premiers mots de la consigne lue, la première dans l'ordre de lecture parmi celles qui ne sont pas déjà prises."""
    cand = [(_v259_norme(t), fitz.Rect(bb)) for t, bb, bi in L]
    for i in range(len(L) - 1):
        if L[i][2] == L[i + 1][2]:
            cand.append((_v259_norme(L[i][0] + " " + L[i + 1][0]), fitz.Rect(L[i][1])))
    out, pris = {}, set()
    for ex, texte in C:
        mc = _v259_norme(texte).split()
        if not mc:
            continue
        tete = " ".join(mc[:min(3, len(mc))])
        cs = [(cle(r.y0, r.x0), i) for i, (n, r) in enumerate(cand) if n.startswith(tete) and i not in pris
              and not any(abs(r.y0 - cand[j][1].y0) < 0.5 and abs(r.x0 - cand[j][1].x0) < 0.5 for j in pris)]
        if cs:
            _, i = min(cs); pris.add(i); out[ex] = cand[i][1]
    return out


def _v259_par_recherche(page, C, cle):
    """La place par la recherche du texte de la consigne (6, 4 puis 3 premiers mots), la première libre dans l'ordre de lecture."""
    out, pris = {}, []
    for ex, texte_c in C:
        mots_c = (texte_c or "").split(); lieu = None
        for n in (6, 4, 3):
            if len(mots_c) < n and n != 3:
                continue
            q = " ".join(mots_c[:n])
            if not q:
                continue
            hits = page.search_for(q)
            libres = sorted([r for r in hits if not any(abs(r.y0 - s_.y0) < 0.5 and abs(r.x0 - s_.x0) < 0.5 for s_ in pris)], key=lambda r: cle(r.y0, r.x0))
            if libres:
                lieu = libres[0]; break
        if lieu is not None:
            pris.append(lieu); out[ex] = lieu
    return out


def _v259_par_dessins(Lfig, pno):
    """La place par les dessins lus : le coin haut-gauche du premier dessin lu de l'exercice."""
    tops = {}
    for (pp, box), e in Lfig.items():
        if pp != pno or e in (None, "aucun"):
            continue
        r = fitz.Rect(box)
        if e not in tops or (r.y0, r.x0) < (tops[e].y0, tops[e].x0):
            tops[e] = r
    return {e: fitz.Rect(r.x0, r.y0 - 0.01, r.x0 + 1, r.y0) for e, r in tops.items()}


def _v259_places_consignes(page, L, reps, C, cle, Lfig, pno, sources=False):
    """La place de la consigne de chaque exercice lu : par le numéro écrit (la ligne de consigne de `reps`), sinon le début de
    ligne, sinon la recherche, sinon les dessins lus. sources=True (v2.69, D5) : rend aussi la source de chaque place
    (« numero », « debut », « recherche », « dessins »)."""
    lieux = {e: b2 for e, (bb, b2) in reps.items()}
    src = {e: "numero" for e in lieux}
    for nom, part in (("debut", _v259_par_debut(L, C, cle)), ("recherche", _v259_par_recherche(page, C, cle)), ("dessins", _v259_par_dessins(Lfig, pno))):
        for e, r in part.items():
            if e not in lieux:
                lieux[e] = r
                src[e] = nom
    noms = {x for x, _ in C}
    out = {e: r for e, r in lieux.items() if e in noms}
    return (out, {e: src[e] for e in out}) if sources else out


def _v261_dans_grille_de_notation(page, U):
    """La boîte est à _V260_DANS au moins dans une grille du module des grilles plus large que haute et à plus de _V261_BORD pt des
    bords gauche et droit de la page. Reconnaissance par la forme et la place : provisoire."""
    C = _v260_cache(page, "grilles")
    if page.number not in C:
        try:
            C[page.number] = [fitz.Rect(g["rect"]) for g in (_grilles.analyser_page(page).get("grilles") or [])] if _grilles else []
        except Exception:
            C[page.number] = []
    A = max(1e-6, U.get_area())
    for G in C[page.number]:
        if G.width > G.height and G.x0 > _V261_BORD and page.rect.width - G.x1 > _V261_BORD and (U & G).get_area() / A >= _V260_DANS:
            return True
    return False


def _v261_entete_repete(S, page, pno, U):
    """Une page précédente a noté une boîte d'en-tête de même largeur et hauteur à 1 pt et de même coin haut-gauche à 3 pt, et la
    boîte n'est pas dans une grille de notation."""
    if not any(p_ < pno and abs(V.width - U.width) <= 1 and abs(V.height - U.height) <= 1 and abs(V.x0 - U.x0) <= 3 and abs(V.y0 - U.y0) <= 3
               for p_, V in S.get("v261_entetes", [])):
        return False
    return not _v261_dans_grille_de_notation(page, U)


# ═══ v2.68 — F014 b5 (DEC-2557 ; DEC-3269 point 2 : le défaut D2 du chantier 6, la 081 page 1) ═══
# Portage de la variante « b5 » de c252/ch6/solutions/fabrique_variantes_ch6_252.py (_ch6_desaccord, appelée avec loin = "strict").
def _v268_desaccord(U, ex, pos, lieux, page, loin=False):
    """Faux si le dessin touche le rectangle de l'exercice lu (de sa place à la place suivante dans l'ordre de lecture ; le dernier jusqu'au bas de
    la page) ou si cet exercice n'est pas placé ; sinon l'exercice dont le rectangle recouvre le plus la hauteur de U (None si aucun)."""
    L = sorted(pos)
    ouv = {}
    for i, (k, e) in enumerate(L):
        if e not in lieux:
            continue
        y0 = lieux[e].y0
        y1 = lieux[L[i + 1][1]].y0 if (i + 1 < len(L) and L[i + 1][1] in lieux) else page.rect.y1
        ouv[e] = (y0, max(y0, y1))
    if ex not in ouv:
        return False
    def ov(e):
        a, b = ouv[e]
        return max(0.0, min(U.y1, b) - max(U.y0, a))
    if ov(ex) > 0.5:
        return False
    best = max(ouv, key=ov)
    if loin and ov(best) > 0.5:
        noms = [e for k, e in L if e in ouv]
        if abs(noms.index(ex) - noms.index(best)) < 2:
            return False          # un exercice voisin : la lecture est crue (la figure peut être au-dessus de sa consigne)
    if loin == "strict" and ov(best) <= 0.5:
        return False              # la figure n'est dans le rectangle d'aucun exercice : la lecture est crue
    return best if ov(best) > 0.5 else None


# ═══ v2.69 — F051 ZA-p (DEC-2857, DEC-1904 « ce qui est à l'intérieur lui appartient » ; DEC-3271 E8) ═══
# Portage de agents274/A48_z7/a48_formes.py (za_p, mode 'u') et du crochet de a48_harnais.py. Écart déclaré, sa réponse E8 : un contenant que la
# lecture n'a pas rangé (« aucun » ou rien) reste tel quel.
def _v269_za_p(figs, aff, groupes, d5=None):
    """Du plus petit au plus grand : un dessin que la lecture a rangé, dont tous les dessins rangés qu'il contient (boîte dedans à 0,5 pt, aire
    plus petite) sont d'un même exercice, prend cet exercice. Puis, d5 donné (v2.69, F069), D5 sur ces exercices. Rend les groupes refaits
    dans l'ordre des dessins."""
    aire = lambda U: max(0.0, U[2] - U[0]) * max(0.0, U[3] - U[1])
    dans = lambda a, b: a[0] >= b[0] - 0.5 and a[1] >= b[1] - 0.5 and a[2] <= b[2] + 0.5 and a[3] <= b[3] + 0.5
    ex = [a["ex"] for a in aff]
    for i in sorted(range(len(aff)), key=lambda i: (aire(aff[i]["U"]), i)):
        if ex[i] is None or aff[i]["lu"] is None:
            continue
        ai = aire(aff[i]["U"])
        placed = [ex[j] for j in range(len(aff)) if j != i and aire(aff[j]["U"]) < ai and dans(aff[j]["U"], aff[i]["U"]) and ex[j] is not None]
        if placed and len(set(placed)) == 1 and ex[i] != placed[0]:
            ex[i] = placed[0]
    if d5 is not None:
        pno, lieux, sources, S = d5
        clemap = _v254_etat()["r1"].get("cle") or {}
        for a in aff:
            t = clemap.get((pno, tuple(round(v, 1) for v in figs[a["fg"]][0]["cadre"])))
            a["tag"] = list(t) if t is not None else None
        try:
            ex = _v269_d5(aff, ex, lieux, sources)
        except RuntimeError as e:
            S.setdefault("v269_d5_pannes", []).append([pno, str(e)])
    ng = {}
    for i, a in enumerate(aff):
        if ex[i] is None or ex[i] == "aucun":
            continue
        ng.setdefault(ex[i], []).append(figs[a["fg"]])
    return ng


# ═══ v2.69 — F069 D5 (DEC-2931 « d5 seul ») : Kx, Cn, Gt2-Au, posés derrière ZA-p dans le même crochet ═══
# Portage de agents276/A88_C2/a82_formes.py (forme('D5')) et de son socle : a80_lib_base.prepare (ord, cv, cand), a80_formes3.conteneurs_t,
# a80_formes (_union, dans_t, inv), a59_formes (aire, inter, meme_tag, couvre, ordre, pts_reels), a59_formes2 (rects_reels). Le pourquoi et les
# mesures : JOURNAL BACKEND v2.69.
_V269_TOL = 0.5
_V269_SOURCES_TEXTE = ("numero", "debut", "recherche")   # une place « dessins » n'est pas une consigne trouvée dans le texte


def _v269_aire(U):
    return max(0.0, U[2] - U[0]) * max(0.0, U[3] - U[1])


def _v269_inter(a, b):
    return max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(0.0, min(a[3], b[3]) - max(a[1], b[1]))


def _v269_meme_tag(a, b):
    return a["tag"] is not None and b["tag"] is not None and tuple(a["tag"]) == tuple(b["tag"])


def _v269_couvre(U, pts):
    """Les exercices dont la consigne (son coin haut-gauche) tombe dans la boîte U (0,5 pt)."""
    return {e for e, (x, y) in pts.items() if U[0] - _V269_TOL <= x <= U[2] + _V269_TOL and U[1] - _V269_TOL <= y <= U[3] + _V269_TOL}


def _v269_dans_t(a, b, tol):
    return a[0] >= b[0] - tol and a[1] >= b[1] - tol and a[2] <= b[2] + tol and a[3] <= b[3] + tol


def _v269_prepare(aff, pts, rects):
    """La géométrie d'une page : l'ordre des dessins (du plus grand au plus petit), les consignes que chaque boîte couvre, les candidats (plus grands, d'autre tag)."""
    P = dict(aff=aff, pts=pts, rects=rects)
    P["ord"] = sorted(range(len(aff)), key=lambda i: (-_v269_aire(aff[i]["U"]), i))
    P["cv"] = [_v269_couvre(a["U"], pts) for a in aff]
    cand = []
    for i in range(len(aff)):
        B = aff[i]["U"]; ab = _v269_aire(B); L = []
        for j in range(len(aff)):
            if j == i or _v269_meme_tag(aff[i], aff[j]) or _v269_aire(aff[j]["U"]) <= ab:
                continue
            L.append((j, _v269_inter(B, aff[j]["U"]), ab))
        cand.append(L)
    P["cand"] = cand
    return P


def _v269_conteneurs_t(P, tol):
    """Un dessin seul, et l'union des morceaux d'un tableau (tag « t », au moins deux morceaux) ; un dessin est un contenu direct d'un tableau s'il est
    dans l'union ET recouvre au moins un morceau."""
    cache = P.setdefault("_contt", {})
    if tol in cache:
        return cache[tol]
    aff = P["aff"]
    C = [dict(kind="d", box=a["U"], membres=[j], directs=[], cle=(_v269_aire(a["U"]), 0, j)) for j, a in enumerate(aff)]
    tabs = {}
    for j, a in enumerate(aff):
        if a["tag"] is not None and a["tag"][0] == "t":
            tabs.setdefault(tuple(a["tag"]), []).append(j)
    for tg, ms in sorted(tabs.items()):
        if len(ms) >= 2:
            bs = [aff[m]["U"] for m in ms]
            ub = (min(b[0] for b in bs), min(b[1] for b in bs), max(b[2] for b in bs), max(b[3] for b in bs))
            C.append(dict(kind="t", box=ub, membres=ms, directs=[], cle=(_v269_aire(ub), 1, ms[0])))
    for i, a in enumerate(aff):
        best = None
        for c in C:
            if i in c["membres"] or any(_v269_meme_tag(a, aff[m]) for m in c["membres"]):
                continue
            if not (_v269_aire(c["box"]) > _v269_aire(a["U"]) and _v269_dans_t(a["U"], c["box"], tol)):
                continue
            if c["kind"] == "t" and not any(_v269_inter(a["U"], aff[m]["U"]) > 0 for m in c["membres"]):
                continue
            if best is None or c["cle"] < best["cle"]:
                best = c
        if best is not None:
            best["directs"].append(i)
    cache[tol] = C
    return C


def _v269_en_litige(P, ex, j):
    """Un dessin rangé est en litige s'il recouvre un dessin plus grand, d'autre tag, rangé dans un autre exercice, ou le rectangle de la consigne d'un autre exercice."""
    aff = P["aff"]; Uj = aff[j]["U"]; aj = _v269_aire(Uj)
    if ex[j] is None:
        return False
    for h in range(len(aff)):
        if h == j or ex[h] is None or ex[h] == ex[j] or _v269_meme_tag(aff[j], aff[h]):
            continue
        if _v269_aire(aff[h]["U"]) > aj and _v269_inter(Uj, aff[h]["U"]) > 0:
            return True
    for e, r in P["rects"].items():
        if e != ex[j] and _v269_aire(r) > 0 and _v269_inter(r, Uj) > 0:
            return True
    return False


def _v269_kx(P, ex):
    """Kx : un dessin rangé par la lecture dont la boîte recouvre le rectangle de la consigne d'exactement un exercice, et contient son bord gauche et sa
    hauteur, prend cet exercice."""
    ex = list(ex); aff = P["aff"]
    for i in P["ord"]:
        if ex[i] is None or aff[i]["lu"] is None:
            continue
        U = aff[i]["U"]
        T = {e for e, r in P["rects"].items() if _v269_aire(r) > 0 and _v269_inter(r, U) > 0}
        if len(T) == 1:
            e = next(iter(T)); r = P["rects"][e]
            if (U[0] - _V269_TOL <= r[0] <= U[2] + _V269_TOL) and r[1] >= U[1] - _V269_TOL and r[3] <= U[3] + _V269_TOL and e != ex[i]:
                ex[i] = e
    return ex


def _v269_cn(P, ex):
    """Cn : un dessin rangé dont le centre est dans la boîte de dessins plus grands, d'autre tag, rangés, qui ne couvrent la consigne d'aucun autre exercice,
    non en litige, tous d'un même autre exercice, prend cet exercice."""
    ex = list(ex); aff = P["aff"]
    for i in P["ord"]:
        if ex[i] is None:
            continue
        U = aff[i]["U"]; p = ((U[0] + U[2]) / 2.0, (U[1] + U[3]) / 2.0); V = []
        for j, x, ab in P["cand"][i]:
            B = aff[j]["U"]
            if ex[j] is None or not (B[0] - _V269_TOL <= p[0] <= B[2] + _V269_TOL and B[1] - _V269_TOL <= p[1] <= B[3] + _V269_TOL):
                continue
            if not (P["cv"][j] <= {ex[j]}):
                continue
            if _v269_en_litige(P, ex, j):
                continue
            V.append(ex[j])
        if V and len(set(V)) == 1 and V[0] != ex[i]:
            ex[i] = V[0]
    return ex


def _v269_gt2au(P, ex, tol=_V269_TOL):
    """Gt2-Au : pour un contenant et ses contenus directs, à la majorité stricte l'exercice du contenant s'impose ; à égalité le contenant décide, sauf
    s'il est en litige ; un tableau : égalité = rien."""
    ex = list(ex)
    for c in sorted(_v269_conteneurs_t(P, tol), key=lambda c: (-_v269_aire(c["box"]), c["kind"], c["membres"][0])):
        if not c["directs"]:
            continue
        ms = [ex[m] for m in c["membres"] if ex[m] is not None]
        if not ms:
            continue
        if c["kind"] == "t" and len(set(ms)) != 1:
            continue
        ec = ms[0]
        if _v269_couvre(c["box"], P["pts"]) - {ec}:
            continue
        V = ms + [ex[k] for k in c["directs"] if ex[k] is not None]
        cnt = _co266.Counter(V); strict = cnt[ec] * 2 > len(V); egal = (not strict) and cnt[ec] == max(cnt.values())
        if egal:
            if c["kind"] == "t":
                continue
            if _v269_en_litige(P, ex, c["membres"][0]):
                continue
        elif not strict:
            continue
        for k in c["directs"]:
            if ex[k] is not None and ex[k] != ec:
                ex[k] = ec
    return ex


def _v269_d5(aff, ex, lieux, sources):
    """D5 = Kx, puis Cn, puis Gt2-Au, sur les exercices des dessins (après ZA-p) ; les consignes : leurs places trouvées dans le texte seulement."""
    cons = {e: [r.x0, r.y0, r.x1, r.y1] for e, r in lieux.items() if sources.get(e) in _V269_SOURCES_TEXTE}
    P = _v269_prepare(aff, {e: (r[0], r[1]) for e, r in cons.items()}, {e: tuple(r) for e, r in cons.items()})
    cur = list(ex)
    for f in (_v269_kx, _v269_cn, _v269_gt2au):
        nxt = f(P, cur)
        if any((x is None) != (y is None) for x, y in zip(cur, nxt)):
            raise RuntimeError("D5 : un dessin rangé deviendrait sans exercice (%s)" % f.__name__)
        cur = nxt
    return cur


def _v269_couper_aux_consignes(page, ims, lieux, Lfig):
    """v2.69 — F085 R2009 (DEC-3158) : une figure non rangée (p<n>-d…, p<n>-t…) dont la boîte contient strictement le haut d'une consigne lue
    est coupée en deux, à ce haut : les morceaux dont le centre est au-dessus restent, les autres forment la figure « <nom>s » ; morceau,
    n_morceaux et entier refaits pour les deux ; répété tant qu'une figure en enjambe encore une. Sauf si un dessin lu la recouvre à 0,8 au
    moins avec un exercice autre que « aucun » (la porte de la 2009). Portage, sans une règle changée, de la fonction du banc
    (A200_LA R2004 ; c288/A208_RD/serveur_R2009/app_v2_62.py l. 3014-3061). Le pourquoi et les mesures : JOURNAL BACKEND v2.69."""
    pno = page.number + 1
    for _ in range(6):
        figs = {}
        for im in ims:
            fg = im.get("figure") or ""
            if fg.startswith("p%d-d" % pno) or fg.startswith("p%d-t" % pno):
                figs.setdefault(fg, []).append(im)
        fait = False
        for fg, Lf in sorted(figs.items()):
            if len(Lf) < 2:
                continue
            U = fitz.Rect(min(i["cadre"][0] for i in Lf), min(i["cadre"][1] for i in Lf), max(i["cadre"][2] for i in Lf), max(i["cadre"][3] for i in Lf))
            ys = sorted({round(l.y0, 1) for l in lieux.values() if U.x0 - 2 <= l.x0 <= U.x1 and U.y0 < l.y0 - 2 and U.y1 > l.y0 + 4})
            if not ys:
                continue
            def _rec(a, b):
                i = a & b
                if i.is_empty or min(a.get_area(), b.get_area()) <= 0:
                    return 0.0
                return i.get_area() / max(1e-6, min(a.get_area(), b.get_area()))
            _cs = [(_rec(fitz.Rect(c), U) * min(1.0, fitz.Rect(c).get_area() / max(1e-6, U.get_area()), U.get_area() / max(1e-6, fitz.Rect(c).get_area())), e) for (pp, c), e in (Lfig or {}).items() if pp == pno]
            if _cs:
                _sc, _e = max(_cs, key=lambda x: x[0])
                if _sc >= 0.8 and _e != "aucun":
                    continue
            y = ys[0]
            haut = [i for i in Lf if (i["cadre"][1] + i["cadre"][3]) / 2 < y]; bas = [i for i in Lf if (i["cadre"][1] + i["cadre"][3]) / 2 >= y]
            if not haut or not bas:
                continue
            for i in bas:
                i["figure"] = fg + "s"
            for G in (haut, bas):
                G.sort(key=lambda i: i["index"])
                for j, im in enumerate(G):
                    im["morceau"] = j + 1; im["n_morceaux"] = len(G); im.pop("entier", None); im.pop("entier_cadre", None)
                if len(G) > 1:
                    U2 = fitz.Rect(min(i["cadre"][0] for i in G), min(i["cadre"][1] for i in G), max(i["cadre"][2] for i in G), max(i["cadre"][3] for i in G)) & page.rect
                    pe = page.get_pixmap(clip=U2, matrix=fitz.Matrix(2, 2))
                    G[0]["entier"] = "data:image/png;base64," + base64.b64encode(pe.tobytes("png")).decode()
                    G[0]["entier_cadre"] = [U2.x0, U2.y0, U2.x1, U2.y1]
            fait = True
            break
        if not fait:
            break
    return ims


def _v254_par_exercice(page, ims, idx_fin):
    """Après la figure par dessin : les dessins d'un même exercice forment une figure (chaque dessin garde ses morceaux).
    L'exercice d'un dessin : celui du dessin lu dont la boîte le recouvre à 0,8 au moins ; « aucun » ou rien : l'exercice dont
    la consigne le précède (sa place : _v259_places_consignes), d'une page à l'autre : un dessin qui précède toute consigne de sa
    page va à l'exercice de la page d'avant, nommé « r<n> » quand un exercice <n> est placé sur la page (v2.62) ; avant la
    première consigne du document : aucun.
    Un dessin d'au plus _V259_PETIT pt de côté dont la boîte contient le centre d'exactement un numéro écrit (_v259_reperes) n'est
    rangé dans aucun exercice et ses images sont retirées ; l'index des images qui le suivent diminue d'autant (v2.59). Un dessin qui
    précède toute consigne d'une page qui en a est noté ; un dessin reporté d'une page d'avant, sur une page qui a une consigne, dont la
    boîte répète une boîte notée d'une page précédente, est retiré de même, sauf dans une grille de notation (v2.61). Sans lecture dans le document jusqu'ici :
    rien ne change. Une page dont la lecture est en panne garde une figure par dessin (v2.55)."""
    S = _v254_etat()["lecture"]
    Lfig = S.get("fig") or {}
    if S.get("doc") is not page.parent or not (Lfig or S.get("consignes")):
        return ims, idx_fin
    pno = page.number + 1
    if pno in S.get("en_panne", ()):
        return ims, idx_fin
    cle = getattr(_grilles, "cle_lecture", None) if _grilles else None
    C = (S.get("consignes") or {}).get(pno) or []
    L_ = _v259_lignes(page)
    reps = _v260_reperes(page, L_, (S.get("exercices") or {}).get(pno) or [], pno)
    # v2.66 — F015 B : les consignes lues de la page, gardées pour voisins_b (la ligne qui recoupe une consigne lue n'est pas listée)
    S.setdefault("vb_reps", {})[pno] = {e: {"numero": [round(v, 2) for v in bb], "ligne": [round(v, 2) for v in b2]} for e, (bb, b2) in reps.items()}
    pos = []
    _v268_lieux = {}
    _v269_sources = {}
    if cle is not None:
        _v268_lieux, _v269_sources = _v259_places_consignes(page, L_, reps, C, cle, Lfig, pno, sources=True)
        for ex, lieu in _v268_lieux.items():
            pos.append((cle(lieu.y0, lieu.x0), ex))
    pos.sort()
    # v2.69 — F085 R2009 : une figure non rangée qui enjambe le haut d'une consigne lue est coupée (voir _v269_couper_aux_consignes)
    ims = _v269_couper_aux_consignes(page, ims, _v268_lieux, Lfig)
    nums = [bb for bb, b2 in reps.values()]
    retires = set()
    avant_page = S.get("dernier")
    figs = {}
    for im in ims:
        fg = im.get("figure") or ""
        if fg.startswith("p%d-d" % pno) or fg.startswith("p%d-t" % pno):
            figs.setdefault(fg, []).append(im)
    def recouvre(a, b):
        i = a & b
        if i.is_empty:
            return 0.0
        return i.get_area() / max(1e-6, min(a.get_area(), b.get_area())) if min(a.get_area(), b.get_area()) > 0 else 0.0
    groupes = {}
    _v269_aff = []
    for fg, Lf in figs.items():
        U = fitz.Rect(min(i["cadre"][0] for i in Lf), min(i["cadre"][1] for i in Lf), max(i["cadre"][2] for i in Lf), max(i["cadre"][3] for i in Lf))
        if (U.x1 - U.x0) <= _V259_PETIT and (U.y1 - U.y0) <= _V259_PETIT and sum(
                1 for n in nums if U.contains(fitz.Point((n.x0 + n.x1) / 2, (n.y0 + n.y1) / 2))) == 1:
            retires.add(fg)
            # v2.60 — la boîte de ce que F-c retire : la zone du numéro, pour la sortie de la page
            S.setdefault("v260_fc", {}).setdefault(pno, []).append([U.x0, U.y0, U.x1, U.y1])
            continue
        if cle is not None and pos and not [1 for k_, e in pos if k_ < cle(U.y0, U.x0)]:
            S.setdefault("v261_entetes", []).append((pno, fitz.Rect(U)))
        ex = None; _v268_src = None
        cands = [(recouvre(fitz.Rect(c), U) * min(1.0, fitz.Rect(c).get_area() / max(1e-6, U.get_area()), U.get_area() / max(1e-6, fitz.Rect(c).get_area())), e) for (pp, c), e in Lfig.items() if pp == pno]
        if cands:
            sc, e = max(cands, key=lambda x: x[0])
            if sc >= 0.8:
                ex = e; _v268_src = "lecture"
        if ex == "aucun":
            ex = None; _v268_src = None
        _v269_lu = ex
        if ex is None and cle is not None:
            km = cle(U.y0, U.x0); av = [e for k_, e in pos if k_ < km]
            if av:
                ex = av[-1]
            elif avant_page is not None:
                if pos and _v261_entete_repete(S, page, pno, U):
                    retires.add(fg)
                    continue
                ex = ("r" + str(avant_page)) if avant_page in {e for k_, e in pos} else avant_page
        # v2.68 — F014 b5 (DEC-2557) : un dessin rangé par la lecture hors du rectangle de son exercice va à l'exercice, à deux exercices
        # ou plus, dont le rectangle le recouvre le plus ; un exercice voisin, ou aucun rectangle : la lecture est crue
        if _v268_src == "lecture" and ex not in (None, "aucun") and cle is not None:
            e2 = _v268_desaccord(U, ex, pos, _v268_lieux, page, "strict")
            if e2 is not False:
                ex = e2 if e2 else None
        _v269_aff.append(dict(fg=fg, U=(U.x0, U.y0, U.x1, U.y1), ex=(None if ex == "aucun" else ex), lu=_v269_lu))
        if ex is None or ex == "aucun":
            continue
        groupes.setdefault(ex, []).append(Lf)
    # v2.69 — F051 ZA-p (ce qui est à l'intérieur lui appartient), puis F069 D5 (voir _v269_za_p, _v269_d5)
    groupes = _v269_za_p(figs, _v269_aff, groupes, d5=(pno, _v268_lieux, _v269_sources, S))
    if pos:
        S["dernier"] = pos[-1][1]
    for ex, LL in groupes.items():
        Lg = sorted([im for Lf in LL for im in Lf], key=lambda i: i["index"])
        fid = "p%d-e%s" % (pno, ex)
        for j, im in enumerate(Lg):
            im["figure"] = fid; im["morceau"] = j + 1; im["n_morceaux"] = len(Lg)
            im.pop("entier", None); im.pop("entier_cadre", None)
        if len(Lg) > 1:
            U = fitz.Rect(min(i["cadre"][0] for i in Lg), min(i["cadre"][1] for i in Lg), max(i["cadre"][2] for i in Lg), max(i["cadre"][3] for i in Lg)) & page.rect
            pe = page.get_pixmap(clip=U, matrix=fitz.Matrix(2, 2))
            Lg[0]["entier"] = "data:image/png;base64," + base64.b64encode(pe.tobytes("png")).decode()
            Lg[0]["entier_cadre"] = [U.x0, U.y0, U.x1, U.y1]
    if retires:
        otes = sorted(im["index"] for im in ims if (im.get("figure") or "") in retires)
        ims = [im for im in ims if (im.get("figure") or "") not in retires]
        for im in ims:
            im["index"] -= sum(1 for k in otes if k < im["index"])
        idx_fin -= len(otes)
    return ims, idx_fin


def _v254_montrer_son_dessin(page, ims, idx_fin):
    """Dans chaque image, les pixels d'un dessin d'une AUTRE figure (son encre, mots rattachés compris, pont d'un pixel) sont
    blanchis, hors pixels du dessin de l'image ; seulement pour un dessin qui a une image à lui ; un dessin de la même figure
    n'est jamais touché. Les deux images (modèle, feuille), par les mêmes pixels de page ; la grille d'une image est celle du
    moteur de rendu."""
    np_ = _np247; K = _V247_K
    R1 = _v254_etat()["r1"]
    if R1.get("doc") is not page.parent:
        return ims, idx_fin
    pno = page.number + 1
    D = R1["pages"].get(pno)
    if not D or "own" not in D[0]:
        return ims, idx_fin
    cle = R1["cle"]
    Hp, Wp = int(np_.ceil(page.rect.height * K)) + 2, int(np_.ceil(page.rect.width * K)) + 2
    qui = np_.full((Hp, Wp), -1, dtype=np_.int32); idx_tag = {}
    for i, dd in enumerate(D):
        idx_tag[dd["tag"]] = i
        bx0, by0, A = dd["own"]
        ys, xs = np_.nonzero(A); ys = ys + by0; xs = xs + bx0
        ok = (ys >= 0) & (xs >= 0) & (ys < Hp) & (xs < Wp); qui[ys[ok], xs[ok]] = i
    fig_de, im_d = {}, []
    for im in ims:
        t = cle.get((pno, tuple(round(v, 1) for v in im["cadre"])))
        i = idx_tag.get(t) if t is not None else None
        im_d.append(i)
        if i is not None:
            fig_de.setdefault(i, set()).add(im.get("figure"))
    for im, i in zip(ims, im_d):
        if i is None:
            continue
        c = im["cadre"]
        _ir = (fitz.Rect(c) * fitz.Matrix(K, K)).irect
        x0, y0, x1, y1 = _ir.x0, _ir.y0, _ir.x1, _ir.y1
        if x0 < 0 or y0 < 0 or x1 > Wp or y1 > Hp or x1 <= x0 or y1 <= y0:
            continue
        z = qui[y0:y1, x0:x1]; f = im.get("figure")
        autres = np_.zeros(z.shape, dtype=bool)
        for j in np_.unique(z[z >= 0]):
            j = int(j)
            if j == i or f in fig_de.get(j, set()) or not fig_de.get(j):
                continue
            autres |= (z == j)
        if not autres.any():
            continue
        un = np_.ones_like(autres, dtype=bool)
        autres = _v247_dilate(autres, 1, un) & ~_v247_dilate(z == i, 1, un)
        for cle_d in ("data", "data_feuille"):
            if not im.get(cle_d) or "," not in im[cle_d]:
                continue
            pm = fitz.Pixmap(base64.b64decode(im[cle_d].split(",", 1)[1]))
            if pm.alpha:
                pm = fitz.Pixmap(pm, 0)
            if pm.n != 3:
                pm = fitz.Pixmap(fitz.csRGB, pm)
            arr = np_.frombuffer(pm.samples, dtype=np_.uint8).reshape(pm.height, pm.width, 3).copy()
            if (pm.height, pm.width) == autres.shape:
                arr[autres] = 255
            else:
                # l'image de feuille : chaque pixel ramené à son point de page, puis à la grille à K du masque
                _m = _V238_PPP_FEUILLE / 72.0
                _irf = (fitz.Rect(c) * fitz.Matrix(_m, _m)).irect
                if (_irf.height, _irf.width) != (pm.height, pm.width):
                    continue
                _vy = ((_irf.y0 + np_.arange(pm.height) + 0.5) / _m * K).astype(int) - y0
                _vx = ((_irf.x0 + np_.arange(pm.width) + 0.5) / _m * K).astype(int) - x0
                _vy = np_.clip(_vy, 0, autres.shape[0] - 1); _vx = np_.clip(_vx, 0, autres.shape[1] - 1)
                arr[autres[_vy][:, _vx]] = 255
            png = fitz.Pixmap(fitz.csRGB, pm.width, pm.height, arr.tobytes(), False).tobytes("png")
            im[cle_d] = "data:image/png;base64," + base64.b64encode(png).decode()
            if cle_d == "data":
                im["sceau"] = _hl.sha256(png).hexdigest()[:12]
    return ims, idx_fin


# ══ v2.63 — LES FIGURES D'UNE PAGE NUMÉROTÉES DANS L'ORDRE DE LECTURE (Z9-v1) ══
# Condition : une page d'au moins deux images, dont les index valent « base + position » et dont chaque figure commence par son
# morceau 1. Sinon la page garde ses numéros, et la page est comptée « abstention » dans la réponse de /parse.
# Effet : les figures de la page (même champ `figure` ; une image sans figure compte seule) sont rangées par la clé de lecture du
# serveur (`cle_lecture` du module des grilles, sur l'ancre de `_v260_ancres`) ; à clé égale, l'ordre d'avant ; les morceaux d'une
# figure gardent leur ordre ; les index sont redonnés base, base + 1, … ; aucun autre champ ne change.
# (voir JOURNAL BACKEND v2.63)
def _v263_ordre_de_lecture(page, ims):
    """Rend (la liste de la page, la voie : "z9v1_renumerote", "z9v1_abstention" ou None) ; les mêmes objets image."""
    n = len(ims)
    if n < 2:
        return ims, None
    cle = getattr(_grilles, "cle_lecture", None) if _grilles else None
    if cle is None:
        return ims, "z9v1_abstention"
    base = ims[0]["index"]
    if [im["index"] for im in ims] != list(range(base, base + n)):
        return ims, "z9v1_abstention"
    unites, par = [], {}
    for im in ims:
        f = im.get("figure")
        if not f:
            unites.append([im])
            continue
        if f not in par:
            par[f] = []
            unites.append(par[f])
        par[f].append(im)
    if any(u[0].get("morceau", 1) != 1 for u in unites):
        return ims, "z9v1_abstention"
    anc = {a["index"]: a["cadre"] for a in _v260_ancres(page, ims)}
    cles = []
    for u in unites:
        a = anc.get(u[0]["index"]) or (u[0].get("entier_cadre") or u[0]["cadre"])
        cles.append(cle(a[1], a[0]))
    ordre = sorted(range(len(unites)), key=lambda i: (cles[i], i))
    neuf = [im for i in ordre for im in unites[i]]
    if [id(a) for a in neuf] == [id(b) for b in ims]:
        return ims, None
    for k, im in enumerate(neuf):
        im["index"] = base + k
    return neuf, "z9v1_renumerote"


# ═══ v2.66 — ÉTAPE B : TROIS CHAMPS QUE LA PAGE LIRA (voisins_b part B · lignes_couleur · la couleur des mots) ═══
# Rien de ce qui suit ne change les images, le texte ni une autre clé de /parse : trois clés de plus, chacune avec sa panne
# déclarée (jamais tue). Le pourquoi, les mesures et les écarts au banc : JOURNAL BACKEND v2.66.
import math as _m266
import collections as _co266

# ── F015 B (V166-LP, DEC-2590) : les phrases à trous à 6 mm au plus d'une image d'exercice, dont aucun mot n'est porté par une
#    image, hors consigne lue et hors « point final isolé » ; rendues par figure. Les étiquettes cursives (E49) s'y ajoutent
#    en v2.67 (DEC-3270). Portage de c252/ch1 : occurrences/mesure_occ_252 (c), sa variante c2, solutions/S2_E_lignes/s2_champ_252,
#    s2_doubles_252, ronde2/B_complete_si_absent/b_champ_252, c249/compte_e/e1_champ_249 (inter), c248 mesure_cursives (ecart_mm).
_VB_CH_P = set(".…_")
_VB_POIDS = {".": 1, "…": 3, "_": 2}
_VB_RE_L = re.compile(r"[A-Za-zÀ-ÿ]")
_VB_RE_3L = re.compile(r"[A-Za-zÀ-ÿ]{3,}")
_VB_RE_EX = re.compile(r"^p(\d+)-e\d+")        # figure d'exercice
_VB_MM = 25.4 / 72.0
_VB_EPS_AIRE = 0.01
_VB_TOL = 0.15


def _vb_ecart_mm(r1, r2):
    dx = max(r2.x0 - r1.x1, r1.x0 - r2.x1, 0.0); dy = max(r2.y0 - r1.y1, r1.y0 - r2.y1, 0.0)
    return _m266.hypot(dx, dy) * _VB_MM


def _vb_proche(r, ims):
    """(distance mm, image) la plus proche ; la PREMIÈRE de distance minimale."""
    return min([(_vb_ecart_mm(r, fitz.Rect(im["cadre"])), im) for im in ims], key=lambda t: t[0], default=(None, None))


def _vb_jetons_page(page):
    """Tous les jetons horizontaux : (bbox, texte, alnum)."""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            if not abs(l["dir"][1]) < 0.1:
                continue
            for s in l["spans"]:
                cur = []
                def ferme():
                    if cur:
                        t = "".join(c["c"] for c in cur)
                        r = fitz.Rect(cur[0]["bbox"])
                        for c in cur[1:]:
                            r |= fitz.Rect(c["bbox"])
                        out.append(dict(bbox=r, texte=t, alnum=any(ch.isalnum() for ch in t)))
                        cur.clear()
                for ch in s["chars"]:
                    if ch["c"].strip() == "":
                        ferme()
                    else:
                        cur.append(ch)
                ferme()
    return out


def _vb_p2(t):
    """Le poids de la plus longue suite collée de « . … _ » dans un mot."""
    best = cur = 0
    for ch in t:
        if ch in _VB_CH_P:
            cur += _VB_POIDS[ch]; best = max(best, cur)
        else:
            cur = 0
    return best


def _vb_rect_items(d):
    it = d.get("items") or []; kinds = [x[0] for x in it]
    out = []
    if it and all(k in ("re", "qu") for k in kinds):
        for x in it:
            out.append(fitz.Rect(x[1]) if x[0] == "re" else fitz.Quad(x[1]).rect)
        return out
    if kinds.count("l") == 4 and len(kinds) == 4:
        if all(abs(x[1].x - x[2].x) <= 0.5 or abs(x[1].y - x[2].y) <= 0.5 for x in it):
            return [fitz.Rect(d["rect"])]
        return []
    if len(kinds) == 8 and kinds.count("l") == 4 and kinds.count("c") == 4:
        return [fitz.Rect(d["rect"])]
    return []


def _vb_est_p1(j):
    return (not j["alnum"]) and j["texte"] and all(ch in _VB_CH_P for ch in j["texte"])


def _vb_cases_et_tirets(page, mots):
    """Une case K (côtés ≥ 6 pt, aire ≤ 40 000 pt²) ne contient le centre d'aucun jeton, sauf des pointillés P1 (variante c2) ;
    un tiret P3 : un trait pointillé horizontal d'au moins 10 pt. Le paramètre `mots` du banc est remplacé, comme en c2, par les
    jetons pleins de la page."""
    pleins = [j for j in _vb_jetons_page(page) if not _vb_est_p1(j)]
    cases, tirets = [], []
    for d in page.get_drawings():
        ds = d.get("dashes"); pointille = bool(ds) and str(ds).strip() not in ("", "[] 0")
        it = d.get("items") or []
        if pointille:
            if len(it) == 1 and it[0][0] == "l":
                a, b = it[0][1], it[0][2]
                if abs(a.y - b.y) <= 1.0 and abs(a.x - b.x) >= 10.0:
                    tirets.append(fitz.Rect(min(a.x, b.x), min(a.y, b.y) - 0.5, max(a.x, b.x), max(a.y, b.y) + 0.5))
            continue
        for r in _vb_rect_items(d):
            if r.width < 6 or r.height < 6 or r.width * r.height > 40000:
                continue
            if any(r.contains(fitz.Point((w["bbox"].x0 + w["bbox"].x1) / 2, (w["bbox"].y0 + w["bbox"].y1) / 2)) for w in pleins):
                continue
            cases.append(r)
    return cases, tirets


def _vb_vov(a, b): return min(a.y1, b.y1) - max(a.y0, b.y0)
def _vb_hgap(a, b): return max(b.x0 - a.x1, a.x0 - b.x1)


def _vb_lignes_page(page):
    """Les lignes (composantes) de la page : liste de dict(elements=[(sorte, Rect, texte)], rect)."""
    J = _vb_jetons_page(page)
    mots = [j for j in J if j["alnum"]]
    p1 = [j for j in J if (not j["alnum"]) and j["texte"] and all(ch in _VB_CH_P for ch in j["texte"])]
    cases, tirets = _vb_cases_et_tirets(page, mots)
    E = [("mot", j["bbox"], j["texte"]) for j in mots] + [("P1", j["bbox"], j["texte"]) for j in p1] + \
        [("K", r, "") for r in cases] + [("P3", r, "") for r in tirets]
    par = list(range(len(E)))
    def tr(a):
        while par[a] != a:
            par[a] = par[par[a]]; a = par[a]
        return a
    def lie(i, j):
        si, ri, _ = E[i]; sj, rj, _ = E[j]
        tx = ("mot", "P1")
        if si in tx and sj in tx:
            hi, hj = ri.height, rj.height
            return _vb_vov(ri, rj) >= 0.5 * min(hi, hj) and _vb_hgap(ri, rj) <= 1.5 * max(hi, hj)
        if si in tx or sj in tx:
            w, o, so = (ri, rj, sj) if si in tx else (rj, ri, si)
            h = w.height
            if so == "K":
                return (0.3 * h <= o.height <= 3 * h) and _vb_vov(w, o) >= 0.5 * min(h, o.height) and _vb_hgap(w, o) <= 1.5 * h
            if so == "P3":
                yc = (o.y0 + o.y1) / 2
                return (w.y0 <= yc <= w.y1 + 0.5 * h) and _vb_hgap(w, o) <= 1.5 * h
        return False
    for i in range(len(E)):
        for j in range(i + 1, len(E)):
            if lie(i, j):
                par[tr(i)] = tr(j)
    grp = {}
    for i in range(len(E)):
        grp.setdefault(tr(i), []).append(i)
    out = []
    for g, idx in grp.items():
        el = sorted((E[i] for i in idx), key=lambda e: (e[1].x0, e[1].y0))
        R = fitz.Rect(el[0][1])
        for e in el[1:]:
            R |= e[1]
        out.append(dict(elements=el, rect=R))
    return out


def _vb_mesure_c(doc, images, corrige_pages, borne=6.0):
    """Les phrases à trou des pages gardées à `borne` mm au plus d'une image d'exercice (toutes si `borne` est None : RL2, v2.69), chacune
    avec sa distance aux images d'exercice et son statut."""
    cors = set(corrige_pages or []); res = []
    for p in doc:
        pno = p.number + 1
        if pno in cors:
            continue
        ims = [im for im in images if im["page"] == pno]
        ex = [im for im in ims if _VB_RE_EX.match(im.get("figure") or "")]
        def dans(r):
            c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
            return [im for im in ims if fitz.Rect(im["cadre"]).contains(c)]
        for L in _vb_lignes_page(p):
            el = L["elements"]
            mots = [e for e in el if e[0] == "mot"]
            lettres = [e for e in mots if _VB_RE_L.search(e[2])]
            if not (len(lettres) >= 2 and any(_VB_RE_3L.search(e[2]) for e in lettres)):
                continue
            k = [e for e in el if e[0] == "K"]; t1 = [e for e in el if e[0] == "P1"]; t3 = [e for e in el if e[0] == "P3"]
            t2 = [e for e in mots if _vb_p2(e[2]) >= 4]
            if borne is None and not (k or t1 or t2 or t3):  # RP1 (DEC-3362, rang 2) : la part B (borne 6 mm) ne demande plus de trou
                continue
            R = L["rect"]
            d_ex, i_ex = _vb_proche(R, ex)
            portes = [e for e in mots if dans(e[1])]
            statut = "aucun mot porté" if not portes else ("toute portée" if len(portes) == len(mots) else "en partie portée")
            texte = " ".join("[ ]" if e[0] == "K" else ("[---]" if e[0] == "P3" else e[2]) for e in el)
            res.append(dict(page=pno, texte=texte, bbox=[round(v, 1) for v in R], n_cases=len(k), n_p1=len(t1), n_p2=len(t2),
                            p1=[dict(texte=e[2]) for e in t1], statut=statut,
                            dist_ex_mm=None if d_ex is None else round(d_ex, 1),
                            image_ex_proche=None if i_ex is None else dict(figure=i_ex.get("figure"))))
    return res if borne is None else [l for l in res if l["dist_ex_mm"] is not None and l["dist_ex_mm"] <= borne]


def _vb_lignes_doc(doc, images, corrige_pages, borne=6.0):
    """Les lignes à trous à `borne` mm au plus d'une image d'exercice (6 par défaut ; toutes si None : RL2), chacune avec ses éléments (sorte, boîte, texte, images qui en portent
    le centre), appariés par (page, rectangle arrondi à 0,1) ; un appariement manquant ou double est une panne déclarée."""
    ims = [dict(page=i["page"], figure=(i["figure"] if i.get("figure") is not None else "(sans figure)"),
                morceau=i.get("morceau"), cadre=i["cadre"]) for i in images]
    L = _vb_mesure_c(doc, ims, corrige_pages, borne)
    if not L:
        return []
    comps = {}
    for pno in sorted({l["page"] for l in L}):
        for c in _vb_lignes_page(doc[pno - 1]):
            k = (pno, tuple(round(v, 1) for v in c["rect"]))
            comps.setdefault(k, []).append(c)
    out = []
    for l in L:
        k = (l["page"], tuple(l["bbox"]))
        if len(comps.get(k, [])) != 1:
            raise RuntimeError("ligne %s non appariée à ses éléments (%d)" % (k, len(comps.get(k, []))))
        cad = [tuple(im["cadre"]) for im in ims if im["page"] == l["page"]]
        figs = [im for im in ims if im["page"] == l["page"]]
        els = []
        for (s, r, t) in comps[k][0]["elements"]:
            bb = [float(r.x0), float(r.y0), float(r.x1), float(r.y1)]
            cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
            dedans = [n for n, c in enumerate(cad) if c[0] <= cx <= c[2] and c[1] <= cy <= c[3]]
            els.append(dict(sorte=s, bbox=bb, texte=t, dans_image=[figs[i]["figure"] + "#%s" % figs[i].get("morceau") for i in dedans]))
        l = dict(l); l["elements"] = els
        out.append(l)
    return out


def _vb_poses_ligne(l):
    """Ce qui se pose d'une ligne « aucun mot porté » (V166) : ses jetons de texte (mot, P1) dont le centre n'est dans aucune image."""
    if l["statut"] != "aucun mot porté":
        return []
    fig = (l.get("image_ex_proche") or {}).get("figure")
    lid = (l["page"], tuple(l["bbox"]))
    return [dict(page=l["page"], texte=e["texte"], bbox=e["bbox"], sorte="ligne", figure=fig, ligne=lid)
            for e in l["elements"] if e["sorte"] in ("mot", "P1") and not e["dans_image"]]


def _vb_reunit(L):
    """Doublons ôtés (même page, même texte, boîte à 0,15 pt) ; le premier gardé."""
    out = []
    for p in L:
        if any(q["page"] == p["page"] and q["texte"] == p["texte"] and all(abs(x - y) <= _VB_TOL for x, y in zip(p["bbox"], q["bbox"]))
               for q in out):
            continue
        out.append(p)
    return out


def _vb_retiree(lig, reps):
    """Une ligne qui recoupe (aire > 0,01) la boîte du numéro ou de la ligne d'une consigne lue de sa page est retirée."""
    for ex, r in (reps or {}).items():
        for k in ("numero", "ligne"):
            b = r.get(k)
            if b:
                a = lig["bbox"]
                x0, y0, x1, y1 = max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])
                if x1 > x0 and y1 > y0 and (x1 - x0) * (y1 - y0) > _VB_EPS_AIRE:
                    return True
    return False


def _vb_point_final_isole(l):
    """La forme « 2 point final isolé » (s2_doubles_252.classe_c) : des pointillés faits d'un seul « . » final, sans case."""
    if l["statut"] == "toute portée":
        return False
    return (not l["n_cases"]) and l["n_p1"] == 1 and l["n_p2"] == 0 and l["p1"][0]["texte"] == "." and l["texte"].endswith(" .")


def _vb_part_b(doc, images, corrige_pages, reps_par_page, e49=()):
    """voisins_b, part B : {figure: [{texte, bbox (0,01), sorte « ligne » ou « etiquette » (E49, v2.67), ligne « page|x0|y0|x1|y1 » ou None, lt}]}."""
    L = _vb_lignes_doc(doc, images, corrige_pages)
    poses = _vb_reunit([p for l in L for p in _vb_poses_ligne(l)])
    Ld = {(l["page"], tuple(l["bbox"])): l for l in L}
    garde = [q for q in poses if not _vb_retiree(Ld[q["ligne"]], reps_par_page.get(q["page"]))
             and not _vb_point_final_isole(Ld[q["ligne"]])]
    out = {}
    # v2.67 — les étiquettes cursives (E49) se réunissent après les lignes gardées, comme au banc (b_champ_252.variantes_b)
    for q in _vb_reunit(garde + list(e49)):
        k = q["ligne"]
        out.setdefault(q["figure"], []).append(dict(texte=q["texte"], bbox=[round(v, 2) for v in q["bbox"]], sorte=q["sorte"],
                                                    ligne=("%d|%s" % (k[0], "|".join("%.2f" % v for v in k[1]))) if k else None,
                                                    lt=Ld[k]["texte"] if k else None))
    return out


# ── F083 R121a (DEC-3156) : la couleur des mots listés, lue dans la couche texte du PDF ; franche si l'écart de canaux est d'au
#    moins 48 pour la moitié au moins des caractères du mot (pris dans la ligne qui en porte le plus). Portage de c287/A205_T2R
#    lib_a205_v2 (r121a) et variantes_analyses_a205 (la boucle).
_V266_CHROMA_MOT = 48


def _v266_couleur_mot(raw, bb):
    par_ligne = _co266.OrderedDict()
    for ib, b in enumerate(raw["blocks"]):
        if b.get("type") != 0:
            continue
        for il, l in enumerate(b["lines"]):
            for s in l["spans"]:
                for c in s["chars"]:
                    cb = c["bbox"]; cx = (cb[0] + cb[2]) / 2.0; cy = (cb[1] + cb[3]) / 2.0
                    if bb[0] <= cx <= bb[2] and bb[1] <= cy <= bb[3]:
                        par_ligne.setdefault((ib, il), []).append(s["color"])
    if not par_ligne:
        return None
    cols = [((c >> 16) & 255, (c >> 8) & 255, c & 255) for c in par_ligne[max(par_ligne, key=lambda k: len(par_ligne[k]))]]
    chrom = [c for c in cols if int(max(c) - min(c)) >= _V266_CHROMA_MOT]
    if not chrom or len(chrom) / float(len(cols)) < 0.5:
        return None
    return "#%02x%02x%02x" % _co266.Counter(chrom).most_common(1)[0][0]


def _v266_couleurs_voisins(doc, vb):
    """Ajoute `couleur` aux entrées de voisins_b dont le mot est franchement coloré dans la source."""
    raws = {}
    for fg, LL in vb.items():
        m = re.match(r"^p(\d+)-", fg or "")
        if not m or not (1 <= int(m.group(1)) <= doc.page_count):
            continue
        pno = int(m.group(1))
        if pno not in raws:
            raws[pno] = doc[pno - 1].get_text("rawdict")
        for e in LL:
            c = _v266_couleur_mot(raws[pno], e["bbox"])
            if c:
                e["couleur"] = c


# ── F088 C1e (DEC-2901, DEC-3159) : TOUTES les lignes du texte (pages jugées corrigé exclues) qui portent au moins un caractère
#    alphanumérique, avec la couleur de chaque caractère quand elle est franche (écart de canaux d'au moins 40), sinon null.
#    Le nom du champ est celui que la page lit ; il porte des lignes de toute couleur. Portage de c287/A204_BU lignes_texte.
_V266_CHROMA_LIGNE = 40


def _v266_lignes_couleur(doc, pages_exclues):
    res = []
    for pno in range(len(doc)):
        if (pno + 1) in pages_exclues:
            continue
        for blk in doc[pno].get_text("rawdict").get("blocks", []):
            for ln in blk.get("lines", []):
                texte, coul = [], []
                for sp in ln.get("spans", []):
                    c = sp.get("color", 0); rgb = ((c >> 16) & 255, (c >> 8) & 255, c & 255)
                    hexc = "#%02x%02x%02x" % rgb if (max(rgb) - min(rgb)) >= _V266_CHROMA_LIGNE else None
                    for ch in sp.get("chars", []):
                        texte.append(ch["c"]); coul.append(hexc)
                if any(t.isalnum() for t in texte):
                    res.append({"p": pno + 1, "t": "".join(texte), "c": coul})
    return res


# ═══ v2.67 — LA PART CURSIVE DE voisins_b (DEC-3270 point 2) : les étiquettes cursives (E49) et les mots cursifs de D6 avec leur cache ═══
# Rien de ce qui suit ne change les images, le texte ni une autre clé de /parse : des entrées de plus dans voisins_b (sorte « etiquette »,
# sorte « cursif » avec `cache` et `voie`), et un `cache` sur les entrées déjà listées qui sont les mêmes mots. L'effacement des images
# sous le cache (F036) n'est pas ici. La reconnaissance des polices cursives est celle de l'essai du 17.10 (R2) ; son faux positif connu
# (Belle Allure Script3i-Fin) attend sa réponse (D1) : cette version n'entre dans aucun dépôt avant.
# Portage de c248/chB/b2_cursives (cursif2_b2_248, mesure_cursives_b2_248), c252/ch1/occurrences/mesure_occ_252 (compte_a, etiquette),
# s2_champ_252 (poses_mot, reunit), b_champ_252 (variantes_b), c252/ch6/ronde3/P_cube/compte_pcube_252 (legendes), outils/d6_champ_256d
# (champ_d6, lignage, cache_encre), outils270/e_270 (champ). Écarts au banc : les composantes d'encre sont étiquetées par _v247_etiqueter
# (8-connexes, comme scipy et le remplissage du banc) ; le document est celui que parse_pdf a ouvert ; les mots cursifs sont mesurés une
# fois pour E49 et D6. Le pourquoi et les mesures : JOURNAL BACKEND v2.67.
_V267_RE_CURSIVE = re.compile(r"cursive", re.I)   # une police nommée « cursive » est cursive
_V267_LETTRES_MIN = 5                             # un span de moins de 5 lettres n'est pas jugé
_V267_ZOOM = 4                                    # rendu du span à 4×, en gris
_V267_ENCRE = 140                                 # encre : gris < 140
_V267_ENCRE_MIN = 30                              # moins de 30 pixels d'encre : non jugé
_V267_SOULIGNE = 8                                # une composante plus large que 8 fois sa hauteur (soulignement) sort du compte
_V267_H_REL = 0.35                                # composante gardée : hauteur ≥ 0,35 × la plus haute
_V267_AIRE_MIN = 6                                # … et aire ≥ 6 pixels
_V267_Q_CURSIF = 0.60                             # span cursif : moins de 0,60 composante par lettre
_V267_PART = 0.5                                  # police cursive : plus de la moitié de ses spans jugés cursifs
_V267_SPANS = 25                                  # au plus 25 spans par police, les plus longs
_V267_ISOLE_EM = 0.6                              # mot isolé : écart à chaque voisin de ligne > 0,6 × la taille
_V267_RECOUV = 0.5                                # voisin de ligne : recouvrement vertical ≥ 0,5 × la plus petite hauteur
_V267_BORD_PT = 0.5                               # voisin à gauche / à droite, à 0,5 pt près
_V267_SEUIL_MM = 6.0                              # à 6 mm au plus d'une figure (arrondi à 0,1 mm, comme le banc)
_V267_PONCT_E49 = ".?!:;…"                        # étiquette : sans ponctuation de fin
_V267_PONCT_D6 = ".,;:!?…"                        # D6, voie (a) : mot isolé non ponctué
_V267_FAM2 = re.compile(r"^p\d+-[ed]")            # la famille des figures de D6 (exercices et dessins)
_V267_COUVERTURE = 0.5                            # pièce-légende : ses mots couvrent au moins la moitié de son cadre
_V267_COPIES_PT = 0.5                             # copies superposées d'un mot, et doublon dans voisins_b : boîtes à 0,5 pt
_V267_LIGNAGE_ZOOM = 2                            # lignage : rendu à 2× sous le mot
_V267_LIGNAGE_ENCRE = 230                         # … encre < 230
_V267_LIGNAGE_PART = 0.6                          # … une rangée de lignage porte de l'encre sur 60 % de sa largeur
_V267_LIGNAGE_N = 3                               # … au moins 3 lignes (ou 3 segments tracés qui traversent le mot)
_V267_TRAIT_DY = 0.5                              # segment horizontal : écart vertical < 0,5 pt
_V267_TRAIT_H = 1.5                               # rectangle-trait : hauteur < 1,5 pt
_V267_CACHE_PPP = 300                             # cache : rendu à 300 ppp autour du mot (une hauteur de marge)
_V267_CACHE_ENCRE = 200                           # … encre < 200


def _v267_nom_propre(f):
    return f.split("+", 1)[1] if len(f) > 7 and f[6] == "+" else f


def _v267_boites(mask):
    """[(x0, y0, x1, y1, aire)] des composantes 8-connexes d'un masque (bornes incluses)."""
    lab = _v247_etiqueter(mask, huit=True)
    ys, xs = _np247.nonzero(lab)
    if not len(ys):
        return []
    ids, inv = _np247.unique(lab[ys, xs], return_inverse=True)
    inv = inv.ravel()
    n = len(ids)
    x0 = _np247.full(n, _np247.iinfo(_np247.int64).max, dtype=_np247.int64); y0 = x0.copy()
    x1 = _np247.full(n, -1, dtype=_np247.int64); y1 = x1.copy()
    _np247.minimum.at(x0, inv, xs); _np247.minimum.at(y0, inv, ys)
    _np247.maximum.at(x1, inv, xs); _np247.maximum.at(y1, inv, ys)
    aire = _np247.bincount(inv, minlength=n)
    return [(int(a), int(b), int(c), int(d), int(e)) for a, b, c, d, e in zip(x0, y0, x1, y1, aire)]


def _v267_composantes_par_lettre(page, bbox, texte, direction):
    """Composantes d'encre gardées par lettre d'un span horizontal (None : non jugé)."""
    if abs(direction[1]) > 0.1:
        return None
    lettres = sum(1 for c in texte if c.isalpha())
    if lettres < _V267_LETTRES_MIN:
        return None
    r = fitz.Rect(bbox) + (1, 1, 1, 1)
    pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(_V267_ZOOM, _V267_ZOOM), colorspace=fitz.csGRAY, alpha=False)
    a = _np247.frombuffer(pix.samples, dtype=_np247.uint8).reshape(pix.height, pix.width)
    enc = a < _V267_ENCRE
    if enc.sum() < _V267_ENCRE_MIN:
        return None
    ok = [(y1 - y0 + 1, ar) for x0, y0, x1, y1, ar in _v267_boites(enc) if (x1 - x0 + 1) <= _V267_SOULIGNE * max(y1 - y0 + 1, 1)]
    if not ok:
        return None
    hmax = max(h for h, ar in ok)
    return len([1 for h, ar in ok if h >= _V267_H_REL * hmax and ar >= _V267_AIRE_MIN]) / lettres


# ═══ v2.69 — LA FORME 2 DE LA RECONNAISSANCE CURSIVE (DEC-3273 R3 ; mesurée : D1) : R3a, le dessin des lettres lu dans la police ═══
# Portage de c288/A241_S126/outils/a241_glyphes.py (analyse_police, part R3a ; traite). Dépend de fontTools (présent par pdf2docx) : s'il manque,
# R3a ne juge rien et la forme 2 rend R2. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_V269_R3A_EPS = 0.02
_V269_R3A_SEUIL = 0.50
_V269_R3A_NMIN = 10


def _v269_r3a_police(doc, xr, paires):
    """R3a d'un programme de police intégré : True (cursive : la moitié au moins des lettres a-z employées touchent les deux bords de leur chasse,
    à 0,02 em), False, ou None (moins de 10 lettres lisibles, ou programme illisible)."""
    import io as _io
    from fontTools.ttLib import TTFont
    from fontTools.pens.boundsPen import BoundsPen
    nm, ext, typ, buf = doc.extract_font(xr)
    if not buf:
        return None
    if buf[:4] in (b"\x00\x01\x00\x00", b"true", b"OTTO", b"ttcf"):
        f = TTFont(_io.BytesIO(buf), lazy=True, fontNumber=0); upm = f["head"].unitsPerEm; order = f.getGlyphOrder()
        cff = ("CFF " in f); glyf = None if cff else f["glyf"]; cs = f["CFF "].cff[f["CFF "].cff.fontNames[0]].CharStrings if cff else None
        def bornes(gname):
            if cff:
                pen = BoundsPen(None); cs[gname].draw(pen); return pen.bounds
            g = glyf[gname]
            if g.numberOfContours == 0:
                return None
            return (g.xMin, g.yMin, g.xMax, g.yMax)
    else:
        from fontTools.cffLib import CFFFontSet
        cf = CFFFontSet(); cf.decompile(_io.BytesIO(buf), None); top = cf[cf.fontNames[0]]; cs = top.CharStrings; order = list(top.charset)
        fm = getattr(top, "FontMatrix", None); upm = int(round(1 / fm[0])) if fm else 1000
        def bornes(gname):
            pen = BoundsPen(None); cs[gname].draw(pen); return pen.bounds
    liees = {}
    for (u, gid), ws in paires.items():
        if gid is None or gid >= len(order) or gid < 1:
            continue
        try:
            bb = bornes(order[gid])
        except Exception:
            continue
        if bb is None or not ws:
            continue
        ws = sorted(ws); adv = ws[len(ws) // 2] * upm
        liees[u] = (bb[0] <= _V269_R3A_EPS * upm) and (bb[2] >= adv - _V269_R3A_EPS * upm)
    if len(liees) < _V269_R3A_NMIN:
        return None
    return bool(sum(liees.values()) / len(liees) >= _V269_R3A_SEUIL)


def _v269_r3a(doc):
    """{police (nom propre) : True | False | None} pour un document : tous les programmes d'un même nom doivent dire oui ; aucun jugé : None.
    Sans fontTools : {} (R3a ne juge rien)."""
    try:
        import fontTools  # noqa: F401
    except Exception:
        return {}
    xrefs = {}
    for p in doc:
        for f in p.get_fonts(full=True):
            xrefs.setdefault(f[0], dict(ext=f[1], base=f[3]))
    paires = _co266.defaultdict(dict)
    for p in doc:
        try:
            tt = p.get_texttrace()
        except Exception:
            tt = []
        for sp in tt:
            for ch in sp["chars"]:
                if 97 <= ch[0] <= 122 and abs(sp["dir"][1]) < 0.1 and sp["size"] > 0:
                    paires[sp["font"]].setdefault((ch[0], ch[1]), []).append((ch[3][2] - ch[3][0]) / sp["size"])
    dec = _co266.defaultdict(list)
    for xr, inf in xrefs.items():
        nom = _v267_nom_propre(inf["base"]); r = None
        if inf["ext"] != "n/a":
            try:
                r = _v269_r3a_police(doc, xr, paires.get(nom, {}))
            except Exception:
                r = None
        dec[nom].append(r)
    out = {}
    for nom, L in dec.items():
        L = [x for x in L if x is not None]
        out[nom] = all(L) if L else None
    return out


def _v267_polices_cursives(doc):
    """{police : dict(nom, cursive, par_nom, part_mesure, spans_juges)} : cursive par son nom, ou par la mesure de ses spans ; v2.69, la forme 2
    (DEC-3273) : sauf si le dessin des lettres (R3a) l'a jugée non cursive."""
    echant = {}
    for p in doc:
        for b in p.get_text("dict")["blocks"]:
            if b["type"] != 0:
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    t = s["text"].strip()
                    if sum(1 for c in t if c.isalpha()) >= _V267_LETTRES_MIN:
                        echant.setdefault(_v267_nom_propre(s["font"]), []).append((p.number, s["bbox"], t, l["dir"]))
    res = {}
    for nom, sp in echant.items():
        votes = []
        for pn, bb, t, dr in sorted(sp, key=lambda x: -len(x[2]))[:_V267_SPANS]:
            q = _v267_composantes_par_lettre(doc[pn], bb, t, dr)
            if q is not None:
                votes.append(q < _V267_Q_CURSIF)
        part = (sum(votes) / len(votes)) if votes else None
        par_nom = bool(_V267_RE_CURSIVE.search(nom))
        res[nom] = dict(nom=nom, cursive=par_nom or (part is not None and part > _V267_PART), par_nom=par_nom,
                        part_mesure=part, spans_juges=len(votes))
    # v2.69 — la forme 2 : R3a « non » l'emporte sur R2 ; R3a non jugée ne contredit pas
    r3a = _v269_r3a(doc) if res else {}
    for nom, v in res.items():
        if v["cursive"] and r3a.get(nom) is False:
            v["cursive"] = False
    return res


def _v267_mots_page(page):
    """Tous les mots de la page : (bbox, texte, police, taille, horizontal)."""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for l in b["lines"]:
            horiz = abs(l["dir"][1]) < 0.1
            for s in l["spans"]:
                nom = _v267_nom_propre(s["font"]); cur = []
                for ch in list(s["chars"]) + [None]:
                    if ch is not None and ch["c"].strip() != "":
                        cur.append(ch)
                        continue
                    if cur:
                        t = "".join(c["c"] for c in cur)
                        if any(x.isalnum() for x in t):
                            r = fitz.Rect(cur[0]["bbox"])
                            for c in cur[1:]:
                                r |= fitz.Rect(c["bbox"])
                            out.append(dict(bbox=tuple(r), texte=t, police=nom, taille=s["size"], horizontal=horiz))
                        cur = []
    return out


def _v267_isole(w, tous):
    """Isolé : l'écart à son voisin de gauche et à son voisin de droite (s'ils existent) dépasse 0,6 × sa taille."""
    r = fitz.Rect(w["bbox"]); g = d = None
    for v in tous:
        if v is w or not v["horizontal"]:
            continue
        q = fitz.Rect(v["bbox"])
        if min(r.y1, q.y1) - max(r.y0, q.y0) < _V267_RECOUV * min(r.height, q.height):
            continue
        if q.x1 <= r.x0 + _V267_BORD_PT:
            g = r.x0 - q.x1 if g is None else min(g, r.x0 - q.x1)
        elif q.x0 >= r.x1 - _V267_BORD_PT:
            d = q.x0 - r.x1 if d is None else min(d, q.x0 - r.x1)
    return all(e / w["taille"] > _V267_ISOLE_EM for e in (g, d) if e is not None)


def _v267_mots_cursifs_doc(doc, images, corrige_pages):
    """Les mots en police cursive des pages gardées, mesurés une fois (E49 et D6 les lisent), et les étiquettes E49 : isolées,
    sans ponctuation de fin, hors de toute image, à 6 mm au plus d'une image d'exercice (la plus proche donne la figure)."""
    polices = _v267_polices_cursives(doc)
    ims_all = [dict(page=i["page"], figure=(i["figure"] if i.get("figure") is not None else "(sans figure)"), cadre=i["cadre"]) for i in images]
    cors = set(corrige_pages or [])
    mots, tous_pp = [], {}
    for p in doc:
        pno = p.number + 1
        if pno in cors:
            continue
        tous = _v267_mots_page(p)
        tous_pp[pno] = tous
        cur = [w for w in tous if polices.get(w["police"], {}).get("cursive")]
        if not cur:
            continue
        ims = [im for im in ims_all if im["page"] == pno]
        ex = [im for im in ims if _VB_RE_EX.match(im["figure"] or "")]
        for w in cur:
            r = fitz.Rect(w["bbox"]); c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
            d_ex, i_ex = _vb_proche(r, ex)
            mots.append(dict(page=pno, mot=w["texte"], bbox=[round(v, 1) for v in w["bbox"]],
                             dans_image=any(fitz.Rect(im["cadre"]).contains(c) for im in ims), isole=_v267_isole(w, tous),
                             dist_ex_mm=None if d_ex is None else round(d_ex, 1), fig_ex=None if i_ex is None else i_ex["figure"]))
    e49 = [dict(page=m["page"], texte=m["mot"], bbox=[float(v) for v in m["bbox"]], sorte="etiquette", figure=m["fig_ex"], ligne=None)
           for m in mots if m["dist_ex_mm"] is not None and not m["dist_ex_mm"] > _V267_SEUIL_MM and not m["dans_image"]
           and m["isole"] and not (m["mot"] and m["mot"][-1] in _V267_PONCT_E49)]
    return dict(polices=polices, mots=mots, tous=tous_pp, e49=e49)


def _v267_centre_dans(bb, c):
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    return c[0] <= cx <= c[2] and c[1] <= cy <= c[3]


def _v267_ecart(r1, r2):
    return _vb_ecart_mm(fitz.Rect(r1), fitz.Rect(r2))


def _v267_aire_union(rects):
    """Aire de l'union de rectangles (x0, y0, x1, y1), par balayage."""
    rects = [r for r in rects if r[2] > r[0] and r[3] > r[1]]
    if not rects:
        return 0.0
    xs = sorted({r[0] for r in rects} | {r[2] for r in rects}); tot = 0.0
    for a, b in zip(xs, xs[1:]):
        cur = 0.0; hi = None
        for y0, y1 in sorted((r[1], r[3]) for r in rects if r[0] <= a and r[2] >= b):
            if hi is None or y0 > hi:
                cur += y1 - y0; hi = y1
            elif y1 > hi:
                cur += y1 - hi; hi = y1
        tot += cur * (b - a)
    return tot


def _v267_legendes(ip, boites):
    """Les pièces-légende d'une page : une image dont les mots couvrent au moins la moitié du cadre, à 6 mm d'une image de la famille."""
    qc = set()
    for rang, fig, c in ip:
        dedans = [bb for bb in boites if _v267_centre_dans(bb, c)]
        aire = (c[2] - c[0]) * (c[3] - c[1])
        if not dedans or aire <= 0:
            continue
        inter = [(max(bb[0], c[0]), max(bb[1], c[1]), min(bb[2], c[2]), min(bb[3], c[3])) for bb in dedans]
        if _v267_aire_union(inter) / aire >= _V267_COUVERTURE:
            qc.add(rang)
    out = set()
    for rang, fig, c in ip:
        if rang not in qc:
            continue
        for r2, f2, c2 in ip:
            if r2 == rang or r2 in qc or not _V267_FAM2.match(f2):
                continue
            if round(_v267_ecart(c, c2), 1) <= _V267_SEUIL_MM:
                out.add(rang)
                break
    return out


def _v267_lignage(page, bb, memo):
    """Un mot écrit sur un lignage (au moins 3 lignes sous lui dans une image, ou 3 traits qui le traversent) : ni posé ni caché."""
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    m = memo.setdefault(page.number, {})
    if "infos" not in m:
        m["infos"] = page.get_image_info(); m["mots"] = page.get_text("words")
    for inf in m["infos"]:
        R = fitz.Rect(inf["bbox"])
        if not (R.x0 <= cx <= R.x1 and R.y0 <= cy <= R.y1):
            continue
        clip = fitz.Rect((R.x0, max(R.y0, bb[1]), R.x1, min(R.y1, bb[3])))
        if clip.is_empty:
            continue
        z = _V267_LIGNAGE_ZOOM
        pm = page.get_pixmap(matrix=fitz.Matrix(z, z), clip=clip, alpha=False)
        A = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].copy()
        for w in m["mots"]:
            r = fitz.Rect(w[:4])
            if r.intersects(clip):
                x0 = int((r.x0 - clip.x0) * z) - 1; y0 = int((r.y0 - clip.y0) * z) - 1; x1 = int((r.x1 - clip.x0) * z) + 1; y1 = int((r.y1 - clip.y0) * z) + 1
                A[max(0, y0):max(0, y1), max(0, x0):max(0, x1)] = 255
        rows = (A.min(axis=2) < _V267_LIGNAGE_ENCRE).mean(axis=1) >= _V267_LIGNAGE_PART
        n = 0; prev = False
        for v in rows:
            if v and not prev:
                n += 1
            prev = v
        if n >= _V267_LIGNAGE_N:
            return True
    if "dessins" not in m:
        m["dessins"] = page.get_drawings()
    n = 0
    for d in m["dessins"]:
        for it in d.get("items", []):
            if it[0] == "l":
                p, q = it[1], it[2]
                if abs(p.y - q.y) < _V267_TRAIT_DY and bb[1] <= p.y <= bb[3] and min(p.x, q.x) <= bb[0] and max(p.x, q.x) >= bb[2]:
                    n += 1
            elif it[0] == "re":
                r = it[1]
                if r.height < _V267_TRAIT_H and bb[1] <= (r.y0 + r.y1) / 2 <= bb[3] and r.x0 <= bb[0] and r.x1 >= bb[2]:
                    n += 1
    return n >= _V267_LIGNAGE_N


def _v267_cache_encre(page, bb):
    """Le cache d'un mot : sa boîte, étendue aux composantes d'encre (300 ppp, une hauteur de marge) qui la touchent sans toucher le bord."""
    k = _V267_CACHE_PPP / 72; h = bb[3] - bb[1]
    Z = fitz.Rect(bb[0] - h, bb[1] - h, bb[2] + h, bb[3] + h) & page.rect
    pm = page.get_pixmap(matrix=fitz.Matrix(k, k), clip=Z, alpha=False)
    A = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].min(axis=2) < _V267_CACHE_ENCRE
    H_, W_ = A.shape
    bx = [(bb[0] - Z.x0) * k, (bb[1] - Z.y0) * k, (bb[2] - Z.x0) * k, (bb[3] - Z.y0) * k]; out = list(bb)
    for x0, y0, x1, y1, _ar in _v267_boites(A):
        touche = not (x1 < bx[0] or x0 > bx[2] or y1 < bx[1] or y0 > bx[3])
        dedans = x0 >= 0 and y0 >= 0 and x1 < W_ - 1 and y1 < H_ - 1
        if touche and dedans:
            out = [min(out[0], Z.x0 + x0 / k), min(out[1], Z.y0 + y0 / k), max(out[2], Z.x0 + (x1 + 1) / k), max(out[3], Z.y0 + (y1 + 1) / k)]
    return out


_R7_SUITE_EM = 1.5  # R7 (rang 7, « (e) » de la 121) : écart au plus entre deux mots de la même ligne pour qu'ils se tiennent bout à bout (en corps) ; fixé avant les mesures


def _r7_suite(bb, texte, tous):
    """R7 — la « suite » d'un mot cursif listé : les mots du texte de la page qui se tiennent bout à bout sur sa ligne (écart d'un mot au suivant <= _R7_SUITE_EM corps, même
    ligne), le mot compris, écrits séparés par une espace ; None si la suite n'a qu'un mot. C'est ce que la page compare au texte que le modèle a écrit pour ne pas reposer deux
    fois une phrase (champ `lt` de l'entrée cursive, que la page ne lisait que pour les lignes). Un signe dessiné (« ⊥ ») n'est pas un mot : il peut rester entre deux mots."""
    W = next((w for w in tous if w["texte"] == texte and all(abs(x - y) < 0.2 for x, y in zip(w["bbox"], bb))), None)
    if W is None:
        return None
    r = fitz.Rect(W["bbox"]); t = W["taille"]; L = []
    for v in tous:
        if v is W or not v["horizontal"]:
            continue
        q = fitz.Rect(v["bbox"])
        if min(r.y1, q.y1) - max(r.y0, q.y0) < _V267_RECOUV * min(r.height, q.height):
            continue
        L.append(v)
    g = []; x = r.x0
    for v in sorted([v for v in L if v["bbox"][2] <= r.x0 + _V267_BORD_PT], key=lambda v: -v["bbox"][2]):
        if x - v["bbox"][2] <= _R7_SUITE_EM * t:
            g.append(v["texte"]); x = v["bbox"][0]
        else:
            break
    d = []; x = r.x1
    for v in sorted([v for v in L if v["bbox"][0] >= r.x1 - _V267_BORD_PT], key=lambda v: v["bbox"][0]):
        if v["bbox"][0] - x <= _R7_SUITE_EM * t:
            d.append(v["texte"]); x = v["bbox"][2]
        else:
            break
    mots = list(reversed(g)) + [W["texte"]] + d
    return " ".join(mots) if len(mots) >= 2 else None


def _v267_ajoute_d6(doc, images, cur, vb):

    """Les mots cursifs de D6 (voie b : dans une image de la famille qui n'est pas une légende ; voie a : isolé, non ponctué, à 6 mm de la
    famille ; a(iii) : sa figure la plus proche est une légende, il va à l'image de la famille à 6 mm d'elle), hors lignage, copies
    superposées réunies, chacun avec son cache ; ajoutés à voisins_b (le cache seul s'ajoute à une entrée qui est le même mot)."""
    par_page = _co266.defaultdict(list)
    for k, im in enumerate(images):
        par_page[im["page"]].append((k, im.get("figure") or "(sans figure)", tuple(im["cadre"])))
    leg = {}; E = []; memo = {}
    for w in cur["mots"]:
        pno = w["page"]; ip = par_page.get(pno, [])
        if pno not in leg:
            leg[pno] = _v267_legendes(ip, [x["bbox"] for x in cur["tous"][pno]])
        bb = w["bbox"]
        nl = [(r, f, c) for r, f, c in ip if _v267_centre_dans(bb, c) and r not in leg[pno]]
        if nl:
            fam = [t for t in nl if _V267_FAM2.match(t[1])]
            if not fam:
                continue
            r, fig, img = fam[0]; voie = "b"
        else:
            fam = [(_v267_ecart(bb, c), r, f, c) for r, f, c in ip if _V267_FAM2.match(f)]
            if not fam:
                continue
            d, r, f, c = min(fam, key=lambda t: t[0])
            if round(d, 1) > _V267_SEUIL_MM:
                continue
            if (not w["isole"] or (w["mot"] and w["mot"][-1] in _V267_PONCT_D6)) and not any(_v267_centre_dans(bb, c0) for r0, f0, c0 in ip):  # R4a (DEC-3362, rang 4)
                continue
            fig, img = f, c; voie = "a"
            if r in leg[pno]:
                q = [(_v267_ecart(c, c2), f2, c2) for r2, f2, c2 in ip if r2 not in leg[pno] and _V267_FAM2.match(f2) and r2 != r]
                q = [t for t in q if round(t[0], 1) <= _V267_SEUIL_MM]
                if q:
                    _, fig, img = min(q, key=lambda t: t[0]); voie = "a(iii)"
        if _v267_lignage(doc[pno - 1], bb, memo):
            continue
        E.append(dict(page=pno, texte=w["mot"], bbox=list(bb), figure=fig, voie=voie, lt=_r7_suite(bb, w["mot"], cur["tous"][pno])))  # R7
    U = []
    for e in E:
        for u in U:
            if u["page"] == e["page"] and u["texte"] == e["texte"] and all(abs(x - y) <= _V267_COPIES_PT for x, y in zip(u["bbox"], e["bbox"])):
                u["cache"] = [min(u["cache"][0], e["bbox"][0]), min(u["cache"][1], e["bbox"][1]), max(u["cache"][2], e["bbox"][2]), max(u["cache"][3], e["bbox"][3])]
                break
        else:
            U.append(dict(e, cache=list(e["bbox"])))
    for u in U:
        cache = [round(v, 2) for v in _v267_cache_encre(doc[u["page"] - 1], u["cache"])]
        Lf = vb.setdefault(u["figure"], [])
        dup = [v for v in Lf if v["texte"] == u["texte"] and all(abs(x - y) <= _V267_COPIES_PT for x, y in zip(v["bbox"], u["bbox"]))]
        if dup:
            dup[0]["cache"] = cache
        else:
            Lf.append(dict(texte=u["texte"], bbox=[round(v, 2) for v in u["bbox"]], sorte="cursif", ligne=None, lt=u.get("lt"), cache=cache, voie=u["voie"]))  # R7 : la suite du mot (None si un seul mot)
    return len(U)


# ═══ v2.69 — F034 S4b (les cases V/F sortent des images) et F033 W12h (la coupe d'un dessin de tableau) ; DEC-2813, DEC-2830 ; DEC-3271 E1 (a), E2 ═══
# Portage de agents273/A25_VF/serveur/a25s_hook.py (_a25_s4b), agents273/A25_VF/a25_cases.py (cases_vf, paires) et agents273/A24_001c/a24_variants_h.py
# (a24_couper, drapeaux de W12h : marge, ligne, encre, héritée). La porte E2 (sa réponse) : seule une case d'une paire V/F est une case de réponse ;
# une case seule (une lettre qui nomme une figure) reste dans son image. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_V269_VF_MIN = 3.0 * 72.0 / 25.4
_V269_VF_MAX = 15.0 * 72.0 / 25.4
_V269_VF_RATIO = 2.0
_V269_VF_LETTRES = ("V", "F")
_V269_W12H_G_MM = 0.5
_V269_W12H_MIETTE_PX = 17
_V269_W12H_EPS = 1e-6
_V269_W12H_LECT = {}


def _v269_vf_ferme(d):
    its = d.get("items") or []
    if not its:
        return False
    if len(its) == 1 and its[0][0] in ("re", "qu"):
        return True
    if any(it[0] not in ("l", "c", "re", "qu") for it in its):
        return False
    if d.get("closePath"):
        return True
    def pts(it):
        if it[0] == "l":
            return it[1], it[2]
        if it[0] == "c":
            return it[1], it[4]
        return None, None
    a, _ = pts(its[0]); _, b = pts(its[-1])
    if a is None or b is None:
        return False
    return abs(a.x - b.x) <= 0.5 and abs(a.y - b.y) <= 0.5


def _v269_vf_petit(r):
    w, h = r.width, r.height
    if not (_V269_VF_MIN <= w <= _V269_VF_MAX and _V269_VF_MIN <= h <= _V269_VF_MAX):
        return False
    return max(w, h) / max(1e-6, min(w, h)) <= _V269_VF_RATIO


def _v269_cases_vf(page):
    """Les cases de réponse V/F de la page : un tracé fermé de 3 à 15 mm (rapport ≤ 2), exactement un mot au centre, « V » ou « F », dans la boîte
    + 1 pt (une case dessinée deux fois comptée une fois) ; et, porte E2, qui appartient à une paire V/F (une V et une F voisines : même rangée ou même
    colonne, recouvrement ≥ 50 % du plus petit côté, écart ≤ 1,5 × le petit côté)."""
    mots = [w for w in page.get_text("words") if w[4].strip()]
    out, vus = [], []
    for d in page.get_drawings():
        r = d.get("rect")
        if r is None:
            continue
        r = fitz.Rect(r)
        if not _v269_vf_petit(r) or not _v269_vf_ferme(d):
            continue
        dedans = [w for w in mots if r.x0 <= (w[0] + w[2]) / 2 <= r.x1 and r.y0 <= (w[1] + w[3]) / 2 <= r.y1]
        if len(dedans) != 1:
            continue
        w = dedans[0]; t = w[4].replace(" ", "").strip()
        if t not in _V269_VF_LETTRES:
            continue
        big = fitz.Rect(r.x0 - 1, r.y0 - 1, r.x1 + 1, r.y1 + 1)
        if not (big.x0 <= w[0] and big.y0 <= w[1] and big.x1 >= w[2] and big.y1 >= w[3]):
            continue
        if any(max(abs(r.x0 - q[0]), abs(r.y0 - q[1]), abs(r.x1 - q[2]), abs(r.y1 - q[3])) <= 0.5 for q in vus):
            continue
        vus.append((r.x0, r.y0, r.x1, r.y1))
        out.append(dict(lettre=t, rect=[r.x0, r.y0, r.x1, r.y1]))
    paire = set()
    for i, a in enumerate(out):
        if a["lettre"] != "V":
            continue
        ra = fitz.Rect(a["rect"])
        for j, b in enumerate(out):
            if b["lettre"] != "F":
                continue
            rb = fitz.Rect(b["rect"]); petit = min(ra.width, ra.height, rb.width, rb.height)
            ov_y = min(ra.y1, rb.y1) - max(ra.y0, rb.y0); ov_x = min(ra.x1, rb.x1) - max(ra.x0, rb.x0)
            gx = max(rb.x0 - ra.x1, ra.x0 - rb.x1); gy = max(rb.y0 - ra.y1, ra.y0 - rb.y1)
            if (ov_y >= 0.5 * petit and gx <= 1.5 * petit) or (ov_x >= 0.5 * petit and gy <= 1.5 * petit):
                paire.add(i); paire.add(j)
    return [c for k, c in enumerate(out) if k in paire]


# ═══ v2.69 — Zk (F052 ; DEC-2858 « Zk seul », DEC-3276 R6) : « les images que S4b a retirées comptent encore pour le champ » ═══
# Sans lui, S4b, en sortant une case Vrai/Faux de son image, fait entrer sa phrase (lettre comprise) dans le champ des mots et la range sous la
# figure de l'exercice suivant (la 098). Portage d'A45 (agents274/A45_z5/a45_lib4.py, images_fantomes) et du journal de S4b du banc
# (agents273/A25_VF/serveur/a25s_hook.py) : toute image qui recouvre une case retenue est notée, retirée, recoupée ou laissée intacte. Seul le
# champ des mots voit ces images fantômes ; les images rendues, la page et le modèle ne changent pas. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
# ═══ v2.69 — F068 RL2 : LES LIGNES DE RANGÉE, LA SÉRIE (E12 et E13 de DEC-3271 ; DEC-2893) ═══
# Une phrase à trou à plus de 6 mm de toute image d'exercice (hors de B) est une ligne de la figure quand la figure est entière, faite d'au
# moins deux rangées, et que CHAQUE rangée porte exactement une phrase à trou à sa droite, à sa hauteur, sans rien entre elles, à 30 mm au
# plus (tout ou rien, E13) ; ses entrées (ce que B en poserait, plus `rangee`) s'ajoutent à voisins_b. Portage d'agents274/A61_bm
# (rangee_a61.classer R1-R5, rangees_de ; serie_a61.classer_serie, variante RL2 ; chaine_rangee_a61.etape_P2r) sans une règle changée.
# La pose (F1-t), la phrase de la demande et la pose au-delà de 6 mm sont la part de la page (E12). Le pourquoi et les mesures : JOURNAL
# BACKEND v2.69.
_V269_RL2_TOL_X = 3.0       # pt : un morceau « à gauche » peut dépasser de 3 pt le début de la ligne
_V269_RL2_G_MM = 30.0       # mm : la ligne à 30 mm au plus de son morceau
_V269_RL2_BORNE_B_MM = 6.0  # mm : la borne de B (DEC-1894, DEC-1898)


def _v269_rl2_complete(l):
    """Le bilan du trou d'une ligne (s2_champ_252.poses_ligne) : tous ses mots hors des images, et son trou présent."""
    el = l["elements"]
    tous_mots = all(not e["dans_image"] for e in el if e["sorte"] == "mot")
    trou = (any(e["sorte"] == "K" and e["dans_image"] for e in el) or any(e["sorte"] == "P1" for e in el)
            or any(e["sorte"] == "mot" and _vb_p2(e["texte"]) >= 4 for e in el)
            or any(e["sorte"] == "P3" and e["dans_image"] for e in el))
    return bool(tous_mots and trou)


def _v269_rl2_vov(a0, a1, b0, b1):
    return min(a1, b1) - max(a0, b0)


def _v269_rl2_rangees(morceaux):
    """Les rangées d'une figure, de haut en bas : un morceau entre dans la rangée du premier quand les intervalles verticaux se recouvrent
    d'au moins la moitié de la plus petite hauteur."""
    rows = []
    for m in sorted(morceaux, key=lambda m: ((m["cadre"][1] + m["cadre"][3]) / 2, m["cadre"][0])):
        for R in rows:
            f = R[0]["cadre"]
            h = min(f[3] - f[1], m["cadre"][3] - m["cadre"][1])
            if h > 0 and _v269_rl2_vov(f[1], f[3], m["cadre"][1], m["cadre"][3]) >= 0.5 * h:
                R.append(m)
                break
        else:
            rows.append([m])
    return rows


def _v269_rl2_morceau(l, ims_page, jetons):
    """Le morceau d'exercice auquel une phrase à trou se rattache, ou None : R1 à R5 (R0, la borne de B, n'est pas jouée ici, comme
    dans la série) — R1 aucun mot porté ; R2 la ligne complète ; R3 un morceau d'exercice à sa gauche, à sa hauteur ; R4 rien entre eux
    (ni morceau, ni jeton qui ne soit à elle) ; R5 à 30 mm au plus."""
    if l["statut"] != "aucun mot porté" or not _v269_rl2_complete(l):
        return None
    Lx0, Ly0, Lx1, Ly1 = l["bbox"]
    cy = (Ly0 + Ly1) / 2
    cands = [m for m in ims_page if _VB_RE_EX.match(m.get("figure") or "") and m["cadre"][2] <= Lx0 + _V269_RL2_TOL_X
             and m["cadre"][1] <= cy <= m["cadre"][3]]
    if not cands:
        return None
    N = min(cands, key=lambda m: max(0.0, Lx0 - m["cadre"][2]))
    gap = max(0.0, Lx0 - N["cadre"][2])
    x0b, x1b = N["cadre"][2], Lx0
    if x1b > x0b:
        for m in ims_page:
            if m is N:
                continue
            c = m["cadre"]
            if min(c[2], x1b) - max(c[0], x0b) > 0 and _v269_rl2_vov(c[1], c[3], Ly0, Ly1) > 0:
                return None
        mien = [e["bbox"] for e in l["elements"] if e["sorte"] in ("mot", "P1")]
        for t in jetons:
            tx, ty = (t[0] + t[2]) / 2, (t[1] + t[3]) / 2
            if x0b <= tx <= x1b and Ly0 <= ty <= Ly1:
                if any(all(abs(a - b) < 0.05 for a, b in zip(t[:4], mb)) for mb in mien):
                    continue
                return None
    if gap > _V269_RL2_G_MM * 72 / 25.4:
        return None
    return N


def _v269_rl2(doc, images, corrige_pages):
    """RL2 : {figure: [entrée, …]} à ajouter à voisins_b — les phrases à trou d'une série (chaque rangée de la figure en porte exactement
    une), à plus de 6 mm, au format de la part B (texte, bbox, sorte, ligne, lt) plus `rangee`."""
    ims = [dict(rang=i, index=i, page=im["page"], figure=im.get("figure"), morceau=im.get("morceau"), cadre=im["cadre"],
                entier_cadre=im.get("entier_cadre")) for i, im in enumerate(images)]
    lignes = _vb_lignes_doc(doc, ims, corrige_pages, borne=None)
    jetons, morceaux = {}, []
    for l in lignes:
        if l["page"] not in jetons:
            jetons[l["page"]] = [[float(j["bbox"].x0), float(j["bbox"].y0), float(j["bbox"].x1), float(j["bbox"].y1), j["texte"], bool(j["alnum"])]
                                 for j in _vb_jetons_page(doc[l["page"] - 1])]
        morceaux.append(_v269_rl2_morceau(l, [im for im in ims if im["page"] == l["page"]], jetons[l["page"]]))
    par = {}
    for i, N in enumerate(morceaux):
        if N is not None:
            par.setdefault((lignes[i]["page"], N["figure"]), []).append(i)
    rangee, serie = {}, set()
    for (page, F), idx in par.items():
        mF = [m for m in ims if m["page"] == page and m["figure"] == F]
        entiere = len(mF) >= 2 and any(m.get("entier_cadre") for m in mF)
        rows = _v269_rl2_rangees(mF)
        for i in idx:
            rangee[i] = None
            for k, Rw in enumerate(rows):
                if any(m.get("morceau") == morceaux[i].get("morceau") for m in Rw):
                    rangee[i] = k
        cpt = _co266.Counter(rangee[i] for i in idx)
        if entiere and len(rows) >= 2 and len(cpt) == len(rows) and all(c == 1 for c in cpt.values()):
            serie.update(idx)
    champ = {}
    for i, l in enumerate(lignes):
        d = l.get("dist_ex_mm")
        if i not in serie or (d is not None and d <= _V269_RL2_BORNE_B_MM):
            continue
        F = morceaux[i]["figure"]
        for q in _vb_poses_ligne(l):
            k = q["ligne"]
            champ.setdefault(F, []).append(dict(texte=q["texte"], bbox=[round(v, 2) for v in q["bbox"]], sorte=q["sorte"],
                                                ligne="%d|%s" % (k[0], "|".join("%.2f" % v for v in k[1])), lt=l["texte"],
                                                rangee=rangee[i]))
    return champ


_V269_S4B_JOURNAL = {}


def _v269_s4b_journal(page):
    """Le journal de S4b du document de la page : les images qui recouvrent une case V/F retenue (page, figure, cadre arrondi à 0,1)."""
    J = _V269_S4B_JOURNAL
    if J.get("parent") is not page.parent:
        J.clear(); J["parent"] = page.parent; J["liste"] = []
    return J["liste"]


def _v269_zk_fantomes(doc):
    """Les images fantômes du journal de S4b pour ce document (page, figure, morceau 0, cadre, sans image entière), données au seul champ des mots."""
    J = _V269_S4B_JOURNAL
    if J.get("parent") is not doc:
        return []
    return [dict(page=j["page"], figure=j.get("figure"), morceau=0, cadre=[float(v) for v in j["cadre"]], entier_cadre=None) for j in J["liste"]]


def _v269_s4b(page, ims, idx_fin):
    """S4b : pour chaque image dont le cadre recouvre à 5 % au moins une case de réponse V/F, l'encre de l'image sans les cases (boîte + 2 px) passe par
    _v254_couper, l'encre des cases pour encre étrangère ; les rectangles rendus remplacent l'image ; une image qui ne porte que des cases est retirée ;
    rien à couper : intacte. Puis _v254_renumerote. Rend (images, index suivant, voies des images neuves)."""
    if not ims:
        return ims, idx_fin, []
    pno = page.number + 1
    cases = [fitz.Rect(c["rect"]) for c in _v269_cases_vf(page)]
    if not cases:
        return ims, idx_fin, []
    K = _V247_K
    base = ims[0]["index"]
    R1 = _v254_etat()["r1"]
    cle = R1["cle"] if R1.get("doc") is page.parent else {}
    out, change, voies = [], False, []
    for im in ims:
        R = fitz.Rect(im["cadre"])
        dedans = [c for c in cases if (R & c).get_area() >= 0.05 * c.get_area()]
        if not dedans:
            out.append(im)
            continue
        # v2.69 — Zk : le journal de S4b, pour le champ des mots (voir _v269_zk_fantomes)
        _v269_s4b_journal(page).append(dict(page=pno, figure=im.get("figure"), cadre=[round(v, 1) for v in im["cadre"]]))
        Rp = R & page.rect
        pm = page.get_pixmap(matrix=fitz.Matrix(K, K), clip=Rp, alpha=False)
        arr = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
        E = (255 - arr.min(axis=2)) >= _V247_ECART
        H, Wd = E.shape
        mask = _np247.zeros_like(E)
        for c in dedans:
            q = fitz.Rect(c) & Rp
            mask[max(0, int((q.y0 - Rp.y0) * K) - 2):min(H, int(_np247.ceil((q.y1 - Rp.y0) * K)) + 2),
                 max(0, int((q.x0 - Rp.x0) * K) - 2):min(Wd, int(_np247.ceil((q.x1 - Rp.x0) * K)) + 2)] = True
        A = E & ~mask
        O = E & mask
        if not A.any():
            change = True
            continue
        feuilles = _v254_couper(A, O, 10)
        rects = [fitz.Rect(Rp.x0 + a / K, Rp.y0 + b / K, Rp.x0 + c / K, Rp.y0 + d / K) for (a, b, c, d), n in feuilles]
        if len(rects) == 1 and max(abs(rects[0][i] - R[i]) for i in range(4)) < 0.01:
            out.append(im)
            continue
        change = True
        tag = cle.get((pno, tuple(round(v, 1) for v in im["cadre"])))
        for rr in rects:
            rr = rr & page.rect
            if rr.is_empty:
                continue
            pix = page.get_pixmap(clip=rr, matrix=fitz.Matrix(2, 2))
            if pix.width < 1 or pix.height < 1:
                continue
            png = pix.tobytes("png")
            extra, voie = _v238_champs(page, rr)
            neuf = dict(im)
            neuf.pop("entier", None)
            neuf.pop("entier_cadre", None)
            neuf.update({"cadre": [rr.x0, rr.y0, rr.x1, rr.y1], "sceau": _hl.sha256(png).hexdigest()[:12],
                         "data": "data:image/png;base64," + base64.b64encode(png).decode(), "w": pix.width, "h": pix.height,
                         "w_mm": round(rr.width * 25.4 / 72, 1), "h_mm": round(rr.height * 25.4 / 72, 1), **extra})
            out.append(neuf)
            voies.append(voie)
            if tag is not None:
                cle[(pno, tuple(round(v, 1) for v in neuf["cadre"]))] = tag
    if not change:
        return ims, idx_fin, []
    out, idx = _v254_renumerote(page, out, idx_fin, base=base)
    return out, idx, voies


def _v269_w12h_lecture(page, K):
    """Les caractères non blancs (boîte en pixels de page) et les lignes (boîte en pixels) de la page, une fois par page et par document."""
    C = _V269_W12H_LECT
    if C.get("parent") is not page.parent:
        C.clear(); C["parent"] = page.parent
    k = page.number
    if k in C:
        return C[k]
    raw = page.get_text("rawdict")
    chars, lignes = [], []
    for b in raw["blocks"]:
        if b.get("type") != 0:
            continue
        for l in b["lines"]:
            lb = l["bbox"]; lignes.append((lb[0] * K, lb[1] * K, lb[2] * K, lb[3] * K))
            for s in l["spans"]:
                for c in s["chars"]:
                    if c["c"].strip() == "":
                        continue
                    cb = c["bbox"]; chars.append((cb[0] * K, cb[1] * K, cb[2] * K, cb[3] * K, c["c"]))
    C[k] = (chars, lignes)
    return C[k]


# ═══ v2.69 — E1 (b) : LE SIGNE ⊥ ÉCRIT COMME UN CARACTÈRE (DEC-3271 E1 ; DEC-2803 : « qu'aucun signe d'une phrase ne devienne une image ») ═══
# Un ⊥ dessiné dans une ligne de texte est écrit « ⊥ » (U+22A5) dans le texte de sa ligne, et effacé des images qui le contiennent. Reconnaissance
# de FORME (un dessin), provisoire par construction : un trait vertical qui tombe au milieu d'un trait horizontal d'au plus 20 pt ; elle ne lit que
# les tracés, aucun mot, aucun éditeur. « Dans une ligne » : les voisins de B-N2 (_v269_perp_voisins, qui appelle _v269_bn2_voisins : une seule
# définition). Seul un ⊥ écrit dans le texte est effacé des images (_v269_perps_ecrits) : un ⊥ non écrit reste dans son image. La police de la
# feuille qui porte le signe est l'affaire de la page. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_V269_PERP_LMAX = 20.0     # pt : le plus long trait d'un ⊥
_V269_PERP_PIED = 1.0      # pt : le pied du trait vertical touche le trait horizontal
_V269_PERP_MILIEU = 0.2    # part de la longueur du trait horizontal : le trait vertical tombe en son milieu
_V269_PERP_CACHE = {}


def _v269_signe_de_ligne(page, q):
    """La règle « ligne » de W12h : la boîte q (en pixels de page, × _V247_K) tient entre les bords d'une ligne de texte de la page, à sa
    hauteur (au plus 1,25 fois la ligne). Sortie de _v269_w12h_couper sans changement ; le ⊥ ne s'en sert plus (_v269_perp_voisins)."""
    K = _V247_K
    _, lignes = _v269_w12h_lecture(page, K); X0, Y0, X1, Y1 = q; cy = 0.5 * (Y0 + Y1)
    for (a, b, c, d) in lignes:
        h = d - b
        if h <= 0:
            continue
        if a < X0 and X1 < c and b <= cy <= d and (Y1 - Y0) <= 1.25 * h:
            return True
    return False


def _v269_perp_voisins(mots, bb):
    """Les deux mots qui encadrent un ⊥ sur sa ligne : les voisins de B-N2 (_v269_bn2_voisins, la boîte du ⊥ tenant lieu de marque), chacun à
    au plus 3 hauteurs de la boîte ; (mot de gauche, mot de droite), ou None si le ⊥ n'est pas dans une ligne de texte."""
    ws = list(mots) + [(bb.x0, bb.y0, bb.x1, bb.y1, "⊥")]
    g, d = _v269_bn2_voisins(ws, len(ws) - 1)
    h = bb.y1 - bb.y0
    if g is None or d is None or bb.x0 - ws[g][2] > 3 * h or ws[d][0] - bb.x1 > 3 * h:
        return None
    return ws[g][4], ws[d][4]


def _v269_perps_ecrits(page):
    """Les ⊥ de la page que _v269_perp_dans_le_texte a écrits dans le texte : les seuls à effacer des images."""
    C = _V269_PERP_CACHE
    if C.get("parent") is not page.parent:
        return []
    return C.get(("ecrits", page.number), [])


def _v269_perps(page):
    """Les ⊥ dessinés dans une ligne de texte de la page : [boîte en points] (une fois par page et par document)."""
    C = _V269_PERP_CACHE
    if C.get("parent") is not page.parent:
        C.clear(); C["parent"] = page.parent
    if page.number in C:
        return C[page.number]
    segs = []
    for d in page.get_drawings():
        for it in d.get("items") or []:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.x - b.x) < 0.6 and abs(a.y - b.y) > 2:
                    segs.append(("v", min(a.y, b.y), max(a.y, b.y), 0.5 * (a.x + b.x)))
                elif abs(a.y - b.y) < 0.6 and abs(a.x - b.x) > 2:
                    segs.append(("h", min(a.x, b.x), max(a.x, b.x), 0.5 * (a.y + b.y)))
            elif it[0] == "re":
                r = it[1]
                if r.width < 1.6 and r.height > 2:
                    segs.append(("v", r.y0, r.y1, r.x0 + r.width / 2))
                elif r.height < 1.6 and r.width > 2:
                    segs.append(("h", r.x0, r.x1, r.y0 + r.height / 2))
    H = [s for s in segs if s[0] == "h" and s[2] - s[1] <= _V269_PERP_LMAX]
    V = [s for s in segs if s[0] == "v" and s[2] - s[1] <= _V269_PERP_LMAX]
    mots = [tuple(w[:5]) for w in page.get_text("words")]
    out, vus = [], set()
    for h in H:
        L = h[2] - h[1]
        for v in V:
            if abs(v[3] - 0.5 * (h[1] + h[2])) <= _V269_PERP_MILIEU * L and abs(v[2] - h[3]) <= _V269_PERP_PIED and 0.5 * L <= v[2] - v[1] <= 2 * L:
                bb = fitz.Rect(h[1], v[1], h[2], h[3])
                k = tuple(round(x, 1) for x in bb)
                if k in vus:
                    continue
                if _v269_perp_voisins(mots, bb) is not None:
                    vus.add(k); out.append(bb)
    C[page.number] = out
    return out


def _v269_perp_dans_le_texte(page, texte):
    """Écrit « ⊥ » dans le texte d'une page tel qu'il sort (ancres posées), à la place de chaque ⊥ de la page : entre les deux mots qui
    l'encadrent sur sa ligne (_v269_perp_voisins), là où le texte les sépare par des blancs d'une même ligne ; s'il les sépare par un
    retour à la ligne, « ⊥ » termine la ligne de gauche (le nombre de lignes ne change pas). Si ce couple de mots se lit dans le texte un
    autre nombre de fois que le nombre de ⊥ qui le portent, rien n'est écrit pour lui (pas de place devinée). Les ⊥ écrits sont notés
    pour la page (_v269_perps_ecrits) : eux seuls sont effacés des images. Une page sans ⊥ : le texte est rendu tel quel. Rend (texte,
    nombre de ⊥ écrits)."""
    P = _v269_perps(page)
    if not P:
        return texte, 0
    mots = [tuple(w[:5]) for w in page.get_text("words")]
    couples = _co266.defaultdict(list)
    for bb in P:
        couples[_v269_perp_voisins(mots, bb)].append(bb)
    ecrits = []
    for (a, b), bbs in couples.items():
        trouves = list(re.finditer(re.escape(a) + r"([ \t]*\n[ \t]*|[ \t]+)" + re.escape(b), texte))
        if len(trouves) != len(bbs):
            continue
        for m in reversed(trouves):
            sep = m.group(1)
            texte = texte[:m.start(1)] + (" ⊥\n" + sep.split("\n", 1)[1] if "\n" in sep else " ⊥ ") + texte[m.end(1):]
        ecrits.extend(bbs)
    _V269_PERP_CACHE[("ecrits", page.number)] = ecrits
    return texte, len(ecrits)


def _v269_w12h_couper(A, O, prof_max=10, tab=False, page=None, bx0=0, by0=0):
    """W12h : pour un dessin de tableau, la coupe de W12 (bord 0,5 mm, miette 3 mm) et quatre règles — une feuille s'étend aux caractères dont toute
    l'encre est en elle (marge) ; une feuille qui tient entre les mots d'une ligne de texte est écartée (ligne) ; une décomposition est refusée si une
    lettre a de l'encre dans et hors d'une feuille (encre) ; deux boîtes qui se partagent une lettre sont réunies et la décomposition refaite (héritée).
    Hors tableau : _v254_couper."""
    feu = _v254_couper(A, O, prof_max)
    if not tab:
        return feu
    K = _V247_K; EPS = _V269_W12H_EPS
    GP = int(_m266.ceil(_V269_W12H_G_MM * 72.0 / 25.4 * K)); MINPX = _V269_W12H_MIETTE_PX
    np = _np247
    IO = np.zeros((O.shape[0] + 1, O.shape[1] + 1), dtype=np.int64); IO[1:, 1:] = O.astype(np.int64).cumsum(0).cumsum(1)
    cnt = lambda b: int(IO[b[3], b[2]] - IO[b[1], b[2]] - IO[b[3], b[0]] + IO[b[1], b[0]])
    def serre(x0, y0, x1, y1):
        s = A[y0:y1, x0:x1]; r = s.any(axis=1); c = s.any(axis=0)
        if not r.any():
            return None
        ya = int(np.argmax(r)); yb = len(r) - int(np.argmax(r[::-1])); xa = int(np.argmax(c)); xb = len(c) - int(np.argmax(c[::-1]))
        return (x0 + xa, y0 + ya, x0 + xb, y0 + yb)
    def plus_large_vide(b):
        x0, y0, x1, y1 = b; s = A[y0:y1, x0:x1]; best = None
        for sens, occ in ((0, s.any(axis=1)), (1, s.any(axis=0))):
            i = 0; n = len(occ)
            while i < n:
                if occ[i]:
                    i += 1
                    continue
                j = i
                while j < n and not occ[j]:
                    j += 1
                if i > 0 and j < n and (j - i) >= GP and (best is None or (j - i) > best[0]):
                    best = (j - i, sens, i, j)
                i = j
        return best
    def decomp(b, out):
        b = serre(*b)
        if b is None:
            return
        x0, y0, x1, y1 = b
        for (a, bb, c, d), n in _v254_couper(A[y0:y1, x0:x1], O[y0:y1, x0:x1], prof_max):
            feuille((x0 + a, y0 + bb, x0 + c, y0 + d), out)
    def feuille(b, out):
        g = plus_large_vide(b)
        if g is None:
            out.append((b, cnt(b)))
            return
        L, sens, i, j = g; x0, y0, x1, y1 = b
        if sens == 0:
            t, u = serre(x0, y0, x1, y0 + i), serre(x0, y0 + j, x1, y1)
        else:
            t, u = serre(x0, y0, x0 + i, y1), serre(x0 + j, y0, x1, y1)
        if t is None or u is None:
            out.append((b, cnt(b)))
            return
        decomp(t, out); decomp(u, out)
    vp = lambda q: (q[0] + bx0, q[1] + by0, q[2] + bx0, q[3] + by0)
    def let_page():
        return _v254_lettres_page(page)
    def ink_of(c, M):
        ia, ib, ic, id_ = int(_m266.floor(c[0] + EPS)), int(_m266.floor(c[1] + EPS)), int(_m266.ceil(c[2] - EPS)), int(_m266.ceil(c[3] - EPS))
        ia, ib = max(ia, 0), max(ib, 0); ic, id_ = min(ic, M.shape[1]), min(id_, M.shape[0])
        if ic <= ia or id_ <= ib:
            return None, (ia, ib, ic, id_)
        yy, xx = np.nonzero(M[ib:id_, ia:ic])
        return (xx + ia, yy + ib), (ia, ib, ic, id_)
    def dans(box, xx, yy):
        return (xx >= box[0]) & (xx < box[2]) & (yy >= box[1]) & (yy < box[3])
    def encre_traversee(q, bq):
        M = let_page(); chars, _ = _v269_w12h_lecture(page, K); X0, Y0, X1, Y1 = q
        for ch_ in chars:
            a, b, c, d, ch = ch_
            if min(c, X1) - max(a, X0) <= EPS or min(d, Y1) - max(b, Y0) <= EPS:
                continue
            ink, _ = ink_of(ch_, M)
            if ink is None or not len(ink[0]):
                continue
            xx, yy = ink; di = dans(q, xx, yy)
            if not (di.any() and (~di).any()):
                continue
            db = dans(bq, xx, yy)
            if db.any() and (~db).any():
                continue
            return True
        return False
    def signe_de_ligne(q):
        return _v269_signe_de_ligne(page, q)
    def etendre(boxes, bq):
        M = let_page(); chars, _ = _v269_w12h_lecture(page, K); out = [list(q) for q in boxes]
        infos = []
        for ch_ in chars:
            ink, ib = ink_of(ch_, M)
            if ink is None or not len(ink[0]):
                continue
            if not (ib[0] >= bq[0] and ib[1] >= bq[1] and ib[2] <= bq[2] and ib[3] <= bq[3]):
                continue
            infos.append((ink, ib))
        for _ in range(3):
            chg = False
            for k, q in enumerate(out):
                for ink, ib in infos:
                    if not dans(q, ink[0], ink[1]).all():
                        continue
                    if ib[0] >= q[0] and ib[1] >= q[1] and ib[2] <= q[2] and ib[3] <= q[3]:
                        continue
                    nq = [min(q[0], ib[0]), min(q[1], ib[1]), max(q[2], ib[2]), max(q[3], ib[3])]
                    if any(j != k and min(nq[2], o[2]) - max(nq[0], o[0]) > 0 and min(nq[3], o[3]) - max(nq[1], o[1]) > 0 for j, o in enumerate(out)):
                        continue
                    out[k] = q = nq; chg = True
            if not chg:
                break
        return out
    def traite(b, n):
        sub = []; decomp(b, sub)
        ok = bool(sub) and sum(c for _, c in sub) < n and all(min(q[2] - q[0], q[3] - q[1]) >= MINPX for q, _ in sub)
        kept = list(sub)
        if ok:
            bq = vp(b)
            for q, _ in sub:
                if encre_traversee(vp(q), bq):
                    ok = False
                    break
            if ok:
                kept = [(q, c) for q, c in kept if not signe_de_ligne(vp(q))]
            if ok and kept:
                boxes = [vp(q) for q, _ in kept]; nb = etendre(boxes, bq); new = []
                for (q, c), o, e in zip(kept, boxes, nb):
                    if list(o) != list(e):
                        qq = (e[0] - bx0, e[1] - by0, e[2] - bx0, e[3] - by0); new.append((qq, cnt(qq)))
                    else:
                        new.append((q, c))
                kept = new
        return ok, kept
    def partenaires(i, b):
        M = let_page(); chars, _ = _v269_w12h_lecture(page, K); bq = vp(b); out = set()
        for ch_ in chars:
            ink, _ = ink_of(ch_, M)
            if ink is None or not len(ink[0]):
                continue
            xx, yy = ink; db = dans(bq, xx, yy)
            if not (db.any() and (~db).any()):
                continue
            for j, (b2, n2) in enumerate(feu):
                if j != i and dans(vp(b2), xx, yy).any():
                    out.add(j)
        return out
    par_idx = {}; consomme = set()
    for idx, (b, n) in enumerate(feu):
        if idx in consomme:
            continue
        if n == 0:
            par_idx[idx] = [(b, n)]
            continue
        fait = False
        P = partenaires(idx, b) - consomme
        if P:
            boxes = [b] + [feu[j][0] for j in sorted(P)]
            U = (min(q[0] for q in boxes), min(q[1] for q in boxes), max(q[2] for q in boxes), max(q[3] for q in boxes))
            libre = not any(j != idx and j not in P and min(U[2], feu[j][0][2]) - max(U[0], feu[j][0][0]) > 0 and min(U[3], feu[j][0][3]) - max(U[1], feu[j][0][1]) > 0 for j in range(len(feu)))
            if libre:
                ok_u, kept_u = traite(U, cnt(U))
                if ok_u:
                    for j in P:
                        par_idx.pop(j, None)
                    consomme |= P; par_idx[idx] = kept_u; fait = True
        if fait:
            continue
        ok, kept = traite(b, n)
        par_idx[idx] = kept if ok else [(b, n)]
    res = []
    for idx in sorted(par_idx):
        res += par_idx[idx]
    return res


# ═══ v2.69 — G2 : L'EFFACEMENT SOUS LE RECTANGLE DES MOTS CURSIFS (F036 Q1 + D6-0i + S ; DEC-3273) ═══
# Portage de outils260/fabrique_rc_260.py (la boucle d'effacement), de ses changements par agents270/A16_094titre/a16_formes.py (encre < 255 ;
# le morceau contient le cache aussi en hauteur), de outils270/e_270.py (d60i) et de agents272/A20_coeur/a20_formes.py (Q1 : COND, q1_keep,
# letters_of, raw_in_image) ; outils260/ra_base_260.py (rendu_seul, dilate : un disque par décalages, tel quel). scipy.ndimage est remplacé par
# l'étiquetage du serveur et des filtres maximum en numpy. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_V269_G2_X = 0.5      # Q1 : part minimale des pixels du tracé brut d'une rangée ou colonne montrés par l'image
_V269_G2_FL = 0.5     # une composante d'encre sombre est « lettre » si moins de FL de ses pixels sont dans le masque dilaté
_V269_G2_DKV = 200    # « sombre » : min(R, V, B) < 200
_V269_G2_TOLB = 5     # tolérance ± 2 rangées ou colonnes


def _v269_g2_rendu_seul(page, k, memo):
    """Les tracés de la page redessinés seuls (contour, couleur, épaisseur) sur une page blanche, rendus à l'échelle k : pixels < 250."""
    cle = (page.number, k)
    if cle in memo:
        return memo[cle]
    A = _v269_rendu_seul(page, [d for d in page.get_drawings() if d.get("rect") is not None], k)
    memo[cle] = A
    return A


def _v269_g2_dilate(M, r):
    """Dilatation par un disque de rayon r, par décalages circulaires (le geste du banc, bords repliés compris)."""
    out = M.copy()
    for dy in range(-r, r + 1):
        for dx in range(-r, r + 1):
            if dx * dx + dy * dy > r * r:
                continue
            out |= _np247.roll(_np247.roll(M, dy, axis=0), dx, axis=1)
    return out


def _v269_g2_max3(M):
    """Filtre maximum 3 × 3 sur un masque (les bords ne voient que l'intérieur)."""
    P = _np247.pad(M, 1, mode="constant", constant_values=False); H, W = M.shape; out = _np247.zeros_like(M)
    for dy in range(3):
        for dx in range(3):
            out |= P[dy:dy + H, dx:dx + W]
    return out


def _v269_g2_lettres(Dk, sub):
    """Les pixels sombres des composantes (8-connexes, sur toute l'image) dont moins de FL des pixels sont dans le masque dilaté."""
    lab = _v247_etiqueter(Dk, huit=True)
    if not lab.any():
        return _np247.zeros(Dk.shape, bool)
    ids, inv = _np247.unique(lab, return_inverse=True)
    inv = inv.reshape(lab.shape)
    taille = _np247.bincount(inv.ravel(), minlength=len(ids)).astype(float)
    dedans = _np247.bincount(inv.ravel(), weights=sub.ravel().astype(float), minlength=len(ids))
    frac = dedans / _np247.maximum(taille, 1.0)
    frac[ids == 0] = 1.0
    return Dk & (frac[inv] < _V269_G2_FL)


def _v269_g2_q1(sub, Rw, ctx, im, cle, box):
    """Q1 : dans la boîte, le masque n'est gardé que là où l'image montre le trait (rangées ou colonnes montrées, ± 2) ; ailleurs il est ôté."""
    x0, y0, x1, y1 = box; key = (id(im), cle)
    if key not in ctx["LET"]:
        Dk = ctx["ORIG"][key].min(axis=2) < _V269_G2_DKV; L = _v269_g2_lettres(Dk, sub)
        ctx["LET"][key] = (L, _v269_g2_max3(Dk & ~L))
    L, pres = ctx["LET"][key]
    Rb = Rw[y0:y1, x0:x1] & ~L[y0:y1, x0:x1]; Pb = pres[y0:y1, x0:x1]
    rden = Rb.sum(axis=1); rnum = (Rb & Pb).sum(axis=1); cden = Rb.sum(axis=0); cnum = (Rb & Pb).sum(axis=0)
    rst = _np247.where(rden > 0, _np247.where(rnum >= _V269_G2_X * rden, 1, -1), 0)
    cst = _np247.where(cden > 0, _np247.where(cnum >= _V269_G2_X * cden, 1, -1), 0)
    def etale(st, v):
        r = _V269_G2_TOLB // 2
        return _np247.convolve((st == v).astype(int), _np247.ones(_V269_G2_TOLB, dtype=int), mode="full")[r:r + len(st)] > 0
    rpos, rneg, cpos, cneg = etale(rst, 1), etale(rst, -1), etale(cst, 1), etale(cst, -1)
    montre = rpos[:, None] | cpos[None, :]; absent = ~montre & (rneg[:, None] | cneg[None, :])
    return ~absent


def _v269_g2_cond(sub, im, cle, k, org, page, box, ctx):
    x0, y0, x1, y1 = box
    if x1 <= x0 or y1 <= y0:
        return sub
    sb = sub[y0:y1, x0:x1].copy()
    if not int(sb.sum()):
        return sub
    H, W = sub.shape; Mp = _v269_g2_rendu_seul(page, k, ctx["RAW"]); Rw = _np247.zeros((H, W), bool)
    ys0, xs0 = max(0, org[1]), max(0, org[0]); ys1, xs1 = min(Mp.shape[0], org[1] + H), min(Mp.shape[1], org[0] + W)
    Rw[ys0 - org[1]:ys1 - org[1], xs0 - org[0]:xs1 - org[0]] = Mp[ys0:ys1, xs0:xs1]
    keep = _v269_g2_q1(sub, Rw, ctx, im, cle, (x0, y0, x1, y1))
    out = sub.copy(); out[y0:y1, x0:x1] = sb & keep
    return out


# ═══ FB8 — L'EFFACEMENT DES MOTS CURSIFS RECONNAÎT LE MOT QUE LE PDF ÉCRIT EN DEUX MORCEAUX (A47_z4 ; DEC-2861, DEC-3311 ; ligne 3.6 du plan) ═══
# Portage de agents274/A47_z4/scripts/a47_lib.py (aire, inter, aire_union, reconnu_FB8) : les corps à l'identique, les noms préfixés _a47_.
# Le défaut : le PDF écrit « onze » en deux morceaux (« onz » en 12 pt, « e » en 14 pt : deux spans). Le champ des mots cursifs les prend un à un ; G2 ne
# laisse effacer un mot de la page que si son coin haut-gauche est celui d'une entrée : « onze » n'est celui d'aucune des deux, il est protégé comme un
# mot voisin et reste dans l'image (0,5 % de son encre effacée au banc).
# La forme : un mot de la page n'est pas protégé si l'union des boîtes des entrées cursives de sa figure couvre au moins 80 % de sa boîte (seuil du banc,
# écrit avant sa mesure). C'est une règle de forme (géométrique) : elle ne lit ni le texte ni la police. Elle s'ajoute à la règle du coin, elle ne la remplace pas.
# Mesurée au banc sur les 129 (une seule famille de PDF porte des mots cursifs) : agents274/A47_z4/RAPPORT_A47.md ; son effet sur D6 complet :
# agents283/A144_Ce2/RAPPORT_A144.md. Les épreuves de ce portage : agentsCODE/PS1/RAPPORT_PS1.md.
def _a47_aire(b):
    return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])


def _a47_inter(a, b):
    return (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))


def _a47_aire_union(rects):
    """Surface de l'union de rectangles (compression de coordonnées, sans double compte)."""
    rects = [r for r in rects if _a47_aire(r) > 0]
    if not rects:
        return 0.0
    xs = sorted({r[0] for r in rects} | {r[2] for r in rects})
    tot = 0.0
    for a, b in zip(xs, xs[1:]):
        iv = sorted((r[1], r[3]) for r in rects if r[0] <= a and r[2] >= b)
        cur = None
        h = 0.0
        for y0, y1 in iv:
            if cur is None:
                cur = [y0, y1]
            elif y0 <= cur[1]:
                cur[1] = max(cur[1], y1)
            else:
                h += cur[1] - cur[0]
                cur = [y0, y1]
        if cur:
            h += cur[1] - cur[0]
        tot += h * (b - a)
    return tot


def _a47_reconnu_fb8(w, curs, seuil=0.8):
    """FB8 : le mot w de la page (x0, y0, x1, y1, texte, …) est reconnu comme mot cursif si l'union des boîtes des entrées cursives curs couvre au moins seuil de sa boîte."""
    W = tuple(w[:4])
    A = _a47_aire(W)
    return A > 0 and _a47_aire_union([_a47_inter(W, e["bbox"]) for e in curs]) >= seuil * A


def _r7_fond_zone(A, Z):
    """R7 (S11, FB8 « à revoir ») : la couleur qui remplace l'encre d'une zone effacée par G2 : le fond de la zone, c'est-à-dire la couleur claire (canal minimal >= 180) la plus
    fréquente parmi ses pixels si elle n'est pas blanche (un canal < 250) et couvre au moins la moitié des pixels clairs ; sinon 255 (comportement d'avant : blanc). Un fond de
    couleur n'est plus blanchi avec le mot ; un fond blanc ne change pas."""
    P = A[Z][:, :3]
    if len(P) == 0:
        return 255
    clair = P[P.min(axis=1) >= 180]
    if len(clair) == 0:
        return 255
    cle = (clair[:, 0].astype(_np247.int64) << 16) | (clair[:, 1].astype(_np247.int64) << 8) | clair[:, 2].astype(_np247.int64)
    val, nb = _np247.unique(cle, return_counts=True); i = int(nb.argmax()); c = int(val[i])
    rgb = ((c >> 16) & 255, (c >> 8) & 255, c & 255)
    if min(rgb) >= 250 or nb[i] < 0.5 * len(clair):
        return 255
    return _np247.array(rgb, dtype=A.dtype)


def _v269_g2_efface(images, vb, doc, ctx):

    """La boucle du banc (fabrique_rc_260, avec les changements de A16 et la condition Q1) : pour chaque mot cursif à cache, l'image posée (l'entière
    de la figure, sinon le morceau qui contient le cache en largeur et en hauteur), la zone du cache moins le masque des tracés gardé par Q1 et moins
    les autres mots ; l'encre y est mise à blanc ; le cache est retiré."""
    for f, LL in (vb or {}).items():
        for v in LL:
            if v.get("sorte") != "cursif" or not v.get("cache"):
                continue
            ims = [x for x in images if x.get("figure") == f and x.get("entier_feuille")]; cle = "entier_feuille"
            if not ims:
                ims = [x for x in images if x.get("figure") == f and x.get("data_feuille") and x["cadre"][0] <= v["cache"][0] + 0.5 and v["cache"][2] <= x["cadre"][2] + 0.5
                       and x["cadre"][1] <= v["cache"][1] + 0.5 and v["cache"][3] <= x["cadre"][3] + 0.5]; cle = "data_feuille"
            if not ims:
                continue
            im = ims[0]; c = im["entier_cadre"] if cle == "entier_feuille" else im["cadre"]
            A = _np247.array(_v260_rgb(im[cle]), copy=True); k = A.shape[1] / (c[2] - c[0])
            org = (int(round(c[0] * k)), int(round(c[1] * k))) if cle == "data_feuille" else ((fitz.Rect(c) * fitz.Matrix(k, k)).irect.x0, (fitz.Rect(c) * fitz.Matrix(k, k)).irect.y0)
            pg = doc[im["page"] - 1]; Mp = _v269_g2_dilate(_v269_g2_rendu_seul(pg, k, ctx["RAW"]), 2)
            k_ = v["cache"]; x0 = max(0, int(_m266.floor(k_[0] * k)) - org[0]); y0 = max(0, int(_m266.floor(k_[1] * k)) - org[1])
            x1 = min(A.shape[1], int(_m266.ceil(k_[2] * k)) - org[0]); y1 = min(A.shape[0], int(_m266.ceil(k_[3] * k)) - org[1])
            Z = _np247.zeros(A.shape[:2], bool); Z[y0:y1, x0:x1] = True
            sub = _np247.zeros(A.shape[:2], bool); ys0, xs0 = max(0, org[1]), max(0, org[0]); ys1, xs1 = min(Mp.shape[0], org[1] + A.shape[0]), min(Mp.shape[1], org[0] + A.shape[1])
            sub[ys0 - org[1]:ys1 - org[1], xs0 - org[0]:xs1 - org[0]] = Mp[ys0:ys1, xs0:xs1]
            sub = _v269_g2_cond(sub, im, cle, k, org, pg, (x0, y0, x1, y1), ctx); Z &= ~sub
            curs = [w for w in LL if w.get("sorte") == "cursif"]
            for w in pg.get_text("words"):
                if any(abs(w[0] - cc["bbox"][0]) < 1 and abs(w[1] - cc["bbox"][1]) < 1 for cc in curs) or _a47_reconnu_fb8(w, curs):
                    continue
                wx0 = int(w[0] * k) - org[0] - 1; wy0 = int(w[1] * k) - org[1] - 1; wx1 = int(_m266.ceil(w[2] * k)) - org[0] + 1; wy1 = int(_m266.ceil(w[3] * k)) - org[1] + 1
                Z[max(0, wy0):max(0, wy1), max(0, wx0):max(0, wx1)] = False
            M = Z & (A.min(axis=2) < 255); A[M] = _r7_fond_zone(A, Z); im[cle] = _v260_png(A); v.pop("cache", None)  # R7 (S11) : le fond, non le blanc


# ═══ FT — APRÈS L'EFFACEMENT D'UN MOT CURSIF, LES TRAITS D'UN PIXEL PARTENT (A70_z23 ; DEC-2900, DEC-3311 ; ligne 3.7 du plan) ═══
# Portage de agents274/A70_z23/a70_formes3.py (masque, forme « FT » ; applique) avec les contextes d'a70_lib2.py (contextes : RC, Zo, sub) et la dilatation sans boucle d'a70_lib3.py.
# Le défaut : la zone effacée (le cache, en pixels, plancher à gauche et en haut) n'a aucune marge ; la colonne de pixels du bord gauche d'une lettre, qui tombe de l'autre côté
# de la ligne de la grille de l'image, reste : 213 des 218 restes mesurés sont à 1 pixel de la zone et larges de 1 pixel.
# La forme : une fois tous les mots d'une image effacés, un morceau d'encre (composante 8-connexe des pixels non blancs) est mis à blanc s'il (i) ne touche aucun rectangle de cache,
# (ii) touche un pixel que l'effacement vient d'ôter (dilaté de 1 pixel, disque de rayon 1), (iii) compte au plus 150 pixels, (iv) ne contient aucun pixel des tracés de la page
# dilatés de 2 pixels (avant Q1) ni des boîtes (± 1 pixel) des mots voisins (ceux de G2, sous FB8). C'est une règle de forme (géométrique) : « les éléments qui se touchent sont
# indissociables » (DEC-1659, DEC-1864), étendue à D6 par DEC-2900. Elle ne repeint rien : un pixel ne passe qu'à blanc.
# Seul écart au texte du banc : les composantes et leurs comptes se calculent avec l'étiquetage du serveur et numpy, non avec scipy.ndimage (comme G2).
# Mesurée au banc sur les 129 (une seule famille de PDF porte des mots cursifs) : agents274/A70_z23/RAPPORT_A70.md et FORMES_129.txt ; D6 complet : agents283/A144_Ce2/RAPPORT_A144.md.
# Les épreuves de ce portage : agentsCODE/PS1/RAPPORT_PS1.md.
_A70_FT_TAILLE = 150


def _a70_dilate(M, r):
    """Dilatation par un disque de rayon r, SANS boucle aux bords (tableau rempli de zéros autour) : a70_lib3.dilate. G2 garde la sienne (_v269_g2_dilate), qui boucle."""
    H, W = M.shape
    P = _np247.zeros((H + 2 * r, W + 2 * r), bool)
    P[r:r + H, r:r + W] = M
    out = P.copy()
    for dy in range(-r, r + 1):
        for dx in range(-r, r + 1):
            if dx * dx + dy * dy > r * r:
                continue
            out |= _np247.roll(_np247.roll(P, dy, axis=0), dx, axis=1)
    return out[r:r + H, r:r + W]


def _a70_pieces(images, f, q):
    """Les images que D6-0i désigne pour un mot de la figure f dont le cache est q : l'entière de feuille si la figure en a une, sinon chaque morceau qui touche le cache, sur sa part (a70_lib.pieces)."""
    ent = [i for i, im in enumerate(images) if im.get("figure") == f and im.get("entier_feuille")]
    if ent:
        return [(ent[0], "entier_feuille", list(q))]
    out = []
    for i, im in enumerate(images):
        c = im["cadre"]
        if im.get("figure") == f and im.get("data_feuille") and c[0] < q[2] and q[0] < c[2] and c[1] < q[3] and q[1] < c[3]:
            out.append((i, "data_feuille", [max(q[0], c[0]), max(q[1], c[1]), min(q[2], c[2]), min(q[3], c[3])]))
    return out


def _a70_ft_masque(A, B, RC, prot):
    """Le masque de la forme FT (a70_formes3.masque) : A l'image après l'effacement, B l'image avant, RC les rectangles de cache, prot = les tracés (sub) et les mots voisins (Zo) ;
    rend les pixels à mettre à blanc."""
    H, Wd = A.shape[:2]
    N = A.min(axis=2) < 255
    E = (B.min(axis=2) < 255) & ~N & RC
    adj = _a70_dilate(E, 1)
    lab = _v247_etiqueter(N, huit=True)
    if not lab.any():
        return _np247.zeros((H, Wd), bool)
    ids, inv = _np247.unique(lab, return_inverse=True)
    inv = inv.reshape(lab.shape).ravel()
    taille = _np247.bincount(inv, minlength=len(ids))

    def dedans(M):
        return _np247.bincount(inv, weights=M.ravel().astype(float), minlength=len(ids)) > 0

    ok = (taille <= _A70_FT_TAILLE) & ~dedans(RC) & ~dedans(prot) & dedans(adj) & (ids != 0)
    return ok[inv].reshape(H, Wd)


def _a70_ft(images, vb, faits, doc, ctx):
    """FT, appelée par G2 après la boucle d'effacement : pour chaque image qu'un cache désigne, les contextes d'a70_lib2.contextes (RC, Zo sous FB8, sub avant Q1),
    puis le masque de FT et sa mise à blanc (a70_formes3.applique). `faits` = [figure, texte, boîte, cache] de chaque mot cursif à cache ; ctx['ORIG'] = les images d'avant l'effacement."""
    C = {}
    for f, _t, _bb, cache in faits:
        curs = [w for w in (vb or {}).get(f, []) if w.get("sorte") == "cursif"]
        for i, cle, qc in _a70_pieces(images, f, cache):
            im = images[i]
            A0 = ctx["ORIG"][(id(im), cle)]
            H, W = A0.shape[:2]
            c = im["entier_cadre"] if cle == "entier_feuille" else im["cadre"]
            k = A0.shape[1] / (c[2] - c[0])
            if cle == "data_feuille":
                org = (int(round(c[0] * k)), int(round(c[1] * k)))
            else:
                r_ = fitz.Rect(c) * fitz.Matrix(k, k)
                org = (r_.irect.x0, r_.irect.y0)
            x0 = max(0, int(_m266.floor(qc[0] * k)) - org[0])
            y0 = max(0, int(_m266.floor(qc[1] * k)) - org[1])
            x1 = min(W, int(_m266.ceil(qc[2] * k)) - org[0])
            y1 = min(H, int(_m266.ceil(qc[3] * k)) - org[1])
            if x1 <= x0 or y1 <= y0:
                continue
            e = C.get((i, cle))
            if e is None:
                pg = doc[im["page"] - 1]
                Zo = _np247.zeros((H, W), bool)
                for w in pg.get_text("words"):
                    if any(abs(w[0] - cc["bbox"][0]) < 1 and abs(w[1] - cc["bbox"][1]) < 1 for cc in curs) or _a47_reconnu_fb8(w, curs):
                        continue
                    wx0 = int(w[0] * k) - org[0] - 1
                    wy0 = int(w[1] * k) - org[1] - 1
                    wx1 = int(_m266.ceil(w[2] * k)) - org[0] + 1
                    wy1 = int(_m266.ceil(w[3] * k)) - org[1] + 1
                    Zo[max(0, wy0):max(0, wy1), max(0, wx0):max(0, wx1)] = True
                Mp = _v269_g2_dilate(_v269_g2_rendu_seul(pg, k, ctx["RAW"]), 2)
                sub = _np247.zeros((H, W), bool)
                ys0, xs0 = max(0, org[1]), max(0, org[0])
                ys1, xs1 = min(Mp.shape[0], org[1] + H), min(Mp.shape[1], org[0] + W)
                sub[ys0 - org[1]:ys1 - org[1], xs0 - org[0]:xs1 - org[0]] = Mp[ys0:ys1, xs0:xs1]
                e = C[(i, cle)] = dict(Zo=Zo, sub=sub, RC=_np247.zeros((H, W), bool))
            e["RC"][y0:y1, x0:x1] = True
    for (i, cle), e in sorted(C.items()):
        im = images[i]
        A = _v260_rgb(im[cle])
        Wm = _a70_ft_masque(A, ctx["ORIG"][(id(im), cle)], e["RC"], e["sub"] | e["Zo"])
        if not Wm.any():
            continue
        A2 = A.copy()
        A2[Wm] = 255
        im[cle] = _v260_png(A2)


def _v269_g2(images, vb, doc):
    """G2 (D6-0i + S + Q1) : sans image entière, chaque morceau qui touche le cache est effacé sur sa part ; puis la boucle d'effacement ; les cache
    des mots cursifs sont retirés de voisins_b. Rend la liste des effacements faits [figure, texte, bbox, cache] (pour les rejouer après une
    recomposition)."""
    faits = [[f, v["texte"], v["bbox"], v["cache"]] for f, LL in (vb or {}).items() for v in LL if v.get("sorte") == "cursif" and v.get("cache")]
    if not faits:
        return []
    ctx = dict(ORIG={}, LET={}, RAW={})
    for im in images:
        for cle in ("entier_feuille", "data_feuille"):
            if im.get(cle):
                ctx["ORIG"][(id(im), cle)] = _v260_rgb(im[cle]).copy()
    split = {}
    for f_, LL in vb.items():
        ent = any(x.get("figure") == f_ and x.get("entier_feuille") for x in images); L2 = []
        for v in LL:
            if v.get("sorte") == "cursif" and v.get("cache") and not ent:
                q = v["cache"]
                for x in images:
                    c = x["cadre"]
                    if x.get("figure") == f_ and x.get("data_feuille") and c[0] < q[2] and q[0] < c[2] and c[1] < q[3] and q[1] < c[3]:
                        L2.append(dict(v, cache=[max(q[0], c[0]), max(q[1], c[1]), min(q[2], c[2]), min(q[3], c[3])]))
            else:
                L2.append(v)
        split[f_] = L2
    _v269_g2_efface(images, split, doc, ctx)
    _a70_ft(images, vb, faits, doc, ctx)
    for f_, LL in vb.items():
        for v in LL:
            if v.get("sorte") == "cursif":
                v.pop("cache", None)
    return faits


def _v269_g2_rejoue(images, faits, doc):
    """Après une recomposition (la gomme, avec la clé), rejoue G2 sur la liste des effacements faits, pour les seules figures qui ont une image
    entière (recomposée à partir des morceaux) ; les morceaux effacés seuls ne sont pas réécrits par la gomme et gardent leur effacement."""
    avec = {x.get("figure") for x in images if x.get("entier_feuille")}
    vb = {}
    for f, t, bb, cache in (faits or []):
        if f in avec:
            vb.setdefault(f, []).append(dict(texte=t, bbox=bb, sorte="cursif", cache=list(cache)))
    return _v269_g2(images, vb, doc)


# ═══ v2.69 — LA MARQUE DU MORCEAU : LA MOITIÉ SERVEUR DE (c) CORRIGÉE (F094 ; DEC-3209, DEC-3213 point 2, DEC-3214 ; codée avant S125 : DEC-3273 R1) ═══
# Portage, par programme, de la marque d'A230 après DEC-3213 (agents288/A230_RD, serveur c2) : sur chaque image, `marque` ("cadre" | "fond" | "bord" | "rien")
# et `marque_couleur`, lus dans les TRACÉS du PDF source, jamais dans les pixels. « cadre » : un tracé fermé de couleur posé sur une partie d'une ligne ;
# « fond » : un aplat de couleur sous la plupart des mots ; « bord » : la couleur ENTOURE le contenu (le bord d'une carte, une étiquette blanche sur un fond) ;
# « rien » : le reste (le bord noir ou gris, un aplat blanc ou gris, un texte seul en couleur, une photo). Priorité : cadre, fond, bord. La couleur d'un TRACÉ
# est celle que le serveur lit déjà pour le texte (_V266_CHROMA_LIGNE) ; un APLAT dessiné compte dès qu'il porte une couleur (_V269_CHROMA_FOND_DESSINE :
# DEC-3271 E16, « pas de seuil, la couleur exacte de la source » ; un gris ou un blanc n'en est pas une), sa couleur vue sur du blanc, l'opacité prise en
# compte (le fond (a), DEC-3216 : la formule de _a231_hex d'A233) ; pour « fond », le plus grand aplat. `marque_couleurs` garde TOUTES les couleurs lues
# (cadre, fond, bord : DEC-3271 E17, « Le fond et le cadre gardent tous deux leur couleur »). Le blanc d'une étiquette : tous les canaux >= 235. La page applique la marque (étape D) ;
# rien ne la lit encore. Ajout pur ; une panne est dite dans `a230_erreur`, jamais levée. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_A230_MARGE = 2.0
_A230_MIN_TRACE = 5.0
_A230_PART_FOND = 0.5
_A230_BLANC = 235
_A230_PART_ETIQUETTE = 0.5
_V269_CHROMA_FOND_DESSINE = 1    # E16 (DEC-3271) : un aplat dessiné compte dès qu'il porte une couleur (écart des canaux >= 1) ; écart 0 = gris ou blanc


def _a230_chroma(c):
    """écart max-min des canaux (0..255) d'une couleur PyMuPDF (None, gris 1 canal, RVB 3, CMJN 4) ; 0 pour 'pas de couleur' ou un gris."""
    if not c:
        return 0
    v = list(c)
    if len(v) == 4:
        k = v[3]
        v = [(1 - v[0]) * (1 - k), (1 - v[1]) * (1 - k), (1 - v[2]) * (1 - k)]
    if len(v) < 3:
        return 0
    w = [int(round(x * 255)) for x in v[:3]]
    return max(w) - min(w)


def _a230_hex(c):
    """« #rrggbb » d'une couleur PyMuPDF (RVB 3 ou CMJN 4), ou None."""
    if not c:
        return None
    v = list(c)
    if len(v) == 4:
        k = v[3]
        v = [(1 - v[0]) * (1 - k), (1 - v[1]) * (1 - k), (1 - v[2]) * (1 - k)]
    if len(v) < 3:
        return None
    return "#%02x%02x%02x" % tuple(int(round(x * 255)) for x in v[:3])


def _v269_sur_blanc(c, op):
    """« #rrggbb » d'une couleur d'aplat PyMuPDF (RVB 3 ou CMJN 4) vue sur du blanc, l'opacité `op` prise en compte (None = opaque) ; None si illisible."""
    if not c:
        return None
    v = list(c)
    if len(v) == 4:
        k = v[3]
        v = [(1 - v[0]) * (1 - k), (1 - v[1]) * (1 - k), (1 - v[2]) * (1 - k)]
    if len(v) < 3:
        return None
    op = 1.0 if op is None else float(op)
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(round((x * op + (1 - op)) * 255)))) for x in v[:3])


def _a230_blanc(c):
    """vrai si la couleur de remplissage est un blanc (tous les canaux >= 235)."""
    if not c:
        return False
    v = list(c)
    if len(v) == 4:
        k = v[3]
        v = [(1 - v[0]) * (1 - k), (1 - v[1]) * (1 - k), (1 - v[2]) * (1 - k)]
    if len(v) < 3:
        return False
    return min(int(round(x * 255)) for x in v[:3]) >= _A230_BLANC


def _a230_ferme(d):
    """un tracé fermé : closePath, un rectangle (re) ou un quadrilatère (qu), ou une boucle dont le dernier point retombe sur le premier."""
    its = d.get("items") or []
    if d.get("closePath"):
        return True
    if any(i[0] in ("re", "qu") for i in its):
        return True
    if len(its) >= 3:
        a, b = its[0], its[-1]
        p0 = a[1] if a[0] in ("l", "c") else None
        p1 = b[-1] if b[0] in ("l", "c") else None
        if p0 is not None and p1 is not None:
            return abs(p0.x - p1.x) < 0.6 and abs(p0.y - p1.y) < 0.6
    return False


def _a230_mots_dans(mots, r):
    """le nombre de mots de la page dont le centre tombe dans le rectangle r."""
    n = 0
    for w in mots:
        cx, cy = (w[0] + w[2]) / 2.0, (w[1] + w[3]) / 2.0
        if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1:
            n += 1
    return n


def _a230_mot_hors_meme_ligne(mots, R, r):
    """vrai s'il existe un mot du morceau R (centre dans R), hors du tracé r (centre hors de r), sur la même ligne que lui (recouvrement vertical >= la moitié de la hauteur du mot)."""
    for w in mots:
        cx, cy = (w[0] + w[2]) / 2.0, (w[1] + w[3]) / 2.0
        if not (R.x0 <= cx <= R.x1 and R.y0 <= cy <= R.y1):
            continue
        if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1:
            continue
        h = max(w[3] - w[1], 1e-6)
        if min(w[3], r.y1) - max(w[1], r.y0) >= 0.5 * h:
            return True
    return False


def _s125_voisin_de_la_page(mots, R, r):
    """S125, R1 (DEC-3230 : « S125 : R8 retenue sous réserve », R8 porte les marques de R1 ; panel `agents288/A240_S125/PANEL_S125_A240.md`) : vrai s'il existe un mot DE LA PAGE,
    hors du tracé r, sur la même ligne que lui (recouvrement vertical >= la moitié de la hauteur du mot) et à 6 mm au plus de lui (la distance de la découpe).
    Remplace, pour le test « cadre ou bord », `_a230_mot_hors_meme_ligne`, qui ne regardait que les mots du morceau : un cadre coupé seul par la découpe
    (son voisin « + » dans un autre morceau) n'était plus vu comme posé sur une ligne et devenait un « bord » (la 052, la 063, la 094)."""
    g6 = 6.0 * 72.0 / 25.4
    for w in mots:
        cx, cy = (w[0] + w[2]) / 2.0, (w[1] + w[3]) / 2.0
        if r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1:
            continue
        h = max(w[3] - w[1], 1e-6)
        if min(w[3], r.y1) - max(w[1], r.y0) >= 0.5 * h and max(0.0, w[0] - r.x1, r.x0 - w[2]) <= g6:
            return True
    return False


def _a230_marque_morceau(dessins, mots, cadre):
    """(marque, détail) d'un morceau : `dessins` = page.get_drawings(), `mots` = page.get_text("words"), `cadre` = [x0, y0, x1, y1] en points de la page.
    détail = {cadre: [..], fond: [..], bord: [..], couleur: '#rrggbb' | None} (les tracés qui ont fait la marque, pour le banc ; la route n'en garde que `marque` et `marque_couleur`)."""
    R = fitz.Rect(cadre)
    aire = max(R.width * R.height, 1e-6)
    M = _A230_MARGE
    cand = []
    for d in dessins:
        r = d.get("rect")
        if r is None:
            continue
        if r.x0 >= R.x0 - M and r.y0 >= R.y0 - M and r.x1 <= R.x1 + M and r.y1 <= R.y1 + M:
            cand.append(d)
    det = {"cadre": [], "fond": [], "bord": [], "couleur": None}
    blancs = [d["rect"] for d in cand if d.get("type") in ("f", "fs") and _a230_blanc(d.get("fill"))]
    for d in cand:
        r = d["rect"]
        if d.get("type") in ("s", "fs") and _a230_chroma(d.get("color")) >= _V266_CHROMA_LIGNE and _a230_ferme(d) and r.width >= _A230_MIN_TRACE and r.height >= _A230_MIN_TRACE and _a230_mots_dans(mots, r) >= 1:
            e = {"rect": [round(v, 1) for v in r], "part_du_morceau": round((r.width * r.height) / aire, 3), "couleur": _a230_hex(d.get("color"))}
            (det["cadre"] if _s125_voisin_de_la_page(mots, R, r) else det["bord"]).append(e)
        if d.get("type") in ("f", "fs") and _a230_chroma(d.get("fill")) >= _V269_CHROMA_FOND_DESSINE:
            inter = fitz.Rect(r)
            inter.intersect(R)
            if inter.is_empty:
                continue
            part = (inter.width * inter.height) / aire
            n_mots = _a230_mots_dans(mots, inter)
            if part >= _A230_PART_FOND and n_mots >= 1:
                n_blanc = 0
                for w in mots:
                    cx, cy = (w[0] + w[2]) / 2.0, (w[1] + w[3]) / 2.0
                    if inter.x0 <= cx <= inter.x1 and inter.y0 <= cy <= inter.y1 and any(b.x0 <= cx <= b.x1 and b.y0 <= cy <= b.y1 and (b.width * b.height) < (r.width * r.height) for b in blancs):
                        n_blanc += 1
                e = {"rect": [round(v, 1) for v in r], "part_du_morceau": round(part, 3), "couleur": _v269_sur_blanc(d.get("fill"), d.get("fill_opacity")), "mots": n_mots, "mots_sur_etiquette_blanche": n_blanc}
                (det["bord"] if n_blanc >= _A230_PART_ETIQUETTE * n_mots else det["fond"]).append(e)
    if det["cadre"]:
        marque, c = "cadre", det["cadre"][0]["couleur"]
    elif det["fond"]:
        marque, c = "fond", max(det["fond"], key=lambda e: e["part_du_morceau"])["couleur"]
    elif det["bord"]:
        marque, c = "bord", det["bord"][0]["couleur"]
    else:
        marque, c = "rien", None
    det["couleur"] = c
    det["couleurs"] = {k: (max(det[k], key=lambda e: e["part_du_morceau"]) if k == "fond" else det[k][0])["couleur"] for k in ("cadre", "fond", "bord") if det[k]}
    return marque, det


# ═══ v2.69 — LE FOND LU DANS UNE IMAGE (E16 second cas, DEC-3271 ; R8, DEC-3276 : « au moins un mot ; seuil 16 ») ═══
# Pour une image que les tracés laissent « rien » et que les images du PDF couvrent au moins à moitié : sa couleur dominante (chaque canal ramené
# à 16 niveaux, la case la plus peuplée, la couleur moyenne de ses pixels) est un fond si elle couvre au moins la moitié de l'image, n'est pas du
# blanc pur, a un écart de canaux d'au moins 16 (le seuil le plus bas qui ne prend aucune tache, cherché à sec sur les 129 : les taches vont
# jusqu'à 15, les couleurs voulues commencent à 20) et si au moins un mot de la page a son centre dans l'image. Le corpus n'a presque pas de
# scans : ce seuil vaut pour ce qui a été vu. Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_V269_FOND_IMAGE_SEUIL = 16
_V269_FOND_IMAGE_PART = 0.5


def _v269_fond_image(im, infos, mots):
    """« #rrggbb » du fond lu dans l'image `im` (cadre, data), ou None ; `infos` = les boîtes des images de la page, `mots` = ses mots."""
    R = fitz.Rect(im["cadre"]); aire = R.width * R.height
    if aire <= 0 or not infos:
        return None
    cov = 0.0
    for q in infos:
        i = fitz.Rect(q); i.intersect(R)
        cov = max(cov, (i.width * i.height) / aire if not i.is_empty else 0)
    if cov < 0.5 or _a230_mots_dans(mots, R) < 1:
        return None
    A = _v260_rgb(im["data"]).astype(int)
    q = A >> 4; key = q[..., 0] * 256 + q[..., 1] * 16 + q[..., 2]
    vals, cnt = _np247.unique(key, return_counts=True); k = vals[cnt.argmax()]; part = cnt.max() / key.size
    moy = A[key == k].mean(axis=0); ecart = int(round(moy.max() - moy.min()))
    if part < _V269_FOND_IMAGE_PART or moy.min() >= 250 or ecart < _V269_FOND_IMAGE_SEUIL:
        return None
    return "#%02x%02x%02x" % tuple(int(round(x)) for x in moy)


def _a230_marques(content, result):
    """ajoute `marque` ("cadre" | "fond" | "bord" | "rien") et `marque_couleur` à chaque image du résultat ; une panne est dite dans `a230_erreur`, jamais levée vers /parse."""
    try:
        doc = fitz.open(stream=content, filetype="pdf")
    except Exception as e:
        result["a230_erreur"] = "ouverture : " + (str(e) or type(e).__name__)[:160]
        return result
    try:
        cache = {}
        infos = {}
        for im in result.get("images") or []:
            try:
                pno = im["page"] - 1
                if pno not in cache:
                    pg = doc[pno]
                    cache[pno] = (pg.get_drawings(), pg.get_text("words"))
                marque, det = _a230_marque_morceau(cache[pno][0], cache[pno][1], im["cadre"])
                im["marque"] = marque
                im["marque_couleur"] = det["couleur"]
                im["marque_couleurs"] = det["couleurs"]
                # S125, R8 (DEC-3230) : les jetons des mots que le tracé d'un morceau « cadre » porte (centre dans le tracé), ligne par ligne ; la page ne remplace qu'eux par l'image
                if marque == "cadre" and det["cadre"]:
                    _r8 = fitz.Rect(det["cadre"][0]["rect"])
                    im["marque_mots"] = [w[4] for w in sorted(cache[pno][1], key=lambda w: (round((w[1] + w[3]) / 2.0 / 4.0), w[0])) if _r8.x0 <= (w[0] + w[2]) / 2.0 <= _r8.x1 and _r8.y0 <= (w[1] + w[3]) / 2.0 <= _r8.y1]
                # v2.69 — E16, R8 : une image que les tracés laissent « rien » peut porter un fond lu dans l'image (voir _v269_fond_image)
                if marque == "rien":
                    if pno not in infos:
                        infos[pno] = [fitz.Rect(x["bbox"]) for x in doc[pno].get_image_info()]
                    _v269_c = _v269_fond_image(im, infos[pno], cache[pno][1])
                    if _v269_c:
                        im["marque"] = "fond"; im["marque_couleur"] = _v269_c; im["marque_couleurs"] = {"fond": _v269_c}; im["marque_lue"] = "image"
            except Exception as e:
                result.setdefault("a230_erreur", "image %s : %s" % (im.get("index"), (str(e) or type(e).__name__)[:120]))
    except Exception as e:
        result["a230_erreur"] = (str(e) or type(e).__name__)[:200]
    finally:
        doc.close()
    return result


# ═══ R9 — LE BARÈME QUE L'IMAGE MONTRE : UN CHAMP `bareme` SUR CHAQUE IMAGE DE /parse (ligne 5.10 du plan ; DEC-3157, DEC-3096) ═══
# `bareme` : les « /N » que le texte du PDF porte sous le cadre de l'image, ou sous celui de sa figure entière, ET dont l'encre est dans l'image envoyée
# (au moins 5 % de pixels sombres dans la boîte du mot « /N »). La page (R9, à la place du bloc v10.110 (2.3)) en retire d'un titre d'exercice le barème
# « ( …… /N ) » que l'image posée dans le même exercice montre déjà. Aucun autre champ n'est changé ; avec ou sans clé ; une panne est dite dans `a225_erreur`.
# Portage de A225_BASE2 (app_v2_62.py du banc, l. 7146-7241), à l'octet pour les sept objets ci-dessous ; la fonction-champ n'écrit que `bareme`
# (`lignes_couleur`, que le banc écrivait aussi, est dans la 2.72 depuis la v2.66).
_A225_SEUIL_ENCRE = 0.05
_A225_RX_V2 = re.compile(r"(?<![0-9A-Za-zÀ-ÿ])/\s*(\d+(?:[.,]\d+)?)")


def _a225_etiquettes(page, cadres):
    """les nombres N des motifs « /N » trouvés sous le cadre de l'image puis sous celui de sa figure entière (règle v2 d'A207 : le caractère de gauche n'est ni lettre ni chiffre)."""
    res = []
    for c in cadres:
        if not c:
            continue
        t = page.get_text("text", clip=fitz.Rect(c)).replace("\xa0", " ")
        for m in _A225_RX_V2.finditer(t):
            if m.group(1) not in res:
                res.append(m.group(1))
    return res


def _a225_pixels(url):
    if not url or not isinstance(url, str):
        return None
    b = url.split(",", 1)[1] if url.startswith("data:") else url
    try:
        return _np247.array(_Im247.open(io.BytesIO(base64.b64decode(b))).convert("L"))
    except Exception:
        return None


def _a225_boites_mot(page, cadre, n):
    mots = page.get_text("words", clip=fitz.Rect(cadre))
    out = []
    re_n = re.compile(r"/\s*" + re.escape(n) + r"(?![\d])")
    for k, w in enumerate(mots):
        t = w[4]
        if re_n.search(t):
            out.append(tuple(w[:4]))
        elif t.endswith("/") and k + 1 < len(mots) and re.match(re.escape(n) + r"(?![\d])", mots[k + 1][4]):
            a, b = w, mots[k + 1]
            out.append((min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])))
    return out


def _a225_part_sombre(pixels, cadre, boite):
    if pixels is None:
        return None
    h_, w_ = pixels.shape
    cw, ch = cadre[2] - cadre[0], cadre[3] - cadre[1]
    if cw <= 0 or ch <= 0:
        return None
    sx, sy = w_ / cw, h_ / ch
    x0 = max(0, int(round((boite[0] - cadre[0]) * sx)))
    x1 = min(w_, int(round((boite[2] - cadre[0]) * sx)))
    y0 = max(0, int(round((boite[1] - cadre[1]) * sy)))
    y1 = min(h_, int(round((boite[3] - cadre[1]) * sy)))
    if x1 <= x0 or y1 <= y0:
        return None
    return float((pixels[y0:y1, x0:x1] < 128).mean())


def _a225_bareme_image(page, im):
    """les « /N » du PDF que l'image envoyée montre : une étiquette de la règle v2 est gardée si, dans `data` (sous le cadre) ou `entier` (sous le cadre de la figure entière), la boîte du mot « /N » porte au moins 5 % de pixels sombres."""
    garde = []
    for n in _a225_etiquettes(page, [im.get("cadre"), im.get("entier_cadre")]):
        tenu = False
        for cadre, champ in ((im.get("cadre"), "data"), (im.get("entier_cadre"), "entier")):
            if not cadre:
                continue
            px = _a225_pixels(im.get(champ))
            for b in _a225_boites_mot(page, cadre, n):
                p = _a225_part_sombre(px, cadre, b)
                if p is not None and p >= _A225_SEUIL_ENCRE:
                    tenu = True
        if tenu:
            garde.append(n)
    return garde


def _a225_bareme_champ(content, result):
    """ajoute `bareme` à chaque image du résultat ; une panne est dite dans `a225_erreur`, jamais levée vers /parse."""
    try:
        doc = fitz.open(stream=content, filetype="pdf")
    except Exception as e:
        result["a225_erreur"] = "ouverture : " + (str(e) or type(e).__name__)[:160]
        return result
    try:
        for im in result.get("images") or []:
            try:
                im["bareme"] = _a225_bareme_image(doc[im["page"] - 1], im)
            except Exception as e:
                im["bareme"] = []
                result.setdefault("a225_erreur", "image %s : %s" % (im.get("index"), (str(e) or type(e).__name__)[:120]))
    except Exception as e:
        result["a225_erreur"] = (str(e) or type(e).__name__)[:200]
    finally:
        doc.close()
    return result


# ═══ PP7 — LA CHAÎNE F-b + l8 + n2 : LES TRAITS DE SOULIGNEMENT QUE LA DÉCOUPE A LAISSÉS SEULS (ligne 4.11 du plan ; DEC-2817, DEC-2827, DEC-2948) ═══
# Quand la découpe sépare une consigne de son soulignement, il reste dans une image un trait seul (le soulignement de « Les saisons » de la 006) ou un trait
# collé à un cadre. F-b repeint la boîte de ce trait à la couleur du fond quand la ligne de texte qu'il souligne n'est PAS dans l'image (C1) et que le trait n'est
# pas le bord d'une case (C2) ; l8 efface en plus le bout de lettre qui croise le trait et le bord du trait que la boîte de F-b laissait ; n2 traite un soulignement
# fait de plusieurs tronçons abouts (014 p. 1) : tous effacés, ou rien. « F-b + l8 + n2 : toute la chaîne, ou rien » (DEC-2827). Une image qui n'a plus aucune
# encre dans `data` reste dans la liste, devient entièrement blanche et est marquée `vide` (DEC-2902 : la page ne la pose jamais). Le serveur ne lit PAS ce que le
# modèle écrira : F-b, l8 et n2 ne lisent que le PDF (lignes de mots, tracés) et les pixels de l'image ; ils agissent donc ici, sur la réponse de /parse.
# Portage, corps recopiés, de ce que le banc a mesuré : agents270/A14_006b/formes.py (A14 : rect_in, lignes_mots, info_page, c2_case, frac_dans, box_px, fond,
# repeint, la règle b de joue), agents270/A5_006/defs.py (est_trait), agents273/A26_R3/l/al_lib.py, al_lib2.py, al_lib3.py (px_pt, band, fond_fb, trait_width, foot,
# band2, box5, l1b, l8), agents273/A26_R3/n/an_lib.py (segments, lignes, meme, chaine, qualifie, tronçon, le pipeline P2) et agents272/N5_feuilles/scripts/n5_fb.py
# pour ce que devient une image retirée (restée dans la liste, blanche ; l'entière reste à sa tête). Seuls changent : les noms (préfixe _pp7_) ; scipy.ndimage.label
# est l'étiquetage 8-connexe du serveur (_v247_etiqueter) et binary_dilation une dilatation 3×3 en numpy (l'égalité pixel à pixel est éprouvée) ; les règles a et d
# de formes.py, les pipelines P0/P1/P3 et les mesures du banc ne sont pas portés ; les dessins d'une page sont lus une fois par page ; un champ n'est décodé que si la
# géométrie retient un trait ; `sceau` est refait quand `data` change (comme _v270_photographie_apres) ; tout ou rien par document : une panne laisse le résultat
# d'aujourd'hui et se dit dans `chaine_fb_erreur`. Le pourquoi et les preuves : JOURNAL BACKEND, entrée de cette version.
import math as _mp7

_PP7_EXO = re.compile(r"^p\d+-e")
_PP7_DELTA = 1.0
_PP7_GAP = 1.0
_PP7_DY = 0.5
_PP7_HMAX = 3.0
_PP7_WMIN = 1.0
_PP7_COUV = 0.90
_PP7_VERT = 3.0


def _pp7_rect_in(T, c, m=1.0):
    return T.x0 >= c[0] - m and T.y0 >= c[1] - m and T.x1 <= c[2] + m and T.y1 <= c[3] + m


def _pp7_est_trait(x0, y0, x1, y1, boites):
    """defs.est_trait (A5) : g1 (h ≤ 3 pt, l ≥ 12 pt, l/h ≥ 10) puis g2 (a, b, c) pour les boîtes de lignes données."""
    w, h = x1 - x0, y1 - y0
    h = max(h, 1e-6)
    if not (h <= 3.0 and w >= 12.0 and w / h >= 10):
        return False
    for l in boites:
        if abs(y0 - l[3]) > 3.0:
            continue
        rec = max(0.0, min(x1, l[2]) - max(x0, l[0]))
        if rec < 0.9 * w:
            continue
        lw = l[2] - l[0]
        if lw > 0 and 0.8 <= w / lw <= 1.25:
            return True
    return False


def _pp7_lignes_mots(pg):
    """formes.lignes_mots : une boîte par ligne de mots (bloc, ligne) de la page."""
    L = {}
    for w in pg.get_text("words"):
        if not w[4].strip():
            continue
        k = (w[5], w[6])
        o = L.setdefault(k, dict(box=[w[0], w[1], w[2], w[3]]))
        b = o["box"]
        b[0] = min(b[0], w[0]); b[1] = min(b[1], w[1]); b[2] = max(b[2], w[2]); b[3] = max(b[3], w[3])
    return list(L.values())


def _pp7_page(ctx, pg):
    """formes.info_page (règle b seulement) + les dessins de la page, lus une fois : {lig, traits, obst, dessins}."""
    k = pg.number
    if k in ctx:
        return ctx[k]
    lig = _pp7_lignes_mots(pg)
    boxes = [l["box"] for l in lig]
    traits = []
    obst = []
    dessins = pg.get_drawings()
    for d in dessins:
        r = d.get("rect")
        if r is None:
            continue
        dash = d.get("dashes") and str(d.get("dashes")).strip() not in ("", "[] 0")
        if dash:
            continue
        blanc = d.get("fill") is not None and tuple(round(v, 2) for v in d["fill"]) == (1.0, 1.0, 1.0) and d.get("color") is None
        if not blanc:
            for it in d.get("items", []):
                pts = []
                if it[0] == "l":
                    pts = [it[1], it[2]]
                elif it[0] == "re":
                    pts = [it[1].tl, it[1].br]
                elif it[0] == "qu":
                    pts = [it[1].ul, it[1].lr, it[1].ur, it[1].ll]
                elif it[0] == "c":
                    pts = list(it[1:5])
                if not pts:
                    continue
                obst.append((min(p.x for p in pts), min(p.y for p in pts), max(p.x for p in pts), max(p.y for p in pts)))
        w, h = r.x1 - r.x0, max(r.y1 - r.y0, 1e-6)
        if not (h <= 3.0 and w >= 12.0 and w / h >= 10):
            continue
        g2 = [i for i, b in enumerate(boxes) if _pp7_est_trait(r.x0, r.y0, r.x1, r.y1, [b])]
        if g2:
            traits.append(dict(r=fitz.Rect(r), L=g2))
    ctx[k] = dict(lig=lig, traits=traits, obst=obst, dessins=dessins, boxes=boxes, seg=None)
    return ctx[k]


def _pp7_c2_case(T, obst):
    x0, y0, x1, y1 = T.x0 - 1, T.y0 - 1, T.x1 + 1, T.y1 + 1
    for (a, b, c, d) in obst:
        if d - b < 4.0:
            continue
        if abs(a - T.x0) < .01 and abs(b - T.y0) < .01 and abs(c - T.x1) < .01 and abs(d - T.y1) < .01:
            continue
        if a <= x1 and c >= x0 and b <= y1 and d >= y0:
            return True
    return False


def _pp7_frac_dans(box, cadre):
    ww = max(0.0, min(box[2], cadre[2]) - max(box[0], cadre[0]))
    hh = max(0.0, min(box[3], cadre[3]) - max(box[1], cadre[1]))
    a = max(1e-6, (box[2] - box[0]) * (box[3] - box[1]))
    return ww * hh / a


def _pp7_box_px(T, cadre, shape, dec=1):
    h, w = shape[:2]
    sx = w / (cadre[2] - cadre[0])
    sy = h / (cadre[3] - cadre[1])
    x0 = int(_mp7.floor((T.x0 - cadre[0]) * sx)) - dec
    x1 = int(_mp7.ceil((T.x1 - cadre[0]) * sx)) + dec
    y0 = int(_mp7.floor((T.y0 - cadre[1]) * sy)) - dec
    y1 = int(_mp7.ceil((T.y1 - cadre[1]) * sy)) + dec
    return max(0, x0), max(0, y0), min(w, x1), min(h, y1)


def _pp7_fond(arr, box, autres):
    h, w = arr.shape[:2]
    x0, y0, x1, y1 = box
    X0, Y0, X1, Y1 = max(0, x0 - 3), max(0, y0 - 3), min(w, x1 + 3), min(h, y1 + 3)
    m = _np247.zeros((h, w), bool)
    m[Y0:Y1, X0:X1] = True
    m[y0:y1, x0:x1] = False
    for (a, b, c, d) in autres:
        m[b:d, a:c] = False
    if not m.any():
        return _np247.array([255, 255, 255], _np247.uint8)
    px = arr[m].astype(_np247.int64)
    k = px[:, 0] * 65536 + px[:, 1] * 256 + px[:, 2]
    u, n = _np247.unique(k, return_counts=True)
    v = int(u[n.argmax()])
    return _np247.array([v >> 16, (v >> 8) & 255, v & 255], _np247.uint8)


def _pp7_repeint(arr, cadre, Ts):
    boxes = [_pp7_box_px(T, cadre, arr.shape, 1) for T in Ts]
    out = arr.copy()
    for i, bx in enumerate(boxes):
        f = _pp7_fond(arr, bx, boxes[:i] + boxes[i + 1:])
        out[bx[1]:bx[3], bx[0]:bx[2]] = f
    return out


def _pp7_ink(a):
    return a.mean(axis=2) < 200


def _pp7_px_pt(shape, cadre):
    h, w = shape[:2]
    return w / (cadre[2] - cadre[0]), h / (cadre[3] - cadre[1])


def _pp7_band(T, cadre, shape, delta):
    h, w = shape[:2]
    sx, sy = _pp7_px_pt(shape, cadre)
    return (max(0, int(_mp7.floor((T.x0 - delta - cadre[0]) * sx))), max(0, int(_mp7.floor((T.y0 - delta - cadre[1]) * sy))),
            min(w, int(_mp7.ceil((T.x1 + delta - cadre[0]) * sx))), min(h, int(_mp7.ceil((T.y1 + delta - cadre[1]) * sy))))


def _pp7_trait_width(P, T):
    """al_lib2.trait_width : la largeur du tracé qui porte T (0 si rempli sans couleur de trait ; 1,0 pt si absente) ; None si aucun dessin ne porte T."""
    for d in P["dessins"]:
        r = d.get("rect")
        if r is None:
            continue
        if abs(r.x0 - T.x0) < .02 and abs(r.y0 - T.y0) < .02 and abs(r.x1 - T.x1) < .02 and abs(r.y1 - T.y1) < .02:
            return 0.0 if d.get("color") is None else float(d.get("width") or 1.0)
    return None


def _pp7_band2(T, cadre, shape, delta, w):
    return _pp7_band(fitz.Rect(T.x0, T.y0 - w / 2.0, T.x1, T.y1 + w / 2.0), cadre, shape, delta)


def _pp7_box5(T, cadre, shape, w):
    h, wd = shape[:2]
    sx, sy = _pp7_px_pt(shape, cadre)
    x0, _, x1, _ = _pp7_box_px(T, cadre, shape, 1)
    return x0, max(0, int(_mp7.floor((T.y0 - w / 2.0 - cadre[1]) * sy)) - 1), x1, min(h, int(_mp7.ceil((T.y1 + w / 2.0 - cadre[1]) * sy)) + 1)


def _pp7_fond_fb(A1, T, cadre):
    bx = _pp7_box_px(T, cadre, A1.shape, 1)
    if bx[2] <= bx[0] or bx[3] <= bx[1]:
        return None
    f = A1[bx[1], bx[0]]
    return None if f.astype(float).mean() < 200 else f


def _pp7_dilate8(m):
    """binary_dilation(m, structure 3×3, une fois, bord à faux)."""
    h, w = m.shape
    p = _np247.pad(m, 1)
    out = _np247.zeros_like(m)
    for dy in (0, 1, 2):
        for dx in (0, 1, 2):
            out |= p[dy:dy + h, dx:dx + w]
    return out


def _pp7_l1b(A0, A1, cadre, Ts, Ws):
    """al_lib2.l1b : les composantes d'encre de l'ORIGINAL qui touchent la boîte de F-b et tiennent dans la bande δ′ sont repeintes à la couleur de F-b."""
    h, w_ = A0.shape[:2]
    lab0 = _v247_etiqueter(_pp7_ink(A0), huit=True)
    A2 = A1.copy()
    light = A1.mean(axis=2) < 250
    for T, w in zip(Ts, Ws):
        f = _pp7_fond_fb(A1, T, cadre)
        if f is None:
            continue
        by = _pp7_band2(T, cadre, A0.shape, _PP7_DELTA, w)
        bxF = _pp7_box_px(T, cadre, A0.shape, 1)
        B_ = _np247.zeros((h, w_), bool)
        B_[by[1]:by[3], by[0]:by[2]] = True
        for lab in _np247.unique(lab0[bxF[1]:bxF[3], bxF[0]:bxF[2]]):
            if lab == 0:
                continue
            m = lab0 == lab
            ys, xs = _np247.nonzero(m)
            y0, y1, x0, x1 = int(ys.min()), int(ys.max()) + 1, int(xs.min()), int(xs.max()) + 1
            if y0 >= by[1] and y1 <= by[3] and x0 >= by[0] and x1 <= by[2]:
                fr = _pp7_dilate8(m) & B_ & light
                A2[fr | (m & B_)] = f
    return A2


def _pp7_l8(A0, A1, cadre, Ts, Ws):
    """al_lib3.l8 : l6 avec le garde-fou « non-encre » (la partie l5 n'est jouée que si le fond de la boîte agrandie n'est pas de l'encre)."""
    boxes = [_pp7_box5(T, cadre, A0.shape, w) for T, w in zip(Ts, Ws)]
    out = A1.copy()
    for i, (T, w) in enumerate(zip(Ts, Ws)):
        f5 = _pp7_fond(A0, boxes[i], boxes[:i] + boxes[i + 1:])
        if _np247.asarray(f5, float).mean() < 200:
            continue
        bx = boxes[i]
        out[bx[1]:bx[3], bx[0]:bx[2]] = f5
    return _pp7_l1b(A0, out, cadre, Ts, Ws)


def _pp7_segments(P):
    """an_lib.segments : les dessins non pointillés de la page, rectangle de hauteur ≤ 3,0 pt et de largeur ≥ 1,0 pt, tracé ou rempli (hors fond blanc)."""
    if P["seg"] is not None:
        return P["seg"]
    out = []
    for d in P["dessins"]:
        r = d.get("rect")
        if r is None:
            continue
        if d.get("dashes") and str(d.get("dashes")).strip() not in ("", "[] 0"):
            continue
        h = r.y1 - r.y0
        w = r.x1 - r.x0
        if h > _PP7_HMAX or w < _PP7_WMIN:
            continue
        col = d.get("color")
        fill = d.get("fill")
        blanc = col is None and fill is not None and tuple(round(v, 2) for v in fill) == (1.0, 1.0, 1.0)
        if blanc or (col is None and fill is None):
            continue
        out.append(dict(r=fitz.Rect(r), w=(0.0 if col is None else float(d.get("width") or 1.0))))
    P["seg"] = out
    return out


def _pp7_meme(a, b):
    return abs(a.x0 - b.x0) < .01 and abs(a.y0 - b.y0) < .01 and abs(a.x1 - b.x1) < .01 and abs(a.y1 - b.y1) < .01


def _pp7_chaine_de(P, T, cadre):
    """an_lib.chaine : les dessins abouts de T (écart ≤ 1,0 pt, même y à 0,5 pt près), transitivement, qui coupent le cadre ; (chaîne, extension gauche, extension droite)."""
    segs = [s for s in _pp7_segments(P) if not _pp7_meme(s["r"], T) and abs(s["r"].y0 - T.y0) <= _PP7_DY and s["r"].x1 > cadre[0] and s["r"].x0 < cadre[2]
            and s["r"].y1 > cadre[1] and s["r"].y0 < cadre[3]]
    lo, hi = T.x0, T.x1
    ch = []
    chg = True
    while chg:
        chg = False
        for s in list(segs):
            if s["r"].x0 <= hi + _PP7_GAP and s["r"].x1 >= lo - _PP7_GAP:
                ch.append(s)
                segs.remove(s)
                lo = min(lo, s["r"].x0)
                hi = max(hi, s["r"].x1)
                chg = True
    gl = max(0.0, T.x0 - max(lo, cadre[0]))
    dr = max(0.0, min(hi, cadre[2]) - T.x1)
    return ch, gl, dr


def _pp7_qualifie(P, s):
    """an_lib.qualifie : un soulignement de texte (≥ 90 % de sa largeur couverte par l'union, en x, des boîtes de lignes dont le bas est à moins de 3,0 pt de son y0)."""
    r = s["r"]
    iv = sorted((max(r.x0, b[0]), min(r.x1, b[2])) for b in P["boxes"] if abs(r.y0 - b[3]) <= _PP7_VERT and min(r.x1, b[2]) > max(r.x0, b[0]))
    tot = 0.0
    cur = None
    for a, b in iv:
        if cur is None or a > cur[1]:
            if cur:
                tot += cur[1] - cur[0]
            cur = [a, b]
        else:
            cur[1] = max(cur[1], b)
    if cur:
        tot += cur[1] - cur[0]
    return tot >= _PP7_COUV * (r.x1 - r.x0)


def _pp7_troncon(P, Ts, cadre):
    """an_lib.tronçon : (à tronçon ?, chaîne agrégée, tous qualifiés ?)."""
    chs = []
    a_tr = False
    for T in Ts:
        ch, gl, dr = _pp7_chaine_de(P, T, cadre)
        if gl > 1.0 or dr > 1.0:
            a_tr = True
        chs += [s for s in ch if not any(_pp7_meme(s["r"], t["r"]) for t in chs)]
    qual = all(_pp7_qualifie(P, s) for s in chs) if chs else True
    return a_tr, chs, qual


def _pp7_p2(P, A0, cadre, Ts, Ws):
    """F-b puis l8, avec n2 : la sortie de la chaîne pour un champ (A0 : l'original ; Ts : les traits que F-b retient)."""
    a_tr, chs, qual = _pp7_troncon(P, Ts, cadre)
    if a_tr and qual:
        Te = list(Ts) + [fitz.Rect(s["r"]) for s in chs]
        We = list(Ws) + [s["w"] for s in chs]
        return _pp7_l8(A0, _pp7_repeint(A0, cadre, Te), cadre, Te, We)
    if a_tr:
        return A0.copy()
    return _pp7_l8(A0, _pp7_repeint(A0, cadre, Ts), cadre, Ts, Ws)


def _pp7_dec(url):
    return _np247.array(_Im247.open(io.BytesIO(base64.b64decode(url.split(",", 1)[1]))).convert("RGB"))


def _pp7_png(arr):
    b = io.BytesIO()
    _Im247.fromarray(arr).save(b, format="PNG")
    return "data:image/png;base64," + base64.b64encode(b.getvalue()).decode(), b.getvalue()


def _pp7_calcule(doc, images):
    """Les unités de la chaîne sur une liste d'images (aucune n'est modifiée) : (les champs à écrire {position: {champ: tableau}}, les positions des images
    devenues sans encre dans `data`, les unités jouées [(position, champ)])."""
    ctx = {}
    ecrire = {}
    vides = []
    unites = []
    for pos, im in enumerate(images):
        if not _PP7_EXO.match(im.get("figure") or ""):
            continue
        sorties = {}
        data_change = False
        for champ, champf, cad in (("entier", "entier_feuille", "entier_cadre"), ("data", "data_feuille", "cadre")):
            if champ == "entier" and not (im.get("entier") and im.get("entier_cadre")):
                continue
            cadre = im.get(cad)
            if not cadre or not im.get(champ):
                continue
            P = _pp7_page(ctx, doc[im["page"] - 1])
            retenus = []
            for t in P["traits"]:
                T = t["r"]
                if not _pp7_rect_in(T, cadre):
                    continue
                c1 = all(_pp7_frac_dans(P["lig"][i]["box"], cadre) < 0.5 for i in t["L"])
                c2 = not _pp7_c2_case(T, P["obst"])
                if c1 and c2:
                    retenus.append(T)
            if not retenus:
                continue
            Ws = [_pp7_trait_width(P, T) or 0.0 for T in retenus]
            for ch in (champ, champf):
                if not im.get(ch):
                    continue
                A0 = _pp7_dec(im[ch])
                sorties[ch] = (A0, _pp7_p2(P, A0, cadre, retenus, Ws))
                unites.append((pos, ch))
            if champ == "data":
                data_change = True
        for ch, (A0, A2) in sorties.items():
            if A2.shape != A0.shape or (A0 != A2).any():
                ecrire.setdefault(pos, {})[ch] = A2
        if data_change and not _pp7_ink(sorties["data"][1]).any():
            vides.append(pos)
            for ch in ("data", "data_feuille"):
                if not im.get(ch):
                    continue
                A0 = sorties[ch][0] if ch in sorties else _pp7_dec(im[ch])
                Z = _np247.full_like(A0, 255)
                if (A0 != Z).any():
                    ecrire.setdefault(pos, {})[ch] = Z
                elif pos in ecrire:
                    ecrire[pos].pop(ch, None)
    return ecrire, vides, unites


def _pp7_applique(doc, images):
    """Calcule d'abord, écrit ensuite (tout ou rien) ; rend {ecrit : [(position, champ)], vides : [positions], unites : [(position, champ)]}."""
    ecrire, vides, unites = _pp7_calcule(doc, images)
    pret = {pos: {ch: _pp7_png(a) for ch, a in champs.items()} for pos, champs in ecrire.items()}
    for pos, champs in pret.items():
        im = images[pos]
        for ch, (url, brut) in champs.items():
            im[ch] = url
            if ch == "data":
                im["sceau"] = _hl.sha256(brut).hexdigest()[:12]
    for pos in vides:
        images[pos]["vide"] = True
    return dict(ecrit=sorted((pos, ch) for pos, champs in pret.items() for ch in champs), vides=list(vides), unites=unites)


def _pp7_chaine(content, result):
    """F-b + l8 + n2 sur les images du résultat ; une panne laisse le résultat tel qu'il était et se dit dans `chaine_fb_erreur`, jamais levée vers /parse."""
    try:
        doc = fitz.open(stream=content, filetype="pdf")
    except Exception as e:
        result["chaine_fb_erreur"] = "ouverture : " + (str(e) or type(e).__name__)[:160]
        return result
    try:
        _pp7_applique(doc, result.get("images") or [])
    except Exception as e:
        result["chaine_fb_erreur"] = (str(e) or type(e).__name__)[:200]
    finally:
        doc.close()
    return result


# ═══ R3b — LA PLACE DES LIGNES DE TEXTE DE CHAQUE IMAGE (DEC-3308 ; DEC-3271 C2 ; DEC-2814) ═══
# R3b ôte d'une image les lignes de texte que le modèle a réécrites. Elle agit dans la page, après la réponse du modèle (que le serveur ne voit
# jamais) ; mais la page ne sait pas où sont les lignes dans l'image : le serveur le lui dit. Pour chaque image, `lignes_image` donne les lignes
# de texte du PDF d'origine dont le centre tombe dans son cadre (« cadre ») et dans son cadre entier (« entier », si elle en a un) : le texte, sa
# boîte et la boîte de chacune de ses lettres, en POINTS du PDF (la page les met à l'échelle de l'image qu'elle pose), et si la ligne est SEULE dans
# un cadre dessiné (`titre_de_cadre` : ce que R3b garde quand le modèle ne l'a pas réécrite en entier). Portage de ce que le banc lit dans le PDF :
# agents273/A26_R3/a26_lib.py (lignes_blocs, aire, dans_poly, frac_boite_dans_poly, cadres_fermes, titre_de_cadre) et agents270/A13_bulles/a13_lib.py
# (pts_path, ferme, traces), corps recopiés. Seuls changent : les noms (préfixe _r3b_) ; les points, rendus tels que le PDF les donne (le banc les arrondit
# en pixels de l'image) ; les tracés de la page, lus une fois par page ; une boîte englobante qui évite de tester un point loin d'un tracé (même résultat) ;
# `frac_boite_dans_poly`, calculée seulement quand aucune autre ligne n'est dans le tracé (même résultat : le banc ne s'en sert qu'avec « aucune autre ») ;
# le diagnostic `info` du banc, qui n'est plus construit. `_a230_ferme` et `_v269_vf_ferme` répondent à d'autres questions (un booléen, une autre
# tolérance) : ils ne sont pas réemployés. Ajout pur : une clé `lignes_image` sur chaque image, rien d'autre ne change dans la réponse ; une panne
# se dit dans `lignes_image_erreur`, jamais levée vers /parse. Le pourquoi et les preuves : JOURNAL BACKEND, entrée de cette version.
_R3B_EPS = 1e-6


def _r3b_pts_path(d):
    P = []
    for it in d["items"]:
        if it[0] == "l":
            seg = [(it[1].x, it[1].y), (it[2].x, it[2].y)]
        elif it[0] == "c":
            p0, c1, c2, p3 = it[1:5]
            seg = [((1 - t) ** 3 * p0.x + 3 * (1 - t) ** 2 * t * c1.x + 3 * (1 - t) * t * t * c2.x + t ** 3 * p3.x, (1 - t) ** 3 * p0.y + 3 * (1 - t) ** 2 * t * c1.y + 3 * (1 - t) * t * t * c2.y + t ** 3 * p3.y) for t in [i / 16 for i in range(17)]]
        elif it[0] == "re":
            R = it[1]
            seg = [(R.x0, R.y0), (R.x1, R.y0), (R.x1, R.y1), (R.x0, R.y1), (R.x0, R.y0)]
        else:
            continue
        P += seg if not P else seg[1:]
    return P


def _r3b_ferme(d):
    P = _r3b_pts_path(d)
    if len(P) < 3:
        return None
    if d.get("closePath") or any(it[0] == "re" for it in d["items"]) or (abs(P[0][0] - P[-1][0]) < 0.5 and abs(P[0][1] - P[-1][1]) < 0.5):
        return P
    return None


def _r3b_aire(P):
    return abs(sum(P[i][0] * P[(i + 1) % len(P)][1] - P[(i + 1) % len(P)][0] * P[i][1] for i in range(len(P)))) / 2


def _r3b_dans_poly(P, x, y):
    c = False
    n = len(P)
    for i in range(n):
        x1, y1 = P[i]
        x2, y2 = P[(i + 1) % n]
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
            c = not c
    return c


def _r3b_frac_boite_dans_poly(P, bb, k=2):
    """aire(boîte bb ∩ polygone P) / aire(P), par rastérisation à k px/pt."""
    x0 = min(p[0] for p in P); y0 = min(p[1] for p in P); x1 = max(p[0] for p in P); y1 = max(p[1] for p in P)
    W = max(1, int((x1 - x0) * k) + 2); H = max(1, int((y1 - y0) * k) + 2)
    im = _Im247.new("1", (W, H), 0)
    _Id249.Draw(im).polygon([((x - x0) * k, (y - y0) * k) for x, y in P], fill=1)
    M = _np247.array(im, bool); tot = int(M.sum())
    if tot == 0:
        return None
    bx0 = int((bb[0] - x0) * k); bx1 = int((bb[2] - x0) * k) + 1; by0 = int((bb[1] - y0) * k); by1 = int((bb[3] - y0) * k) + 1
    Z = _np247.zeros_like(M)
    Z[max(0, by0):max(0, by1), max(0, bx0):max(0, bx1)] = True
    return float((M & Z).sum() / tot)


def _r3b_traces(dessins, cad):
    """les tracés de la page (à contour) dont la boîte touche le cadre."""
    out = []
    for d in dessins:
        r = d["rect"]
        if d.get("color") is None:
            continue
        if r.x1 > cad[0] and r.x0 < cad[2] and r.y1 > cad[1] and r.y0 < cad[3]:
            out.append(d)
    return out


def _r3b_cadres_fermes(dessins, cad):
    """tracés fermés du cadre de l'image, d'aire au plus 90 % de celle du cadre ; chacun avec sa boîte englobante."""
    A_cad = (cad[2] - cad[0]) * (cad[3] - cad[1]); out = []
    for d in _r3b_traces(dessins, cad):
        P = _r3b_ferme(d)
        if P is None or len(P) < 3:
            continue
        a = _r3b_aire(P)
        if a <= 0.90 * A_cad:
            out.append(dict(P=P, aire=a, bb=(min(p[0] for p in P), min(p[1] for p in P), max(p[0] for p in P), max(p[1] for p in P))))
    return out


def _r3b_dedans(F, x, y):
    """le point est-il dans le tracé F ? Un point à plus de _R3B_EPS de la boîte englobante n'y est pas ; sinon, le test du banc."""
    b = F["bb"]
    if x < b[0] - _R3B_EPS or x > b[2] + _R3B_EPS or y < b[1] - _R3B_EPS or y > b[3] + _R3B_EPS:
        return False
    return _r3b_dans_poly(F["P"], x, y)


def _r3b_titre_de_cadre(l, Lb, cadres):
    """TC(l) du banc : un tracé fermé (aire <= 90 % du cadre) contient le centre de l, aucune autre ligne à lettres n'a son centre dedans, et la
    boîte de l recouvre au plus 50 % du tracé."""
    cx = (l["pb"][0] + l["pb"][2]) / 2; cy = (l["pb"][1] + l["pb"][3]) / 2
    for F in cadres:
        if not _r3b_dedans(F, cx, cy):
            continue
        autres = [m for m in Lb if m is not l and m["chars"] and _r3b_dedans(F, (m["pb"][0] + m["pb"][2]) / 2, (m["pb"][1] + m["pb"][3]) / 2)]
        if autres:
            continue
        fr = _r3b_frac_boite_dans_poly(F["P"], l["pb"])
        if fr is not None and fr <= 0.5:
            return True
    return False


def _r3b_lignes(blocs, cad):
    """les lignes de texte (rawdict) dont le centre est dans le cadre : texte (sans blancs aux bouts), boîte `pb`, boîtes des lettres non blanches."""
    L = []
    for b in blocs:
        for l in b.get("lines", []):
            x0, y0, x1, y1 = l["bbox"]
            if cad[0] <= (x0 + x1) / 2 <= cad[2] and cad[1] <= (y0 + y1) / 2 <= cad[3]:
                t_ = "".join(q["c"] for sp in l["spans"] for q in sp["chars"])
                ch = [[q["bbox"][0], q["bbox"][1], q["bbox"][2], q["bbox"][3]] for sp in l["spans"] for q in sp["chars"] if q["c"].strip()]
                L.append({"texte": t_.strip(), "pb": [x0, y0, x1, y1], "chars": ch})
    return L


def _r3b_lignes_de(blocs, dessins, cad):
    """la liste rendue pour un cadre : [{texte, boite, lettres, titre_de_cadre}, …] ; `dessins()` rend les tracés de la page (lus à la première demande)."""
    Lb = _r3b_lignes(blocs, cad)
    if not Lb:
        return []
    cadres = _r3b_cadres_fermes(dessins(), cad)
    return [{"texte": l["texte"], "boite": l["pb"], "lettres": l["chars"], "titre_de_cadre": bool(cadres) and _r3b_titre_de_cadre(l, Lb, cadres)} for l in Lb]


def _r3b_lignes_images(content, result):
    """ajoute `lignes_image` ({cadre: […], entier: […] si l'image a un cadre entier}) à chaque image du résultat, d'après le PDF d'origine ;
    une panne est dite dans `lignes_image_erreur`, jamais levée vers /parse."""
    try:
        doc = fitz.open(stream=content, filetype="pdf")
    except Exception as e:
        result["lignes_image_erreur"] = "ouverture : " + (str(e) or type(e).__name__)[:160]
        return result
    try:
        par_page = {}
        for im in result.get("images") or []:
            par_page.setdefault(im.get("page"), []).append(im)
        for p1, ims in par_page.items():
            try:
                page = doc[p1 - 1]
                blocs = page.get_text("rawdict")["blocks"]
                lus = []

                def dessins(_page=page, _lus=lus):
                    if not _lus:
                        _lus.append(_page.get_drawings())
                    return _lus[0]
            except Exception as e:
                result.setdefault("lignes_image_erreur", "page %s : %s" % (p1, (str(e) or type(e).__name__)[:120]))
                continue
            for im in ims:
                try:
                    champs = {"cadre": im["cadre"]}
                    if im.get("entier_cadre"):
                        champs["entier"] = im["entier_cadre"]
                    im["lignes_image"] = {k: _r3b_lignes_de(blocs, dessins, cad) for k, cad in champs.items()}
                except Exception as e:
                    result.setdefault("lignes_image_erreur", "image %s : %s" % (im.get("index"), (str(e) or type(e).__name__)[:120]))
    except Exception as e:
        result["lignes_image_erreur"] = (str(e) or type(e).__name__)[:200]
    finally:
        doc.close()
    return result


# ═══ v2.69 — LES ENDROITS DU SIGNE : B-N2 (F021) AVEC Wc2 (F028), LA PORTE E3, SANS PLANCHER (DEC-2690, DEC-2691, DEC-2708, DEC-2732, DEC-3271) ═══
# Là où la réponse attendue est un signe écrit entre deux nombres, chacun dans son propre cadre (F-2C), la feuille que reçoit l'élève porte les deux
# cadres ôtés (le tracé rendu seul, dilaté de 2 px, hors de l'encre des nombres : Wc2), la marque (« . », « .. ») effacée, et une seule case dans le
# blanc entre les deux nombres (1 pt, gris 107). Portage de outils/b_n2_256.py (le champ, second passage ; les gestes Z et K), de s5_porte_252.py
# (voisins, dessine_case) et de outils260/ra_c_260.py (gestes_v, Wc2) ; rien n'est changé d'une règle. Ajouts : la porte E3 (sa réponse : une marque
# qui a sa propre cellule, un cadre qui la contient sans les nombres, est « la case déjà là » : rien ne change) ; aucun plancher (E4) ; une
# abstention comptée là où le banc criait (F-0C). Le porteur : les gestes en données simples (`endroits_signe`), rejoués par /parse après sa
# recomposition. Les F-1C restent à leur forme (N3-e, ER). Le pourquoi et les mesures : JOURNAL BACKEND v2.69.
_V269_BN2_MM = 72 / 25.4
_V269_BN2_K = 300 / 72          # le rendu d'un tracé seul et de l'encre (300 ppp, comme au banc)
_V269_BN2_GRIS = (107, 107, 107)


def _v269_rendu_seul(page, dessins, k):
    """Les tracés donnés redessinés seuls (contour, couleur, épaisseur) sur une page blanche de la taille de la source, rendus à l'échelle k :
    pixels < 250 (le rendu d'un tracé seul du banc, outils260/ra_base_260.py ; G2 s'en sert pour tous les tracés de la page)."""
    D = fitz.open(); p = D.new_page(width=page.rect.width, height=page.rect.height)
    for d in dessins:
        sh = p.new_shape()
        for it in d["items"]:
            if it[0] == "l":
                sh.draw_line(it[1], it[2])
            elif it[0] == "c":
                sh.draw_bezier(it[1], it[2], it[3], it[4])
            elif it[0] == "re":
                sh.draw_rect(it[1])
            elif it[0] == "qu":
                sh.draw_quad(it[1])
        sh.finish(color=d.get("color") or (0, 0, 0), fill=None, width=d.get("width") or 0.75, closePath=bool(d.get("closePath")),
                  lineCap=max(d.get("lineCap") or [0]) if isinstance(d.get("lineCap"), (list, tuple)) else (d.get("lineCap") or 0), lineJoin=d.get("lineJoin") or 0)
        sh.commit()
    pm = p.get_pixmap(matrix=fitz.Matrix(k, k), alpha=False)
    return _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].min(axis=2) < 250


def _v269_bn2_voisins(mots, i):
    """Le mot immédiatement à gauche et à droite de mots[i] sur la même ligne (recouvrement vertical ≥ 50 % de la plus petite hauteur)."""
    m = mots[i]; hm = m[3] - m[1]; g = d = None
    for j, w in enumerate(mots):
        if j == i:
            continue
        hw = w[3] - w[1]
        if min(m[3], w[3]) - max(m[1], w[1]) < 0.5 * min(hm, hw):
            continue
        cw, cm = (w[0] + w[2]) / 2, (m[0] + m[2]) / 2
        if cw < cm and w[2] <= m[0] + 0.5 * (m[2] - m[0]):
            if g is None or w[2] > mots[g][2]:
                g = j
        elif cw > cm and w[0] >= m[2] - 0.5 * (m[2] - m[0]):
            if d is None or w[0] < mots[d][0]:
                d = j
    return g, d


def _v269_bn2_encre(page, bb):
    """La boîte serrée de l'encre (gris < 200, 300 ppp) dans la boîte bb ; bb elle-même si aucune encre."""
    k = _V269_BN2_K; pm = page.get_pixmap(matrix=fitz.Matrix(k, k), clip=fitz.Rect(bb), alpha=False)
    A = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].min(axis=2) < 200; ys, xs = _np247.nonzero(A)
    if not len(xs):
        return list(bb)
    return [bb[0] + xs.min() / k, bb[1] + ys.min() / k, bb[0] + (xs.max() + 1) / k, bb[1] + (ys.max() + 1) / k]


def _v269_bn2_endroits(page, dessins):
    """Les endroits d'un signe entre deux nombres d'une page (le champ du second passage de B-N2) : la marque « . » ou « .. », ses deux voisins
    à chiffre à moins de 3 hauteurs ; la famille F-2C (chaque nombre dans son cadre, qui ne contient pas la marque), F-1C (un cadre les contient
    tous trois), F-0C ; `cellule` vrai si un cadre contient la marque sans aucun des deux nombres (la porte E3)."""
    ws = page.get_text("words")
    cad = [(j, d) for j, d in enumerate(dessins) if d.get("rect") and d.get("color") is not None and 6 <= d["rect"].width <= 300 and 6 <= d["rect"].height <= 80
           and any(it[0] in ("l", "c", "re", "qu") for it in d["items"])]
    ctr = lambda b: ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
    dans = lambda R, p: R.x0 <= p[0] <= R.x1 and R.y0 <= p[1] <= R.y1
    E = []
    for i, w in enumerate(ws):
        if not re.fullmatch(r"\.{1,2}", w[4].strip()):
            continue
        h = w[3] - w[1]; gi, di = _v269_bn2_voisins(ws, i)
        if gi is None or di is None:
            continue
        gw, dw = ws[gi], ws[di]
        if (w[0] - gw[2]) > 3 * h or (dw[0] - w[2]) > 3 * h or (w[2] - w[0]) > 3 * h:
            continue
        if not (re.search(r"\d", gw[4]) and re.search(r"\d", dw[4])):
            continue
        cm = ctr(w[:4]); cg = ctr(gw[:4]); cd = ctr(dw[:4])
        fg = [(j, d) for j, d in cad if dans(d["rect"], cg) and not dans(d["rect"], cm)]
        fd = [(j, d) for j, d in cad if dans(d["rect"], cd) and not dans(d["rect"], cm)]
        f1 = [(j, d) for j, d in cad if dans(d["rect"], cg) and dans(d["rect"], cd) and dans(d["rect"], cm)]
        fam = "F-2C" if (fg and fd) else ("F-1C" if f1 else "F-0C")
        cellule = any(dans(d["rect"], cm) and not dans(d["rect"], cg) and not dans(d["rect"], cd) for j, d in cad)
        E.append(dict(page=page.number + 1, marque=list(w[:4]), g=list(gw[:4]), d=list(dw[:4]), tg=gw[4], td=dw[4], fam=fam, cellule=cellule,
                      cg=(min(fg, key=lambda x: x[1]["rect"].get_area()) if fg else None), cd=(min(fd, key=lambda x: x[1]["rect"].get_area()) if fd else None),
                      cf=(min(f1, key=lambda x: x[1]["rect"].get_area()) if f1 else None)))
    return E


def _v269_bn2_gestes(page, e):
    """Les gestes d'un endroit F-2C (gestes_v de ra_c_260.py, variante Wc2), en données simples : W2 pour chaque cadre (l'indice de son tracé dans
    la page, la garde de l'encre des deux nombres), Z (la boîte d'encre de la marque ± 0,6 pt), K (la case : de l'encre du nombre de gauche + 2 mm à
    celle du nombre de droite − 2 mm, sur la hauteur commune des cadres − 1 mm ; ± 0,3 pt si moins de 1 pt)."""
    MM = _V269_BN2_MM
    eg = _v269_bn2_encre(page, e["g"]); ed = _v269_bn2_encre(page, e["d"]); em = _v269_bn2_encre(page, e["marque"])
    G = [dict(geste="W2", trace=j, garde=[eg, ed]) for j, _d in (e["cg"], e["cd"])]
    G.append(dict(geste="Z", boite=[em[0] - 0.6, em[1] - 0.6, em[2] + 0.6, em[3] + 0.6]))
    rg, rd = e["cg"][1]["rect"], e["cd"][1]["rect"]
    y0 = max(rg.y0, rd.y0) + MM; y1 = min(rg.y1, rd.y1) - MM
    k_ = [eg[2] + 2 * MM, y0, ed[0] - 2 * MM, y1]
    if k_[2] - k_[0] < 1.0:
        k_ = [eg[2] + 0.3, y0, ed[0] - 0.3, y1]
    G.append(dict(geste="K", boite=k_))
    return G


def _v269_bn2_appliquer(C, org, k, page, dessins, G, memo):
    """Les gestes, dans l'ordre, sur une image (tableau RVB) dont l'origine en pixels de page est org, à l'échelle k (le crochet du banc)."""
    A = _np247.array(C, copy=True)
    for x in G:
        if x["geste"] == "S":
            # ligne 5.3 (D2-b) : les pixels de la source recopiés dans la boîte d'encre d'un « . » du réseau (crochet S de b_sn_256.py)
            rr = x["boite"]; pm = page.get_pixmap(matrix=fitz.Matrix(k, k), clip=fitz.Rect(rr), alpha=False)
            B_ = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
            ir = (fitz.Rect(rr) * fitz.Matrix(k, k)).irect; sx0, sy0 = ir.x0 - org[0], ir.y0 - org[1]
            H, Wd = A.shape[:2]; h_, w_ = min(B_.shape[0], H - sy0), min(B_.shape[1], Wd - sx0)
            if h_ > 0 and w_ > 0 and sx0 >= 0 and sy0 >= 0:
                A[sy0:sy0 + h_, sx0:sx0 + w_] = B_[:h_, :w_]
        elif x["geste"] == "W2":
            cle = (page.number, x["trace"], k)
            if cle not in memo:
                # v2.69 — R9 (B), DEC-3279 : le masque rendu à l'échelle k de l'image (une tête : k = 300/72, le même qu'avant ; un morceau seul : son échelle réelle)
                memo[cle] = _v269_g2_dilate(_v269_rendu_seul(page, [dessins[x["trace"]]], k), 2)
            Mp = memo[cle]; H, Wd = A.shape[:2]
            sub = _np247.zeros((H, Wd), bool); y0, x0 = org[1], org[0]
            ys0, xs0 = max(0, y0), max(0, x0); ys1, xs1 = min(Mp.shape[0], y0 + H), min(Mp.shape[1], x0 + Wd)
            if ys1 > ys0 and xs1 > xs0:
                sub[ys0 - y0:ys1 - y0, xs0 - x0:xs1 - x0] = Mp[ys0:ys1, xs0:xs1]
            for bb in x["garde"]:
                gx0 = int(_m266.floor(bb[0] * k)) - org[0] - 1; gy0 = int(_m266.floor(bb[1] * k)) - org[1] - 1
                gx1 = int(_m266.ceil(bb[2] * k)) - org[0] + 1; gy1 = int(_m266.ceil(bb[3] * k)) - org[1] + 1
                sub[max(0, gy0):max(0, gy1), max(0, gx0):max(0, gx1)] = False
            A[sub] = 255
        elif x["geste"] == "Z":
            r = x["boite"]
            x0 = int(_m266.floor(r[0] * k)) - org[0]; y0 = int(_m266.floor(r[1] * k)) - org[1]; x1 = int(_m266.ceil(r[2] * k)) - org[0]; y1 = int(_m266.ceil(r[3] * k)) - org[1]
            A[max(0, y0):max(0, y1), max(0, x0):max(0, x1)] = 255
        elif x["geste"] == "K":
            r = x["boite"]; im = _Im247.fromarray(A); dr = _Id249.Draw(im)
            x0 = round(r[0] * k) - org[0]; y0 = round(r[1] * k) - org[1]; x1 = round(r[2] * k) - org[0]; y1 = round(r[3] * k) - org[1]
            dr.rectangle([x0, y0, x1, y1], outline=_V269_BN2_GRIS, width=max(1, int(round(1.0 * k))))
            A = _np247.array(im)
        elif x["geste"] == "W2e":
            # ligne 5.9 (ER) : le cadre ôté, protégé des tracés voisins, leurs croisements retracés (crochet_a12 du banc, branche W2 avec _autres)
            H, Wd = A.shape[:2]; Au = [dessins[i] for i in x["autres"]]
            Mu = _v269_g2_dilate(_v269_rendu_seul(page, [dessins[x["trace"]]], k), 2); M = Mu.copy()
            if Au:
                Ka = _v269_rendu_seul(page, Au, k); M = M & ~Ka
            A[_l59_croque(M, org, H, Wd)] = 255
            if Au:
                cross = _l59_croque(Mu, org, H, Wd) & _l59_croque(Ka, org, H, Wd)
                Col = _l59_croque(_l59_rendu_couleur(page, Au, k), org, H, Wd, 255, _np247.uint8); A[cross] = Col[cross]
        elif x["geste"] == "Ee":
            # ligne 5.9 (ER) : l'intérieur du cadre effacé, sauf ce qu'un tracé voisin y laisse (crochet_a12 du banc, branche E)
            H, Wd = A.shape[:2]; r_ = x["rect"]
            ex0 = max(0, int(_m266.floor(r_[0] * k)) - org[0]); ey0 = max(0, int(_m266.floor(r_[1] * k)) - org[1])
            ex1 = min(Wd, int(_m266.ceil(r_[2] * k)) - org[0]); ey1 = min(H, int(_m266.ceil(r_[3] * k)) - org[1])
            keep = _l59_protege(A, page, [dessins[i] for i in x["autres"]], org, k)
            sub = _np247.zeros((H, Wd), bool); sub[ey0:ey1, ex0:ex1] = True; A[sub & ~keep] = 255
    return A


def _v269_bn2_poser(images, faits, doc, voies):
    """Pose les gestes portés sur les images : la tête d'une figure (son `entier_feuille`, au rendu de feuille) ou le morceau seul (son
    `data_feuille`) ; les gestes d'une même image dans l'ordre où ils ont été trouvés."""
    m = _V238_PPP_FEUILLE / 72.0; memo = {}; dessins = {}; par = {}
    for f in faits:
        if f["voie"] in voies:
            par.setdefault((f["page"], f["figure"], f["voie"], f.get("index")), []).extend(f["gestes"])
    for (pn, fig, voie, idx), G in par.items():
        if voie == "tête":
            cible = [im for im in images if im.get("page") == pn and im.get("figure") == fig and im.get("entier_feuille") and im.get("entier_cadre")]
            if not cible:
                continue
            im = cible[0]; ch = "entier_feuille"; ir = (fitz.Rect(im["entier_cadre"]) * fitz.Matrix(m, m)).irect; org, k = (ir.x0, ir.y0), m
            A = _v260_rgb(im[ch])
        else:
            cible = [im for im in images if im.get("page") == pn and im.get("index") == idx and im.get("data_feuille")]
            if not cible:
                continue
            im = cible[0]; ch = "data_feuille"; A = _v260_rgb(im[ch]); c = im["cadre"]
            k = A.shape[1] / (c[2] - c[0]); org = (int(round(c[0] * k)), int(round(c[1] * k)))
        pg = doc[pn - 1]
        if pn not in dessins:
            dessins[pn] = pg.get_drawings()
        im[ch] = _v260_png(_v269_bn2_appliquer(A, org, k, pg, dessins[pn], G, memo))


# ═══ ligne 5.9 du plan — ER « EFFACER ET RÉÉCRIRE », LA PART DU SERVEUR (DEC-2812, DEC-2818, DEC-2805, DEC-3271 C4 ; forme « ER natif » d'A32) ═══
# Là où la réponse attendue est un signe écrit entre deux nombres DANS UN SEUL CADRE (F-1C : « 8 + 2 [?] 7 + 3 » ; la 032 et la 050), la feuille de l'élève
# ne garde pas le calcul de la source en image : le serveur EFFACE le calcul de l'image entière de feuille de la figure (le cadre de l'expression ôté,
# son intérieur effacé sauf ce qu'un tracé voisin y laisse) et le DÉCLARE (`calculs_er` : le texte de chaque terme, sa place, sa taille, la place de
# la case, la police de la source) ; c'est la PAGE qui réécrit le calcul en texte, dans la police de la feuille, et pose la case. Portage de
# agents273/A27_N3e4/a27_bloc.py (er_texte, er_layout_all, protege_voisins, _ink), d'a27_banc.py (dessins_autres, prepare, le crochet : gestes W2 protégé
# et E), d'outils260/ra_a_260.py (plan, masque_px), d'agents274/A32_r/a32_mkbanc.py (gestes_fig_ERn, decl_ERn) et du passage d'a32_analyse.py qui fait
# `calculs_er` ; même logique, mêmes seuils (TOL 40, 15,3 mm, 2 mm, 14 pt au moins), noms en `_l59_`. CE QUI CHANGE PAR RAPPORT AU BANC, nommé :
#   · les gestes sont des données simples (`W2e` : l'indice du tracé du cadre et ceux des tracés voisins ; `Ee` : le rectangle et les voisins), rejouées
#     par /parse après sa recomposition (`endroits_signe`), comme B-N2 ; les helpers que le serveur a déjà remplacent leurs sosies du banc
#     (`_v269_rendu_seul`, `_v269_g2_dilate`, `_v269_bn2_encre`, `_V269_BN2_MM`) ;
#   · les calculs de mesure d'A27 que ni les gestes ni la déclaration ne lisent ne sont pas portés (libre, n30, hy du plan, m0, m1, split, ink_y, new_end) ;
#   · la police est celle du woff2 gardé (Andika v27 latin 400), lue dans `police/Andika_400_mem.ttf` à côté du serveur : un fichier de plus à déployer avec lui ;
#     s'il manque, ou si une chaîne ne s'écrit pas, le bloc entier s'abstient et le dit (`endroits_signe.er.erreur`) : rien n'est effacé sans être déclaré ;
#   · la voie « morceau seul » (une expression dont la figure n'a pas d'image entière) n'est pas mesurée pour ER natif : abstention comptée, rien n'est effacé ;
#   · la porte E3 de la 2.69 précède le tri par famille : elle écarte aussi les F-1C qui ont leur cellule.
# COUPLAGE (règle du 09.08) : sans la page qui lit `calculs_er`, le calcul effacé ne serait écrit nulle part : cette forme ne se livre qu'avec elle.
# Le pourquoi et les mesures : JOURNAL BACKEND, entrée de la version qui porte ce geste.
_L59_TOL = 40                                                      # le test de couleur de protège-voisins (TOL du banc)
_L59_CASE = 15.3 * _V269_BN2_MM                                    # la grande case (CASE du banc : 15,3 mm)
_L59_MARGE = 2 * _V269_BN2_MM                                      # les marges texte-case (MARGE du banc : 2 mm)
_L59_POLICE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "police", "Andika_400_mem.ttf")
_L59_INK = {}


def _l59_ink(chaine, taille):
    """Décalages de l'encre (gris < 200) de la chaîne en Andika 400, en pt, depuis son origine : (x0, y0, x1, y1) ; rendu à 300 ppp sur une page blanche (_ink du banc)."""
    cle = (chaine, round(taille, 3))
    if cle not in _L59_INK:
        kk = 300 / 72; d_ = fitz.open(); pg = d_.new_page(width=300, height=80)
        pg.insert_text((20, 50), chaine, fontsize=taille, fontname="Andika", fontfile=_L59_POLICE, color=(0, 0, 0))
        pm = pg.get_pixmap(matrix=fitz.Matrix(kk, kk), alpha=False)
        a = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].min(axis=2) < 200
        d_.close()
        ys, xs = _np247.nonzero(a)
        if not len(xs):
            raise RuntimeError("la chaîne %r ne s'écrit pas (aucune encre)" % chaine)
        _L59_INK[cle] = (xs.min() / kk - 20, ys.min() / kk - 50, (xs.max() + 1) / kk - 20, (ys.max() + 1) / kk - 50)
    return _L59_INK[cle]


def _l59_rendu_couleur(page, dessins, k):
    """Les tracés redessinés seuls EN COULEUR sur une page blanche, rendus à l'échelle k : un tableau RVB (rendu_couleur du banc)."""
    D_ = fitz.open(); p = D_.new_page(width=page.rect.width, height=page.rect.height)
    for d in dessins:
        sh = p.new_shape()
        for it in d["items"]:
            if it[0] == "l":
                sh.draw_line(it[1], it[2])
            elif it[0] == "c":
                sh.draw_bezier(it[1], it[2], it[3], it[4])
            elif it[0] == "re":
                sh.draw_rect(it[1])
            elif it[0] == "qu":
                sh.draw_quad(it[1])
        sh.finish(color=d.get("color") or (0, 0, 0), fill=None, width=d.get("width") or 0.75, closePath=bool(d.get("closePath")),
                  lineCap=max(d.get("lineCap") or [0]) if isinstance(d.get("lineCap"), (list, tuple)) else (d.get("lineCap") or 0), lineJoin=d.get("lineJoin") or 0)
        sh.commit()
    pm = p.get_pixmap(matrix=fitz.Matrix(k, k), alpha=False)
    out = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].copy()
    D_.close()
    return out


def _l59_croque(Mp, org, H, Wd, fill=False, dt=bool):
    """La partie d'un tableau de la page entière (origine des pixels org) qui tombe dans une image de H x Wd pixels (croque du banc)."""
    sub = _np247.full((H, Wd) + Mp.shape[2:], fill, dt); y0, x0 = org[1], org[0]; ys0, xs0 = max(0, y0), max(0, x0); ys1, xs1 = min(Mp.shape[0], y0 + H), min(Mp.shape[1], x0 + Wd)
    if ys1 > ys0 and xs1 > xs0:
        sub[ys0 - y0:ys1 - y0, xs0 - x0:xs1 - x0] = Mp[ys0:ys1, xs0:xs1]
    return sub


def _l59_masque_px(page, dessins, org, k, H, Wd, r=0):
    """Le masque des tracés donnés (rendus seuls, dilatés de r px) aux coordonnées de l'image (masque_px de ra_a_260.py)."""
    Mp = _v269_rendu_seul(page, dessins, k)
    if r:
        Mp = _v269_g2_dilate(Mp, r)
    return _l59_croque(Mp, org, H, Wd)


def _l59_protege(A, page, autres, org, k):
    """Les pixels qu'un tracé voisin (qui franchit la bande) a le droit de garder : dans son masque rendu seul (halo 0) ET de sa couleur à TOL (protege_voisins du banc)."""
    H, Wd = A.shape[:2]; K = _np247.zeros((H, Wd), bool)
    if not autres:
        return K
    Af = A.astype(float)
    for da in autres:
        col = da.get("color")
        if col is None:
            continue
        v = 255.0 - _np247.array(col, float) * 255.0; Mi = _l59_masque_px(page, [da], org, k, H, Wd, 0)
        a_ = _np247.clip(((255.0 - Af) @ v) / max(float(v @ v), 1e-9), 0, 1.15); res = _np247.abs(Af - (255.0 - a_[..., None] * v)).max(axis=2); K |= Mi & (res <= _L59_TOL)
    return K


def _l59_plan(page, dessins, E):
    """Les expressions F-1C d'une figure, annotées, et leurs bandes (plan de ra_a_260.py) : le cadre de l'expression est celui que l'endroit a trouvé (`cf`)."""
    P_ = []
    for e in E:
        eg = _v269_bn2_encre(page, e["g"]); ed = _v269_bn2_encre(page, e["d"]); em = _v269_bn2_encre(page, e["marque"])
        cf = e.get("cf")
        if cf is None:
            P_.append(dict(e=e, nc="cadre non retrouvé")); continue
        cad = [dessins[cf[0]]]
        y0 = min(c["rect"].y0 for c in cad) - 1; y1 = max(c["rect"].y1 for c in cad) + 1
        blanc = ed[0] - eg[2]; D = max(0.0, _L59_CASE + 2 * _L59_MARGE - blanc)
        P_.append(dict(e=e, eg=eg, ed=ed, em=em, cad=cad, cad_j=[cf[0]], bande=(y0, y1), D=D, blanc=blanc))
    ok = [p for p in P_ if "nc" not in p]; bandes = []
    for p in sorted(ok, key=lambda p: p["bande"][0]):
        for b in bandes:
            if p["bande"][0] < b["y1"] and p["bande"][1] > b["y0"]:
                b["L"].append(p); b["y0"] = min(b["y0"], p["bande"][0]); b["y1"] = max(b["y1"], p["bande"][1]); break
        else:
            bandes.append(dict(y0=p["bande"][0], y1=p["bande"][1], L=[p]))
    return P_, bandes


def _l59_proche(a, b, t=0.3):
    return abs(a.x0 - b.x0) < t and abs(a.y0 - b.y0) < t and abs(a.x1 - b.x1) < t and abs(a.y1 - b.y1) < t


def _l59_autres(page, band, canevas, cads):
    """Les INDICES (dans page.get_drawings()) des tracés voisins qui franchissent la bande, hors les cadres de ses expressions (dessins_autres du banc, forme C1)."""
    B_ = fitz.Rect(canevas[0], band["y0"], canevas[1], band["y1"]); drs = page.get_drawings(); out = []
    blancs = [(i, fitz.Rect(d["rect"])) for i, d in enumerate(drs) if d.get("color") and min(d["color"]) >= 0.99 and "s" in (d.get("type") or "")]
    for i, d in enumerate(drs):
        col = d.get("color")
        if col is None or min(col) >= 0.99 or "s" not in (d.get("type") or ""):
            continue
        R = fitz.Rect(d["rect"])
        if R.x1 < B_.x0 or R.x0 > B_.x1 or R.y1 < B_.y0 or R.y0 > B_.y1:
            continue
        if any(_l59_proche(R, fitz.Rect(c["rect"])) for c in cads):
            continue
        if any((c["rect"].x0 - 1 <= R.x0 and c["rect"].y0 - 1 <= R.y0 and R.x1 <= c["rect"].x1 + 1 and R.y1 <= c["rect"].y1 + 1) for c in cads):
            continue
        if any(j > i and _l59_proche(R, rj) for j, rj in blancs):
            continue
        if not (R.y0 < band["y0"] - 0.5 or R.y1 > band["y1"] + 0.5):
            continue
        m = _v269_rendu_seul(page, [d], 2.0); sub = m[int(_m266.floor(B_.y0 * 2)):int(_m266.ceil(B_.y1 * 2)), int(_m266.floor(B_.x0 * 2)):int(_m266.ceil(B_.x1 * 2))]
        if sub.any():
            out.append(i)
    return out


def _l59_texte(pg, p):
    """Lit dans le PDF le texte de l'expression (mots du cadre hors la marque), côté gauche / droit, l'origine du premier glyphe, la taille, la fin d'encre d'origine (er_texte du banc)."""
    e = p["e"]; R = fitz.Rect(p["cad"][0]["rect"]); mx = (e["marque"][0] + e["marque"][2]) / 2
    ws = [w for w in pg.get_text("words") if R.x0 - 0.5 <= (w[0] + w[2]) / 2 <= R.x1 + 0.5 and R.y0 - 0.5 <= (w[1] + w[3]) / 2 <= R.y1 + 0.5]
    est_marque = lambda w: abs(w[0] - e["marque"][0]) < 0.6 and abs(w[2] - e["marque"][2]) < 0.6 and abs(w[1] - e["marque"][1]) < 0.6
    mots = [w for w in ws if not est_marque(w)]
    Lw = sorted([w for w in mots if (w[0] + w[2]) / 2 < mx], key=lambda w: w[0]); Rw = sorted([w for w in mots if (w[0] + w[2]) / 2 > mx], key=lambda w: w[0])
    if not Lw or not Rw:
        return None, "un côté vide (mots du cadre hors la marque : %d à gauche, %d à droite)" % (len(Lw), len(Rw))
    d = pg.get_text("rawdict", clip=R + (-1, -1, 1, 1)); spans = [s for b in d["blocks"] for l in b.get("lines", []) for s in l["spans"] if any(c["c"].strip() for c in s["chars"])]
    chars = [(c, s) for s in spans for c in s["chars"] if c["c"].strip() and R.x0 - 0.5 <= (c["bbox"][0] + c["bbox"][2]) / 2 <= R.x1 + 0.5 and R.y0 - 0.5 <= (c["bbox"][1] + c["bbox"][3]) / 2 <= R.y1 + 0.5]
    if not chars:
        return None, "aucun glyphe de police dans le cadre (texte en image ?)"
    w0 = Lw[0]; cs = [(c, s) for c, s in chars if w0[0] - 0.3 <= (c["bbox"][0] + c["bbox"][2]) / 2 <= w0[2] + 0.3 and w0[1] - 0.3 <= (c["bbox"][1] + c["bbox"][3]) / 2 <= w0[3] + 0.3]
    if not cs:
        return None, "premier glyphe introuvable"
    c0, s0 = min(cs, key=lambda cs_: cs_[0]["bbox"][0]); taille = max(14.0, max(s["size"] for c, s in chars))     # DEC-2756 / DEC-2764 : 14 pt au moins jusqu'en 8H
    bbR = [min(w[0] for w in Rw), min(w[1] for w in Rw), max(w[2] for w in Rw), max(w[3] for w in Rw)]
    return dict(L=" ".join(w[4] for w in Lw), R=" ".join(w[4] for w in Rw), ox=c0["origin"][0], oy=c0["origin"][1], size=taille, police_source=sorted(set((s["font"], round(s["size"], 2)) for c, s in chars)),
                old_end=_v269_bn2_encre(pg, bbR)[2], cadre=[R.x0, R.y0, R.x1, R.y1], bb_first=list(w0[:4])), None


def _l59_layout(pg, P_, bandes):
    """La mise en page d'ER, par bande, de gauche à droite : décalage cumulé s ; la bande entière s'abstient si une expression s'abstient (er_layout_all du banc)."""
    out = []; MM1 = 1.0 * _V269_BN2_MM
    for b in bandes:
        shift = 0.0; ok = True; lots = []
        for p in sorted(b["L"], key=lambda p: p["eg"][0]):
            t, why = _l59_texte(pg, p)
            if t is None:
                ok = False; p["nc"] = "ER abstention (%s)" % why; continue
            a = _l59_ink(t["L"], t["size"]); r = _l59_ink(t["R"], t["size"]); oxL = t["ox"] + shift; cx0 = oxL + a[2] + _L59_MARGE; cx1 = cx0 + _L59_CASE; oxR = cx1 + _L59_MARGE - r[0]; new_end = oxR + r[2]
            hy0 = t["oy"] + min(a[1], r[1]) - MM1; hy1 = t["oy"] + max(a[3], r[3]) + MM1; depl = new_end - t["old_end"]
            p["er"] = dict(t, oxL=oxL, oxR=oxR, s=shift, new_end=new_end, ink_L=(oxL + a[0], oxL + a[2]), ink_R=(oxR + r[0], oxR + r[2]))
            p["D"] = depl - shift; p["case"] = [cx0, hy0, cx1, hy1]; shift = depl; lots.append(p)
        if not ok:
            for p in b["L"]:
                if "nc" not in p:
                    p["nc"] = "ER abstention (la bande entière : une expression de la bande s'abstient)"
            continue
        b["L"] = lots; b["growth"] = shift; out.append(b)
    return out


def _l59_decl(bandes):
    """La déclaration d'une figure : `calculs_er[figure]`, une entrée par expression, en points PDF (decl_ERn du banc, puis le passage d'a32_analyse.py ; arrondi à 4 décimales)."""
    out = []
    f = lambda v: round(float(v), 4)
    for b in bandes:
        for p in sorted(b["L"], key=lambda p: p["er"]["oxL"]):
            er = p["er"]
            ps = [list(x) for x in er["police_source"]]
            pol = ", ".join("%s %.2f pt" % tuple(x) for x in ps)
            tg, td, tail = er["L"], er["R"], f(er["size"]); oxL, oxR, oy = f(er["oxL"]), f(er["oxR"]), f(er["oy"]); iL = [f(v) for v in er["ink_L"]]; iR = [f(v) for v in er["ink_R"]]
            out.append(dict(termes=[dict(texte=tg, x=oxL, base=oy, taille=tail, encre=[iL[0], iL[1]], poids=400),
                                    dict(texte=td, x=oxR, base=oy, taille=tail, encre=[iR[0], iR[1]], poids=400)],
                            case=[f(v) for v in p["case"]],
                            source=dict(police=pol, x=f(er["ox"]), base=oy, fin_encre=f(er["old_end"]), cadre=[f(v) for v in er["cadre"]], decalage=f(er["s"]))))
    return out


def _l59_er(images, doc, E):
    """Les expressions F-1C `E` (endroits de `_v269_bn2_endroits`, la porte E3 déjà passée) : rend {faits : les gestes à poser sur les têtes, calculs : la déclaration,
    bilan : ce qui a été fait, ce qui s'est abstenu et pourquoi}. Ne lève jamais : une panne efface tout le bloc et se dit dans bilan['erreur']."""
    bilan = dict(expressions=len(E), traitees=0, abstentions=[], erreur=None)
    try:
        ctr = lambda b: ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
        dans = lambda c, p: c[0] <= p[0] <= c[2] and c[1] <= p[1] <= c[3]
        par = {}
        for e in E:
            cg, cd = ctr(e["g"]), ctr(e["d"])
            im = next((x for x in images if x.get("page") == e["page"] and dans(x["cadre"], cg) and dans(x["cadre"], cd)), None) or \
                next((x for x in images if x.get("page") == e["page"] and (dans(x["cadre"], cg) or dans(x["cadre"], cd))), None)
            if im is None:
                bilan["abstentions"].append(dict(g=e["tg"], d=e["td"], page=e["page"], raison="aucune image ne porte l'expression")); continue
            if not [x for x in images if x.get("figure") == im.get("figure") and x.get("page") == im.get("page") and x.get("entier_feuille")]:
                bilan["abstentions"].append(dict(g=e["tg"], d=e["td"], page=e["page"], raison="morceau seul : la voie n'est pas mesurée pour ER natif")); continue
            par.setdefault((e["page"], im.get("figure")), []).append(e)
        faits = []; calc = {}
        for (pn, fig), EE in par.items():
            pg = doc[pn - 1]; dessins = pg.get_drawings()
            P_, bandes = _l59_plan(pg, dessins, EE)
            ext0 = max([sum(p["D"] for p in b["L"]) for b in bandes] + [0])
            t = [x for x in images if x.get("page") == pn and x.get("figure") == fig and x.get("entier_feuille") and x.get("entier_cadre")][0]; c0 = list(t["entier_cadre"])
            for b in bandes:
                cads = [c for p in b["L"] for c in p["cad"]]; b["autres"] = _l59_autres(pg, b, (c0[0], c0[2] + ext0), cads)
            bandes = _l59_layout(pg, P_, bandes)
            for p in P_:
                if "nc" in p:
                    bilan["abstentions"].append(dict(g=p["e"]["tg"], d=p["e"]["td"], page=pn, raison=str(p["nc"])))
            if not bandes:
                continue
            G = []
            for b in bandes:
                for p in b["L"]:
                    for c, j in zip(p["cad"], p["cad_j"]):
                        R = c["rect"]
                        G.append(dict(geste="W2e", trace=j, autres=list(b["autres"])))
                        G.append(dict(geste="Ee", rect=[R.x0, R.y0, R.x1, R.y1], autres=list(b["autres"])))
            faits.append(dict(page=pn, figure=fig, index=None, voie="tête", termes=[x for b in bandes for p in b["L"] for x in (p["e"]["tg"], p["e"]["td"])], gestes=G))
            calc.setdefault(fig, []).extend(_l59_decl(bandes))
            bilan["traitees"] += sum(len(b["L"]) for b in bandes)
        return dict(faits=faits, calculs=calc, bilan=bilan)
    except Exception as ex:
        bilan["erreur"] = "%s: %s" % (type(ex).__name__, ex); bilan["traitees"] = 0
        return dict(faits=[], calculs={}, bilan=bilan)


def _v269_bn2(images, doc):
    """Dans parse_pdf, après la composition des images entières : les endroits F-2C reçoivent leurs gestes (sur la tête ou le morceau seul de la
    figure qui porte les deux nombres) ; rend {gestes : la liste portée, comptes : par famille, porte E3, sans image}."""
    # ligne 5.9 (ER) : les endroits F-1C sont mis de côté (er_E), puis traités par _l59_er avant la pose
    faits = []; cpt = {"F-2C": 0, "F-1C": 0, "F-0C": 0, "porte_E3": 0, "sans_image": 0}; er_E = []
    ctr = lambda b: ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
    dans = lambda c, p: c[0] <= p[0] <= c[2] and c[1] <= p[1] <= c[3]
    for pg in doc:
        for e in _v269_bn2_endroits(pg, pg.get_drawings()):
            if e["cellule"]:
                cpt["porte_E3"] += 1
                continue
            cpt[e["fam"]] += 1
            if e["fam"] == "F-1C":
                er_E.append(e)
            if e["fam"] != "F-2C":
                continue
            cg, cd = ctr(e["g"]), ctr(e["d"])
            im = next((x for x in images if x.get("page") == e["page"] and dans(x["cadre"], cg) and dans(x["cadre"], cd)), None) or \
                next((x for x in images if x.get("page") == e["page"] and (dans(x["cadre"], cg) or dans(x["cadre"], cd))), None)
            if im is None:
                cpt["sans_image"] += 1
                continue
            tete = [x for x in images if x.get("figure") == im.get("figure") and x.get("page") == im.get("page") and x.get("entier_feuille")]
            faits.append(dict(page=e["page"], figure=im.get("figure"), index=(None if tete else im.get("index")), voie=("tête" if tete else "morceau"),
                              termes=[e["tg"], e["td"]], gestes=_v269_bn2_gestes(pg, e)))
    # ligne 5.9 (ER) : les F-1C reçoivent leurs gestes ; la déclaration part dans `calculs_er`, que parse_pdf sort du résultat
    er = _l59_er(images, doc, er_E) if er_E else None
    if er:
        faits += er["faits"]
    _v269_bn2_poser(images, faits, doc, ("tête", "morceau"))
    # 2.10 pour ER (DEC-2902, DEC-3342, DEC-3344) : une image entière de feuille qu'ER a rendue blanche est marquée `entier_feuille_vide` ;
    # la page ne l'affiche pas et garde les calculs réécrits (`vide` ne convient pas : il retirerait la figure et ses calculs)
    if er:
        for _f in er["faits"]:
            for _im in images:
                if _im.get("page") == _f["page"] and _im.get("figure") == _f["figure"] and _im.get("entier_feuille"):
                    if not ((255 - _v260_rgb(_im["entier_feuille"]).min(axis=2).astype(_np247.int16)) >= _V270_ECART).any():
                        _im["entier_feuille_vide"] = True
    res = {"gestes": faits, "comptes": cpt}
    if er:
        res["er"] = er["bilan"]
        if er["calculs"]:
            res["calculs_er"] = er["calculs"]
    return res


# ═══ ligne 5.3 du plan — D2-b : LE QUADRILLAGE À POINTS DE LA SOURCE EST GARDÉ (DEC-2716 ; DEC-2717 ; fiche V2_2_a-R71) ═══
# Les « . » d'un réseau (mots exactement « . », groupés par hauteur de mot à 0,5 pt, rangées et colonnes regroupées à 3 pt ; au moins 3 rangées, 3 colonnes,
# 9 points, et 60 % des cases rangée × colonne occupées : le critère de S5, écrit avant) sont de la ponctuation pour la découpe : ils sortent des dessins, et
# l'image de la figure les perd (la 117, fiche 2, exercice 4 : 5 points sur 36 présents). À la composition de l'image entière de feuille, les pixels de la source
# sont donc recopiés dans la boîte d'encre de chaque point (+ 0,6 pt), sur la tête de la figure dont le cadre entier contient le point. Portage d'outils/d2b_256.py
# (`encre`, `joue`, le crochet « S » de b_sn_256.py) et de `points_reseau` d'outils/d2_cause_256.py ; même logique, mêmes seuils. CE QUI CHANGE PAR RAPPORT AU BANC,
# nommé : les gestes sont des données simples (`S` : la boîte), rejouées par /parse après sa recomposition (`endroits_signe`), comme B-N2 ; une panne se déclare
# (`endroits_signe.d2b.erreur`) et /parse continue ; la clé `d2b` n'existe que s'il y a un réseau ; un point sans encre ou sans figure entière est compté, jamais posé.
# Le pourquoi et les mesures : JOURNAL BACKEND, entrée de la version qui porte ce geste.
def _l53_points_reseau(doc):
    """Les « . » d'un réseau : (page, boîte arrondie à 0,1 pt) de chacun (points_reseau d'outils/d2_cause_256.py)."""
    L = []
    for p in doc:
        ws = [w for w in p.get_text("words") if w[4].strip() == "."]
        par_h = {}
        for w in ws:
            par_h.setdefault(round((w[3] - w[1]) * 2) / 2, []).append(w)
        for hh, G in par_h.items():
            def grp(vals):
                vals = sorted(vals); out = []
                for v in vals:
                    if out and v - out[-1][-1] <= 3:
                        out[-1].append(v)
                    else:
                        out.append([v])
                return out
            ys = grp([(w[1] + w[3]) / 2 for w in G]); xs = grp([(w[0] + w[2]) / 2 for w in G])
            if len(ys) >= 3 and len(xs) >= 3 and len(G) >= 9 and len(G) / (len(ys) * len(xs)) >= 0.6:
                for w in G:
                    L.append((p.number + 1, [round(w[0], 1), round(w[1], 1), round(w[2], 1), round(w[3], 1)]))
    return L


def _l53_encre(page, k):
    """La boîte serrée de l'encre du point (gris < 200, rendu 300 ppp) dans la boîte du mot ; None s'il n'y a pas d'encre (encre d'outils/d2b_256.py)."""
    kk = 300 / 72; pm = page.get_pixmap(matrix=fitz.Matrix(kk, kk), clip=fitz.Rect(k), alpha=False)
    A = _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3].min(axis=2) < 200
    ys, xs = _np247.nonzero(A)
    if not len(xs):
        return None
    return [k[0] + xs.min() / kk, k[1] + ys.min() / kk, k[0] + (xs.max() + 1) / kk, k[1] + (ys.max() + 1) / kk]


def _l53_d2b(images, doc, endroits):
    """Pose les gestes « S » des points d'un réseau sur les têtes des figures et les ajoute aux `endroits` (que /parse rejoue après sa recomposition) ;
    rend le bilan (None s'il n'y a aucun réseau). Ne lève jamais : une panne se dit dans `endroits['d2b']['erreur']`."""
    bilan = dict(points=0, joues=0, sans_encre=0, sans_figure_entiere=0, erreur=None)
    try:
        Lp = _l53_points_reseau(doc)
        if not Lp:
            return None
        bilan["points"] = len(Lp); par = {}
        for pn, k in Lp:
            e = _l53_encre(doc[pn - 1], k)
            if e is None:
                bilan["sans_encre"] += 1; continue
            cx, cy = (e[0] + e[2]) / 2, (e[1] + e[3]) / 2
            t = [im for im in images if im.get("page") == pn and im.get("entier_feuille") and im.get("entier_cadre")
                 and im["entier_cadre"][0] <= cx <= im["entier_cadre"][2] and im["entier_cadre"][1] <= cy <= im["entier_cadre"][3]]
            if not t:
                bilan["sans_figure_entiere"] += 1; continue
            par.setdefault((pn, t[0].get("figure")), []).append(dict(geste="S", boite=[float(e[0] - 0.6), float(e[1] - 0.6), float(e[2] + 0.6), float(e[3] + 0.6)]))
            bilan["joues"] += 1
        faits = [dict(page=pn, figure=fig, index=None, voie="tête", termes=[], gestes=G) for (pn, fig), G in par.items()]
        _v269_bn2_poser(images, faits, doc, ("tête",))
        endroits["gestes"] = list(endroits.get("gestes") or []) + faits
    except Exception as ex:
        bilan["erreur"] = "%s: %s" % (type(ex).__name__, ex)
    endroits["d2b"] = bilan
    return bilan


def _v269_bn2_rejoue(images, endroits, doc):
    """Après une recomposition (la gomme, avec la clé) : rejoue les gestes des têtes (les `data_feuille` des morceaux seuls gardent les leurs)."""
    _v269_bn2_poser(images, (endroits or {}).get("gestes") or [], doc, ("tête",))


def parse_pdf(content: bytes, filename: str, api_key=None, corrige_pages=None):
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
    _lis2_prepare(content)
    # v2.38 — les deux chiffres de la divergence (image d'origine / photographie), comptés pendant la lecture et rendus
    # dans la réponse : un compte qui n'est pas rendu n'est pas un compte.
    _v238_compte = {}
    images = []
    full_text = []
    grilles_inventaire = []
    formes_inventaire = []
    idx = 0
    zones_libres = []          # V2.21 — cadres nus vides, déclarés en zones
    n_receptacles = 0          # V2.21 — « … + … = … » écartés de la voie image
    n_c11 = 0                  # v2.34 — cadres sans aucun dessin (C11)
    n_ancres = 0               # v2.36 — marqueurs de position posés
    n_ancres_hors = 0          # v2.36 — figures sans ancre, comptées à part
    # v2.52 — le jugement corrigé : sa règle, lue par Sonnet 5.5 ; ce que l'outil fait d'une page jugée corrigé ne change
    # pas (le `continue` ci-dessous). `corrige_pages` donné : pages déjà jugées (la symbiose sur une découpe).
    _v252_compte = _v252_compte_vide()
    corrige = set(corrige_pages) if corrige_pages is not None else _v252_pages_de_corrige(doc, api_key, _v252_compte)
    # v2.54 — l'état de la découpe de ce document, propre à ce fil ; la marque de la lecture des exercices
    _v254_etat_neuf()
    _v254_lecture = _v254_lecture_vide()
    _s245_ajouts = {}        # v2.45 — index ajouté -> grille d'origine (S1)
    _s245_grilles_pos = []   # v2.45 — grilles des pages de contenu (S2)
    _v247_compte = {"figures": 0, "morceaux": 0, "mots_attaches": 0, "tableaux": 0, "tableaux_ajoutes": 0,
                    "tableaux_non_cherches": 0, "fenetres_sans_dessin": 0, "traits_absorbes": 0,
                    "morceaux_trop_petits": 0, "pages_en_repli": 0, "derniere_erreur": ""}   # v2.47

    for pno, page in enumerate(doc):
        if (pno + 1) in corrige:      # page jugée corrigé : rien n'en part
            continue
        pw, ph = page.rect.width, page.rect.height
        page_area = pw * ph

        solids, thins, rasters, v21_zones_page, micros_page = _collect_page_regions(page)
        for _z in v21_zones_page:
            _z["page"] = pno + 1
            zones_libres.append(_z)

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
        blocs_page = None               # v2.36 — blocs de la page, pour l'ancre
        i0_page = len(images)           # v2.36 — première figure de cette page
        if GRILLES_ACTIVES:
            try:
                analyse = _grilles.analyser_page(page)
                grilles_page = analyse["grilles"]
                zones_grilles = analyse["zones_grilles"]
                texte_page, blocs_page = _grilles.texte_avec_grilles(
                    page, grilles_page, blocs_formes, rendre_blocs=True)
            except Exception:
                texte_page, blocs_page = _texte_gras(page, rendre_blocs=True)
        else:
            texte_page, blocs_page = _texte_gras(page, rendre_blocs=True)
            if blocs_formes:
                texte_page += "\n" + "\n".join(b[2] for b in sorted(blocs_formes))
                blocs_page = None       # v2.36 — ajout hors blocs : on s'abstient

        # v2.5 — les lignes fines ne partent que si la page parle de mesure ou de tracé (mots-clés ci-dessous) : sur une
        # fiche de français, les lignes de réponse aux longueurs variées traverseraient le filtre « familles ».
        # (voir JOURNAL BACKEND v2.5)
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
            # v2.33 — même doctrine en aval : le plancher des 24 pt écarte les miettes de traits ; appliqué aux images il
            # jetait des figures collées, qui gardent le plancher de 8 pt.
            _img33 = any(not (ra & r).is_empty
                         and (ra & r).get_area() >= 0.9 * min(ra.get_area(), r.get_area())
                         for ra in rasters)
            full = (r.width >= 8 and r.height >= 8) if _img33 else (
                r.width >= 24 and r.height >= 24)
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
            # v2.47 — V2.10 gardée, restreinte : un tableau qui porte un chiffre ou une étiquette est une figure ; un tableau
            # qui ne porte que des lignes de texte reste du texte.
            if dans_tableau and not any(
                    not (r & ra).is_empty
                    and (r & ra).get_area() >= 0.5 * min(r.get_area(),
                                                         ra.get_area())
                    for ra in rasters) and not _v247_porte_un_glyphe(page, r):
                continue
            keep.append(r)

        # V2.10 — le quadrillage part d'un seul tenant. Jusqu'ici ses traits
        # réguliers étaient éliminés comme du bruit et seules les zones
        # coloriées survivaient : le modèle recevait des taches sans repère,
        # et l'élève une figure sans quadrillage à compter. On ajoute donc le
        # cadre complet, et on retire les morceaux qu'il contient — ils y sont
        # déjà, en place.
        cadres_grille = set()      # v2.32 — cadres nes du module des grilles
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
            # v2.33 — LE CADRE SUIT LA SCÈNE, PAS LA TABLE. Condition : un objet (dessiné ou collé) touche la bande par le
            # dessus ou le dessous, dans son emprise horizontale. Effet : le cadre l'emporte. Jeu : 12 pt.
            _haut33 = fitz.Rect(cadre.x0 - 3, cadre.y0 - 12, cadre.x1 + 3, cadre.y0)
            _bas33 = fitz.Rect(cadre.x0 - 3, cadre.y1, cadre.x1 + 3, cadre.y1 + 12)
            for _z33 in (_haut33, _bas33):
                for _d33 in page.get_drawings():
                    _r33 = _d33.get("rect")
                    if not _r33:
                        continue
                    _o33 = fitz.Rect(_r33)
                    if _o33.width < 4 or _o33.width > 250 or _o33.height < 3:
                        continue
                    if (_o33.intersects(_z33) and _o33.x0 >= cadre.x0 - 5
                            and _o33.x1 <= cadre.x1 + 5):
                        cadre = cadre | _o33
                for _ra33 in rasters:
                    if (_ra33.intersects(_z33) and not (_ra33 & cadre).is_empty
                            and _ra33.x0 >= cadre.x0 - 5 and _ra33.x1 <= cadre.x1 + 5):
                        cadre = cadre | _ra33
            cadres_grille.add((round(cadre.x0, 1), round(cadre.y0, 1),
                               round(cadre.x1, 1), round(cadre.y1, 1)))
            keep.append(cadre)

        # ══ v2.17 — UNE FORME À MOITIÉ DANS LA ZONE Y EST PRISE ENTIÈRE ═════
        # Une forme vectorielle est absorbée par le cadre de découpe si : 40 % au moins de sa surface est déjà dans le
        # cadre (MAJORITÉ) ; elle fait 75 pt au plus (TAILLE) ; elle ne porte PAS de texte encore dehors (une étiquette
        # ou un badge est un rectangle fermé comme un rond : l'avaler imprimerait son texte deux fois).
        # Ce qu'il ne fait pas : une forme dedans à moins de 40 % reste fragmentée. (voir JOURNAL BACKEND v2.17)
        _MAX_FORME = 75.0        # côté maximal d'une forme absorbable, en points
        _PART_MINI = 0.40        # part déjà dans le cadre pour qu'elle en fasse partie
        try:
            _mots_page = [(fitz.Rect(w[:4]), w[4]) for w in page.get_text("words")]
            _formes = [fitz.Rect(_d["rect"]) for _d in page.get_drawings()]
            _formes = [_r for _r in _formes
                       if 1 < _r.width <= _MAX_FORME and 1 < _r.height <= _MAX_FORME]
            _elargies = []
            for _z in keep:
                _out = fitz.Rect(_z)
                for _r in _formes:
                    _i = _r & _z
                    if _i.is_empty or _r.get_area() <= 0:
                        continue
                    if _i.get_area() / _r.get_area() < _PART_MINI:
                        continue
                    if any(_r.contains(_w) and not fitz.Rect(_z).contains(_w)
                           for _w, _t in _mots_page):
                        continue
                    _out |= _r
                _elargies.append(_out & page.rect)
            keep = _elargies
        except Exception:
            pass          # une page illisible ne doit pas faire tomber l'extraction

        # ══ v2.16 — LES RANGÉES DE LECTURE S'ANCRENT SUR LES ZONES, PLUS SUR UNE GRILLE FIXE ═══
        # La rangée s'ancre sur la PREMIÈRE zone rencontrée : les zones sont triées par hauteur ; tant qu'une zone commence à
        # moins de 24 pt du début de la rangée en cours, elle est de cette rangée ; chaque rangée est lue de gauche à droite.
        # (voir JOURNAL BACKEND v2.16)
        _TOL_RANGEE = 24
        _tri = sorted(keep, key=lambda r: (r.y0, r.x0))
        _rangees, _cour, _debut = [], [], None
        for _r in _tri:
            if _debut is None or _r.y0 - _debut <= _TOL_RANGEE:
                if _debut is None:
                    _debut = _r.y0
                _cour.append(_r)
            else:
                _rangees.append(_cour)
                _cour, _debut = [_r], _r.y0
        if _cour:
            _rangees.append(_cour)
        keep = []
        for _rg in _rangees:
            _rg.sort(key=lambda r: r.x0)
            keep.extend(_rg)

        # ══ LA ZONE EST PHOTOGRAPHIÉE, PAS EXTRAITE ══
        # CONDITION : une zone a été retenue comme figure. EFFET : elle est RENDUE — photographiée telle qu'elle se voit
        # sur la page (traits, texte et images ensemble), et non extraite du PDF comme objet image ; la même image sert au
        # modèle (miniature) et à la feuille. Matrix(2,2) sur 72 ppp = 144 ppp. (voir JOURNAL BACKEND v2.38)
        page_words = page.get_text("words")
        _v247_fenetres = []   # v2.47 — les fenêtres retenues de la page, découpées après les portes
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
                # V2.22 — LE DÉCOR SE SIGNALE, NE SE SUPPRIME PAS : petite figure ancrée aux marges (badge de niveau,
                # cartouche d'éditeur, pastille « Fiche N », ruban latéral), hauteur ≥ 8 mm exigée. (voir JOURNAL BACKEND v2.22)
                # V2.26 — les micros qui TOUCHENT ce cadre (≤ 3 pt) l'étendent.
                for _mi26 in micros_page:
                    if fitz.Rect(clip.x0 - 3, clip.y0 - 3,
                                 clip.x1 + 3, clip.y1 + 3).intersects(_mi26):
                        clip.include_rect(_mi26)
                _w_mm_ = clip.width * 25.4 / 72
                _h_mm_ = clip.height * 25.4 / 72
                _marge_ = (clip.y0 < 0.08 * ph) or (clip.x0 < 0.06 * pw) or (clip.x1 > 0.94 * pw)
                _decor_ = _marge_ and ((_w_mm_ <= 45 and 8 <= _h_mm_ <= 25)
                                       or (_w_mm_ <= 20 and _h_mm_ <= 80 and clip.x1 > 0.94 * pw))
                # V2.21 — un réceptacle de réponse ne part pas en image (B7) ;
                # son contenu est déjà dans le texte, le frontend pose la case.
                _t_clip = page.get_text(clip=clip)
                # v2.32 — UNE GRILLE DÉCLARÉE N'EST PAS UN RÉCEPTACLE. Condition : le cadre vient du module des grilles.
                # Effet : la porte des réceptacles ne s'applique pas — la nature d'une grille est jugée par le module, pas
                # deux fois. (voir JOURNAL BACKEND v2.32)
                _ne_grille = (round(r.x0, 1), round(r.y0, 1),
                              round(r.x1, 1), round(r.y1, 1)) in cadres_grille
                # v2.33 — symétrique de la 2.32 : une case à remplir est une zone VIDE, jamais une image ; une scène qui
                # tient un nombre n'est pas une case. (voir JOURNAL BACKEND v2.33)
                _imgr33 = any(not (ra & clip).is_empty
                              and (ra & clip).get_area() >= 0.5 * min(ra.get_area(), clip.get_area())
                              for ra in rasters)
                if not _ne_grille and not _imgr33 and _v21_est_receptacle(_t_clip):
                    n_receptacles += 1
                    continue
                import re as _re21
                _mots_clip = len(_re21.findall(r"[A-Za-zàâçéèêëîïôöûüù]{2,}", _t_clip))
                # v2.34 — C11 : un cadre sans dessin n'est pas transmis (voir _v234_porte_un_dessin) ; une grille déclarée
                # n'est jamais une fausse figure : sa nature est jugée par le module des grilles, pas deux fois.
                if not _ne_grille and not _v234_porte_un_dessin(page, clip, rasters):
                    n_c11 += 1
                    continue
                # v2.47 — la fenêtre a passé les portes : elle est découpée après (voir _v247_*)
                _v247_fenetres.append({"clip": fitz.Rect(clip), "mots_clip": _mots_clip,
                                       "decor": bool(_decor_), "tableau": bool(_ne_grille),
                                       "thin": bool(is_thin)})
            except Exception:
                pass

        # v2.52 — E1 : les fenêtres d'encre des images en ligne entrent ici, quand les portes ont rendu les fenêtres de la page
        # (C11 comprise) : couverte à 50 % ou plus par elles, une fenêtre d'encre est le même dessin et ne double rien ;
        # sinon elle est une figure de plus, découpée comme les autres — une image (elle rejoint les rasters), jamais regroupée.
        for _r252 in _v252_fenetres_encre(page):
            if _v252_couverte(_r252, [g["clip"] for g in _v247_fenetres]) < 0.5:
                import re as _re252
                rasters.append(_r252)
                _v247_fenetres.append({"clip": fitz.Rect(_r252), "tableau": False, "decor": False, "thin": False,
                                       "mots_clip": len(_re252.findall(r"[A-Za-zàâçéèêëîïôöûüù]{2,}", page.get_text(clip=_r252)))})
        # ══ v2.47 — LA DÉCOUPE EN MORCEAUX (bloc _v247_* ci-dessus) ══
        # Condition : la page a des fenêtres retenues. Effet : chaque fenêtre devient une FIGURE faite de MORCEAUX, chacun
        # une image de la liste avec sa place et sa marque ; les tableaux de la page deviennent des figures ; les mots à
        # 6 mm partent avec leur figure. En cas d'accroc, la page retombe sur la découpe 2.46 (une image par fenêtre),
        # comptée `pages_en_repli` — rien n'est perdu.
        _v247_appendus_page = 0
        try:
            _v247_mots, _v247_L, _v247_texte = _v247_lignes(page)
            _v247_cases, _v247_decor = _v247_cases_et_decor(page)
            _v247_tabs, _v247_nc = _v247_tableaux(page, _v247_cases)
            _v247_compte["tableaux_non_cherches"] += _v247_nc
            _v247_tabs = _v247_fusionner_tableaux(_v247_tabs)   # les cellules d'un quadrillage sont un seul tableau
            _v247_compte["tableaux"] += len(_v247_tabs)
            for _t in _v247_tabs:
                _t = _t & page.rect
                if _t.width < 8 or _t.height < 8 or (_t.width * _t.height) > 0.85 * page_area:
                    continue
                # un tableau qui ne porte ni image ni chiffre ni étiquette est du texte : pas une figure
                if not (any(not (_t & _ra).is_empty for _ra in rasters) or _v247_porte_un_glyphe(page, _t)):
                    _v247_compte["tableaux_de_texte"] = _v247_compte.get("tableaux_de_texte", 0) + 1
                    continue
                _couv = [g for g in _v247_fenetres
                         if (not (_t & g["clip"]).is_empty) and (_t & g["clip"]).get_area() >= 0.6 * _t.get_area()]
                if _couv:
                    for g in _couv:
                        g["tableau"] = True
                else:
                    _v247_fenetres.append({"clip": _t, "mots_clip": 0, "decor": False,
                                           "tableau": True, "thin": False})
                    _v247_compte["tableaux_ajoutes"] += 1
            _v247_res, _v247_n_mots = _v247_decouper_page(
                page, _v247_fenetres, _v247_tabs, _v247_mots, _v247_L, _v247_texte,
                _v247_cases, _v247_decor)
            _v247_compte["mots_attaches"] += _v247_n_mots
            # v2.69 — F13 : la page où la découpe a tourné reçoit ses traits servis (voir le bloc v2.69)
            _v260_cache(page, "a196_pf")[page.number] = True
            # ce qui est à l'intérieur d'un tableau lui appartient : chaque morceau d'une fenêtre
            # ordinaire dont le centre est dans une fenêtre-tableau devient un morceau du tableau
            _v247_tabs_idx = [_j for _j, _t in enumerate(_v247_fenetres) if _t["tableau"]]
            _v247_pieces_par_fig = {}
            _v247_verses = set()
            for _k, g in enumerate(_v247_fenetres):
                for _p in _v247_res[_k]["pieces"]:
                    _cible = _k
                    if not g["tableau"]:
                        _c = fitz.Point((_p.x0 + _p.x1) / 2, (_p.y0 + _p.y1) / 2)
                        for _j in _v247_tabs_idx:
                            if _v247_fenetres[_j]["clip"].contains(_c):
                                _cible = _j
                                _v247_verses.add(_k)
                                break
                    _v247_pieces_par_fig.setdefault(_cible, []).append(_p)
            _v247_parent = {_k: 0 for _k in _v247_verses if not _v247_pieces_par_fig.get(_k)}
            _v247_compte["fenetres_versees_au_tableau"] = _v247_compte.get("fenetres_versees_au_tableau", 0) + len(_v247_parent)
            _v247_couverts = set()
            for _k, g in enumerate(_v247_fenetres):
                if not g["thin"] or _k in _v247_parent:
                    continue
                for _j, g2 in enumerate(_v247_fenetres):
                    if _j == _k or g2["thin"] or _j in _v247_parent:
                        continue
                    if any(((p & g["clip"]).get_area() >= 0.8 * g["clip"].get_area())
                           for p in _v247_pieces_par_fig.get(_j, []) if not (p & g["clip"]).is_empty):
                        _v247_couverts.add(_k)
                        break
            _v247_compte["traits_absorbes"] += len(_v247_couverts)
            for _k, g in enumerate(_v247_fenetres):
                if _k in _v247_parent or _k in _v247_couverts:
                    continue
                # v2.54 — l'absorption dessin par dessin
                _pieces = _v254_absorber_par_dessin(page, [p & page.rect for p in _v247_pieces_par_fig.get(_k, [])
                                                           if not (p & page.rect).is_empty])
                if not _pieces:
                    _v247_compte["fenetres_sans_dessin"] += 1
                    continue
                _pieces.sort(key=lambda r: (round(r.y0), r.x0))
                _fig_id = "p%d-f%d" % (pno + 1, _k)
                _entier = fitz.Rect(min(p.x0 for p in _pieces), min(p.y0 for p in _pieces),
                                    max(p.x1 for p in _pieces), max(p.y1 for p in _pieces))
                _entrees = []
                for _p in _pieces:
                    # v2.54 — un morceau à sa taille exacte
                    _p = _p & page.rect
                    pix = page.get_pixmap(clip=_p, matrix=fitz.Matrix(2, 2))
                    if pix.width < 1 or pix.height < 1:
                        _v247_compte["morceaux_trop_petits"] += 1
                        continue
                    _png30 = pix.tobytes("png")
                    _v238_extra, _v238_voie = _v238_champs(page, _p)
                    _entree = {
                        "cadre": [_p.x0, _p.y0, _p.x1, _p.y1],
                        "sceau": _hl.sha256(_png30).hexdigest()[:12],
                        **({"mots_taches": g["mots_clip"]} if g["mots_clip"] >= 6 else {}),
                        **({"decor_probable": True} if g["decor"] else {}),
                        **({"tableau": True} if g["tableau"] else {}),
                        "index": idx,
                        "page": pno + 1,
                        "data": "data:image/png;base64," + base64.b64encode(_png30).decode(),
                        "w": pix.width, "h": pix.height,
                        "w_mm": round(_p.width * 25.4 / 72, 1),
                        "h_mm": round(_p.height * 25.4 / 72, 1),
                        "figure": _fig_id,
                        "morceau": len(_entrees) + 1,
                        "n_morceaux": 0,
                        **_v238_extra,
                    }
                    images.append(_entree)
                    _entrees.append(_entree)
                    _v238_compte[_v238_voie] = _v238_compte.get(_v238_voie, 0) + 1
                    idx += 1
                if not _entrees:
                    _v247_compte["fenetres_sans_dessin"] += 1
                    continue
                for _e in _entrees:
                    _e["n_morceaux"] = len(_entrees)
                if len(_entrees) > 1:
                    _pe = page.get_pixmap(clip=_entier & page.rect, matrix=fitz.Matrix(2, 2))
                    _entrees[0]["entier"] = "data:image/png;base64," + base64.b64encode(_pe.tobytes("png")).decode()
                    _entrees[0]["entier_cadre"] = [_entier.x0, _entier.y0, _entier.x1, _entier.y1]
                _v247_compte["figures"] += 1
                _v247_compte["morceaux"] += len(_entrees)
                _v247_appendus_page += len(_entrees)
        except Exception as _e247:
            _v247_compte["pages_en_repli"] += 1
            _v247_compte["derniere_erreur"] = (str(_e247) or type(_e247).__name__)[:160]
            for g in _v247_fenetres:
                if g["tableau"] and g.get("mots_clip") == 0 and _v247_appendus_page == 0 and False:
                    continue
                try:
                    clip = g["clip"]
                    pix = page.get_pixmap(clip=clip, matrix=fitz.Matrix(2, 2))
                    if pix.width < 8 or pix.height < 8:
                        continue
                    _png30 = pix.tobytes("png")
                    _v238_extra, _v238_voie = _v238_champs(page, clip)
                    images.append({
                        "cadre": [clip.x0, clip.y0, clip.x1, clip.y1],
                        "sceau": _hl.sha256(_png30).hexdigest()[:12],
                        **({"mots_taches": g["mots_clip"]} if g["mots_clip"] >= 6 else {}),
                        **({"decor_probable": True} if g["decor"] else {}),
                        "index": idx, "page": pno + 1,
                        "data": "data:image/png;base64," + base64.b64encode(_png30).decode(),
                        "w": pix.width, "h": pix.height,
                        "w_mm": round(clip.width * 25.4 / 72, 1),
                        "h_mm": round(clip.height * 25.4 / 72, 1),
                        "figure": "p%d-repli%d" % (pno + 1, idx), "morceau": 1, "n_morceaux": 1,
                        **_v238_extra,
                    })
                    _v238_compte[_v238_voie] = _v238_compte.get(_v238_voie, 0) + 1
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

        # v2.54 — après la découpe de la page, avant les ancres : une figure par dessin ; les morceaux d'un tableau qui ne se
        # touchent pas séparés (v2.62) ; la lecture des exercices de la page ; une figure par exercice ; les morceaux d'un tableau
        # qui ne se touchent pas séparés ; chaque image ne montre que son dessin.
        images[i0_page:], idx = _v254_par_dessin(page, images[i0_page:], idx)
        images[i0_page:], idx = _v254_separer(page, images[i0_page:], idx)
        _v254_lire_la_page(page, images[i0_page:], api_key, _v254_lecture)
        images[i0_page:], idx = _v254_par_exercice(page, images[i0_page:], idx)
        images[i0_page:], idx = _v254_separer(page, images[i0_page:], idx)
        # v2.69 — F034 S4b : les cases de réponse V/F sortent des images (voir _v269_s4b)
        images[i0_page:], idx, _v269_voies = _v269_s4b(page, images[i0_page:], idx)
        for _v269_v in _v269_voies:
            _v238_compte[_v269_v] = _v238_compte.get(_v269_v, 0) + 1
        images[i0_page:], idx = _v254_montrer_son_dessin(page, images[i0_page:], idx)
        # v2.60 — à la sortie de la page : les restes de pastille, les surlignages, les glyphes orphelins
        images[i0_page:], idx, _v260_voies = _v260_sortie_page(page, images[i0_page:], idx)
        for _v260_v in _v260_voies:
            _v238_compte[_v260_v] = _v238_compte.get(_v260_v, 0) + 1
        # v2.63 — Z9-v1 : les figures de la page dans l'ordre de lecture (voir JOURNAL BACKEND v2.63)
        images[i0_page:], _v263_voie = _v263_ordre_de_lecture(page, images[i0_page:])
        if _v263_voie:
            _v238_compte[_v263_voie] = _v238_compte.get(_v263_voie, 0) + 1

        # ══ v2.39 — LA DISTANCE SOUS UNE FIGURE EST CELLE DE LA SOURCE ══
        # L'espace entre le bas de la figure et la première chose qui vient en dessous (texte, autre figure, tableau) est
        # reporté tel quel ; rien en dessous, ou distance inconnue : aucune marge n'est ajoutée. Les boîtes complètes sont
        # relues ici : les blocs rendus par `_texte_gras` n'ont pas de bord droit, et « en dessous » demande la largeur.
        # (voir JOURNAL BACKEND v2.39)
        _v239_distance_sous(page, images[i0_page:])

        # v2.60 — les figures de CETTE page reçoivent leur marqueur : une figure d'exercice après la consigne de son exercice
        # trouvée dans le texte de la page, toute autre au coin de son image entière (voir JOURNAL BACKEND v2.60)
        # v2.64 — S3 : le marqueur d'une figure d'exercice juste après la dernière ligne de sa consigne (voir JOURNAL BACKEND v2.64)
        _v264_anc, _v264_lieux, _v264_L, _v264_exs = _v260_ancres(page, images[i0_page:], avec_contexte=True)
        texte_page, _n_pose, _n_hors = poser_ancres(
            texte_page, blocs_page, _v264_anc,
            getattr(_grilles, "cle_lecture", None) if _grilles else None,
            _v264_apres_consigne(texte_page, blocs_page, _v264_anc, _v264_lieux, _v264_L, _v264_exs, page.number + 1))
        n_ancres += _n_pose
        n_ancres_hors += _n_hors

        # v2.69 — E1 (b) : le ⊥ dessiné d'une ligne écrit « ⊥ » dans le texte de la page, quelle que soit la route (voir _v269_perp_dans_le_texte)
        texte_page, _ = _v269_perp_dans_le_texte(page, texte_page)
        full_text.append(texte_page)

        # ══ v2.45 — S1 : LES SUPPORTS NON TRANSMIS DEVIENNENT DES FIGURES ══
        # (voir l'en-tête ; le texte de la page est DÉJÀ posé : rien n'y bouge)
        if GRILLES_ACTIVES and grilles_page:
            try:
                _s245_candidats = []
                for _g45 in grilles_page:
                    _r45 = fitz.Rect(_g45["rect"]) & page.rect
                    if _r45.is_empty:
                        continue
                    # v2.54 — une grille est déjà transmise quand la moitié au moins de son encre est sous les images de la page
                    _couv45 = _v254_encre_portee(page, _r45, [f["cadre"] for f in images if f.get("page") == pno + 1]) >= 0.5
                    if _couv45:
                        continue
                    # v2.46 — la frontière prend la place du compteur de mots : une grille qui porte un dessin ou une
                    # couleur part en figure, telle quelle ; toute autre grille reste au texte.
                    if not _v246_porte_dessin_ou_couleur(page, _r45):
                        continue
                    _s245_candidats.append((_r45.y0, _r45.x0, _g45, _r45))
                for _y45, _x45, _g45, _r45 in sorted(_s245_candidats,
                                                     key=lambda c: (c[0], c[1])):
                    clip45 = fitz.Rect(_r45)
                    # le cadre suit le DESSIN : bord droit étendu aux mots chevauchants
                    _cl45 = (_g45.get("case_l_mm") or 8.0) * 72 / 25.4
                    _mx45 = clip45.x1
                    for _w45 in page.get_text("words"):
                        _wr45 = fitz.Rect(_w45[:4])
                        _cy45 = (_wr45.y0 + _wr45.y1) / 2
                        if (clip45.y0 <= _cy45 <= clip45.y1
                                and _wr45.x1 > clip45.x1 - 0.2 * _cl45
                                and _wr45.x0 < clip45.x1 + 1.2 * _cl45):
                            _mx45 = max(_mx45, _wr45.x1 + 0.34 * _cl45)
                    clip45.x1 = _mx45
                    clip45 = fitz.Rect(clip45.x0 - 1.5, clip45.y0 - 1.5,
                                       clip45.x1 + 1.5, clip45.y1 + 1.5) & page.rect
                    pix45 = page.get_pixmap(clip=clip45, matrix=fitz.Matrix(2, 2))
                    if pix45.width < 8 or pix45.height < 8:
                        continue
                    _png45 = pix45.tobytes("png")
                    _t45c = page.get_text("text", clip=clip45)
                    _mots45 = len(re.findall(r"[A-Za-zàâçéèêëîïôöûüù]{2,}", _t45c))
                    _extra45, _voie45 = _v238_champs(page, clip45)
                    # v2.54 — une grille qui contient une figure agrandit son plus grand morceau à son cadre, au lieu de s'ajouter
                    _b45 = fitz.Rect(clip45.x0 - 0.5, clip45.y0 - 0.5, clip45.x1 + 0.5, clip45.y1 + 0.5)
                    _in45 = [f for f in images if f.get("page") == pno + 1 and f.get("figure") and _b45.contains(fitz.Rect(f["cadre"]))]
                    if _in45:
                        _f45 = max(_in45, key=lambda f: _v254_aire(f["cadre"]))
                        _f45.pop("data_feuille", None)
                        _f45.update({"cadre": [clip45.x0, clip45.y0, clip45.x1, clip45.y1], "sceau": _hl.sha256(_png45).hexdigest()[:12],
                                     "data": "data:image/png;base64," + base64.b64encode(_png45).decode(), "w": pix45.width, "h": pix45.height,
                                     "w_mm": round(clip45.width * 25.4 / 72, 1), "h_mm": round(clip45.height * 25.4 / 72, 1), "s1_etendu": True})
                        _f45.update(_extra45)
                        continue
                    _fig45 = {"cadre": [clip45.x0, clip45.y0, clip45.x1, clip45.y1],
                              "sceau": _hl.sha256(_png45).hexdigest()[:12]}
                    if _mots45 >= 6:
                        _fig45["mots_taches"] = _mots45
                    _fig45.update({"index": idx, "page": pno + 1,
                                   "data": "data:image/png;base64,"
                                           + base64.b64encode(_png45).decode(),
                                   "w": pix45.width, "h": pix45.height,
                                   "w_mm": round(clip45.width * 25.4 / 72, 1),
                                   "h_mm": round(clip45.height * 25.4 / 72, 1)})
                    _fig45.update(_extra45)
                    images.append(_fig45)
                    _v238_compte[_voie45] = _v238_compte.get(_voie45, 0) + 1
                    _s245_ajouts[idx] = {"page": pno + 1,
                                         "rect": list(fitz.Rect(_g45["rect"]))}
                    idx += 1
            except Exception:
                pass
        # v2.61 — la pastille prise dans une grande image, effacée après les supports ajoutés ci-dessus
        _v261_pastilles_grandes(page, images[i0_page:])
        # v2.69 — E1 (b) : le ⊥ écrit en texte (lui seul) est effacé des images qui le contiennent, par l'effaceur des pastilles
        for _v269_bb in _v269_perps_ecrits(page):
            _v269_cx, _v269_cy = 0.5 * (_v269_bb.x0 + _v269_bb.x1), 0.5 * (_v269_bb.y0 + _v269_bb.y1)
            for _v269_im in images[i0_page:]:
                _v269_c = _v269_im["cadre"]
                if _v269_c[0] <= _v269_cx <= _v269_c[2] and _v269_c[1] <= _v269_cy <= _v269_c[3]:
                    _v260_efface_pastille(_v269_im, _v269_bb)
        if GRILLES_ACTIVES and grilles_page:
            for _g45 in grilles_page:
                _s245_grilles_pos.append({
                    "page": pno + 1, "rect": list(fitz.Rect(_g45["rect"])),
                    "nature": _g45.get("nature"),
                    "case_l_mm": _g45.get("case_l_mm"),
                    "case_h_mm": _g45.get("case_h_mm")})

    # ══ v2.45 — S2 : CHAQUE SUPPORT REÇOIT SON TEXTE SOUS LA MARQUE ══════
    # Le champ `text` ne bouge pas ; seuls les BLOCS sortent, dans l'ordre de lecture. En cas d'accroc, le champ dit
    # pourquoi au lieu de se taire. (voir JOURNAL BACKEND v2.45)
    _s245_marques = ""
    try:
        _S245_TOK = re.compile(r"\[TAILORY_IMG_(\d+)\]")
        _S245_DECL = re.compile(r"^\[(quadrillage|tableau) .*\]")
        _S245_COMPT = re.compile(r"^\[comptage non transmis[^\]]*\]")
        _S245_CELL = re.compile(r"^[^|]*\|[^|]*\|")
        _s245_texte = "\n".join(full_text)
        _s245_transmises = [f for f in images if f["index"] not in _s245_ajouts]
        _s245_pages_contenu = sorted({f["page"] for f in images}
                                     | {g["page"] for g in _s245_grilles_pos})

        def _s245_centre_dans(bb, cadre, marge=1.0):
            cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
            return (cadre[0] - marge <= cx <= cadre[2] + marge
                    and cadre[1] - marge <= cy <= cadre[3] + marge)

        def _s245_couvert(rect, cadre, seuil=0.8):
            r = fitz.Rect(rect) & fitz.Rect(cadre)
            if r.is_empty:
                return False
            a = fitz.Rect(rect).get_area()
            return a > 0 and r.get_area() >= seuil * a

        _s245_recolte = {}
        for _p45 in _s245_pages_contenu:
            if _p45 in corrige or _p45 > len(doc):
                continue
            _page45 = doc[_p45 - 1]
            _gr45 = [fitz.Rect(g["rect"]) for g in _s245_grilles_pos
                     if g["page"] == _p45]
            _d45 = _page45.get_text("dict")
            for _b45 in _d45["blocks"]:
                if _b45.get("type") != 0:
                    continue
                for _li45 in _b45["lines"]:
                    _t45 = " ".join(s["text"].strip() for s in _li45["spans"]
                                    if s["text"].strip())
                    _t45 = re.sub(r"\s+", " ", _t45).strip()
                    if not _t45:
                        continue
                    _bb45 = _li45["bbox"]
                    if any(_s245_centre_dans(_bb45, [r.x0, r.y0, r.x1, r.y1])
                           for r in _gr45):
                        continue
                    for _f45 in _s245_transmises:
                        if (_f45["page"] == _p45
                                and _s245_centre_dans(_bb45, _f45["cadre"])):
                            _s245_recolte.setdefault(_p45, []).append(
                                [_t45, _f45["index"], 1])
                            break

        def _s245_sans_espace(t):
            return t.replace(" ", "")

        def _s245_apparie(nu, pno1):
            pool = _s245_recolte.get(pno1, [])
            nu_c = _s245_sans_espace(nu)
            for e in pool:
                if e[2] and _s245_sans_espace(e[0]) == nu_c:
                    e[2] = 0
                    return e[1]
            for sup in {e[1] for e in pool if e[2]}:
                libres = [e for e in pool if e[2] and e[1] == sup]
                reste = nu
                pris = []
                change = True
                while reste and change:
                    change = False
                    for e in libres:
                        ec = _s245_sans_espace(e[0])
                        if e[2] and _s245_sans_espace(reste).startswith(ec):
                            pris.append(e)
                            rc = _s245_sans_espace(reste)[len(ec):]
                            j = 0
                            k = 0
                            while (j < len(reste)
                                   and k < len(_s245_sans_espace(reste)) - len(rc)):
                                if reste[j] != " ":
                                    k += 1
                                j += 1
                            reste = reste[j:].strip()
                            e[2] = 0
                            change = True
                            break
                if not reste and pris:
                    return sup
                for e in pris:
                    e[2] = 1
            return None

        _s245_file = sorted(_s245_grilles_pos,
                            key=lambda g: (g["page"], g["rect"][1], g["rect"][0]))

        def _s245_grille_suivante(pno1):
            for g in _s245_file:
                if g["page"] == pno1:
                    _s245_file.remove(g)
                    return g
            return None

        def _s245_support_de(g):
            if g is None:
                return None, None
            for _ia45, _ga45 in _s245_ajouts.items():
                if (_ga45.get("page") == g["page"]
                        and list(fitz.Rect(_ga45["rect"])) == g["rect"]):
                    return "ajout", _ia45
            for f in _s245_transmises:
                if f["page"] == g["page"] and _s245_couvert(g["rect"], f["cadre"]):
                    return "couvert", f["index"]
            return "libre", None

        _s245_page_fig = {f["index"]: f["page"] for f in _s245_transmises}
        _s245_sortie = []
        _s245_bloc_tab = None
        _s245_page_cour = (_s245_pages_contenu[0]
                           if _s245_pages_contenu else 1)
        _s245_ouverte = [None]
        _s245_derniere_decl = [None]

        def _s245_ferme():
            if _s245_ouverte[0] is not None:
                _s245_sortie.append("⟦/support⟧")
                _s245_ouverte[0] = None

        def _s245_ouvre(n):
            if _s245_ouverte[0] != n:
                _s245_ferme()
                _s245_sortie.append("⟦support %d — à lire seulement⟧" % n)
                _s245_ouverte[0] = n

        for _brut45 in _s245_texte.split("\n"):
            _toks45 = _S245_TOK.findall(_brut45)
            if _toks45:
                _pg45 = max(_s245_page_fig.get(int(t), _s245_page_cour)
                            for t in _toks45)
                if _pg45 > _s245_page_cour:
                    _s245_page_cour = _pg45
            _nu45 = _S245_TOK.sub("", _brut45).strip()
            _nu45 = re.sub(r"\s+", " ", _nu45)
            _sort45, _n45 = "garde", None
            if _S245_DECL.match(_nu45):
                _g45x = _s245_grille_suivante(_s245_page_cour)
                _typ45, _n45 = _s245_support_de(_g45x)
                _sort45 = "support" if _typ45 in ("ajout", "couvert") else "garde"
                _s245_bloc_tab = ((_sort45, _n45)
                                  if _g45x and _g45x.get("nature") == "tableau"
                                  else None)
                _s245_derniere_decl[0] = (_typ45, _n45)
            elif _S245_COMPT.match(_nu45):
                if (_s245_derniere_decl[0]
                        and _s245_derniere_decl[0][0] in ("ajout", "couvert")):
                    _sort45, _n45 = "support", _s245_derniere_decl[0][1]
                else:
                    _sort45 = "garde"
            elif _S245_CELL.match(_nu45) and _s245_bloc_tab:
                _sort45, _n45 = _s245_bloc_tab
            else:
                if _nu45:
                    _s245_bloc_tab = None
                    _sup45 = _s245_apparie(_nu45, _s245_page_cour)
                    if _sup45 is not None:
                        _sort45, _n45 = "support", _sup45
            if _sort45 == "garde" or _n45 is None:
                _s245_ferme()
            else:
                _s245_ouvre(_n45)
                _s245_sortie.append(_nu45)
        _s245_ferme()
        _s245_b = "\n".join(_s245_sortie)
        _s245_marques = "\n".join(re.findall(
            r"(⟦support \d+ — à lire seulement⟧.*?⟦/support⟧)", _s245_b, re.S))
    except Exception as _e45:
        _s245_marques = "[supports_marques indisponibles : %s]" % _e45

    # v2.60 — l'image entière d'une figure est faite de ses morceaux posés ; les pages scannées sont déclarées
    # v2.62 — l'image entière de feuille ; les lignes du texte des pages jugées corrigé
    # v2.69 — F13 : les traits pointillés fins jugés décor sont servis comme des images, à la fin de la liste
    images = _a196_ajouter_traits(doc, images)
    _v260_entier_compose(images)
    _v262_entier_feuille(images)
    # v2.69 — B-N2 + Wc2 : les endroits du signe (voir _v269_bn2) ; une panne se déclare, /parse continue
    try:
        _v269_endroits = _v269_bn2(images, doc)
        _l59_calculs = _v269_endroits.pop("calculs_er", None)
        _v269_bn2_err = None
    except Exception as _e269b:
        _v269_endroits, _v269_bn2_err = {"gestes": [], "comptes": {}}, "%s: %s" % (type(_e269b).__name__, _e269b)
        _l59_calculs = None
    # ligne 5.3 (D2-b) : le quadrillage à points de la source est gardé sur l'image entière de feuille ; une panne se déclare, /parse continue
    try:
        _l53_d2b(images, doc, _v269_endroits)
    except Exception as _e53:
        _v269_bn2_err = ((_v269_bn2_err + " ; ") if _v269_bn2_err else "") + "D2-b — %s: %s" % (type(_e53).__name__, _e53)
    _v260_scannees = [p.number + 1 for p in doc if _v260_est_scannee(p)]
    _v262_corrige_lignes = {str(p): [l for l in doc[p - 1].get_text().split("\n") if l.strip()] for p in sorted(corrige)}
    # v2.66 — étape B : les trois champs que la page lira (voir le bloc v2.66) ; une panne se déclare, /parse continue
    # v2.67 — la part cursive de voisins_b (voir le bloc v2.67) : les mots cursifs mesurés une fois, E49 réunies aux lignes,
    # puis les mots de D6 avec leur cache ; une panne se déclare dans voisins_b_erreur, les lignes restent
    try:
        _v267_cur = _v267_mots_cursifs_doc(doc, images, corrige)
        _v267_err = None
    except Exception as _e267:
        _v267_cur, _v267_err = None, "cursif — %s: %s" % (type(_e267).__name__, _e267)
    try:
        # v2.69 — Zk : le champ des mots voit aussi les images que S4b a touchées (voir _v269_zk_fantomes)
        _v266_vb = _vb_part_b(doc, images + _v269_zk_fantomes(doc), corrige, _v254_etat()["lecture"].get("vb_reps") or {},
                              _v267_cur["e49"] if _v267_cur else ())
        _v266_vb_err = None
    except Exception as _e266:
        _v266_vb, _v266_vb_err = {}, "%s: %s" % (type(_e266).__name__, _e266)
    if _v267_cur is not None:
        try:
            _v267_ajoute_d6(doc, images, _v267_cur, _v266_vb)
        except Exception as _e267d:
            _v267_err = "cursif D6 — %s: %s" % (type(_e267d).__name__, _e267d)
    if _v267_err:
        _v266_vb_err = (_v266_vb_err + " ; " if _v266_vb_err else "") + _v267_err
    # v2.69 — F068 RL2 : les lignes de rangée (voir _v269_rl2), avant la couleur des mots ; une entrée identique déjà là ou une panne se
    # déclare dans voisins_b_erreur, et aucune entrée n'est ajoutée
    try:
        _v269_rl2_ch = _v269_rl2(doc, images, corrige)
        for _v269_f, _v269_E in _v269_rl2_ch.items():
            for _v269_q in _v269_E:
                if any(v.get("texte") == _v269_q["texte"] and v.get("ligne") == _v269_q["ligne"]
                       and all(abs(a - b) <= 0.5 for a, b in zip(v["bbox"], _v269_q["bbox"])) for v in _v266_vb.get(_v269_f, [])):
                    raise RuntimeError("une entrée identique existe déjà dans voisins_b[%s] (%r)" % (_v269_f, _v269_q["texte"]))
        for _v269_f, _v269_E in _v269_rl2_ch.items():
            _v266_vb.setdefault(_v269_f, []).extend(_v269_E)
    except Exception as _e269r:
        _v266_vb_err = (_v266_vb_err + " ; " if _v266_vb_err else "") + "RL2 — %s: %s" % (type(_e269r).__name__, _e269r)
    try:
        _v266_couleurs_voisins(doc, _v266_vb)
    except Exception as _e266c:
        _v266_vb_err = (_v266_vb_err + " ; " if _v266_vb_err else "") + "couleur — %s: %s" % (type(_e266c).__name__, _e266c)
    # v2.69 — G2 : l'effacement sous le rectangle des mots cursifs (voir _v269_g2) ; une panne se déclare, /parse continue
    try:
        _v269_effacements = _v269_g2(images, _v266_vb, doc)
        _v269_g2_err = None
    except Exception as _e269g:
        _v269_effacements, _v269_g2_err = [], "%s: %s" % (type(_e269g).__name__, _e269g)
    try:
        _v266_lc = _v266_lignes_couleur(doc, set(corrige))
        _v266_lc_err = None
    except Exception as _e266l:
        _v266_lc, _v266_lc_err = [], "%s: %s" % (type(_e266l).__name__, _e266l)
    # BLOC 5, LIGNE 5.1 — S4d-c4 : le champ `couleur_ou_dessin` posé sur chaque image (voir _s4d_declare) ; une panne se déclare, /parse continue
    try:
        _s4d_declare(doc, images)
        _s4d_err = None
    except Exception as _es4d:
        _s4d_err = "%s: %s" % (type(_es4d).__name__, _es4d)
    _v254_etat_neuf()
    doc.close()
    return {
        "filename": filename,
        "pdf_mode": True,
        # v2.52 — les pages jugées corrigé, déclarées (jamais en silence)
        "corrige_pages": sorted(corrige),
        "corrige_lecture": _v252_compte,   # v2.52 — la marque machine du lecteur
        "corrige_lignes": _v262_corrige_lignes,   # v2.62 — les lignes du texte de chaque page jugée corrigé
        "exercices_lecture": _v254_lecture,   # v2.54 — la marque machine de la lecture des exercices
        # V2.21 — zones de production déclarées + réceptacles écartés, comptés
        "zones_libres": zones_libres,
        "receptacles_ecartes": n_receptacles,
        "cadres_sans_dessin": n_c11,          # v2.34 — C11
        # v2.36 — R1 : ce qui a été ancré, et ce qui ne l'a pas été. L'absence
        # se compte, elle ne se tait pas (A6/A21/A22).
        "ancres_posees": n_ancres,
        "ancres_absentes": n_ancres_hors,

        # v2.38 — LA DIVERGENCE, CHIFFRÉE ET RENDUE (voir `_v238_champs`).
        # Ces deux nombres disent, pour CE document, combien de figures la
        # feuille recevra en image d'origine et combien en photographie à
        # 300 ppp. `figures_aucune` compte les figures pour lesquelles NI l'une
        # NI l'autre n'a pu être fabriquée : elles retombent sur `data`, et
        # elles ne se taisent pas.
        "figures_image_origine": _v238_compte.get("origine", 0),
        "figures_photographie": _v238_compte.get("feuille", 0),
        "figures_aucune": _v238_compte.get("aucun", 0),
        # v2.63 — Z9-v1 : les pages renumérotées dans l'ordre de lecture, et celles où la règle s'est abstenue.
        "z9v1_pages_renumerotees": _v238_compte.get("z9v1_renumerote", 0),
        "z9v1_pages_abstention": _v238_compte.get("z9v1_abstention", 0),

        "num_exercises": 0,
        "num_images": len(images),
        "images": images,
        "decoupe": _v247_compte,   # v2.47 — le recompte des morceaux, rendu
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
        # v2.45 — S2 : les passages des supports sous la marque, dans l'ordre de lecture
        "supports_marques": _s245_marques,
        # v2.60 — les pages qu'une image couvre à 85 % et qui portent moins de 20 mots
        "pages_scannees": _v260_scannees,
        # v2.66 — étape B : voisins_b (part B, F015 ; la couleur des mots, F083), lignes_couleur (C1e, F088), leurs pannes
        "voisins_b": _v266_vb,
        "voisins_b_erreur": _v266_vb_err,
        "lignes_couleur": _v266_lc,
        "lignes_couleur_erreur": _v266_lc_err,
        # v2.69 — G2 : les effacements faits (rejoués par /parse après sa recomposition), et leur panne
        "effacements_cursifs": _v269_effacements,
        # v2.69 — B-N2 + Wc2 : les gestes des endroits du signe (rejoués par /parse après sa recomposition), et leur panne
        "endroits_signe": _v269_endroits,
        "endroits_signe_erreur": _v269_bn2_err,
        # ligne 5.9 (ER) : la déclaration des calculs effacés des images (la page les réécrit) ; la clé n'existe que s'il y en a une
        **({"calculs_er": _l59_calculs} if _l59_calculs else {}),
        "effacements_cursifs_erreur": _v269_g2_err,
        "couleur_ou_dessin_erreur": _s4d_err,
    }

# 8.6 (DEC-1249 : « L'outil n'accepte que des PDF ») — la rasterisation des images non-web par LibreOffice servait la seule route Word, retirée : plus aucun appelant, bloc retiré par CX1.


# ─────────────────────────────────────────────
# V2.15 — LE SERVEUR ANNONCE LES FORMATS QU'IL SAIT LIRE : la page l'affiche, elle n'a pas de liste à elle.
# Porte du dépôt : ce que /formats annonce à la page. (voir JOURNAL BACKEND v2.15 et v2.42)
FORMATS_PORTE = ("pdf", "txt")


@app.get("/formats")
async def formats_acceptes():
    """Ce que la porte du dépôt accepte. La page s'en sert pour son filtre
    de fichier et son message de refus. (voir JOURNAL v2.42)"""
    return {"formats": list(FORMATS_PORTE)}


# ─────────────────────────────────────────────
# ENDPOINT : /parse — PDF seul (tout autre fichier est refusé) → structure JSON pédagogique
# ─────────────────────────────────────────────
# ═══ v2.43 — LA SYMBIOSE : LE MODÈLE VISE, LE CODE DÉCIDE, LE SERVEUR REDÉCOUPE ═══
# Condition : /parse reçoit un champ facultatif `api_key` ET le document est un PDF. Sans clé, RIEN ne change.
# Effet : (1) le modèle juge L'USAGE des tracés et des textes visibles, par tuiles numérotées, UN jugement par famille
#         de jumeaux ; (2) les décors jugés « ne sert pas » sont caviardés (tracé entier touché ; texte retiré dans sa
#         boîte) ; (3) v2.70 : bornée par le code et par la tâche de l'élève, la gomme agit AVANT la photographie : les images
#         sont refaites d'après le PDF effacé ; il n'y a plus de ramassage de « récupérés ». Un champ additif `symbiose` dit
#         tout (tuiles jouées/non jugées, familles, jetons, le bornage, les images refaites) ; une tuile illisible est NON
#         JUGÉE et DITE, jamais complétée ; une panne de l'étage rend le parse d'aujourd'hui, avec l'erreur DITE.
#         (voir JOURNAL BACKEND v2.43 et v2.70)
# v2.51 — la gomme juge avec Opus 5.5.
_S243_MODELE = "claude-opus-5-5"
_S243_TUILE_MAX = 110
# DEC-2622 (04.10.2026) : la règle de la gomme (« Ce qui doit rester sur la feuille, c'est ce qui sert à l'élève. […]
# Le reste part. »), faite sienne par Catherine ; rejoint DEC-0125 et DEC-1670. (voir JOURNAL BACKEND v2.65)
_S243_DEMANDE = (
    "Voici une bande découpée dans une page de fiche d'école. Des ÉLÉMENTS y "
    "sont entourés en orange et numérotés T1, T2… (des tracés — traits, "
    "cadres, cases, fonds — et des lignes de texte). La règle qui décide, "
    "mot pour mot : « Ce qui doit rester sur la feuille, c'est ce qui sert à "
    "l'élève. Une case où il écrit, une bande qu'il lit, une figure qu'il "
    "regarde. Le reste part. » Pour CHAQUE numéro de la liste, rends un "
    "jugement d'USAGE : \"sert\" (à l'élève) · \"ne sert pas\" (décor, "
    "habillage d'éditeur, cartouche, pastille, cadre décoratif, consigne que "
    "l'outil réécrit ailleurs) · \"ne vois pas\". Réponds UNIQUEMENT par un "
    "bloc JSON : {\"verdicts\":[{\"n\":1,\"v\":\"sert\"}]} — chaque numéro, "
    "EXACTEMENT une fois. La liste : ")


def _s243_famille(bb):
    return (round((bb[2] - bb[0]) * 2) / 2, round((bb[3] - bb[1]) * 2) / 2)


def _s243_items_de_page(page):
    """(tracés uniques {bbox: famille}, lignes de texte visibles [(bbox, txt)])."""
    traces = {}
    for g in page.get_drawings():
        r = g["rect"]
        bb = (round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2))
        traces[bb] = _s243_famille(bb)
    textes = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for l in b.get("lines", []):
            t = "".join(s.get("text", "") for s in l.get("spans", []))
            if t.strip():
                textes.append((tuple(l["bbox"]), t.strip()))
    return traces, textes


def _s243_tuile_png(page, lot):
    """L'image de la tuile, éléments numérotés. lot : [(num, bbox)]."""
    y0 = max(0.0, min(b[1] for _n, b in lot) - 6)
    y1 = min(page.rect.height, max(b[3] for _n, b in lot) + 6)
    clip = fitz.Rect(0, y0, page.rect.width, y1)
    g = min(4.0, 1500.0 / max(clip.width, clip.height))
    pm = page.get_pixmap(clip=clip, matrix=fitz.Matrix(g, g))
    d = fitz.open()
    p2 = d.new_page(width=pm.width, height=pm.height)
    p2.insert_image(fitz.Rect(0, 0, pm.width, pm.height), pixmap=pm)
    for num, bb in lot:
        rr = fitz.Rect((bb[0] - clip.x0) * g - 1, (bb[1] - clip.y0) * g - 1,
                       (bb[2] - clip.x0) * g + 1, (bb[3] - clip.y0) * g + 1)
        p2.draw_rect(rr, color=(1, 0.45, 0), width=1.2)
        p2.insert_text((rr.x0 + 1, max(12, rr.y0 - 3)), "T%d" % num,
                       fontsize=11, color=(0.85, 0.1, 0))
    png = p2.get_pixmap(matrix=fitz.Matrix(1, 1)).tobytes("png")
    d.close()
    return png


def _s243_depouille(texte, numeros):
    """La partition exacte des numéros, ou la faute — jamais un trou."""
    try:
        deb = texte.index("{")
        fin = texte.rindex("}")
        rep = json.loads(texte[deb:fin + 1])
    except (ValueError, json.JSONDecodeError):
        return None, "réponse illisible"
    if not isinstance(rep, dict) or not isinstance(rep.get("verdicts"), list):
        return None, "réponse illisible"
    vus = {}
    for v in rep["verdicts"]:
        try:
            n = int(v["n"])
            verdict = str(v["v"]).strip().lower()
        except (KeyError, TypeError, ValueError):
            return None, "entrée illisible"
        if verdict not in ("sert", "ne sert pas", "ne vois pas") or n in vus:
            return None, "verdict inconnu ou numéro en double"
        vus[n] = verdict
    if any(n not in vus for n in numeros) or any(n not in numeros for n in vus):
        return None, "numéros manquants ou inconnus"
    return vus, None


def _s243_juge_tuile(api_key, png, liste_txt, numeros):
    """UN appel au modèle des tirages pour UNE tuile. Rend (verdicts | None, faute | None, jetons_in, jetons_out).
    Une passe — coupée = non jugée. Une panne d'appel (réseau, surcharge, clé) est une faute de TUILE comme une
    autre : rendue, jamais levée."""
    try:
        client = anthropic.Anthropic(api_key=api_key, timeout=600.0)
        # v2.51 — l'effort « high » seulement quand le modèle est Opus 5.5. La réflexion compte dans max_tokens (8 000) : une
        # réponse coupée est une tuile non jugée, dite et reprise une fois (v2.43) — rien n'est effacé.
        rep = client.messages.create(
            model=_S243_MODELE, max_tokens=8000,
            **({"output_config": {"effort": "high"}} if _S243_MODELE == "claude-opus-5-5" else {}),
            messages=[{"role": "user", "content": [
                {"type": "text", "text": _S243_DEMANDE + liste_txt},
                {"type": "image", "source": {"type": "base64",
                                             "media_type": "image/png",
                                             "data": base64.b64encode(png).decode()}}]}])
    except Exception as e:
        return None, "panne d'appel : %s" % type(e).__name__, 0, 0
    j_in = getattr(rep.usage, "input_tokens", 0) or 0
    j_out = getattr(rep.usage, "output_tokens", 0) or 0
    if rep.stop_reason == "max_tokens":
        return None, "réponse coupée", j_in, j_out
    texte = "".join(b.text for b in rep.content if getattr(b, "type", "") == "text")
    verdicts, faute = _s243_depouille(texte, numeros)
    return verdicts, faute, j_in, j_out


def _s243_point_sur_trace(path):
    it = path["items"][0]
    if it[0] == "l":
        return ((it[1].x + it[2].x) / 2, (it[1].y + it[2].y) / 2)
    if it[0] == "re":
        return (it[1].x0, (it[1].y0 + it[1].y1) / 2)
    if it[0] == "qu":
        return (it[1].ul.x, it[1].ul.y)
    return (it[1].x, it[1].y)


def _s243_points_candidats(path):
    """Les points possibles où poser la gomme, tous SUR la géométrie de CE
    tracé. Le premier est le point d'hier (_s243_point_sur_trace)."""
    pts = [_s243_point_sur_trace(path)]
    for it in path["items"]:
        if it[0] == "l":
            pts += [(it[1].x, it[1].y), (it[2].x, it[2].y)]
        elif it[0] == "re":
            r = it[1]
            pts += [((r.x0 + r.x1) / 2, r.y0), (r.x1, (r.y0 + r.y1) / 2),
                    ((r.x0 + r.x1) / 2, r.y1), (r.x0, r.y0), (r.x1, r.y0),
                    (r.x0, r.y1), (r.x1, r.y1)]
        elif it[0] == "qu":
            q = it[1]
            pts += [(q.ul.x, q.ul.y), (q.ur.x, q.ur.y),
                    (q.ll.x, q.ll.y), (q.lr.x, q.lr.y)]
        else:
            pts.append((it[1].x, it[1].y))
    return pts


def _s243_point_choisi(path, a_eviter):
    """Le point de gomme qui ne touche AUCUNE boîte à préserver ; à défaut le point d'hier, et l'appelant le compte
    (gommes_sans_abri). Condition : la gomme REMOVE_IF_TOUCHED emporte TOUT tracé touché par son point, pas seulement
    sa cible. Effet : le point se pose à l'écart des tracés préservés. (voir JOURNAL BACKEND v2.43)"""
    for (x, y) in _s243_points_candidats(path):
        r = (x - .3, y - .3, x + .3, y + .3)
        if not any(_s243_proche(r, b, 0) for b in a_eviter):
            return (x, y), True
    return _s243_point_sur_trace(path), False


def _s243_proche(a, b, d):
    return not (a[2] + d < b[0] or b[2] + d < a[0]
                or a[3] + d < b[1] or b[3] + d < a[1])


# ═══ v2.70 — BLOC 1, LIGNE 1.12 : LA GRILLE DE NOTATION RECONNUE PAR SA DISPOSITION (DEC-2804 : « D5 sur tous les morceaux ∧ un quadrillage de
# cases » ; DEC-2439 : une grille de notation n'est jamais retirée). Recopié du banc par programme, sans changement de sens : D5 (agents270/A9_grille2,
# a9_play.py : F, N, Bh ; a8_play.py : inside, ncols) et le quadrillage Q = Q1 ∨ Q2 (agents272/A22_grille_cases/a22_quad.py). Une FORME (une disposition),
# provisoire par construction. Le pourquoi et la mesure (23 grilles sur 23, 0 faux sur les 129) : JOURNAL BACKEND v2.70.
def _v270g_lignes_page(page):
    L = []
    for b in page.get_text('dict')['blocks']:
        for l in b.get('lines', []):
            t = ''.join(s['text'] for s in l['spans']).strip()
            if t: L.append((tuple(l['bbox']), t, bool(l['spans'][0]['flags'] & 16)))
    return L
def _v270g_inside(L, rc):
    out = []
    for bb, t, _g in L:
        cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
        if rc[0] - 1 <= cx <= rc[2] + 1 and rc[1] - 1 <= cy <= rc[3] + 1: out.append((bb, t))
    return out
def _v270g_ncols(L):
    iv = sorted((bb[0], bb[2]) for bb, _ in L)
    n = 0; end = None
    for a, b in iv:
        if end is None or a - end >= 8: n += 1; end = b
        else: end = max(end, b)
    return n
def _v270g_d5(page, rc, L):
    """D5 = F ∧ N ∧ Bh (a9_play.py)."""
    H = page.rect.height; I = _v270g_inside(L, rc)
    F = (rc[3] - rc[1]) <= 0.20 * H and (rc[2] - rc[0]) >= 80 and len(I) >= 3 and _v270g_ncols(I) >= 3 and all(len(t) <= 20 for _, t in I)
    if not F: return False
    if any(re.match(r'^[\d\s.,]+$', t) for _, t in I) or any(re.search(r'[….…_]{2,}|…|\.{2,}|_{2,}', t) for _, t in I) or any(re.search(r'[=+×÷]', t) for _, t in I): return False
    for bb, t, gras in L:
        if not gras or bb[2] < rc[0] or bb[0] > rc[2]: continue
        if rc[0] - 1 <= (bb[0] + bb[2]) / 2 <= rc[2] + 1 and rc[1] - 1 <= (bb[1] + bb[3]) / 2 <= rc[3] + 1: continue
        if rc[1] - 30 <= bb[3] <= rc[1] + 1: return False
    return True
_V270G_TAU = 1.5
_V270G_MIN_ABS = 10.0
_V270G_MIN_REL = 0.20
_V270G_GRIS = 190
_V270G_TROU = 2
_V270G_BANDE_MAX = 6
_V270G_SCALE = 2.0

def _v270g_ells(R):
    return max(_V270G_MIN_ABS, _V270G_MIN_REL * (R[2] - R[0])), max(_V270G_MIN_ABS, _V270G_MIN_REL * (R[3] - R[1]))   # (ell_H, ell_V)

def _v270g_rplus(R, pg):
    r = pg.rect
    return (max(r.x0, R[0] - _V270G_TAU), max(r.y0, R[1] - _V270G_TAU), min(r.x1, R[2] + _V270G_TAU), min(r.y1, R[3] + _V270G_TAU))

# ---------- pipeline commun : traits (c, a, b) -> decision
def _v270g_lignes(segs, ell):
    """segs : liste de (c, a, b), c = ordonnee (H) ou abscisse (V), [a,b] etendue. Rend [(c_moyen, [(a,b) troncons >= ell])]."""
    segs = sorted(segs)
    groupes = []; last = None
    for s in segs:
        if groupes and s[0] - last <= _V270G_TAU: groupes[-1].append(s)
        else: groupes.append([s])
        last = s[0]
    out = []
    for g in groupes:
        c = sum(s[0] for s in g) / len(g)
        iv = sorted((s[1], s[2]) for s in g)
        pieces = []
        for a, b in iv:
            if pieces and a - pieces[-1][1] <= _V270G_TAU: pieces[-1][1] = max(pieces[-1][1], b)
            else: pieces.append([a, b])
        pieces = [(a, b) for a, b in pieces if (b - a) >= ell]
        if pieces: out.append((c, pieces))
    return out

def _v270g_decide(H, V, ellH, ellV):
    HL = _v270g_lignes(H, ellH); VL = _v270g_lignes(V, ellV)
    info = dict(nH=len(HL), nV=len(VL), maxcols=0)
    if len(HL) < 2 or len(VL) < 3: return False, info
    def croise(hl, vl):
        y = hl[0]; x = vl[0]
        okx = any(a - _V270G_TAU <= x <= b + _V270G_TAU for a, b in hl[1])
        oky = any(a - _V270G_TAU <= y <= b + _V270G_TAU for a, b in vl[1])
        return okx and oky
    M = [[croise(h, v) for v in VL] for h in HL]
    best = 0
    for i in range(len(HL)):
        for k in range(i + 1, len(HL)):
            n = sum(1 for j in range(len(VL)) if M[i][j] and M[k][j])
            if n > best: best = n
    info['maxcols'] = best
    return best >= 3, info

# ---------- Q1
class _v270g_PageQ:
    def __init__(self, pg):
        self.pg = pg
        assert abs(pg.rect.x0) < 1e-6 and abs(pg.rect.y0) < 1e-6, ('origine de page non nulle', tuple(pg.rect))
        H = []; V = []
        for d in pg.get_drawings():
            t = d.get('type', '') or ''
            trace = 's' in t
            for it in d['items']:
                k = it[0]
                if k == 'l':
                    p1, p2 = it[1], it[2]
                    dx = p2.x - p1.x; dy = p2.y - p1.y
                    if abs(dy) <= 0.6: H.append(((p1.y + p2.y) / 2, min(p1.x, p2.x), max(p1.x, p2.x)))
                    elif abs(dx) <= 0.6: V.append(((p1.x + p2.x) / 2, min(p1.y, p2.y), max(p1.y, p2.y)))
                elif k in ('re', 'qu'):
                    if k == 're': r = it[1]
                    else:
                        q = it[1]
                        xs = [q.ul.x, q.ur.x, q.ll.x, q.lr.x]; ys = [q.ul.y, q.ur.y, q.ll.y, q.lr.y]
                        if abs(q.ul.x - q.ll.x) > 0.01 or abs(q.ur.x - q.lr.x) > 0.01 or abs(q.ul.y - q.ur.y) > 0.01 or abs(q.ll.y - q.lr.y) > 0.01: continue
                        r = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
                    w = r.x1 - r.x0; h = r.y1 - r.y0
                    if trace:
                        H.append((r.y0, r.x0, r.x1)); H.append((r.y1, r.x0, r.x1)); V.append((r.x0, r.y0, r.y1)); V.append((r.x1, r.y0, r.y1))
                    elif min(w, h) <= 1.5:
                        if w >= h: H.append(((r.y0 + r.y1) / 2, r.x0, r.x1))
                        else: V.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
                # 'c' ignore
        self.H = _np247.array(H, dtype=float).reshape(-1, 3); self.V = _np247.array(V, dtype=float).reshape(-1, 3)
        self._gray = None

    def q1(self, R):
        P = _v270g_rplus(R, self.pg); eH, eV = _v270g_ells(R)
        h = self.H
        m = (h[:, 0] >= P[1]) & (h[:, 0] <= P[3]) & (h[:, 2] >= P[0]) & (h[:, 1] <= P[2])
        Hs = [(c, max(a, P[0]), min(b, P[2])) for c, a, b in h[m]]
        v = self.V
        m = (v[:, 0] >= P[0]) & (v[:, 0] <= P[2]) & (v[:, 2] >= P[1]) & (v[:, 1] <= P[3])
        Vs = [(c, max(a, P[1]), min(b, P[3])) for c, a, b in v[m]]
        return _v270g_decide(Hs, Vs, eH, eV)

    # ---------- Q2
    def gray(self):
        if self._gray is None:
            pix = self.pg.get_pixmap(matrix=fitz.Matrix(_V270G_SCALE, _V270G_SCALE), colorspace=fitz.csGRAY, alpha=False, annots=False)
            assert pix.n == 1
            a = _np247.frombuffer(pix.samples, dtype=_np247.uint8)
            self._gray = a.reshape(pix.height, pix.stride)[:, :pix.width]
        return self._gray

    def q2(self, R):
        P = _v270g_rplus(R, self.pg); eH, eV = _v270g_ells(R)
        g = self.gray()
        j0 = int(_np247.floor(P[0] * _V270G_SCALE)); j1 = int(_np247.ceil(P[2] * _V270G_SCALE)); i0 = int(_np247.floor(P[1] * _V270G_SCALE)); i1 = int(_np247.ceil(P[3] * _V270G_SCALE))
        j1 = min(j1, g.shape[1]); i1 = min(i1, g.shape[0])
        if j1 - j0 < 2 or i1 - i0 < 2: return False, dict(nH=0, nV=0, maxcols=0, vide=True)
        M = g[i0:i1, j0:j1] <= _V270G_GRIS
        # traits horizontaux : coordonnee = ordonnee ; verticaux : on transpose
        Hs = _v270g_bandes(M, int(_np247.ceil(eH * _V270G_SCALE)), i0, j0)
        Vs = _v270g_bandes(M.T, int(_np247.ceil(eV * _V270G_SCALE)), j0, i0)
        return _v270g_decide(Hs, Vs, eH, eV)

def _v270g_runs_lignes(M):
    h, w = M.shape
    Pm = _np247.zeros((h, w + 2), dtype=_np247.int8); Pm[:, 1:w + 1] = M
    d = _np247.diff(Pm, axis=1)
    st = _np247.argwhere(d == 1); en = _np247.argwhere(d == -1)
    assert st.shape == en.shape
    return st[:, 0], st[:, 1], en[:, 1]

def _v270g_bandes(M, minpx, off_c, off_e):
    """M : masque (_v270g_lignes x colonnes) ; extrait les _v270g_bandes de suites sombres le long des _v270g_lignes. Rend [(c_pt, a_pt, b_pt)]."""
    if not M.any(): return []
    rows, a, b = _v270g_runs_lignes(M)
    # fusion des trous <= _V270G_TROU dans chaque rangee
    merged = []   # (row, a, b)
    cur = None
    for r, x0, x1 in zip(rows.tolist(), a.tolist(), b.tolist()):
        if cur is not None and cur[0] == r and x0 - cur[2] <= _V270G_TROU: cur[2] = x1
        else:
            if cur is not None and (cur[2] - cur[1]) >= minpx: merged.append(tuple(cur))
            cur = [r, x0, x1]
    if cur is not None and (cur[2] - cur[1]) >= minpx: merged.append(tuple(cur))
    # _v270g_bandes : rangees contigues (ecart <= 1) dont les suites se recouvrent (>= 50 % de la plus courte)
    actives = []; finies = []
    for r, x0, x1 in merged:
        placed = False
        for bd in actives:
            if bd['last'] == r - 1 or (bd['last'] == r - 1):
                ov = min(x1, bd['x1']) - max(x0, bd['x0'])
                if ov > 0 and ov >= 0.5 * min(x1 - x0, bd['x1'] - bd['x0']) and bd['upd'] != r:
                    bd['y1'] = r; bd['last'] = r; bd['upd'] = r; bd['x0'] = min(bd['x0'], x0); bd['x1'] = max(bd['x1'], x1); placed = True; break
        if not placed: actives.append(dict(y0=r, y1=r, last=r, upd=r, x0=x0, x1=x1))
        # retire les _v270g_bandes terminees
        keep = []
        for bd in actives:
            if bd['last'] < r - 1: finies.append(bd)
            else: keep.append(bd)
        actives = keep
    finies += actives
    out = []
    for bd in finies:
        if bd['y1'] - bd['y0'] + 1 > _V270G_BANDE_MAX: continue
        c = off_c / _V270G_SCALE + (bd['y0'] + bd['y1'] + 1) / 2.0 / _V270G_SCALE
        out.append((c, off_e / _V270G_SCALE + bd['x0'] / _V270G_SCALE, off_e / _V270G_SCALE + bd['x1'] / _V270G_SCALE))
    return out

def _v270_grille_de_notation(page, rc):
    """DEC-2804 : D5 ∧ (Q1 ∨ Q2) sur le rectangle rc (en points de page) ; mémorisé par page."""
    C = _v260_cache(page, "v270_grille")
    k = (page.number, tuple(round(float(v), 2) for v in rc))
    if k not in C:
        if "_L" not in C or C.get("_p") != page.number: C["_L"] = _v270g_lignes_page(page); C["_Q"] = _v270g_PageQ(page); C["_p"] = page.number
        C[k] = _v270g_d5(page, rc, C["_L"]) and (C["_Q"].q1(rc)[0] or C["_Q"].q2(rc)[0])
    return C[k]


# ═══ BLOC 5, LIGNE 5.1 — S4d-c4, PART SERVEUR : LE CHAMP « couleur_ou_dessin » POSÉ SUR CHAQUE IMAGE ═══
# Ses décisions : DEC-2807 (« S4d-c4 : l'application de DEC-2412 (la règle des 60 % retirée), de DEC-1238 et de DEC-2778 ; tenu en deux temps. »),
# DEC-3264 j (« S4d-c4 se code sans la grille (DEC-2412, DEC-2931 « d5 seul »), à condition que la version contienne D5. »), DEC-3276 R5 (« un tableau
# qui n'est que des chiffres ou des mots se reconstruit (DEC-1238) »). La page (10.442, PP2) lit `i.couleur_ou_dessin!==false` : une valeur absente vaut
# « couleur ou dessin » (l'ancien comportement) ; `false` = un support qui n'est que des mots, des chiffres ou du gris : le tableau que le modèle refait reste.
# Recopié du banc par programme (ps9_fragments.py), sans changement de sens : le classifieur de T5 (agents267/T5_tableaux/mesure_t5.py : classe, verdict),
# S4d-c2 (outils270/fabrique_a2_270.py : un tracé en courbes qui entoure un mot de 5 mm ou plus est un cadre, pas un dessin ; cd, rect_figure) et la boucle
# de S4d-c3 (outils270/fabrique_a3_270.py : une figure reconnue par D5 reste image). Une FORME (une apparence : remplissage, pixels chromatiques, image posée,
# courbes, obliques, disposition d'une grille de notation), provisoire par construction. Écarts au banc, tous déclarés : les noms ; `t5.classe` et `t5.verdict`
# deviennent des appels directs ; le D5 d'A9 est le `_v270g_d5` déjà porté par la ligne 1.12 (même fonction, appelée sur les lignes de la page, gardées par
# page comme `lines_cache` au banc) ; les pages sont lues dans le document déjà ouvert par `parse_pdf` ; une panne se déclare (couleur_ou_dessin_erreur), /parse continue.
# La déclaration lit le PDF d'ORIGINE : elle ne dépend ni de la gomme, ni du modèle.

_S4D_MM = 72 / 25.4

# VITESSE (21.10.2026, DEC-3337 Q13 a : « même résultat, prouvé égal au banc ») : la valeur posée sur chaque image est celle du texte du banc ; seul le calcul change.
# Pourquoi on peut le calculer par morceaux : `couleur_ou_dessin` est vrai si UN de ces faits est vrai (le classifieur de T5, puis S4d-c2, puis D5 de S4d-c3) :
# une image posée dont le centre est dans le rectangle (raster) · un tracé en courbes dont le centre y est et qui n'entoure pas un mot de 5 mm ou plus ·
# un tracé rempli d'une couleur (écart de composantes > 0,15) plus petit que le rectangle · un trait oblique d'au moins 8 pt · 0,05 % de pixels chromatiques
# (écart > 40) dans le rectangle rendu à 72 ppp · une grille de notation (D5) ; le reste du dictionnaire du banc (couleur_fill, mots, traits_h, traits_v, vrai_tableau...) ne
# sert qu'au verdict en mots (GRIS_SEUL, GRILLE_NUE, TEXTE_SEUL), qui rend `faux` dans les trois cas : il n'est plus calculé. Ce qui est refait pour chaque image au banc
# (get_drawings trois fois, get_images et get_image_rects, le texte de la page, le comptage des pixels en Python) est fait UNE fois par page ; ce qui est décidé pour un
# rectangle (une figure en compte souvent plusieurs, avec le même rectangle) l'est une fois ; le comptage des pixels (le plus cher) n'a lieu que si rien d'autre n'a répondu.
# Le calcul de chaque fait appelle les mêmes fonctions de PyMuPDF sur les mêmes valeurs (rect.contains, get_area, get_pixmap, get_text) : rien n'est réécrit à la main.
# Écart connu, dit : si un calcul ne servant qu'au verdict en mots échouait au banc (un rendu impossible), le banc déclarait la panne ; ici le calcul n'a plus lieu.

class _S4dPage:
    """Ce que la déclaration lit d'UNE page, calculé une fois : les centres des images posées, les tracés qui comptent, les mots (pour les cadres)."""
    def __init__(self, page):
        self.page = page
        self._centres = None
        self._traces = None
        self._mots = None

    def centres_images(self):
        if self._centres is None:
            c = []
            for img in self.page.get_images(full=True):
                for ra in self.page.get_image_rects(img[0]):
                    c.append(fitz.Point((ra.x0 + ra.x1) / 2, (ra.y0 + ra.y1) / 2))
            self._centres = c
        return self._centres

    def traces(self):
        # ne gardent que les tracés qui peuvent compter : en courbes, remplis d'une couleur, ou portant un trait oblique ; chacun : [centre, rect, courbe, rempli_couleur, aire, oblique, cadre]
        if self._traces is None:
            rel = []
            for d in self.page.get_drawings():
                rr = d.get('rect')
                if rr is None: continue
                rr = fitz.Rect(rr)
                items = d.get('items', [])
                courbe = any(it[0] == 'c' for it in items)
                fill = d.get('fill')
                couleur = False
                if fill and not rr.is_empty:
                    comp = [float(v) for v in fill]
                    if not (all(v <= 0.03 for v in comp) or all(v >= 0.97 for v in comp)) and not (min(rr.width, rr.height) * 25.4 / 72 <= 1.0):
                        couleur = (max(comp) - min(comp) > 0.15)
                oblique = False
                for it in items:
                    if it[0] == 'l':
                        a, b = it[1], it[2]
                        if abs(a.x - b.x) >= 1 and abs(a.y - b.y) >= 1 and ((a.x - b.x) ** 2 + (a.y - b.y) ** 2) ** 0.5 >= 8: oblique = True; break
                if courbe or couleur or oblique:
                    rel.append([fitz.Point((rr.x0 + rr.x1) / 2, (rr.y0 + rr.y1) / 2), rr, courbe, couleur, rr.get_area(), oblique, None])
            self._traces = rel
        return self._traces

    def _cadre(self, t):
        # un tracé en courbes d'au moins 5 mm sur 5 mm qui entoure un mot est un cadre, pas un dessin (S4d-c2)
        if t[6] is None:
            if self._mots is None:
                self._mots = [fitz.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2) for w in self.page.get_text('words') if w[4].strip()]
            rr = t[1]
            t[6] = bool(rr.width >= 5 * _S4D_MM and rr.height >= 5 * _S4D_MM and any(rr.contains(p) for p in self._mots))
        return t[6]

    def sans_pixels(self, rect):
        for p in self.centres_images():
            if rect.contains(p): return True
        aire = rect.get_area()
        for t in self.traces():
            if not rect.contains(t[0]): continue
            if t[5]: return True
            if t[3] and t[4] < aire * 1.01: return True
            if t[2] and not self._cadre(t): return True
        return False

    def pixels_chromatiques(self, rect):
        pm = self.page.get_pixmap(clip=rect, matrix=fitz.Matrix(1, 1), alpha=False)
        s = pm.samples; n = pm.width * pm.height; ch = pm.n
        if ch == 3 and len(s) == n * 3:
            a = _np247.frombuffer(s, dtype=_np247.uint8).reshape(n, 3).astype(_np247.int16)
            c = int(((a.max(axis=1) - a.min(axis=1)) > 40).sum())
        else:
            c = 0
            for i in range(0, n * ch, ch):
                a = s[i:i + 3]
                if max(a) - min(a) > 40: c += 1
        return c / max(n, 1)

def _s4d_rect_figure(par_figure, ims, im):
    if not im.get('figure'): return fitz.Rect(im['cadre']) if im.get('cadre') else None
    fr = par_figure.get((im['figure'], im.get('page')), [])
    e = [x['entier_cadre'] for x in fr if x.get('entier_cadre')]
    if e: return fitz.Rect(e[0])
    cs = [x['cadre'] for x in fr if x.get('cadre')]
    if not cs: return None
    r = fitz.Rect(cs[0])
    for c in cs[1:]: r |= fitz.Rect(c)
    return r

def _s4d_declare(doc, images):
    """Pose `couleur_ou_dessin` (un booléen) sur chaque image de `images`, EN PLACE ; rend le nombre d'images déclarées. La boucle de fabrique_a3_270.py (S4d-c3), calculée par page (voir VITESSE ci-dessus)."""
    lignes = {}; n = 0
    par_figure = {}
    for x in images: par_figure.setdefault((x.get('figure'), x.get('page')), []).append(x)
    pages = {}; memo = {}; memo_d5 = {}; rect_figure = {}
    for im in images:
        if im.get('figure'):
            kf = (im['figure'], im.get('page'))
            if kf not in rect_figure: rect_figure[kf] = _s4d_rect_figure(par_figure, images, im)
            r = rect_figure[kf]
        else:
            r = _s4d_rect_figure(par_figure, images, im)
        if r is None or not im.get('page'):
            im['couleur_ou_dessin'] = True; n += 1; continue
        page = doc[im['page'] - 1]
        if page.number not in pages:
            if len(pages) >= 3: pages.pop(next(iter(pages)))
            pages[page.number] = _S4dPage(page)
        P = pages[page.number]
        rc = r & page.rect
        cle = (page.number, tuple(rc))
        if cle not in memo: memo[cle] = P.sans_pixels(rc) or P.pixels_chromatiques(rc) >= 0.0005
        v = memo[cle]
        if not v and im.get('figure') is not None:
            rects = [r] + [fitz.Rect(x['cadre']) for x in par_figure.get((im.get('figure'), im.get('page')), []) if x.get('tableau') and x.get('cadre')]
            kd = (page.number, tuple(tuple(x) for x in rects))
            if kd not in memo_d5:
                if page.number not in lignes: lignes[page.number] = _v270g_lignes_page(page)
                memo_d5[kd] = any(_v270g_d5(page, tuple(x), lignes[page.number]) for x in rects)
            if memo_d5[kd]: v = True
        im['couleur_ou_dessin'] = v; n += 1
    return n


# ═══ v2.70 — BLOC 1, LIGNE 1.1 : LA GOMME BORNÉE PAR LE CODE ; LE DÉCOR EFFACÉ AVANT LA PHOTOGRAPHIE (F6′r + Fa2c2)
# Ses décisions : DEC-1676 (« Le décor doit partir avant que la photographie soit prise. Le modèle ne peut pas effacer un trait à
# l'intérieur d'une image. ») ; DEC-2891 (F6, « la gomme bornée par le code ») ; DEC-2934 (« F6′r et Fa2c2 retenues ») ;
# DEC-3281 (ses règles figées dans l'outil, avec preuves) ; DEC-3283. Seulement AVEC la clé : sans elle, la gomme n'a pas lieu.
#  (G1) le code DÉSIGNE le décor qu'il permet d'effacer — une FORME (une apparence), provisoire par construction (protocole § 2.1) :
#       DF les cadres à tirets d'au moins 30 × 8 mm ; DB les bandeaux d'objectifs (cadre clair en haut de page, à puces, sans ligne de
#       consigne ; la consigne se reconnaît à une LISTE de verbes, provisoire) ; R2 la carte blanche sans texte, son ombre grise
#       translucide et son coin ; R3 la languette « Fiche… » collée au bord et son médaillon ; garde de famille (une famille de tracés
#       n'est désignée que si tous ses membres le sont). PAS les pastilles numérotées (« PAS les pastilles », F6′r).
#  (G2) la gomme BORNÉE : un élément que le modèle juge « ne sert pas » n'est effacé que s'il est désigné ; sinon il « sert ».
#  (G3) plus de « récupérés » : le ramassage après la gomme ajoutait du décor (12 morceaux sur 13, DEC-3017).
#  (G4) Fa2 : l'ombre en mosaïque (≥ 20 tuiles ≤ 6 pt qui se touchent) d'une carte effacée part avec elle.
#  (G5) photographier APRÈS : les rectangles de la découpe sont gardés ; `data` et `data_feuille` sont refaits d'après le PDF effacé,
#       par différence ; Fc2 : un pixel blanc de l'image servie ne devient jamais de l'encre ; `sceau`, `entier`, `entier_feuille` refaits.
# Recopié du banc sans changement de sens (CHANTIER_MESURES_CH1_05-10/BLOC1_GOMME/NOTES_LECTURE_BANC_F6.md) : a50_def (DF),
# a33_detect_b1prime (DB, sa part B1), a75_desig (R2, R3, garde ; R1 non porté, comme a91_forms13r), a75_f6p (bornage, fusion),
# a76_lib et a76_forms (Fa2, Fc2), a50_forme (la différence). DEUX ÉCARTS DÉCLARÉS : le cas manuel de la 062 (boîtes tapées à la main
# au banc) n'est PAS porté ; scipy, absent du serveur, est remplacé par numpy (la dilatation 3 × 3 répétée trois fois = le carré
# de 7 ; les composantes de tuiles = union-find, rangées par leur plus petit indice comme scipy). Le pourquoi et les mesures :
# JOURNAL BACKEND v2.70.
import unicodedata as _ud270
import copy as _cp270
_V270_ACTIVE = True                  # l'interrupteur des mesures (éteint = la gomme de la 2.69)
_V270_PTMM = 25.4 / 72
_V270_K = 2.0                        # 144 ppp : le rendu de `data`
_V270_ECART = 25                     # encre = écart au blanc >= 25
_V270_TUILE_PT = 6.0                 # Fa2 : une tuile d'ombre fait 6 pt de côté au plus
_V270_CONTACT_PT = 0.75              # Fa2 : deux tuiles de la même ombre se touchent à 0,75 pt
_V270_MOSAIQUE_MIN = 20              # Fa2 : une ombre en mosaïque a 20 tuiles au moins


def _v270_r2(r): return tuple(round(float(v), 2) for v in (r.x0, r.y0, r.x1, r.y1)) if hasattr(r, 'x0') else tuple(round(float(v), 2) for v in r)
def _v270_proche(a, b, tol=0.011): return all(abs(x - y) <= tol for x, y in zip(a, b))


# ── (G1) DF : les cadres à tirets d'au moins 30 × 8 mm (a50_def.df_dv, sa part DF) ──
def _v270_dash(d): return str(d.get('dashes') or '').strip()
def _v270_a_des_tirets(d):
    s = _v270_dash(d); return bool(s) and s != '[] 0'
def _v270_df(page):
    df = {}
    for d in page.get_drawings():
        if not _v270_a_des_tirets(d): continue
        r = d.get('rect')
        if r is None: continue
        r = fitz.Rect(r); wmm, hmm = r.width * _V270_PTMM, r.height * _V270_PTMM
        if wmm >= 30 and hmm >= 8:
            df.setdefault((_v270_r2(r), _v270_dash(d)), dict(rect=_v270_r2(r)))
    return list(df.values())


# ── (G1) DB : les bandeaux d'objectifs (a33_detect_b1prime.detecte_page, sa part B1) ──
def _v270_sans_accent(s):
    return ''.join(c for c in _ud270.normalize('NFD', s) if _ud270.category(c) != 'Mn').lower()
# (la liste de verbes à l'impératif d'a33_detect_b1prime, recopiée par programme, provisoire par construction)
_V270_VERBES = ("dessine trace ecris colorie entoure relie complete calcule coche barre souligne cherche range classe associe compte lis decoupe colle mesure reproduis observe "
           "indique reponds resous numerote surligne repasse ajoute enleve decompose compare transforme trouve cite nomme dis fais utilise place pose fabrique construis invente "
           "imagine prolonge termine remplace choisis regarde ecoute").split()
# la liste déclarée finit par « … » : les formes du pluriel (vous) des mêmes verbes, écrites une par une
_V270_PLUR = ("dessinez tracez ecrivez coloriez entourez reliez completez calculez cochez barrez soulignez cherchez rangez classez associez comptez decoupez collez mesurez reproduisez observez "
         "indiquez repondez resolvez numerotez surlignez repassez ajoutez enlevez decomposez comparez transformez trouvez citez nommez dites faites utilisez placez posez fabriquez "
         "construisez inventez imaginez prolongez terminez remplacez choisissez regardez ecoutez lisez").split()
_V270_IMP = set(_V270_VERBES) | set(_V270_PLUR)
_V270_PUCES = set('•●▪◦‣·–-*○■▫□◆◇')
def _v270_est_puce(ch):
    return ch in _V270_PUCES or 0xE000 <= ord(ch) <= 0xF8FF
_V270_RE_NUM = re.compile(r'^\s*\d+\s*[.)]')
_V270_RE_BLANC = re.compile(r'(\.{3,}|…|_{3,})')
def _v270_lignes_b1(page):
    out = []
    for b in page.get_text('dict')['blocks']:
        if b['type'] != 0: continue
        for l in b.get('lines', []):
            t = ''.join(s.get('text', '') for s in l.get('spans', []))
            if t.strip():
                out.append(dict(bb=tuple(l['bbox']), txt=t.strip()))
    return out
def _v270_proche_rect(a, b, tol):
    return all(abs(x - y) <= tol for x, y in zip((a.x0, a.y0, a.x1, a.y1), (b.x0, b.y0, b.x1, b.y1)))
def _v270_dedans(r, big, tol=0.05):
    return r.x0 >= big.x0 - tol and r.y0 >= big.y0 - tol and r.x1 <= big.x1 + tol and r.y1 <= big.y1 + tol
def _v270_analyse_ligne(txt):
    """(est_puce, texte_apres_puce, motif_consigne | None)"""
    t = txt.strip()
    puce = bool(t) and _v270_est_puce(t[0])
    rest = t[1:].strip() if puce else t
    n = _v270_sans_accent(rest)
    mot = re.split(r"[^a-z]+", n.strip())
    mot = mot[0] if mot and mot[0] else ''
    if mot in _V270_IMP: return puce, rest, 'verbe:' + mot
    if _V270_RE_NUM.match(rest): return puce, rest, 'numero'
    if _V270_RE_BLANC.search(rest): return puce, rest, 'a_remplir'
    return puce, rest, None
def _v270_b1(page):
    """Les bandeaux B1 de la page : [dict(cadre, elements=dict(tracés, lignes))]."""
    H = page.rect.height
    drs = page.get_drawings(); L = _v270_lignes_b1(page)
    cands = []
    for d in drs:
        f = d.get('fill'); r = d['rect']
        if f is None or min(f) < 0.95: continue
        if r.width < 60 or r.height < 14 or r.width < 1.5 * r.height: continue
        if r.y0 > 0.25 * H or r.y1 > 0.35 * H: continue
        contour = ('s' in d['type']) or any((e.get('fill') is None and e.get('color') is not None and _v270_proche_rect(e['rect'], r, 0.6)) for e in drs)
        if not contour: continue
        if any(_v270_proche_rect(c, r, 0.6) for c in cands): continue
        cands.append(fitz.Rect(r))
    b1 = []
    for r in cands:
        dans = [l for l in L if _v270_dedans(fitz.Rect(l['bb']), r, 2.0)]
        puces_txt = []; consignes = []
        for l in dans:
            p, rest, cons = _v270_analyse_ligne(l['txt'])
            if p: puces_txt.append(l)
            if cons: consignes.append(l)
        puces_dess = []
        for d in drs:
            f = d.get('fill'); q = d['rect']
            if f is None or not (2 <= q.width <= 8 and 2 <= q.height <= 8 and abs(q.width - q.height) <= 2): continue
            if not _v270_dedans(q, r, 0.5): continue
            for l in dans:
                bb = l['bb']
                if 0 <= bb[0] - q.x1 <= 20 and min(bb[3], q.y1) - max(bb[1], q.y0) >= 0.5 * q.height:
                    puces_dess.append(q); break
        if (bool(puces_txt) or bool(puces_dess)) and not consignes:
            exp = fitz.Rect(r.x0 - 1, r.y0 - 1, r.x1 + 8, r.y1 + 8)
            b1.append(dict(cadre=tuple(round(v, 2) for v in r), elements=dict(
                tracés=[tuple(round(v, 2) for v in q['rect']) for q in drs if _v270_dedans(q['rect'], exp)],
                lignes=[tuple(round(v, 2) for v in l['bb']) for l in dans])))
    return b1


# ── (G1) DF + DB : les cibles de F6 (a50_main.terrain_doc et cibles_depuis_terrain, sortes DF et DB ; sans le cas manuel) ──
def _v270_uniq(L):
    out = []
    for x in L:
        if not any(_v270_proche(x, y, 0.011) for y in out): out.append(x)
    return out
def _v270_cibles_f6(doc, corrige):
    tr, li = {}, {}
    for pno in range(1, len(doc) + 1):
        if pno in corrige: continue
        page = doc[pno - 1]
        b = [e['rect'] for e in _v270_df(page)]; l = []
        for o in _v270_b1(page):
            b += [tuple(x) for x in o['elements']['tracés']]; l += [tuple(x) for x in o['elements']['lignes']]
        b, l = _v270_uniq(b), _v270_uniq(l)
        if b: tr[pno] = b
        if l: li[pno] = l
    return tr, li


# ── (G1) R2 et R3, la garde de famille (a75_desig, sans R1 comme a91_forms13r) ──
def _v270_famille(bb): return (round((bb[2] - bb[0]) * 2) / 2, round((bb[3] - bb[1]) * 2) / 2)
def _v270_aire(b): return max(0.0, b[2] - b[0]) * max(0.0, b[3] - b[1])
def _v270_inter(a, b): return max(0.0, min(a[2], b[2]) - max(a[0], b[0])) * max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
def _v270_dans(l, boite, tol=2.0):
    """la ligne l est dans la boîte (à tol près) : son centre dedans ET 80 % de son aire dedans."""
    cx, cy = (l[0] + l[2]) / 2, (l[1] + l[3]) / 2
    if not (boite[0] - tol <= cx <= boite[2] + tol and boite[1] - tol <= cy <= boite[3] + tol): return False
    return _v270_inter(l, (boite[0] - tol, boite[1] - tol, boite[2] + tol, boite[3] + tol)) >= 0.8 * max(_v270_aire(l), 1e-6)
def _v270_blanc(c): return c is not None and min(c) >= 0.99
def _v270_gris(c): return c is not None and (max(c) - min(c)) <= 0.02 and 0.3 <= c[0] <= 0.7
def _v270_traces_par_boite(page):
    out = {}
    for d in page.get_drawings():
        bb = tuple(round(float(v), 2) for v in d['rect']); e = out.setdefault(bb, dict(bb=bb, fill=None, opacite=None, seqno=d.get('seqno', 0), n_items=0))
        e['seqno'] = min(e['seqno'], d.get('seqno', 0)); e['n_items'] = max(e['n_items'], len(d.get('items') or []))
        if d.get('fill') is not None:
            if e['fill'] is None: e['fill'] = tuple(d['fill'])
            o = d.get('fill_opacity'); o = 1.0 if o is None else o
            e['opacite'] = o if e['opacite'] is None else min(e['opacite'], o)
    return out
def _v270_candidats_page(page):
    """{classe : set de bb de tracés}, {classe : set de boîtes de lignes} pour UNE page, avant la garde de famille (R1 vide)."""
    T = _v270_traces_par_boite(page); L = [(tuple(bb), t) for bb, t in _s243_items_de_page(page)[1]]; Wp, Hp = page.rect.width, page.rect.height
    tr = {'R1': set(), 'R2': set(), 'R3': set()}; li = {'R1': set(), 'R2': set(), 'R3': set()}
    for bb, e in T.items():
        if e['fill'] is None or not _v270_blanc(e['fill']) or _v270_aire(bb) > 40000 or any(_v270_dans(l, bb) for l, t in L): continue
        for bs, es in T.items():
            if bs == bb or not _v270_gris(es['fill']) or es['opacite'] is None or es['opacite'] > 0.6 or es['seqno'] >= e['seqno']: continue
            if abs((bs[2] - bs[0]) - (bb[2] - bb[0])) > 1.0 or abs((bs[3] - bs[1]) - (bb[3] - bb[1])) > 1.0: continue
            if abs(bs[0] - bb[0]) > 10 or abs(bs[1] - bb[1]) > 10: continue
            coins = {bf for bf, ef in T.items() if bf not in (bb, bs) and ef['fill'] is not None and ef['seqno'] > es['seqno']
                     and bf[0] >= bb[0] - 1 and bf[1] >= bb[1] - 1 and bf[2] <= bb[2] + 1 and bf[3] <= bb[3] + 1 and _v270_aire(bf) <= 0.15 * _v270_aire(bb)}
            if not _v270_carte_porte(page, bb, e['seqno'], coins | {bb}): continue
            tr['R2'] |= {bb, bs} | coins
            tr.setdefault('_cartes', set()).add(bb)
    for bb, e in T.items():
        w, h = bb[2] - bb[0], bb[3] - bb[1]
        if e['fill'] is None or w <= 0 or not (h / w >= 4 and 20 <= w <= 60): continue
        if not ((bb[2] >= Wp - 25 or bb[0] <= 25) and (bb[1] <= 1 or bb[3] >= Hp - 1)): continue
        vert = [l for l, t in L if _v270_dans(l, bb)]
        if not vert: continue
        tr['R3'].add(bb); li['R3'] |= {l for l, t in L if _v270_dans(l, bb)}
        for bm, em in T.items():
            if bm == bb: continue
            wm, hm = bm[2] - bm[0], bm[3] - bm[1]
            if not (25 <= wm <= 60 and 25 <= hm <= 60 and 0.75 <= wm / hm <= 1.33) or _v270_inter(bm, bb) < 0.5 * _v270_aire(bm): continue
            ms = [l for l, t in L if _v270_dans(l, bm)]
            if not ms: continue
            tr['R3'].add(bm); li['R3'] |= set(ms)
    if _V270_FC2:
        for bb in _v270_cartes_fc2(page, T, L):
            tr['R2'].add(bb); tr.setdefault('_cartes', set()).add(bb)
            for bf in _v270_charge_utile(T, bb): tr['R2'].add(bf)
    return tr, li
def _v270_carte_porte(page, bb, seqno, exclus):
    """Une carte d'éditeur porte quelque chose peint par-dessus (son mot, en tracés ou en image) ; une case vide à ombre n'est jamais une
    carte : c'est une place de réponse, elle ne part jamais (DEC-3283)."""
    R = fitz.Rect(bb)
    for d in page.get_drawings():
        if d.get('seqno', -1) > seqno and fitz.Rect(d['rect']) in R + (-1, -1, 1, 1) and tuple(round(float(v), 2) for v in d['rect']) not in exclus: return True
    for t, r in page.get_bboxlog():
        if t in ('fill-image', 'fill-imgmask') and fitz.Rect(r).intersects(R) and (fitz.Rect(r) & R).get_area() > 0.5 * fitz.Rect(r).get_area(): return True
    return False
# ── BLOC 1, LIGNE 1.5 — LE MOT « EXERCICES » ET SA CARTE (DEC-2973 : « F1 + FG1 + FC2 retenues ») ; recopié du banc, sans changement de sens :
# FG1 (agents278/A106_G3/a106_gommes.py : la gomme touche chaque sous-tracé), FC2 (a106_cartes.py : une carte reconnue par son coin replié, avec
# son contenu), F1 (a106_forms13._images_cartes, a106_run_v2, agents277/A100_E1/a100_formes.retire_images_page : l'image posée dans une carte
# désignée part avec elle, vérifiée). Le pourquoi et la mesure : JOURNAL BACKEND v2.70.
_V270_FG1 = True
_V270_FC2 = True
_V270_F1 = True
_V270_HORS_TACHE = True
def _v270_debut_fin(it):
    k = it[0]
    if k == 'l': return it[1], it[2]
    if k == 'c': return it[1], it[4]
    return None, None
def _v270_sous_traces(path):
    """FG1 : les éléments de l'objet coupés en sous-tracés : 're' et 'qu' sont un sous-tracé ; sinon rupture de plus de 0,05 pt."""
    subs, cur, last = [], [], None
    for it in path['items']:
        if it[0] in ('re', 'qu'):
            if cur: subs.append(cur); cur = []
            subs.append([it]); last = None; continue
        d, f = _v270_debut_fin(it)
        if d is None:
            cur.append(it); continue
        if last is not None and _m266.hypot(last.x - d.x, last.y - d.y) > 0.05 and cur:
            subs.append(cur); cur = []
        cur.append(it); last = f
    if cur: subs.append(cur)
    return subs
def _v270_contenu(bb, c, tol=1.0, seuil=0.95):
    a = _v270_aire(bb)
    if a <= 0: return False
    ix = max(0.0, min(bb[2], c[2] + tol) - max(bb[0], c[0] - tol)) * max(0.0, min(bb[3], c[3] + tol) - max(bb[1], c[1] - tol))
    return ix >= seuil * a
def _v270_gris_clair(c): return c is not None and (max(c) - min(c)) <= 0.02 and 0.6 <= c[0] <= 0.95
def _v270_cartes_fc2(page, T, L):
    """FC2 : un coin replié gris clair, collé au bord droit ou bas, peint après la carte, d'aire 0,5-15 %, contenu dans la carte, avec un jumeau de contour."""
    cand = {bb: e for bb, e in T.items() if e['fill'] is not None and _v270_blanc(e['fill']) and _v270_aire(bb) <= 40000 and not any(_v270_dans(l, bb) for l, t in L)}
    contours = {tuple(round(float(v), 2) for v in d['rect']) for d in page.get_drawings() if d.get('color') is not None}; out = []
    for bb, e in cand.items():
        ac = _v270_aire(bb)
        for bf, ef in T.items():
            if bf == bb or ef['fill'] is None or not _v270_gris_clair(ef['fill']) or ef['seqno'] <= e['seqno']: continue
            if not (0.005 * ac <= _v270_aire(bf) <= 0.15 * ac) or not _v270_contenu(bf, bb, 1.0, 0.95): continue
            if not (abs(bf[2] - bb[2]) <= 1.0 or abs(bf[3] - bb[3]) <= 1.0): continue
            if bf not in contours: continue
            out.append(bb); break
    return sorted(set(out))
def _v270_charge_utile(T, bb):
    """le contenu de la carte : tout tracé contenu (>= 95 %, 1 pt) dans sa boîte, peint après elle, qui n'est ni elle ni de sa taille (à 1,5 pt)."""
    e = T[bb]; out = []
    for bf, ef in T.items():
        if bf == bb or ef['seqno'] <= e['seqno'] or not _v270_contenu(bf, bb, 1.0, 0.95): continue
        if (bf[2] - bf[0]) >= (bb[2] - bb[0]) - 1.5 and (bf[3] - bf[1]) >= (bb[3] - bb[1]) - 1.5: continue
        out.append(bf)
    return out
def _v270_images_dans(page, bb):
    """F1 : images PDF (côté > 6 pt) contenues (>= 95 %, 1 pt) dans la boîte bb : [(bbox, digest12)]."""
    out = []
    for i in page.get_image_info(hashes=True):
        r = i['bbox']
        if max(r[2] - r[0], r[3] - r[1]) <= 6.0: continue
        if _v270_contenu(tuple(r), bb, 1.0, 0.95): out.append((tuple(round(float(v), 2) for v in r), (i.get('digest') or b'').hex()[:12]))
    return out
def _v270_cle_image(i): return (tuple(round(float(v), 1) for v in i['bbox']), (i.get('digest') or b'').hex() if isinstance(i.get('digest'), (bytes, bytearray)) else str(i.get('digest')))
def _v270_liste_images(page): return sorted(_v270_cle_image(i) for i in page.get_image_info(hashes=True))
def _v270_retire_images_page(doc, pno, cibles):
    """F1 : retire les images visées de la page pno ; d'abord sur une copie : ne touche la vraie page que si la liste des images après = avant moins les cibles."""
    page = doc[pno - 1]; avant = _v270_liste_images(page); visees = []
    for bb, dg in cibles:
        trouve = [k for k in avant if all(abs(a - b) <= 0.15 for a, b in zip(k[0], bb)) and (dg is None or k[1].startswith(dg))]
        if not trouve: return 0
        visees.append(trouve[0])
    attendu = list(avant)
    for k in visees:
        if k in attendu: attendu.remove(k)
    tmp = fitz.open(); tmp.insert_pdf(doc, from_page=pno - 1, to_page=pno - 1); tp = tmp[0]
    for k in visees:
        r = fitz.Rect(*k[0]) + (0.5, 0.5, -0.5, -0.5)
        if r.width > 0.5 and r.height > 0.5: tp.add_redact_annot(r, fill=False)
    tp.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_NONE)
    apres = _v270_liste_images(tp); tmp.close()
    if apres != attendu: return 0
    for k in visees:
        r = fitz.Rect(*k[0]) + (0.5, 0.5, -0.5, -0.5)
        if r.width > 0.5 and r.height > 0.5: page.add_redact_annot(r, fill=False)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_NONE)
    return len(visees)


def _v270_designe_doc(doc, pages):
    """({page: [bb tracés]}, {page: [boîtes lignes]}, nombre de familles refusées) ; garde de famille sur `pages`."""
    cand = {}; tout = {}
    for p in pages:
        cand[p] = _v270_candidats_page(doc[p - 1])
        for bb in _v270_traces_par_boite(doc[p - 1]): tout.setdefault(_v270_famille(bb), []).append((p, bb))
    des_tr = {p: set() for p in pages}; des_li = {p: set() for p in pages}; refusees = set()
    for p in pages:
        tr, li = cand[p]
        for cl in ('R1', 'R2', 'R3'):
            for bb in tr[cl]:
                fam = _v270_famille(bb)
                if all(any(bb2 in cand[p2][0][c2] for c2 in ('R1', 'R2', 'R3')) for p2, bb2 in tout[fam]): des_tr[p].add(bb)
                else: refusees.add((fam, cl))
        for cl in ('R1', 'R3'):
            for l in li[cl]: des_li[p].add(l)
    for p in pages:
        tr, li = cand[p]
        for cl in ('R1', 'R3'):
            for l in list(li[cl]):
                if not any(_v270_dans(l, bb) for bb in des_tr[p] if bb in tr[cl]): des_li[p].discard(l)
    return {p: sorted(v) for p, v in des_tr.items() if v}, {p: sorted(v) for p, v in des_li.items() if v}, len(refusees)
def _v270_fusion(cib, extra_tr, extra_li):
    tr = {p: list(v) for p, v in cib[0].items()}; li = {p: list(v) for p, v in cib[1].items()}
    for p, v in extra_tr.items(): tr[p] = tr.get(p, []) + [b for b in v if not any(_v270_proche(b, c) for c in tr.get(p, []))]
    for p, v in extra_li.items(): li[p] = li.get(p, []) + [b for b in v if not any(_v270_proche(b, c) for c in li.get(p, []))]
    return tr, li
def _v270_hors_tache(page, images, p, bb, panne):
    """DEC-3289, ses mots : « un badge est un décor car il ne fait pas parti de la tâche de l'élève ». Un élément est hors de la tâche
    s'il est dans une image que la lecture des exercices range à aucun exercice (« p<n>-d… », « p<n>-t… ») sur une page qu'elle a lue ;
    une image d'exercice (« p<n>-e… ») en fait partie ; une grille de notation n'en sort jamais (DEC-2439 ; reconnue par sa disposition, DEC-2804 : _v270_grille_de_notation)."""
    if p in panne:
        return False
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    ims = [im for im in images if int(im["page"]) == p and im["cadre"][0] <= cx <= im["cadre"][2] and im["cadre"][1] <= cy <= im["cadre"][3]]
    figs = [str(im.get("figure") or "") for im in ims]
    if not figs or any(f.startswith("p%d-e" % p) for f in figs) or not all(f.startswith(("p%d-d" % p, "p%d-t" % p)) for f in figs):
        return False
    return not any(_v270_grille_de_notation(page, im["cadre"]) for im in ims)


def _v270_designe(des, p, bb):
    """(G2) l'élément de boîte bb, page p, est-il désigné par le code (tracé ou ligne) ?"""
    return any(_v270_proche(bb, x) for x in des[0].get(p, [])) or any(_v270_proche(bb, x) for x in des[1].get(p, []))


# ── (G4) Fa2 : l'ombre en mosaïque part avec sa carte (a76_lib, a76_forms) ──
def _v270_tuiles(page):
    out = []
    for i, (t, r) in enumerate(page.get_bboxlog()):
        if t in ('fill-image', 'fill-imgmask'):
            R = fitz.Rect(r)
            if not R.is_empty and max(R.width, R.height) <= _V270_TUILE_PT: out.append((i, t, tuple(R)))
    return out
def _v270_mosaiques(page):
    """composantes (≥ 20) des tuiles qui se touchent à 0,75 pt près ; union-find (scipy.sparse.csgraph au banc), rangées par leur plus petit indice."""
    ti = _v270_tuiles(page)
    if len(ti) < _V270_MOSAIQUE_MIN: return []
    A = _np247.array([t[2] for t in ti], float); h = _V270_CONTACT_PT / 2.0
    x0, y0, x1, y1 = A[:, 0] - h, A[:, 1] - h, A[:, 2] + h, A[:, 3] + h
    ov = (x0[:, None] < x1[None, :]) & (x0[None, :] < x1[:, None]) & (y0[:, None] < y1[None, :]) & (y0[None, :] < y1[:, None])
    par = list(range(len(ti)))
    def racine(k):
        while par[k] != k:
            par[k] = par[par[k]]; k = par[k]
        return k
    for i, j in zip(*_np247.nonzero(ov)):
        a, b = racine(int(i)), racine(int(j))
        if a != b: par[max(a, b)] = min(a, b)
    groupes = {}
    for k in range(len(ti)): groupes.setdefault(racine(k), []).append(k)
    return [[ti[k] for k in g] for g in sorted(groupes.values(), key=min) if len(g) >= _V270_MOSAIQUE_MIN]
def _v270_boite_union(L):
    u = fitz.Rect(L[0])
    for b in L[1:]: u |= fitz.Rect(b)
    return tuple(u)
def _v270_ombres_de_carte(page):
    mos = _v270_mosaiques(page); dr = page.get_drawings(); out = []
    for m in mos:
        bx = fitz.Rect(_v270_boite_union([t[2] for t in m])); dern = max(t[0] for t in m); carte = None
        for d in sorted((d for d in dr if d.get('fill') is not None and d.get('seqno', -1) > dern), key=lambda d: d['seqno']):
            r = fitz.Rect(d['rect']); inter = r & bx
            if not inter.is_empty and inter.get_area() >= 0.5 * bx.get_area(): carte = dict(boite=tuple(r), fill=d['fill']); break
        out.append(dict(carte=carte, tuiles=[t[2] for t in m]))
    return out
def _v270_carte_effacee(de_page, carte, tol=0.02):
    for d in de_page.get_drawings():
        if d.get('fill') is not None and all(abs(x - y) <= tol for x, y in zip(tuple(d['rect']), carte['boite'])) and all(abs(a - b) < 0.01 for a, b in zip(d['fill'], carte['fill'])): return False
    return True
def _v270_retire_tuiles(doc, pno, boites):
    page = doc[pno - 1]; n = 0
    for b in boites:
        cx, cy = (b[0] + b[2]) / 2, (b[1] + b[3]) / 2; page.add_redact_annot(fitz.Rect(cx - .2, cy - .2, cx + .2, cy + .2), fill=False); n += 1
    if n: page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE, graphics=fitz.PDF_REDACT_LINE_ART_NONE, text=fitz.PDF_REDACT_TEXT_NONE)
    return n


# ── (G5) photographier après, par différence, Fc2 (a50_forme.pages_changees, a76_forms.fusion_diff_Fc sur l'image servie) ──
def _v270_rend(page, z=_V270_K):
    pm = page.get_pixmap(matrix=fitz.Matrix(z, z), colorspace=fitz.csRGB, alpha=False)
    return _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, 3)
def _v270_dilate(m, it=3):
    """binary_dilation(m, carré 3 × 3, it fois) = la dilatation par le carré de 2·it + 1, bord à 0 (égalité prouvée contre scipy)."""
    h, w = m.shape; p = _np247.zeros((h + 2 * it, w + 2 * it), bool); p[it:it + h, it:it + w] = m; out = _np247.zeros_like(m)
    for dy in range(2 * it + 1):
        for dx in range(2 * it + 1): out |= p[dy:dy + h, dx:dx + w]
    return out
def _v270_pages_changees(do, de, pages):
    out = {}
    for pno in pages:
        a = _v270_rend(do[pno - 1]); b = _v270_rend(de[pno - 1])
        h, w = min(a.shape[0], b.shape[0]), min(a.shape[1], b.shape[1])
        m = (a[:h, :w] != b[:h, :w]).any(axis=2)
        out[pno] = _v270_dilate(m) if m.any() else m
    return out
def _v270_arr(pm): return _np247.frombuffer(pm.samples, dtype=_np247.uint8).reshape(pm.height, pm.width, pm.n)[:, :, :3]
def _v270_png_arr(url): return _np247.asarray(_Im247.open(io.BytesIO(base64.b64decode(url.split(',', 1)[1]))).convert('RGB'))
def _v270_photographie_apres(do, de, images):
    """(G5) rend (images neuves, nombre d'images refaites, nombre de champs non rejouables) ; les rectangles ne changent pas."""
    out = _cp270.deepcopy(images); pages = sorted({int(im['page']) for im in images}); ch = _v270_pages_changees(do, de, pages); n_ref = 0; n_nr = 0
    for im in out:
        pno = int(im['page']); M_ = ch[pno]
        if not M_.any(): continue
        ir = (fitz.Rect(im['cadre']) * fitz.Matrix(_V270_K, _V270_K)).irect; x0, y0, x1, y1 = ir.x0, ir.y0, ir.x1, ir.y1
        if not M_[max(0, y0):max(0, y1), max(0, x0):max(0, x1)].any(): continue
        modif = False
        for champ, z in (('data', _V270_K), ('data_feuille', _V238_PPP_FEUILLE / 72.0)):
            v = im.get(champ)
            if not v or ',' not in v: continue
            Va = _v270_png_arr(v)
            pa = _v270_arr(do[pno - 1].get_pixmap(clip=fitz.Rect(im['cadre']), matrix=fitz.Matrix(z, z))); pb = _v270_arr(de[pno - 1].get_pixmap(clip=fitz.Rect(im['cadre']), matrix=fitz.Matrix(z, z)))
            if Va.shape != pa.shape or pa.shape != pb.shape: n_nr += 1; continue
            mask = (pa != pb).any(axis=2)
            mask &= ~(((255 - Va.min(axis=2).astype(_np247.int16)) < _V270_ECART) & ((255 - pb.min(axis=2).astype(_np247.int16)) >= _V270_ECART))   # Fc2 : ce qui était blanc reste blanc
            if not mask.any(): continue
            V2 = Va.copy(); V2[mask] = pb[mask]; h_, w_ = V2.shape[:2]
            if champ == 'data' and ((255 - Va.min(axis=2).astype(_np247.int16)) >= _V270_ECART).any() and not ((255 - V2.min(axis=2).astype(_np247.int16)) >= _V270_ECART).any():
                # une image devenue blanche reste dans la liste et n'est jamais posée [DEC-2902]
                im['vide'] = True
            png = fitz.Pixmap(fitz.csRGB, w_, h_, V2.tobytes(), False).tobytes('png'); im[champ] = 'data:image/png;base64,' + base64.b64encode(png).decode()
            if champ == 'data': im['sceau'] = _hl.sha256(png).hexdigest()[:12]
            modif = True
        if modif: n_ref += 1
    _v260_entier_compose(out); _v262_entier_feuille(out)
    return out, n_ref, n_nr


def _s243_symbiose(content, filename, result, api_key):
    """L'étage entier. Ne lève JAMAIS vers /parse : toute panne rend le
    résultat d'aujourd'hui avec l'erreur DITE dans `symbiose`."""
    info = {"tuiles_jouees": 0, "tuiles_non_jugees": 0, "familles_jugees": 0,
            "textes_juges": 0, "decors_traces": 0, "decors_textes": 0,
            "recuperees": 0, "jetons_entree": 0, "jetons_sortie": 0}
    try:
        images = result.get("images") or []
        pages_ex = sorted({int(f["page"]) for f in images})
        if not pages_ex:
            info["erreur"] = "aucune découpe — étage sans objet"
            result["symbiose"] = info
            return result
        doc = fitz.open(stream=content, filetype="pdf")

        # ── (1) LE MODÈLE VISE — un jugement par famille de jumeaux ──
        familles_doc = {}          # famille -> bbox représentant (page, bbox)
        textes_doc = []            # (page, bbox, txt)
        for pno in pages_ex:
            traces, textes = _s243_items_de_page(doc[pno - 1])
            for bb, fam in traces.items():
                familles_doc.setdefault(fam, (pno, bb))
            textes_doc += [(pno, bb, t) for bb, t in textes]
        info["familles_jugees"] = len(familles_doc)
        info["textes_juges"] = len(textes_doc)
        elems = [("trace", fam, pno, bb) for fam, (pno, bb) in familles_doc.items()] \
            + [("texte", None, pno, bb) for pno, bb, _t in textes_doc]
        verdict_fam, verdict_txt = {}, {}
        par_page = {}
        for e in elems:
            par_page.setdefault(e[2], []).append(e)
        num = 0
        for pno, lot_page in sorted(par_page.items()):
            lot_page.sort(key=lambda e: ((e[3][1] + e[3][3]) / 2, e[3][0]))
            for i0 in range(0, len(lot_page), _S243_TUILE_MAX):
                lot = lot_page[i0:i0 + _S243_TUILE_MAX]
                nums = list(range(num + 1, num + 1 + len(lot)))
                num += len(lot)
                png = _s243_tuile_png(doc[pno - 1],
                                      list(zip(nums, [e[3] for e in lot])))
                liste_txt = " · ".join(
                    "T%d (%.0f×%.0f pt)" % (n, e[3][2] - e[3][0], e[3][3] - e[3][1])
                    for n, e in zip(nums, lot))
                verdicts, faute, j_in, j_out = _s243_juge_tuile(
                    api_key, png, liste_txt, nums)
                info["jetons_entree"] += j_in
                info["jetons_sortie"] += j_out
                # v2.43 — UNE reprise par tuile, bornée et dite : une forme cassée se rejoue une fois ; si elle casse
                # encore, la tuile est NON JUGÉE et DITE. Condition : première passe en faute. Effet : un second appel, jamais plus.
                if verdicts is None:
                    verdicts, faute2, j_in2, j_out2 = _s243_juge_tuile(
                        api_key, png, liste_txt, nums)
                    info["jetons_entree"] += j_in2
                    info["jetons_sortie"] += j_out2
                    info.setdefault("reprises", 0)
                    info["reprises"] += 1
                    faute = faute2 if verdicts is None else faute
                if verdicts is None:
                    info["tuiles_non_jugees"] += 1
                    info.setdefault("fautes", []).append(faute)
                    continue
                info["tuiles_jouees"] += 1
                for n, e in zip(nums, lot):
                    if e[0] == "trace":
                        verdict_fam[e[1]] = verdicts[n]
                    else:
                        verdict_txt[(e[2], e[3])] = verdicts[n]

        # ── v2.70 (G2) — LA GOMME BORNÉE PAR LE CODE (DEC-2891, DEC-2934) : un « ne sert pas » sur un élément que le code ne
        # désigne pas devient « sert » ; la désignation est lue sur le PDF d'origine (DF, DB, R2, R3, garde de famille : le bloc v2.70).
        if _V270_ACTIVE:
            _do0 = fitz.open(stream=content, filetype="pdf")
            _cor = set(int(p) for p in (result.get("corrige_pages") or []))
            _dtr, _dli, _refus = _v270_designe_doc(_do0, pages_ex)
            _des = _v270_fusion(_v270_cibles_f6(_do0, _cor), _dtr, _dli)
            _do0.close()
            info["v270_designes"] = sum(len(v) for v in _des[0].values()) + sum(len(v) for v in _des[1].values())
            info["v270_familles_refusees"] = _refus
            info["v270_gardes"] = 0
            _panne = {int(x[0]) for x in ((result.get("exercices_lecture") or {}).get("pages_en_panne") or [])}
            info["v270_ramenes"] = 0
            for e in elems:
                d_, k_ = (verdict_fam, e[1]) if e[0] == "trace" else (verdict_txt, (e[2], e[3]))
                if d_.get(k_) != "ne sert pas":
                    continue
                if _v270_designe(_des, e[2], e[3]) or (_V270_HORS_TACHE and _v270_hors_tache(doc[e[2] - 1], images, e[2], e[3], _panne)):
                    info["v270_gardes"] += 1
                else:
                    d_[k_] = "sert"
                    info["v270_ramenes"] += 1

        # ── (2) LE CODE DÉCIDE — le caviardage des décors ──
        for pno in pages_ex:
            page = doc[pno - 1]
            n_tr = 0
            dessins = page.get_drawings()
            a_eviter, cibles = [], []
            for path in dessins:
                r = path["rect"]
                bb = (round(r.x0, 2), round(r.y0, 2), round(r.x1, 2), round(r.y1, 2))
                if verdict_fam.get(_s243_famille(bb)) == "ne sert pas":
                    cibles.append(path)
                else:
                    a_eviter.append((r.x0 - .3, r.y0 - .3, r.x1 + .3, r.y1 + .3))
            for path in cibles:
                for _sub in (_v270_sous_traces(path) if (_V270_ACTIVE and _V270_FG1) else [None]):
                    (x, y), abri = _s243_point_choisi(path if _sub is None else {"items": _sub, "rect": path["rect"]}, a_eviter)
                    if not abri:
                        info["gommes_sans_abri"] = info.get("gommes_sans_abri", 0) + 1
                    page.add_redact_annot(fitz.Rect(x - .3, y - .3, x + .3, y + .3))
                n_tr += 1
            if n_tr:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                      graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED,
                                      text=fitz.PDF_REDACT_TEXT_NONE)
                # v2.43 — la gomme DÉCLARE ce qu'elle a réellement emporté : attendu = cibles posées ; emporté = tracés
                # réellement partis.
                emp = len(dessins) - len(page.get_drawings())
                info["caviardage_attendu"] = info.get("caviardage_attendu", 0) + n_tr
                info["caviardage_emporte"] = info.get("caviardage_emporte", 0) + emp
            info["decors_traces"] += n_tr
            n_tx = 0
            for (p2, bb), v in verdict_txt.items():
                if p2 == pno and v == "ne sert pas":
                    page.add_redact_annot(fitz.Rect(*bb))
                    n_tx += 1
            if n_tx:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                                      graphics=fitz.PDF_REDACT_LINE_ART_NONE)
            info["decors_textes"] += n_tx

        # ── v2.70 (G4, G5) — Fa2 puis PHOTOGRAPHIER APRÈS (DEC-1676) : l'ombre en mosaïque d'une carte effacée part avec elle ; les
        # pixels des images sont refaits d'après le PDF effacé, par différence, sans jamais ajouter d'encre (Fc2).
        if _V270_ACTIVE:
            _do = fitz.open(stream=content, filetype="pdf")
            _de = fitz.open(stream=doc.tobytes(), filetype="pdf")
            _nt = 0
            for _p in range(1, len(_do) + 1):
                for _o in [o for o in _v270_ombres_de_carte(_do[_p - 1]) if o["carte"] is not None]:
                    if _v270_carte_effacee(_de[_p - 1], _o["carte"]):
                        _nt += _v270_retire_tuiles(_de, _p, list(_o["tuiles"]))
            if _nt:
                _de = fitz.open(stream=_de.tobytes(), filetype="pdf")
            if _V270_F1:
                _dtr1 = _v270_designe_doc(_do, list(range(1, len(_do) + 1)))[0]; _nf1 = 0
                for _p in range(1, len(_do) + 1):
                    _cartes = [bb for bb in (_v270_candidats_page(_do[_p - 1])[0].get('_cartes') or ()) if any(_v270_proche(bb, x) for x in _dtr1.get(_p, []))]
                    _cib = []
                    for bb in sorted(_cartes): _cib += _v270_images_dans(_do[_p - 1], bb)
                    if _cib: _nf1 += _v270_retire_images_page(_de, _p, _cib)
                if _nf1:
                    _de = fitz.open(stream=_de.tobytes(), filetype="pdf")
                info["v270_images_de_cartes"] = _nf1
            images, info["v270_images_refaites"], info["v270_non_rejouables"] = _v270_photographie_apres(_do, _de, images)
            info["v270_tuiles_retirees"] = _nt
            info["v270_images_vides"] = sum(1 for _im in images if _im.get("vide"))
            _do.close()
            _de.close()
        doc.close()
        result["images"] = images
        result["num_images"] = len(images)
        result["symbiose"] = info
        return result
    except Exception as e:
        info["erreur"] = "étage en panne — parse d'aujourd'hui rendu : %s" % e
        result["symbiose"] = info
        return result


@app.post("/parse")
def parse_document(file: UploadFile = File(...),
                         api_key: str = Form(None), request: Request = None):  # v2.84 — la requête en dernier et facultative : un appel direct (banc, batterie) reste possible ; par le réseau, FastAPI la donne toujours
    """
    Reçoit un PDF (extraction PyMuPDF, puis le pipeline PDF). Tout autre fichier est refusé (415) : l'outil n'accepte que des PDF (DEC-1249).
    Retourne une structure JSON avec blocs texte, images, tableaux,
    type d'exercice détecté, et positions relatives.
    """
    content = file.file.read()
    filename = file.filename or "document"
    ext = filename.rsplit(".", 1)[-1].lower()

    # PDF : extraction directe PyMuPDF (images raster + figures vectorielles)
    # v2.43 — avec une clé, l'étage de la symbiose (décors jugés, gommés avant la photographie : v2.70) ;
    # sans clé, rien ne change. (voir JOURNAL BACKEND v2.43)
    if ext == "pdf":
        # v2.83 — LA PORTE : plus de 30 pages, refusé avant la découpe (DEC-3492)
        try:
            _p_n = fitz.open(stream=content, filetype="pdf").page_count
        except Exception:
            _p_n = 0
        if _p_n > _P_PAGES_MAX:
            raise HTTPException(413, "Le document a %d pages : la limite est de 30. Divisez-le en plusieurs documents." % _p_n)
        result = parse_pdf(content, filename, api_key=api_key)
        if api_key:
            result = _s243_symbiose(content, filename, result, api_key)
            # v2.60 — après la symbiose aussi, l'image entière est faite des morceaux posés ; v2.62 : et l'entière de feuille
            _v260_entier_compose(result.get("images") or [])
            _v262_entier_feuille(result.get("images") or [])
            # v2.69 — B-N2 + Wc2 : la recomposition a refait les images entières ; les gestes des endroits du signe y sont rejoués (voir _v269_bn2_rejoue)
            try:
                _v269_bn2_rejoue(result.get("images") or [], result.get("endroits_signe"), fitz.open(stream=content, filetype="pdf"))
            except Exception as _e269br:
                result["endroits_signe_erreur"] = ((result.get("endroits_signe_erreur") or "") + " ; rejeu après la gomme — %s: %s" % (type(_e269br).__name__, _e269br)).lstrip(" ;")
            # v2.69 — G2 : la recomposition a refait les images entières ; l'effacement des mots cursifs y est rejoué (voir _v269_g2_rejoue)
            try:
                _v269_g2_rejoue(result.get("images") or [], result.get("effacements_cursifs"), fitz.open(stream=content, filetype="pdf"))
            except Exception as _e269r:
                result["effacements_cursifs_erreur"] = ((result.get("effacements_cursifs_erreur") or "") + " ; rejeu après la gomme — %s: %s" % (type(_e269r).__name__, _e269r)).lstrip(" ;")
        # PP7 — la chaîne F-b + l8 + n2 (voir _pp7_chaine) : avec ou sans clé, après la gomme et ses recompositions, AVANT la marque du morceau, `lignes_image` et `bareme` (les champs qui lisent l'image la lisent après la chaîne) ; une panne est dite dans `chaine_fb_erreur`
        result = _pp7_chaine(content, result)
        # v2.69 — la marque du morceau, (c) corrigée (voir _a230_marques) : avec ou sans clé ; une panne est dite dans `a230_erreur`
        result = _a230_marques(content, result)
        # R3b — la place des lignes de texte de chaque image (voir _r3b_lignes_images) : avec ou sans clé ; une panne est dite dans `lignes_image_erreur`
        result = _r3b_lignes_images(content, result)
        # R9 (DEC-3157) — le champ `bareme` de chaque image : les « /N » du PDF que l'image envoyée montre ; avec ou sans clé ; une panne est dite dans `a225_erreur`
        result = _a225_bareme_champ(content, result)
        # v2.83 — LA PORTE : la dépense de l'analyse s'inscrit au compte du code
        if request is not None:  # v2.84 — un appel direct (sans la porte) n'inscrit rien
            _p_depense(getattr(request.state, "tailory_code", "?"), result)
        return result

    # DEC-1249 (« L'outil n'accepte que des PDF ») et la réponse de Catherine du 21.10.2026 (Q15) : tout ce qui n'est pas un PDF est refusé clairement, avant toute lecture.
    raise HTTPException(
        415, "Ce fichier (« %s ») n'est pas un PDF, et l'outil n'accepte que des PDF. "
             "Dans votre traitement de texte, choisissez « Enregistrer sous » ou « Exporter en PDF », puis déposez le PDF obtenu." % filename)


import unicodedata


# ─────────────────────────────────────────────
# ENDPOINT : /health
# ─────────────────────────────────────────────

@app.post("/figures_vues")
async def figures_vues(file: UploadFile = File(...)):
    # v2.30 — LA VOIX DU DIAGNOSTIC. Condition : un PDF reçu. Effet : la liste
    # des figures telle que parse_pdf la fabrique (MÊME chemin de code, aucune
    # découpe parallèle), sans les images elles-mêmes — ordre, page, cadre en
    # pleines décimales, sceau du contenu, et l'empreinte du fichier reçu.
    # (voir JOURNAL BACKEND v2.30)
    content = await file.read()
    if not (file.filename or "").lower().endswith(".pdf"):
        raise HTTPException(status_code=400,
                            detail="figures_vues ne lit que des PDF")
    r = parse_pdf(content, file.filename)
    return {
        "version": VERSION, "empreinte": EMPREINTE, "lecteur_pdf": LECTEUR_PDF,
        "fichier": {"nom": file.filename,
                    "sha256": _hl.sha256(content).hexdigest()[:12],
                    "octets": len(content)},
        "nombre": len(r["images"]),
        "figures": [{"ordre": i["index"], "page": i["page"],
                     "cadre": i["cadre"], "sceau": i["sceau"],
                     "w": i["w"], "h": i["h"]}
                    for i in r["images"]],
    }


@app.get("/health")
def health():
    # V2.10 — la version du module grilles est remontée telle qu'elle est
    # DANS le fichier déployé, jamais recopiée à la main : c'est ce qui
    # permet de savoir, en ligne, laquelle tourne vraiment.
    return {"status": "ok", "version": VERSION, "empreinte": EMPREINTE, "lecteur_pdf": LECTEUR_PDF,
            # v2.44 — change à CHAQUE démarrage du processus : c'est ce qui permet
            # à la page de savoir que sa demande est morte avec l'ancien serveur.
            "instance": DEMARRAGE,
            # v2.44 — information seulement : change au REMPLACEMENT de l'instance,
            # pas au redémarrage du processus. Ne juge rien.
            "instance_hebergeur": INSTANCE_HEBERGEUR,
            "grilles": (getattr(_grilles, "VERSION", "inconnue")
                        if GRILLES_ACTIVES else "absent"),
            "formes": (getattr(_formes, "VERSION", "inconnue")
                       if FORMES_ACTIVES else "absent")}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
