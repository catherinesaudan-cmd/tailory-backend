#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GRILLES ET FORMES COMPOSITES — module serveur, version 2.14
===========================================================

VERSION 2.14 — 17 aout 2026 (reprise 89). Le 2.12 est du 11 aout 2026.
Fichier livre sous le nom `grilles_v2_12.py` ; **a renommer `grilles.py`**
au deploiement, a cote d'`app.py` (c'est sous ce nom qu'il est importe).
La version est lisible en ligne : `/health` la renvoie.

Journal de la journee du 27.07 (trois etats successifs, celui-ci est le
dernier — les deux premiers ne doivent plus servir) :
  1. quadrillages et tableaux dessines en traits ;
  2. + quadrillages arrives en IMAGE, decouverts sur l'evaluation de
     geometrie 5P, ou le quadrillage de l'exercice 8 est un dessin importe ;
  3. + l'interdiction de reporter une mesure sur la feuille eleve, collee a
     la mesure elle-meme (arbitrage de Catherine du 27.07).


Ce module repare le dernier defaut grave ouvert : les quadrillages et les
tableaux d'une source arrivent aujourd'hui au modele soit en miettes, soit
en file de nombres. Trois symptomes mesures sur le cas minimal :

  1. le quadrillage disparait entierement (ses traits sont manges par le
     filtre « familles de lignes identiques ») : il ne reste que les taches
     coloriees, sans reference pour compter — d'ou les aires inventees ;
  2. un tableau de texte ressort ligne apres ligne, colonne apres colonne,
     dans un ordre qui ne dit plus quelle valeur va avec quelle entree ;
  3. le meme tableau part AUSSI en figure : le modele le voit deux fois,
     une fois juste et une fois faux.

Principe de la reparation : le serveur ne resout rien, il MESURE et il le
dit. Une grille detectee produit trois choses au plus :

  - sa forme      : « quadrillage 8 colonnes x 6 lignes, case 10,0 mm » ;
  - son contenu   : pour un tableau de texte, les lignes reconstituees ;
  - son comptage  : le nombre de cases pleines de chaque zone coloriee,
                    et UNIQUEMENT si ces zones epousent la grille. Une
                    forme oblique, un triangle, une zone a cheval : rien
                    n'est annonce. Un comptage douteux est plus dangereux
                    qu'un comptage absent.

Autonome : n'a besoin que de PyMuPDF et de la bibliotheque standard.

Entree publique :
    analyser_page(page) -> dict
        {"grilles": [...], "zones_grilles": [fitz.Rect, ...], "texte": str}
"""

from __future__ import annotations

import math
import re

import fitz

try:
    import numpy as _np
except Exception:          # numpy absent : les quadrillages en image
    _np = None             # ne sont pas mesures, le reste fonctionne

VERSION = "2.14"

# ── Reglages, tous exprimes en unites lisibles ──────────────────────────
PT_MM = 25.4 / 72.0          # 1 point PDF en millimetres

MIN_LIGNES = 2               # en deca, ce n'est pas une grille
MIN_COLONNES = 2
MIN_CASE_MM = 3.0            # une case plus petite n'est pas manipulable
TAUX_TEXTE_TABLEAU = 0.40    # au-dela, la grille porte du contenu ecrit
# Un quadrillage a compter, c'est beaucoup de petites cases. Trois figures
# alignees sur une page en forment un aussi, aux yeux d'un detecteur de
# tableaux : les piscines du cas minimal ressortaient en « quadrillage
# 2 x 2, case 75 x 40 mm ». Un quadrillage annonce a tort est pire que pas
# de quadrillage du tout — d'ou ces deux bornes.
CASE_MAX_MM = 30.0           # au-dela, ce n'est pas une case a compter
CASES_MIN_QUADRILLAGE = 9    # moins de 9 cases : mise en page, pas grille
TOLERANCE_CASE = 0.15        # 15 % d'une case : marge d'alignement admise
COUVERTURE_MIN = 0.55        # part d'une case a couvrir pour la dire pleine
ECART_REGULARITE = 0.08      # 8 % d'ecart max entre cases d'un quadrillage


# ── Outils ──────────────────────────────────────────────────────────────

def _mm(valeur_pt: float) -> float:
    return round(valeur_pt * PT_MM, 1)


def _cellules(table) -> list:
    """Rectangles des cellules d'une table PyMuPDF, en liste plate."""
    plates = []
    for cellule in getattr(table, "cells", []) or []:
        if cellule is None:
            continue
        plates.append(fitz.Rect(cellule))
    return plates


def _pas_moyen(bornes: list[float]) -> tuple[float, bool]:
    """Pas moyen d'une suite de bornes, et regularite de cette suite."""
    if len(bornes) < 2:
        return 0.0, False
    ecarts = [b - a for a, b in zip(bornes, bornes[1:])]
    moyen = sum(ecarts) / len(ecarts)
    if moyen <= 0:
        return 0.0, False
    regulier = all(abs(e - moyen) <= ECART_REGULARITE * moyen for e in ecarts)
    return moyen, regulier


def _bornes(table) -> tuple[list[float], list[float]]:
    """Bornes verticales et horizontales deduites des cellules."""
    xs, ys = set(), set()
    for c in _cellules(table):
        xs.add(round(c.x0, 1)); xs.add(round(c.x1, 1))
        ys.add(round(c.y0, 1)); ys.add(round(c.y1, 1))
    return sorted(xs), sorted(ys)


# ── Fonds de cellules et zones coloriees ────────────────────────────────

def _fonds_colores(page) -> list[dict]:
    """
    Tracas remplis d'une couleur, hors blanc pur et hors fonds de page.
    Chacun revient avec son rectangle et sa couleur arrondie, qui sert de
    cle pour regrouper les cases d'une meme figure.
    """
    aire_page = page.rect.width * page.rect.height
    trouves = []
    for dessin in page.get_drawings():
        remplissage = dessin.get("fill")
        if remplissage is None:
            continue
        rect = dessin.get("rect")
        if rect is None or rect.is_empty:
            continue
        if rect.width * rect.height > 0.85 * aire_page:
            continue                       # fond de page
        couleur = tuple(round(float(v), 2) for v in remplissage)
        if all(v >= 0.97 for v in couleur):
            continue                       # blanc : ce n'est pas un coloriage
        trouves.append({"rect": fitz.Rect(rect), "couleur": couleur,
                        "orthogonal": _a_bords_droits(dessin)})
    return trouves


def _a_bords_droits(dessin) -> bool:
    """
    Vrai si le contour du trace ne comporte que des bords horizontaux ou
    verticaux. Un triangle, un disque, une figure oblique repondent non —
    et leur cadre englobant, lui, tomberait parfaitement sur la grille.
    C'est ce piege qui produisait des aires fausses : le cadre d'un
    triangle 4 x 3 « mesure » 12 cases alors que la figure en couvre 6.
    """
    for item in dessin.get("items", []):
        genre = item[0]
        if genre == "re":
            continue
        if genre == "l":
            p1, p2 = item[1], item[2]
            if abs(p1.x - p2.x) <= 0.5 or abs(p1.y - p2.y) <= 0.5:
                continue
            return False                   # segment oblique
        return False                       # courbe, quad, tout le reste
    return True


def _etiquette_voisine(page, rect: fitz.Rect) -> str | None:
    """
    Lettre isolee (A, B, C…) posee dans une zone coloriee : c'est ainsi que
    les sources nomment leurs figures. Une lettre seule, rien d'autre.
    """
    candidates = []
    for mot in page.get_text("words"):
        texte = mot[4].strip()
        if len(texte) != 1 or not texte.isalpha():
            continue
        centre = fitz.Point((mot[0] + mot[2]) / 2, (mot[1] + mot[3]) / 2)
        if rect.contains(centre):
            candidates.append(texte.upper())
    if len(set(candidates)) == 1:
        return candidates[0]
    return None


def _compter_zones(page, grille: dict, fonds: list[dict]) -> tuple[list[dict], str | None]:
    """
    Compte les cases pleines de chaque zone coloriee posee sur la grille.

    Retourne (zones, refus). Si `refus` n'est pas None, aucun comptage
    n'est transmis et la raison est nommee : c'est la regle du projet,
    une mesure incertaine ne sort pas en silence.
    """
    cadre = grille["rect"]
    largeur_case = grille["case_l_pt"]
    hauteur_case = grille["case_h_pt"]
    if largeur_case <= 0 or hauteur_case <= 0:
        return [], "cases irregulieres"

    dedans = [f for f in fonds
              if not (f["rect"] & cadre).is_empty
              and (f["rect"] & cadre).get_area() >= 0.5 * f["rect"].get_area()]
    if not dedans:
        return [], None

    tol_l = TOLERANCE_CASE * largeur_case
    tol_h = TOLERANCE_CASE * hauteur_case

    cases_pleines: dict[tuple[int, int], tuple] = {}   # (col, lig) -> couleur
    for fond in dedans:
        rect = fond["rect"]
        if not fond["orthogonal"]:
            return [], "figure a bords obliques ou courbes"
        # Un fond doit epouser la grille : ses quatre bords tombent sur une
        # bordure, a la tolerance pres. Sinon, la figure est oblique ou a
        # cheval, et le comptage n'a plus de sens.
        for valeur, origine, pas, tol in (
                (rect.x0, cadre.x0, largeur_case, tol_l),
                (rect.x1, cadre.x0, largeur_case, tol_l),
                (rect.y0, cadre.y0, hauteur_case, tol_h),
                (rect.y1, cadre.y0, hauteur_case, tol_h)):
            reste = abs((valeur - origine) / pas - round((valeur - origine) / pas))
            if reste * pas > tol:
                return [], "figure non alignee sur le quadrillage"
        col_deb = int(math.floor((rect.x0 - cadre.x0) / largeur_case + 0.5))
        col_fin = int(math.floor((rect.x1 - cadre.x0) / largeur_case + 0.5))
        lig_deb = int(math.floor((rect.y0 - cadre.y0) / hauteur_case + 0.5))
        lig_fin = int(math.floor((rect.y1 - cadre.y0) / hauteur_case + 0.5))
        for col in range(col_deb, col_fin):
            for lig in range(lig_deb, lig_fin):
                if 0 <= col < grille["colonnes"] and 0 <= lig < grille["lignes"]:
                    cases_pleines[(col, lig)] = fond["couleur"]

    # Une figure, ce sont des cases qui se touchent par un cote ET qui
    # portent la meme couleur. Deux figures grises posees cote a cote sans
    # se toucher restent deux figures ; un L reste une seule figure.
    zones = []
    restantes = dict(cases_pleines)
    while restantes:
        depart, couleur = next(iter(restantes.items()))
        cases, a_voir = set(), [depart]
        while a_voir:
            case = a_voir.pop()
            if case in cases or restantes.get(case) != couleur:
                continue
            cases.add(case)
            restantes.pop(case, None)
            col, lig = case
            a_voir.extend([(col + 1, lig), (col - 1, lig),
                           (col, lig + 1), (col, lig - 1)])
        if not cases:
            continue
        enveloppe = fitz.Rect(
            cadre.x0 + min(c for c, _ in cases) * largeur_case,
            cadre.y0 + min(l for _, l in cases) * hauteur_case,
            cadre.x0 + (max(c for c, _ in cases) + 1) * largeur_case,
            cadre.y0 + (max(l for _, l in cases) + 1) * hauteur_case)
        etendue = ((max(c for c, _ in cases) - min(c for c, _ in cases) + 1)
                   * (max(l for _, l in cases) - min(l for _, l in cases) + 1))
        zones.append({
            "etiquette": _etiquette_voisine(page, enveloppe),
            "cases": len(cases),
            "composite": len(cases) != etendue,
            "rect": enveloppe,
        })

    zones.sort(key=lambda z: (round(z["rect"].y0), z["rect"].x0))
    return zones, None


# ── Quadrillages sans texte : detection par les traits ──────────────────
#
# Le detecteur de tableaux de PyMuPDF s'appuie sur le contenu ecrit : une
# grille vide, celle qu'on donne justement a compter, lui reste invisible.
# On la retrouve donc par sa geometrie : des traits paralleles reguliers
# qui se croisent. C'est le cas d'usage central du chantier.

MIN_TRAITS = 3               # 3 traits = 2 cases : le minimum d'une grille
LONGUEUR_MIN_PT = 20.0


def _segments(page) -> tuple[list, list]:
    """Segments horizontaux et verticaux de la page, traits fins compris."""
    horizontaux, verticaux = [], []
    for dessin in page.get_drawings():
        for item in dessin.get("items", []):
            if item[0] == "l":
                p1, p2 = item[1], item[2]
                x0, y0, x1, y1 = p1.x, p1.y, p2.x, p2.y
            elif item[0] == "re":
                r = item[1]
                if min(r.width, r.height) > 1.5:
                    continue          # rectangle plein, pas un trait
                x0, y0, x1, y1 = r.x0, r.y0, r.x1, r.y1
            else:
                continue
            if abs(y1 - y0) <= 1.0 and abs(x1 - x0) >= LONGUEUR_MIN_PT:
                horizontaux.append((min(x0, x1), max(x0, x1), (y0 + y1) / 2))
            elif abs(x1 - x0) <= 1.0 and abs(y1 - y0) >= LONGUEUR_MIN_PT:
                verticaux.append((min(y0, y1), max(y0, y1), (x0 + x1) / 2))
    return horizontaux, verticaux


def _suite_reguliere(valeurs: list[float]) -> list[float] | None:
    """
    Plus longue suite a pas constant dans une liste de coordonnees.
    Retourne None si elle compte moins de MIN_TRAITS elements.
    """
    uniques = []
    for v in sorted(valeurs):
        if not uniques or abs(v - uniques[-1]) > 1.5:
            uniques.append(v)
    if len(uniques) < MIN_TRAITS:
        return None
    meilleure = None
    for depart in range(len(uniques) - MIN_TRAITS + 1):
        for suivant in range(depart + 1, len(uniques)):
            pas = uniques[suivant] - uniques[depart]
            if pas <= 0:
                continue
            suite = [uniques[depart], uniques[suivant]]
            for v in uniques[suivant + 1:]:
                if abs(v - (suite[-1] + pas)) <= ECART_REGULARITE * pas:
                    suite.append(v)
            if len(suite) >= MIN_TRAITS and (meilleure is None
                                             or len(suite) > len(meilleure)):
                meilleure = suite
    return meilleure


def _quadrillages_geometriques(page) -> list[dict]:
    """Grilles reconnues a leurs seuls traits, sans aucun texte requis."""
    horizontaux, verticaux = _segments(page)
    if len(horizontaux) < MIN_TRAITS or len(verticaux) < MIN_TRAITS:
        return []

    ys = _suite_reguliere([h[2] for h in horizontaux])
    xs = _suite_reguliere([v[2] for v in verticaux])
    if not ys or not xs:
        return []

    cadre = fitz.Rect(xs[0], ys[0], xs[-1], ys[-1])
    if cadre.is_empty:
        return []

    # Les traits doivent vraiment former une grille : chaque horizontale
    # retenue traverse la largeur du cadre, chaque verticale sa hauteur.
    def _traverse(debut, fin, borne_a, borne_b):
        return debut <= borne_a + 2 and fin >= borne_b - 2

    horiz_ok = sum(1 for h in horizontaux
                   if any(abs(h[2] - y) <= 1.5 for y in ys)
                   and _traverse(h[0], h[1], cadre.x0, cadre.x1))
    vert_ok = sum(1 for v in verticaux
                  if any(abs(v[2] - x) <= 1.5 for x in xs)
                  and _traverse(v[0], v[1], cadre.y0, cadre.y1))
    if horiz_ok < MIN_TRAITS or vert_ok < MIN_TRAITS:
        return []

    case_l = (xs[-1] - xs[0]) / (len(xs) - 1)
    case_h = (ys[-1] - ys[0]) / (len(ys) - 1)
    if _mm(case_l) < MIN_CASE_MM or _mm(case_h) < MIN_CASE_MM:
        return []

    return [{
        "rect": cadre,
        "lignes": len(ys) - 1,
        "colonnes": len(xs) - 1,
        "case_l_pt": case_l,
        "case_h_pt": case_h,
        "case_l_mm": _mm(case_l),
        "case_h_mm": _mm(case_h),
        "l_mm": _mm(cadre.width),
        "h_mm": _mm(cadre.height),
        "regulier": True,
        "taux_texte": 0.0,
        "contenu": [],
        "nature": "quadrillage",
        "zones": [],
        "comptage_refuse": None,
    }]


# ── Quadrillages arrives en IMAGE ───────────────────────────────────────
#
# Le detecteur geometrique lit des traits ; il ne voit rien quand le
# quadrillage est un dessin importe, ce qui est le cas le plus frequent
# dans les evaluations reelles (l'exercice « reproduis ce dessin » de
# l'evaluation de geometrie 5P, par exemple). On le retrouve alors dans
# les pixels : des rangees de points sombres alignes et regulierement
# espaces. On mesure sa forme, jamais son contenu — un dessin pose sur un
# quadrillage n'est pas un coloriage de cases, et le compter serait
# exactement le genre de mesure douteuse que ce module refuse.

PART_MIN_RASTER = 0.45       # part de la traversee a noircir pour etre un trait
SEUIL_SOMBRE = 160
ECART_TOLERE_RASTER = 0.12
COUPURE_BLOC = 2.5


def _pics(projection, seuil: float) -> list[int]:
    """Positions des rangees denses, une seule par trait meme s'il est epais."""
    dedans = projection >= seuil
    pics, debut = [], None
    for i, actif in enumerate(dedans):
        if actif and debut is None:
            debut = i
        elif not actif and debut is not None:
            pics.append((debut + i - 1) // 2)
            debut = None
    if debut is not None:
        pics.append((debut + len(dedans) - 1) // 2)
    return pics


def _suites_pics(pics: list[int]) -> list[list[int]]:
    """Decoupe une liste de pics en suites a pas constant."""
    if len(pics) < MIN_TRAITS:
        return []
    ecarts = [b - a for a, b in zip(pics, pics[1:])]
    if not ecarts:
        return []
    pas = float(_np.median(ecarts))
    if pas <= 1:
        return []
    suites, courante = [], [pics[0]]
    for precedent, actuel in zip(pics, pics[1:]):
        ecart = actuel - precedent
        if abs(ecart - pas) <= ECART_TOLERE_RASTER * pas:
            courante.append(actuel)
        elif ecart > COUPURE_BLOC * pas:
            suites.append(courante)
            courante = [actuel]
        else:
            courante.append(actuel)      # trait manquant ou double : on suit
    suites.append(courante)
    return [s for s in suites if len(s) >= MIN_TRAITS]


def _quadrillages_raster(page) -> list[dict]:
    """Quadrillages caches dans les images de la page."""
    if _np is None:
        return []
    trouves = []
    doc = page.parent
    for info in page.get_images(full=True):
        try:
            rects = page.get_image_rects(info[0])
            if not rects:
                continue
            cadre = fitz.Rect(rects[0])
            if cadre.width < 40 or cadre.height < 40:
                continue
            pix = fitz.Pixmap(doc, info[0])
            donnees = _np.frombuffer(pix.samples, dtype=_np.uint8)
            donnees = donnees.reshape(pix.height, pix.width, pix.n)
            gris = (donnees[:, :, :3].mean(axis=2).astype(_np.uint8)
                    if pix.n >= 3 else donnees[:, :, 0])
            sombre = (gris < SEUIL_SOMBRE).astype(_np.int32)

            xs_suites = _suites_pics(
                _pics(sombre.sum(axis=0), PART_MIN_RASTER * pix.height))
            ys_suites = _suites_pics(
                _pics(sombre.sum(axis=1), PART_MIN_RASTER * pix.width))
            if not xs_suites or not ys_suites:
                continue
            xs = max(xs_suites, key=len)
            ys = max(ys_suites, key=len)

            case_l = (xs[-1] - xs[0]) / (len(xs) - 1) * cadre.width / pix.width
            case_h = (ys[-1] - ys[0]) / (len(ys) - 1) * cadre.height / pix.height
            if _mm(case_l) < MIN_CASE_MM or _mm(case_h) < MIN_CASE_MM:
                continue

            trouves.append({
                "rect": cadre,
                "lignes": len(ys) - 1,
                "colonnes": len(xs) - 1,
                "case_l_pt": case_l,
                "case_h_pt": case_h,
                "case_l_mm": _mm(case_l),
                "case_h_mm": _mm(case_h),
                "l_mm": _mm(cadre.width),
                "h_mm": _mm(cadre.height),
                "regulier": True,
                "taux_texte": 0.0,
                "contenu": [],
                "nature": "quadrillage_image",
                "zones": [],
                "comptage_refuse": None,
            })
        except Exception:
            continue
    return trouves


MOTS_COURANTS_MAX = 3        # v2.11 — seuil au milieu du plateau (vrais 0 · faux 8)
_RE_MOT = re.compile("[A-Za-z\u00e0\u00e2\u00e7\u00e9\u00e8\u00ea\u00eb\u00ee\u00ef\u00f4\u00f6\u00fb\u00fc]{2,}")


def _texte_courant(page, rect) -> int:
    """v2.11 — un quadrillage qui porte du TEXTE COURANT n'en est pas un.
    Condition : des mots d'au moins deux lettres dans le cadre. Effet : le
    compte ; le refus se joue au seuil MOTS_COURANTS_MAX, aux deux portes
    d'analyser_page. (voir JOURNAL BACKEND, entree module grilles 2.11)"""
    n = 0
    for w in page.get_text("words"):
        if fitz.Rect(w[:4]).intersects(rect) and _RE_MOT.search(w[4]):
            n += 1
    return n


BANDE_COLONNES_MIN = 6       # v2.12 — en deca, c'est un en-tete, pas un parcours
_RE_LETTRE = re.compile("[A-Za-z\u00e0\u00e2\u00e7\u00e9\u00e8\u00ea\u00eb\u00ee\u00ef\u00f4\u00f6\u00fb\u00fc]")


def _est_bande(table) -> bool:
    """v2.12 — UNE BANDE D'UNE SEULE LIGNE EST UNE FIGURE (piste de jeu, frise,
    bande numerique). Condition : une ligne, au moins BANDE_COLONNES_MIN
    colonnes, et AUCUNE lettre dans ses cases (un chiffre est une case de
    piste, un mot est un intitule de colonne — meme doctrine que le critere
    de mots de la 2.11). Effet : elle franchit la porte MIN_LIGNES et part en
    QUADRILLAGE, donc en figure d'un seul tenant, avec son pion et sa case
    coloriee. (voir JOURNAL BACKEND, entree module grilles 2.12)"""
    if table.row_count != 1 or table.col_count < BANDE_COLONNES_MIN:
        return False
    try:
        cases = table.extract()[0]
    except Exception:
        return False
    return not any(_RE_LETTRE.search((c or "")) for c in cases)


def _quadrillage_plausible(grille: dict) -> bool:
    """
    Un quadrillage a compter porte beaucoup de petites cases. Quelques
    grandes cases, ce sont des figures alignees ou une mise en page.
    """
    if (grille["case_l_mm"] > CASE_MAX_MM
            or grille["case_h_mm"] > CASE_MAX_MM):
        return False
    return grille["colonnes"] * grille["lignes"] >= CASES_MIN_QUADRILLAGE


# ── Detection des grilles ───────────────────────────────────────────────

def analyser_page(page) -> dict:
    """
    Analyse les grilles d'une page.

    Retourne :
      grilles       liste decrite ci-dessous
      zones_grilles rectangles a NE PAS transmettre en figure decoupee
                    (les tableaux de texte ; le quadrillage de mesure, lui,
                    reste une figure — mais d'un seul tenant)
      texte         le texte de la page, tableaux reconstitues en lignes
    """
    grilles = []
    try:
        trouvees = page.find_tables(strategy="lines").tables
    except Exception:
        trouvees = []

    fonds = _fonds_colores(page)

    for table in trouvees:
        lignes, colonnes = table.row_count, table.col_count
        bande = _est_bande(table)          # v2.12
        if (lignes < MIN_LIGNES and not bande) or colonnes < MIN_COLONNES:
            continue
        cadre = fitz.Rect(table.bbox)
        if cadre.is_empty:
            continue

        xs, ys = _bornes(table)
        pas_l, regulier_l = _pas_moyen(xs)
        pas_h, regulier_h = _pas_moyen(ys)
        case_l = pas_l if pas_l > 0 else cadre.width / max(colonnes, 1)
        case_h = pas_h if pas_h > 0 else cadre.height / max(lignes, 1)
        if _mm(case_l) < MIN_CASE_MM or _mm(case_h) < MIN_CASE_MM:
            continue

        contenu = table.extract()
        total = sum(len(r) for r in contenu) or 1
        remplies = sum(1 for r in contenu for c in r if (c or "").strip())
        taux = remplies / total

        grille = {
            "rect": cadre,
            "lignes": lignes,
            "colonnes": colonnes,
            "case_l_pt": case_l,
            "case_h_pt": case_h,
            "case_l_mm": _mm(case_l),
            "case_h_mm": _mm(case_h),
            "l_mm": _mm(cadre.width),
            "h_mm": _mm(cadre.height),
            "regulier": bool(regulier_l and regulier_h),
            "taux_texte": round(taux, 2),
            "contenu": contenu,
            "nature": ("quadrillage" if bande            # v2.12
                       else "tableau" if taux >= TAUX_TEXTE_TABLEAU
                       else "quadrillage"),
            "zones": [],
            "comptage_refuse": None,
        }

        if (grille["nature"] == "quadrillage" and not bande
                and not _quadrillage_plausible(grille)):
            continue
        if bande and (grille["case_l_mm"] > CASE_MAX_MM
                      or grille["case_h_mm"] > CASE_MAX_MM):
            continue          # v2.12 — des cases geantes ne font pas une piste
        if (grille["nature"].startswith("quadrillage")
                and _texte_courant(page, cadre) >= MOTS_COURANTS_MAX):
            continue          # v2.11 — cadre decoratif a texte, pas une grille

        grilles.append(grille)

    # Quadrillages vides : invisibles au detecteur de tableaux, retrouves
    # par leurs traits. On n'ajoute que ceux qui ne recouvrent pas une
    # grille deja connue.
    for candidat in (_quadrillages_geometriques(page)
                     + _quadrillages_raster(page)):
        double = any(not (candidat["rect"] & g["rect"]).is_empty
                     and (candidat["rect"] & g["rect"]).get_area()
                     >= 0.5 * min(candidat["rect"].get_area(),
                                  g["rect"].get_area())
                     for g in grilles)
        if (not double and _quadrillage_plausible(candidat)
                and _texte_courant(page, candidat["rect"]) < MOTS_COURANTS_MAX):
            grilles.append(candidat)   # v2.11 — meme porte que la branche des tables

    for grille in grilles:
        if grille["nature"] == "quadrillage":
            zones, refus = _compter_zones(page, grille, fonds)
            grille["zones"] = zones
            grille["comptage_refuse"] = refus

    zones_a_masquer = [g["rect"] for g in grilles if g["nature"] == "tableau"]
    return {
        "grilles": grilles,
        "zones_grilles": zones_a_masquer,
        "texte": texte_avec_grilles(page, grilles),   # sans blocs ajoutes
    }


# ── Restitution ─────────────────────────────────────────────────────────

def decrire(grille: dict) -> str:
    """Une ligne de description factuelle, destinee au modele."""
    if grille["nature"] == "tableau":
        return (f"[tableau {grille['lignes']} lignes x "
                f"{grille['colonnes']} colonnes]")
    if grille["nature"] == "quadrillage_image":
        return (f"[quadrillage en image {grille['colonnes']} colonnes x "
                f"{grille['lignes']} lignes, "
                f"{grille['colonnes'] * grille['lignes']} cases, "
                f"case {grille['case_l_mm']} x {grille['case_h_mm']} mm]")
    forme = (f"[quadrillage {grille['colonnes']} colonnes x "
             f"{grille['lignes']} lignes, "
             f"{grille['colonnes'] * grille['lignes']} cases, "
             f"case {grille['case_l_mm']} x {grille['case_h_mm']} mm]")
    if grille["comptage_refuse"]:
        return forme + f" [comptage non transmis : {grille['comptage_refuse']}]"
    morceaux = []
    for zone in grille["zones"]:
        nom = f"figure {zone['etiquette']}" if zone["etiquette"] else "zone coloriee"
        mot = "cases" if zone["cases"] > 1 else "case"
        forme_zone = ", forme composite" if zone["composite"] else ""
        morceaux.append(f"{nom} : {zone['cases']} {mot} pleines{forme_zone}")
    if morceaux:
        # L'interdiction voyage AVEC la mesure, au point exact ou elle
        # apparait : c'est la reponse de l'exercice, elle sert a la fiche
        # enseignant et a rien d'autre. Une regle posee loin, dans un bloc
        # de prompt conditionnel, ne serait pas toujours envoyee.
        forme += (" [mesure du document, RESERVEE A LA FICHE ENSEIGNANT — "
                  + " ; ".join(morceaux)
                  + ". Ces nombres sont la reponse de l'exercice : ne les "
                    "reporte JAMAIS sur la feuille eleve, ni comme reponse, "
                    "ni dans un exemple, ni dans une consigne, ni dans une "
                    "zone pre-remplie.]")
    return forme


def _tableau_en_lignes(grille: dict) -> str:
    lignes = []
    for rangee in grille["contenu"]:
        cellules = [(c or "").replace("\n", " ").strip() for c in rangee]
        lignes.append(" | ".join(cellules))
    return "\n".join(lignes)


BANDE_PT = 8.0   # hauteur de bande du tri de lecture, inchangee depuis v2.10


def cle_lecture(y0, x0):
    """LA cle d'ordre de lecture du projet, et la seule (v2.14, 17.08.2026).

    Elle etait ecrite en dur dans le tri des pages a grilles ; l'ancre de
    lecture du serveur, elle, comparait des altitudes brutes. Deux ordres
    differents pour une meme question : mesure du 17.08, 9 marqueurs poses
    entre le numero d'un exercice et sa consigne, parce que le numero est un
    point plus bas que sa consigne. Une seule cle desormais, appelee des deux
    endroits. (voir JOURNAL BACKEND v2.37)"""
    return (round(y0 / BANDE_PT), x0)


def texte_avec_grilles(page, grilles: list[dict], blocs_sup=None,
                       rendre_blocs=False):
    """
    Texte de la page ou chaque grille est restituee a sa place, dans le
    bon ordre de lecture, en remplacement de la file de mots qui sortait
    jusqu'ici. Le texte contenu dans une grille n'apparait qu'une fois.

    v2.13 - rendre_blocs=True rend AUSSI la liste des blocs assembles, dans
    l'ordre du texte, chacun avec sa position : (y0, x0, texte). Le texte rendu
    est le MEME, caractere pour caractere. Quand les blocs ne peuvent pas etre
    rendus fidelement, la liste vaut None : l'abstention se dit, elle ne se
    devine pas.
    """
    blocs_sup = list(blocs_sup or [])
    # v2.14 (17.08.2026) — LE RETOUR ANTICIPE EST SUPPRIME. Il rendait
    # `page.get_text("text")` : les blocs dans l'ordre INTERNE du PDF, celui du
    # dessin, pas celui de la lecture. Mesure du 17.08 sur 18 documents et 46
    # pages : 11 documents et 26 pages assemblaient un texte ou jusqu'a un bloc
    # sur trois REMONTE vers le haut de la page. Le modele lisait un titre de
    # haut de page apres une consigne du bas. Le tri qui suit existait deja et
    # tournait sur les pages a grilles, ou l'ordre est parfait : il est
    # generalise, il n'est pas invente. (voir JOURNAL BACKEND v2.37)

    blocs = []
    for bloc in page.get_text("blocks"):
        rect = fitz.Rect(bloc[:4])
        texte = (bloc[4] or "").strip()
        if not texte:
            continue
        dans_grille = any(
            not (rect & g["rect"]).is_empty
            and (rect & g["rect"]).get_area() >= 0.5 * rect.get_area()
            for g in grilles)
        if dans_grille:
            continue
        blocs.append((rect.y0, rect.x0, texte, rect.y1, False))

    for grille in grilles:
        entete = decrire(grille)
        corps = _tableau_en_lignes(grille) if grille["nature"] == "tableau" else ""
        blocs.append((grille["rect"].y0, grille["rect"].x0,
                      entete + ("\n" + corps if corps else ""),
                      grille["rect"].y1, True))

    # Contenus fournis par d'autres modules (formes composites), places a
    # leur position de lecture comme le reste.
    blocs.extend(blocs_sup)

    blocs.sort(key=lambda b: cle_lecture(b[0], b[1]))
    texte = "\n".join(b[2] for b in blocs)
    if rendre_blocs:
        return texte, list(blocs)
    return texte
