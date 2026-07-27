#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FORMES COMPOSITES — module serveur, version 2.10
================================================

VERSION 2.10 — 27 juillet 2026.
Fichier livre sous le nom `formes_v2_10.py` ; **a renommer `formes.py`** au
deploiement, a cote d'`app.py`. La version est lisible sur `/health`.

Repare le defaut de la piscine B, ouvert depuis l'essai 55 et longtemps
impute a tort a un modele.

CE QUI SE PASSE AUJOURD'HUI. Une figure faite de deux rectangles — une
piscine en L, une terrasse, une piece a plusieurs pans — est extraite par
son ENCOMBREMENT : le plus petit rectangle qui la contient. Sur le cas
minimal, la piscine B, 40 x 80 plus 110 x 20, soit 5 400 mm2, ressort en
150 x 80, soit 12 000 mm2. Rien ne dit au modele que la figure est creuse.
Il mesure ce qu'on lui donne, et la plus grande piscine change de nom :
l'eleve repond faux, sur les deux modeles, par construction.

CE QUE FAIT LA REPARATION. Le meme principe que pour les grilles : le
serveur ne resout rien, il mesure et il le dit.

  - la FORME est decrite : « figure composite : 2 rectangles, 40 x 80 mm et
    110 x 20 mm ; encombrement 150 x 80 mm ». Ce n'est pas la reponse de
    l'exercice, c'est la description de la figure — sans elle, le modele
    croit voir un bloc plein ;
  - l'AIRE reelle, elle, EST la reponse. Elle est donc transmise sous la
    meme reserve que les comptages de cases : fiche enseignant seulement,
    jamais la feuille eleve (arbitrage de Catherine du 27.07) ;
  - au moindre doute, rien n'est calcule. Un pan oblique, un arrondi, deux
    figures voisines qu'on ne peut pas dire reunies : la forme est signalee
    comme non decomposable et la raison est transmise. Un chiffre faux est
    pire qu'un chiffre absent.

Autonome : PyMuPDF et la bibliotheque standard.

Entree publique :
    analyser_formes(page) -> list[dict]
"""

from __future__ import annotations

import fitz

VERSION = "2.10"

PT_MM = 25.4 / 72.0

CONTACT_PT = 2.0        # au-dela, deux rectangles ne se touchent plus
VOISINAGE_PT = 10.0     # marge du regroupement du pipeline : zone de doute
MIN_COTE_MM = 5.0       # plus petit qu'un timbre : pas une figure a mesurer
MIN_MEMBRES = 2         # une seule piece, ce n'est pas un composite
ECART_TAILLE = 0.05     # 5 % : deux cases « de meme taille »


def _mm(v: float) -> float:
    return round(v * PT_MM, 1)


def _est_rectangle(dessin) -> bool:
    """
    Vrai si le trace est un rectangle a bords droits. Un « re », ou quatre
    segments tous horizontaux ou verticaux. Tout le reste — oblique, courbe,
    arrondi — repond non, et interdira le calcul d'aire.
    """
    items = dessin.get("items", [])
    if not items:
        return False
    for item in items:
        genre = item[0]
        if genre == "re":
            continue
        if genre == "l":
            p1, p2 = item[1], item[2]
            if abs(p1.x - p2.x) <= 0.5 or abs(p1.y - p2.y) <= 0.5:
                continue
            return False
        return False
    return True


def _rectangles(page) -> tuple[list[fitz.Rect], bool]:
    """
    Rectangles traces de la page, et un drapeau disant si des formes non
    rectangulaires ont ete rencontrees au passage.
    """
    aire_page = page.rect.width * page.rect.height
    trouves, autre_forme = [], False
    for dessin in page.get_drawings():
        rect = dessin.get("rect")
        if rect is None or rect.is_empty:
            continue
        if rect.width * rect.height > 0.85 * aire_page:
            continue                       # fond de page
        if _mm(rect.width) < MIN_COTE_MM or _mm(rect.height) < MIN_COTE_MM:
            continue                       # trait, tiret, micro-trace
        if _est_rectangle(dessin):
            trouves.append(fitz.Rect(rect))
        else:
            autre_forme = True
    return trouves, autre_forme


def _distance(a: fitz.Rect, b: fitz.Rect) -> float:
    """Ecart entre deux rectangles ; 0 s'ils se touchent ou se recouvrent."""
    dx = max(0.0, max(a.x0, b.x0) - min(a.x1, b.x1))
    dy = max(0.0, max(a.y0, b.y0) - min(a.y1, b.y1))
    return max(dx, dy)


def _grouper(rects: list[fitz.Rect], seuil: float) -> list[list[fitz.Rect]]:
    """Groupes de rectangles relies de proche en proche."""
    restants = list(rects)
    groupes = []
    while restants:
        groupe = [restants.pop()]
        change = True
        while change:
            change = False
            for candidat in list(restants):
                if any(_distance(candidat, membre) <= seuil for membre in groupe):
                    groupe.append(candidat)
                    restants.remove(candidat)
                    change = True
        groupes.append(groupe)
    return groupes


def _est_serie_de_cases(groupe: list[fitz.Rect]) -> bool:
    """
    Une rangee de cases a remplir, toutes de meme taille et alignees, n'est
    pas une figure composite : c'est un dispositif de reponse. La decrire
    comme une piscine en escalier serait un contresens.
    """
    if len(groupe) < 2:
        return False
    largeurs = [r.width for r in groupe]
    hauteurs = [r.height for r in groupe]
    meme_taille = (
        max(largeurs) - min(largeurs) <= ECART_TAILLE * max(largeurs)
        and max(hauteurs) - min(hauteurs) <= ECART_TAILLE * max(hauteurs))
    if not meme_taille:
        return False
    en_rangee = max(r.y0 for r in groupe) - min(r.y0 for r in groupe) <= 2.0
    en_colonne = max(r.x0 for r in groupe) - min(r.x0 for r in groupe) <= 2.0
    if en_rangee or en_colonne:
        return True

    # Pavage : des cases identiques posees sur une trame reguliere, comme
    # un coloriage case par case sur un quadrillage. Les additionner en
    # « figure composite » ferait double emploi avec le comptage de cases,
    # qui est deja fait, mieux, par le module des grilles.
    pas_l, pas_h = max(largeurs), max(hauteurs)
    origine_x = min(r.x0 for r in groupe)
    origine_y = min(r.y0 for r in groupe)
    sur_trame = all(
        abs((r.x0 - origine_x) / pas_l - round((r.x0 - origine_x) / pas_l))
        * pas_l <= ECART_TAILLE * pas_l
        and abs((r.y0 - origine_y) / pas_h - round((r.y0 - origine_y) / pas_h))
        * pas_h <= ECART_TAILLE * pas_h
        for r in groupe)
    return sur_trame


def _aire_union_mm2(rects: list[fitz.Rect]) -> float:
    """
    Aire de la reunion, recouvrements comptes une seule fois. Balayage par
    bandes : les bords verticaux decoupent le plan en tranches, et dans
    chaque tranche on additionne les hauteurs occupees.
    """
    if not rects:
        return 0.0
    xs = sorted({v for r in rects for v in (r.x0, r.x1)})
    total = 0.0
    for gauche, droite in zip(xs, xs[1:]):
        largeur = droite - gauche
        if largeur <= 0:
            continue
        segments = sorted((r.y0, r.y1) for r in rects
                          if r.x0 <= gauche and r.x1 >= droite)
        haut = 0.0
        fin_courante = None
        for debut, fin in segments:
            if fin_courante is None or debut > fin_courante:
                haut += fin - debut
                fin_courante = fin
            elif fin > fin_courante:
                haut += fin - fin_courante
                fin_courante = fin
        total += largeur * haut
    return round(total * PT_MM * PT_MM)


def _etiquette(page, cadre: fitz.Rect) -> str | None:
    """Lettre isolee posee dans la figure — c'est ainsi qu'on la nomme."""
    vues = []
    for mot in page.get_text("words"):
        texte = mot[4].strip()
        if len(texte) != 1 or not texte.isalpha():
            continue
        centre = fitz.Point((mot[0] + mot[2]) / 2, (mot[1] + mot[3]) / 2)
        if cadre.contains(centre):
            vues.append(texte.upper())
    return vues[0] if len(set(vues)) == 1 and vues else None


def analyser_formes(page) -> list[dict]:
    """Formes composites de la page, decrites et mesurees quand c'est sur."""
    rects, autre_forme = _rectangles(page)
    if len(rects) < MIN_MEMBRES:
        return []

    resultats = []
    for groupe in _grouper(rects, CONTACT_PT):
        if len(groupe) < MIN_MEMBRES:
            continue
        if _est_serie_de_cases(groupe):
            continue

        cadre = fitz.Rect(groupe[0])
        for r in groupe[1:]:
            cadre |= r

        # Zone de doute : un rectangle qui n'est pas au contact mais qui
        # entre dans la marge de regroupement du pipeline. Figure unique ou
        # deux figures voisines ? La machine ne trancherait qu'au hasard.
        voisin_douteux = any(
            CONTACT_PT < _distance(autre, membre) <= VOISINAGE_PT
            for membre in groupe for autre in rects if autre not in groupe)

        refus = None
        if autre_forme:
            refus = "la page porte des formes non rectangulaires"
        elif voisin_douteux:
            refus = "une figure voisine touche presque celle-ci"

        resultats.append({
            "rect": cadre,
            "etiquette": _etiquette(page, cadre),
            "membres": [{"l_mm": _mm(r.width), "h_mm": _mm(r.height)}
                        for r in sorted(groupe, key=lambda r: (r.y0, r.x0))],
            "encombrement_l_mm": _mm(cadre.width),
            "encombrement_h_mm": _mm(cadre.height),
            "encombrement_mm2": round(_mm(cadre.width) * _mm(cadre.height)),
            "aire_mm2": None if refus else _aire_union_mm2(groupe),
            "aire_refusee": refus,
        })
    return resultats


def decrire(forme: dict) -> str:
    """Une ligne factuelle, destinee au modele."""
    nom = f"figure {forme['etiquette']}" if forme["etiquette"] else "figure"
    pieces = " et ".join(f"{m['l_mm']} x {m['h_mm']} mm" for m in forme["membres"])
    texte = (f"[{nom} composite : {len(forme['membres'])} rectangles, {pieces} ; "
             f"encombrement {forme['encombrement_l_mm']} x "
             f"{forme['encombrement_h_mm']} mm — la figure NE REMPLIT PAS son "
             f"encombrement, ne la traite jamais comme un rectangle plein]")
    if forme["aire_refusee"]:
        return texte + f" [aire non calculee : {forme['aire_refusee']}]"
    return texte + (
        f" [mesure du document, RESERVEE A LA FICHE ENSEIGNANT — aire reelle "
        f"{forme['aire_mm2']} mm2, contre {forme['encombrement_mm2']} mm2 pour "
        f"l'encombrement. C'est la reponse de l'exercice : ne la reporte JAMAIS "
        f"sur la feuille eleve, ni comme reponse, ni dans un exemple, ni dans "
        f"une consigne, ni dans une zone pre-remplie.]")
