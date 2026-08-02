#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Contre-épreuve v2.17 : le VRAI défaut a-t-il disparu, sans dégât nouveau ?"""
import sys, types, importlib.util
for nom in ('pdf2docx','docx','mammoth','docx.shared','docx.enum','docx.enum.text','docx.oxml','docx.oxml.ns'):
    if nom not in sys.modules: sys.modules[nom]=types.ModuleType(nom)
class X:
    def __init__(self,*a,**k):pass
    def __getattr__(self,n):return X()
    def __call__(self,*a,**k):return X()
for nom in ('pdf2docx','docx','mammoth','docx.shared','docx.enum','docx.enum.text','docx.oxml','docx.oxml.ns'):
    for a in ('Converter','Document','Inches','Pt','Cm','Emu','RGBColor','WD_ALIGN_PARAGRAPH','OxmlElement','qn','convert_to_html','WD_BREAK'):
        setattr(sys.modules[nom],a,X)
import fitz
def zones(mod, srcs):
    spec=importlib.util.spec_from_file_location("bk",mod); bk=importlib.util.module_from_spec(spec); spec.loader.exec_module(bk)
    out={}
    for src in srcs:
        bk.TV_ZONES_MESUREES.clear(); bk.parse_pdf(open(src,'rb').read(),src.split('/')[-1])
        out[src]=list(bk.TV_ZONES_MESUREES)
    return out
SRCS=sys.argv[3:]
A=zones(sys.argv[1],SRCS); B=zones(sys.argv[2],SRCS)
def bilan(Z,titre):
    amp=gris=voisin=mots=0; surf=[]
    for src,pages in Z.items():
        doc=fitz.open(src)
        for pno,zs in pages:
            page=doc[pno-1]
            formes=[fitz.Rect(d['rect']) for d in page.get_drawings() if 1<fitz.Rect(d['rect']).width<=75 and 1<fitz.Rect(d['rect']).height<=75]
            mm=[(fitz.Rect(w[:4]),w[4]) for w in page.get_text("words")]
            for z in zs:
                zr=fitz.Rect(z)
                for r in formes:
                    i=r&zr
                    if i.is_empty or i.get_area()<=0 or zr.contains(r): continue
                    d=max(zr.y0-r.y0,r.y1-zr.y1,zr.x0-r.x0,r.x1-zr.x1)
                    if d<1.0: continue
                    p=i.get_area()/r.get_area()
                    if p>=0.50: amp+=1
                    elif p>=0.25: gris+=1
                    else: voisin+=1
    print(f"\n   {titre}")
    print(f"     formes AMPUTÉES (≥50 % dedans, coupées) ... {amp}")
    print(f"     zone grise (25–50 %) ...................... {gris}")
    print(f"     voisines effleurées (<25 %) ............... {voisin}")
    return amp
print("══ CONTRE-ÉPREUVE v2.17 ══")
a=bilan(A,"AVANT (v2.16)"); b=bilan(B,"APRÈS (v2.17)")
# mots avalés : mots présents dans une zone après et pas avant
mots_nouv=0; surf=[]
for src in SRCS:
    doc=fitz.open(src)
    for (pno,za),(pnb,zb) in zip(A[src],B[src]):
        page=doc[pno-1]
        mm=[(fitz.Rect(w[:4]),w[4]) for w in page.get_text("words")]
        for ra,rb in zip(za,zb):
            RA,RB=fitz.Rect(ra),fitz.Rect(rb)
            av={t for r,t in mm if RA.contains(r)}; ap={t for r,t in mm if RB.contains(r)}
            mots_nouv+=len(ap-av)
            if RA.get_area()>0: surf.append(RB.get_area()/RA.get_area()-1)
print(f"\n   COÛT MESURÉ")
print(f"     mots de texte entrés dans une image ....... {mots_nouv}")
print(f"     agrandissement moyen des zones ............ {sum(surf)/len(surf)*100:.1f} %")
print(f"     nombre de zones identique ................. "
      + ("oui" if all(len(x[1])==len(y[1]) for x,y in zip(sum(A.values(),[]),sum(B.values(),[]))) else "NON"))
print(f"\n   VERDICT : {a} formes amputées avant → {b} après.")
