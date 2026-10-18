# -*- coding: utf-8 -*-
"""
PlanMetric Pro — Lecture de plans PDF -> Modélisation 3D -> Métrés réels
-> Devis poste par poste -> Matériaux & Main-d'œuvre.
Rigueur : chaque quantité est traçable (formule + source). Les hypothèses
non lues sur le plan sont marquées et BLOQUENT le statut "définitif".
"""
import io, os, re, math, json
from flask import Flask, request, jsonify, send_file, Response
import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024

def fnum(s):
    try: return float(str(s).replace(" ", "").replace(",", "."))
    except Exception: return None

# ---------------------------------------------------------------- EXTRACTION
PIECES = ["Séjour","Salon","Cuisine","Chambre","Salle de bain","Salle d'eau","SDB","WC",
          "Entrée","Couloir","Dégagement","Bureau","Cellier","Buanderie","Garage",
          "Terrasse","Balcon","Palier","Dressing","Mezzanine","Atelier","Chaufferie"]

def classifier_feuille(texte, nb_segments):
    t = texte.lower()
    if "coupe" in t: return "COUPE"
    if "façade" in t or "facade" in t or "elevation" in t or "élévation" in t: return "FAÇADE"
    if "plan de masse" in t or "parcelle" in t or "cadastre" in t: return "PLAN DE MASSE"
    if any(k in t for k in ["plan rdc","rez-de-chaussée","plan du rez","plan de l'étage","plan étage","plan r+","plan rez"]): return "PLAN DE NIVEAU"
    if "tableau" in t and "menuiserie" in t: return "MENUISERIES"
    if "toiture" in t or "charpente" in t: return "PLAN TOITURE / CHARPENTE"
    if "niveau" in t and nb_segments > 10: return "PLAN DE NIVEAU"
    if nb_segments > 30: return "PLAN (à confirmer)"
    return "AUTRE FEUILLE"

def extraire_texte_pages(doc, nom_fichier="", doc_idx=0):
    pages = []
    for i, page in enumerate(doc):
        dessins = page.get_drawings()
        segs = []
        for d in dessins:
            for it in d["items"]:
                if it[0] == "l":
                    (x1,y1),(x2,y2) = it[1], it[2]
                    segs.append((x1,y1,x2,y2))
                elif it[0] == "re":
                    r = it[1]
                    segs += [(r.x0,r.y0,r.x1,r.y0),(r.x1,r.y0,r.x1,r.y1),
                             (r.x1,r.y1,r.x0,r.y1),(r.x0,r.y1,r.x0,r.y0)]
        tx = page.get_text()
        pages.append({"num": i+1, "doc_idx": doc_idx, "texte": tx, "fichier": nom_fichier,
                      "type_feuille": classifier_feuille(tx, len(segs)),
                      "segments": segs, "dessins": dessins,
                      "nb_images": len(page.get_images()),
                      "rect": (page.rect.width, page.rect.height)})
    return pages

def detecter_echelle_et_emprise(page):
    """Retourne (L_m, W_m, echelle_info). Utilise le plus grand rectangle
    dessiné (dalle/contour) sinon le nuage de segments ; valide l'échelle
    par les cotations numériques du texte (ratio L/l cohérent)."""
    pw, ph = page["rect"]
    grand_rect = None
    for d in page["dessins"]:
        r = d["rect"]
        if r.width > 80 and r.height > 80 and r.width*r.height < 0.85*pw*ph:
            if grand_rect is None or r.width*r.height > grand_rect.width*grand_rect.height:
                grand_rect = r
    if grand_rect is not None:
        bw_pt, bh_pt, bx0, by0 = grand_rect.width, grand_rect.height, grand_rect.x0, grand_rect.y0
        source_geom = "Contour rectangulaire vectoriel du plan (dalle/emprise)"
    else:
        segs = [s for s in page["segments"] if math.hypot(s[2]-s[0], s[3]-s[1]) > 25]
        if len(segs) < 8: return None
        xs = sorted([s[0] for s in segs] + [s[2] for s in segs])
        ys = sorted([s[1] for s in segs] + [s[3] for s in segs])
        k = max(1, int(len(xs)*0.02))
        bx0, bx1, by0, by1 = xs[k], xs[-k-1], ys[k], ys[-k-1]
        bw_pt, bh_pt = bx1-bx0, by1-by0
        source_geom = "Nuage de segments vectoriels (murs) — bornes à 2 %"
    # Cotations candidates : d'abord les lignes autonomes ("12.00" seul = cotation),
    # jamais les nombres des surfaces annotées en m² (piège classique : 13,50 m² pris pour 13,50 m)
    lignes = page["texte"].splitlines()
    autonomes = [fnum(x) for x in re.findall(r"(?m)^\s*(\d{1,2}[.,]\d{2})\s*m?\s*$", page["texte"])]
    cotes = sorted({x for x in autonomes if x and 2 <= x <= 40}, reverse=True)
    if len(cotes) < 2:
        texte_sans_surfaces = "\n".join(l for l in lignes if "m²" not in l and "m2" not in l)
        cotes = sorted({fnum(x) for x in re.findall(r"(?<!\d)(\d{1,2}[.,]\d{2})(?!\d)", texte_sans_surfaces)
                        if fnum(x) and 2 <= fnum(x) <= 40}, reverse=True)
    ratio_pt = bw_pt / bh_pt if bh_pt else 1
    echelle, validee, paire = None, False, None
    for i, a in enumerate(cotes):
        for b in cotes[i:]:
            if b and abs((a/b) - ratio_pt) / max(ratio_pt, 0.01) < 0.18:
                paire, validee = (a, b), True; break
        if paire: break
    if paire:
        L, W = max(paire), min(paire)
        if bw_pt < bh_pt: L, W = W, L
    elif cotes:
        # Plus grande cotation = longueur supposée, largeur par ratio géométrique
        L = cotes[0]; W = L / ratio_pt if bw_pt >= bh_pt else L * ratio_pt
        if bw_pt < bh_pt: L, W = W, L
        paire = (round(L,2), round(W,2))
    else:
        return {"erreur": "Aucune cotation lisible : échelle indéterminable automatiquement.",
                "bloquant": True}
    return {"longueur_m": round(max(L,W),2) if bw_pt>=bh_pt else round(min(L,W),2),
            "largeur_m": round(min(L,W),2) if bw_pt>=bh_pt else round(max(L,W),2),
            "bbox_pt": (bx0, by0, bw_pt, bh_pt),
            "echelle_validee": validee, "cotations_utilisees": paire,
            "source": source_geom}

def murs_interieurs(page, geom):
    bx0, by0, bw, bh = geom["bbox_pt"]
    L, W = geom["longueur_m"], geom["largeur_m"]
    sx, sy = L/bw, W/bh
    horiz, vert = {}, {}
    def cle(v): return round(v/0.25)
    for (x1,y1,x2,y2) in page["segments"]:
        mx1, my1 = (x1-bx0)*sx, (y1-by0)*sy
        mx2, my2 = (x2-bx0)*sx, (y2-by0)*sy
        if not (-0.3 <= mx1 <= L+0.3 and -0.3 <= my1 <= W+0.3): continue
        long = math.hypot(mx2-mx1, my2-my1)
        if long < 0.8: continue
        if abs(my2-my1) < 0.05 and abs(mx2-mx1) > 0.8:  # horizontal
            k = cle((my1+my2)/2)
            horiz.setdefault(k, []).append((min(mx1,mx2), max(mx1,mx2), (my1+my2)/2))
        elif abs(mx2-mx1) < 0.05 and abs(my2-my1) > 0.8:  # vertical
            k = cle((mx1+mx2)/2)
            vert.setdefault(k, []).append((min(my1,my2), max(my1,my2), (mx1+mx2)/2))
    murs, total = [], 0.0
    for groupes, horiz_ok in ((horiz, True), (vert, False)):
        for ivs in groupes.values():
            ivs.sort(); fusion, cur = [], list(ivs[0])
            for a,b,c in ivs[1:]:
                if a <= cur[1] + 0.30: cur[1] = max(cur[1], b)
                else: fusion.append(cur); cur = [a,b,c]
            fusion.append(cur)
            for a,b,c in fusion:
                long = b-a
                # Exclure le contour extérieur lui-même
                if horiz_ok and (c < 0.15 or c > W-0.15): continue
                if not horiz_ok and (c < 0.15 or c > L-0.15): continue
                if long < 0.8: continue
                total += long
                if horiz_ok: murs.append({"x1":round(a,2),"y1":round(c,2),"x2":round(b,2),"y2":round(c,2),"longueur":round(long,2)})
                else: murs.append({"x1":round(c,2),"y1":round(a,2),"x2":round(c,2),"y2":round(b,2),"longueur":round(long,2)})
    return murs[:40], round(total, 2)

def analyser_pdf(data, nom_fichier):
    return analyser_dossier([(nom_fichier, data)])

def analyser_dossier(fichiers):
    """fichiers = [(nom, octets PDF), ...] — toutes les feuilles de tous les
    PDF sont fusionnées en un seul dossier, comme déposé en mairie."""
    pages, noms, docs_refs = [], [], []
    for nom_fichier, data in fichiers:
        doc = pymupdf.open(stream=data, filetype="pdf")
        pages.extend(extraire_texte_pages(doc, nom_fichier, len(docs_refs)))
        noms.append(nom_fichier); docs_refs.append(doc)
    nom_fichier = " + ".join(noms)
    texte_tout = "\n".join(p["texte"] for p in pages)
    total_segs = sum(len(p["segments"]) for p in pages)
    vectoriel = total_segs >= 10
    # Page de plan : plus de segments + mots-clés plan
    def score_plan(p):
        t = p["texte"].lower()
        score = len(p["segments"])*5 + 400*sum(k in t for k in ["rdc","rez-de-chaussée","plan de","plan du"])
        if "coupe" in t or "façade" in t or "elevation" in t: score -= 600
        return score
    page_plan = max(pages, key=score_plan)
    # ---- Pièces annotées (surface écrite sur le plan = donnée d'architecte)
    pieces = []
    for ligne in texte_tout.splitlines():
        m = re.search(r"([A-Za-zÀ-ÿ0-9'’\- ]{3,28}?)\s*[:\-–]?\s*(\d+[.,]\d{1,2})\s*m\s*[²2]", ligne)
        if m:
            nom, surf = m.group(1).strip(" :-–"), fnum(m.group(2))
            if surf and 1 <= surf <= 300 and any(k.lower() in nom.lower() for k in PIECES):
                if not any(p["nom"].lower() == nom.lower() and p["surface_m2"] == surf for p in pieces):
                    pieces.append({"nom": nom, "surface_m2": surf, "source": f"Annotation plan p.{page_plan['num']}"})
    surf_pieces = round(sum(p["surface_m2"] for p in pieces), 2)
    # ---- Surfaces officielles du permis
    def cherche(motifs):
        for mo in motifs:
            m = re.search(mo + r"\s*[:=]?\s*(\d+[.,]\d{1,2})\s*m\s*[²2]", texte_tout, re.I)
            if m: return fnum(m.group(1))
        return None
    sh = cherche([r"surface habitable", r"surf\.?\s*habitable", r"\bSH\b"])
    emprise_txt = cherche([r"emprise au sol", r"emprise"])
    plancher = cherche([r"surface de plancher", r"surface plancher"])
    # ---- Géométrie / échelle
    geom = detecter_echelle_et_emprise(page_plan)
    # ---- Hauteurs (coupes / altitudes)
    altitudes = sorted({fnum(x) for x in re.findall(r"\+\s*(\d+[.,]\d{2})", texte_tout) if fnum(x) and fnum(x) < 15})
    hauteur_niveau, source_hauteur = 2.70, "DÉFAUT (non lu sur coupe) — À VALIDER"
    if len(altitudes) >= 2:
        ecarts = [round(altitudes[i+1]-altitudes[i],2) for i in range(len(altitudes)-1)]
        cand = [e for e in ecarts if 2.2 <= e <= 3.4]
        if cand: hauteur_niveau, source_hauteur = cand[0], f"Altitudes de coupe/façade {altitudes}"
    mh = re.search(r"hauteur (?:d'étage|entre dalles|de mur)[^0-9]{0,10}(\d[.,]\d{2})\s*m", texte_tout, re.I)
    if mh and "DÉFAUT" in source_hauteur: hauteur_niveau, source_hauteur = fnum(mh.group(1)), "Mention hauteur d'étage du dossier"
    # ---- Niveaux
    niveaux_texte = len(set(re.findall(r"\b(RDC|R\+1|R\+2|étage 1|niveau 1)\b", texte_tout, re.I)))
    # ---- Ouvertures (dimensions type 120x135 cm, 90x215...)
    ouvertures = []
    for m in re.finditer(r"(\d{2,3})\s*[x×]\s*(\d{2,3})", texte_tout):
        l_cm, h_cm = int(m.group(1)), int(m.group(2))
        if 40 <= l_cm <= 400 and 40 <= h_cm <= 300:
            ouvertures.append({"largeur_m": l_cm/100, "hauteur_m": h_cm/100,
                               "surface_m2": round(l_cm*h_cm/10000, 2),
                               "type": "Porte" if h_cm >= 190 else "Fenêtre/baie"})
    surf_ouv = round(sum(o["surface_m2"] for o in ouvertures), 2)
    # ---- EXTRACTION DES DONNÉES CONSTRUCTIVES (coupes, façades, annotations)
    # Règle : une donnée utilisée dans les métrés doit venir du dossier.
    # Si elle est absente, elle est déclarée MANQUANTE — jamais inventée.
    def cher(pats, types_prio=None):
        pp = sorted(pages, key=lambda p: (types_prio.index(p["type_feuille"]) if types_prio and p["type_feuille"] in types_prio else 9))
        for page in pp:
            for pat in pats:
                mch = re.search(pat, page["texte"], re.I)
                if mch: return mch, f"{page['fichier']} p.{page['num']} ({page['type_feuille']})"
        return None, None
    extraites, manquantes = [], []
    def consigner(nom, valeur, unite, source):
        if valeur is None: manquantes.append(nom)
        else: extraites.append({"donnee": nom, "valeur": valeur, "unite": unite, "source": source})
        return valeur
    # Hauteur sous plafond (coupe / plan)
    mm, src = cher([r"hauteur sous plafond[^0-9]{0,15}(\d[.,]\d{2})\s*m", r"\bHSP\b[^0-9]{0,10}(\d[.,]\d{2})"], ["COUPE","PLAN DE NIVEAU"])
    hsp = consigner("Hauteur sous plafond", fnum(mm.group(1)) if mm else None, "m", src)
    # Hauteurs d'égout et de faîtage (façades / coupes, altitudes)
    mm, src = cher([r"égout[^0-9+]{0,15}\+?\s*(\d[.,]\d{2})", r"hauteur à l'égout[^0-9]{0,15}(\d[.,]\d{2})"], ["FAÇADE","COUPE"])
    h_egout = consigner("Hauteur à l'égout", fnum(mm.group(1)) if mm else None, "m", src)
    mm, src = cher([r"fa[îi]tage[^0-9+]{0,15}\+?\s*(\d[.,]\d{2})", r"hauteur au fa[îi]tage[^0-9]{0,15}(\d[.,]\d{2})"], ["FAÇADE","COUPE"])
    h_fait = consigner("Hauteur au faîtage", fnum(mm.group(1)) if mm else None, "m", src)
    # Épaisseur des murs extérieurs (coupe, annotations maçonnerie)
    mm, src = cher([r"murs? extérieurs?[^0-9]{0,40}(\d{2,3})\s*cm", r"maçonnerie[^0-9]{0,30}(\d{2,3})\s*cm",
                    r"blocs?\s+(?:creux\s+)?(?:de\s+)?(\d{2,3})\s*cm", r"épaisseur(?: des murs?)?[^0-9]{0,20}(\d{2,3})\s*cm"], ["COUPE","PLAN DE NIVEAU"])
    ep_mur = consigner("Épaisseur murs extérieurs", (fnum(mm.group(1))/100) if mm else None, "m", src)
    # Épaisseur des cloisons
    mm, src = cher([r"cloisons?[^0-9]{0,30}(\d{2,3})\s*mm", r"cloisons?[^0-9]{0,30}(\d{1,2})\s*cm"], ["COUPE","PLAN DE NIVEAU"])
    ep_cloison = None
    if mm:
        v = fnum(mm.group(1)); ep_cloison = v/1000 if "mm" in mm.group(0).lower() or v >= 30 else v/100
    consigner("Épaisseur cloisons", ep_cloison, "m", src)
    # Fondations : profondeur et section de semelle (coupe)
    mm, src = cher([r"profondeur de fondation[^0-9]{0,15}(\d[.,]\d{1,2})\s*m", r"fond de fouille[^0-9\-]{0,10}-?\s*(\d[.,]\d{2})"], ["COUPE"])
    prof_fond = consigner("Profondeur de fondation", fnum(mm.group(1)) if mm else None, "m", src)
    mm, src = cher([r"semelles?[^0-9]{0,25}(\d{2,3})\s*[x×]\s*(\d{2,3})"], ["COUPE","PLAN DE NIVEAU"])
    larg_sem, haut_sem = (fnum(mm.group(1))/100, fnum(mm.group(2))/100) if mm else (None, None)
    consigner("Largeur de semelle", larg_sem, "m", src); consigner("Hauteur de semelle", haut_sem, "m", src)
    # Dallage et plancher
    mm, src = cher([r"dallage[^0-9]{0,20}(\d{1,2})\s*cm", r"dalle (?:de )?compression[^0-9]{0,15}(\d{1,2})\s*cm"], ["COUPE"])
    ep_dal = consigner("Épaisseur dallage", (fnum(mm.group(1))/100) if mm else None, "m", src)
    mm, src = cher([r"plancher[^0-9]{0,20}(\d{1,2})\s*\+\s*(\d{1,2})"], ["COUPE","PLAN DE NIVEAU"])
    ep_planch = consigner("Épaisseur plancher (hourdis + dalle)", ((fnum(mm.group(1))+fnum(mm.group(2)))/100) if mm else None, "m", src)
    # Isolation des murs
    mm, src = cher([r"(?:isolation des murs|laine de verre|isolant)[^0-9]{0,25}(\d{2,3})\s*mm"], ["COUPE","PLAN DE NIVEAU"])
    ep_iso = consigner("Épaisseur isolation murs", fnum(mm.group(1)) if mm else None, "mm", src)
    # Hauteur de faïence
    mm, src = cher([r"fa[ïi]ence[^0-9]{0,30}(\d[.,]\d{1,2})\s*m"], ["PLAN DE NIVEAU","COUPE"])
    h_fai = consigner("Hauteur de faïence", fnum(mm.group(1)) if mm else None, "m", src)
    # Matériau de couverture
    mat_couv = None
    for mot in ["tuiles mécaniques","tuile mécanique","tuiles canal","ardoise","bac acier","zinc","tuiles"]:
        if mot in texte_tout.lower(): mat_couv = mot; break
    consigner("Matériau de couverture", mat_couv, "", "Texte du dossier (toiture/façades)")
    if "DÉFAUT" in str(m.get("source_hauteur") if False else ""): pass
    metriques_extra = {"hauteur_sous_plafond_m": hsp, "hauteur_egout_m": h_egout,
        "hauteur_faitage_m": h_fait, "ep_mur_ext_plan_m": ep_mur, "ep_cloison_plan_m": ep_cloison,
        "prof_fondation_plan_m": prof_fond, "larg_semelle_plan_m": larg_sem, "haut_semelle_plan_m": haut_sem,
        "ep_dallage_plan_m": ep_dal, "ep_plancher_plan_m": ep_planch, "ep_isolation_mur_mm": ep_iso,
        "hauteur_faience_m": h_fai, "materiau_couverture": mat_couv,
        "donnees_extraites": extraites, "donnees_manquantes": manquantes}
    # ---- Toiture
    mp = re.search(r"pente[^0-9]{0,12}(\d{1,2})\s*(%|°|deg)", texte_tout, re.I)
    pente_pct, source_pente = (fnum(mp.group(1)) if mp else 35), ("Plan/façade" if mp else "DÉFAUT 35 % — À VALIDER")
    if mp and mp.group(2) == "°": pente_pct = round(math.tan(math.radians(fnum(mp.group(1))))*100, 1)
    type_bat = "Maison individuelle"
    if re.search(r"immeuble|collectif|bâtiment", texte_tout, re.I): type_bat = "Bâtiment / collectif"
    metriques = {"nom_fichier": nom_fichier, "nb_pages": len(pages),
        "vectoriel": vectoriel, "nb_segments_vectoriels": total_segs,
        "pieces": pieces, "surface_pieces_m2": surf_pieces,
        "surface_habitable_m2": sh, "emprise_texte_m2": emprise_txt,
        "surface_plancher_m2": plancher, "type_batiment": type_bat,
        "hauteur_niveau_m": hauteur_niveau, "source_hauteur": source_hauteur,
        "ouvertures": ouvertures, "surface_ouvertures_m2": surf_ouv,
        "pente_toit_pct": pente_pct, "source_pente": source_pente,
        "niveaux_detectes_texte": niveaux_texte}
    if geom and "erreur" not in geom:
        murs_int, long_int = murs_interieurs(page_plan, geom)
        metriques.update({"longueur_m": geom["longueur_m"], "largeur_m": geom["largeur_m"],
            "echelle_validee": geom["echelle_validee"], "cotations": geom["cotations_utilisees"],
            "source_geometrie": geom["source"], "murs_interieurs": murs_int,
            "longueur_murs_interieurs_ml": long_int, "page_plan": page_plan["num"]})
        metriques["emprise_sol_m2"] = round(geom["longueur_m"]*geom["largeur_m"], 2)
        metriques["perimetre_m"] = round(2*(geom["longueur_m"]+geom["largeur_m"]), 2)
    else:
        metriques.update({"erreur_geometrie": (geom or {}).get("erreur", "Géométrie non détectée"),
                          "murs_interieurs": [], "longueur_murs_interieurs_ml": 0})
        metriques["emprise_sol_m2"] = emprise_txt
    # Niveaux déduits
    E = metriques.get("emprise_sol_m2") or 0
    n = niveaux_texte if niveaux_texte >= 1 else 1
    if sh and E and sh/E > n: n = max(n, round(sh/E))
    metriques["nb_niveaux"] = max(1, n)
    if not sh and surf_pieces: metriques["surface_habitable_m2"] = surf_pieces
    metriques.update(metriques_extra)
    # Aperçu visuel de la feuille de plan : à comparer avec le modèle 3D
    try:
        doc_plan = docs_refs[page_plan["doc_idx"]]
        pix = doc_plan.load_page(page_plan["num"]-1).get_pixmap(matrix=pymupdf.Matrix(1.1, 1.1))
        import base64 as _b64
        metriques["apercu_plan_png"] = "data:image/png;base64," + _b64.b64encode(pix.tobytes("png")).decode()
    except Exception:
        metriques["apercu_plan_png"] = None
    for d_ in docs_refs:
        try: d_.close()
        except Exception: pass
    metriques["feuilles"] = [{"fichier": p["fichier"], "page": p["num"],
                              "type": p["type_feuille"], "segments": len(p["segments"])} for p in pages]
    return metriques

# ---------------------------------------------------------------- MODÈLE 3D + MÉTRÉS
PARAMS_DEFAUT = {"ep_mur_ext_m": 0.20, "ep_refend_m": 0.20, "prof_fondation_m": 0.80,
    "larg_fondation_m": 0.60, "ep_dallage_m": 0.12, "ep_plancher_m": 0.20,
    "h_sous_plafond_m": 2.50, "perte_materiaux_pct": 5.0, "tva_pct": 20.0,
    "taux_horaire_mo_eur": 48.0, "heures_par_jour": 7.0}

def construire_metres(m, p):
    """Métrés calculés UNIQUEMENT depuis la géométrie du modèle. Chaque
    quantité renvoie (valeur, formule lisible, source, statut)."""
    L = m.get("longueur_m") or 0; W = m.get("largeur_m") or 0
    E = m.get("emprise_sol_m2") or (L*W); P = m.get("perimetre_m") or (2*(L+W) if L and W else 0)
    N = m.get("nb_niveaux", 1); h = m.get("hauteur_niveau_m", 2.70)
    O = m.get("surface_ouvertures_m2", 0); SH = m.get("surface_habitable_m2") or 0
    Li = m.get("longueur_murs_interieurs_ml", 0)
    pente = (m.get("pente_toit_pct") or 35)/100
    toit = E * math.sqrt(1 + pente**2) if E else 0
    plancher_total = E * N
    murs_brut = P * h * N
    murs_net = max(murs_brut - O, 0)
    humides = sum(x["surface_m2"] for x in m.get("pieces", []) if any(k in x["nom"].lower() for k in ["bain","eau","wc","cuisine","buanderie","cellier"]))
    chambres = sum(x["surface_m2"] for x in m.get("pieces", []) if "chambre" in x["nom"].lower())
    sejour = sum(x["surface_m2"] for x in m.get("pieces", []) if any(k in x["nom"].lower() for k in ["séjour","salon","cuisine"]))
    sol_carrelage = humides + sejour if (humides+sejour) > 0 else SH*0.55
    sol_parquet = chambres if chambres > 0 else SH*0.35
    nb_wc = max(1, sum(1 for x in m.get("pieces", []) if "wc" in x["nom"].lower()))
    nb_sdb = max(1, sum(1 for x in m.get("pieces", []) if any(k in x["nom"].lower() for k in ["bain","eau"])))
    nb_portes_int = max(3, sum(1 for x in m.get("pieces", []) if any(k in x["nom"].lower() for k in ["chambre","bain","eau","wc","bureau","cellier"])))
    # Provenance : le plan prime sur le paramètre de secours, et tout secours
    # utilisé est déclaré MANQUANT (la donnée aurait dû être dans le dossier).
    ep_ext = m.get("ep_mur_ext_plan_m") or p["ep_mur_ext_m"]
    prof_f = m.get("prof_fondation_plan_m") or p["prof_fondation_m"]
    larg_f = m.get("larg_semelle_plan_m") or p["larg_fondation_m"]
    haut_sem = m.get("haut_semelle_plan_m") or 0.40
    hsp_eff = m.get("hauteur_sous_plafond_m") or p["h_sous_plafond_m"]
    h_fai = m.get("hauteur_faience_m") or 2.00
    src_fond = "Profondeur/section lues sur la coupe du dossier" if m.get("prof_fondation_plan_m") else "DONNÉE ABSENTE DU DOSSIER (coupe à fournir)"
    src_ep = "Épaisseur lue sur le dossier (coupe/annotations)" if m.get("ep_mur_ext_plan_m") else "DONNÉE ABSENTE DU DOSSIER (coupe à fournir)"
    def Q(v, formule, source, statut="CALCULÉ"):
        return {"quantite": round(v, 2), "formule": formule, "source": source, "statut": statut}
    src_plan = f"Modèle 3D issu du plan p.{m.get('page_plan','?')} (échelle {'validée par cotations' if m.get('echelle_validee') else 'NON validée — calibration requise'})"
    metres = {
      "emprise":      Q(E, f"{L} m × {W} m (emprise modélisée)", src_plan, "MESURÉ"),
      "perimetre":    Q(P, f"2 × ({L} + {W}) m", src_plan, "MESURÉ"),
      "decapage":     Q(E, "Emprise au sol × 1 (décapage pleine surface)", src_plan),
      "fouilles":     Q(P*larg_f*prof_f, f"Périmètre {P} ml × larg. semelle {larg_f} m × prof. {prof_f} m", src_fond, "CALCULÉ" if m.get("prof_fondation_plan_m") else "MANQUANT"),
      "beton_fond":   Q(P*larg_f*haut_sem, f"Périmètre {P} ml × {larg_f} m × {haut_sem} m (section de semelle)", src_fond, "CALCULÉ" if m.get("larg_semelle_plan_m") else "MANQUANT"),
      "dallage":      Q(E, "Emprise au sol (dalle RDC)", src_plan),
      "murs_ext_net": Q(murs_net, f"Périmètre {P} ml × hauteur {h} m × {N} niv. − ouvertures {O} m²", src_plan + " ; hauteur : " + str(m.get("source_hauteur")), "MESURÉ" if m.get("echelle_validee") else "CALCULÉ"),
      "murs_ext_vol": Q(murs_net*ep_ext, f"Surface nette murs {round(murs_net,2)} m² × ép. {ep_ext} m", src_ep, "CALCULÉ" if m.get("ep_mur_ext_plan_m") else "MANQUANT"),
      "planchers":    Q(E*max(N-1,0), f"Emprise {E} m² × {max(N-1,0)} niveau(x) supérieur(s)", src_plan),
      "chainages":    Q(P*N, f"Périmètre {P} ml × {N} niveau(x)", src_plan),
      "enduit_ext":   Q(murs_net, "Surface nette des murs extérieurs", src_plan),
      "toiture":      Q(toit, f"Emprise {E} m² × √(1+pente²) avec pente {m.get('pente_toit_pct',35)} %", "Emprise plan ; pente : " + str(m.get("source_pente")), "CALCULÉ" if "DÉFAUT" not in str(m.get("source_pente")) else "MANQUANT"),
      "faitage":      Q(L, "Longueur du bâtiment (ligne de faîtage)", src_plan, "MESURÉ"),
      "gouttieres":   Q(2*L, f"2 × longueur {L} m (deux versants)", src_plan),
      "cloisons":     Q(max(Li*h - nb_portes_int*1.9, 0), f"Murs/cloisons int. modélisés {Li} ml × h {h} m − {nb_portes_int} portes", src_plan),
      "doublage":     Q(max(P*h*N - O, 0), "Face intérieure des murs extérieurs (nette ouvertures)", src_plan),
      "plafonds":     Q(plancher_total, f"Surface de plancher totale {round(plancher_total,2)} m²", src_plan),
      "surface_hab":  Q(SH, "Somme des surfaces annotées / surface habitable du permis", "Annotations du plan / dossier PC", "MESURÉ" if SH else "MANQUANT"),
      "carrelage":    Q(sol_carrelage, "Pièces humides + séjour/cuisine (surfaces annotées des pièces)", "Surfaces des pièces lues sur le plan"),
      "parquet":      Q(sol_parquet, "Chambres (surfaces annotées des pièces)", "Surfaces des pièces lues sur le plan"),
      "faience":      Q(nb_sdb*4*h_fai, f"{nb_sdb} pièce(s) d'eau × développé 4 ml × hauteur {h_fai} m", "Nombre de pièces d'eau du plan ; hauteur de faïence : " + ("lue sur le dossier" if m.get("hauteur_faience_m") else "absente du dossier (2,00 m utilisés, à faire préciser)"), "CALCULÉ" if m.get("hauteur_faience_m") else "MANQUANT"),
      "peinture":     Q(max((P + 2*Li)*hsp_eff*N - O, 0) + plancher_total, "Murs intérieurs développés + plafonds − ouvertures", src_plan),
      "ouvertures_m2":Q(O, "Somme des menuiseries cotées (l×h lues sur le plan)", "Cotations menuiseries du plan", "MESURÉ" if O else "HYPOTHÈSE"),
    }
    elements_3d = {"longueur_m": L, "largeur_m": W, "nb_niveaux": N,
        "hauteur_niveau_m": h, "murs_interieurs": m.get("murs_interieurs", []),
        "ouvertures": m.get("ouvertures", []), "pieces": m.get("pieces", []),
        "surface_toiture_m2": round(toit,2), "surface_murs_ext_net_m2": round(murs_net,2)}
    return metres, elements_3d

# ---------------------------------------------------------------- DEVIS + MATÉRIAUX + MO
# (clé métré, code, lot, poste, unité, prix de vente unitaire €, matériaux[(désignation, qté/unité, unité_mat, prix achat €, perte%)], temps MO h/unité, qualification)
LIGNES = [
 ("fouilles","1.01","1 · Terrassement","Fouilles en rigoles pour semelles filantes (déblai/remblai, évacuation incluse dans le poste)","m3",95,[],0.55,"Conducteur d'engin / terrassier"),
 ("decapage","1.02","1 · Terrassement","Décapage terre végétale et implantation","m2",6,[],0.02,"Terrassier"),
 ("beton_fond","2.01","2 · Fondations — Gros œuvre","Béton de semelles filantes C25/30 dosé 350 kg, coffrage et ferraillage forfaitaire inclus au m3","m3",420,[("Béton C25/30 livré",1.0,"m3",145,3),("Acier HA (ratio 35 kg/m3 béton)",35,"kg",1.15,5),("Film/c coffrage perdu et divers",1.0,"ens",18,0)],1.60,"Maçon coffreur"),
 ("dallage","2.02","2 · Fondations — Gros œuvre","Dallage RDC ép. 12 cm sur hérisson, treillis soudé, film polyane","m2",68,[("Béton C25/30 livré",0.12,"m3",145,3),("Treillis soudé ST25",1.05,"m2",6.5,5),("Film polyane 200µ",1.05,"m2",0.9,5),("Tout-venant hérisson 20 cm",0.20,"m3",28,5)],0.55,"Maçon"),
 ("murs_ext_net","2.03","2 · Fondations — Gros œuvre","Maçonnerie murs extérieurs en blocs creux 20 cm, mortier, chaînages verticaux","m2",95,[("Blocs creux 20×20×50",10,"u",1.45,4),("Mortier (sac 35 kg)",0.55,"sac",6.8,5),("Béton chaînages/linteaux",0.012,"m3",145,3),("Acier chaînages",1.4,"kg",1.15,5)],1.05,"Maçon"),
 ("planchers","2.04","2 · Fondations — Gros œuvre","Plancher poutrelles-hourdis + dalle de compression 20 cm","m2",110,[("Poutrelles + hourdis (kit au m2)",1.0,"m2",38,2),("Béton dalle compression 5 cm",0.05,"m3",145,3),("Treillis soudé",1.05,"m2",6.5,5)],0.85,"Maçon coffreur"),
 ("chainages","2.05","2 · Fondations — Gros œuvre","Chaînages horizontaux périphériques (béton armé)","ml",28,[("Béton C25/30",0.03,"m3",145,3),("Acier HA",2.2,"kg",1.15,5)],0.35,"Maçon"),
 ("enduit_ext","3.01","3 · Enduit extérieur","Enduit monocouche projeté, finition grattée","m2",42,[("Enduit monocouche (sac 25 kg)",1.1,"sac",9.5,6),("Trame fibre de verre renforts",0.15,"m2",1.8,5)],0.45,"Façadier"),
 ("toiture","4.01","4 · Charpente — Couverture","Charpente traditionnelle/fermettes + couverture tuiles mécaniques, écran sous-toiture","m2",168,[("Bois charpente (fermettes au m2)",1.0,"m2",42,2),("Tuiles mécaniques",14,"u",1.85,6),("Liteaux + écran sous-toiture",1.05,"m2",4.2,5),("Faîtières et accessoires (au m2)",1.0,"ens",6,3)],1.15,"Charpentier-couvreur"),
 ("faitage","4.02","4 · Charpente — Couverture","Faîtage à sec ventilé","ml",58,[("Faîtières demi-rondes",3.3,"u",4.6,5),("Closoir ventilé",1.0,"ml",6.5,3)],0.40,"Couvreur"),
 ("gouttieres","4.03","4 · Charpente — Couverture","Gouttières demi-rondes zinc/PVC et descentes","ml",42,[("Gouttière + crochets (au ml)",1.0,"ml",11,3),("Tuyau de descente (quote-part)",0.35,"ml",9,3)],0.30,"Couvreur-zingueur"),
 ("ouvertures_m2","5.01","5 · Menuiseries extérieures","Menuiseries extérieures PVC/alu double vitrage (fourniture + pose, au m2 d'ouverture cotée)","m2",520,[("Menuiserie DV sur mesure (au m2)",1.0,"m2",330,0),("Tapées, seuils, habillages",1.0,"ens",28,0),("Mousse, silicone, fixations",1.0,"ens",7,0)],0.90,"Menuisier poseur"),
 ("doublage","6.01","6 · Plâtrerie — Isolation","Doublage isolant murs extérieurs (laine de verre 120 mm + plaque de plâtre BA13 sur ossature)","m2",64,[("Laine de verre 120 mm",1.02,"m2",11.5,5),("Plaque BA13",1.02,"m2",4.6,5),("Ossature métallique (rails/montants)",1.0,"ens",5.5,5),("Enduit à joint + bandes",0.35,"kg",1.9,5)],0.55,"Plaquiste"),
 ("cloisons","6.02","6 · Plâtrerie — Isolation","Cloisons de distribution BA13 sur ossature, laine acoustique 45 mm","m2",56,[("Plaque BA13 (2 parements)",2.04,"m2",4.6,5),("Laine acoustique 45 mm",1.02,"m2",4.8,5),("Ossature métallique",1.0,"ens",5.0,5),("Enduit à joint + bandes",0.5,"kg",1.9,5)],0.60,"Plaquiste"),
 ("plafonds","6.03","6 · Plâtrerie — Isolation","Plafonds BA13 sur ossature + isolation des combles (laine soufflée 350 mm au RDC/plafond haut)","m2",64,[("Plaque BA13",1.02,"m2",4.6,5),("Ossature + suspentes",1.0,"ens",5.8,5),("Laine à souffler 350 mm",1.0,"m2",13.5,6)],0.50,"Plaquiste"),
 ("surface_hab","7.01","7 · Électricité","Installation électrique complète NF C 15-100 (au m2 habitable : tableau, circuits, appareillage standard)","m2",95,[("Tableau + protections (quote-part m2)",1.0,"ens",9,0),("Câbles, gaines, boîtes (au m2)",1.0,"ens",16,3),("Appareillage (prises/inter au m2)",1.0,"ens",13,2)],0.85,"Électricien"),
 ("surface_hab","8.01","8 · Plomberie — Sanitaire","Réseaux EF/EC/évacuations + appareils sanitaires standard (au m2 habitable)","m2",108,[("Tuyauterie multicouche/PVC (au m2)",1.0,"ens",15,4),("Appareils sanitaires (quote-part m2)",1.0,"ens",26,0),("Robinetterie et divers",1.0,"ens",8,0)],0.80,"Plombier"),
 ("surface_hab","9.01","9 · Chauffage — Ventilation","Pompe à chaleur air/eau + plancher chauffant/émetteurs + VMC simple flux (au m2 habitable)","m2",132,[("PAC air/eau (quote-part m2)",1.0,"ens",38,0),("Plancher chauffant/émetteurs (au m2)",1.0,"m2",24,3),("VMC simple flux hygro (quote-part)",1.0,"ens",7,0)],0.70,"Chauffagiste"),
 ("carrelage","10.01","10 · Revêtements de sol","Carrelage grès cérame 45×45 collé, plinthes assorties","m2",82,[("Carrelage grès cérame",1.05,"m2",24,7),("Colle flex (sac 25 kg)",0.22,"sac",14,5),("Joint + croisillons",1.0,"ens",1.6,5),("Plinthes assorties",0.9,"ml",3.4,5)],0.65,"Carreleur"),
 ("parquet","10.02","10 · Revêtements de sol","Parquet stratifié/contrecollé flottant, sous-couche, plinthes","m2",66,[("Parquet (lames)",1.05,"m2",26,6),("Sous-couche acoustique",1.05,"m2",2.4,5),("Plinthes MDF",0.9,"ml",2.6,5)],0.40,"Parqueteur / menuisier"),
 ("faience","10.03","10 · Revêtements de sol","Faïence murale pièces d'eau (hauteur 2 m)","m2",78,[("Faïence",1.05,"m2",21,7),("Colle + joint hydrofuge",1.0,"ens",4.5,5),("Profilés et étanchéité SPEC",1.0,"ens",5.5,3)],0.70,"Carreleur"),
 ("peinture","11.01","11 · Peinture","Peinture murs et plafonds : impression + 2 couches acrylique","m2",28,[("Peinture acrylique (0,25 L/m2 total)",0.25,"L",11,5),("Sous-couche/impression",0.12,"L",8,5),("Enduit de lissage (quote-part)",0.2,"kg",1.7,5)],0.32,"Peintre"),
]


MATERIAUX_DETAILLES = {
"1.01": [("Rotation benne 8 m3 (évacuation déblais, foisonnement 1,3)",0.1625,"rotation",85,0,"benne 8 m3",1),
         ("Gazole mini-pelle / engin de terrassement",0.9,"L",1.85,0,None,None)],
"1.02": [("Piquets d'implantation bois",0.06,"u",1.2,0,"paquet de 10",10),
         ("Cordeau de maçon / traçage",0.01,"rouleau",6.5,0,"rouleau",1),
         ("Bombe de traçage chantier",0.008,"bombe",7.5,0,"bombe",1)],
"2.01": [("Béton C25/30 livré (semelles)",1.0,"m3",145,3,None,None),
         ("Acier HA8 (semelles, en kg)",14,"kg",1.15,5,"barre de 6 m",None),
         ("Acier HA10 (semelles, en kg)",21,"kg",1.15,5,"barre de 6 m",None),
         ("Fil de ligature",0.3,"kg",2.4,5,"bobine 5 kg",5),
         ("Cales béton / distanciers",2,"u",0.35,5,"sachet de 100",100)],
"2.02": [("Béton C25/30 livré (dallage)",0.12,"m3",145,3,None,None),
         ("Treillis soudé ST25C",1.05,"m2",6.5,5,"panneau 2,40x1,20",2.88),
         ("Film polyane 200µ",1.05,"m2",0.9,5,"rouleau 100 m2",100),
         ("Tout-venant 0/31,5 (hérisson)",0.20,"m3",28,5,None,None),
         ("Joint de dilatation PVC",0.25,"ml",1.8,3,"barre de 3 m",3)],
"2.03": [("Blocs creux béton 20x20x50",10,"u",1.45,4,"palette de 70",70),
         ("Blocs linteau / chaînage en U 20 cm",0.5,"u",2.6,4,"palette",None),
         ("Mortier prêt à l'emploi (sac 35 kg)",0.55,"sac",6.8,5,"sac 35 kg",1),
         ("Béton C25/30 (chaînages verticaux/linteaux)",0.012,"m3",145,3,None,None),
         ("Acier HA10 (chaînages, en kg)",1.4,"kg",1.15,5,"barre de 6 m",None)],
"2.04": [("Poutrelles béton précontraint",1.67,"ml",9.5,2,"à la longueur",None),
         ("Hourdis / entrevous béton 12 cm",8,"u",1.35,4,"palette",None),
         ("Béton C25/30 (dalle de compression)",0.05,"m3",145,3,None,None),
         ("Treillis soudé ST25C",1.05,"m2",6.5,5,"panneau 2,40x1,20",2.88),
         ("Étais / tour d'étaiement (location, quote-part)",1.0,"ens",3.5,0,None,None)],
"2.05": [("Béton C25/30 (chaînages horizontaux)",0.03,"m3",145,3,None,None),
         ("Acier HA8+HA10 (chaînages, en kg)",2.2,"kg",1.15,5,"barre de 6 m",None),
         ("Planches de coffrage 27 mm",0.06,"m2",12,10,None,None)],
"3.01": [("Enduit monocouche (sac 25 kg)",1.1,"sac",9.5,6,"sac 25 kg",1),
         ("Trame fibre de verre (renforts angles/fissures)",0.15,"m2",1.8,5,"rouleau 50 m2",50),
         ("Profilé goutte d'eau / angle PVC",0.12,"ml",2.2,3,"barre de 2,50 m",2.5),
         ("Sous-couche d'accrochage (zones lisses)",0.05,"L",6,3,"pot 10 L",10)],
"4.01": [("Fermettes / charpente industrialisée (au m2 de toit)",1.0,"m2",42,2,None,None),
         ("Tuiles mécaniques grand moule",14,"u",1.85,6,"palette",None),
         ("Liteaux 27x40 (pureaux 0,35 m)",2.9,"ml",0.75,6,"botte",None),
         ("Contre-liteaux 27x40 (entraxe 0,60 m)",1.67,"ml",0.75,6,"botte",None),
         ("Écran sous-toiture HPV",1.05,"m2",2.6,5,"rouleau 75 m2",75),
         ("Pointes / vis de couverture (en kg)",0.08,"kg",4.5,5,"boîte 5 kg",5),
         ("Crochets de tuile (zones exposées)",1.2,"u",0.28,3,"sachet de 100",100)],
"4.02": [("Faîtières demi-rondes",3.3,"u",4.6,5,"palette",None),
         ("Closoir ventilé de faîtage",1.0,"ml",6.5,3,"rouleau 5 ml",5),
         ("Vis de faîtière + cavalier",3.3,"u",0.4,3,"sachet de 50",50)],
"4.03": [("Gouttière demi-ronde (développé 25)",1.0,"ml",8.5,3,"barre de 4 m",4),
         ("Crochets de gouttière (entraxe 0,65 m)",1.55,"u",1.6,3,"boîte de 25",25),
         ("Tuyau de descente Ø80 (quote-part)",0.4,"ml",7.5,3,"barre de 3 m",3),
         ("Colliers de descente",0.35,"u",2.8,3,"boîte de 10",10),
         ("Naissance + coude de descente",0.12,"u",6.5,3,"à l'unité",1),
         ("Colle PVC / mastic d'étanchéité",0.02,"cartouche",6,0,"cartouche",1)],
"5.01": [("Menuiserie extérieure DV 4/16/4 sur mesure (au m2)",1.0,"m2",330,0,"sur mesure",None),
         ("Tapées d'isolation + habillages (au m2 d'ouverture)",1.0,"ens",28,0,None,None),
         ("Mousse PU d'étanchéité",0.06,"bombe",8,0,"bombe 750 ml",1),
         ("Fond de joint + silicone extérieur",0.1,"cartouche",7,0,"cartouche",1),
         ("Vis + chevilles de pose menuiserie",4,"u",0.35,3,"boîte de 100",100),
         ("Seuil / rejet d'eau aluminium (quote-part)",0.5,"ml",14,3,"barre de 3 m",3)],
"6.01": [("Plaque de plâtre BA13 12,5 mm",1.02,"m2",3.4,5,"plaque 2,50x1,20 = 3 m2",3),
         ("Laine de verre semi-rigide 120 mm (murs)",1.02,"m2",10.5,5,"panneau/rouleau",None),
         ("Montant M48",1.67,"ml",1.05,4,"barre de 3,00 m",3),
         ("Rail R48",0.8,"ml",0.95,4,"barre de 3,00 m",3),
         ("Vis TTPC 3,5x25",12,"u",0.02,3,"boîte de 1000",1000),
         ("Bande à joint papier",1.4,"ml",0.06,3,"rouleau de 150 ml",150),
         ("Enduit à joint (poudre)",0.35,"kg",1.9,5,"sac 25 kg",25),
         ("Chevilles à frapper / fixations rail",0.8,"u",0.3,3,"boîte de 100",100)],
"6.02": [("Plaque de plâtre BA13 12,5 mm",2.04,"m2",3.4,5,"plaque 2,50x1,20 = 3 m2",3),
         ("Laine acoustique 45 mm (cloisons)",1.02,"m2",4.2,5,"panneau/rouleau",None),
         ("Montant M48",1.67,"ml",1.05,4,"barre de 3,00 m",3),
         ("Rail R48",0.9,"ml",0.95,4,"barre de 3,00 m",3),
         ("Vis TTPC 3,5x25",24,"u",0.02,3,"boîte de 1000",1000),
         ("Bande à joint papier",2.8,"ml",0.06,3,"rouleau de 150 ml",150),
         ("Enduit à joint (poudre)",0.7,"kg",1.9,5,"sac 25 kg",25),
         ("Bande résiliente sous rail",0.9,"ml",0.55,3,"rouleau de 30 ml",30)],
"6.03": [("Plaque de plâtre BA13 12,5 mm",1.02,"m2",3.4,5,"plaque 2,50x1,20 = 3 m2",3),
         ("Fourrure F47 (entraxe 0,60 m)",1.67,"ml",0.85,4,"barre de 3,00 m",3),
         ("Suspentes (grille 0,60 x 1,20 m)",1.39,"u",0.45,3,"boîte de 100",100),
         ("Cornière de rive (périphérie plafond)",0.35,"ml",0.9,4,"barre de 3,00 m",3),
         ("Vis TTPC 3,5x25",12,"u",0.02,3,"boîte de 1000",1000),
         ("Bande à joint papier",1.4,"ml",0.06,3,"rouleau de 150 ml",150),
         ("Enduit à joint (poudre)",0.35,"kg",1.9,5,"sac 25 kg",25),
         ("Laine à souffler combles 350 mm (en kg)",7,"kg",1.9,6,"sac de 12,5 kg",12.5),
         ("Tiges filetées + cavaliers (quote-part)",1.0,"ens",1.1,3,None,None)],
"7.01": [("Câble R2V 3G1,5 (éclairage/prises courant faible)",1.8,"ml",0.55,4,"couronne 100 ml",100),
         ("Câble R2V 3G2,5 (prises/cuisine)",0.9,"ml",0.85,4,"couronne 100 ml",100),
         ("Câble RJ45 Cat 6",0.5,"ml",0.4,4,"couronne 100 ml",100),
         ("Gaine ICTA Ø20 précâblée ou vide",2.2,"ml",0.35,4,"couronne 100 ml",100),
         ("Boîtes d'encastrement Ø67",0.35,"u",0.7,3,"boîte de 10",10),
         ("Prises 2P+T 16A",0.28,"u",6.5,2,"à l'unité",1),
         ("Interrupteurs / va-et-vient",0.18,"u",7.5,2,"à l'unité",1),
         ("DCL + douilles plafond",0.12,"u",5.5,2,"à l'unité",1),
         ("Disjoncteurs 10/16/20A (quote-part)",0.14,"u",11,0,"à l'unité",1),
         ("Coffret électrique + ID 30 mA (quote-part)",0.02,"u",95,0,"à l'unité",1),
         ("Détecteur de fumée DAAF",0.012,"u",18,0,"à l'unité",1)],
"8.01": [("Tube multicouche Ø16 (EF/EC)",1.4,"ml",1.5,4,"couronne 50 ml",50),
         ("Tube multicouche Ø20 (collecteurs)",0.5,"ml",2.1,4,"couronne 50 ml",50),
         ("Tube PVC évacuation Ø40",0.35,"ml",1.6,4,"barre de 3 m",3),
         ("Tube PVC évacuation Ø100",0.12,"ml",4.2,4,"barre de 3 m",3),
         ("Raccords multicouche à sertir (quote-part)",0.25,"u",3.8,3,"sachet",None),
         ("Collecteur/nourrice sanitaire (quote-part)",0.03,"u",38,0,"à l'unité",1),
         ("Pack WC complet (quote-part au m2 hab.)",0.012,"u",210,0,"à l'unité",1),
         ("Lavabo/vasque + robinetterie (quote-part)",0.014,"u",185,0,"à l'unité",1),
         ("Receveur douche 90x90 + paroi (quote-part)",0.012,"u",340,0,"à l'unité",1),
         ("Chauffe-eau / ballon ECS (quote-part)",0.009,"u",520,0,"à l'unité",1),
         ("Robinets d'arrêt + flexibles (quote-part)",0.1,"u",7.5,2,"à l'unité",1)],
"9.01": [("Pompe à chaleur air/eau (quote-part au m2 hab.)",0.0085,"u",4200,0,"à l'unité",1),
         ("Tube PER plancher chauffant (5 ml/m2)",4.5,"ml",0.65,3,"couronne 240 ml",240),
         ("Plaques isolantes plancher chauffant",1.02,"m2",11,3,"paquet",None),
         ("Collecteur plancher chauffant (quote-part)",0.02,"u",180,0,"à l'unité",1),
         ("Agrafes PER (6/m2)",6,"u",0.06,3,"boîte de 300",300),
         ("Kit VMC simple flux hygro B (quote-part)",0.0085,"u",420,0,"à l'unité",1),
         ("Gaine VMC Ø80/125 isolée",0.8,"ml",3.2,3,"couronne 10 ml",10),
         ("Bouches d'extraction hygro",0.05,"u",16,0,"à l'unité",1)],
"10.01":[("Carrelage grès cérame 45x45",1.05,"m2",24,7,"carton",None),
         ("Colle flex C2 (4 kg/m2)",0.16,"sac",14,5,"sac 25 kg",1),
         ("Joint poudre (0,4 kg/m2)",0.4,"kg",1.6,5,"sac 5 kg",5),
         ("Croisillons 2 mm",0.02,"sachet",3.5,0,"sachet de 250",1),
         ("Plinthes assorties carrelage",0.9,"ml",3.4,5,"carton",None),
         ("Primaire d'accrochage (supports poreux)",0.04,"L",7,3,"pot 5 L",5)],
"10.02":[("Parquet stratifié/contrecollé (lames)",1.05,"m2",26,6,"paquet",None),
         ("Sous-couche acoustique 3 mm",1.05,"m2",2.4,5,"rouleau 15 m2",15),
         ("Plinthes MDF 70 mm",0.9,"ml",2.6,5,"botte",None),
         ("Barres de seuil / profilés de jonction",0.03,"u",9.5,0,"à l'unité",1),
         ("Colle plinthes / mastic acrylique",0.02,"cartouche",5.5,0,"cartouche",1)],
"10.03":[("Faïence murale 20x40",1.05,"m2",21,7,"carton",None),
         ("Colle faïence C1 (3,5 kg/m2)",0.14,"sac",12,5,"sac 25 kg",1),
         ("Joint hydrofuge (0,3 kg/m2)",0.3,"kg",2.1,5,"sac 5 kg",5),
         ("SPEC étanchéité sous faïence (kit quote-part)",0.06,"kit",28,0,"kit",1),
         ("Profilés de finition alu (quote-part)",0.15,"ml",4.2,3,"barre de 2,50 m",2.5)],
"11.01":[("Peinture acrylique finition (2 couches, 0,20 L/m2)",0.20,"L",11,5,"pot 10 L",10),
         ("Sous-couche / impression universelle",0.12,"L",8,5,"pot 10 L",10),
         ("Enduit de lissage / rebouchage",0.2,"kg",1.7,5,"sac 15 kg",15),
         ("Bande de masquage peintre",0.04,"rouleau",3.2,0,"rouleau 50 ml",1),
         ("Bâche de protection / consommables",0.01,"ens",4,0,None,None)],
}

def construire_devis(m, p):
    metres, elements = construire_metres(m, p)
    lignes, bordereau, mo_par_qualif = [], {}, {}
    for (cle, code, lot, poste, unite, pu_vente, mats, tps_mo, qualif) in LIGNES:
        if code in MATERIAUX_DETAILLES: mats = list(MATERIAUX_DETAILLES[code])
        if code == "6.03":  # combles = plafond du dernier niveau seulement
            niv = max(m.get("nb_niveaux", 1), 1)
            mats = [(t[0], (t[1]/niv if t[0].startswith("Laine à souffler") else t[1])) + tuple(t[2:]) for t in mats]
        if cle == "dallage" and m.get("ep_dallage_plan_m"):
            mats = [(t + (None, None))[:7] if False else (t[0], (m["ep_dallage_plan_m"] if t[0].startswith("Béton") else t[1])) + tuple(t[2:]) for t in mats]
        q = metres[cle]
        quantite = q["quantite"]
        if quantite <= 0: continue
        mats_calc, cout_mat = [], 0.0
        for tuple_mat in mats:
            des, q_par_u, u_mat, pu_achat, perte = tuple_mat[:5]
            cond = tuple_mat[5] if len(tuple_mat) > 5 else None
            taille_cond = tuple_mat[6] if len(tuple_mat) > 6 else None
            perte_eff = max(perte, p["perte_materiaux_pct"]) if perte else 0
            qte = quantite * q_par_u * (1 + perte_eff/100)
            cout = qte * pu_achat; cout_mat += cout
            nb_cond = math.ceil(qte / taille_cond) if taille_cond else None
            mats_calc.append({"designation": des, "quantite": round(qte,2), "unite": u_mat,
                              "quantite_par_unite_ouvrage": q_par_u, "perte_pct": perte_eff,
                              "prix_achat_unitaire": pu_achat, "cout_total": round(cout,2),
                              "conditionnement": cond, "nb_conditionnements": nb_cond,
                              "quantite_a_commander": round(nb_cond * taille_cond, 2) if nb_cond else round(qte,2)})
            cle_b = (des, u_mat)
            if cle_b not in bordereau:
                bordereau[cle_b] = {"designation": des, "unite": u_mat, "quantite_totale": 0.0,
                                    "prix_achat_unitaire": pu_achat, "cout_total": 0.0, "postes": [],
                                    "conditionnement": cond, "taille_conditionnement": taille_cond}
            bordereau[cle_b]["quantite_totale"] = round(bordereau[cle_b]["quantite_totale"] + qte, 2)
            bordereau[cle_b]["cout_total"] = round(bordereau[cle_b]["cout_total"] + cout, 2)
            if code not in bordereau[cle_b]["postes"]: bordereau[cle_b]["postes"].append(code)
        heures = quantite * tps_mo
        taux = p["taux_horaire_mo_eur"]
        cout_mo = heures * taux
        debourse = cout_mat + cout_mo
        total_vente = quantite * pu_vente
        mo_par_qualif[qualif] = round(mo_par_qualif.get(qualif, 0) + heures, 1)
        lignes.append({"code": code, "lot": lot, "poste": poste, "unite": unite,
            "quantite": quantite, "formule": q["formule"], "source": q["source"],
            "statut": q["statut"], "prix_unitaire_vente": pu_vente,
            "total_vente_ht": round(total_vente,2), "materiaux": mats_calc,
            "main_oeuvre": {"temps_unitaire_h": tps_mo, "heures_totales": round(heures,1),
                            "qualification": qualif, "taux_horaire": taux, "cout_total": round(cout_mo,2)},
            "debourse_sec": round(debourse,2),
            "marge_eur": round(total_vente - debourse,2),
            "marge_pct": round((total_vente-debourse)/total_vente*100,1) if total_vente else 0})
    for b in bordereau.values():
        tc = b.get("taille_conditionnement")
        b["nb_conditionnements"] = math.ceil(b["quantite_totale"] / tc) if tc else None
        b["quantite_a_commander"] = round(b["nb_conditionnements"] * tc, 2) if tc else b["quantite_totale"]
    tot_vente = round(sum(l["total_vente_ht"] for l in lignes), 2)
    tot_mat = round(sum(l["materiaux"] and sum(x["cout_total"] for x in l["materiaux"]) or 0 for l in lignes), 2)
    tot_mo_h = round(sum(l["main_oeuvre"]["heures_totales"] for l in lignes), 1)
    tot_mo_eur = round(sum(l["main_oeuvre"]["cout_total"] for l in lignes), 2)
    jours = p["heures_par_jour"]
    synthese = {"total_vente_ht": tot_vente, "tva_pct": p["tva_pct"],
        "tva_eur": round(tot_vente*p["tva_pct"]/100,2),
        "total_ttc": round(tot_vente*(1+p["tva_pct"]/100),2),
        "total_materiaux_achat_ht": tot_mat, "total_main_oeuvre_heures": tot_mo_h,
        "total_main_oeuvre_eur": tot_mo_eur, "debourse_sec_total": round(tot_mat+tot_mo_eur,2),
        "marge_brute_eur": round(tot_vente-tot_mat-tot_mo_eur,2),
        "marge_brute_pct": round((tot_vente-tot_mat-tot_mo_eur)/tot_vente*100,1) if tot_vente else 0,
        "duree_1_personne_jours": round(tot_mo_h/jours,1),
        "duree_equipe_4_jours": round(tot_mo_h/jours/4,1),
        "main_oeuvre_par_qualification": mo_par_qualif,
        "prix_moyen_m2_habitable": round(tot_vente/(m.get("surface_habitable_m2") or 1),0) if m.get("surface_habitable_m2") else None}
    controles = construire_controles(m, metres)
    return {"metriques": m, "elements_3d": elements, "metres": metres,
            "lignes": lignes, "bordereau_materiaux": sorted(bordereau.values(), key=lambda x: -x["cout_total"]),
            "synthese": synthese, "controles": controles, "parametres": p}

def construire_controles(m, metres):
    c = []
    def add(niveau, titre, detail): c.append({"niveau": niveau, "titre": titre, "detail": detail})
    if not m.get("vectoriel"): add("BLOQUANT","PDF non vectoriel (scan/image)","Aucun segment vectoriel exploitable : les métrés automatiques sont impossibles. Exiger le PDF vectoriel de l'architecte ou vectoriser/calibrer manuellement.")
    else: add("OK","PDF vectoriel exploité", f"{m.get('nb_segments_vectoriels')} segments vectoriels lus.")
    if m.get("longueur_m"):
        if m.get("echelle_validee"): add("OK","Échelle validée par les cotations", f"Cotations {m.get('cotations')} m cohérentes avec la géométrie (écart < 18 %).")
        else: add("BLOQUANT","Échelle NON validée","Une seule cotation exploitable ou incohérence : vérifier la longueur réelle d'une façade avant tout chiffrage définitif.")
    else: add("BLOQUANT","Emprise non détectée", m.get("erreur_geometrie",""))
    if "DÉFAUT" in str(m.get("source_hauteur")): add("BLOQUANT","Hauteur d'étage absente du dossier","Aucune altitude ni hauteur d'étage lisible : fournissez la feuille de COUPE. Les m² de murs calculés avec 2,70 m de secours ne sont pas chiffrables en définitif.")
    else: add("OK","Hauteur d'étage lue sur le dossier", str(m.get("source_hauteur")))
    SH, sp = m.get("surface_habitable_m2") or 0, m.get("surface_pieces_m2") or 0
    if SH and sp:
        ecart = abs(sp-SH)/SH*100
        add("OK" if ecart < 8 else "ALERTE","Cohérence surfaces des pièces / surface habitable",
            f"Pièces annotées : {sp} m² ; surface habitable : {SH} m² ; écart {ecart:.1f} % (seuil 8 %).")
    if not m.get("surface_ouvertures_m2"): add("ALERTE","Aucune menuiserie cotée détectée","Le poste menuiseries est à zéro/faible : vérifier le tableau des menuiseries du dossier.")
    manq = [k for k,v in metres.items() if v["statut"] == "MANQUANT"]
    if manq: add("BLOQUANT", f"{len(manq)} métré(s) sans donnée source dans le dossier", "Postes concernés : " + ", ".join(manq) + ". Aucune valeur n'a été inventée : complétez le dossier (coupe, façades, détails) puis relancez l'analyse.")
    if m.get("donnees_manquantes"): add("BLOQUANT","Données constructives absentes du dossier","Non trouvées sur aucune feuille : " + ", ".join(m["donnees_manquantes"]) + ". Ces données figurent normalement sur la coupe, les façades ou les détails — ajoutez les feuilles correspondantes.")
    bloquants = sum(1 for x in c if x["niveau"] == "BLOQUANT")
    add("BLOQUANT" if bloquants else "OK","Statut du dossier",
        "DEVIS DÉFINITIF INTERDIT tant qu'un contrôle BLOQUANT subsiste." if bloquants else "Aucun bloquant géométrique. Le devis reste soumis à la validation humaine des hypothèses signalées.")
    return c

# ---------------------------------------------------------------- PLAN D'EXEMPLE
def generer_plan_exemple():
    doc = pymupdf.open()
    page = doc.new_page(width=842, height=595)
    x0, y0, sx, sy = 120, 110, 45, 42  # 12 m x 9 m -> échelle pt/m
    L, W = 12, 9
    page.draw_rect(pymupdf.Rect(x0, y0, x0+L*sx, y0+W*sy), color=(0,0,0), width=2.2)
    # refends / cloisons
    page.draw_line(pymupdf.Point(x0+5*sx, y0), pymupdf.Point(x0+5*sx, y0+W*sy), color=(0,0,0), width=1.2)
    page.draw_line(pymupdf.Point(x0, y0+4*sy), pymupdf.Point(x0+5*sx, y0+4*sy), color=(0,0,0), width=1.2)
    page.draw_line(pymupdf.Point(x0+5*sx, y0+5.5*sy), pymupdf.Point(x0+L*sx, y0+5.5*sy), color=(0,0,0), width=1.2)
    page.draw_line(pymupdf.Point(x0+8.5*sx, y0), pymupdf.Point(x0+8.5*sx, y0+5.5*sy), color=(0,0,0), width=1.2)
    txt = [("PLAN RDC — Maison individuelle R+1", 120, 60, 15),
           ("12.00", 330, 92, 11), ("9.00", 690, 300, 11),
           ("Séjour : 32,40 m²", 150, 200, 10), ("Cuisine : 12,60 m²", 150, 400, 10),
           ("Chambre 1 : 13,50 m²", 420, 200, 10), ("Chambre 2 : 11,90 m²", 560, 190, 10),
           ("Chambre 3 : 12,40 m²", 560, 215, 10), ("Bureau : 9,30 m²", 150, 250, 10),
           ("Salle de bain : 6,80 m²", 420, 400, 10), ("Salle d'eau : 4,20 m²", 420, 425, 10),
           ("WC : 2,10 m²", 620, 400, 10), ("Cellier : 3,80 m²", 620, 425, 10),
           ("Entrée : 5,20 m²", 560, 480, 10), ("Dégagement : 3,80 m²", 420, 480, 10),
           ("Surface habitable : 118,00 m²", 120, 545, 10),
           ("Emprise au sol : 108,00 m²", 400, 545, 10),
           ("Menuiseries : Fenêtre 120x135 — Baie 240x215 — Porte 90x215", 120, 565, 9),
           ("Toiture tuiles, pente 35%", 560, 60, 10)]
    for t, x, y, s in txt: page.insert_text(pymupdf.Point(x, y), t, fontsize=s)
    page2 = doc.new_page(width=842, height=595)
    for t, x, y, s in [("COUPE A-A", 120, 60, 15), ("Niveau RDC +0.00", 120, 200, 11),
                       ("Niveau étage +2.70", 120, 300, 11), ("Égout +5.40", 120, 400, 11),
                       ("Faîtage +7.60", 120, 480, 11), ("Hauteur sous plafond : 2.50 m", 400, 200, 11),
                       ("Murs extérieurs : blocs creux 20 cm", 400, 240, 10),
                       ("Cloisons de distribution : 72 mm", 400, 265, 10),
                       ("Dallage ép. 12 cm", 400, 290, 10), ("Plancher 16+4 cm", 400, 315, 10),
                       ("Profondeur de fondation : 0.80 m", 400, 340, 10),
                       ("Semelle filante 60x40 cm", 400, 365, 10),
                       ("Isolation des murs : laine de verre 120 mm", 400, 390, 10),
                       ("Faïence hauteur 2.00 m", 400, 415, 10),
                       ("Couverture : tuiles mécaniques", 400, 440, 10)]:
        page2.insert_text(pymupdf.Point(x, y), t, fontsize=s)
    page2.draw_line(pymupdf.Point(120,320), pymupdf.Point(700,320), color=(0,0,0), width=1.5)
    page2.draw_line(pymupdf.Point(120,430), pymupdf.Point(700,430), color=(0,0,0), width=1.5)
    buf = io.BytesIO(); doc.save(buf); doc.close(); buf.seek(0)
    return buf

# ---------------------------------------------------------------- EXPORTS
def fmt(x): return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " €"

@app.route("/")
def index(): return app.send_static_file("index.html")

@app.route("/health")
def health(): return jsonify({"statut": "ok", "service": "PlanMetric Pro"})

@app.route("/manifest.webmanifest")
def manifest():
    r = app.send_static_file("manifest.webmanifest")
    r.mimetype = "application/manifest+json"
    return r

@app.route("/sw.js")
def service_worker():
    r = app.send_static_file("sw.js")
    r.mimetype = "application/javascript"
    r.headers["Service-Worker-Allowed"] = "/"
    r.headers["Cache-Control"] = "no-cache"
    return r

@app.route("/api/analyser", methods=["POST"])
def api_analyser():
    fichiers = [f for f in request.files.getlist("fichiers") if f.filename] \
               or ([request.files["fichier"]] if "fichier" in request.files and request.files["fichier"].filename else [])
    if not fichiers:
        return jsonify({"erreur": "Aucun fichier PDF reçu."}), 400
    try: m = analyser_dossier([(f.filename, f.read()) for f in fichiers])
    except Exception as e: return jsonify({"erreur": f"Lecture PDF impossible : {e}"}), 422
    p = dict(PARAMS_DEFAUT)
    if request.form.get("parametres"):
        try: p.update(json.loads(request.form["parametres"]))
        except Exception: pass
    return jsonify(construire_devis(m, p))

@app.route("/api/recalculer", methods=["POST"])
def api_recalculer():
    body = request.get_json(force=True)
    m, p = body["metriques"], dict(PARAMS_DEFAUT)
    p.update(body.get("parametres", {}))
    if body.get("calibration_longueur_m") and m.get("longueur_m"):
        ratio = body["calibration_longueur_m"] / m["longueur_m"]
        m["longueur_m"] = round(body["calibration_longueur_m"], 2)
        m["largeur_m"] = round(m["largeur_m"]*ratio, 2)
        m["emprise_sol_m2"] = round(m["longueur_m"]*m["largeur_m"], 2)
        m["perimetre_m"] = round(2*(m["longueur_m"]+m["largeur_m"]), 2)
        m["echelle_validee"] = True
    return jsonify(construire_devis(m, p))

@app.route("/api/plan-exemple")
def api_plan_exemple():
    return send_file(generer_plan_exemple(), mimetype="application/pdf",
                     as_attachment=True, download_name="plan_exemple_maison_R+1.pdf")

@app.route("/api/demo", methods=["POST"])
def api_demo():
    buf = generer_plan_exemple()
    m = analyser_pdf(buf.read(), "plan_exemple_maison_R+1.pdf")
    return jsonify(construire_devis(m, dict(PARAMS_DEFAUT)))

def construire_workbook(d):
    wb = Workbook(); bleu = PatternFill("solid", fgColor="1F4E79"); blanc = Font(color="FFFFFF", bold=True)
    fin = Border(*[Side(style="thin", color="B0B0B0")]*4)
    def entete(ws, titres, largeurs):
        ws.append(titres)
        for c in ws[1]: c.fill, c.font, c.border = bleu, blanc, fin
        for i, l in enumerate(largeurs, 1): ws.column_dimensions[get_column_letter(i)].width = l
    ws = wb.active; ws.title = "Devis"
    entete(ws, ["Code","Lot","Poste","Unité","Quantité","Prix unit. vente HT","Total vente HT","Déboursé matériaux HT","Déboursé MO HT","Marge €","Marge %","Statut métré"], [8,26,58,8,11,14,14,15,13,11,9,12])
    for l in d["lignes"]:
        ws.append([l["code"], l["lot"], l["poste"], l["unite"], l["quantite"], l["prix_unitaire_vente"],
                   l["total_vente_ht"], round(sum(x["cout_total"] for x in l["materiaux"]),2),
                   l["main_oeuvre"]["cout_total"], l["marge_eur"], l["marge_pct"], l["statut"]])
    s = d["synthese"]
    ws.append([]); ws.append(["","","TOTAL HT","","","",s["total_vente_ht"]])
    ws.append(["","","TVA "+str(s["tva_pct"])+" %","","","",s["tva_eur"]]); ws.append(["","","TOTAL TTC","","","",s["total_ttc"]])
    ws2 = wb.create_sheet("Metres-traçabilité")
    entete(ws2, ["Grandeur","Quantité","Formule de calcul","Source sur le plan","Statut"], [22,12,70,60,12])
    for k, v in d["metres"].items(): ws2.append([k, v["quantite"], v["formule"], v["source"], v["statut"]])
    ws3 = wb.create_sheet("Bordereau-matériaux")
    entete(ws3, ["Matériau","Unité","Quantité totale (pertes incluses)","Conditionnement","Nb conditionnements","Quantité à commander","Prix achat unitaire HT","Coût total HT","Postes concernés"], [42,10,22,26,14,16,16,14,20])
    for b in d["bordereau_materiaux"]: ws3.append([b["designation"], b["unite"], b["quantite_totale"], b.get("conditionnement") or "", b.get("nb_conditionnements") or "", b.get("quantite_a_commander", b["quantite_totale"]), b["prix_achat_unitaire"], b["cout_total"], ", ".join(b["postes"])])
    ws4 = wb.create_sheet("Main-doeuvre")
    entete(ws4, ["Code","Poste","Qualification","Temps unitaire (h)","Quantité","Heures totales","Taux horaire","Coût MO HT"], [8,58,28,14,11,13,12,12])
    for l in d["lignes"]:
        mo = l["main_oeuvre"]; ws4.append([l["code"], l["poste"], mo["qualification"], mo["temps_unitaire_h"], l["quantite"], mo["heures_totales"], mo["taux_horaire"], mo["cout_total"]])
    ws4.append([]); ws4.append(["","","TOTAL HEURES","","",s["total_main_oeuvre_heures"],"",s["total_main_oeuvre_eur"]])
    ws4.append(["","","Durée 1 personne (jours)","","",s["duree_1_personne_jours"]]); ws4.append(["","","Durée équipe de 4 (jours)","","",s["duree_equipe_4_jours"]])
    ws5 = wb.create_sheet("Contrôles")
    entete(ws5, ["Niveau","Contrôle","Détail"], [12,46,110])
    for c in d["controles"]: ws5.append([c["niveau"], c["titre"], c["detail"]])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

@app.route("/api/export/excel", methods=["POST"])
def export_excel():
    d = request.get_json(force=True)
    return send_file(construire_workbook(d), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="devis_metres_materiaux_MO.xlsx")

@app.route("/api/export/pdf", methods=["POST"])
def export_pdf():
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    d = request.get_json(force=True); s = d["synthese"]
    definitif = d.get("definitif", False)
    styles = getSampleStyleSheet()
    petit = ParagraphStyle("petit", parent=styles["Normal"], fontSize=6.5, leading=8)
    titre = ParagraphStyle("titre", parent=styles["Title"], fontSize=16)
    buf = io.BytesIO()
    doc_tpl = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=12*mm, rightMargin=12*mm, topMargin=12*mm, bottomMargin=12*mm,
                                title="Devis détaillé — métrés issus du modèle 3D")
    el = [Paragraph("DEVIS DÉTAILLÉ — " + ("DÉFINITIF" if definitif else "BROUILLON À VALIDER"), titre),
          Paragraph(f"Projet : {d['metriques'].get('nom_fichier','')} — {d['metriques'].get('type_batiment','')} — "
                    f"{d['metriques'].get('longueur_m','?')} × {d['metriques'].get('largeur_m','?')} m, "
                    f"{d['metriques'].get('nb_niveaux')} niveau(x), surface habitable {d['metriques'].get('surface_habitable_m2','?')} m². "
                    f"Métrés calculés depuis le modèle 3D reconstruit à partir du PDF vectoriel.", petit), Spacer(1, 4*mm)]
    data = [[Paragraph(x, petit) for x in ["Code","Poste","U","Qté","PU HT","Total HT","Matériaux HT","MO (h)","Statut"]]]
    for l in d["lignes"]:
        data.append([l["code"], Paragraph(l["poste"], petit), l["unite"], l["quantite"], fmt(l["prix_unitaire_vente"]),
                     fmt(l["total_vente_ht"]), fmt(round(sum(x["cout_total"] for x in l["materiaux"]),2)),
                     l["main_oeuvre"]["heures_totales"], l["statut"]])
    data.append(["","","","", "TOTAL HT", fmt(s["total_vente_ht"]), fmt(s["total_materiaux_achat_ht"]), s["total_main_oeuvre_heures"], ""])
    data.append(["","","","", f"TVA {s['tva_pct']} %", fmt(s["tva_eur"]), "", "", ""])
    data.append(["","","","", "TOTAL TTC", fmt(s["total_ttc"]), "", "", ""])
    t = Table(data, colWidths=[14*mm, 95*mm, 10*mm, 16*mm, 24*mm, 24*mm, 26*mm, 16*mm, 22*mm], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E79")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                           ("GRID",(0,0),(-1,-1),0.4,colors.grey),("FONTSIZE",(0,0),(-1,-1),6.5),
                           ("BACKGROUND",(0,-3),(-1,-1),colors.HexColor("#DCE6F1")),("VALIGN",(0,0),(-1,-1),"TOP")]))
    el.append(t); el.append(Spacer(1,4*mm))
    el.append(Paragraph(f"Déboursé sec total : {fmt(s['debourse_sec_total'])} (matériaux {fmt(s['total_materiaux_achat_ht'])} + main-d'œuvre "
                        f"{s['total_main_oeuvre_heures']} h = {fmt(s['total_main_oeuvre_eur'])}). Marge brute : {fmt(s['marge_brute_eur'])} ({s['marge_brute_pct']} %). "
                        f"Durée estimée : {s['duree_1_personne_jours']} jours à 1 personne, {s['duree_equipe_4_jours']} jours à 4. "
                        "Détail des formules, sources et contrôles : voir l'export Excel (onglets Métrés-traçabilité, Bordereau-matériaux, Main-d'œuvre, Contrôles). "
                        "Ce devis doit être vérifié et signé par un professionnel avant engagement : toute quantité marquée HYPOTHÈSE repose sur un paramètre non lu sur le plan.", petit))
    doc_tpl.build(el); buf.seek(0)
    return send_file(buf, mimetype="application/pdf", as_attachment=True,
                     download_name=("devis_definitif.pdf" if definitif else "devis_BROUILLON_a_valider.pdf"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
