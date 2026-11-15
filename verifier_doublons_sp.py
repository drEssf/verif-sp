#!/usr/bin/env python3
"""
Vérifie si la synthèse de presse (SP) du jour reprend des articles
déjà présents dans les SP récentes (fichiers Word .docx ou PDF).

Installation (une seule fois) :
    pip install python-docx pypdf

Utilisation :
    python verifier_doublons_sp.py SP_2026-11-15.docx --dossier ./archives_SP --derniers 7
    python verifier_doublons_sp.py SP_du_jour.pdf --dossier ./archives_SP --seuil 0.4 --csv rapport.csv

Options :
    --dossier   dossier contenant les SP précédentes (défaut : dossier du fichier du jour)
    --derniers  nombre de SP récentes à comparer (défaut : 7, triées par date de modification)
    --seuil     similarité de 0 à 1 à partir de laquelle on signale un doublon (défaut : 0.5)
    --csv       exporte le rapport dans un fichier CSV
"""
import argparse
import csv
import re
import sys
import unicodedata
from collections import namedtuple
from pathlib import Path

Article = namedtuple("Article", "source titre texte")


# ---------- 1. Lecture des fichiers ----------

def lire_docx(chemin):
    """Retourne une liste de (texte, est_titre) pour chaque paragraphe."""
    from docx import Document
    blocs = []
    for p in Document(str(chemin)).paragraphs:
        t = p.text.strip()
        if not t:
            continue
        nom_style = (p.style.name if p.style is not None else "").lower()
        est_titre = nom_style.startswith(("heading", "titre", "title"))
        blocs.append((t, est_titre))
    return blocs


def lire_pdf(chemin):
    """Retourne une liste de (ligne, False). Pas de détection de titres en PDF."""
    from pypdf import PdfReader
    texte = "\n".join((page.extract_text() or "") for page in PdfReader(str(chemin)).pages)
    return [(l.strip(), False) for l in texte.splitlines() if l.strip()]


def lire_fichier(chemin):
    ext = chemin.suffix.lower()
    if ext == ".docx":
        return lire_docx(chemin)
    if ext == ".pdf":
        return lire_pdf(chemin)
    raise ValueError(f"Format non géré : {chemin.name}")


# ---------- 2. Découpage en articles ----------

def decouper(blocs, source, taille_fenetre=120):
    """
    Si le document a des titres (styles Titre/Heading dans Word) : 1 article par titre.
    Sinon (PDF, Word sans styles) : fenêtres de ~120 mots.
    """
    if any(est_titre for _, est_titre in blocs):
        articles, titre, tampon = [], None, []
        for texte, est_titre in blocs:
            if est_titre:
                if tampon:
                    articles.append(Article(source, titre or "(sans titre)", " ".join(tampon)))
                titre, tampon = texte, []
            else:
                tampon.append(texte)
        if tampon:
            articles.append(Article(source, titre or "(sans titre)", " ".join(tampon)))
        return articles

    mots = " ".join(t for t, _ in blocs).split()
    return [
        Article(source, " ".join(mots[i:i + 8]) + "…", " ".join(mots[i:i + taille_fenetre]))
        for i in range(0, len(mots), taille_fenetre)
    ]


# ---------- 3. Comparaison ----------

def normaliser(texte):
    """minuscules, sans accents, mots alphanumériques uniquement."""
    texte = unicodedata.normalize("NFD", texte.lower())
    texte = "".join(c for c in texte if unicodedata.category(c) != "Mn")
    return re.findall(r"[a-z0-9]+", texte)


def shingles(mots, n=3):
    """Ensemble de séquences de n mots consécutifs."""
    return {tuple(mots[i:i + n]) for i in range(len(mots) - n + 1)}


def similarite(a, b):
    """Part de la plus petite empreinte retrouvée dans l'autre (0 à 1).
    Détecte aussi un article repris en version raccourcie ou légèrement reformulée."""
    if not a or not b:
        return 0.0
    return len(a & b) / min(len(a), len(b))


def main():
    ap = argparse.ArgumentParser(description="Détecte les articles déjà inclus dans les SP précédentes.")
    ap.add_argument("fichier_du_jour")
    ap.add_argument("--dossier")
    ap.add_argument("--derniers", type=int, default=7)
    ap.add_argument("--seuil", type=float, default=0.5)
    ap.add_argument("--csv")
    args = ap.parse_args()

    f_jour = Path(args.fichier_du_jour).resolve()
    dossier = Path(args.dossier).resolve() if args.dossier else f_jour.parent

    historique = sorted(
        (p for p in dossier.iterdir()
         if p.suffix.lower() in (".docx", ".pdf") and p.resolve() != f_jour and not p.name.startswith("~$")),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )[:args.derniers]

    if not historique:
        sys.exit("Aucune SP précédente trouvée dans " + str(dossier))

    print(f"SP du jour : {f_jour.name}")
    print("Comparée à  : " + ", ".join(p.name for p in historique) + "\n")

    # Indexation des anciennes SP
    anciens = []
    for p in historique:
        try:
            for art in decouper(lire_fichier(p), p.name):
                anciens.append((art, shingles(normaliser(art.texte))))
        except Exception as e:
            print(f"  ! Impossible de lire {p.name} : {e}")

    # Comparaison article par article
    rapport = []
    for art in decouper(lire_fichier(f_jour), f_jour.name):
        sh = shingles(normaliser(art.texte))
        if len(sh) < 5:          # trop court pour être fiable
            continue
        meilleur = max(((similarite(sh, sh_old), old) for old, sh_old in anciens),
                       key=lambda x: x[0], default=(0, None))
        score, old = meilleur
        if score >= args.seuil:
            rapport.append((score, art, old))

    if not rapport:
        print("✅ Aucun doublon détecté.")
    else:
        print(f"⚠️  {len(rapport)} doublon(s) possible(s) :\n")
        for score, art, old in sorted(rapport, key=lambda x: -x[0]):
            print(f"  [{score:.0%}] {art.titre[:90]}")
            print(f"         déjà dans : {old.source}  →  « {old.titre[:70]} »\n")

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["similarite", "article_du_jour", "deja_dans", "titre_ancien"])
            for score, art, old in rapport:
                w.writerow([f"{score:.2f}", art.titre, old.source, old.titre])
        print(f"Rapport CSV : {args.csv}")


if __name__ == "__main__":
    main()
