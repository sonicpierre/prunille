"""
Script de génération des fiches techniques Word à partir du JSON.
Utilise docxtpl pour injecter les données dans les templates Word.

Usage:
    python src/generate_fiches.py [chemin_dossier_input]
"""

import json
import os
import sys
import math
from datetime import date
from docxtpl import DocxTemplate, RichText

COULEUR = "0B0BF4"

def get_input_dir() -> str:
    """
    Retourne le chemin absolu vers le dossier input.
    Utilise le premier argument de la ligne de commande s'il est fourni,
    sinon ../input relatif au script (ou à l'exécutable PyInstaller).
    """
    if len(sys.argv) > 1:
        return os.path.abspath(sys.argv[1])

    if getattr(sys, 'frozen', False):
        # Exécutable PyInstaller : dossier de l'exe
        base_path = os.path.dirname(sys.executable)
    else:
        # Script normal (.py) : dossier du script
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.abspath(os.path.join(base_path, '..', 'input'))

# ---------------------------------------------------------------------------
# Chemins
# ---------------------------------------------------------------------------

INPUT_DIR = get_input_dir()
TEMPLATE_DIR = os.path.join(INPUT_DIR, 'template')
OUTPUT_DIR = os.path.join(INPUT_DIR, 'fiche_technique')
JSON_PATH = os.path.join(INPUT_DIR, 'data_handmade.json')

TEMPLATE_SACHET = os.path.join(TEMPLATE_DIR, 'Fic Tech Log Sachet Preforme.docx')
TEMPLATE_FILM = os.path.join(TEMPLATE_DIR, 'Fic Tech Log Film.docx')


# ---------------------------------------------------------------------------
# Mapping JSON → Jinja2 variable names
# ---------------------------------------------------------------------------

FIELD_MAP = {
    'Code article': 'code_article',
    'Fournisseur': 'fournisseur',
    'Métrage bobine': 'metrage_bobine',
    'Diamètre mandrin intérieur': 'interieur_mandrin',
    'sens déroulement': 'sens_deroulement',
    'Nb couleurs sans vernis': 'nb_couleurs_sans_vernis',
    'Vernis': 'vernis',
    'Procédé impression': 'procede_impression',
    'Impression extérieur ou cello email': 'impression_exterieur_ou_cello_email',
    'Hauteur/Laize': 'hauteur_laize',
    'Largeur/Coupe': 'largeur_coupe',
    'Soufflet': 'soufflet',
    'Masse sachet à vide (g)': 'masse_sachet',
    'Zip': 'zip',
    'Position zip VS haut du sachet': 'position_zip',
    'Encoche ouverture': 'encoche_ouverture',
    'Position encoche VS haut du sachet': 'Position_encoche',
    'Type soudure fond': 'type_soudure_fond',
    'Angle arrondis': 'angle_arrondis',
    'Trou Européen': 'trou_europeen',
    'Produit': 'Produit',
    'Rapport surface/volume': 'Rapport_surface_volume',
}



# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def format_value(val):
    """Formate une valeur pour l'affichage dans le template."""
    if val is None:
        return ''
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val)


def build_nom_article(record: dict) -> str:
    """Construit le nom de l'article : Client + Type sachet + Grammage + Produit."""
    parts = [
        format_value(record.get('Client')),
        format_value(record.get('Type de sachet')),
        format_value(record.get('Grammage')),
        format_value(record.get('Produit')),
    ]
    return ' '.join(p for p in parts if p)

def build_ph(record: dict) -> str:
    """PH of the product"""
    if record.get('Acide\nPH<4.5')=="Oui":
        return "oui (pH≤4.5)"
    else:
        return "non (pH>4.5)"

def build_pasteurisation(record: dict) -> str:
    """Pasteurisation of the product"""
    if record.get('Pasteurisation')=="Oui":
        return "oui (6 heures à 80 °C)"
    else:
        return "non"
    
def is_missing(val) -> bool:
    return val is None or (isinstance(val, float) and math.isnan(val)) or val == ''

def build_vernis(record: dict) -> str:
    if not is_missing(record.get('Vernis')):
        return str(int(record.get('Nb couleurs sans vernis')))+ " + Vernis " + record.get('Vernis')
    else:
        return record.get('Nb couleurs sans vernis')



def build_context(record: dict) -> dict:
    """Transforme un enregistrement JSON en contexte Jinja2 pour le template."""
    context = {}
    for json_key, tpl_var in FIELD_MAP.items():
        context[tpl_var] = format_value(record.get(json_key))

    # Champ calculé
    context['nom_article'] = build_nom_article(record)

    context['ph'] = build_ph(record)

    context['Pasteurisation'] = build_pasteurisation(record)

    context['nb_couleurs_sans_vernis']= build_vernis(record)

    # Décomposition de Structure en mat_1, mat_2, mat_3
    structure = record.get('Structure') or ''
    parts = [p.strip() for p in structure.split('/')] if structure else []
    context['mat_1'] = parts[0] if len(parts) > 0 else ''
    context['mat_2'] = parts[1] if len(parts) > 1 else ''
    context['mat_3'] = parts[2] if len(parts) > 2 else ''

    # Date du jour
    context['date'] = date.today().strftime('%d/%m/%Y')

    # Tous les textes insérés en couleur et en gras (nécessite {{r var }} dans le template)
    return {k: RichText(v, color=COULEUR, bold=True) for k, v in context.items()}


def get_template_info(record: dict) -> tuple[str, str] | None:
    """
    Retourne (chemin_template, nom_template) selon le type de sachet.
    Retourne None si le record doit être ignoré.
    """
    type_sachet = record.get('Type de sachet')
    code_article = record.get('Code article')
    if type_sachet is None or code_article is None:
        return None

    if type_sachet.startswith('P'):
        return TEMPLATE_SACHET, 'Sachet Préformé'
    else:
        return TEMPLATE_FILM, 'Film'


def generate_fiche(record: dict, output_dir: str) -> str | None:
    """
    Génère un fichier Word pour un enregistrement.
    Retourne le chemin du fichier créé, ou None si ignoré.
    """
    template_info = get_template_info(record)
    if template_info is None:
        return None

    template_path, template_name = template_info
    context = build_context(record)
    code = record.get('Code article', 'UNKNOWN')

    # Charger le template et rendre
    tpl = DocxTemplate(template_path)
    tpl.render(context, autoescape=True)

    # Nom du fichier de sortie
    filename = f"{code}_Fic Tech Log {template_name}.docx"
    output_path = os.path.join(output_dir, filename)
    tpl.save(output_path)

    return output_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print(f"Lecture du JSON : {JSON_PATH}")

    if not os.path.exists(JSON_PATH):
        print(f"ERREUR : Fichier introuvable : {JSON_PATH}")
        sys.exit(1)

    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        records = json.load(f)

    print(f"{len(records)} enregistrements chargés")

    # Créer le dossier de sortie
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    generated = 0
    skipped = 0

    for record in records:
        try:
            result = generate_fiche(record, OUTPUT_DIR)
            if result:
                code = record.get('Code article', '?')
                print(f"  ✓ {code} → {os.path.basename(result)}")
                generated += 1
            else:
                skipped += 1
        except PermissionError:
            code = record.get('Code article', '?')
            print(f"  ✗ {code} → ERREUR : fichier verrouillé (ouvert dans Word ?)")
            skipped += 1

    print(f"\nTerminé : {generated} fiches générées, {skipped} ignorées")
    print(f"Dossier de sortie : {OUTPUT_DIR}")


if __name__ == '__main__':
    main()
