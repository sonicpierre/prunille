"""
Script de lecture du fichier data_handmade.xlsx
Lit toutes les lignes de données (hors titres) de la sheet "Base emb (FT)"
et produit un fichier JSON avec un élément par ligne.

Structure du fichier :
  - Ligne 4 : titres des colonnes (à partir de la colonne B)
  - Lignes 5+ : données
"""

import pandas as pd
import json
import os
import sys


def get_workbook_path() -> str:
    """Retourne le chemin absolu vers data_handmade.xlsx."""
    return os.path.abspath(
        os.path.join(os.path.dirname(__file__), '..', 'input', 'data_handmade.xlsx')
    )


SHEET_NAME = "Base emb (FT)"
HEADER_ROW_INDEX = 3  # Ligne 4 en index 0-based = 3


def get_output_path() -> str:
    """Retourne le chemin absolu vers le fichier JSON de sortie."""
    return os.path.abspath(
        os.path.join(os.path.dirname(__file__), '..', 'input', 'output_inge_pack', 'data_handmade.json')
    )


def read_all_data_rows(filepath: str) -> list[dict]:
    """
    Ouvre le fichier Excel et retourne toutes les lignes de données
    sous forme d'une liste de dictionnaires {titre_colonne: valeur}.

    Les titres sont lus en ligne 4, les données à partir de la ligne 5,
    à partir de la colonne B.
    """
    df = pd.read_excel(
        filepath,
        sheet_name=SHEET_NAME,
        header=HEADER_ROW_INDEX,  # Ligne 4 = en-têtes (index 3)
    )

    # Supprimer la première colonne (colonne A, généralement vide)
    if str(df.columns[0]).startswith('Unnamed'):
        df = df.iloc[:, 1:]

    # Supprimer les colonnes entièrement sans titre (Unnamed)
    df = df.loc[:, ~df.columns.astype(str).str.startswith('Unnamed')]

    # Supprimer les lignes entièrement vides
    df = df.dropna(how='all')

    # Convertir chaque ligne en dictionnaire
    rows = df.to_dict(orient='records')

    return rows


def serialize_value(val):
    """Convertit les valeurs non sérialisables en JSON (NaT, Timestamp, NaN)."""
    if pd.isna(val):
        return None
    if isinstance(val, pd.Timestamp):
        return val.isoformat()
    return val


def main():
    filepath = get_workbook_path()
    output_path = get_output_path()
    print(f"Lecture du fichier : {os.path.basename(filepath)}")
    print(f"Sheet : {SHEET_NAME}\n")

    if not os.path.exists(filepath):
        print(f"ERREUR : Le fichier n'existe pas : {filepath}")
        sys.exit(1)

    try:
        rows = read_all_data_rows(filepath)
    except PermissionError:
        print("ERREUR : Le fichier est ouvert dans un autre programme (Excel ?).")
        print("Veuillez le fermer et relancer le script.")
        sys.exit(1)

    # Nettoyer les valeurs non sérialisables
    clean_rows = [
        {k: serialize_value(v) for k, v in row.items()}
        for row in rows
    ]

    # Écrire le JSON
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(clean_rows, f, ensure_ascii=False, indent=2)

    print(f"{len(clean_rows)} lignes exportées vers :")
    print(f"  {output_path}")


if __name__ == "__main__":
    main()
