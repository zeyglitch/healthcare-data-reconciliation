import pandas as pd
import numpy as np
import logging
from datetime import datetime
from pathlib import Path
import sys

# Configuration des dossiers (Chemins relatifs ou absolus à adapter)
DOSSIER_IMPORT = Path(__file__).parent
DOSSIER_EXPORT = Path(__file__).parent

# Création des dossiers s'ils n'existent pas
DOSSIER_IMPORT.mkdir(parents=True, exist_ok=True)
DOSSIER_EXPORT.mkdir(parents=True, exist_ok=True)

# Configuration du système de journalisation (Logging)
logging.basicConfig(
    filename=DOSSIER_EXPORT / f'conciliation_DPI_DIM_{datetime.now().strftime("%Y%m%d")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def trouver_et_charger_fichier(dossier, mot_cle, dtype_dict, colonnes_dates):
    """
    Scanne un répertoire pour trouver un fichier par mot-clé, détecte son format (.csv ou .xlsx)
    et le charge dans un DataFrame Pandas avec les contraintes de typage.
    """
    # Recherche tous les fichiers dans le dossier contenant le mot-clé (insensible à la casse)
    fichiers_trouves = [f for f in dossier.iterdir() if f.is_file() and mot_cle.lower() in f.name.lower() and not f.name.startswith('~')]
    
    if len(fichiers_trouves) == 0:
        raise FileNotFoundError(f"Aucun fichier contenant le mot-clé '{mot_cle}' n'a été trouvé dans le dossier {dossier}.")
    elif len(fichiers_trouves) > 1:
        noms = [f.name for f in fichiers_trouves]
        raise ValueError(f"Conflit : Plusieurs fichiers trouvés pour '{mot_cle}' : {noms}. Veuillez nettoyer le dossier d'import.")
    
    fichier = fichiers_trouves[0]
    logging.info(f"Fichier identifié pour {mot_cle} : {fichier.name}")
    
    # Chargement conditionnel basé sur l'extension
    extension = fichier.suffix.lower()
    
    if extension == '.csv':
        # Hypothèse d'un encodage standard et d'un séparateur classique (à ajuster si nécessaire)
        return pd.read_csv(fichier, sep=',', dtype=dtype_dict, parse_dates=colonnes_dates, encoding='utf-8')
    elif extension in ['.xlsx', '.xls']:
        return pd.read_excel(fichier, dtype=dtype_dict, parse_dates=colonnes_dates)
    else:
        raise TypeError(f"Le fichier {fichier.name} a une extension non supportée ({extension}). Seuls CSV et Excel sont acceptés.")

def clean_keys_strict(df, key_columns, expected_length=10):
    for col in key_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.zfill(expected_length)
    return df

def validate_and_map_columns(df, source_name, mapping_dict):
    missing_cols = [col for col in mapping_dict.keys() if col not in df.columns]
    if missing_cols:
        raise KeyError(f"Structure invalide pour {source_name}. Colonnes manquantes : {missing_cols}")
    return df.rename(columns=mapping_dict)

def extract_unique_stays(df, source_name, date_col=None):
    cols_to_keep = ['IPP', 'IEP']
    
    if date_col and date_col in df.columns:
        cols_to_keep.append(date_col)
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
        df_unique = df[cols_to_keep].groupby(['IPP', 'IEP'], as_index=False).min()
        df_unique.rename(columns={date_col: f'Date_Ref_{source_name}'}, inplace=True)
    else:
        df_unique = df[['IPP', 'IEP']].drop_duplicates()
        df_unique[f'Date_Ref_{source_name}'] = np.nan

    df_unique[f'Source_{source_name}'] = True
    return df_unique

def rapprocher_fichiers():
    logging.info("--- Début du cycle de conciliation ---")
    
    dtype_hexa = {'IPP': str, 'NDA': str}
    dtype_orbis = {'IPP_PATIENT': str, 'NUM_SEJOUR': str}
    dtype_dxcare = {'ID_PAT': str, 'NUM_VENUE': str}

    # 1. Extraction dynamique (CSV ou XLSX) basée sur les mots-clés
    try:
        df_hexa = trouver_et_charger_fichier(DOSSIER_IMPORT, "hexa", dtype_hexa, ['DATE_ADM'])
        df_orbis = trouver_et_charger_fichier(DOSSIER_IMPORT, "orbis", dtype_orbis, ['DATE_DOCUMENT'])
        df_dxcare = trouver_et_charger_fichier(DOSSIER_IMPORT, "dxcare", dtype_dxcare, ['DATE_REALISATION'])
        logging.info(f"Volumétrie brute -> Hexagone: {len(df_hexa)}, Orbis: {len(df_orbis)}, DxCare: {len(df_dxcare)}")
    except Exception as e:
        logging.critical(f"Arrêt du processus. Raison : {str(e)}")
        sys.exit(1) # Arrêt propre du script avec code d'erreur

    # 2. Dictionnaires de Mapping Spécifiques aux Éditeurs
    map_hexa = {'NDA': 'IEP'}
    map_orbis = {'IPP_PATIENT': 'IPP', 'NUM_SEJOUR': 'IEP'}
    map_dxcare = {'ID_PAT': 'IPP', 'NUM_VENUE': 'IEP'}

    # 3. Standardisation et Contrôle d'Intégrité
    try:
        df_hexa = validate_and_map_columns(df_hexa, "Hexagone", map_hexa)
        df_orbis = validate_and_map_columns(df_orbis, "Orbis", map_orbis)
        df_dxcare = validate_and_map_columns(df_dxcare, "DxCare", map_dxcare)
    except KeyError as e:
        logging.critical(e)
        sys.exit(1)

    for df in [df_hexa, df_orbis, df_dxcare]:
        clean_keys_strict(df, ['IPP', 'IEP'], expected_length=10)

    # 4. Aplatissement et conservation du contexte temporel rigoureux
    hexa_keys = extract_unique_stays(df_hexa, "GAM", "DATE_ADM")
    orbis_keys = extract_unique_stays(df_orbis, "Orbis", "DATE_DOCUMENT")
    dxcare_keys = extract_unique_stays(df_dxcare, "DxCare", "DATE_REALISATION")

    # 5. Fusions externes successives
    df_merged = pd.merge(hexa_keys, orbis_keys, on=['IPP', 'IEP'], how='outer')
    df_merged = pd.merge(df_merged, dxcare_keys, on=['IPP', 'IEP'], how='outer')

    for col in ['Source_GAM', 'Source_Orbis', 'Source_DxCare']:
        df_merged[col] = df_merged[col].fillna(False).astype(bool)

    # 6. Moteur de Règles Métier
    conditions = [
        (~df_merged['Source_GAM'] & (df_merged['Source_Orbis'] | df_merged['Source_DxCare'])),
        (df_merged['Source_GAM'] & ~df_merged['Source_Orbis'] & ~df_merged['Source_DxCare']),
        (df_merged['Source_GAM'] & (~df_merged['Source_Orbis'] | ~df_merged['Source_DxCare'])),
        (df_merged['Source_GAM'] & df_merged['Source_Orbis'] & df_merged['Source_DxCare'])
    ]

    statuts = [
        "Alerte : Dossier fantôme GAM",
        "Alerte : Aucune trace clinique",
        "Info : Manque dans un logiciel",
        "OK"
    ]
    
    df_merged['Statut'] = np.select(conditions, statuts, default="[Inconnu]")

    df_merged['Logiciels_trouves'] = np.where(df_merged['Source_GAM'], "Hexagone ", "") + \
                                     np.where(df_merged['Source_Orbis'], "Orbis ", "") + \
                                     np.where(df_merged['Source_DxCare'], "DxCare", "")
    df_merged['Logiciels_trouves'] = df_merged['Logiciels_trouves'].str.strip()

    # 7. Préparation de l'Export
    colonnes_export = ['Statut', 'IPP', 'IEP', 'Logiciels_trouves', 'Date_Ref_GAM', 'Date_Ref_Orbis', 'Date_Ref_DxCare']
    df_final = df_merged[colonnes_export].copy()
    df_final.sort_values(by=['Statut', 'IEP'], inplace=True)
    
    output_file = DOSSIER_EXPORT / f'Rapport_Exhaustivite_{datetime.now().strftime("%Y%m%d")}.xlsx'
    
    try:
        df_final.to_excel(output_file, index=False)
        logging.info(f"Processus terminé avec succès. Lignes : {len(df_final)}. Fichier : {output_file}")
    except PermissionError:
        logging.critical("Échec : Le fichier Excel de destination est ouvert par un autre programme.")
    except Exception as e:
        logging.error(f"Échec de l'écriture Excel : {e}")

if __name__ == "__main__":
    rapprocher_fichiers()