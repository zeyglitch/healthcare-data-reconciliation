import openpyxl
from datetime import datetime, timedelta

def creer_fichier_orbis(nom_fichier, entetes, donnees):
    """Crée le fichier Orbis (sans lignes d'en-tête spéciales)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(entetes)
    for ligne in donnees:
        ws.append(ligne)
    wb.save(nom_fichier)
    print(f"[OK] Fichier Orbis genere avec {len(donnees)} lignes (hors entetes) : {nom_fichier}")

def creer_fichier_hexa(nom_fichier, titre_rapport, entetes, donnees):
    """Crée un fichier Hexagone avec le format réel : titre du rapport (ligne 1), ligne vide (ligne 2), puis colonnes (ligne 3)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([titre_rapport])  # Ligne 1 : titre du rapport
    ws.append([])               # Ligne 2 : vide
    ws.append(entetes)          # Ligne 3 : les vrais en-têtes
    for ligne in donnees:
        ws.append(ligne)
    wb.save(nom_fichier)
    print(f"[OK] Fichier Hexagone genere avec {len(donnees)} lignes (hors entetes) : {nom_fichier}")

# ==========================================
# INITIALISATION DES LISTES DE DONNÉES
# ==========================================
orbis_data = []
hexa_2_data = [] # Hospit Standard
hexa_3_data = [] # Nouveau-né
hexa_4_data = [] # Orthogénie
hexa_5_data = [] # Chimio
hexa_6_data = [] # Dialyse P
hexa_7_data = [] # Hémodialyse

# Variables de base pour itération temporelle
date_base = datetime(2025, 11, 1)

# ==========================================
# 1. GÉNÉRATION DES CAS PARFAITS (MATCH ORBIS <-> HEXAGONE)
# ==========================================

# 1.1 Hospit Standard (Fichier 2) - 25 patients
for i in range(1, 26):
    id_hosp = f"428690{i:03d}"
    nom = f"HOSPIT_OK_{i}"
    orbis_data.append([id_hosp, "05/11/2025", "", "07/01/2026", "", nom, "JEAN", "01/01/1980", "", "", "", "", "05/11/2025", "09/12/2025", "", "Faux", "1800", "", "", "", "", "", "", "pas d'acte", "OK"])
    hexa_2_data.append(["07/01/2026 12:00", f"{nom}/JEAN", nom, "M", "01/01/1980", "pas de bloc", f"H {id_hosp}", "1800", "", "05/11/2025", "", "1800"])

# 1.2 Nouveau-Né (Fichier 3) - 25 patients
for i in range(1, 26):
    id_hosp = f"428691{i:03d}"
    nom = f"BEBE_OK_{i}"
    orbis_data.append([id_hosp, "08/12/2025", "", "15/12/2025", "", nom, "LUC", "08/12/2025", "", "", "", "", "08/12/2025", "15/12/2025", "", "Faux", "2301", "", "", "", "", "", "", "", "OK"])
    hexa_3_data.append(["15/12/2025 10:00", f"{nom}/LUC", nom, "M", "08/12/2025", "nouveau ne", f"N {id_hosp}", "2301", "", "08/12/2025", "", "2301"])

# 1.3 Orthogénie (Fichier 4) - 25 patients
for i in range(1, 26):
    id_hosp = f"428692{i:03d}"
    nom = f"ORTHO_OK_{i}"
    orbis_data.append([id_hosp, "10/12/2025", "", "11/12/2025", "", nom, "MARIE", "07/11/1980", "", "", "", "", "10/12/2025", "11/12/2025", "", "Faux", "1800", "", "", "", "", "", "", "", "OK"])
    hexa_4_data.append(["11/12/2025 16:00", f"{nom}/MARIE", nom, "F", "07/11/1980", "ortho", f"I {id_hosp}", "1800", "", "10/12/2025", "", "1800"])

# 1.4 Séances Chimio (Fichier 5) - 25 patients (UM 1400)
for i in range(1, 26):
    id_hosp = f"426095{i:03d}"
    nom = f"CHIMIO_OK_{i}"
    d_str = (date_base + timedelta(days=i)).strftime("%d/%m/%Y")
    orbis_data.append([id_hosp, d_str, "", d_str, "", nom, "LUCIEN", "25/07/1994", "", "", "", "", d_str, d_str, "", "Faux", "1400", "", "", "", "", "", "", "", "Séance"])
    hexa_5_data.append([f"{nom}/LUCIEN", nom, "M", "25/07/1994", id_hosp, f"{d_str} 09:12", "", "Chimio", "", "1400", "", "", ""])

# 1.5 Séances Dialyse P (Fichier 6) - 25 patients (UM 1401)
for i in range(1, 26):
    id_hosp = f"426096{i:03d}"
    nom = f"DIALYSEP_OK_{i}"
    d_str = (date_base + timedelta(days=i)).strftime("%d/%m/%Y")
    orbis_data.append([id_hosp, d_str, "", d_str, "", nom, "PAUL", "12/12/1985", "", "", "", "", d_str, d_str, "", "Faux", "1401", "", "", "", "", "", "", "", "Séance"])
    hexa_6_data.append([f"{nom}/PAUL", nom, "M", "12/12/1985", id_hosp, f"{d_str} 08:30", "", "Dialyse P", "", "1401", "", "", ""])

# 1.6 Séances Hémodialyse (Fichier 7) - 25 patients (UM 1048)
for i in range(1, 26):
    id_hosp = f"426097{i:03d}"
    nom = f"HEMO_OK_{i}"
    d_str = (date_base + timedelta(days=i)).strftime("%d/%m/%Y")
    orbis_data.append([id_hosp, d_str, "", d_str, "", nom, "MARC", "08/08/1970", "", "", "", "", d_str, d_str, "", "Faux", "1048", "", "", "", "", "", "", "", "Séance"])
    hexa_7_data.append([f"{nom}/MARC", nom, "M", "08/08/1970", id_hosp, f"{d_str} 14:00", "", "Hémodialyse", "", "1048", "", "", ""])


# ==========================================
# 2. INJECTION DES ANOMALIES MULTIPLES
# ==========================================

# -----------------------------------------------
# 2.1 DOSSIER DANS ORBIS UNIQUEMENT (Tri 1 : Hospit)
# -----------------------------------------------
for i in range(1, 11):
    id_hosp = f"500000{i:03d}"
    nom = f"ORBIS_HOSP_ONLY_{i}"
    orbis_data.append([id_hosp, "01/01/2026", "", "05/01/2026", "", nom, "LUC", "10/10/1990", "", "", "", "", "01/01/2026", "05/01/2026", "", "Faux", "1800", "", "", "", "", "", "", "", "Erreur GAM"])

# -----------------------------------------------
# 2.2 DOSSIER DANS ORBIS UNIQUEMENT (Tri 2 : Séances)
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"510000{i:03d}"
    nom = f"ORBIS_SEANCE_ONLY_{i}"
    orbis_data.append([id_hosp, "15/01/2026", "", "15/01/2026", "", nom, "MARC", "10/10/1990", "", "", "", "", "15/01/2026", "15/01/2026", "", "Faux", "1400", "", "", "", "", "", "", "", "Oubli GAM"])

# -----------------------------------------------
# 2.3 DOSSIERS FANTÔMES (Hexagone Uniquement)
# -----------------------------------------------
for i in range(1, 6):
    hexa_2_data.append(["07/01/2026 12:00", f"FANTOME_HOSP_{i}/PIERRE", f"FANTOME_HOSP_{i}", "M", "01/01/1980", "Fantome", f"H 900002{i:03d}", "1800", "", "05/11/2025", "", "1800"])
    hexa_3_data.append(["15/12/2025 10:00", f"FANTOME_BEBE_{i}/LUC", f"FANTOME_BEBE_{i}", "M", "08/12/2025", "Fantome", f"N 900003{i:03d}", "2301", "", "08/12/2025", "", "2301"])
    hexa_4_data.append(["11/12/2025 16:00", f"FANTOME_ORTHO_{i}/MARIE", f"FANTOME_ORTHO_{i}", "F", "07/11/1980", "Fantome", f"I 900004{i:03d}", "1800", "", "10/12/2025", "", "1800"])
    hexa_5_data.append([f"FANTOME_CHIMIO_{i}/JEAN", f"FANTOME_CHIMIO_{i}", "M", "01/01/1999", f"900005{i:03d}", "02/01/2026 10:00", "", "Fantome", "", "1400", "", "", ""])
    hexa_6_data.append([f"FANTOME_DIALYSE_{i}/JEAN", f"FANTOME_DIALYSE_{i}", "M", "01/01/1999", f"900006{i:03d}", "03/01/2026 10:00", "", "Fantome", "", "1401", "", "", ""])
    hexa_7_data.append([f"FANTOME_HEMO_{i}/JEAN", f"FANTOME_HEMO_{i}", "M", "01/01/1999", f"900007{i:03d}", "04/01/2026 10:00", "", "Fantome", "", "1048", "", "", ""])

# -----------------------------------------------
# 2.4 SÉJOURS SANS DATE DE SORTIE — MANQUANTE PARTOUT (Tri 3)
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"990000{i:03d}"
    nom = f"SANS_SORTIE_PARTOUT_{i}"
    orbis_data.append([id_hosp, "15/01/2026", "", "", "", nom, "ALICE", "05/05/1975", "", "", "", "", "15/01/2026", "", "", "Faux", "1800", "", "", "", "", "", "", "", "En cours"])
    hexa_2_data.append(["", f"{nom}/ALICE", nom, "F", "05/05/1975", "Pas de sortie", f"H {id_hosp}", "1800", "", "15/01/2026", "", "1800"])

# -----------------------------------------------
# 2.5 DATE DE SORTIE MANQUANTE DANS ORBIS, PRÉSENTE DANS HEXA HOSPIT (Tri 3)
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"991000{i:03d}"
    nom = f"SORTIE_MANQUE_ORBIS_{i}"
    orbis_data.append([id_hosp, "20/12/2025", "", "", "", nom, "SOPHIE", "15/03/1965", "", "", "", "", "20/12/2025", "", "", "Faux", "1800", "", "", "", "", "", "", "", "A corriger"])
    hexa_2_data.append(["28/12/2025 14:00", f"{nom}/SOPHIE", nom, "F", "15/03/1965", "", f"H {id_hosp}", "1800", "", "20/12/2025", "", "1800"])

# -----------------------------------------------
# 2.6 DATE DE SORTIE MANQUANTE DANS ORBIS, PRÉSENTE DANS HEXA NOUVEAU-NÉ (Tri 3)
# -----------------------------------------------
for i in range(1, 4):
    id_hosp = f"991100{i:03d}"
    nom = f"SORTIE_MANQUE_ORBIS_NN_{i}"
    orbis_data.append([id_hosp, "10/12/2025", "", "", "", nom, "BEBE_LUC", "10/12/2025", "", "", "", "", "10/12/2025", "", "", "Faux", "2301", "", "", "", "", "", "", "", "A corriger"])
    hexa_3_data.append(["18/12/2025 11:00", f"{nom}/BEBE_LUC", nom, "M", "10/12/2025", "", f"N {id_hosp}", "2301", "", "10/12/2025", "", "2301"])

# -----------------------------------------------
# 2.7 DATE DE SORTIE MANQUANTE DANS HEXA, PRÉSENTE DANS ORBIS (Tri 3)
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"992000{i:03d}"
    nom = f"SORTIE_MANQUE_HEXA_{i}"
    orbis_data.append([id_hosp, "05/12/2025", "", "12/12/2025", "", nom, "THOMAS", "22/08/1950", "", "", "", "", "05/12/2025", "12/12/2025", "", "Faux", "1800", "", "", "", "", "", "", "", "A corriger"])
    hexa_2_data.append(["", f"{nom}/THOMAS", nom, "M", "22/08/1950", "", f"H {id_hosp}", "1800", "", "05/12/2025", "", "1800"])

# -----------------------------------------------
# 2.8 SÉANCES MULTIPLES - 1 MANQUANTE DANS ORBIS (Tri 2)
#     Patient avec 4 séances Hexa Chimio, seulement 3 dans Orbis
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"880100{i:03d}"
    nom = f"SEANCE_MANQUE_ORBIS_{i}"
    dates_seances = [(date_base + timedelta(days=7*j+i)).strftime("%d/%m/%Y") for j in range(4)]
    
    # 4 séances dans Hexa Chimio
    for d_str in dates_seances:
        hexa_5_data.append([f"{nom}/PIERRE", nom, "M", "10/03/1975", id_hosp, f"{d_str} 09:00", "", "Chimio", "", "1400", "", "", ""])
    
    # Seulement 3 séances dans Orbis (la 4ème manque)
    for d_str in dates_seances[:3]:
        orbis_data.append([id_hosp, d_str, "", d_str, "", nom, "PIERRE", "10/03/1975", "", "", "", "", d_str, d_str, "", "Faux", "1400", "", "", "", "", "", "", "", "Séance"])

# -----------------------------------------------
# 2.9 SÉANCES MULTIPLES - 1 MANQUANTE DANS HEXA (Tri 2)
#     Patient avec 4 séances Orbis, seulement 3 dans Hexa Dialyse
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"880200{i:03d}"
    nom = f"SEANCE_MANQUE_HEXA_{i}"
    dates_seances = [(date_base + timedelta(days=7*j+i)).strftime("%d/%m/%Y") for j in range(4)]
    
    # 4 séances dans Orbis
    for d_str in dates_seances:
        orbis_data.append([id_hosp, d_str, "", d_str, "", nom, "JULIEN", "18/06/1988", "", "", "", "", d_str, d_str, "", "Faux", "1401", "", "", "", "", "", "", "", "Séance"])
    
    # Seulement 3 séances dans Hexa Dialyse (la 4ème manque)
    for d_str in dates_seances[:3]:
        hexa_6_data.append([f"{nom}/JULIEN", nom, "M", "18/06/1988", id_hosp, f"{d_str} 08:30", "", "Dialyse P", "", "1401", "", "", ""])

# -----------------------------------------------
# 2.10 DÉSYNCHRONISATION DE DATE DE SÉANCE (Tri 2)
#      Même NDA mais dates différentes -> 2 anomalies (1 manquante de chaque côté)
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"880000{i:03d}"
    nom = f"DESYNC_CHIMIO_{i}"
    # Date Orbis = 10/01/2026
    orbis_data.append([id_hosp, "10/01/2026", "", "10/01/2026", "", nom, "PAUL", "10/10/1980", "", "", "", "", "10/01/2026", "10/01/2026", "", "Faux", "1400", "", "", "", "", "", "", "", "Date diff"])
    # Date Hexagone = 11/01/2026
    hexa_5_data.append([f"{nom}/PAUL", nom, "M", "10/10/1980", id_hosp, "11/01/2026 09:00", "", "Decalage", "", "1400", "", "", ""])

# -----------------------------------------------
# 2.11 PATIENTS AVEC UM 1402 (EXCLUSION)
#      Doivent être totalement ignorés par le programme
# -----------------------------------------------
for i in range(1, 6):
    id_hosp = f"770000{i:03d}"
    nom = f"EXCLU_UM1402_{i}"
    orbis_data.append([id_hosp, "01/12/2025", "", "03/12/2025", "", nom, "EVE", "01/01/1960", "", "", "", "", "01/12/2025", "03/12/2025", "", "Faux", "1402", "", "", "", "", "", "", "", ""])

# -----------------------------------------------
# 2.12 DATE DE SORTIE MANQUANTE DANS ORBIS, PRÉSENTE DANS HEXA ORTHOGÉNIE (Tri 3)
# -----------------------------------------------
for i in range(1, 4):
    id_hosp = f"991200{i:03d}"
    nom = f"SORTIE_MANQUE_ORBIS_ORTHO_{i}"
    orbis_data.append([id_hosp, "10/12/2025", "", "", "", nom, "CLAIRE", "07/11/1980", "", "", "", "", "10/12/2025", "", "", "Faux", "1800", "", "", "", "", "", "", "", "A corriger"])
    hexa_4_data.append(["12/12/2025 16:00", f"{nom}/CLAIRE", nom, "F", "07/11/1980", "", f"I {id_hosp}", "1800", "", "10/12/2025", "", "1800"])

# -----------------------------------------------
# 2.13 SÉANCES HÉMODIALYSE - MULTIPLES MANQUANTES DANS HEXA (Tri 2)
#      Patient avec 5 séances Orbis, seulement 2 dans Hexa Hémo
# -----------------------------------------------
for i in range(1, 4):
    id_hosp = f"880300{i:03d}"
    nom = f"HEMO_MANQUE_MULTI_{i}"
    dates_seances = [(date_base + timedelta(days=7*j+i)).strftime("%d/%m/%Y") for j in range(5)]
    
    # 5 séances dans Orbis
    for d_str in dates_seances:
        orbis_data.append([id_hosp, d_str, "", d_str, "", nom, "RENE", "01/01/1945", "", "", "", "", d_str, d_str, "", "Faux", "1048", "", "", "", "", "", "", "", "Séance"])
    
    # Seulement 2 séances dans Hexa Hémodialyse
    for d_str in dates_seances[:2]:
        hexa_7_data.append([f"{nom}/RENE", nom, "M", "01/01/1945", id_hosp, f"{d_str} 14:00", "", "Hémodialyse", "", "1048", "", "", ""])


# ==========================================
# EXPORT DES FICHIERS EXCEL
# ==========================================
orbis_headers = ["N° Hospit", "Entrée le", "Prov", "Sortie le", "Dest", "Nom", "Prénom", "Né(e) le", "Diagnostic prin", "Diag. Relié", "Val. GHS", "Durée", "Entrée le", "Sortie le", "Durée", "Exclu", "UM", "ILP", "Age", "ME", "MS", "Prov.", "Dest.", "Comm. codif. in", "Comm. ctrl. DIM"]
hexa_hospit_headers = ["Date", "Nom/Prénom", "Nom de Naissance", "Sexe", "Date de naissance", "Commentaire", "Dossier", "UF", "Mode", "Date entrée", "Mode", "UF"]
hexa_seance_headers = ["Nom/Prénom", "Nom de Naissance", "Sexe", "Date de naissance", "N° Dossier", "Date", "Mod/Nat", "Commentaire", "Type", "UF", "Héber", "Bâtiment", "Chambre/Lit"]

print("Lancement du moteur de generation des tests...\n")
creer_fichier_orbis("1_Export_Orbis.xlsx", orbis_headers, orbis_data)
creer_fichier_hexa("2_Export_Hexa_Hospit.xlsx", "Liste des sortants", hexa_hospit_headers, hexa_2_data)
creer_fichier_hexa("3_Export_Hexa_NouveauNe.xlsx", "Liste des sortants", hexa_hospit_headers, hexa_3_data)
creer_fichier_hexa("4_Export_Hexa_Orthogenie.xlsx", "Liste des sortants", hexa_hospit_headers, hexa_4_data)
creer_fichier_hexa("5_Export_Hexa_Chimio.xlsx", "Liste des mouvements par séjour", hexa_seance_headers, hexa_5_data)
creer_fichier_hexa("6_Export_Hexa_DialyseP.xlsx", "Liste des mouvements par séjour", hexa_seance_headers, hexa_6_data)
creer_fichier_hexa("7_Export_Hexa_Hemodialyse.xlsx", "Liste des mouvements par séjour", hexa_seance_headers, hexa_7_data)
print(f"\nOperation terminee. {len(orbis_data)} lignes Orbis generees.")