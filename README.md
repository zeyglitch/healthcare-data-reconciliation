# Contrôle d'exhaustivité PMSI — MCO (Orbis ↔ Hexagone)

Outil Python d'automatisation du contrôle croisé entre les exports du dossier patient (Orbis) et le logiciel administratif (Hexagone), pour un Département d'Information Médicale (DIM) hospitalier — activité de Médecine, Chirurgie et Obstétrique (MCO).

## Contexte

Développé pendant un stage de 3ème année d'école d'ingénieur (ISIS Castres, partenaire INSA) au Département d'Information Médicale du Centre Hospitalier de Cahors, entre avril et juin 2026. Le concept général de l'applicatif a été défini par le DIM du Centre Hospitalier de Cahors ; le code présenté ici en est la réalisation technique, publiée avec leur accord.

Le contrôle d'exhaustivité mensuel — vérifier que chaque séjour identifié dans Orbis est bien facturé dans Hexagone, et inversement — était jusqu'alors réalisé manuellement via une série de RECHERCHEV Excel, à raison d'environ **une journée de travail par mois**.

## Résultats

- Traitement croisé de **7 fichiers d'export** pour l'activité MCO.
- Jusqu'à **45 000 lignes traitées par mois**.
- Contrôle mensuel ramené de ~1 jour à quelques minutes.
- **Validé et déployé en production** sur les postes des codeurs du DIM.
- Distribution sous forme d'exécutable Windows autonome (`.exe`) via PyInstaller, pour utilisation sans installation Python.

## Fonctionnalités

- **Analyse croisée automatique** entre exports Orbis et Hexagone.
- **Tri intelligent** : séparation des hospitalisations classiques (MCO) et des séances (chimiothérapie, dialyse, radiothérapie).
- **Rapports Excel formatés** : mise en couleur des anomalies, filtres actifs, colonnes prêtes à l'exploitation par le codeur.
- **Interface graphique** simple (Tkinter) pour sélection des fichiers et lancement du traitement.

## Stack technique

- Python 3.9+
- Pandas (traitement des données)
- openpyxl (génération Excel avec mise en forme)
- Tkinter (interface graphique)
- PyInstaller (distribution `.exe`)

## Confidentialité et publication

- **Publication autorisée** par le maître de stage.
- Code **entièrement générique** : aucun élément de configuration ni identifiant du Centre Hospitalier n'est présent dans le code source.
- Les jeux de données inclus dans `data_test/` sont **entièrement synthétiques** — aucun IPP, NDA, NIP, ni aucune donnée patient réelle. Générés pour la démonstration.
- Le nom exact de l'établissement n'est intentionnellement pas mentionné.

## Structure du projet

```
├── src/                  # Scripts Python (logique métier + interface)
├── data_test/
│   └── import_test/      # Fichiers d'essai synthétiques
├── docs/                 # Guide développeur
├── mode_logiciel/        # Mode d'emploi utilisateur
├── requirements.txt      # Dépendances Python
└── .gitignore
```

## Installation

```bash
python -m venv venv
# Windows PowerShell :
venv\Scripts\Activate.ps1
# macOS / Linux :
source venv/bin/activate

pip install -r requirements.txt
```

## Utilisation

### Interface graphique

```bash
python src/interface.py
```

Sélectionner les fichiers d'export Orbis et Hexagone, indiquer le dossier de sortie, cliquer sur **Lancer**.

### Génération d'un exécutable Windows autonome

Pour distribuer l'outil à un utilisateur qui n'a pas Python installé :

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name "Conciliation_DIM" src/interface.py
```

L'exécutable `Conciliation_DIM.exe` est généré dans `dist/`.

## Licence

MIT — voir [LICENSE](LICENSE).
