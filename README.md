# 🩺 Conciliation DIM - Aide à l'Exhaustivité

Bienvenue dans l'outil de conciliation dédié au département d'Information Médicale (DIM). Ce petit logiciel a été conçu pour simplifier la vie des codeurs en automatisant la comparaison entre les exports **Orbis** et **Hexagone**.

L'objectif ? Repérer en un clin d'œil les séjours ou séances qui manquent à l'appel et s'assurer que tout est bien facturé !

## 🌟 Ce que fait l'outil pour vous

- **Analyse croisée** : Compare automatiquement vos fichiers exports.
- **Tri intelligent** : Sépare les hospitalisations classiques des séances (Chimio, Dialyse...).
- **Rapports clairs** : Génère des fichiers Excel tout propres, avec des couleurs et des filtres, prêts à être exploités.
- **Gain de temps** : Plus besoin de faire des RECHERCHEV complexes à la main.

## 📁 Comment est organisé le projet ?

Pour que tout soit bien rangé, nous avons séparé les choses :
- `src/` : Contient la "mécanique" du logiciel (les scripts Python).
- `data_test/` : C'est ici que vous pouvez mettre vos fichiers pour faire des essais (`import_test` pour l'entrée, `export_test` pour les résultats).
- `docs/` : Vous y trouverez le **Guide Développeur** si vous avez besoin de modifier le fonctionnement.

## 🛠️ Installation (pour la première fois)

1. Assurez-vous d'avoir **Python** installé.
2. Dans votre terminal, installez les outils nécessaires :
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Comment l'utiliser ?

### Mode Classique (Interface Graphique)
C'est la méthode la plus simple. Lancez cette commande depuis le dossier principal :
```bash
python src/interface.py
```
Une fenêtre s'ouvrira pour vous permettre de choisir vos fichiers et de cliquer sur "Lancer".

### Créer un fichier .exe
Si vous voulez donner l'outil à un collègue qui n'a pas Python, vous pouvez transformer le script en application Windows :
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name "Conciliation_DIM" src/interface.py
```
Le résultat sera dans le dossier `dist/`.
