# Générateur de fiches "Visa de l'Expert-Comptable"

Application de bureau (Tkinter) en 3 étapes :

1. **Importer** un fichier Excel contenant un ou plusieurs tableaux
   `Critères / Choix` (un tableau = un client, empilés dans la même
   feuille) → génère un JSON brut.
2. **Structurer** ce JSON via l'API **Gemini** pour obtenir un objet
   propre `{"client 1": {...}, "client 2": {...}}`.
3. **Générer** un fichier Word par client, à partir du gabarit
   `template.docx` (copie de votre modèle ABC COMPANY, avec les
   informations spécifiques remplacées automatiquement).

## Installation

```bash
pip install -r requirements.txt
```

Nécessite Python 3.9+ avec Tkinter (inclus par défaut sur Windows/macOS ;
sur Linux : `sudo apt install python3-tk`).

## Clé API Gemini

Obtenez une clé sur https://aistudio.google.com/apikey, puis soit :
- collez-la dans le champ "Clé API Gemini" en haut de la fenêtre, soit
- définissez la variable d'environnement `GEMINI_API_KEY` avant de lancer l'app.

## Lancer l'application

```bash
python app.py
```

## Fichiers du dossier

- `app.py` — l'application (interface + logique).
- `template.docx` — le gabarit Word (copie de votre modèle ABC COMPANY
  avec des variables `{{ raison_sociale }}`, `{{ nom_pdg }}`, etc. à la
  place des informations du client). **Doit rester à côté de `app.py`.**
- `requirements.txt` — dépendances Python.

## Format Excel attendu

Chaque tableau doit avoir une ligne d'en-tête contenant "Critères" et
"Choix" (peu importe la colonne exacte), suivie des lignes :

```
Raison sociale               | ...
Adresse de la société        | ...
Nom de PDG                   | ...
Exercice                     | 2026
Trimestre                    | T3
Deux types d'attestations    | Avec retard
Montant                      | ...
Signataire                   | Expert-comptable / Associé
Nom & Prenom                 | ...
Lieu                         | ...
Date de signature            | ...
```

Pour un deuxième client, il suffit d'ajouter un nouveau tableau
`Critères / Choix` plus bas dans la même feuille (avec au moins une
ligne vide entre les deux tableaux).

## Limite actuelle importante

Le gabarit `template.docx` ne couvre que le texte du modèle que vous
avez fourni (cas "**Avec retard**" pour le champ "Deux types
d'attestations"). Si le deuxième type d'attestation a un texte
différent, envoyez-moi ce second modèle Word et j'ajouterai la
logique pour choisir automatiquement le bon gabarit selon
`type_attestation`.

## Où sont générées les fiches ?

Par défaut dans `~/Fiches_Visa_Generees` (modifiable via le bouton
"Dossier de sortie..." dans l'application), un fichier
`Visa_<Raison sociale>.docx` par client.
