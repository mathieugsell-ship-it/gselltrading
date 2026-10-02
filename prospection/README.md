# Prospection — fiches Infopro « DRINGEND » (dossier `harworth/`)

`Prospection_Harworth.xlsx` : fichier de phoning (en anglais) généré à partir des PDF de `harworth/`
(108 projets, 204 contacts ; colonne A = nom du PDF source).

## Régénérer après ajout de nouveaux PDF

```bash
pip install pdfplumber openpyxl
cd harworth && python3 ../prospection/scripts/parse.py . > /tmp/parsed.json && cd ..
python3 prospection/scripts/build.py /tmp/parsed.json prospection/Prospection_Harworth.xlsx
```

## Suivi des appels

Les comptes rendus d'appels sont stockés dans `call_log.json` et réappliqués à chaque
génération : régénérer le fichier ne perd donc pas le suivi. Chaque entrée est rattachée
à sa ligne par **fichier PDF + réf. projet Infopro + rôle + société** (jamais la société
seule) et ne peut modifier que les champs de contact et de suivi (Contact person, Phone,
Email, Last call result, Status, Attempts, Last call date, Next action, Follow-up date,
Notes / person reached). Une entrée qui ne correspond pas à exactement une ligne fait
échouer la génération.

Deux modes : `update` remplace la valeur d'un champ (compte rendu d'appel), `append`
ajoute du texte à la suite des notes existantes sans rien effacer (recherches préalables).

Attention : une saisie faite directement dans Excel n'est pas reportée dans
`call_log.json` — elle serait écrasée par une régénération.
