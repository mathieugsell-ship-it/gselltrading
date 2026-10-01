# Prospection — fiches Infopro « DRINGEND » (dossier `harworth/`)

`Prospection_Harworth.xlsx` : fichier de phoning (en anglais) généré à partir des PDF de `harworth/`
(108 projets, 204 contacts ; colonne A = nom du PDF source).

## Régénérer après ajout de nouveaux PDF

```bash
pip install pdfplumber openpyxl
cd harworth && python3 ../prospection/scripts/parse.py . > /tmp/parsed.json && cd ..
python3 prospection/scripts/build.py /tmp/parsed.json prospection/Prospection_Harworth.xlsx
```

Attention : régénérer recrée un fichier vierge — les colonnes de suivi d'appel déjà
remplies ne sont pas reprises.
