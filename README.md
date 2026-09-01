# GEN150

GEN150 transforme uniquement les feuilles APSF explicitement configurées d’un classeur
Excel en fichiers CSV prêts à charger en base de données. Il prend en charge plusieurs
familles de feuilles dans un même classeur, sans appliquer chaque parser à toutes les
feuilles.

## Architecture

```text
GEN150/
├── main.py
├── requirements.txt
├── config/
│   ├── settings.yml
│   ├── families.yml
│   └── workbooks.yml
├── input/
├── output/
│   ├── success/
│   ├── rejected/
│   ├── logs/
│   └── manifests/
├── src/
│   ├── config_loader.py
│   ├── config_validator.py
│   ├── date_utils.py
│   ├── exporter.py
│   ├── numeric_utils.py
│   ├── runner.py
│   ├── settings.py
│   ├── sheet_router.py
│   ├── text_utils.py
│   ├── validators.py
│   ├── workbook_loader.py
│   └── parsers/
│       ├── base_parser.py
│       ├── apsf_activity_parser.py
│       └── apsf_competitor_summary_parser.py
└── tests/
```

## Installation

Python 3.11 ou plus récent est recommandé.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Les dépendances d’exécution sont limitées à `openpyxl` et `PyYAML`.

## Exécution

Traiter tous les fichiers `.xlsx` et `.xlsm` du dossier `input/` :

```powershell
python main.py
```

Traiter un classeur précis :

```powershell
python main.py "D:\donnees\SCC_Statistiques APSF Globales Mars 2026.xlsx"
```

Surcharger les dossiers de configuration et de sortie :

```powershell
python main.py "D:\donnees\apsf.xlsx" --config-dir "D:\config\gen150" --output-dir "D:\exports\gen150"
```

Le code de retour vaut `0` si toutes les feuilles traitées réussissent, `1` si au moins
une feuille est rejetée et `2` si le lancement ou la configuration échoue.

## Routage des feuilles

Au chargement du classeur, GEN150 construit une seule fois un index normalisé des noms de
feuilles. La comparaison supprime les espaces de bord, réduit les espaces multiples et
ignore la casse. Le nom Excel original reste utilisé pour accéder à la feuille.

Pour chaque famille, le routeur consulte seulement `family_sheets`. Une feuille :

- n’est routée que si elle appartient explicitement à une famille active ;
- n’est traitée qu’une fois ;
- est ignorée sans rejet métier si son état est `hidden` ou `veryHidden` ;
- est ignorée si elle n’est déclarée dans aucune famille.

Une collision de noms après normalisation ou l’appartenance d’une feuille à deux familles
est une erreur de configuration.

## Configuration

### `settings.yml`

Ce fichier configure les dossiers d’entrée et de sortie, le séparateur et l’encodage CSV,
le niveau de log et les politiques d’erreurs Excel.

Les politiques disponibles sont :

- `warning` : journaliser l’anomalie et conserver la valeur d’erreur, sans la convertir ;
- `reject` : rejeter la feuille entière.

`error_policy` concerne les erreurs Excel telles que `#REF!`, `#VALUE!` et `#DIV/0!`.
`missing_formula_cache_policy` concerne une formule qui ne possède pas de résultat calculé
en cache dans le fichier. GEN150 ne recalcule jamais Excel.

### `families.yml`

Chaque famille déclare son parser, son activation et la liste obligatoire
`family_sheets`. Les noms normalisés doivent être uniques entre familles.

### `workbooks.yml`

`defaults.sheets` contient la structure commune des feuilles. Une entrée de `workbooks`
peut surcharger cette structure pour un nom de fichier précis ; les dictionnaires sont
fusionnés récursivement lorsque `inherit_defaults: true`.

Les limites métier, la colonne de libellé, les blocs de catégories, les lignes de section,
les lignes de période et les colonnes de données résident dans ce fichier, pas dans le
moteur principal. Les ancres YAML évitent de répéter la structure commune aux 17 feuilles
d’activité.

## Famille `apsf_activity`

Le parser lit uniquement les lignes 1 à 117 et les colonnes A à I définies dans la
configuration. Il produit au maximum une ligne CSV par ligne source et n’effectue aucun
unpivot.

Les lignes vides, les libellés `TOTAL`, les lignes d’unité et les lignes de présentation
sans donnée sont ignorés. Les marqueurs `PRÊTS AFFECTÉS` et `PRÊTS NON AFFECTÉS` mettent à
jour `type_pret` mais ne sont jamais exportés. La taxonomie de catégories et de produits
est définie dans `hierarchy`. Un libellé commençant par `LOA_` reste fusionné dans
`produit` ; aucune colonne de détail n’est créée.

Les colonnes B à I sont associées aux huit mesures par `data_columns`. Dans le format APSF
livré, les années sont lues explicitement dans `year_cells.reference` et
`year_cells.comparison` (`B7/C7`). Elles peuvent aussi être fixées avec
`annee_reference`/`annee_comparaison` ou, en dernier recours, recherchées dans
`title_area`. La date de reporting est le dernier jour réel du mois trouvé dans le titre ;
une vraie cellule de date Excel est également acceptée.

Les sections peuvent être mappées par `sections`. Avec `infer_sections: true`, un libellé
de présentation inconnu situé après la zone de titre devient le contexte de section ;
`default_section` s’applique uniquement tant qu’aucun contexte n’a été rencontré.

## Famille `apsf_competitor_summary`

Les deux feuilles de synthèse utilisent le même parser. Leurs différences sont entièrement
décrites par `label_column`, `category_blocks` et `sections`.

Chaque bloc de catégorie possède ses propres cellules de période. `Mar-26`, `Mar-25`,
`Dec-25` et `Dec-22` deviennent respectivement `2026-03`, `2025-03`, `2025-12` et
`2022-12`.

La ligne `Part de marché` est ignorée. La ligne `Marché` n’est pas exportée : ses trois
valeurs sont conservées par couple `(section, categorie)` puis propagées aux organismes de
ce même couple. Par défaut, une catégorie avec des organismes mais sans contexte Marché
est rejetée.

Pour `PRODUCTION_NETTE`, `ENCOURS_SAIN`, `CREANCES_EN_SOUFFRANCE` et `ENCOURS_BRUT`, les
valeurs alimentent les champs `_mdh` et `_pct`. Pour `CES_SUR_ENCOURS_BRUT`, elles alimentent
uniquement les champs de ratio `_pct` et `_points`. Un format Excel avec une virgule de
mise à l’échelle, par exemple `#,##0,`, entraîne la conversion en MDH par division par
1 000.

## Nombres et formules

Les nombres sont convertis avec `Decimal`, puis écrits sans séparateur de milliers, sans
notation scientifique et sans zéros décimaux inutiles. Une cellule vide reste vide, zéro
reste `0` et `NS` reste `NS`.

Chaque classeur est chargé deux fois :

- `data_only=True` fournit les valeurs calculées mises en cache ;
- `data_only=False` permet l’audit des formules et références erronées.

Une erreur Excel n’est jamais remplacée par zéro.

## Sorties

Une feuille validée génère un CSV dans `output/success/`, au format :

```text
<classeur_normalise>__<feuille_normalisee>.csv
```

Le séparateur est `;`, l’encodage est `utf-8-sig` et aucune colonne d’index n’est écrite.
Une feuille dont les validations échouent produit son CSV candidat et un rapport JSON dans
`output/rejected/`. Une erreur empêchant le parsing produit seulement le rapport JSON.

Un manifeste JSON par classeur résume le parser, la famille, le statut, le nombre de lignes,
les avertissements et le chemin de chaque sortie. Les logs se trouvent dans `output/logs/`.

## Validations

Avant l’export en succès, GEN150 contrôle notamment :

- l’en-tête exact et l’absence de colonnes supplémentaires ;
- les lignes `TOTAL` et techniques ;
- la notation scientifique ;
- les doublons de clé métier ;
- la continuité de `ordre_ligne` par organisme et section ;
- l’absence de `Marché` et `Part de marché` comme organismes ;
- la cohérence de la propagation du contexte Marché ;
- la séparation stricte entre champs monétaires et champs de ratio.

## Ajouter une famille

1. Créer une classe dérivée de `BaseParser` dans `src/parsers/` avec son en-tête exact et
   sa méthode `process()`.
2. Importer la classe et ajouter l’association famille/classe dans `FAMILY_REGISTRY`, dans
   `src/parsers/__init__.py`.
3. Ajouter l’entrée de famille et sa liste `family_sheets` dans `families.yml`.
4. Ajouter les structures des feuilles dans `workbooks.yml`.
5. Ajouter les validations propres à la famille et ses tests.

Le routeur, le chargement du classeur, l’exporteur et le runner n’ont pas à être modifiés.

## Tests

La suite utilise la bibliothèque standard `unittest` et ne requiert pas de dépendance de
test supplémentaire :

```powershell
python -m unittest discover -s tests -v
```

Elle couvre la configuration, le routage, les feuilles masquées et non configurées, les
limites, la taxonomie, les totaux, les valeurs vides/zéro/`NS`, les erreurs Excel, les
périodes, la mise à l’échelle MDH, les pourcentages, les points, la propagation Marché, les
en-têtes et un flux complet Excel → CSV → manifeste.
