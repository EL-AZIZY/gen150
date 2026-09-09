# GEN150

GEN150 transforme les feuilles APSF et de suivi production explicitement configurées d’un classeur
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
│       ├── apsf_competitor_summary_parser.py
│       ├── apsf_competitor_quarterly_parser.py
│       ├── flat_table_parser.py
│       ├── production_marque_penetration_parser.py
│       ├── production_marque_concession_parser.py
│       ├── production_marque_pivot_staging_parser.py
│       ├── production_marque_variation_tpr_parser.py
│       └── pivot_utils.py
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
- est toujours ignorée si son état est `hidden` ou `veryHidden`, même si elle est configurée ;
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
`family_sheets`. Seuls les onglets visibles sont traités ; l'ancien flag `include_hidden`
n'autorise plus le traitement des onglets masqués. Les noms normalisés doivent être
uniques entre familles.

### `workbooks.yml`

`defaults.sheets` contient la structure commune des feuilles. Une entrée de `workbooks`
peut surcharger cette structure pour un nom de fichier précis ; les dictionnaires sont
fusionnés récursivement lorsque `inherit_defaults: true`.

`enabled_families` limite chaque classeur à ses propres familles. Cela évite qu'un
nom générique commun, comme `Feuil2`, soit traité avec les règles d'un autre classeur.

Les limites métier, la colonne de libellé, les lignes de section, les lignes de période et,
selon la famille, les colonnes de données résident dans ce fichier, pas dans le moteur
principal. Les ancres YAML évitent de répéter les structures communes.

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

Les deux synthèses concurrentielles utilisent `ApsfCompetitorSummaryParser` avec
le format `summary`. Le mode optionnel `staging` reste disponible pour les
configurations personnalisées ; aucune feuille livrée ne l’utilise par défaut.

Les deux feuilles de synthèse utilisent le même traitement. Leurs différences sont entièrement
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

## Famille `apsf_competitor_quarterly`

Le parser du commit v2 est restauré. Seule `Syn prin concurents par trim` est
configurée : section `PRODUCTION_NETTE` en ligne 9, catégories en 11, trimestres
en 13, données de 16 à 30. Les blocs E:H, J:M, O:R, T:W et Y:AB correspondent
aux véhicules, prêts personnels, équipement domestique, revolving et global.
Les formules auxiliaires des lignes 33/36 et colonnes AD:AH ne sont pas des sections.
Les quatre vues masquées restent exclues et le routeur demeure inchangé.

Le CSV contient une ligne par organisme/catégorie, avec quatre valeurs
trimestrielles et quatre valeurs Marché propagées. Les formats Excel pilotent la
conversion MDH ou pourcentage. Les lignes Part de marché et TOTAL sont exclues.

B2 est un intitulé sans date. `annee: 2026` fournit explicitement l’année de
reporting du classeur livré, indépendamment des variantes de son nom. Sans date
dans la zone de titre, le 31 décembre de cette année est utilisé par convention ;
une date dans le titre ou une `date_reporting` explicite reste prioritaire.
Ce repli est propre à cette famille et n’affecte pas les autres parsers.
Mettre à jour `annee` pour un nouveau millésime. Cette métadonnée ne garantit pas
le millésime des liens externes historiques présents dans les formules Excel.

## Famille `production_marque_details`

Cette famille couvre les trois tableaux simples `Détails`, `Feuil2` et `Feuil4`.
`FlatTableParser` choisit comme en-tête la première ligne non vide, puis conserve
les colonnes contiguës depuis A jusqu’au premier en-tête vide. Il détecte ainsi
la ligne 1 pour `Détails` et `Feuil2`, et la ligne 3 pour `Feuil4`. Les noms sont
normalisés en snake_case. La valeur auxiliaire H1=1011 de `Détails`, séparée par
G1 vide, est exclue. Les noms vides ou dupliqués après normalisation sont rejetés.

`Détails` conserve ses contrôles métier : les champs requis sont `marque`,
`nombre`, `ville`, `mois`, `annee`, `source`, dans n’importe quel ordre. `Feuil2`
et `Feuil4` conservent chacune leur propre schéma source sans recevoir de colonnes
artificielles. `Feuil2` produit quatre colonnes et 13 lignes ; `Feuil4` produit
deux colonnes et 45 lignes, total compris.

`nombre` devient un `Decimal`, `mois` et `annee` des entiers. Le parser conserve
le texte brut, notamment la casse de `ville`, ainsi que les champs vides. Il ne
corrige pas l’année 2821 et ne contrôle pas la plage d’année. Les politiques
existantes d’erreurs Excel et de caches de formules restent applicables.

Par défaut, une ligne entièrement vide arrête la lecture. Dans la configuration
livrée, `stop_on_empty_row: false` applique le choix confirmé par l’utilisateur :
ignorer les lignes entièrement vides et poursuivre jusqu’à la fin. Le fichier
contient 74 443 lignes de données, dont 19 942 après la ligne vide 54 503.
Cette ligne vide porte les seuls VILLE/Source absents du fichier réel. Les tests
synthétiques vérifient des champs VILLE/Source vides sur des lignes de données ;
le test réel vérifie l’année 2821 et la conservation des lignes après le blanc.

## Famille `production_marque_penetration`

Les 30 feuilles `marq Men`/`marq An`, globales et pour 14 villes, ainsi que `ville`, partagent
`ProductionMarquePenetrationParser`. Chaque feuille possède explicitement `ville`
(GLOBAL pour les vues globales) et `periodicite` (MENSUEL/ANNUEL). Les noms Excel
exacts sont configurés, notamment `marq An EL Jadida`. Aucune métadonnée n’est
déduite du nom de la feuille au moment du parsing.

La configuration porte `dimension: MARQUE` pour les 30 feuilles par marque et
`dimension: VILLE` pour `ville`. Pour la dimension VILLE, le libellé source
alimente `ville` et `marque` reste vide. Le CSV est en format long : une seule
colonne `axe` contient les cinq groupes de la ligne 7, `sous_axe` contient leurs
en-têtes de la ligne 8 et `valeur` contient la mesure correspondante. Une marque
produit donc 11 lignes. Les dates Excel sont rendues avec le mois en français,
par exemple `avril 2026`. Le format Excel des sous-axes du taux est conservé sous
la forme `TPR avr 2026`, `TPR mars 2026` ou `TPR avr 2025`.
Le parser vérifie l’en-tête correspondant en A8, commence en ligne 9 et s’arrête
au premier libellé vide ou `Total général`, sans exporter cette ligne. Une limite `data_end_row` explicite reste possible, sans être nécessaire
pour les 30 feuilles. C/D, F/G et I/J fournissent AIVAM/SOFAC courant, précédent
et année précédente. L/M/N contiennent les taux de pénétration et P/Q les
variations. Les champs `_pct` utilisent la conversion partagée selon le format
Excel ; le parser exporte les valeurs en cache sans recalculer les mesures.

Les formules F7 (`=C7-1`) et I7 (`=C5-365`) sont identiques entre Men et An.
L’inspection des filtres de `synthése Par marque` confirme cependant que E:G
sélectionne mars 2026 et Q:S janvier à mars 2026 ; I:K sélectionne avril 2025
et U:W janvier à avril 2025. La vue globale An est cumulative. Cette preuve ne
couvre pas tous les filtres de chaque ville : cette limite est notée dans le parser.

La feuille `ville` fournit aussi la date C5 référencée par les formules des autres
feuilles. Elle contient deux tableaux de même structure dans le même groupe : le
tableau mensuel utilise les lignes 9 à 36 et le tableau annuel les lignes 48 à 75.
Chaque tableau s’arrête à son propre `Total général`, qui n’est pas exporté. Les
28 villes produisent donc 308 lignes mensuelles et 308 lignes annuelles.

## Famille `production_marque_concession`

La feuille `Concession` contient quatre groupes dans le fichier livré : BAMOTORS,
SOPRIAM, TOYOTA MAROC et JAMEL. Les titres TPR sont répétés pour les deux années,
soit huit occurrences. Le parser recherche les titres, puis associe les tableaux
situés dans leur zone aux en-têtes `Étiquettes de lignes`, AIVAM, SOFAC et Total
général. Les années proviennent des cellules `année` au-dessus de chaque tableau.
Les positions, l’ordre des colonnes et le nombre de marques ne sont pas figés.

Une ligne CSV représente une marque ou le total d’un groupe pour une année :
`groupe_concessionnaire, annee, marque, aivam, sofac, total, row_type`.
Le total est conservé avec `row_type=total`, puis la lecture de ce tableau s’arrête.
Le fichier livré produit 24 lignes, dont huit totaux, pour 2025 et 2026.

## Famille `production_marque_pivot_staging`

Cette famille couvre les 15 synthèses par marque et `Synthése par ville`.
`Table` reste hors périmètre de la famille des tableaux simples : sa première
ligne non vide est un titre et la feuille contient plusieurs tableaux séparés.

Les ancres sont `Étiquettes de lignes`, sauf pour `Synthése par ville`, dont les
six blocs utilisent `VILLE` suivi des en-têtes AIVAM/SOFAC. Les 15 synthèses par
marque ont aussi six blocs, avec des largeurs variables ; certains n’ont que la
colonne AIVAM.

Le CSV est long et structuré : `sheet, bloc_index, axe, periodicite, annee, mois,
dimension, marque, ville, organisme, valeur`. Chaque cellule de mesure
génère une ligne. `axe` vient de la première cellule non vide placée au-dessus du
bloc (`Mois M`, `Mois M-1`, `Mois M N-1`, `ANNEE N M`, `ANNEE N M-1` ou
`ANNEE N-1`) et `periodicite` vaut `MENSUEL` ou `ANNUEL`.

Pour une synthèse par marque, le libellé de ligne alimente `marque` et le filtre
VILLE alimente `ville`. Pour `Synthése par ville`, le libellé alimente `ville`,
`marque` reste nul et `dimension` vaut `VILLE`. Les en-têtes AIVAM/SOFAC
alimentent `organisme`. Les valeurs vides, les répétitions et les totaux sont
conservés, sans déduplication ni réconciliation. Les périodes absentes restent
nulles. Dans `mois`, l'affichage
`(Plusieurs éléments)` est remplacé par les choix réellement enregistrés dans
le filtre du tableau croisé, par exemple `1,2,3`. Ces choix sont lus directement
dans les définitions XML du classeur et ne sont pas inscrits dans la
configuration. L’ordre des blocs est celui de leurs ancres, par ligne puis
colonne.

`pivot_utils.py` factorise la découverte des ancres, des colonnes contiguës,
du contexte de période et des lignes pour les deux nouvelles familles.
`context_lookback` fixe la fenêtre de recherche des périodes au-dessus des ancres
(huit lignes par défaut), sans figer les coordonnées des tableaux.

La configuration applique `blank_label_policy: keep_with_values` : une ligne
sans libellé est conservée si elle contient au moins une mesure, y compris zéro.
Le libellé manquant reste vide ; aucune marque n’est inventée. Chaque bloc s’arrête
seulement quand son libellé et toutes ses mesures sont vides. Cela permet notamment
de conserver M8 vide/O8=0 et de lire la suite du bloc annuel de la synthèse globale.
L’option `stop` reste disponible pour les configurations nécessitant un arrêt
au premier libellé vide.

## Famille `production_marque_variation_tpr`

Cette famille traite les deux tableaux de la feuille `Table` dans un seul CSV.
La configuration fixe leurs zones et leurs colonnes, tandis que les dates et les
valeurs restent lues dans le classeur.

Le CSV est au format long `ville, axe, valeur`. Les intitulés des deux périodes
et `Variation` deviennent le contenu de la colonne `axe`, par exemple `avr-26`,
`mars-26` et `Variation`. Le parser compare les dates d'en-tête et associe chaque
valeur au bon axe, même lorsque leur ordre est inversé entre les deux tableaux.
Les valeurs au format pourcentage sont converties en pourcentage et les variations
restent en points, avec deux décimales dans le CSV.

La vue détaillée ignore sa ligne de séparation vide et continue jusqu'à la fin
de sa zone configurée. Les valeurs Excel `NS` et les erreurs mises en cache sont
conservées selon les politiques générales du projet.


### Compatibilité des caches de tableaux croisés

Le fichier production déclenche une erreur interne d’openpyxl 3.1.5 pendant la
lecture des caches de tableaux croisés (`Nested.from_tree() missing ... 'node'`).
Le chargeur tente d’abord le chargement normal. Pour cette erreur précise,
il réessaie avec une copie en mémoire sans déclaration de caches ni relations
vers les objets PivotTable. Un avertissement est écrit dans le journal.

Les cellules, formules, valeurs en cache, formats, fusions et états masqués sont
conservés. Le fichier Excel original n’est jamais réécrit. Cette copie sert
uniquement à l’extraction : elle ne doit pas être utilisée pour enregistrer ou
modifier les tableaux croisés. Les autres erreurs continuent à être remontées.

Le traitement direct du fichier original a été vérifié : 51 feuilles exportées,
aucune rejetée, dont 74 443 lignes pour Détails. Les tests couvrent aussi le
chargement normal, les erreurs sans rapport avec les caches, la fermeture des
ressources et la conservation du fichier source et du XML des cellules.

## Nombres et formules

Les nombres sont convertis avec `Decimal` et écrits sans notation scientifique. Dans les
familles métier, les montants et nombres de dossiers sont arrondis à l'entier le plus
proche (`2870.929` devient `2871`) ; les pourcentages et points sont arrondis à une
décimale (`20.889` devient `20.9`). L'arrondi à `0.5` s'effectue comme `ARRONDI` dans
Excel, en s'éloignant de zéro. La colonne `value` de le mode staging reste brute afin
de préserver sa fonction de réconciliation. Une cellule vide reste vide et `NS` reste
`NS`.

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
- la notation scientifique ;
- les doublons de clé métier ;
- la continuité de `ordre_ligne` par organisme et section ;
- l’absence de `Marché` et `Part de marché` comme organismes ;
- la cohérence de la propagation du contexte Marché ;
- la séparation stricte entre champs monétaires et champs de ratio.

Ces règles sont propres aux familles métier. Le mode staging contrôle seulement son
en-tête, la notation scientifique et l’unicité de sa clé technique ; elle accepte les
lignes `Marché`, `Part de marché` et `TOTAL`.

## Ajouter une famille

1. Créer une classe dérivée de `BaseParser` dans `src/parsers/` avec son en-tête exact et
   sa méthode `process()`.
2. Importer la classe et ajouter l’association famille/classe dans `FAMILY_REGISTRY`, dans
   `src/parsers/__init__.py`.
3. Ajouter l’entrée de famille et sa liste `family_sheets` dans `families.yml`.
4. Ajouter les structures des feuilles dans `workbooks.yml`.
5. Ajouter les validations propres à la famille et ses tests.

En règle générale, le routeur, le chargement du classeur, l’exporteur et le runner n’ont
pas à être modifiés pour ajouter une famille.

## Tests

La suite utilise la bibliothèque standard `unittest` et ne requiert pas de dépendance de
test supplémentaire :

```powershell
python -m unittest discover -s tests -v
```

Elle couvre la configuration, le routage, les feuilles masquées et non configurées, les
limites, la taxonomie, les totaux, les valeurs vides/zéro/`NS`, les erreurs Excel, les
périodes, la mise à l’échelle MDH, les pourcentages, les points, la propagation Marché,
la découverte des blocs fusionnés, les périodes décalées, les en-têtes dynamiques,
les limites de lecture, les dimensions MARQUE/VILLE, les groupes concessionnaires,
les blocs de tableaux croisés et les flux Excel → CSV → manifeste.
