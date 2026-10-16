# sbae2xls_civ_2022_2025

Côte d'Ivoire, projet PROMIRE / ONU-REDD : données d'activité, estimations de superficie et émissions forestières 2021-2025 à partir d'un échantillonnage systématique (SBAE) interprété dans Collect Earth Online (CEO).

Les notebooks calculent l'activité de chaque échantillon (stable forêt, stable non-forêt, déforestation, dégradation, gain), estiment les superficies et leurs incertitudes avec l'estimateur stratifié (STR), puis calculent les émissions et absorptions. Les résultats sont exportés dans des **classeurs Excel dynamiques** : les entrées sont modifiables et toutes les cellules de résultats sont des formules.

Les calculs sont faits au niveau national, par zone phytogéographique et pour le territoire du **PRE** (Programme de Réductions des Émissions).

## Contenu

| Fichier | Rôle |
|:--|:--|
| `1_CIV_activite_CEO_2021_2025.ipynb` | Lit l'export CEO, rattache chaque échantillon à sa strate et au PRE, calcule l'activité 2021-2025. Produit `results/CIV_echantillons_activite_2021_2025.csv`. |
| `2_CIV_estimation_superficies_STR_2021_2025.ipynb` | Estimations de superficie et incertitudes (STR), national, par zone et pour le PRE. Produit les classeurs `CIV_estimations_superficie_STR_2021_2025_par_zone*.xlsx`. |
| `3_CIV_emissions_2021_2025.ipynb` | Émissions de la déforestation (biomasse, sols minéraux, combustion CH₄ et N₂O), de la dégradation et absorptions des gains, avec propagation des incertitudes (GIEC, approche 1). Produit les classeurs `CIV_emissions_deforestation_2021_2025*.xlsx`. |
| `sbae_excel.py` | Module : classeur Excel des estimations de superficie (utilisé par le notebook 2). |
| `sbae_emissions_excel.py` | Module : classeur Excel des émissions et absorptions (utilisé par le notebook 3). |
| `sepal_environment.yml` | Environnement conda. |

Les notebooks s'exécutent dans l'ordre 1, 2, 3, depuis la racine du dépôt.

## Installation

Sur SEPAL ou sur un poste avec conda (ou mamba) :

```bash
conda env create -f sepal_environment.yml
conda activate sbae2xls_civ
python -m ipykernel install --user --name sbae2xls_civ --display-name "sbae2xls_civ"
```

Ouvrir ensuite les notebooks dans Jupyter et choisir le noyau `sbae2xls_civ`.

## Données d'entrée

Les données ne sont pas incluses dans le dépôt. Elles doivent être placées dans le dossier `data/` :

| Fichier | Contenu |
|:--|:--|
| `data/ceo-RCI_MONITORING_21-25-plot-data_QA_QC-rev-2.csv` | Export CEO des échantillons interprétés (encodage latin-1). Colonnes `plotid`, `center_lon`, `center_lat`, et une colonne par classe et par année valant 100 pour la classe choisie et 0 sinon : `Occupation_<année>:<classe>` (niveau 1), `..._Niveau2_<année>:<classe>` (niveau 2), `..._Niveau3_<année>:<classe>` (niveau 3). |
| `data/allWithStrata_civ_20_26.shp` | Points des échantillons avec leur strate (colonne `mode`, valeurs 1, 2, 3). |
| `data/ZnePhytoGeo_CI.shp` | Zones phytogéographiques (colonne `Zone`). |
| `data/shapefile_PRE_repaired_geometry.shp` | Limite du PRE (polygone). |

Les noms de fichiers sont définis au début de chaque notebook et peuvent être modifiés.

Les superficies cartographiées des strates (nationales, par zone, et par zone dans le PRE) sont saisies dans les notebooks 2 et 3. Les facteurs d'émission (stocks de carbone, facteurs des sols, paramètres de combustion, taux de croissance) sont saisis dans le notebook 3, avec leur source. Toutes ces valeurs restent modifiables dans les classeurs produits.

## Résultats

Les fichiers sont écrits dans le dossier `results/` :

- `CIV_echantillons_activite_2021_2025.csv` : une ligne par échantillon (strate, appartenance au PRE, activité, classes de chaque année) ;
- `CIV_estimations_superficie_STR_2021_2025_par_zone.xlsx` et `..._PRE.xlsx` : superficies par classe d'activité, erreurs standard, incertitudes et intervalles de confiance à 90 % ;
- `CIV_emissions_deforestation_2021_2025.xlsx` et `..._PRE.xlsx` : données d'activité par transition, facteurs, émissions et absorptions annuelles, bilan et incertitudes.

## Méthodes et références

- Estimateur stratifié et post-stratification par zone : Cochran (1977) *Sampling Techniques* ; Olofsson et al. (2014) *Good practices for estimating area and assessing accuracy of land change*, Remote Sensing of Environment 148.
- Émissions, absorptions et propagation des incertitudes : GIEC (2006) Lignes directrices pour les inventaires nationaux de gaz à effet de serre, vol. 1 ch. 3 et vol. 4 ch. 2, 4 et 5 ; GIEC (2019) Affinement.
- Facteurs d'émission nationaux : NERF 2024 de la Côte d'Ivoire et inventaire forestier national (IFN).
