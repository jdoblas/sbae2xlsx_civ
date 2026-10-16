# sbae2xls_civ_2022_2025

*Français ci-dessous · [English version below](#english-version)*

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

---

## English version

Côte d'Ivoire, PROMIRE project / UN-REDD: activity data, area estimates and forest emissions 2021-2025 from a systematic sample (SBAE) interpreted in Collect Earth Online (CEO).

The notebooks compute the activity of each sample (stable forest, stable non-forest, deforestation, degradation, gain), estimate the areas and their uncertainties with the stratified estimator (STR), then compute the emissions and removals. The results are exported to **dynamic Excel workbooks**: the inputs are editable and all result cells are formulas.

The calculations are made at the national level, by phytogeographic zone and for the territory of the **ERP** (Emission Reductions Program, *PRE* in French).

Every notebook and module exists in French and in English. The English files carry the suffix `_en` and perform exactly the same calculations: the English workbooks have the same formulas and the same values as the French ones; only the labels, sheet names and named ranges are translated. Values that come from the input data, such as the CEO class names (e.g. `Foret_Dense`) and the zone names (e.g. `Secteur Ombrophile`), are kept as they are.

### Contents

| French | English | Role |
|:--|:--|:--|
| `1_CIV_activite_CEO_2021_2025.ipynb` | `1_CIV_activity_CEO_2021_2025_en.ipynb` | Reads the CEO export, links each sample to its stratum and to the ERP, computes the 2021-2025 activity. Produces the samples CSV. |
| `2_CIV_estimation_superficies_STR_2021_2025.ipynb` | `2_CIV_area_estimation_STR_2021_2025_en.ipynb` | Area estimates and uncertainties (STR), national, by zone and for the ERP. Produces the area workbooks. |
| `3_CIV_emissions_2021_2025.ipynb` | `3_CIV_emissions_2021_2025_en.ipynb` | Emissions from deforestation (biomass, mineral soils, CH₄ and N₂O from biomass burning) and degradation, and removals from gains, with uncertainty propagation (IPCC approach 1). Produces the emissions workbooks. |
| `sbae_excel.py` | `sbae_excel_en.py` | Module: Excel workbook of the area estimates (used by notebook 2). |
| `sbae_emissions_excel.py` | `sbae_emissions_excel_en.py` | Module: Excel workbook of the emissions and removals (used by notebook 3). |
| `sepal_environment.yml` | | Conda environment (shared). |

The notebooks are run in the order 1, 2, 3, from the root of the repository. The French and English chains are independent: each one reads and writes its own files.

### Installation

On SEPAL or on a computer with conda (or mamba):

```bash
conda env create -f sepal_environment.yml
conda activate sbae2xls_civ
python -m ipykernel install --user --name sbae2xls_civ --display-name "sbae2xls_civ"
```

Then open the notebooks in Jupyter and select the `sbae2xls_civ` kernel.

### Input data

The data are not included in the repository. They must be placed in the `data/` folder:

| File | Content |
|:--|:--|
| `data/ceo-RCI_MONITORING_21-25-plot-data_QA_QC-rev-2.csv` | CEO export of the interpreted samples (latin-1 encoding). Columns `plotid`, `center_lon`, `center_lat`, and one column per class and per year equal to 100 for the chosen class and 0 otherwise: `Occupation_<year>:<class>` (level 1), `..._Niveau2_<year>:<class>` (level 2), `..._Niveau3_<year>:<class>` (level 3). |
| `data/allWithStrata_civ_20_26.shp` | Sample points with their stratum (column `mode`, values 1, 2, 3). |
| `data/ZnePhytoGeo_CI.shp` | Phytogeographic zones (column `Zone`). |
| `data/shapefile_PRE_repaired_geometry.shp` | ERP boundary (polygon). |

The file names are defined at the beginning of each notebook and can be changed.

The mapped stratum areas (national, by zone, and by zone within the ERP) are entered in notebooks 2 and 3. The emission factors (carbon stocks, soil factors, burning parameters, growth rates) are entered in notebook 3, with their source. All these values remain editable in the workbooks produced.

### Results

The files are written to the `results/` folder:

| French chain | English chain | Content |
|:--|:--|:--|
| `CIV_echantillons_activite_2021_2025.csv` | `CIV_activity_samples_2021_2025_en.csv` | One row per sample (stratum, ERP membership, activity, classes of each year); the two files are identical. |
| `CIV_estimations_superficie_STR_2021_2025_par_zone.xlsx` and `..._PRE.xlsx` | `CIV_area_estimates_STR_2021_2025_by_zone_en.xlsx` and `..._ERP_en.xlsx` | Areas by activity class, standard errors, uncertainties and 90% confidence intervals. |
| `CIV_emissions_deforestation_2021_2025.xlsx` and `..._PRE.xlsx` | `CIV_emissions_deforestation_2021_2025_en.xlsx` and `..._ERP_en.xlsx` | Activity data by transition, factors, annual emissions and removals, balance and uncertainties. |

### Methods and references

- Stratified estimator and post-stratification by zone: Cochran (1977) *Sampling Techniques*; Olofsson et al. (2014) *Good practices for estimating area and assessing accuracy of land change*, Remote Sensing of Environment 148.
- Emissions, removals and uncertainty propagation: IPCC (2006) Guidelines for National Greenhouse Gas Inventories, vol. 1 ch. 3 and vol. 4 ch. 2, 4 and 5; IPCC (2019) Refinement.
- National emission factors: FREL 2024 (*NERF*) of Côte d'Ivoire and national forest inventory (NFI, *IFN*).
