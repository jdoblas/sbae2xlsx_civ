# sbae2xls_civ_2022_2025

*Français ci-dessous · [English version below](#english-version)*

Côte d'Ivoire, projet PROMIRE / ONU-REDD : données d'activité, estimations de superficie et émissions forestières 2021-2025 à partir d'un échantillonnage systématique (SBAE) interprété dans Collect Earth Online (CEO).

Les notebooks calculent l'activité de chaque échantillon (stable forêt, stable non-forêt, déforestation, dégradation, gain), estiment les superficies et leurs incertitudes avec l'estimateur stratifié (STR), puis calculent les émissions et absorptions. Les résultats sont exportés dans des **classeurs Excel dynamiques** : les entrées sont modifiables et toutes les cellules de résultats sont des formules.

Les calculs sont faits au niveau national, par zone phytogéographique et pour le territoire du **PRE** (Programme de Réductions des Émissions).

## Contenu

| Fichier | Rôle |
|:--|:--|
| `1_CEO_output_formatting.ipynb` | Lit l'export CEO, rattache chaque échantillon à sa strate et au PRE, calcule l'activité 2021-2025. Produit `results/CIV_echantillons_activite_2021_2025.csv`. |
| `2_area_estimation.ipynb` | Estimations de superficie et incertitudes (STR), national, par zone et pour le PRE. Produit les classeurs `2_area_estimation*.xlsx`. |
| `3_emissions_and_removals.ipynb` | Émissions de la déforestation (biomasse, sols minéraux, combustion CH₄ et N₂O), de la dégradation et absorptions des gains, avec propagation des incertitudes (GIEC, approche 1). Produit les classeurs `3_emissions_and_removals*.xlsx`. |
| `sbae_excel.py` | Module : classeur Excel des estimations de superficie (utilisé par le notebook 2). |
| `sbae_emissions_excel.py` | Module : classeur Excel des émissions et absorptions (utilisé par le notebook 3). |
| `sepal_environment.yml` | Environnement conda. |

Les notebooks s'exécutent dans l'ordre 1, 2, 3, depuis la racine du dépôt.

## Utilisation dans Sepal

1. Cloner le dépôt dans votre compte utilisateur Sepal (par exemple depuis un terminal Sepal) :

   ```bash
   git clone https://github.com/jdoblas/sbae2xlsx_civ.git
   cd sbae2xlsx_civ
   ```

2. Déposer les fichiers de données d'entrée dans le dossier `data/` du dépôt cloné (voir [Données d'entrée](#données-dentrée)).

3. Ouvrir les notebooks dans Sepal Jupyter et choisir l'un des noyaux **Sampling Based Area Estimation** (n'importe lequel convient). Exécuter les notebooks dans l'ordre 1, 2, 3.

4. Les résultats sont écrits dans le dossier `results/`. Les classeurs `.xlsx` peuvent ensuite être téléchargés et utilisés localement sur l'ordinateur de l'utilisateur.

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
- `2_area_estimation.xlsx` et `..._PRE.xlsx` : superficies par classe d'activité, erreurs standard, incertitudes et intervalles de confiance à 90 % ;
- `3_emissions_and_removals.xlsx` et `..._PRE.xlsx` : données d'activité par transition, facteurs, émissions et absorptions annuelles, bilan et incertitudes.

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
| `1_CEO_output_formatting.ipynb` | `1_CEO_output_formatting_en.ipynb` | Reads the CEO export, links each sample to its stratum and to the ERP, computes the 2021-2025 activity. Produces the samples CSV. |
| `2_area_estimation.ipynb` | `2_area_estimation_en.ipynb` | Area estimates and uncertainties (STR), national, by zone and for the ERP. Produces the area workbooks. |
| `3_emissions_and_removals.ipynb` | `3_emissions_and_removals_en.ipynb` | Emissions from deforestation (biomass, mineral soils, CH₄ and N₂O from biomass burning) and degradation, and removals from gains, with uncertainty propagation (IPCC approach 1). Produces the emissions workbooks. |
| `sbae_excel.py` | `sbae_excel_en.py` | Module: Excel workbook of the area estimates (used by notebook 2). |
| `sbae_emissions_excel.py` | `sbae_emissions_excel_en.py` | Module: Excel workbook of the emissions and removals (used by notebook 3). |
| `sepal_environment.yml` | | Conda environment (shared). |

The notebooks are run in the order 1, 2, 3, from the root of the repository. The French and English chains are independent: each one reads and writes its own files.

### Usage in Sepal

1. Clone the repository into your Sepal user account (for example from a Sepal terminal):

   ```bash
   git clone https://github.com/jdoblas/sbae2xlsx_civ.git
   cd sbae2xlsx_civ
   ```

2. Put the input data files into the `data/` folder of the cloned repository (see [Input data](#input-data)).

3. Open the notebooks in Sepal Jupyter and select any of the **Sampling Based Area Estimation** kernels. Run the notebooks in the order 1, 2, 3.

4. The results are written to the `results/` folder. The `.xlsx` workbooks can then be downloaded and used locally on the user's computer.


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
| `2_area_estimation.xlsx` and `..._PRE.xlsx` | `2_area_estimation_en.xlsx` and `..._ERP_en.xlsx` | Areas by activity class, standard errors, uncertainties and 90% confidence intervals. |
| `3_emissions_and_removals.xlsx` and `..._PRE.xlsx` | `3_emissions_and_removals_en.xlsx` and `..._ERP_en.xlsx` | Activity data by transition, factors, annual emissions and removals, balance and uncertainties. |

### Methods and references

- Stratified estimator and post-stratification by zone: Cochran (1977) *Sampling Techniques*; Olofsson et al. (2014) *Good practices for estimating area and assessing accuracy of land change*, Remote Sensing of Environment 148.
- Emissions, removals and uncertainty propagation: IPCC (2006) Guidelines for National Greenhouse Gas Inventories, vol. 1 ch. 3 and vol. 4 ch. 2, 4 and 5; IPCC (2019) Refinement.
- National emission factors: FREL 2024 (*NERF*) of Côte d'Ivoire and national forest inventory (NFI, *IFN*).
