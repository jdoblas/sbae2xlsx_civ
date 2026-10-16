"""Classeur Excel « vivant » pour les estimations de superficie SBAE (estimateur stratifié STR), national et par zone.

Toutes les cellules de résultats sont des formules Excel : le classeur fonctionne de manière autonome.
Les seules valeurs saisies sont les entrées (cellules jaunes) :

* feuille « Échantillons » : une ligne par échantillon (strate, zone, classe) -> effectifs n_hk et n_hkz par NB.SI.ENS ;
* feuille « Entrées »    : niveau de confiance, superficies nationales des strates A_h, superficies A_hz par zone.

Modifier une entrée (reclasser un échantillon, changer une superficie, le niveau de confiance...) recalcule
l'ensemble des estimations, des tableaux et des graphiques.

Utilisation depuis un notebook ::

    from sbae_excel import build_workbook
    out = build_workbook('../results/estimations.xlsx', plots=df_valid, area_hz=area_hz[[1, 2, 3]],
                         area_h=strata_sizes.set_index('stratum')['area_ha'])
    out['zone'], out['national']   # valeurs calculées en Python (identiques aux formules), pour contrôle

Références : Cochran (1977) ; Olofsson et al. (2014), Remote Sensing of Environment 148, 42-57.
"""
import datetime
import re
from statistics import NormalDist

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

# ---------------------------------------------------------------- noms des feuilles
S_SYN = 'Synthèse'
S_IN = 'Entrées'
S_RZ = 'Résultats par zone'
S_TC = 'Tableau croisé'
S_NZ = 'National vs zones'
S_NAT = 'SBAE national'
S_GR = 'Graphiques'
S_EQ = 'Équations'
S_PL = 'Échantillons'

def territory_labels(territory=None):
    """Libellés du territoire couvert par l'estimateur global : national par défaut, ou un territoire nommé (ex. 'PRE')."""
    if territory is None:
        return dict(adj='national', Adj='National', fem='nationale', fem_pl='nationales')
    return dict(adj=territory, Adj=territory, fem=f'du {territory}', fem_pl=f'du {territory}')


def sheet_descriptions(s_nz=S_NZ, s_nat=S_NAT, adj='national'):
    return [
    (S_IN, 'Entrées modifiables : niveau de confiance, superficies A_h et A_hz ; poids et effectifs calculés'),
    (S_RZ, 'Superficie, erreur standard et incertitude par zone et par classe'),
    (S_TC, 'Tableaux croisés classe × zone : superficies, erreurs standard, incertitudes, effectifs'),
    (s_nz, f'SBAE {adj} comparé à la somme des SBAE par zone'),
    (s_nat, f"Résultats de l'estimateur STR {adj}"),
    (S_GR, 'Graphiques des estimations (± incertitude)'),
    (S_EQ, 'Équations des estimateurs et leur mise en œuvre dans le classeur'),
    (S_PL, 'Échantillons : strate, zone et classe de référence (modifiables)'),
    ]


SHEET_DESCRIPTIONS = sheet_descriptions()

# ---------------------------------------------------------------- palette et formats numériques
GREEN, GREEN_L, BAND, GRID, GREY, INK = '#1F4E3D', '#DDEBE3', '#F3F8F5', '#D5DED9', '#6B7280', '#111827'
INPUT_BG, INPUT_INK = '#FFF4CC', '#1D4ED8'
FONT = 'Calibri'
HA, INT, PROP, PCT, DPCT, RATIO, W4 = '#,##0', '#,##0', '0.00%', '0.0%', '+0.0%;-0.0%;0.0%', '0.00', '0.0000'
DHA = '+#,##0;-#,##0;0'
CHART_COLORS = ['#1F4E3D', '#5B9C7A', '#B7D7C3', '#D9A441', '#7A6C9C', '#C0504D', '#4C72B0', '#8C8C8C']
SCALE_CI = {'type': '3_color_scale', 'min_type': 'num', 'min_value': 0.05, 'mid_type': 'num', 'mid_value': 0.25,
            'max_type': 'num', 'max_value': 0.6, 'min_color': '#C6EFCE', 'mid_color': '#FFF2CC', 'max_color': '#F8CBAD'}

RT_TOKEN = re.compile(r'([_^])(\{[^}]*\}|[A-Za-zΣ0-9]+)|([^_^]+)')


# ---------------------------------------------------------------- références de cellules
def ref(sheet, r, c):
    """Référence absolue vers une cellule d'une autre feuille, p. ex. 'Entrées'!$C$6."""
    return f"'{sheet}'!{xl_rowcol_to_cell(r, c, True, True)}"


def rng(sheet, r1, c1, r2, c2):
    """Référence absolue vers une plage d'une autre feuille."""
    return f"'{sheet}'!{xl_rowcol_to_cell(r1, c1, True, True)}:{xl_rowcol_to_cell(r2, c2, True, True)}"


def a1(r, c):
    """Référence relative sur la même feuille."""
    return xl_rowcol_to_cell(r, c)


def _clean(v):
    """Valeur en cache acceptable par xlsxwriter (pas de NaN / inf)."""
    if isinstance(v, (float, np.floating)) and not np.isfinite(v):
        return ''
    return v


# ---------------------------------------------------------------- calcul (valeurs en cache des formules)
def str_estimates(n_kzh, n_kh, area_zh, area_h, z):
    """Estimateur STR, par zone (post-stratification zone x strate) et national.

    n_kzh : effectifs [classe, zone, strate] ; n_kh : effectifs nationaux [classe, strate] ;
    area_zh : superficies [zone, strate] ; area_h : superficies nationales [strate] ; z : quantile normal.
    """
    with np.errstate(divide='ignore', invalid='ignore'):
        A_z = area_zh.sum(axis=1)
        W_zh = area_zh / A_z[:, None]
        n_zh = n_kzh.sum(axis=0)
        P = n_kzh / n_zh
        p_kz = np.nansum(W_zh * P, axis=-1)
        var_kz = np.nansum(W_zh ** 2 * P * (1 - P) / (n_zh - 1), axis=-1)
        zone = dict(n=n_kzh.sum(axis=-1), p=p_kz, area=p_kz * A_z, se=np.sqrt(var_kz) * A_z)
        zone['ci'] = z * zone['se']

        A = area_h.sum()
        W_h = area_h / A
        n_h = n_kh.sum(axis=0)
        P = n_kh / n_h
        p_k = np.nansum(W_h * P, axis=-1)
        var_k = np.nansum(W_h ** 2 * P * (1 - P) / (n_h - 1), axis=-1)
        nat = dict(n=n_kh.sum(axis=-1), p=p_k, area=p_k * A, se=np.sqrt(var_k) * A)
        nat['ci'] = z * nat['se']

        agg = dict(area=zone['area'].sum(axis=1), se=np.sqrt((zone['se'] ** 2).sum(axis=1)))
        agg['ci'] = z * agg['se']
    return dict(zone=zone, national=nat, sum_zones=agg, A_z=A_z, W_zh=W_zh, n_zh=n_zh, A=A, W_h=W_h, n_h=n_h)


# ---------------------------------------------------------------- styles
class _Styles:
    def __init__(self, wb):
        self.wb = wb
        self._cache = {}
        self.title = self.fmt(bold=True, font_size=18, font_color=GREEN)
        self.sub = self.fmt(italic=True, font_color=GREY)
        self.h2 = self.fmt(bold=True, font_size=12, font_color=GREEN, bottom=2, bottom_color=GREEN)
        self.hdr = self.fmt(bold=True, font_color='white', bg_color=GREEN, align='center', text_wrap=True,
                            border=1, border_color='white')
        self.hdr_l = self.fmt(bold=True, font_color='white', bg_color=GREEN, align='left', text_wrap=True,
                              border=1, border_color='white')
        self.hdr2 = self.fmt(bold=True, font_color=GREEN, bg_color=GREEN_L, align='center', text_wrap=True,
                             border=1, border_color='white')
        self.note = self.fmt(italic=True, font_color=GREY, text_wrap=True, valign='top')
        self.note1 = self.fmt(italic=True, font_color=GREY)
        self.text = self.fmt(text_wrap=True, valign='top')
        self.key = self.fmt(bold=True, font_color=GREEN, border=1, border_color=GRID, bg_color=BAND)
        self.link = self.fmt(font_color='#1D4ED8', underline=1)
        self.label_font = {
            self.h2: dict(bold=True, font_size=12, font_color=GREEN),
            self.hdr: dict(bold=True, font_color='white'),
            self.hdr_l: dict(bold=True, font_color='white'),
            self.hdr2: dict(bold=True, font_color=GREEN),
            self.note: dict(italic=True, font_color=GREY),
            self.note1: dict(italic=True, font_color=GREY),
            self.key: dict(bold=True, font_color=GREEN),
            self.text: dict(),
        }

    def fmt(self, **kw):
        key = tuple(sorted(kw.items()))
        if key not in self._cache:
            self._cache[key] = self.wb.add_format(
                {'font_name': FONT, 'font_size': 10, 'valign': 'vcenter', 'font_color': INK, **kw})
        return self._cache[key]

    def cell(self, num=None, band=False, bold=False, italic=False, color=None, align=None, top=False, inp=False):
        """Format de cellule de tableau ; inp=True : cellule d'entrée modifiable (jaune, texte bleu)."""
        kw = {'border': 1, 'border_color': GRID}
        if num:
            kw['num_format'] = num
        if band:
            kw['bg_color'] = BAND
        if inp:
            kw.update(bg_color=INPUT_BG, font_color=INPUT_INK)
        if bold:
            kw['bold'] = True
        if italic:
            kw['italic'] = True
        if color:
            kw['font_color'] = color
        if align:
            kw['align'] = align
        if top:
            kw.update(top=2, top_color=GREEN)
        return self.fmt(**kw)


# ---------------------------------------------------------------- constructeur
class WorkbookBase:
    """Utilitaires d'écriture partagés par les classeurs (attend self.wb, self.st, self.title, self.subtitle,
    self.confidence)."""

    def rich(self, ws, r, c, text, cell_format, **font):
        """Texte riche : `_x` / `_{...}` en indice, `^x` / `^{...}` en exposant."""
        st = self.st
        base = {'font_name': FONT, 'font_size': 10, 'font_color': INK, **font}
        f_base, f_sub, f_sup = st.fmt(**base), st.fmt(**base, font_script=2), st.fmt(**base, font_script=1)
        parts = []
        for m in RT_TOKEN.finditer(text):
            if m.group(3):
                parts += [f_base, m.group(3)]
            else:
                parts += [f_sub if m.group(1) == '_' else f_sup, m.group(2).strip('{}')]
        if len(parts) == 2:                      # write_rich_string exige au moins deux fragments
            parts += [f_base, ' ']
        ws.write_rich_string(r, c, *parts, cell_format)

    def label(self, ws, r, c, text, f, last_r=None, last_c=None):
        """Libellé (fusionné si last_r/last_c), en texte riche s'il contient des indices."""
        if last_c is not None:
            ws.merge_range(r, c, last_r, last_c, '', f)
        if isinstance(text, str) and re.search(r'[_^]', text):
            self.rich(ws, r, c, text, f, **self.st.label_font.get(f, {'bold': True}))
        else:
            ws.write(r, c, text, f)

    def h2(self, ws, r, c1, c2, text):
        self.label(ws, r, c1, text, self.st.h2, r, c2)

    def formula(self, ws, r, c, f, fmt, value):
        ws.write_formula(r, c, f, fmt, _clean(value))

    def ci_label(self, prefix='Incertitude', suffix=' (ha)'):
        """En-tête dynamique suivant le niveau de confiance, p. ex. « Incertitude 90% (ha) »."""
        return (f'="{prefix} "&TEXT(Niveau_confiance,"0%")&"{suffix}"', f'{prefix} {self.confidence:.0%}{suffix}')

    def sheet(self, name, title, widths, landscape=True):
        ws = self.wb.add_worksheet(name)
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 2)
        for i, w in enumerate(widths, start=1):
            ws.set_column(i, i, w)
        ws.set_row(0, 8)
        ws.write(1, 1, title, self.st.title)
        ws.set_row(1, 28)
        ws.write(2, 1, self.subtitle, self.st.sub)
        if landscape:
            ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.set_footer('&L&8' + self.title.replace('&', '&&') + '&R&8&A - page &P / &N')
        return ws


class _Builder(WorkbookBase):
    def __init__(self, path, plots, area_hz, area_h, confidence, low_n, change_exclude, title, subtitle,
                 data_source, zone_rule, area_source, cols, territory=None):
        self.st = None
        self.lab = territory_labels(territory)
        self.S_NZ = f"{self.lab['Adj']} vs zones"
        self.S_NAT = f"SBAE {self.lab['adj']}"
        self.sheet_descriptions = sheet_descriptions(self.S_NZ, self.S_NAT, self.lab['adj'])
        self.path = path
        self.cols = cols
        sc, zc, cc = cols['stratum'], cols['zone'], cols['class']
        self.plots = plots.reset_index(drop=True)
        self.strata = [int(h) if float(h).is_integer() else h for h in sorted(area_h.index)]
        self.zones = list(area_hz.index)
        self.H, self.Z = len(self.strata), len(self.zones)
        self.A_zh = area_hz.loc[self.zones, self.strata].to_numpy(float)
        self.A_h = area_h.loc[self.strata].to_numpy(float)
        self.confidence, self.low_n = confidence, low_n
        self.zval = round(NormalDist().inv_cdf(1 - (1 - confidence) / 2), 3)     # 1,645 (90 %), 1,960 (95 %)
        self.title, self.subtitle = title, subtitle
        self.data_source, self.zone_rule, self.area_source = data_source, zone_rule, area_source

        classes = sorted(self.plots[cc].dropna().unique())
        n_kh = (self.plots.groupby([cc, sc]).size().unstack(sc, fill_value=0)
                .reindex(index=classes, columns=self.strata, fill_value=0).to_numpy(float))
        n_kzh = np.stack([
            (self.plots[self.plots[zc] == zn].groupby([cc, sc]).size().unstack(sc, fill_value=0)
             .reindex(index=classes, columns=self.strata, fill_value=0).to_numpy(float))
            for zn in self.zones], axis=1)
        est = str_estimates(n_kzh, n_kh, self.A_zh, self.A_h, self.zval)
        order = np.argsort(-est['national']['area'], kind='stable')       # classes : superficie nationale décroissante
        self.classes = [classes[i] for i in order]
        self.K = len(self.classes)
        self.n_kh, self.n_kzh = n_kh[order], n_kzh[order]
        self.est = str_estimates(self.n_kzh, self.n_kh, self.A_zh, self.A_h, self.zval)
        self.change = [k for k in self.classes if k not in change_exclude]
        self._layout()

    # ------------------------------------------------------------ positions des blocs
    def _layout(self):
        H, Z, K = self.H, self.Z, self.K
        L = {}
        # Entrées
        L['conf'], L['z'], L['seuil'] = 5, 6, 7
        L['ah_hdr'], L['ah'], L['wh'], L['nh'] = 10, 11, 12, 13
        L['az_h2'] = 15
        L['az_hdr'] = 16
        L['az0'] = 17
        L['az_tot'] = L['az0'] + Z
        L['w_h2'] = L['az_tot'] + 3
        L['w_hdr'] = L['w_h2'] + 1
        L['w0'] = L['w_hdr'] + 1
        L['nk_h2'] = L['w0'] + Z + 2
        L['nk_hdr'] = L['nk_h2'] + 1
        L['nk0'] = L['nk_hdr'] + 1
        L['nh_row'] = L['nk0'] + K
        L['nkz_h2'] = L['nh_row'] + 3
        L['nkz_hdr1'] = L['nkz_h2'] + 1
        L['nkz_hdr2'] = L['nkz_hdr1'] + 1
        L['nkz0'] = L['nkz_hdr2'] + 1
        L['nhz'] = L['nkz0'] + K
        # Résultats par zone / SBAE national
        L['rz0'] = 5
        L['nat0'] = 5
        # Tableau croisé : 5 blocs empilés
        r = 4
        for name, has_total in [('tc_area', True), ('tc_se', False), ('tc_ci', False), ('tc_pct', False), ('tc_n', True)]:
            L[name] = r + 2                              # première ligne de données
            r = L[name] + K + (1 if has_total else 0) + 2
        # National vs zones
        L['nz0'] = 6
        self.L = L

    def rz_row(self, zi, ki):
        return self.L['rz0'] + zi * self.K + ki

    def nkz_col(self, zi, hi):
        return 2 + zi * self.H + hi

    # ------------------------------------------------------------ construction
    def build(self):
        self.wb = xlsxwriter.Workbook(self.path)
        self.st = _Styles(self.wb)
        self.wb.set_properties({'title': self.title, 'subject': 'SBAE - estimateur stratifié (STR)', 'company': 'FAO'})
        self.wb.set_calc_mode('auto')
        self.wb.set_size(1600, 1000)
        L = self.L
        self.wb.define_name('Niveau_confiance', '=' + ref(S_IN, L['conf'], 2))
        self.wb.define_name('z_IC', '=' + ref(S_IN, L['z'], 2))
        self.wb.define_name('Seuil_n', '=' + ref(S_IN, L['seuil'], 2))
        self.sheet_synthese()
        self.sheet_entrees()
        self.sheet_resultats_zone()
        self.sheet_tableau_croise()
        self.sheet_national_vs_zones()
        self.sheet_national()
        self.sheet_graphiques()
        self.sheet_equations()
        self.sheet_echantillons()
        self.wb.worksheets()[0].activate()
        self.wb.close()

    # ============================================================ Synthèse
    def sheet_synthese(self):
        st, L, e = self.st, self.L, self.est
        ws = self.sheet(S_SYN, 'Estimations de superficie par zone — classeur dynamique', [46, 20, 18, 18, 60])
        r = 4
        self.h2(ws, r, 1, 5, 'Paramètres')
        zones_tot = ref(S_IN, L['az_tot'], 2 + self.H)
        A_tot = ref(S_IN, L['ah'], 2 + self.H)
        params = [
            ('Données de référence (CEO)', self.data_source, None, None),
            ('Échantillons valides', f"=COUNTA('{S_PL}'!A:A)-1", INT, len(self.plots)),
            ('Zones phytogéographiques', self.Z, INT, None),
            ('Superficie cartographiée ' + self.lab['fem'] + '  A (ha)', '=' + A_tot, HA, e['A']),
            ('Superficie cartographiée des zones  Σ A_z (ha)', '=' + zones_tot, HA, e['A_z'].sum()),
            ('Écart  Σ A_z / A − 1', f'={zones_tot}/{A_tot}-1', '+0.000%;-0.000%', e['A_z'].sum() / e['A'] - 1),
            ('Échantillons hors zone', self.zone_rule, None, None),
            ('Niveau de confiance', '=Niveau_confiance', '0%', self.confidence),
            ('Quantile normal  z', '=z_IC', '0.000', self.zval),
            ('Méthode', 'Cochran (1977) ; Olofsson et al. (2014)', None, None),
        ]
        for i, (k, v, nf, cached) in enumerate(params, start=r + 1):
            self.label(ws, i, 1, k, st.key)
            ws.merge_range(i, 2, i, 5, '', st.cell(nf, align='left'))
            if isinstance(v, str) and v.startswith('='):
                self.formula(ws, i, 2, v, st.cell(nf, align='left'), cached)
            elif nf:
                ws.write_number(i, 2, v, st.cell(nf, align='left'))
            else:
                ws.write(i, 2, v, st.cell(align='left'))
        r += len(params) + 2

        self.h2(ws, r, 1, 4, 'Échantillons et superficie par zone')
        r += 1
        for c, h in enumerate(['Zone', 'Échantillons  n_z', 'Superficie  A_z (ha)', 'Part de Σ A_z'], start=1):
            self.label(ws, r, c, h, st.hdr if c > 1 else st.hdr_l)
        first = r + 1
        for zi, zn in enumerate(self.zones):
            rr, band = first + zi, zi % 2 == 1
            self.formula(ws, rr, 1, '=' + ref(S_IN, L['az0'] + zi, 1), st.cell(band=band), zn)
            nhz = rng(S_IN, L['nhz'], self.nkz_col(zi, 0), L['nhz'], self.nkz_col(zi, self.H - 1))
            self.formula(ws, rr, 2, f'=SUM({nhz})', st.cell(INT, band=band), e['n_zh'][zi].sum())
            self.formula(ws, rr, 3, '=' + ref(S_IN, L['az0'] + zi, 2 + self.H), st.cell(HA, band=band), e['A_z'][zi])
            self.formula(ws, rr, 4, f'={a1(rr, 3)}/{a1(first + self.Z, 3)}', st.cell(PCT, band=band),
                         e['A_z'][zi] / e['A_z'].sum())
        rr = first + self.Z
        ws.write(rr, 1, 'Total', st.cell(bold=True, top=True))
        for c, nf, v in [(2, INT, e['n_zh'].sum()), (3, HA, e['A_z'].sum()), (4, PCT, 1.0)]:
            self.formula(ws, rr, c, f'=SUM({a1(first, c)}:{a1(rr - 1, c)})', st.cell(nf, bold=True, top=True), v)
        r = rr + 3

        self.h2(ws, r, 1, 5, 'Mode d\'emploi')
        r += 1
        ws.write(r, 1, 'Cellule d\'entrée (modifiable)', st.cell(inp=True))
        ws.merge_range(r, 2, r, 5, 'Valeur saisie : niveau de confiance, superficies, données des échantillons.', st.text)
        r += 1
        ws.write(r, 1, 'Cellule calculée (formule)', st.cell())
        ws.merge_range(r, 2, r, 5, 'Toutes les autres cellules sont des formules et se recalculent automatiquement.', st.text)
        r += 1
        self.label(ws, r, 1,
                   'Scénarios : modifier le niveau de confiance ou les superficies A_h / A_hz dans « Entrées », ou '
                   'reclasser des échantillons dans « Échantillons » (colonnes strate, zone, classe). Les effectifs, poids, '
                   'estimations, comparaisons et graphiques sont mis à jour. Les noms de zones et de classes de la '
                   'feuille « Entrées » doivent correspondre exactement à ceux de « Échantillons ». Pour revenir aux '
                   "valeurs d'origine, relancer le notebook.", st.note, r + 2, 5)
        r += 4

        self.h2(ws, r, 1, 5, 'Contenu du classeur')
        for i, (name, desc) in enumerate(self.sheet_descriptions, start=r + 1):
            ws.write_url(i, 1, f"internal:'{name}'!A1", st.link, name)
            ws.merge_range(i, 2, i, 5, '', st.text)
            self.label(ws, i, 2, desc, st.text)
        r += len(self.sheet_descriptions) + 2
        ws.merge_range(r, 1, r + 1, 5,
                       f'Les combinaisons zone × classe comptant moins de « Seuil n » échantillons (Entrées ; {self.low_n} par '
                       'défaut) sont signalées dans « Résultats par zone » : leurs erreurs standard et intervalles de '
                       'confiance sont à interpréter avec prudence.', st.note)

    # ============================================================ Entrées
    def sheet_entrees(self):
        st, L, e, H, Z, K = self.st, self.L, self.est, self.H, self.Z, self.K
        sc, zc, cc = (f"'{S_PL}'!${c}:${c}" for c in 'BCD')
        ncols = 2 + Z * H
        ws = self.sheet(S_IN, "Entrées et effectifs d'échantillonnage", [34] + [13] * (Z * H + 1))
        ws.write(3, 1, 'Cellules jaunes : valeurs modifiables. Les autres cellules sont calculées.', st.note1)

        # paramètres
        self.h2(ws, 4, 1, 3, 'Paramètres')
        ws.write(L['conf'], 1, 'Niveau de confiance', st.key)
        ws.write_number(L['conf'], 2, self.confidence, st.cell('0%', inp=True))
        ws.data_validation(L['conf'], 2, L['conf'], 2, {'validate': 'decimal', 'criteria': 'between',
                                                         'minimum': 0.5, 'maximum': 0.999,
                                                         'error_message': 'Niveau de confiance entre 50 % et 99,9 %.'})
        ws.write(L['z'], 1, 'Quantile normal bilatéral  z', st.key)
        self.formula(ws, L['z'], 2, '=ROUND(_xlfn.NORM.S.INV(1-(1-Niveau_confiance)/2),3)', st.cell('0.000'), self.zval)
        ws.write(L['seuil'], 1, 'Seuil « petit effectif »  n', st.key)
        ws.write_number(L['seuil'], 2, self.low_n, st.cell(INT, inp=True))
        ws.write(L['z'], 3, '(90 % → 1,645 ; 95 % → 1,960)', st.note1)

        # strates nationales
        self.h2(ws, L['ah_hdr'] - 1, 1, 2 + H, 'Strates — superficie cartographiée ' + self.lab['fem'])
        ws.write(L['ah_hdr'], 1, '', st.hdr_l)
        for hi, h in enumerate(self.strata):
            ws.write(L['ah_hdr'], 2 + hi, h, st.fmt(bold=True, font_color='white', bg_color=GREEN, align='center',
                                                    border=1, border_color='white', num_format='"Strate "0'))
        ws.write(L['ah_hdr'], 2 + H, 'Total', st.hdr2)
        self.label(ws, L['ah'], 1, 'A_h (ha)', st.key)
        self.label(ws, L['wh'], 1, 'W_h = A_h / A', st.key)
        self.label(ws, L['nh'], 1, 'n_h  (échantillons)', st.key)
        for hi in range(H):
            c = 2 + hi
            ws.write_number(L['ah'], c, self.A_h[hi], st.cell(HA, inp=True))
            self.formula(ws, L['wh'], c, f'={a1(L["ah"], c)}/{xl_rowcol_to_cell(L["ah"], 2 + H, True, True)}',
                         st.cell(W4), e['W_h'][hi])
            self.formula(ws, L['nh'], c, f'={a1(L["nh_row"], c)}', st.cell(INT), e['n_h'][hi])
        for row, nf, v in [(L['ah'], HA, e['A']), (L['wh'], W4, 1.0), (L['nh'], INT, e['n_h'].sum())]:
            self.formula(ws, row, 2 + H, f'=SUM({a1(row, 2)}:{a1(row, 1 + H)})', st.cell(nf, bold=True), v)

        def strata_header(r, first_label, total_label):
            self.label(ws, r, 1, first_label, st.hdr_l)
            for hi in range(H):
                self.formula(ws, r, 2 + hi, '=' + xl_rowcol_to_cell(L['ah_hdr'], 2 + hi, True, True),
                             st.fmt(bold=True, font_color='white', bg_color=GREEN, align='center', border=1,
                                    border_color='white', num_format='"Strate "0'), self.strata[hi])
            self.label(ws, r, 2 + H, total_label, st.hdr2)

        # A_hz
        self.h2(ws, L['az_h2'], 1, 2 + H, 'Superficie cartographiée par zone  A_hz (ha)')
        strata_header(L['az_hdr'], 'Zone', 'A_z')
        for zi, zn in enumerate(self.zones):
            r = L['az0'] + zi
            ws.write_string(r, 1, zn, st.cell(bold=True, inp=True))
            for hi in range(H):
                ws.write_number(r, 2 + hi, self.A_zh[zi, hi], st.cell(HA, inp=True))
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell(HA, bold=True), e['A_z'][zi])
        r = L['az_tot']
        self.label(ws, r, 1, 'Total  Σ A_z', st.cell(bold=True, top=True))
        for hi in range(H + 1):
            c = 2 + hi
            v = self.A_zh[:, hi].sum() if hi < H else e['A_z'].sum()
            self.formula(ws, r, c, f'=SUM({a1(L["az0"], c)}:{a1(r - 1, c)})', st.cell(HA, bold=True, top=True), v)
        if self.area_source:
            ws.merge_range(r + 1, 1, r + 1, 2 + Z * H, f'Source : {self.area_source}', st.note1)

        # W_hz
        self.h2(ws, L['w_h2'], 1, 2 + H, 'Poids  W_hz = A_hz / A_z')
        strata_header(L['w_hdr'], 'Zone', 'Σ_h W_hz')
        for zi, zn in enumerate(self.zones):
            r = L['w0'] + zi
            band = zi % 2 == 1
            self.formula(ws, r, 1, '=' + xl_rowcol_to_cell(L['az0'] + zi, 1, True, True), st.cell(band=band, bold=True), zn)
            for hi in range(H):
                self.formula(ws, r, 2 + hi,
                             f'={a1(L["az0"] + zi, 2 + hi)}/{xl_rowcol_to_cell(L["az0"] + zi, 2 + H, True, True)}',
                             st.cell(W4, band=band), e['W_zh'][zi, hi])
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell(W4, band=band, bold=True), 1.0)

        # n_hk national
        self.h2(ws, L['nk_h2'], 1, 2 + H, "Effectifs nationaux  n_hk  (NB.SI.ENS sur « Échantillons »)")
        strata_header(L['nk_hdr'], 'Classe', 'n_k')
        for ki, k in enumerate(self.classes):
            r = L['nk0'] + ki
            band = ki % 2 == 1
            ws.write_string(r, 1, k, st.cell(bold=True, inp=True))
            for hi in range(H):
                h_ref = xl_rowcol_to_cell(L['ah_hdr'], 2 + hi, True, True)
                k_ref = xl_rowcol_to_cell(r, 1, False, True)
                self.formula(ws, r, 2 + hi, f'=COUNTIFS({sc},{h_ref},{cc},{k_ref})', st.cell(INT, band=band),
                             self.n_kh[ki, hi])
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell(INT, band=band, bold=True),
                         self.n_kh[ki].sum())
        r = L['nh_row']
        self.label(ws, r, 1, 'n_h', st.cell(bold=True, top=True))
        for hi in range(H + 1):
            c = 2 + hi
            v = self.n_kh[:, hi].sum() if hi < H else self.n_kh.sum()
            self.formula(ws, r, c, f'=SUM({a1(L["nk0"], c)}:{a1(r - 1, c)})', st.cell(INT, bold=True, top=True), v)

        # n_hkz
        self.h2(ws, L['nkz_h2'], 1, ncols - 1, "Effectifs par zone  n_hkz  (NB.SI.ENS sur « Échantillons »)")
        ws.merge_range(L['nkz_hdr1'], 1, L['nkz_hdr2'], 1, 'Classe', st.hdr_l)
        for zi, zn in enumerate(self.zones):
            c1, c2 = self.nkz_col(zi, 0), self.nkz_col(zi, H - 1)
            ws.merge_range(L['nkz_hdr1'], c1, L['nkz_hdr1'], c2, '', st.hdr)
            self.formula(ws, L['nkz_hdr1'], c1, '=' + xl_rowcol_to_cell(L['az0'] + zi, 1, True, True), st.hdr, zn)
            for hi in range(H):
                self.formula(ws, L['nkz_hdr2'], c1 + hi, '=' + xl_rowcol_to_cell(L['ah_hdr'], 2 + hi, True, True),
                             st.fmt(bold=True, font_color=GREEN, bg_color=GREEN_L, align='center', border=1,
                                    border_color='white', num_format='"h = "0'), self.strata[hi])
        for ki, k in enumerate(self.classes):
            r = L['nkz0'] + ki
            band = ki % 2 == 1
            self.formula(ws, r, 1, '=' + xl_rowcol_to_cell(L['nk0'] + ki, 1, True, True), st.cell(band=band, bold=True), k)
            for zi in range(Z):
                z_ref = xl_rowcol_to_cell(L['nkz_hdr1'], self.nkz_col(zi, 0), True, True)
                for hi in range(H):
                    c = self.nkz_col(zi, hi)
                    h_ref = xl_rowcol_to_cell(L['nkz_hdr2'], c, True, True)
                    k_ref = xl_rowcol_to_cell(r, 1, False, True)
                    self.formula(ws, r, c, f'=COUNTIFS({sc},{h_ref},{zc},{z_ref},{cc},{k_ref})',
                                 st.cell(INT, band=band), self.n_kzh[ki, zi, hi])
        r = L['nhz']
        self.label(ws, r, 1, 'n_hz', st.cell(bold=True, top=True))
        for zi in range(Z):
            for hi in range(H):
                c = self.nkz_col(zi, hi)
                self.formula(ws, r, c, f'=SUM({a1(L["nkz0"], c)}:{a1(r - 1, c)})', st.cell(INT, bold=True, top=True),
                             e['n_zh'][zi, hi])
        self.label(ws, r + 1, 1, "Une cellule n_hz égale à 0 ou 1 rend l'estimateur non calculable (#DIV/0!) "
                   "pour la zone concernée.", st.note1)
        ws.freeze_panes(4, 2)

    # ============================================================ Résultats par zone
    def _ci_headers(self):
        return [self.ci_label('Incertitude', ' (ha)'), self.ci_label('Incertitude', ' (% de Â)')]

    def sheet_resultats_zone(self):
        st, L, e, H = self.st, self.L, self.est, self.H
        cols = [('Zone', 26), ('Classe', 10), ('Échantillons  n', 11), ('Proportion', 12), ('Superficie  Â (ha)', 15),
                ('Erreur standard (ha)', 14), (None, 14), (None, 12), ('Borne inf. (ha)', 14), ('Borne sup. (ha)', 14),
                ('Part de la zone', 11), ('Remarque', 32)]
        ws = self.sheet(S_RZ, 'Résultats par zone phytogéographique', [w for _, w in cols])
        r0 = L['rz0'] - 1
        ws.set_row(r0, 32)
        ci_h = self._ci_headers()
        for c, (h, _) in enumerate(cols, start=1):
            f = st.hdr if c > 2 else st.hdr_l
            if h is None:
                self.formula(ws, r0, c, ci_h[c - 7][0], f, ci_h[c - 7][1])
            else:
                ws.write(r0, c, h, f)
        z = e['zone']
        for zi in range(self.Z):
            band = zi % 2 == 1
            Wz = rng(S_IN, L['w0'] + zi, 2, L['w0'] + zi, 1 + H)
            NH = rng(S_IN, L['nhz'], self.nkz_col(zi, 0), L['nhz'], self.nkz_col(zi, H - 1))
            Az = ref(S_IN, L['az0'] + zi, 2 + H)
            for ki in range(self.K):
                r = self.rz_row(zi, ki)
                top = ki == 0
                N = rng(S_IN, L['nkz0'] + ki, self.nkz_col(zi, 0), L['nkz0'] + ki, self.nkz_col(zi, H - 1))
                R = lambda c: a1(r, c)
                a, se, ci, n = z['area'][ki, zi], z['se'][ki, zi], z['ci'][ki, zi], z['n'][ki, zi]
                cf = lambda nf, **kw: st.cell(nf, band=band, top=top, **kw)
                if top:
                    self.formula(ws, r, 1, '=' + ref(S_IN, L['az0'] + zi, 1), cf(None, bold=True), self.zones[zi])
                else:
                    ws.write_blank(r, 1, None, cf(None))
                self.formula(ws, r, 2, '=' + ref(S_IN, L['nk0'] + ki, 1), cf(None, bold=True), self.classes[ki])
                self.formula(ws, r, 3, f'=SUM({N})', cf(INT), n)
                self.formula(ws, r, 4, f'=SUMPRODUCT({Wz},{N}/{NH})', cf(PROP), z['p'][ki, zi])
                self.formula(ws, r, 5, f'={Az}*{R(4)}', cf(HA, bold=True), a)
                self.formula(ws, r, 6, f'={Az}*SQRT(SUMPRODUCT({Wz}^2*({N}/{NH})*(1-{N}/{NH})/({NH}-1)))', cf(HA), se)
                self.formula(ws, r, 7, f'=z_IC*{R(6)}', cf(HA), ci)
                self.formula(ws, r, 8, f'=IF({R(5)}>0,{R(7)}/{R(5)},"")', cf(PCT), ci / a if a > 0 else '')
                self.formula(ws, r, 9, f'={R(5)}-{R(7)}', cf(HA), a - ci)
                self.formula(ws, r, 10, f'={R(5)}+{R(7)}', cf(HA), a + ci)
                self.formula(ws, r, 11, f'={R(5)}/{Az}', cf(PCT), a / e['A_z'][zi])
                self.formula(ws, r, 12, f'=IF({R(3)}<Seuil_n,"n < "&Seuil_n&" : interpréter avec prudence","")', cf(None),
                             f'n < {self.low_n} : interpréter avec prudence' if n < self.low_n else '')
        last = self.rz_row(self.Z - 1, self.K - 1)
        ws.freeze_panes(L['rz0'], 3)
        ws.autofilter(r0, 1, last, len(cols))
        ws.conditional_format(L['rz0'], 3, last, 12, {
            'type': 'formula', 'criteria': f'=${xl_col_to_name(3)}{L["rz0"] + 1}<Seuil_n',
            'format': self.wb.add_format({'italic': True, 'font_color': GREY})})
        ws.conditional_format(L['rz0'], 8, last, 8, SCALE_CI)
        ws.conditional_format(L['rz0'], 11, last, 11, {'type': 'data_bar', 'bar_color': '#CFE3D6', 'bar_solid': True,
                                                       'min_type': 'num', 'min_value': 0, 'max_type': 'num', 'max_value': 1})
        self.label(ws, last + 2, 1, 'Incertitude relative (% de Â) : vert < 5 %, jaune ≈ 25 %, rouge > 60 %.   Part de la zone = Â_kz / A_z.',
                   st.note1)

    # ============================================================ Tableau croisé
    def sheet_tableau_croise(self):
        st, L, e, Z, K = self.st, self.L, self.est, self.Z, self.K
        ws = self.sheet(S_TC, 'Tableaux croisés classe × zone', [12] + [17] * Z + [17])
        z, agg = e['zone'], e['sum_zones']
        RZ = lambda zi, ki, c: ref(S_RZ, self.rz_row(zi, ki), c)
        sz = 2 + Z

        def block(key, title, src_col, vals, nf, sum_kind, sum_vals, note=None):
            first = L[key]
            hdr = first - 1
            if isinstance(title, tuple):
                ws.merge_range(hdr - 1, 1, hdr - 1, sz, '', st.h2)
                self.formula(ws, hdr - 1, 1, title[0], st.h2, title[1])
            else:
                self.h2(ws, hdr - 1, 1, sz, title)
            ws.set_row(hdr, 32)
            ws.write(hdr, 1, 'Classe', st.hdr_l)
            for zi in range(Z):
                self.formula(ws, hdr, 2 + zi, '=' + ref(S_IN, L['az0'] + zi, 1), st.hdr, self.zones[zi])
            ws.write(hdr, sz, 'Σ zones', st.hdr2)
            for ki in range(K):
                r, band = first + ki, ki % 2 == 1
                self.formula(ws, r, 1, '=' + ref(S_IN, L['nk0'] + ki, 1), st.cell(band=band, bold=True), self.classes[ki])
                for zi in range(Z):
                    self.formula(ws, r, 2 + zi, '=' + RZ(zi, ki, src_col), st.cell(nf, band=band), vals[ki, zi])
                row = f'{a1(r, 2)}:{a1(r, 1 + Z)}'
                f = {'sum': f'=SUM({row})',
                     'sumsq': f'=SQRT(SUMSQ({row}))',
                     'ci': f'=z_IC*{a1(L["tc_se"] + ki, sz)}',
                     'pct': f'=IF({a1(L["tc_area"] + ki, sz)}>0,{a1(L["tc_ci"] + ki, sz)}/{a1(L["tc_area"] + ki, sz)},"")',
                     }[sum_kind]
                self.formula(ws, r, sz, f, st.cell(nf, band=band, bold=True), sum_vals[ki])
            last = first + K - 1
            if sum_kind == 'sum':
                r = last + 1
                ws.write(r, 1, 'Total', st.cell(bold=True, top=True))
                for c in range(2, sz + 1):
                    v = vals[:, c - 2].sum() if c < sz else np.nansum(sum_vals)
                    self.formula(ws, r, c, f'=SUM({a1(first, c)}:{a1(last, c)})', st.cell(nf, bold=True, top=True), v)
                last = r
            if note:
                self.label(ws, last + 1, 1, note, st.note1)

        with np.errstate(divide='ignore', invalid='ignore'):
            pct = np.where(z['area'] > 0, z['ci'] / z['area'], np.nan)
            pct_agg = np.where(agg['area'] > 0, agg['ci'] / agg['area'], np.nan)
        block('tc_area', 'Superficie estimée  Â_kz (ha)', 5, z['area'], HA, 'sum', agg['area'],
              'La ligne Total de chaque zone est égale à sa superficie cartographiée A_z.')
        block('tc_se', 'Erreur standard  SE(Â_kz) (ha)', 6, z['se'], HA, 'sumsq', agg['se'],
              'Σ zones : √( Σ_z SE(Â_kz)² ) — zones estimées sur des échantillons disjoints.')
        block('tc_ci', (f'="Incertitude à "&TEXT(Niveau_confiance,"0%")&" (ha)"',
                        f"Incertitude à {self.confidence:.0%} (ha)"), 7, z['ci'], HA, 'ci', agg['ci'])
        block('tc_pct', (f'="Incertitude relative à "&TEXT(Niveau_confiance,"0%")&" (% de la superficie estimée)"',
                         f"Incertitude relative à {self.confidence:.0%} (% de la superficie estimée)"),
              8, pct, PCT, 'pct', pct_agg)
        block('tc_n', "Nombre d'échantillons  n_kz", 3, z['n'], INT, 'sum', z['n'].sum(axis=1))
        ws.conditional_format(L['tc_pct'], 2, L['tc_pct'] + K - 1, sz, SCALE_CI)
        ws.freeze_panes(0, 2)

    # ============================================================ National vs zones
    def sheet_national_vs_zones(self):
        st, L, e, K = self.st, self.L, self.est, self.K
        nat, agg = e['national'], e['sum_zones']
        ws = self.sheet(self.S_NZ, 'SBAE ' + self.lab['adj'] + ' vs somme des SBAE par zone', [10, 10, 15, 13, 13, 11, 15, 13, 13, 11, 14, 11, 10])
        r0 = L['nz0'] - 2
        ws.merge_range(r0, 1, r0 + 1, 1, 'Classe', st.hdr_l)
        ws.merge_range(r0, 2, r0 + 1, 2, 'Échantillons  n', st.hdr)
        self.label(ws, r0, 3, 'SBAE ' + self.lab['adj'] + '  Â_k', st.hdr, r0, 6)
        self.label(ws, r0, 7, 'Somme des SBAE par zone  Â_k^Σz', st.hdr, r0, 10)
        ws.merge_range(r0, 11, r0, 13, 'Comparaison', st.hdr)
        ws.set_row(r0 + 1, 32)
        ci_h = self._ci_headers()
        ci_pct = self.ci_label('Incertitude', ' (%)')
        subs = ['Superficie (ha)', 'Erreur std. (ha)', ci_h[0], ci_pct] * 2 + ['Écart Δ (ha)', 'Écart Δ (%)', 'Rapport ρ des SE']
        for c, h in enumerate(subs, start=3):
            if isinstance(h, tuple):
                self.formula(ws, r0 + 1, c, h[0], st.hdr2, h[1])
            else:
                ws.write(r0 + 1, c, h, st.hdr2)
        first = L['nz0']
        for ki in range(K):
            r, band = first + ki, ki % 2 == 1
            R = lambda c: a1(r, c)
            NAT = lambda c: ref(self.S_NAT, L['nat0'] + ki, c)
            TC = lambda key: ref(S_TC, L[key] + ki, 2 + self.Z)
            an, sn, cn = nat['area'][ki], nat['se'][ki], nat['ci'][ki]
            az, sz_, cz = agg['area'][ki], agg['se'][ki], agg['ci'][ki]
            cells = [
                (1, '=' + ref(S_IN, L['nk0'] + ki, 1), None, self.classes[ki], True),
                (2, '=' + NAT(2), INT, nat['n'][ki], False),
                (3, '=' + NAT(4), HA, an, True),
                (4, '=' + NAT(5), HA, sn, False),
                (5, '=' + NAT(6), HA, cn, False),
                (6, f'=IF({R(3)}>0,{R(5)}/{R(3)},"")', PCT, cn / an if an > 0 else '', False),
                (7, '=' + TC('tc_area'), HA, az, True),
                (8, '=' + TC('tc_se'), HA, sz_, False),
                (9, f'=z_IC*{R(8)}', HA, cz, False),
                (10, f'=IF({R(7)}>0,{R(9)}/{R(7)},"")', PCT, cz / az if az > 0 else '', False),
                (11, f'={R(7)}-{R(3)}', DHA, az - an, False),
                (12, f'=IF({R(3)}>0,{R(11)}/{R(3)},"")', DPCT, (az - an) / an if an > 0 else '', False),
                (13, f'=IF({R(4)}>0,{R(8)}/{R(4)},"")', RATIO, sz_ / sn if sn > 0 else '', False),
            ]
            for c, f, nf, v, bold in cells:
                self.formula(ws, r, c, f, st.cell(nf, band=band, bold=bold), v)
        last = first + K - 1
        rr = last + 1
        ws.write(rr, 1, 'Total', st.cell(bold=True, top=True))
        for c, nf, v in [(2, INT, nat['n'].sum()), (3, HA, nat['area'].sum()), (7, HA, agg['area'].sum())]:
            self.formula(ws, rr, c, f'=SUM({a1(first, c)}:{a1(last, c)})', st.cell(nf, bold=True, top=True), v)
        self.formula(ws, rr, 11, f'={a1(rr, 7)}-{a1(rr, 3)}', st.cell(DHA, bold=True, top=True),
                     agg['area'].sum() - nat['area'].sum())
        self.formula(ws, rr, 12, f'={a1(rr, 11)}/{a1(rr, 3)}', st.cell(DPCT, bold=True, top=True),
                     agg['area'].sum() / nat['area'].sum() - 1)
        for c in (4, 5, 6, 8, 9, 10, 13):
            ws.write_blank(rr, c, None, st.cell(top=True))
        ws.freeze_panes(first, 2)
        ws.conditional_format(first, 12, last, 12, {'type': '3_color_scale', 'min_type': 'num', 'min_value': -0.2,
                                                    'mid_type': 'num', 'mid_value': 0, 'max_type': 'num', 'max_value': 0.2,
                                                    'min_color': '#F8CBAD', 'mid_color': '#FFFFFF', 'max_color': '#F8CBAD'})
        ws.conditional_format(first, 13, last, 13, {'type': 'cell', 'criteria': '<', 'value': 1,
                                                    'format': self.wb.add_format({'font_color': '#166534', 'bold': True})})
        ws.conditional_format(first, 13, last, 13, {'type': 'cell', 'criteria': '>', 'value': 1,
                                                    'format': self.wb.add_format({'font_color': '#B91C1C', 'bold': True})})
        for c in (6, 10):
            ws.conditional_format(first, c, last, c, SCALE_CI)
        notes = [
            'Â_k^Σz = Σ_z Â_kz ;  SE(Â_k^Σz) = √( Σ_z SE(Â_kz)² ) — les zones sont estimées sur des échantillons disjoints.',
            'Δ = Â_k^Σz − Â_k ;  ρ = SE(Â_k^Σz) / SE(Â_k) : ρ < 1 (vert) indique que la post-stratification par zone '
            'réduit l\'incertitude.',
            'Les superficies totales diffèrent (A ' + self.lab['fem'] + ' vs Σ A_z des zones, voir « Synthèse ») : cet écart de '
            'couverture explique une partie de Δ.',
        ]
        for i, t in enumerate(notes):
            self.label(ws, rr + 2 + i, 1, t, st.note, rr + 2 + i, 13)

    # ============================================================ SBAE national
    def sheet_national(self):
        st, L, e, H, K = self.st, self.L, self.est, self.H, self.K
        nat = e['national']
        cols = [('Classe', 10), ('Échantillons  n', 11), ('Proportion', 12), ('Superficie  Â (ha)', 15),
                ('Erreur standard (ha)', 14), (None, 14), (None, 12), ('Borne inf. (ha)', 14), ('Borne sup. (ha)', 14)]
        ws = self.sheet(self.S_NAT, 'SBAE ' + self.lab['adj'] + ' — estimateur stratifié (STR)', [w for _, w in cols])
        r0 = L['nat0'] - 1
        ws.set_row(r0, 32)
        ci_h = self._ci_headers()
        for c, (h, _) in enumerate(cols, start=1):
            f = st.hdr if c > 1 else st.hdr_l
            if h is None:
                self.formula(ws, r0, c, ci_h[c - 6][0], f, ci_h[c - 6][1])
            else:
                ws.write(r0, c, h, f)
        W = rng(S_IN, L['wh'], 2, L['wh'], 1 + H)
        NH = rng(S_IN, L['nh_row'], 2, L['nh_row'], 1 + H)
        A = ref(S_IN, L['ah'], 2 + H)
        for ki in range(K):
            r, band = L['nat0'] + ki, ki % 2 == 1
            R = lambda c: a1(r, c)
            N = rng(S_IN, L['nk0'] + ki, 2, L['nk0'] + ki, 1 + H)
            a, se, ci = nat['area'][ki], nat['se'][ki], nat['ci'][ki]
            cells = [
                (1, '=' + ref(S_IN, L['nk0'] + ki, 1), None, self.classes[ki], True),
                (2, f'=SUM({N})', INT, nat['n'][ki], False),
                (3, f'=SUMPRODUCT({W},{N}/{NH})', PROP, nat['p'][ki], False),
                (4, f'={A}*{R(3)}', HA, a, True),
                (5, f'={A}*SQRT(SUMPRODUCT({W}^2*({N}/{NH})*(1-{N}/{NH})/({NH}-1)))', HA, se, False),
                (6, f'=z_IC*{R(5)}', HA, ci, False),
                (7, f'=IF({R(4)}>0,{R(6)}/{R(4)},"")', PCT, ci / a if a > 0 else '', False),
                (8, f'={R(4)}-{R(6)}', HA, a - ci, False),
                (9, f'={R(4)}+{R(6)}', HA, a + ci, False),
            ]
            for c, f, nf, v, bold in cells:
                self.formula(ws, r, c, f, st.cell(nf, band=band, bold=bold), v)
        last = L['nat0'] + K - 1
        rr = last + 1
        ws.write(rr, 1, 'Total', st.cell(bold=True, top=True))
        for c, nf, v in [(2, INT, nat['n'].sum()), (3, PROP, nat['p'].sum()), (4, HA, nat['area'].sum())]:
            self.formula(ws, rr, c, f'=SUM({a1(L["nat0"], c)}:{a1(last, c)})', st.cell(nf, bold=True, top=True), v)
        for c in range(5, 10):
            ws.write_blank(rr, c, None, st.cell(top=True))
        ws.conditional_format(L['nat0'], 7, last, 7, SCALE_CI)
        ws.freeze_panes(L['nat0'], 2)

        r = rr + 3
        self.h2(ws, r, 1, 5, 'Strates : superficie et poids')
        r += 1
        for c, h in enumerate(['Strate  h', 'Échantillons  n_h', 'Superficie  A_h (ha)', 'Poids  W_h'], start=1):
            self.label(ws, r, c, h, st.hdr)
        for hi in range(H):
            rr, band = r + 1 + hi, hi % 2 == 1
            self.formula(ws, rr, 1, '=' + ref(S_IN, L['ah_hdr'], 2 + hi), st.cell(band=band, bold=True, align='center'),
                         self.strata[hi])
            self.formula(ws, rr, 2, '=' + ref(S_IN, L['nh'], 2 + hi), st.cell(INT, band=band), e['n_h'][hi])
            self.formula(ws, rr, 3, '=' + ref(S_IN, L['ah'], 2 + hi), st.cell(HA, band=band), self.A_h[hi])
            self.formula(ws, rr, 4, '=' + ref(S_IN, L['wh'], 2 + hi), st.cell(W4, band=band), e['W_h'][hi])
        rr = r + 1 + H
        ws.write(rr, 1, 'Total', st.cell(bold=True, top=True))
        for c, nf, v in [(2, INT, e['n_h'].sum()), (3, HA, e['A']), (4, W4, 1.0)]:
            self.formula(ws, rr, c, f'=SUM({a1(r + 1, c)}:{a1(rr - 1, c)})', st.cell(nf, bold=True, top=True), v)

    # ============================================================ Graphiques
    def sheet_graphiques(self):
        st, L, e, Z, K = self.st, self.L, self.est, self.Z, self.K
        ws = self.sheet(S_GR, 'Graphiques', [14] + [16] * (2 * Z + 2))
        ws.write(3, 1, 'Les graphiques sont liés aux formules du classeur et se mettent à jour automatiquement.', st.note1)
        conf = f'{self.confidence:.0%}'
        row_g = 5 + ((Z + 1) // 2) * 17 + 17 + 4          # début des données des graphiques
        nC = len(self.change)
        t0 = row_g + 2 * nC + 8                           # titres dynamiques des graphiques
        self.h2(ws, t0 - 1, 1, 6, 'Titres des graphiques (suivent le niveau de confiance)')
        ic_txt = '" (± incertitude à "&TEXT(Niveau_confiance,"0%")&")"'

        def title_cell(i, text_formula, cached):
            self.formula(ws, t0 + i, 1, '=' + text_formula, st.fmt(italic=True, font_color=GREY), cached)
            return [S_GR, t0 + i, 1]

        def style(chart, title, w=620, h=330, legend=True):
            chart.set_title({'name': title, 'name_font': {'size': 11, 'color': GREEN, 'bold': True}})
            chart.set_legend({'position': 'bottom'} if legend else {'none': True})
            chart.set_size({'width': w, 'height': h})
            chart.set_chartarea({'border': {'none': True}})

        def err(sheet, r1, r2, c):
            rg = f"='{sheet}'!${xl_col_to_name(c)}${r1 + 1}:${xl_col_to_name(c)}${r2 + 1}"
            return {'type': 'custom', 'plus_values': rg, 'minus_values': rg, 'line': {'color': '#4B5563', 'width': 0.75}}

        # 1. une figure par zone : toutes les classes (± incertitude)
        for zi in range(Z):
            r1, r2 = self.rz_row(zi, 0), self.rz_row(zi, K - 1)
            ch = self.wb.add_chart({'type': 'bar'})
            ch.add_series({'name': self.zones[zi],
                           'categories': [S_RZ, r1, 2, r2, 2], 'values': [S_RZ, r1, 5, r2, 5],
                           'fill': {'color': CHART_COLORS[0]}, 'border': {'none': True}, 'gap': 60,
                           'y_error_bars': err(S_RZ, r1, r2, 7)})
            ch.set_y_axis({'reverse': True, 'crossing': 'max'})
            ch.set_x_axis({'name': 'Superficie (ha)', 'num_format': '#,##0',
                           'major_gridlines': {'visible': True, 'line': {'color': '#E5E7EB'}}})
            style(ch, title_cell(zi, f'{ref(S_IN, L["az0"] + zi, 1)}&{ic_txt}', f'{self.zones[zi]} (± incertitude à {conf})'),
                  legend=False)
            ws.insert_chart(5 + (zi // 2) * 17, 1 + (zi % 2) * 6, ch, {'x_offset': 5, 'y_offset': 5})

        # données des graphiques de synthèse (formules)
        self.h2(ws, row_g, 1, 2 + 2 * Z, 'Données des graphiques (liées aux autres feuilles)')
        hdr = row_g + 1
        ws.write(hdr, 1, 'Classe', st.hdr_l)
        for zi in range(Z):
            self.formula(ws, hdr, 2 + zi, '=' + ref(S_IN, L['az0'] + zi, 1), st.hdr, self.zones[zi])
            self.formula(ws, hdr, 2 + Z + zi, f'="Incertitude "&{ref(S_IN, L["az0"] + zi, 1)}', st.hdr2, f'Incertitude {self.zones[zi]}')
        d0 = hdr + 1
        for i, k in enumerate(self.change):
            ki = self.classes.index(k)
            r, band = d0 + i, i % 2 == 1
            self.formula(ws, r, 1, '=' + ref(S_IN, L['nk0'] + ki, 1), st.cell(band=band, bold=True), k)
            for zi in range(Z):
                self.formula(ws, r, 2 + zi, '=' + ref(S_TC, L['tc_area'] + ki, 2 + zi), st.cell(HA, band=band),
                             e['zone']['area'][ki, zi])
                self.formula(ws, r, 2 + Z + zi, '=' + ref(S_TC, L['tc_ci'] + ki, 2 + zi), st.cell(HA, band=band),
                             e['zone']['ci'][ki, zi])
        n0 = d0 + nC + 2
        for c, h in enumerate(['Classe', 'SBAE ' + self.lab['adj'], 'Incertitude ' + self.lab['adj'], 'Somme des zones', 'Incertitude zones'], start=1):
            ws.write(n0, c, h, st.hdr if c > 1 else st.hdr_l)
        for i, k in enumerate(self.change):
            ki = self.classes.index(k)
            r, band = n0 + 1 + i, i % 2 == 1
            self.formula(ws, r, 1, '=' + ref(S_IN, L['nk0'] + ki, 1), st.cell(band=band, bold=True), k)
            for c, col, v in [(2, 3, e['national']['area'][ki]), (3, 5, e['national']['ci'][ki]),
                              (4, 7, e['sum_zones']['area'][ki]), (5, 9, e['sum_zones']['ci'][ki])]:
                self.formula(ws, r, c, '=' + ref(self.S_NZ, L['nz0'] + ki, col), st.cell(HA, band=band), v)

        # 2. classes de changement par zone
        ch = self.wb.add_chart({'type': 'column'})
        for zi in range(Z):
            ch.add_series({'name': [S_GR, hdr, 2 + zi],
                           'categories': [S_GR, d0, 1, d0 + nC - 1, 1], 'values': [S_GR, d0, 2 + zi, d0 + nC - 1, 2 + zi],
                           'fill': {'color': CHART_COLORS[zi % len(CHART_COLORS)]}, 'border': {'none': True},
                           'y_error_bars': err(S_GR, d0, d0 + nC - 1, 2 + Z + zi)})
        ch.set_y_axis({'name': 'Superficie (ha)', 'num_format': '#,##0',
                       'major_gridlines': {'visible': True, 'line': {'color': '#E5E7EB'}}})
        style(ch, title_cell(Z, f'"Classes de changement par zone"&{ic_txt}', f'Classes de changement par zone (± incertitude à {conf})'))
        ws.insert_chart(5 + ((Z + 1) // 2) * 17, 1, ch, {'x_offset': 5, 'y_offset': 5})

        # 3. national vs somme des zones
        ch = self.wb.add_chart({'type': 'column'})
        for c, ci_c, color in [(2, 3, '#4C72B0'), (4, 5, '#55A868')]:
            ch.add_series({'name': [S_GR, n0, c],
                           'categories': [S_GR, n0 + 1, 1, n0 + nC, 1], 'values': [S_GR, n0 + 1, c, n0 + nC, c],
                           'fill': {'color': color}, 'border': {'none': True}, 'gap': 80,
                           'y_error_bars': err(S_GR, n0 + 1, n0 + nC, ci_c)})
        ch.set_y_axis({'name': 'Superficie (ha)', 'num_format': '#,##0',
                       'major_gridlines': {'visible': True, 'line': {'color': '#E5E7EB'}}})
        style(ch, title_cell(Z + 1, f'"SBAE {self.lab["adj"]} vs somme des zones"&{ic_txt}',
                             f'SBAE {self.lab["adj"]} vs somme des zones (± incertitude à {conf})'))
        ws.insert_chart(5 + ((Z + 1) // 2) * 17, 7, ch, {'x_offset': 5, 'y_offset': 5})
        ws.set_h_pagebreaks([5 + ((Z + 1) // 2) * 17, row_g])

    # ============================================================ Équations
    def sheet_equations(self):
        st = self.st
        ws = self.sheet(S_EQ, 'Équations des estimateurs', [5, 70, 52, 52])
        F_EQ = st.fmt(border=1, border_color=GRID, indent=1)
        F_EQ_B = st.fmt(border=1, border_color=GRID, indent=1, bg_color=BAND)
        F_NUM = st.fmt(font_color=GREY, align='center', border=1, border_color=GRID)
        F_DESC = st.fmt(text_wrap=True, border=1, border_color=GRID, font_color='#374151')
        F_XL = st.fmt(text_wrap=True, border=1, border_color=GRID, font_color='#1F4E3D', font_name='Consolas', font_size=9)
        sections = [
            ('1 · SBAE ' + self.lab['adj'] + ' — estimateur stratifié (STR)', [
                ('p̂_{hk} = n_{hk} / n_{h}', "Proportion d'échantillon de la classe k dans la strate h", 'Entrées : n_hk, n_h (NB.SI.ENS)'),
                ('W_{h} = A_{h} / A ,   A = Σ_{h} A_{h}', 'Poids en superficie de la strate h', 'Entrées : W_h = A_h / Total'),
                ('p̂_{k} = Σ_{h} W_{h} · p̂_{hk}', 'Proportion de superficie de la classe k', 'SOMMEPROD(W_h ; n_hk / n_h)'),
                ('SE(p̂_{k}) = √[ Σ_{h} W_{h}^{2} · p̂_{hk} (1 − p̂_{hk}) / (n_{h} − 1) ]', 'Erreur standard de la proportion',
                 'RACINE(SOMMEPROD(W_h^2 · p̂_hk · (1 − p̂_hk) / (n_h − 1)))'),
                ('Â_{k} = A · p̂_{k} ,   SE(Â_{k}) = A · SE(p̂_{k})', 'Superficie estimée et son erreur standard', 'SBAE ' + self.lab['adj'] + ' : colonnes Â et SE'),
                ('U(Â_{k}) = z · SE(Â_{k}) ,   IC(Â_{k}) = Â_{k} ± U(Â_{k})', 'Incertitude U (z = 1,645 pour 90 %) et intervalle de confiance', 'z_IC = ARRONDI(LOI.NORMALE.STANDARD.INVERSE.N(1 − (1 − niveau) / 2) ; 3)'),
            ]),
            ('2 · SBAE par zone — post-stratification zone × strate', [
                ('W_{hz} = A_{hz} / A_{z} ,   A_{z} = Σ_{h} A_{hz}', 'Poids de la strate h dans la zone z', 'Entrées : bloc W_hz'),
                ('p̂_{hkz} = n_{hkz} / n_{hz}', 'Proportion de la classe k parmi les échantillons de la strate h situés dans la zone z',
                 'Entrées : n_hkz, n_hz (NB.SI.ENS)'),
                ('p̂_{kz} = Σ_{h} W_{hz} · p̂_{hkz}', 'Proportion de superficie de la classe k dans la zone z', 'SOMMEPROD(W_hz ; n_hkz / n_hz)'),
                ('SE(p̂_{kz}) = √[ Σ_{h} W_{hz}^{2} · p̂_{hkz} (1 − p̂_{hkz}) / (n_{hz} − 1) ]', 'Erreur standard de la proportion dans la zone z',
                 'RACINE(SOMMEPROD(W_hz^2 · p̂_hkz · (1 − p̂_hkz) / (n_hz − 1)))'),
                ('Â_{kz} = A_{z} · p̂_{kz} ,   SE(Â_{kz}) = A_{z} · SE(p̂_{kz})', 'Superficie estimée dans la zone z et son erreur standard',
                 'Résultats par zone : colonnes Â et SE'),
                ('Σ_{k} Â_{kz} = A_{z}', "La somme des classes d'une zone est égale à sa superficie cartographiée", 'Tableau croisé : ligne Total'),
            ]),
            ('3 · Agrégation des zones et comparaison avec le SBAE ' + self.lab['adj'], [
                ('Â_{k}^{Σz} = Σ_{z} Â_{kz}', 'Somme des estimations par zone', 'SOMME(Â_kz) — Tableau croisé, colonne Σ zones'),
                ('SE(Â_{k}^{Σz}) = √[ Σ_{z} SE(Â_{kz})^{2} ]', "Zones estimées sur des échantillons disjoints : les variances s'additionnent",
                 'RACINE(SOMME.CARRES(SE(Â_kz)))'),
                ('Δ_{k} = Â_{k}^{Σz} − Â_{k} ,   Δ_{k} (%) = 100 · Δ_{k} / Â_{k}', 'Écart entre la somme des zones et le SBAE ' + self.lab['adj'],
                 self.S_NZ + ' : colonnes Δ'),
                ('ρ_{k} = SE(Â_{k}^{Σz}) / SE(Â_{k})', "Rapport des erreurs standard (ρ < 1 : la post-stratification par zone réduit l'incertitude)",
                 self.S_NZ + ' : colonne ρ'),
            ]),
        ]
        r, no = 4, 1
        for title, eqs in sections:
            self.h2(ws, r, 1, 4, title)
            ws.set_row(r, 22)
            r += 1
            for c, h in enumerate(['#', 'Équation', 'Description', 'Mise en œuvre dans le classeur'], start=1):
                ws.write(r, c, h, st.hdr if c == 1 else st.hdr_l)
            r += 1
            for i, (expr, desc, xl) in enumerate(eqs):
                ws.set_row(r, 30)
                ws.write(r, 1, f'({no})', F_NUM)
                self.rich(ws, r, 2, expr, F_EQ_B if i % 2 else F_EQ, font_name='Cambria', font_size=13)
                ws.write(r, 3, desc, F_DESC)
                self.rich(ws, r, 4, xl, F_XL, font_name='Consolas', font_size=9, font_color=GREEN)
                r, no = r + 1, no + 1
            r += 1
        self.h2(ws, r, 1, 4, 'Symboles')
        r += 1
        for c, h in enumerate(['', 'Symbole', 'Signification'], start=1):
            ws.write(r, c, h, st.hdr_l)
        ws.merge_range(r, 3, r, 4, 'Signification', st.hdr_l)
        r += 1
        symbols = [
            ('h = 1, …, H', f"strate du plan d'échantillonnage (H = {self.H})"),
            ('z', 'zone phytogéographique'),
            ('k', "classe d'activité / de changement"),
            ('A_{h} ,  A', 'superficie cartographiée de la strate h ; superficie totale ' + self.lab['fem']),
            ('A_{hz} ,  A_{z}', 'superficie cartographiée de la strate h dans la zone z ; superficie de la zone z'),
            ('n_{h} ,  n_{hk}', 'échantillons de la strate h ; parmi eux, échantillons de la classe k (référence CEO)'),
            ('n_{hz} ,  n_{hkz}', 'échantillons de la strate h situés dans la zone z ; parmi eux, échantillons de la classe k'),
            ('SE ,  U ,  IC ,  z', 'erreur standard ; incertitude (z · SE) ; intervalle de confiance (Â ± U) ; quantile normal (Entrées)'),
        ]
        for i, (sym, desc) in enumerate(symbols):
            ws.set_row(r, 24)
            ws.write_blank(r, 1, None, F_NUM)
            self.rich(ws, r, 2, sym, F_EQ_B if i % 2 else F_EQ, font_name='Cambria', font_size=12)
            ws.merge_range(r, 3, r, 4, desc, F_DESC)
            r += 1
        ws.merge_range(r + 1, 1, r + 2, 4,
                       'Références : Cochran, W. G. (1977) Sampling Techniques, 3e éd., Wiley. — Olofsson, P. et al. (2014) '
                       'Good practices for estimating area and assessing accuracy of land change, Remote Sensing of '
                       'Environment 148, 42-57.', st.note)

    # ============================================================ Échantillons
    def sheet_echantillons(self):
        st, c = self.st, self.cols
        ws = self.wb.add_worksheet(S_PL)
        # Ligne 1 = en-têtes : les colonnes entières B:D sont lues par NB.SI.ENS (feuille « Entrées »)
        main = [c['id'], c['stratum'], c['zone'], c['class']]
        extra = [x for x in c['extra'] if x in self.plots.columns]
        headers = ['Échantillon', 'Strate', 'Zone', 'Classe'] + extra
        widths = [12, 9, 24, 10] + [13] * len(extra)
        for j, (h, w) in enumerate(zip(headers, widths)):
            ws.write(0, j, h, st.hdr if j else st.hdr_l)
            ws.set_column(j, j, w)
        f_id = st.fmt(font_color=GREY)
        f_in = st.fmt(bg_color=INPUT_BG, font_color=INPUT_INK)
        f_num = st.fmt(num_format='0.00000', font_color=GREY)
        data = self.plots
        ids = data[main[0]].tolist() if main[0] in data.columns else list(range(1, len(data) + 1))
        ws.write_column(1, 0, ids, f_id)
        ws.write_column(1, 1, [int(v) if float(v).is_integer() else v for v in data[main[1]]], f_in)
        ws.write_column(1, 2, data[main[2]].astype(str).tolist(), f_in)
        ws.write_column(1, 3, data[main[3]].astype(str).tolist(), f_in)
        for j, col in enumerate(extra, start=4):
            ws.write_column(1, j, [_clean(v) for v in data[col]], f_num)
        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, len(data), len(headers) - 1)
        ws.data_validation(1, 1, len(data) + 5000, 1, {'validate': 'list', 'source': [str(h) for h in self.strata]})
        ws.data_validation(1, 2, len(data) + 5000, 2, {'validate': 'list',
                                                        'source': '=' + rng(S_IN, self.L['az0'], 1, self.L['az0'] + self.Z - 1, 1)})
        ws.data_validation(1, 3, len(data) + 5000, 3, {'validate': 'list',
                                                        'source': '=' + rng(S_IN, self.L['nk0'], 1, self.L['nk0'] + self.K - 1, 1)})

    # ------------------------------------------------------------ résultats Python (contrôle)
    def frames(self):
        e = self.est
        z = e['zone']
        zone = pd.DataFrame([
            {'zone': zn, 'activity': k, 'n_sample': z['n'][ki, zi], 'proportion': z['p'][ki, zi],
             'area_ha': z['area'][ki, zi], 'se_ha': z['se'][ki, zi], 'ci_ha': z['ci'][ki, zi]}
            for zi, zn in enumerate(self.zones) for ki, k in enumerate(self.classes)])
        n = e['national']
        national = pd.DataFrame({'activity': self.classes, 'n_sample': n['n'], 'proportion': n['p'],
                                 'area_ha': n['area'], 'se_ha': n['se'], 'ci_ha': n['ci']})
        s = e['sum_zones']
        sum_zones = pd.DataFrame({'activity': self.classes, 'area_ha': s['area'], 'se_ha': s['se'], 'ci_ha': s['ci']})
        return {'zone': zone, 'national': national, 'sum_zones': sum_zones}


def build_workbook(path, plots, area_hz, area_h, confidence=0.90, low_n=10, change_exclude=('SNF', 'SF'),
                   title="Côte d'Ivoire — SBAE par zone phytogéographique",
                   subtitle=None, data_source='', zone_rule='', area_source='',
                   stratum_col='strata', zone_col='zone', class_col='activity', id_col='plotid',
                   extra_cols=('center_lon', 'center_lat'), territory=None):
    """Écrit le classeur Excel dynamique et renvoie les estimations calculées en Python (pour contrôle).

    plots   : DataFrame, une ligne par échantillon valide (colonnes strate, zone, classe ; identifiant facultatif).
    area_hz : DataFrame indexé par zone, une colonne par strate (superficies cartographiées, ha).
    area_h  : Series indexée par strate (superficies cartographiées nationales, ha).
    Les zones et les strates sont prises dans area_hz / area_h, les classes dans plots.
    territory : None (estimateur global « national ») ou nom du territoire couvert par area_h (ex. 'PRE') ;
                ne change que les libellés et les noms de feuilles.
    """
    if subtitle is None:
        subtitle = (f"Projet PROMIRE / ONU-REDD · SBAE, estimateur stratifié (STR) · "
                    f"généré le {datetime.date.today():%d/%m/%Y}")
    cols = dict(stratum=stratum_col, zone=zone_col, **{'class': class_col}, id=id_col, extra=list(extra_cols))
    b = _Builder(path, plots, area_hz, area_h, confidence, low_n, set(change_exclude), title, subtitle,
                 data_source, zone_rule, area_source, cols, territory)
    b.build()
    out = b.frames()
    out['path'] = path
    return out
