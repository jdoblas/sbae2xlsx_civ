"""Classeur Excel « vivant » des émissions et absorptions forestières, par zone phytogéographique et au niveau
national, à partir des échantillons SBAE et des facteurs d'émission nationaux (NERF 2024).

Composantes :

    DEF      déforestation - biomasse et matière organique morte (FE net = stock forestier − stock post-conversion)
    DEF SOC  déforestation - carbone organique des sols minéraux (GIEC éq. 2.25, D = 20 ans)
    DEF CH4  déforestation - combustion de la biomasse, CH₄ et N₂O (GIEC éq. 2.27 :
    DEF N2O    L = A_brûlée · MB · Cf · Gef · 10⁻³, puis × PRG ; MB = AGB + litière + bois mort, en tms/ha)
    DEG      dégradation - FE = stock de la forêt initiale − stock de la forêt finale
    RSCF     renforcement des stocks de carbone forestier - gains (non-forêt -> forêt), taux de croissance annuel

Chaîne de calcul (toutes les cellules de résultats sont des formules Excel) :

    Échantillons -> Correspondances -> Activité DEF / DEG / RSCF (estimateur stratifié par zone)
      -> Stocks de carbone (stocks, biomasse combustible, facteurs SOC, taux de croissance)
      -> Émissions DEF, Émissions DEF SOC, Émissions DEF CH4 N2O, Émissions DEG, Absorptions RSCF
      -> Bilan, Résultats, Contrôles

Les entrées modifiables sont en jaune. Les valeurs calculées en Python (EmissionModel) sont écrites en cache des
formules et renvoyées pour contrôle.

Références : GIEC (2006, 2019) vol. 1 ch. 3, vol. 4 ch. 2, 4 et 5 ; Cochran (1977) ; Olofsson et al. (2014).
"""
import datetime
from statistics import NormalDist

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from sbae_excel import (BAND, CHART_COLORS, GREEN, GREY, GRID, HA, INPUT_BG, INPUT_INK, INT, WorkbookBase, _clean,
                        _Styles, a1, ref, rng, territory_labels)

# ---------------------------------------------------------------- feuilles
S_SYN = 'Synthèse'
S_PAR = 'Paramètres'
S_COR = 'Correspondances'
S_STK = 'Stocks de carbone'
S_ADEF = 'Activité DEF'
S_ADEG = 'Activité DEG'
S_AGAIN = 'Activité RSCF'
S_EM = 'Émissions DEF'
S_SOC = 'Émissions DEF SOC'
S_FIRE = 'Émissions DEF CH4 N2O'
S_DEG = 'Émissions DEG'
S_RSCF = 'Absorptions RSCF'
S_BIL = 'Bilan'
S_RES = 'Résultats'
S_CTL = 'Contrôles'
S_GR = 'Graphiques'
S_EQ = 'Équations'
S_ECH = 'Échantillons'

SHEETS = [
    (S_PAR, 'Paramètres généraux, paramètres des sols, superficies A_hz, poids W_hz et effectifs n_hz'),
    (S_COR, 'Correspondance des classes CEO avec les catégories de stocks, de sols et de gains (modifiable)'),
    (S_STK, 'Stocks de carbone, facteurs d\'émission, facteurs SOC et taux de croissance'),
    (S_ADEF, 'Superficie déforestée par zone et par transition (estimateur stratifié)'),
    (S_ADEG, 'Superficie dégradée par zone et par transition'),
    (S_AGAIN, 'Superficie de gain (non-forêt → forêt) par zone et par catégorie'),
    (S_EM, 'Déforestation : émissions de la biomasse et de la MOM (FE brut et net)'),
    (S_SOC, 'Déforestation : émissions du carbone organique des sols minéraux'),
    (S_FIRE, 'Déforestation : émissions de CH₄ et de N₂O de la combustion de la biomasse'),
    (S_DEG, 'Dégradation : émissions'),
    (S_RSCF, 'Renforcement des stocks de carbone forestier : absorptions'),
    (S_BIL, 'Bilan : moyennes annuelles par composante et série par année'),
    (S_RES, 'Résultats par zone, par catégorie forestière et par utilisation post-déforestation'),
    (S_CTL, 'Contrôles de cohérence et estimateur direct'),
    (S_GR, 'Graphiques (± incertitude)'),
    (S_EQ, 'Équations et mise en œuvre dans le classeur'),
    (S_ECH, 'Échantillons valides : strate, zone, activité, classes CEO 2021 et 2025 (modifiables)'),
]

UNASSIGNED = 'NON ATTRIBUÉ'
PCT1 = '0.0%;-0.0%;"–"'
HA0 = '#,##0;-#,##0;"–"'
T0 = '#,##0;-#,##0;"–"'
EF2 = '#,##0.00;-#,##0.00;"–"'
F3 = '0.000;-0.000;"–"'
DM2 = '0.00'
DIFF = '0.000000;-0.000000;"0"'
SCI = '0.00E+00;-0.00E+00;"–"'
POOLS = ['agb', 'bgb', 'litter', 'dw']
COMPONENTS = [('def', 'Déforestation — biomasse et MOM (FE net)'), ('soc', 'Déforestation — sols minéraux'),
              ('ch4', 'Déforestation — combustion de la biomasse, CH₄'),
              ('n2o', 'Déforestation — combustion de la biomasse, N₂O'),
              ('deg', 'Dégradation'), ('tot', 'Total des émissions'), ('rscf', 'Renforcement des stocks (RSCF)'),
              ('net', 'Bilan net (émissions + absorptions)')]
EMISSIONS = ['def', 'soc', 'ch4', 'n2o', 'deg']     # composantes du total des émissions
GASES = [('ch4', 'CH₄'), ('n2o', 'N₂O')]
FIRE_DEFAULT = 0.5      # fraction brûlée d'une transition absente de la table « fire_fraction »


def z_of(conf):
    """Quantile normal bilatéral arrondi à 3 décimales (1,645 pour 90 %, 1,960 pour 95 %)."""
    return round(NormalDist().inv_cdf(1 - (1 - conf) / 2), 3)


def u_ratio(ux2, total):
    return np.sqrt(ux2) / abs(total) if total else 0.0


# ================================================================ modèle (valeurs en cache des formules)
class EmissionModel:
    """Calcule en Python toutes les grandeurs écrites en formules dans le classeur."""

    def __init__(self, samples, area_hz, cats, maps, forest_stocks, post_stocks, soc_factors, removal_rates, p,
                 fire_fraction=None):
        # combustion de la biomasse : valeurs du NERF 2024 par défaut
        p = dict(dict(cf_litter=0.37, comb=0.32, u_comb=0.0, gef_ch4=6.8, u_gef_ch4=0.0, gef_n2o=0.2, u_gef_n2o=0.0,
                      gwp_ch4=28.0, gwp_n2o=265.0, u_fire=0.0), **p)
        self.p = p
        self.fire_fraction = dict(fire_fraction or {})
        self.zones = list(area_hz.index)
        self.strata = sorted(area_hz.columns)
        self.A_zh = area_hz.loc[self.zones, self.strata].to_numpy(float)
        self.A_z = self.A_zh.sum(axis=1)
        self.W_zh = self.A_zh / self.A_z[:, None]
        self.forest_cats, self.post_cats, self.gain_cats = list(cats['forest']), list(cats['post']), list(cats['gain'])
        self.forest_map, self.post_map, self.gain_map = dict(maps['forest']), dict(maps['post']), dict(maps['gain'])
        self.z, self.z_fe = z_of(p['confidence']), z_of(p['ef_confidence'])
        self.k_fe = self.z / self.z_fe
        self.T = p['final_year'] - p['initial_year']
        self.coh = (self.T + 1) / 2
        self.cf, self.cf_low = p['cf'], p['cf_low']
        self.u_cf = abs(self.cf_low - self.cf) / self.cf

        s = samples.copy()
        act = s['activity']

        def mapped(mask, classes, mapping):
            return np.where(mask, classes.map(mapping).fillna(UNASSIGNED), '')

        s['fcat_i'] = mapped(act.isin(['Def', 'Deg']), s['class_initial'], self.forest_map)
        s['pcat'] = mapped(act == 'Def', s['class_final'], self.post_map)
        s['fcat_f'] = mapped(act == 'Deg', s['class_final'], self.forest_map)
        s['gcat'] = mapped(act == 'Gain', s['class_final'], self.gain_map)
        s['key_def'] = np.where(act == 'Def', s['zone'] + '|' + s['fcat_i'] + '|' + s['pcat'], '')
        s['key_deg'] = np.where(act == 'Deg', s['zone'] + '|' + s['fcat_i'] + '|' + s['fcat_f'], '')
        s['key_gain'] = np.where(act == 'Gain', s['zone'] + '|' + s['gcat'], '')
        self.samples = s
        self.n_zh = (s.groupby(['zone', 'strata']).size().unstack(fill_value=0)
                     .reindex(index=self.zones, columns=self.strata, fill_value=0).to_numpy(float))

        self.forest_stocks = self._forest_stocks(forest_stocks)
        self.post_stocks = self._post_stocks(post_stocks)
        self.soc_factors = self._soc_factors(soc_factors)
        self.removal_rates = self._removal_rates(removal_rates)

        self.def_rows = self._def_rows()
        self.soc_rows = self._soc_rows()
        self.fire_rows = self._fire_rows()
        self.deg_rows = self._deg_rows()
        self.gain_rows = self._gain_rows()
        self.zone_act = {a: self._zone_direct(a) for a in ('Def', 'Deg', 'Gain')}
        self._direct()
        self._bilan()

    # ------------------------------------------------------------ stocks et facteurs
    def _pool_u(self, u):
        return np.sqrt(u ** 2 + self.u_cf ** 2)

    def _forest_stocks(self, fs):
        out = {}
        rows = fs.to_dict('records')
        by_key = {(r['category'], r['zone']): r for r in rows}
        for r in rows:
            o = dict(r)
            comps = ('lying', 'stump', 'standing', 'bgb_standing')
            o['dw'] = sum(r[f'dw_{c}'] for c in comps)
            o['u_dw'] = (np.sqrt(sum((r[f'dw_{c}'] * r[f'e_{c}']) ** 2 for c in comps)) / o['dw']) if o['dw'] else 0.0
            if r['litter_rule'] == 'ipcc':
                o['litter'] = r['litter_tc']
                o['u_litter'] = (r['litter_upper_tc'] - r['litter_tc']) / r['litter_tc']
            elif r['litter_rule'] == 'scaled':
                d = by_key[(r['litter_ref'], r['zone'])]
                o['litter'] = d['litter_tc'] * r['agb'] / d['agb']
                o['u_litter'] = np.sqrt(d['u_agb'] ** 2 + r['u_agb'] ** 2
                                        + ((d['litter_upper_tc'] - d['litter_tc']) / d['litter_tc']) ** 2)
            else:
                o['litter'], o['u_litter'] = 0.0, 0.0
            o['c_agb'], o['uc_agb'] = r['agb'] * self.cf, self._pool_u(r['u_agb'])
            o['c_bgb'], o['uc_bgb'] = r['bgb'] * self.cf, self._pool_u(r['u_bgb'])
            o['c_litter'], o['uc_litter'] = o['litter'], o['u_litter']
            o['c_dw'], o['uc_dw'] = o['dw'] * self.cf, self._pool_u(o['u_dw'])
            o['c_tot'] = sum(o[f'c_{q}'] for q in POOLS)
            o['u_tot'] = np.sqrt(sum((o[f'uc_{q}'] * o[f'c_{q}']) ** 2 for q in POOLS)) / o['c_tot']
            o['ef'] = o['c_tot'] * 44 / 12
            # biomasse combustible (tms/ha) : AGB + litière + bois mort, sans les racines (GIEC 2006 éq. 2.27)
            o['litter_dm'] = o['litter'] / self.p['cf_litter']
            o['mb'] = r['agb'] + o['litter_dm'] + o['dw']
            o['u_mb'] = (np.sqrt((r['u_agb'] * r['agb']) ** 2 + (o['u_litter'] * o['litter_dm']) ** 2
                                 + (o['u_dw'] * o['dw']) ** 2) / o['mb']) if o['mb'] else 0.0
            out[(r['category'], r['zone'])] = o
        return out

    def _post_stocks(self, ps):
        out = {}
        rows = ps.to_dict('records')
        by_cat = {r['category']: r for r in rows}
        for r in rows:
            o = dict(r)
            if r['rule'] == 'mean':
                src = [by_cat[c] for c in r['mean_of']]
                o['agb'] = np.mean([x['agb'] for x in src])
                o['bgb'] = np.mean([x['bgb'] for x in src])
                sa, sb = sum(x['agb'] for x in src), sum(x['bgb'] for x in src)
                o['u_agb'] = np.sqrt(sum((x['agb'] * x['u_agb']) ** 2 for x in src)) / sa if sa else 0.0
                o['u_bgb'] = np.sqrt(sum((x['bgb'] * x['u_bgb']) ** 2 for x in src)) / sb if sb else 0.0
            o['c_agb'], o['uc_agb'] = o['agb'] * self.cf, self._pool_u(o['u_agb'])
            o['c_bgb'], o['uc_bgb'] = o['bgb'] * self.cf, self._pool_u(o['u_bgb'])
            o['c_tot'] = o['c_agb'] + o['c_bgb']
            o['u_tot'] = (np.sqrt((o['uc_agb'] * o['c_agb']) ** 2 + (o['uc_bgb'] * o['c_bgb']) ** 2) / o['c_tot']
                          if o['c_tot'] else 0.0)
            o['ef'] = o['c_tot'] * 44 / 12
            out[r['category']] = o
        return out

    def _soc_factors(self, sf):
        out = {}
        for r in sf.to_dict('records'):
            o = dict(r)
            o['f'] = r['flu'] * r['fmg'] * r['fi']
            o['u_f'] = np.sqrt(r['u_flu'] ** 2 + r['u_fmg'] ** 2 + r['u_fi'] ** 2) if o['f'] else 0.0
            out[r['category']] = o
        return out

    def _removal_rates(self, rr):
        out = {}
        for r in rr.to_dict('records'):
            o = dict(r)
            o['g'] = -r['rate'] * (1 + r['r']) * 44 / 12
            out[(r['category'], r['zone'])] = o
        return out

    # ------------------------------------------------------------ estimateur stratifié par zone
    def _str(self, zi, counts):
        with np.errstate(divide='ignore', invalid='ignore'):
            P = counts / self.n_zh[zi]
            W = self.W_zh[zi]
            p = np.nansum(W * P)
            var = np.nansum(W ** 2 * P * (1 - P) / (self.n_zh[zi] - 1))
        area = p * self.A_z[zi]
        se = np.sqrt(var) * self.A_z[zi]
        return p, area, se, (self.z * se / area if area > 0 else 0.0)

    def _estimate(self, keycol, items):
        """items : liste de dict (zone, zi, key, + libellés) -> DataFrame avec superficies et incertitudes."""
        s = self.samples
        cnt = (s[s[keycol] != ''].groupby([keycol, 'strata']).size().unstack(fill_value=0)
               .reindex(columns=self.strata, fill_value=0))
        out = []
        for it in items:
            n_h = cnt.loc[it['key']].to_numpy(float) if it['key'] in cnt.index else np.zeros(len(self.strata))
            p, area, se, u = self._str(it['zi'], n_h)
            out.append(dict(it, n_h=n_h, n=n_h.sum(), p=p, area=area, se=se, u_area=u, pm_area=u * area,
                            area_year=area / self.T))
        return pd.DataFrame(out)

    def _zone_direct(self, act):
        s = self.samples
        cnt = (s[s.activity == act].groupby(['zone', 'strata']).size().unstack(fill_value=0)
               .reindex(index=self.zones, columns=self.strata, fill_value=0))
        out = []
        for zi, zn in enumerate(self.zones):
            n_h = cnt.loc[zn].to_numpy(float)
            p, area, se, u = self._str(zi, n_h)
            out.append(dict(zone=zn, n_h=n_h, n=n_h.sum(), p=p, area=area, se=se, u_area=u, area_year=area / self.T))
        df = pd.DataFrame(out)
        nat_area, nat_se = df.area.sum(), np.sqrt((df.se ** 2).sum())
        df.attrs['national'] = dict(area=nat_area, se=nat_se, u_area=self.z * nat_se / nat_area if nat_area else 0.0,
                                    area_year=nat_area / self.T, n_h=sum(df.n_h), n=df.n.sum())
        return df

    # ------------------------------------------------------------ composantes
    def _def_rows(self):
        items = [dict(zone=zn, zi=zi, forest=f, post=u, key=f'{zn}|{f}|{u}')
                 for zi, zn in enumerate(self.zones) for f in self.forest_cats for u in self.post_cats]
        rows = self._estimate('key_def', items)
        k, T = self.k_fe, self.T
        rec = []
        for _, r in rows.iterrows():
            fs, ps = self.forest_stocks[(r.forest, r.zone)], self.post_stocks[r.post]
            efg, ug, st, us = fs['ef'], fs['u_tot'] * k, ps['ef'], ps['u_tot'] * k
            efn = efg - st
            un = np.sqrt((ug * efg) ** 2 + (us * st) ** 2) / abs(efn) if efn != 0 else 0.0
            eg, en = r.area * efg, r.area * efn
            ueg = np.sqrt(r.u_area ** 2 + ug ** 2) if r.area > 0 else 0.0
            uen = np.sqrt(r.u_area ** 2 + un ** 2) if r.area > 0 else 0.0
            rec.append(dict(ef_gross=efg, u_ef_gross=ug, stock_post=st, u_stock_post=us, ef_net=efn, u_ef_net=un,
                            em_gross=eg, u_em_gross=ueg, pm_gross=ueg * eg, ux2_gross=(ueg * eg) ** 2,
                            em_net=en, u_em_net=uen, pm_net=uen * en, ux2_net=(uen * en) ** 2,
                            em_gross_year=eg / T, em_net_year=en / T))
        return pd.concat([rows, pd.DataFrame(rec)], axis=1)

    def _soc_rows(self):
        k, T, p = self.k_fe, self.T, self.p
        rec = []
        for _, r in self.def_rows.iterrows():
            sf = self.soc_factors[r.post]
            f, uf = sf['f'], sf['u_f'] * k
            df_ = f - 1
            u_df = uf * f / abs(df_) if df_ != 0 else 0.0
            dsoc = p['socref'] * df_
            u_dsoc = np.sqrt((p['u_socref'] * k) ** 2 + u_df ** 2) if df_ != 0 else 0.0
            rate = -dsoc / p['soc_d'] * 44 / 12
            em = r.area * rate * self.coh
            u = np.sqrt(r.u_area ** 2 + u_dsoc ** 2) if r.area > 0 else 0.0
            rec.append(dict(zone=r.zone, zi=r.zi, forest=r.forest, post=r.post, area=r.area, u_area=r.u_area,
                            socref=p['socref'], f=f, u_f=uf, dsoc=dsoc, u_dsoc=u_dsoc, rate=rate, em=em, u_em=u,
                            pm=u * abs(em), ux2=(u * em) ** 2, em_year=em / T))
        return pd.DataFrame(rec)

    def _fire_rows(self):
        """Combustion de la biomasse lors de la déforestation : L = A · fraction brûlée · MB · Cf · Gef · 10⁻³ (t gaz)."""
        k, T, p = self.k_fe, self.T, self.p
        rec = []
        for _, r in self.def_rows.iterrows():
            fs = self.forest_stocks[(r.forest, r.zone)]
            frac = self.fire_fraction.get((r.forest, r.post), FIRE_DEFAULT)
            mb, u_mb = fs['mb'], fs['u_mb'] * k
            burnt = r.area * frac
            o = dict(zone=r.zone, zi=r.zi, forest=r.forest, post=r.post, key=r.key, area=r.area, u_area=r.u_area,
                     frac=frac, burnt=burnt, mb=mb, u_mb=u_mb, comb=p['comb'])
            for g, _ in GASES:
                t = burnt * mb * p['comb'] * p[f'gef_{g}'] * 1e-3
                em = t * p[f'gwp_{g}']
                u = (np.sqrt(r.u_area ** 2 + u_mb ** 2 + (p['u_comb'] * k) ** 2 + (p[f'u_gef_{g}'] * k) ** 2
                             + (p['u_fire'] * k) ** 2) if r.area > 0 and frac > 0 else 0.0)
                o.update({f'gef_{g}': p[f'gef_{g}'], f't_{g}': t, f'em_{g}': em, f'u_{g}': u, f'pm_{g}': u * abs(em),
                          f'ux2_{g}': (u * em) ** 2, f'em_{g}_year': em / T})
            rec.append(o)
        return pd.DataFrame(rec)

    def _deg_rows(self):
        items = [dict(zone=zn, zi=zi, f_i=fi, f_f=ff, key=f'{zn}|{fi}|{ff}')
                 for zi, zn in enumerate(self.zones) for fi in self.forest_cats for ff in self.forest_cats]
        rows = self._estimate('key_deg', items)
        k, T = self.k_fe, self.T
        rec = []
        for _, r in rows.iterrows():
            a, b = self.forest_stocks[(r.f_i, r.zone)], self.forest_stocks[(r.f_f, r.zone)]
            si, ui, sf_, uf = a['ef'], a['u_tot'] * k, b['ef'], b['u_tot'] * k
            ef = si - sf_
            uef = np.sqrt((ui * si) ** 2 + (uf * sf_) ** 2) / abs(ef) if ef != 0 else 0.0
            em = r.area * ef
            u = np.sqrt(r.u_area ** 2 + uef ** 2) if r.area > 0 else 0.0
            rec.append(dict(s_i=si, u_s_i=ui, s_f=sf_, u_s_f=uf, ef=ef, u_ef=uef, em=em, u_em=u, pm=u * abs(em),
                            ux2=(u * em) ** 2, em_year=em / T))
        return pd.concat([rows, pd.DataFrame(rec)], axis=1)

    def _gain_rows(self):
        items = [dict(zone=zn, zi=zi, gain=g, key=f'{zn}|{g}') for zi, zn in enumerate(self.zones) for g in self.gain_cats]
        rows = self._estimate('key_gain', items)
        k, T = self.k_fe, self.T
        rec = []
        for _, r in rows.iterrows():
            rr = self.removal_rates[(r.gain, r.zone)]
            ug = rr['u_rate'] * k
            em = r.area * rr['g'] * self.coh
            u = np.sqrt(r.u_area ** 2 + ug ** 2) if r.area > 0 else 0.0
            rec.append(dict(rate=rr['rate'], u_rate=ug, r=rr['r'], g=rr['g'], em=em, u_em=u, pm=u * abs(em),
                            ux2=(u * em) ** 2, em_year=em / T))
        return pd.concat([rows, pd.DataFrame(rec)], axis=1)

    def _direct(self):
        """Estimateur direct des émissions nettes DEF : moyenne stratifiée du FE net de chaque échantillon."""
        ef = dict(zip(self.def_rows.key, self.def_rows.ef_net))
        s = self.samples
        s['y'] = s['key_def'].map(ef).fillna(0.0)
        s['y2'] = s['y'] ** 2
        g = s.groupby(['zone', 'strata'])
        sy = g['y'].sum().unstack(fill_value=0).reindex(index=self.zones, columns=self.strata, fill_value=0).to_numpy()
        sy2 = g['y2'].sum().unstack(fill_value=0).reindex(index=self.zones, columns=self.strata, fill_value=0).to_numpy()
        n = self.n_zh
        mean = sy / n
        var = (sy2 - n * mean ** 2) / (n - 1)
        self.direct = pd.DataFrame({'zone': self.zones, 'sum_y': list(sy), 'sum_y2': list(sy2),
                                    'em_net': self.A_z * (self.W_zh * mean).sum(axis=1),
                                    'se': self.A_z * np.sqrt((self.W_zh ** 2 * var / n).sum(axis=1))})

    def _bilan(self):
        """Moyennes annuelles par composante (national et par zone) et série par année."""
        T = self.T
        src = {'def': (self.def_rows, 'em_net', 'ux2_net'), 'soc': (self.soc_rows, 'em', 'ux2'),
               'ch4': (self.fire_rows, 'em_ch4', 'ux2_ch4'), 'n2o': (self.fire_rows, 'em_n2o', 'ux2_n2o'),
               'deg': (self.deg_rows, 'em', 'ux2'), 'rscf': (self.gain_rows, 'em', 'ux2')}

        def comp(df, col, ux):
            tot, ux2 = df[col].sum(), df[ux].sum()
            u = u_ratio(ux2, tot)
            return dict(year=tot / T, u=u, pm=u * abs(tot / T), period=tot)

        def combine(parts):
            val = sum(x['year'] for x in parts)
            pm = np.sqrt(sum(x['pm'] ** 2 for x in parts))
            return dict(year=val, u=pm / abs(val) if val else 0.0, pm=pm, period=val * T)

        nat = {c: comp(*src[c]) for c in src}
        nat['tot'] = combine([nat[c] for c in EMISSIONS])
        nat['net'] = combine([nat['tot'], nat['rscf']])
        self.bilan = nat
        zones = {}
        for zn in self.zones:
            d = {c: comp(df[df.zone == zn], col, ux) for c, (df, col, ux) in src.items()}
            d['tot'] = combine([d[c] for c in EMISSIONS])
            d['net'] = combine([d['tot'], d['rscf']])
            zones[zn] = d
        self.bilan_zone = zones
        years = []
        for t in range(1, T + 1):
            f = 2 * t / (T + 1)
            y = {'def': nat['def']['year'], 'soc': nat['soc']['year'] * f, 'ch4': nat['ch4']['year'],
                 'n2o': nat['n2o']['year'], 'deg': nat['deg']['year'], 'rscf': nat['rscf']['year'] * f}
            pm = {c: nat[c]['u'] * abs(y[c]) for c in y}
            y['tot'] = sum(y[c] for c in EMISSIONS)
            y['pm_tot'] = np.sqrt(sum(pm[c] ** 2 for c in EMISSIONS))
            y['net'] = y['tot'] + y['rscf']
            y['pm_net'] = np.sqrt(y['pm_tot'] ** 2 + pm['rscf'] ** 2)
            y['year'] = self.p['initial_year'] + t
            years.append(y)
        self.years = pd.DataFrame(years)


# ================================================================ constructeur du classeur
class _EmissionBuilder(WorkbookBase):
    def __init__(self, path, model, title, subtitle, params, sources, extra_cols, territory=None):
        self.path, self.m = path, model
        self.lab = territory_labels(territory)
        self.title, self.subtitle = title, subtitle
        self.p = params
        self.confidence = params['confidence']
        self.sources = sources
        self.extra_cols = list(extra_cols)
        self.Z, self.H = len(model.zones), len(model.strata)
        self._layout()

    # ------------------------------------------------------------ positions
    def _layout(self):
        Z, m = self.Z, self.m
        L = dict(conf=5, z=6, conf_fe=7, z_fe=8, k_fe=9, y0=10, y1=11, years=12, coh=13, seuil=14, cf=15, cf_low=16,
                 u_cf=17, co2c=18, socref=19, u_socref=20, soc_d=21, cf_litter=22, comb=23, u_comb=24, gef_ch4=25,
                 u_gef_ch4=26, gef_n2o=27, u_gef_n2o=28, gwp_ch4=29, gwp_n2o=30, u_fire=31)
        L['az_h2'] = 33
        L['az_hdr'] = 34
        L['az0'] = 35
        L['az_tot'] = L['az0'] + Z
        L['w_h2'] = L['az_tot'] + 3
        L['w_hdr'] = L['w_h2'] + 1
        L['w0'] = L['w_hdr'] + 1
        L['n_h2'] = L['w0'] + Z + 1
        L['n_hdr'] = L['n_h2'] + 1
        L['n0'] = L['n_hdr'] + 1
        L['n_tot'] = L['n0'] + Z
        self.L = L
        # feuilles d'activité : une ligne par transition à partir de la ligne 6 (index 5), puis tableau des zones
        self.act = {}
        for name, rows, labels, keycol, code in [
                (S_ADEF, m.def_rows, ['forest', 'post'], 'I', 'Def'),
                (S_ADEG, m.deg_rows, ['f_i', 'f_f'], 'M', 'Deg'),
                (S_AGAIN, m.gain_rows, ['gain'], 'O', 'Gain')]:
            k = len(labels)
            c = dict(zone=1, key=2 + k, s0=3 + k)
            c.update(n=3 + k + self.H, p=4 + k + self.H, area=5 + k + self.H, se=6 + k + self.H, u=7 + k + self.H,
                     pm=8 + k + self.H, ay=9 + k + self.H)
            first, last = 6, 6 + len(rows) - 1
            zd_h2 = last + 4
            self.act[name] = dict(rows=rows, labels=labels, keycol=keycol, code=code, cols=c, first=first, last=last,
                                  total=last + 1, zd_h2=zd_h2, zd_hdr=zd_h2 + 1, zd0=zd_h2 + 2, zd_nat=zd_h2 + 2 + Z)
        # Bilan
        self.bil = dict(first=6, rows={c: 6 + i for i, (c, _) in enumerate(COMPONENTS)})
        self.bil['area0'] = self.bil['first'] + len(COMPONENTS) + 3
        self.bil['year_h2'] = self.bil['area0'] + 3 + 3
        self.bil['year0'] = self.bil['year_h2'] + 2

    # ------------------------------------------------------------ utilitaires
    def hdr_row(self, ws, r, c0, headers, height=32, left=1):
        ws.set_row(r, height)
        for j, h in enumerate(headers):
            f = self.st.hdr_l if j < left else self.st.hdr
            if isinstance(h, tuple):
                self.formula(ws, r, c0 + j, h[0], f, h[1])
            else:
                self.label(ws, r, c0 + j, h, f)

    def inp(self, ws, r, c, v, nf=None):
        if isinstance(v, str):
            ws.write_string(r, c, v, self.st.cell(nf, inp=True))
        else:
            ws.write_number(r, c, float(v), self.st.cell(nf, inp=True))

    def note(self, ws, r, c1, c2, text, rows=1):
        self.label(ws, r, c1, text, self.st.note, r + rows - 1, c2)

    def key_fmt(self, band, top):
        kw = dict(font_color=GREY, font_size=7, border=1, border_color=GRID)
        if band:
            kw['bg_color'] = BAND
        if top:
            kw.update(top=2, top_color=GREEN)
        return self.st.fmt(**kw)

    def col(self, sheet, c, r1, r2):
        return rng(sheet, r1, c, r2, c)

    def zone_ref(self, zi):
        return ref(S_PAR, self.L['az0'] + zi, 1)

    # ------------------------------------------------------------ construction
    def build(self):
        self.wb = xlsxwriter.Workbook(self.path)
        self.st = _Styles(self.wb)
        self.st.label_font[self.st.text] = {}
        self.wb.set_properties({'title': self.title, 'subject': 'Émissions et absorptions forestières - SBAE',
                                'company': 'FAO'})
        self.wb.set_calc_mode('auto')
        self.wb.set_size(1600, 1000)
        L = self.L
        for name, row in [('Niveau_confiance', 'conf'), ('z_IC', 'z'), ('Niveau_confiance_FE', 'conf_fe'),
                          ('z_FE', 'z_fe'), ('Facteur_U_FE', 'k_fe'), ('Annee_initiale', 'y0'), ('Annee_finale', 'y1'),
                          ('Duree_periode', 'years'), ('Facteur_cohortes', 'coh'), ('Seuil_n', 'seuil'), ('CF', 'cf'),
                          ('CF_bas', 'cf_low'), ('U_CF', 'u_cf'), ('CO2_C', 'co2c'), ('SOC_ref', 'socref'),
                          ('U_SOC_ref', 'u_socref'), ('SOC_D', 'soc_d'), ('CF_litiere', 'cf_litter'),
                          ('Facteur_combustion', 'comb'), ('U_Facteur_combustion', 'u_comb'), ('Gef_CH4', 'gef_ch4'),
                          ('U_Gef_CH4', 'u_gef_ch4'), ('Gef_N2O', 'gef_n2o'), ('U_Gef_N2O', 'u_gef_n2o'),
                          ('PRG_CH4', 'gwp_ch4'), ('PRG_N2O', 'gwp_n2o'), ('U_Fraction_brulee', 'u_fire')]:
            self.wb.define_name(name, '=' + ref(S_PAR, L[row], 2))
        self.sheet_synthese()
        self.sheet_parametres()
        self.sheet_correspondances()
        self.sheet_stocks()
        for name, title in [(S_ADEF, 'Superficie déforestée par transition'),
                            (S_ADEG, 'Superficie dégradée par transition'),
                            (S_AGAIN, 'Superficie de gain par catégorie')]:
            self.sheet_activity(name, title)
        self.sheet_emissions_def()
        self.sheet_soc()
        self.sheet_fire()
        self.sheet_deg()
        self.sheet_rscf()
        self.sheet_bilan()
        self.sheet_resultats()
        self.sheet_controles()
        self.sheet_graphiques()
        self.sheet_equations()
        self.sheet_echantillons()
        self.wb.worksheets()[0].activate()
        self.wb.close()

    # ============================================================ Synthèse
    def sheet_synthese(self):
        st, m = self.st, self.m
        ws = self.sheet(S_SYN, 'Émissions et absorptions forestières — classeur dynamique', [46, 18, 14, 16, 14, 40])
        r = 4
        self.h2(ws, r, 1, 6, 'Paramètres')
        s = m.samples
        par = [
            ('Données de référence (CEO)', self.sources.get('data', ''), None, None),
            ('Période', '=Annee_initiale&" – "&Annee_finale&" ("&Duree_periode&" ans)"', None,
             f'{self.p["initial_year"]} – {self.p["final_year"]} ({m.T} ans)'),
            ('Échantillons valides', f"=COUNTA('{S_ECH}'!A:A)-1", INT, len(s)),
            ('Échantillons Def / Deg / Gain',
             f"=COUNTIF('{S_ECH}'!D:D,\"Def\")&\" / \"&COUNTIF('{S_ECH}'!D:D,\"Deg\")&\" / \"&COUNTIF('{S_ECH}'!D:D,\"Gain\")",
             None, f'{(s.activity == "Def").sum()} / {(s.activity == "Deg").sum()} / {(s.activity == "Gain").sum()}'),
            ('Niveau de confiance', '=Niveau_confiance', '0%', self.p['confidence']),
            ('Composantes', 'DEF biomasse et MOM, DEF sols minéraux, DEF combustion (CH₄, N₂O), dégradation, '
                            'renforcement des stocks', None, None),
            ('Méthode', 'GIEC 2006/2019 (approche 1) ; estimateur stratifié par zone (Cochran 1977)', None, None),
        ]
        for i, (k, v, nf, cached) in enumerate(par, start=r + 1):
            self.label(ws, i, 1, k, st.key)
            ws.merge_range(i, 2, i, 6, '', st.cell(nf, align='left'))
            if isinstance(v, str) and v.startswith('='):
                self.formula(ws, i, 2, v, st.cell(nf, align='left'), cached)
            else:
                ws.write(i, 2, v, st.cell(align='left'))
        r += len(par) + 2

        self.h2(ws, r, 1, 6, 'Résultats nationaux — moyennes annuelles')
        r += 1
        self.hdr_row(ws, r, 1, ['', 'Valeur', 'Incertitude (%)', '± (valeur)', 'Unité'])
        b, bil = self.bil, m.bilan
        lines = []
        for name, code in [('Superficie déforestée', 'Def'), ('Superficie dégradée', 'Deg'),
                           ('Superficie de gain', 'Gain')]:
            rr = b['area0'] + ['Def', 'Deg', 'Gain'].index(code)
            nat = m.zone_act[code].attrs['national']
            lines.append((name, ref(S_BIL, rr, 2), nat['area_year'], ref(S_BIL, rr, 3), nat['u_area'], 'ha/an'))
        for code, name in COMPONENTS:
            rr = b['rows'][code]
            lines.append((name, ref(S_BIL, rr, 2), bil[code]['year'], ref(S_BIL, rr, 3), bil[code]['u'], 'tCO₂/an'))
        for i, (name, fv, vv, fu, vu, unit) in enumerate(lines, start=r + 1):
            band = (i - r) % 2 == 0
            bold = name.startswith(('Total', 'Bilan'))
            ws.write(i, 1, name, st.cell(band=band, bold=True))
            self.formula(ws, i, 2, '=' + fv, st.cell(T0, band=band, bold=bold), vv)
            self.formula(ws, i, 3, '=' + fu, st.cell(PCT1, band=band), vu)
            self.formula(ws, i, 4, f'=ABS({a1(i, 2)})*{a1(i, 3)}', st.cell(T0, band=band), abs(vv) * vu)
            ws.write(i, 5, unit, st.cell(band=band))
        r += len(lines) + 1
        self.note(ws, r, 1, 6, 'Absorptions en valeurs négatives. Incertitudes : demi-largeur de l\'intervalle de confiance, en % '
                  'de la valeur et en valeur absolue (±). Composantes combinées selon l\'approche 1 du GIEC (voir Bilan).', 2)
        r += 3

        self.h2(ws, r, 1, 6, 'Mode d\'emploi')
        r += 1
        ws.write(r, 1, 'Cellule d\'entrée (modifiable)', st.cell(inp=True))
        ws.merge_range(r, 2, r, 6, 'Paramètres, superficies, correspondances, stocks, facteurs, données des échantillons.',
                       st.text)
        r += 1
        ws.write(r, 1, 'Cellule calculée (formule)', st.cell())
        ws.merge_range(r, 2, r, 6, 'Toutes les autres cellules sont des formules et se recalculent automatiquement.', st.text)
        r += 1
        self.note(ws, r, 1, 6,
                  'Scénarios : modifier les paramètres et les superficies A_hz (Paramètres), les correspondances des classes '
                  'CEO (Correspondances), les stocks, facteurs SOC et taux de croissance (Stocks de carbone) ou les classes '
                  'des échantillons (Échantillons). Les noms de zones et de catégories doivent rester identiques dans toutes '
                  'les feuilles. Pour revenir aux valeurs d\'origine, relancer le notebook.', 3)
        r += 4
        self.h2(ws, r, 1, 6, 'Contenu du classeur')
        for i, (name, desc) in enumerate(SHEETS, start=r + 1):
            ws.write_url(i, 1, f"internal:'{name}'!A1", st.link, name)
            ws.merge_range(i, 2, i, 6, '', st.text)
            self.label(ws, i, 2, desc, st.text)

    # ============================================================ Paramètres
    def sheet_parametres(self):
        st, L, m, H, Z = self.st, self.L, self.m, self.H, self.Z
        ws = self.sheet(S_PAR, 'Paramètres et superficies cartographiées', [44, 16, 16, 16, 16, 50])
        ws.write(3, 1, 'Cellules jaunes : valeurs modifiables. Les autres cellules sont calculées.', st.note1)
        self.h2(ws, 4, 1, 6, 'Paramètres généraux')
        p = self.p
        items = [
            ('conf', 'Niveau de confiance des résultats', p['confidence'], '0%', 'Intervalle bilatéral'),
            ('z', 'Quantile normal  z', '=ROUND(_xlfn.NORM.S.INV(1-(1-Niveau_confiance)/2),3)', '0.000', m.z),
            ('conf_fe', 'Niveau de confiance des incertitudes saisies', p['ef_confidence'], '0%',
             'Stocks, facteurs SOC, taux de croissance (IC 90 %)'),
            ('z_fe', 'Quantile normal  z_FE', '=ROUND(_xlfn.NORM.S.INV(1-(1-Niveau_confiance_FE)/2),3)', '0.000', m.z_fe),
            ('k_fe', 'Conversion des incertitudes saisies  z / z_FE', '=z_IC/z_FE', '0.0000', m.k_fe),
            ('y0', 'Année initiale', p['initial_year'], '0', 'État de référence des échantillons'),
            ('y1', 'Année finale', p['final_year'], '0', ''),
            ('years', 'Durée de la période  T (années)', '=Annee_finale-Annee_initiale', '0', m.T),
            ('coh', 'Facteur de cohortes  (T + 1) / 2', '=(Duree_periode+1)/2', '0.00', m.coh),
            ('seuil', 'Seuil « petit effectif »  n', p['low_n'], INT, 'Signale les transitions peu échantillonnées'),
            ('cf', 'Fraction de carbone de la biomasse  CF', m.cf, '0.000', 'GIEC 2006, tableau 4.3'),
            ('cf_low', 'Borne basse de CF', m.cf_low, '0.000', 'GIEC 2006, tableau 4.3 (0,44 – 0,49)'),
            ('u_cf', 'Incertitude de CF  U_CF', '=ABS(CF_bas-CF)/CF', PCT1, m.u_cf),
            ('co2c', 'Rapport CO₂ / C', '=44/12', '0.0000', 44 / 12),
            ('socref', 'Stock de référence des sols  SOC_ref (tC/ha)', p['socref'], '0.0',
             'Sols minéraux LAC, climat tropical humide (GIEC tableau 2.3)'),
            ('u_socref', 'Incertitude de SOC_ref', p['u_socref'], PCT1, 'Non renseignée : à compléter'),
            ('soc_d', 'Durée de transition des sols  D (années)', p['soc_d'], '0', 'GIEC : 20 ans par défaut'),
            ('cf_litter', 'Fraction de carbone de la litière', m.p['cf_litter'], '0.00',
             'GIEC 2006 : 0,37 ; litière tC/ha → tms/ha pour la biomasse combustible'),
            ('comb', 'Facteur de combustion  Cf', m.p['comb'], '0.00', 'GIEC 2006 tableau 2.6'),
            ('u_comb', 'Incertitude de Cf', m.p['u_comb'], PCT1, 'Non renseignée : à compléter'),
            ('gef_ch4', 'Facteur d\'émission  Gef CH₄ (g/kg ms brûlée)', m.p['gef_ch4'], '0.00',
             'GIEC 2006 tableau 2.5, forêt tropicale'),
            ('u_gef_ch4', 'Incertitude de Gef CH₄', m.p['u_gef_ch4'], PCT1, 'Non renseignée : à compléter'),
            ('gef_n2o', 'Facteur d\'émission  Gef N₂O (g/kg ms brûlée)', m.p['gef_n2o'], '0.00',
             'GIEC 2006 tableau 2.5, forêt tropicale'),
            ('u_gef_n2o', 'Incertitude de Gef N₂O', m.p['u_gef_n2o'], PCT1, 'Non renseignée : à compléter'),
            ('gwp_ch4', 'PRG du CH₄ (tCO₂e / t)', m.p['gwp_ch4'], '0', 'GIEC AR5'),
            ('gwp_n2o', 'PRG du N₂O (tCO₂e / t)', m.p['gwp_n2o'], '0', 'GIEC AR5'),
            ('u_fire', 'Incertitude de la fraction brûlée', m.p['u_fire'], PCT1,
             'Fractions par transition : Correspondances ; incertitude à compléter'),
        ]
        for key, label, v, nf, extra in items:
            r = L[key]
            self.label(ws, r, 1, label, st.key)
            if isinstance(v, str) and v.startswith('='):
                self.formula(ws, r, 2, v, st.cell(nf), extra)
            else:
                self.inp(ws, r, 2, v, nf)
                if extra:
                    ws.write(r, 3, extra, st.note1)
        strata_fmt = st.fmt(bold=True, font_color='white', bg_color=GREEN, align='center', border=1,
                            border_color='white', num_format='"Strate "0')

        def header(r, first, last):
            self.label(ws, r, 1, first, st.hdr_l)
            for hi, h in enumerate(m.strata):
                ws.write_number(r, 2 + hi, h, strata_fmt)
            self.label(ws, r, 2 + H, last, st.hdr2)

        self.h2(ws, L['az_h2'], 1, 2 + H, 'Superficie cartographiée par zone et par strate  A_hz (ha)')
        header(L['az_hdr'], 'Zone', 'A_z')
        for zi, zn in enumerate(m.zones):
            r = L['az0'] + zi
            ws.write_string(r, 1, zn, st.cell(bold=True, inp=True))
            for hi in range(H):
                self.inp(ws, r, 2 + hi, m.A_zh[zi, hi], HA)
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell(HA, bold=True), m.A_z[zi])
        r = L['az_tot']
        self.label(ws, r, 1, 'Total  Σ A_z', st.cell(bold=True, top=True))
        for c in range(2, 3 + H):
            v = m.A_zh[:, c - 2].sum() if c < 2 + H else m.A_z.sum()
            self.formula(ws, r, c, f'=SUM({a1(L["az0"], c)}:{a1(r - 1, c)})', st.cell(HA, bold=True, top=True), v)
        if self.sources.get('areas'):
            ws.write(r + 1, 1, f'Source : {self.sources["areas"]}', st.note1)

        self.h2(ws, L['w_h2'], 1, 2 + H, 'Poids  W_hz = A_hz / A_z')
        header(L['w_hdr'], 'Zone', 'Σ_h W_hz')
        for zi, zn in enumerate(m.zones):
            r, band = L['w0'] + zi, zi % 2 == 1
            self.formula(ws, r, 1, '=' + xl_rowcol_to_cell(L['az0'] + zi, 1, True, True), st.cell(band=band, bold=True), zn)
            for hi in range(H):
                self.formula(ws, r, 2 + hi,
                             f'={a1(L["az0"] + zi, 2 + hi)}/{xl_rowcol_to_cell(L["az0"] + zi, 2 + H, True, True)}',
                             st.cell('0.0000', band=band), m.W_zh[zi, hi])
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell('0.0000', band=band, bold=True), 1.0)

        self.h2(ws, L['n_h2'], 1, 2 + H, "Échantillons par zone et par strate  n_hz  (NB.SI.ENS sur « Échantillons »)")
        header(L['n_hdr'], 'Zone', 'n_z')
        for zi, zn in enumerate(m.zones):
            r, band = L['n0'] + zi, zi % 2 == 1
            z_ref = xl_rowcol_to_cell(L['az0'] + zi, 1, True, True)
            self.formula(ws, r, 1, '=' + z_ref, st.cell(band=band, bold=True), zn)
            for hi in range(H):
                h_ref = xl_rowcol_to_cell(L['az_hdr'], 2 + hi, True, True)
                self.formula(ws, r, 2 + hi, f"=COUNTIFS('{S_ECH}'!$C:$C,{z_ref},'{S_ECH}'!$B:$B,{h_ref})",
                             st.cell(INT, band=band), m.n_zh[zi, hi])
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell(INT, band=band, bold=True),
                         m.n_zh[zi].sum())
        r = L['n_tot']
        ws.write(r, 1, 'Total', st.cell(bold=True, top=True))
        for c in range(2, 3 + H):
            v = m.n_zh[:, c - 2].sum() if c < 2 + H else m.n_zh.sum()
            self.formula(ws, r, c, f'=SUM({a1(L["n0"], c)}:{a1(r - 1, c)})', st.cell(INT, bold=True, top=True), v)
        ws.freeze_panes(4, 2)

    # ============================================================ Correspondances
    def sheet_correspondances(self):
        st, m = self.st, self.m
        ws = self.sheet(S_COR, 'Correspondance des classes CEO (niveau 2) avec les catégories',
                        [30, 30, 3, 28, 26, 3, 30, 26, 3, 36])
        ws.write(3, 1, 'Les catégories sont modifiables (listes déroulantes). Une classe absente des tables reçoit '
                       '« NON ATTRIBUÉ » (voir Contrôles).', st.note1)
        r0, d0 = 5, 6
        tables = [
            (1, 'Classe CEO forestière (2021 ou 2025)', 'Catégorie forestière (stock)', m.forest_map, 'Corr_foret'),
            (4, 'Classe CEO post-déforestation (2025)', 'Utilisation (stock et sols)', m.post_map, 'Corr_post'),
            (7, 'Classe CEO forestière des gains (2025)', 'Catégorie de gain (croissance)', m.gain_map, 'Corr_gain'),
        ]
        lists = [('Catégories forestières', m.forest_cats), ('Utilisations post-déforestation', m.post_cats),
                 ('Catégories de gain', m.gain_cats)]
        # listes des catégories (sources des listes déroulantes)
        self.hdr_row(ws, r0, 10, ['Catégories disponibles'])
        r = d0
        bounds = []
        for title, items in lists:
            ws.write(r, 10, title, st.key)
            for i, c in enumerate(items, start=1):
                ws.write_string(r + i, 10, c, st.cell())
            bounds.append((r + 1, r + len(items)))
            r += len(items) + 2
        last = r
        for (c0, h1, h2, mp, name), (l0, l1) in zip(tables, bounds):
            self.hdr_row(ws, r0, c0, [h1, h2], left=2)
            for i, (k, v) in enumerate(mp.items()):
                ws.write_string(d0 + i, c0, k, st.cell(bold=True))
                self.inp(ws, d0 + i, c0 + 1, v)
            e = d0 + len(mp) - 1
            self.wb.define_name(name, '=' + rng(S_COR, d0, c0, e, c0 + 1))
            ws.data_validation(d0, c0 + 1, e, c0 + 1, {'validate': 'list', 'source': '=' + rng(S_COR, l0, 10, l1, 10)})
            last = max(last, e)
        # fraction brûlée par transition de déforestation (combustion de la biomasse)
        r = last + 3
        self.h2(ws, r, 1, 5, 'Fraction de la superficie déforestée brûlée, par transition (combustion de la biomasse)')
        self.hdr_row(ws, r + 1, 1, ['Catégorie forestière 2021', 'Utilisation post-déforestation', '',
                                    'Fraction brûlée', 'Clé'], left=2)
        f0 = r + 2
        pairs = [(f, u) for f in m.forest_cats for u in m.post_cats]
        for i, (f, u) in enumerate(pairs):
            rr = f0 + i
            ws.write_string(rr, 1, f, st.cell(bold=True))
            ws.write_string(rr, 2, u, st.cell())
            self.inp(ws, rr, 4, m.fire_fraction.get((f, u), FIRE_DEFAULT), '0.00')
            self.formula(ws, rr, 5, f'={a1(rr, 1)}&"|"&{a1(rr, 2)}', self.key_fmt(False, False), f'{f}|{u}')
        f1 = f0 + len(pairs) - 1
        self.wb.define_name('Feu_fraction', '=' + rng(S_COR, f0, 4, f1, 4))
        self.wb.define_name('Feu_cle', '=' + rng(S_COR, f0, 5, f1, 5))
        last = f1 + 1
        notes = list(self.sources.get('mapping_notes', [])) + [
            'Fraction brûlée : valeurs du NERF 2024 — 0,5 de la superficie déforestée, 0 pour les '
            'conversions en jachère / fourré ou en terres graminéennes et pour l\'agroforêt convertie en cacao ou en cultures '
            'pérennes. Une transition absente de la table reçoit 0,5.']
        for i, t in enumerate(notes):
            self.note(ws, last + 2 + i, 1, 10, t)

    # ============================================================ Stocks de carbone
    def sheet_stocks(self):
        st, m = self.st, self.m
        ws = self.sheet(S_STK, 'Stocks de carbone, facteurs SOC et taux de croissance', [28, 24] + [11] * 20)
        ws.write(3, 1, 'Biomasse en tonnes de matière sèche (tms/ha), carbone en tC/ha. Incertitudes = demi-largeur de '
                       'l\'IC au « Niveau de confiance des incertitudes saisies » (Paramètres).', st.note1)
        fr = list(m.forest_stocks.values())
        self.stk = {}
        # --- A. saisies forêt
        r = 5
        self.h2(ws, r, 1, 20, 'A. Stocks forestiers — saisies (IFN Tier 2 et GIEC Tier 1)')
        r += 1
        hdr = ['Catégorie', 'Zone', 'AGB (tms/ha)', 'U AGB', 'BGB (tms/ha)', 'U BGB', 'Litière (tC/ha)',
               'Litière borne sup. (tC/ha)', 'BM couché (tms/ha)', 'Souches (tms/ha)', 'BM sur pied (tms/ha)',
               'BGB BM sur pied (tms/ha)', 'U BM couché', 'U souches', 'U BM sur pied', 'U BGB BM sur pied',
               'Règle litière', 'Source']
        self.hdr_row(ws, r, 1, hdr, height=44)
        a0 = r + 1
        row_of = {(o['category'], o['zone']): a0 + i for i, o in enumerate(fr)}
        for i, o in enumerate(fr):
            rr = a0 + i
            ws.write_string(rr, 1, o['category'], st.cell(bold=True))
            ws.write_string(rr, 2, o['zone'], st.cell())
            self.inp(ws, rr, 3, o['agb'], DM2)
            self.inp(ws, rr, 4, o['u_agb'], PCT1)
            self.inp(ws, rr, 5, o['bgb'], DM2)
            self.inp(ws, rr, 6, o['u_bgb'], PCT1)
            if o['litter_rule'] == 'ipcc':
                self.inp(ws, rr, 7, o['litter_tc'], DM2)
                self.inp(ws, rr, 8, o['litter_upper_tc'], DM2)
            elif o['litter_rule'] == 'scaled':
                ref_r = row_of[(o['litter_ref'], o['zone'])]
                self.formula(ws, rr, 7, f'={a1(ref_r, 7)}*{a1(rr, 3)}/{a1(ref_r, 3)}', st.cell(DM2), o['litter'])
                ws.write_blank(rr, 8, None, st.cell())
            else:
                self.inp(ws, rr, 7, 0, DM2)
                ws.write_blank(rr, 8, None, st.cell())
            for j, c in enumerate(('lying', 'stump', 'standing', 'bgb_standing')):
                self.inp(ws, rr, 9 + j, o[f'dw_{c}'], DM2)
                self.inp(ws, rr, 13 + j, o[f'e_{c}'], PCT1)
            ws.write_string(rr, 17, {'ipcc': 'GIEC Tier 1', 'scaled': 'Proportionnelle à l\'AGB',
                                     'none': 'Aucune'}[o['litter_rule']], st.cell())
            ws.write_string(rr, 18, o.get('source', ''), st.cell())
        a1_ = a0 + len(fr) - 1
        # --- B. calculs forêt
        r = a1_ + 3
        self.h2(ws, r, 1, 20, 'B. Stocks forestiers — carbone par réservoir et stock total (FE brut de déforestation)')
        r += 1
        hdr = ['Catégorie', 'Zone', 'Clé', 'C AGB (tC/ha)', 'U', 'C BGB (tC/ha)', 'U', 'C litière (tC/ha)', 'U',
               'Bois mort (tms/ha)', 'U bois mort (mesure)', 'C bois mort (tC/ha)', 'U', 'C total (tC/ha)', 'U C total',
               'Stock (tCO₂/ha)', 'U stock', 'Biomasse combustible MB (tms/ha)', 'U MB']
        self.hdr_row(ws, r, 1, hdr, height=44)
        b0 = r + 1
        for i, o in enumerate(fr):
            rr, a = b0 + i, a0 + i
            band = i % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            ws.write_string(rr, 1, o['category'], cf(None, bold=True))
            ws.write_string(rr, 2, o['zone'], cf(None))
            self.formula(ws, rr, 3, f'={a1(rr, 1)}&"|"&{a1(rr, 2)}', self.key_fmt(band, False),
                         f'{o["category"]}|{o["zone"]}')
            self.formula(ws, rr, 4, f'={a1(a, 3)}*CF', cf(DM2), o['c_agb'])
            self.formula(ws, rr, 5, f'=SQRT({a1(a, 4)}^2+U_CF^2)', cf(PCT1), o['uc_agb'])
            self.formula(ws, rr, 6, f'={a1(a, 5)}*CF', cf(DM2), o['c_bgb'])
            self.formula(ws, rr, 7, f'=SQRT({a1(a, 6)}^2+U_CF^2)', cf(PCT1), o['uc_bgb'])
            self.formula(ws, rr, 8, f'={a1(a, 7)}', cf(DM2), o['c_litter'])
            if o['litter_rule'] == 'ipcc':
                fu = f'=({a1(a, 8)}-{a1(a, 7)})/{a1(a, 7)}'
            elif o['litter_rule'] == 'scaled':
                ra = row_of[(o['litter_ref'], o['zone'])]
                fu = f'=SQRT({a1(ra, 4)}^2+{a1(a, 4)}^2+(({a1(ra, 8)}-{a1(ra, 7)})/{a1(ra, 7)})^2)'
            else:
                fu = '=0'
            self.formula(ws, rr, 9, fu, cf(PCT1), o['u_litter'])
            self.formula(ws, rr, 10, f'=SUM({a1(a, 9)}:{a1(a, 12)})', cf(DM2), o['dw'])
            self.formula(ws, rr, 11, f'=IF({a1(rr, 10)}>0,SQRT(SUMPRODUCT(({a1(a, 9)}:{a1(a, 12)}*{a1(a, 13)}:{a1(a, 16)})^2))'
                                     f'/{a1(rr, 10)},0)', cf(PCT1), o['u_dw'])
            self.formula(ws, rr, 12, f'={a1(rr, 10)}*CF', cf(DM2), o['c_dw'])
            self.formula(ws, rr, 13, f'=SQRT({a1(rr, 11)}^2+U_CF^2)', cf(PCT1), o['uc_dw'])
            self.formula(ws, rr, 14, f'={a1(rr, 4)}+{a1(rr, 6)}+{a1(rr, 8)}+{a1(rr, 12)}', cf(DM2, bold=True), o['c_tot'])
            self.formula(ws, rr, 15, f'=SQRT(({a1(rr, 5)}*{a1(rr, 4)})^2+({a1(rr, 7)}*{a1(rr, 6)})^2+({a1(rr, 9)}*{a1(rr, 8)})^2'
                                     f'+({a1(rr, 13)}*{a1(rr, 12)})^2)/{a1(rr, 14)}', cf(PCT1), o['u_tot'])
            self.formula(ws, rr, 16, f'={a1(rr, 14)}*CO2_C', cf(EF2, bold=True), o['ef'])
            self.formula(ws, rr, 17, f'={a1(rr, 15)}', cf(PCT1), o['u_tot'])
            # MB = AGB + litière (tC/ha / CF litière) + bois mort, en tms/ha
            self.formula(ws, rr, 18, f'={a1(a, 3)}+{a1(rr, 8)}/CF_litiere+{a1(rr, 10)}', cf(DM2, bold=True), o['mb'])
            self.formula(ws, rr, 19, f'=IF({a1(rr, 18)}>0,SQRT(({a1(a, 4)}*{a1(a, 3)})^2+({a1(rr, 9)}*{a1(rr, 8)}/CF_litiere)^2'
                                     f'+({a1(rr, 11)}*{a1(rr, 10)})^2)/{a1(rr, 18)},0)', cf(PCT1), o['u_mb'])
        b1 = b0 + len(fr) - 1
        self.stk.update(f_key=rng(S_STK, b0, 3, b1, 3), f_ef=rng(S_STK, b0, 16, b1, 16), f_u=rng(S_STK, b0, 17, b1, 17),
                        f_mb=rng(S_STK, b0, 18, b1, 18), f_umb=rng(S_STK, b0, 19, b1, 19),
                        f_rows={(o['category'], o['zone']): b0 + i for i, o in enumerate(fr)})
        # --- C. stocks post-conversion
        pr = list(m.post_stocks.values())
        r = b1 + 3
        self.h2(ws, r, 1, 20, 'C. Stocks de l\'utilisation post-déforestation (biomasse aérienne et souterraine)')
        r += 1
        hdr = ['Catégorie', 'Règle', 'AGB (tms/ha)', 'U AGB', 'BGB (tms/ha)', 'U BGB', 'C AGB (tC/ha)', 'U', 'C BGB (tC/ha)',
               'U', 'C total (tC/ha)', 'U C total', 'Stock (tCO₂/ha)', 'U stock', 'Source']
        self.hdr_row(ws, r, 1, hdr, height=44)
        c0 = r + 1
        prow = {o['category']: c0 + i for i, o in enumerate(pr)}
        for i, o in enumerate(pr):
            rr, band = c0 + i, i % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            ws.write_string(rr, 1, o['category'], cf(None, bold=True))
            if o['rule'] == 'mean':
                srcs = [prow[c] for c in o['mean_of']]
                ws.write_string(rr, 2, ('Moyenne de ' if len(srcs) > 1 else 'Identique à ') + ' et '.join(o['mean_of']),
                                cf(None))
                for c_v, c_u, v, u in [(3, 4, 'agb', 'u_agb'), (5, 6, 'bgb', 'u_bgb')]:
                    self.formula(ws, rr, c_v, '=AVERAGE(' + ','.join(a1(x, c_v) for x in srcs) + ')', cf(DM2), o[v])
                    self.formula(ws, rr, c_u, '=IFERROR(SQRT(' + '+'.join(f'({a1(x, c_v)}*{a1(x, c_u)})^2' for x in srcs)
                                 + ')/(' + '+'.join(a1(x, c_v) for x in srcs) + '),0)', cf(PCT1), o[u])
            else:
                ws.write_string(rr, 2, 'Saisie' if o['rule'] == 'input' else 'Sans biomasse (0)', cf(None))
                self.inp(ws, rr, 3, o['agb'], DM2)
                self.inp(ws, rr, 4, o['u_agb'], PCT1)
                self.inp(ws, rr, 5, o['bgb'], DM2)
                self.inp(ws, rr, 6, o['u_bgb'], PCT1)
            self.formula(ws, rr, 7, f'={a1(rr, 3)}*CF', cf(DM2), o['c_agb'])
            self.formula(ws, rr, 8, f'=SQRT({a1(rr, 4)}^2+U_CF^2)', cf(PCT1), o['uc_agb'])
            self.formula(ws, rr, 9, f'={a1(rr, 5)}*CF', cf(DM2), o['c_bgb'])
            self.formula(ws, rr, 10, f'=SQRT({a1(rr, 6)}^2+U_CF^2)', cf(PCT1), o['uc_bgb'])
            self.formula(ws, rr, 11, f'={a1(rr, 7)}+{a1(rr, 9)}', cf(DM2, bold=True), o['c_tot'])
            self.formula(ws, rr, 12, f'=IF({a1(rr, 11)}>0,SQRT(({a1(rr, 8)}*{a1(rr, 7)})^2+({a1(rr, 10)}*{a1(rr, 9)})^2)'
                                     f'/{a1(rr, 11)},0)', cf(PCT1), o['u_tot'])
            self.formula(ws, rr, 13, f'={a1(rr, 11)}*CO2_C', cf(EF2, bold=True), o['ef'])
            self.formula(ws, rr, 14, f'={a1(rr, 12)}', cf(PCT1), o['u_tot'])
            ws.write_string(rr, 15, o.get('source', ''), cf(None))
        c1 = c0 + len(pr) - 1
        self.stk.update(p_key=rng(S_STK, c0, 1, c1, 1), p_ef=rng(S_STK, c0, 13, c1, 13), p_u=rng(S_STK, c0, 14, c1, 14),
                        p_rows=prow)
        # --- D. facteurs SOC
        sr = list(m.soc_factors.values())
        r = c1 + 3
        self.h2(ws, r, 1, 20, 'D. Facteurs de variation du carbone des sols (utilisation post-déforestation ; forêt = 1)')
        r += 1
        hdr = ['Utilisation', 'F_LU', 'U F_LU', 'F_MG', 'U F_MG', 'F_I', 'U F_I', 'F = F_LU·F_MG·F_I', 'U F', 'Source']
        self.hdr_row(ws, r, 1, hdr, height=36)
        d0 = r + 1
        for i, o in enumerate(sr):
            rr, band = d0 + i, i % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            ws.write_string(rr, 1, o['category'], cf(None, bold=True))
            for j, (v, u) in enumerate([('flu', 'u_flu'), ('fmg', 'u_fmg'), ('fi', 'u_fi')]):
                self.inp(ws, rr, 2 + 2 * j, o[v], '0.00')
                self.inp(ws, rr, 3 + 2 * j, o[u], PCT1)
            self.formula(ws, rr, 8, f'={a1(rr, 2)}*{a1(rr, 4)}*{a1(rr, 6)}', cf('0.0000', bold=True), o['f'])
            self.formula(ws, rr, 9, f'=IF({a1(rr, 8)}<>0,SQRT({a1(rr, 3)}^2+{a1(rr, 5)}^2+{a1(rr, 7)}^2),0)', cf(PCT1), o['u_f'])
            ws.write_string(rr, 10, o.get('source', ''), cf(None))
        d1 = d0 + len(sr) - 1
        self.stk.update(s_key=rng(S_STK, d0, 1, d1, 1), s_f=rng(S_STK, d0, 8, d1, 8), s_u=rng(S_STK, d0, 9, d1, 9))
        # --- E. taux de croissance
        gr = list(m.removal_rates.values())
        r = d1 + 3
        self.h2(ws, r, 1, 20, 'E. Taux de croissance des gains (non-forêt → forêt)')
        r += 1
        hdr = ['Catégorie de gain', 'Zone', 'Clé', 'Croissance AGB (tC/ha/an)', 'U', 'R (racines/tiges)',
               'G (tCO₂/ha/an)', 'U G', 'Source']
        self.hdr_row(ws, r, 1, hdr, height=36)
        e0 = r + 1
        for i, o in enumerate(gr):
            rr, band = e0 + i, i % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            ws.write_string(rr, 1, o['category'], cf(None, bold=True))
            ws.write_string(rr, 2, o['zone'], cf(None))
            self.formula(ws, rr, 3, f'={a1(rr, 1)}&"|"&{a1(rr, 2)}', self.key_fmt(band, False), f'{o["category"]}|{o["zone"]}')
            self.inp(ws, rr, 4, o['rate'], '0.000')
            self.inp(ws, rr, 5, o['u_rate'], PCT1)
            self.inp(ws, rr, 6, o['r'], '0.000')
            self.formula(ws, rr, 7, f'=-{a1(rr, 4)}*(1+{a1(rr, 6)})*CO2_C', cf(EF2, bold=True), o['g'])
            self.formula(ws, rr, 8, f'={a1(rr, 5)}', cf(PCT1), o['u_rate'])
            ws.write_string(rr, 9, o.get('source', ''), cf(None))
        e1 = e0 + len(gr) - 1
        self.stk.update(g_key=rng(S_STK, e0, 3, e1, 3), g_rate=rng(S_STK, e0, 4, e1, 4), g_r=rng(S_STK, e0, 6, e1, 6),
                        g_g=rng(S_STK, e0, 7, e1, 7), g_u=rng(S_STK, e0, 8, e1, 8))
        r = e1 + 2
        self.note(ws, r, 1, 19, 'A-B : C AGB = AGB·CF ; C BGB = BGB·CF ; litière en tC/ha (GIEC Tier 1, ou proportionnelle à '
                  'l\'AGB pour la forêt secondaire) ; C bois mort = Σ composantes·CF ; U de chaque réservoir combinée avec U_CF ; '
                  'U du total = √Σ(U·C)² / C. MB (biomasse combustible) = AGB + litière / CF litière + bois mort, en tms/ha, '
                  'sans les racines (GIEC 2006 éq. 2.27) ; U MB = √Σ(U·x)² / MB. D : F = 1 pour la forêt ; '
                  'U F = √(U_LU² + U_MG² + U_I²) (produit). E : G = − croissance AGB · (1 + R) · 44/12 ; R = 0 : croissance '
                  'aérienne seule.', 3)
        ws.freeze_panes(0, 3)

    # ============================================================ feuilles d'activité
    def sheet_activity(self, name, title):
        st, L, m, H = self.st, self.L, self.m, self.H
        A = self.act[name]
        rows, labels, c = A['rows'], A['labels'], A['cols']
        lab_hdr = {'forest': 'Forêt 2021', 'post': 'Utilisation post-déforestation', 'f_i': 'Forêt 2021',
                   'f_f': 'Forêt 2025', 'gain': 'Catégorie de gain'}
        widths = [24] + [26] * len(labels) + [8] + [8] * H + [8, 11, 14, 13, 11, 13, 14]
        ws = self.sheet(name, title + ' (estimateur stratifié par zone)', widths)
        r0 = A['first'] - 1
        strata_hdr = st.fmt(bold=True, font_color='white', bg_color=GREEN, align='center', border=1, border_color='white',
                            num_format='"n  h="0')
        tail = ['n', 'Proportion', 'Superficie période (ha)', 'Erreur standard (ha)', self.ci_label('Incertitude', ' (%)'),
                '± (ha)', 'Superficie annuelle (ha/an)']

        def header(r, with_labels):
            ws.set_row(r, 40)
            self.label(ws, r, 1, 'Zone', st.hdr_l)
            for j, lb in enumerate(labels):
                ws.write(r, 2 + j, lab_hdr[lb] if with_labels else '', st.hdr_l)
            ws.write(r, c['key'], 'Clé' if with_labels else '', st.hdr)
            for hi, h in enumerate(m.strata):
                ws.write_number(r, c['s0'] + hi, h, strata_hdr)
            for j, h in enumerate(tail):
                if isinstance(h, tuple):
                    self.formula(ws, r, c['n'] + j, h[0], st.hdr, h[1])
                else:
                    ws.write(r, c['n'] + j, h, st.hdr)

        def estimate_cells(r, zi, N, v, cf):
            Wz = rng(S_PAR, L['w0'] + zi, 2, L['w0'] + zi, 1 + H)
            NH = rng(S_PAR, L['n0'] + zi, 2, L['n0'] + zi, 1 + H)
            Az = ref(S_PAR, L['az0'] + zi, 2 + H)
            self.formula(ws, r, c['n'], f'=SUM({N})', cf(T0, bold=True), v['n'])
            self.formula(ws, r, c['p'], f'=SUMPRODUCT({Wz},{N}/{NH})', cf('0.0000%;-0.0000%;"–"'), v['p'])
            self.formula(ws, r, c['area'], f'={Az}*{a1(r, c["p"])}', cf(HA0, bold=True), v['area'])
            self.formula(ws, r, c['se'], f'={Az}*SQRT(SUMPRODUCT({Wz}^2*({N}/{NH})*(1-{N}/{NH})/({NH}-1)))', cf(HA0), v['se'])
            self.formula(ws, r, c['u'], f'=IF({a1(r, c["area"])}>0,z_IC*{a1(r, c["se"])}/{a1(r, c["area"])},0)', cf(PCT1),
                         v['u_area'])
            self.formula(ws, r, c['pm'], f'={a1(r, c["u"])}*{a1(r, c["area"])}', cf(HA0), v['u_area'] * v['area'])
            self.formula(ws, r, c['ay'], f'={a1(r, c["area"])}/Duree_periode', cf(HA0), v['area_year'])

        header(r0, True)
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r, zi = A['first'] + i, row.zi
            band, top = zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            self.formula(ws, r, 1, '=' + self.zone_ref(zi), cf(None, bold=top), row.zone)
            ws.write_row(r, 2, [row[lb] for lb in labels], cf(None))
            key_f = '=' + '&"|"&'.join(a1(r, cc) for cc in [1] + list(range(2, 2 + len(labels))))
            self.formula(ws, r, c['key'], key_f, self.key_fmt(band, top), row.key)
            for hi in range(H):
                h_ref = xl_rowcol_to_cell(r0, c['s0'] + hi, True, False)
                self.formula(ws, r, c['s0'] + hi, f"=COUNTIFS('{S_ECH}'!${A['keycol']}:${A['keycol']},"
                                                  f"{xl_rowcol_to_cell(r, c['key'], False, True)},'{S_ECH}'!$B:$B,{h_ref})",
                             cf(T0), row.n_h[hi])
            estimate_cells(r, zi, f'{a1(r, c["s0"])}:{a1(r, c["s0"] + H - 1)}', row, cf)
        last, rt = A['last'], A['total']
        ws.write(rt, 1, 'Total (toutes transitions)', st.cell(bold=True, top=True))
        for cc in range(2, c['ay'] + 1):
            ws.write_blank(rt, cc, None, st.cell(top=True))
        for cc, v, nf in [(c['n'], rows.n.sum(), T0), (c['area'], rows.area.sum(), HA0), (c['ay'], rows.area_year.sum(), HA0)]:
            self.formula(ws, rt, cc, f'=SUM({a1(A["first"], cc)}:{a1(last, cc)})', st.cell(nf, bold=True, top=True), v)
        n_col = xl_col_to_name(c['n'])
        ws.conditional_format(A['first'], 1, last, c['ay'], {
            'type': 'formula', 'criteria': f'=${n_col}{A["first"] + 1}=0', 'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.conditional_format(A['first'], c['n'], last, c['n'], {
            'type': 'formula', 'criteria': f'=AND(${n_col}{A["first"] + 1}>0,${n_col}{A["first"] + 1}<Seuil_n)',
            'format': self.wb.add_format({'bg_color': '#FDE7C8'})})
        ws.autofilter(r0, 1, last, c['ay'])
        ws.freeze_panes(A['first'], c['key'] + 1)

        # superficie totale par zone : estimation directe de l'activité
        code = A['code']
        self.h2(ws, A['zd_h2'], 1, c['ay'], f'Superficie totale par zone — estimation directe de l\'activité « {code} »')
        header(A['zd_hdr'], False)
        zd = m.zone_act[code]
        for zi, zr in zd.iterrows():
            r, band = A['zd0'] + zi, zi % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            self.formula(ws, r, 1, '=' + self.zone_ref(zi), cf(None, bold=True), zr.zone)
            for cc in range(2, c['s0']):
                ws.write_blank(r, cc, None, cf(None))
            for hi in range(H):
                h_ref = xl_rowcol_to_cell(A['zd_hdr'], c['s0'] + hi, True, False)
                self.formula(ws, r, c['s0'] + hi, f"=COUNTIFS('{S_ECH}'!$C:$C,{self.zone_ref(zi)},'{S_ECH}'!$D:$D,"
                                                  f"\"{code}\",'{S_ECH}'!$B:$B,{h_ref})", cf(T0), zr.n_h[hi])
            estimate_cells(r, zi, f'{a1(r, c["s0"])}:{a1(r, c["s0"] + H - 1)}', zr, cf)
        r = A['zd_nat']
        z0, z1 = A['zd0'], A['zd0'] + self.Z - 1
        nat = zd.attrs['national']
        ws.write(r, 1, self.lab['Adj'] + ' (somme des zones)', st.cell(bold=True, top=True))
        for cc in range(2, c['ay'] + 1):
            ws.write_blank(r, cc, None, st.cell(top=True))
        tf = lambda nf: st.cell(nf, bold=True, top=True)
        for hi in range(H):
            cc = c['s0'] + hi
            self.formula(ws, r, cc, f'=SUM({a1(z0, cc)}:{a1(z1, cc)})', tf(T0), nat['n_h'][hi])
        self.formula(ws, r, c['n'], f'=SUM({a1(z0, c["n"])}:{a1(z1, c["n"])})', tf(T0), nat['n'])
        self.formula(ws, r, c['area'], f'=SUM({a1(z0, c["area"])}:{a1(z1, c["area"])})', tf(HA0), nat['area'])
        self.formula(ws, r, c['se'], f'=SQRT(SUMSQ({a1(z0, c["se"])}:{a1(z1, c["se"])}))', tf(HA0), nat['se'])
        self.formula(ws, r, c['u'], f'=IF({a1(r, c["area"])}>0,z_IC*{a1(r, c["se"])}/{a1(r, c["area"])},0)', tf(PCT1),
                     nat['u_area'])
        self.formula(ws, r, c['pm'], f'={a1(r, c["u"])}*{a1(r, c["area"])}', tf(HA0), nat['u_area'] * nat['area'])
        self.formula(ws, r, c['ay'], f'={a1(r, c["area"])}/Duree_periode', tf(HA0), nat['area_year'])
        self.note(ws, r + 2, 1, c['ay'], 'Chaque transition est un domaine de la zone : p̂ = Σ_h W_hz · n_hkz / n_hz (Cochran 1977). '
                  'Lignes grisées : aucun échantillon ; en orange : moins de « Seuil n » échantillons. La somme des transitions '
                  'd\'une zone est égale à l\'estimation directe si tous les échantillons sont attribués (voir Contrôles).', 2)

    def act_cell(self, name, i, col):
        A = self.act[name]
        return ref(name, A['first'] + i, A['cols'][col])

    # ============================================================ émissions : utilitaires communs
    def emission_sheet(self, name, title, widths, groups, headers):
        ws = self.sheet(name, title, widths)
        r0 = 4
        for c1, c2, t in groups:
            if c1 == c2:
                ws.write(r0, c1, t, self.st.hdr)
            else:
                ws.merge_range(r0, c1, r0, c2, t, self.st.hdr)
        ws.set_row(r0 + 1, 30)
        for j, h in enumerate(headers):
            ws.write(r0 + 1, 1 + j, h, self.st.hdr2 if j >= 3 else self.st.hdr_l)
        return ws

    def emission_total(self, ws, first, last, ncols, specs, label=None):
        """specs : liste de (col, kind, valeur) ; kind = sum | u:(col_val, col_ux2) | pm:(col_val, col_u)."""
        label = label or 'Total ' + self.lab['adj']
        rt = last + 1
        ws.write(rt, 1, label, self.st.cell(bold=True, top=True))
        for c in range(2, ncols + 1):
            ws.write_blank(rt, c, None, self.st.cell(top=True))
        for c, kind, nf, v in specs:
            if kind == 'sum':
                f = f'=SUM({a1(first, c)}:{a1(last, c)})'
            elif kind[0] == 'u':
                f = f'=IF({a1(rt, kind[1])}<>0,SQRT({a1(rt, kind[2])})/ABS({a1(rt, kind[1])}),0)'
            else:
                f = f'=ABS({a1(rt, kind[1])})*{a1(rt, kind[2])}'
            self.formula(ws, rt, c, f, self.st.cell(nf, bold=True, top=True), v)
        return rt

    # ============================================================ Émissions DEF
    def sheet_emissions_def(self):
        st, m = self.st, self.m
        rows = m.def_rows
        groups = [(1, 4, 'Transition'), (5, 6, 'Données d\'activité'), (7, 8, 'FE brut (stock forestier)'),
                  (9, 10, 'Stock post-déforestation'), (11, 12, 'FE net'), (13, 16, 'Émissions brutes (période)'),
                  (17, 20, 'Émissions nettes (période)'), (21, 22, 'Moyenne annuelle')]
        headers = ['Zone', 'Forêt 2021', 'Utilisation post-déforestation', 'Clé', 'Superficie (ha)', 'U (%)',
                   'tCO₂/ha', 'U (%)', 'tCO₂/ha', 'U (%)', 'tCO₂/ha', 'U (%)', 'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²',
                   'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²', 'Brutes (tCO₂/an)', 'Nettes (tCO₂/an)']
        ws = self.emission_sheet(S_EM, 'Déforestation — émissions de la biomasse et de la MOM',
                                 [22, 24, 24, 8, 13, 9, 10, 9, 10, 9, 10, 9, 14, 9, 13, 10, 14, 9, 13, 10, 14, 14],
                                 groups, headers)
        first = 6
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r = first + i
            band, top = row.zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            R = lambda c: a1(r, c)
            self.formula(ws, r, 1, '=' + self.act_cell(S_ADEF, i, 'zone'), cf(None, bold=top), row.zone)
            self.formula(ws, r, 2, '=' + ref(S_ADEF, first + i, 2), cf(None), row.forest)
            self.formula(ws, r, 3, '=' + ref(S_ADEF, first + i, 3), cf(None), row.post)
            self.formula(ws, r, 4, f'={R(1)}&"|"&{R(2)}&"|"&{R(3)}', self.key_fmt(band, top), row.key)
            self.formula(ws, r, 5, '=' + self.act_cell(S_ADEF, i, 'area'), cf(HA0), row.area)
            self.formula(ws, r, 6, '=' + self.act_cell(S_ADEF, i, 'u'), cf(PCT1), row.u_area)
            fkey = f'{R(2)}&"|"&{R(1)}'
            self.formula(ws, r, 7, f'=INDEX({self.stk["f_ef"]},MATCH({fkey},{self.stk["f_key"]},0))', cf(EF2), row.ef_gross)
            self.formula(ws, r, 8, f'=INDEX({self.stk["f_u"]},MATCH({fkey},{self.stk["f_key"]},0))*Facteur_U_FE', cf(PCT1),
                         row.u_ef_gross)
            self.formula(ws, r, 9, f'=INDEX({self.stk["p_ef"]},MATCH({R(3)},{self.stk["p_key"]},0))', cf(EF2), row.stock_post)
            self.formula(ws, r, 10, f'=INDEX({self.stk["p_u"]},MATCH({R(3)},{self.stk["p_key"]},0))*Facteur_U_FE', cf(PCT1),
                         row.u_stock_post)
            self.formula(ws, r, 11, f'={R(7)}-{R(9)}', cf(EF2, bold=True), row.ef_net)
            self.formula(ws, r, 12, f'=IF({R(11)}<>0,SQRT(({R(8)}*{R(7)})^2+({R(10)}*{R(9)})^2)/ABS({R(11)}),0)', cf(PCT1),
                         row.u_ef_net)
            for c0, (v, u, pm, ux2), efc, efu in [(13, ('em_gross', 'u_em_gross', 'pm_gross', 'ux2_gross'), 7, 8),
                                                  (17, ('em_net', 'u_em_net', 'pm_net', 'ux2_net'), 11, 12)]:
                self.formula(ws, r, c0, f'={R(5)}*{R(efc)}', cf(T0, bold=c0 == 17), row[v])
                self.formula(ws, r, c0 + 1, f'=IF({R(5)}>0,SQRT({R(6)}^2+{R(efu)}^2),0)', cf(PCT1), row[u])
                self.formula(ws, r, c0 + 2, f'=ABS({R(c0)})*{R(c0 + 1)}', cf(T0), abs(row[pm]))
                self.formula(ws, r, c0 + 3, f'=({R(c0 + 1)}*{R(c0)})^2', cf(SCI), row[ux2])
            self.formula(ws, r, 21, f'={R(13)}/Duree_periode', cf(T0), row.em_gross_year)
            self.formula(ws, r, 22, f'={R(17)}/Duree_periode', cf(T0, bold=True), row.em_net_year)
        last = first + len(rows) - 1
        tg, tn = rows.em_gross.sum(), rows.em_net.sum()
        ug, un = u_ratio(rows.ux2_gross.sum(), tg), u_ratio(rows.ux2_net.sum(), tn)
        rt = self.emission_total(ws, first, last, 22, [
            (5, 'sum', HA0, rows.area.sum()), (13, 'sum', T0, tg), (16, 'sum', SCI, rows.ux2_gross.sum()),
            (14, ('u', 13, 16), PCT1, ug), (15, ('pm', 13, 14), T0, abs(tg) * ug),
            (17, 'sum', T0, tn), (20, 'sum', SCI, rows.ux2_net.sum()), (18, ('u', 17, 20), PCT1, un),
            (19, ('pm', 17, 18), T0, abs(tn) * un), (21, 'sum', T0, rows.em_gross_year.sum()),
            (22, 'sum', T0, rows.em_net_year.sum())])
        self.em = {S_EM: dict(first=first, last=last, total=rt, val=17, ux2=20, year=22, gross=13, ux2_gross=16,
                              year_gross=21, area=5, u_area=6)}
        ws.conditional_format(first, 1, last, 22, {'type': 'formula', 'criteria': f'=$E{first + 1}=0',
                                                   'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.autofilter(5, 1, last, 22)
        ws.freeze_panes(first, 4)
        self.wb.define_name('Cle_transitions', '=' + rng(S_EM, first, 4, last, 4))
        self.wb.define_name('FE_net_transitions', '=' + rng(S_EM, first, 11, last, 11))
        self.note(ws, rt + 2, 1, 22,
                  'FE brut = stock forestier (catégorie, zone) ; FE net = FE brut − stock post-déforestation. Incertitudes des '
                  'stocks converties au niveau de confiance des résultats (× Facteur_U_FE). Émissions = superficie × FE ; '
                  'U = √(U_AD² + U_FE²) ; total : U = √Σ(U·x)² / |Σx| (GIEC 2006 vol. 1 éq. 3.1 et 3.2). Émissions immédiates : '
                  'moyenne annuelle = total / T.', 2)

    # ============================================================ Émissions DEF SOC
    def sheet_soc(self):
        st, m = self.st, self.m
        rows = m.soc_rows
        groups = [(1, 3, 'Transition'), (4, 5, 'Données d\'activité'), (6, 6, 'Sol'), (7, 8, 'Facteur F'),
                  (9, 11, 'Variation du stock des sols'), (12, 15, 'Émissions (période)'), (16, 16, 'Annuel')]
        headers = ['Zone', 'Forêt 2021', 'Utilisation post-déforestation', 'Superficie (ha)', 'U (%)',
                   'SOC_ref (tC/ha)', 'F', 'U (%)', 'ΔSOC (tC/ha)', 'U (%)', 'Taux (tCO₂/ha/an)', 'tCO₂', 'U (%)',
                   '± (tCO₂)', '(U·x)²', 'tCO₂/an']
        ws = self.emission_sheet(S_SOC, 'Déforestation — émissions du carbone organique des sols minéraux',
                                 [22, 24, 24, 13, 9, 11, 9, 9, 11, 9, 12, 14, 9, 13, 10, 14], groups, headers)
        first = 6
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r = first + i
            band, top = row.zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            R = lambda c: a1(r, c)
            em_r = self.em[S_EM]['first'] + i
            self.formula(ws, r, 1, '=' + ref(S_EM, em_r, 1), cf(None, bold=top), row.zone)
            self.formula(ws, r, 2, '=' + ref(S_EM, em_r, 2), cf(None), row.forest)
            self.formula(ws, r, 3, '=' + ref(S_EM, em_r, 3), cf(None), row.post)
            self.formula(ws, r, 4, '=' + ref(S_EM, em_r, 5), cf(HA0), row.area)
            self.formula(ws, r, 5, '=' + ref(S_EM, em_r, 6), cf(PCT1), row.u_area)
            self.formula(ws, r, 6, '=SOC_ref', cf('0.0'), row.socref)
            self.formula(ws, r, 7, f'=INDEX({self.stk["s_f"]},MATCH({R(3)},{self.stk["s_key"]},0))', cf('0.0000'), row.f)
            self.formula(ws, r, 8, f'=INDEX({self.stk["s_u"]},MATCH({R(3)},{self.stk["s_key"]},0))*Facteur_U_FE', cf(PCT1),
                         row.u_f)
            self.formula(ws, r, 9, f'={R(6)}*({R(7)}-1)', cf(EF2), row.dsoc)
            self.formula(ws, r, 10, f'=IF({R(7)}<>1,SQRT((U_SOC_ref*Facteur_U_FE)^2+({R(8)}*{R(7)}/ABS({R(7)}-1))^2),0)',
                         cf(PCT1), row.u_dsoc)
            self.formula(ws, r, 11, f'=-{R(9)}/SOC_D*CO2_C', cf(EF2), row.rate)
            self.formula(ws, r, 12, f'={R(4)}*{R(11)}*Facteur_cohortes', cf(T0, bold=True), row.em)
            self.formula(ws, r, 13, f'=IF({R(4)}>0,SQRT({R(5)}^2+{R(10)}^2),0)', cf(PCT1), row.u_em)
            self.formula(ws, r, 14, f'=ABS({R(12)})*{R(13)}', cf(T0), row.pm)
            self.formula(ws, r, 15, f'=({R(13)}*{R(12)})^2', cf(SCI), row.ux2)
            self.formula(ws, r, 16, f'={R(12)}/Duree_periode', cf(T0, bold=True), row.em_year)
        last = first + len(rows) - 1
        t, u = rows.em.sum(), u_ratio(rows.ux2.sum(), rows.em.sum())
        rt = self.emission_total(ws, first, last, 16, [
            (4, 'sum', HA0, rows.area.sum()), (12, 'sum', T0, t), (15, 'sum', SCI, rows.ux2.sum()),
            (13, ('u', 12, 15), PCT1, u), (14, ('pm', 12, 13), T0, abs(t) * u), (16, 'sum', T0, rows.em_year.sum())])
        self.em[S_SOC] = dict(first=first, last=last, total=rt, val=12, ux2=15, year=16)
        ws.conditional_format(first, 1, last, 16, {'type': 'formula', 'criteria': f'=$D{first + 1}=0',
                                                   'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.autofilter(5, 1, last, 16)
        ws.freeze_panes(first, 4)
        self.note(ws, rt + 2, 1, 16,
                  'GIEC 2006 vol. 4 éq. 2.25 : ΔSOC = SOC_ref · (F − 1), avec F = F_LU·F_MG·F_I de l\'utilisation finale et '
                  'F = 1 pour la forêt ; la variation s\'étale sur D ans. Chaque hectare déforesté émet −ΔSOC/D · 44/12 par an. '
                  'Avec une superficie annuelle constante (A/T), les cohortes s\'accumulent : l\'année t émet (A/T)·t·taux, soit '
                  'A · taux · (T+1)/2 sur la période. U = √(U_AD² + U_ΔSOC²) ; U_ΔSOC combine U_SOC_ref et l\'incertitude de '
                  '(F − 1).', 3)

    # ============================================================ Émissions DEF CH4 N2O (combustion de la biomasse)
    def sheet_fire(self):
        st, m = self.st, self.m
        rows = m.fire_rows
        groups = [(1, 3, 'Transition'), (4, 5, 'Données d\'activité'), (6, 7, 'Utilisation du feu'),
                  (8, 10, 'Biomasse brûlée'), (11, 17, 'CH₄'), (18, 24, 'N₂O')]
        gas_hdr = ['Gef (g/kg)', 't gaz', 'tCO₂e', 'U (%)', '± (tCO₂e)', '(U·x)²', 'tCO₂e/an']
        headers = ['Zone', 'Forêt 2021', 'Utilisation post-déforestation', 'Superficie (ha)', 'U (%)', 'Fraction brûlée',
                   'Superficie brûlée (ha)', 'MB (tms/ha)', 'U (%)', 'Cf'] + gas_hdr + gas_hdr
        ws = self.emission_sheet(S_FIRE, 'Déforestation — CH₄ et N₂O de la combustion de la biomasse',
                                 [22, 24, 24, 13, 9, 9, 13, 10, 9, 7] + [9, 11, 13, 9, 12, 10, 12] * 2, groups, headers)
        first = 6
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r = first + i
            band, top = row.zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            R = lambda c: a1(r, c)
            em_r = self.em[S_EM]['first'] + i
            self.formula(ws, r, 1, '=' + ref(S_EM, em_r, 1), cf(None, bold=top), row.zone)
            self.formula(ws, r, 2, '=' + ref(S_EM, em_r, 2), cf(None), row.forest)
            self.formula(ws, r, 3, '=' + ref(S_EM, em_r, 3), cf(None), row.post)
            self.formula(ws, r, 4, '=' + ref(S_EM, em_r, 5), cf(HA0), row.area)
            self.formula(ws, r, 5, '=' + ref(S_EM, em_r, 6), cf(PCT1), row.u_area)
            self.formula(ws, r, 6, f'=IFERROR(INDEX(Feu_fraction,MATCH({R(2)}&"|"&{R(3)},Feu_cle,0)),{FIRE_DEFAULT})',
                         cf('0.00'), row.frac)
            self.formula(ws, r, 7, f'={R(4)}*{R(6)}', cf(HA0), row.burnt)
            fkey = f'{R(2)}&"|"&{R(1)}'
            self.formula(ws, r, 8, f'=INDEX({self.stk["f_mb"]},MATCH({fkey},{self.stk["f_key"]},0))', cf(DM2), row.mb)
            self.formula(ws, r, 9, f'=INDEX({self.stk["f_umb"]},MATCH({fkey},{self.stk["f_key"]},0))*Facteur_U_FE',
                         cf(PCT1), row.u_mb)
            self.formula(ws, r, 10, '=Facteur_combustion', cf('0.00'), row.comb)
            for c0, (g, _), gef, ugef, gwp in [(11, GASES[0], 'Gef_CH4', 'U_Gef_CH4', 'PRG_CH4'),
                                               (18, GASES[1], 'Gef_N2O', 'U_Gef_N2O', 'PRG_N2O')]:
                self.formula(ws, r, c0, f'={gef}', cf('0.00'), row[f'gef_{g}'])
                self.formula(ws, r, c0 + 1, f'={R(7)}*{R(8)}*{R(10)}*{R(c0)}*0.001', cf(EF2), row[f't_{g}'])
                self.formula(ws, r, c0 + 2, f'={R(c0 + 1)}*{gwp}', cf(T0, bold=True), row[f'em_{g}'])
                self.formula(ws, r, c0 + 3, f'=IF(AND({R(4)}>0,{R(6)}>0),SQRT({R(5)}^2+{R(9)}^2'
                                            f'+(U_Facteur_combustion*Facteur_U_FE)^2+({ugef}*Facteur_U_FE)^2'
                                            f'+(U_Fraction_brulee*Facteur_U_FE)^2),0)', cf(PCT1), row[f'u_{g}'])
                self.formula(ws, r, c0 + 4, f'=ABS({R(c0 + 2)})*{R(c0 + 3)}', cf(T0), row[f'pm_{g}'])
                self.formula(ws, r, c0 + 5, f'=({R(c0 + 3)}*{R(c0 + 2)})^2', cf(SCI), row[f'ux2_{g}'])
                self.formula(ws, r, c0 + 6, f'={R(c0 + 2)}/Duree_periode', cf(T0, bold=True), row[f'em_{g}_year'])
        last = first + len(rows) - 1
        specs = [(4, 'sum', HA0, rows.area.sum()), (7, 'sum', HA0, rows.burnt.sum())]
        for c0, (g, _) in [(11, GASES[0]), (18, GASES[1])]:
            t, u = rows[f'em_{g}'].sum(), u_ratio(rows[f'ux2_{g}'].sum(), rows[f'em_{g}'].sum())
            specs += [(c0 + 1, 'sum', EF2, rows[f't_{g}'].sum()), (c0 + 2, 'sum', T0, t),
                      (c0 + 5, 'sum', SCI, rows[f'ux2_{g}'].sum()), (c0 + 3, ('u', c0 + 2, c0 + 5), PCT1, u),
                      (c0 + 4, ('pm', c0 + 2, c0 + 3), T0, abs(t) * u), (c0 + 6, 'sum', T0, rows[f'em_{g}_year'].sum())]
        rt = self.emission_total(ws, first, last, 24, specs)
        for c0, (g, _) in [(11, GASES[0]), (18, GASES[1])]:
            self.em[g] = dict(sheet=S_FIRE, first=first, last=last, total=rt, val=c0 + 2, ux2=c0 + 5, year=c0 + 6)
        ws.conditional_format(first, 1, last, 24, {'type': 'formula', 'criteria': f'=$E{first + 1}=0',
                                                   'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.autofilter(5, 1, last, 24)
        ws.freeze_panes(first, 4)
        self.note(ws, rt + 2, 1, 24,
                  'GIEC 2006 vol. 4 éq. 2.27 : L = A · fraction brûlée · MB · '
                  'Cf · Gef · 10⁻³ (t de gaz), puis × PRG (AR5). MB = AGB + litière + bois mort (tms/ha, sans les racines) pour '
                  'toutes les transitions ; fractions brûlées : Correspondances. U = √(U_AD² + U_MB² + U_Cf² + U_Gef² + '
                  'U_fraction²) ; les incertitudes de Cf, Gef et de la fraction brûlée sont à compléter (Paramètres). '
                  'Émissions immédiates : moyenne annuelle = total / T.', 3)

    # ============================================================ Émissions DEG
    def sheet_deg(self):
        st, m = self.st, self.m
        rows = m.deg_rows
        groups = [(1, 3, 'Transition'), (4, 5, 'Données d\'activité'), (6, 7, 'Stock forêt 2021'),
                  (8, 9, 'Stock forêt 2025'), (10, 11, 'FE dégradation'), (12, 15, 'Émissions (période)'),
                  (16, 16, 'Annuel')]
        headers = ['Zone', 'Forêt 2021', 'Forêt 2025', 'Superficie (ha)', 'U (%)', 'tCO₂/ha', 'U (%)', 'tCO₂/ha', 'U (%)',
                   'tCO₂/ha', 'U (%)', 'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²', 'tCO₂/an']
        ws = self.emission_sheet(S_DEG, 'Dégradation — émissions',
                                 [22, 26, 26, 13, 9, 11, 9, 11, 9, 11, 9, 14, 9, 13, 10, 14], groups, headers)
        first = 6
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r = first + i
            band, top = row.zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            R = lambda c: a1(r, c)
            self.formula(ws, r, 1, '=' + self.act_cell(S_ADEG, i, 'zone'), cf(None, bold=top), row.zone)
            self.formula(ws, r, 2, '=' + ref(S_ADEG, self.act[S_ADEG]['first'] + i, 2), cf(None), row.f_i)
            self.formula(ws, r, 3, '=' + ref(S_ADEG, self.act[S_ADEG]['first'] + i, 3), cf(None), row.f_f)
            self.formula(ws, r, 4, '=' + self.act_cell(S_ADEG, i, 'area'), cf(HA0), row.area)
            self.formula(ws, r, 5, '=' + self.act_cell(S_ADEG, i, 'u'), cf(PCT1), row.u_area)
            for c0, cat, v, u in [(6, R(2), 's_i', 'u_s_i'), (8, R(3), 's_f', 'u_s_f')]:
                k = f'{cat}&"|"&{R(1)}'
                self.formula(ws, r, c0, f'=INDEX({self.stk["f_ef"]},MATCH({k},{self.stk["f_key"]},0))', cf(EF2), row[v])
                self.formula(ws, r, c0 + 1, f'=INDEX({self.stk["f_u"]},MATCH({k},{self.stk["f_key"]},0))*Facteur_U_FE',
                             cf(PCT1), row[u])
            self.formula(ws, r, 10, f'={R(6)}-{R(8)}', cf(EF2, bold=True), row.ef)
            self.formula(ws, r, 11, f'=IF({R(10)}<>0,SQRT(({R(7)}*{R(6)})^2+({R(9)}*{R(8)})^2)/ABS({R(10)}),0)', cf(PCT1),
                         row.u_ef)
            self.formula(ws, r, 12, f'={R(4)}*{R(10)}', cf(T0, bold=True), row.em)
            self.formula(ws, r, 13, f'=IF({R(4)}>0,SQRT({R(5)}^2+{R(11)}^2),0)', cf(PCT1), row.u_em)
            self.formula(ws, r, 14, f'=ABS({R(12)})*{R(13)}', cf(T0), row.pm)
            self.formula(ws, r, 15, f'=({R(13)}*{R(12)})^2', cf(SCI), row.ux2)
            self.formula(ws, r, 16, f'={R(12)}/Duree_periode', cf(T0, bold=True), row.em_year)
        last = first + len(rows) - 1
        t, u = rows.em.sum(), u_ratio(rows.ux2.sum(), rows.em.sum())
        rt = self.emission_total(ws, first, last, 16, [
            (4, 'sum', HA0, rows.area.sum()), (12, 'sum', T0, t), (15, 'sum', SCI, rows.ux2.sum()),
            (13, ('u', 12, 15), PCT1, u), (14, ('pm', 12, 13), T0, abs(t) * u), (16, 'sum', T0, rows.em_year.sum())])
        self.em[S_DEG] = dict(first=first, last=last, total=rt, val=12, ux2=15, year=16)
        ws.conditional_format(first, 1, last, 16, {'type': 'formula', 'criteria': f'=$D{first + 1}=0',
                                                   'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.autofilter(5, 1, last, 16)
        ws.freeze_panes(first, 4)
        self.note(ws, rt + 2, 1, 16,
                  'FE dégradation = stock (forêt 2021, zone) − stock (forêt 2025, zone), réservoirs biomasse, litière et bois '
                  'mort ; U = √[(U_i·S_i)² + (U_f·S_f)²] / |S_i − S_f|. Émissions immédiates : '
                  'moyenne annuelle = total / T. Les catégories identiques en 2021 et 2025 ont un FE nul.', 2)

    # ============================================================ Absorptions RSCF
    def sheet_rscf(self):
        st, m = self.st, self.m
        rows = m.gain_rows
        groups = [(1, 2, 'Gain'), (3, 4, 'Données d\'activité'), (5, 9, 'Facteur d\'absorption'),
                  (10, 13, 'Absorptions (période)'), (14, 14, 'Annuel')]
        headers = ['Zone', 'Catégorie de gain', 'Superficie (ha)', 'U (%)', 'Croissance AGB (tC/ha/an)', 'U (%)', 'R',
                   'G (tCO₂/ha/an)', 'U (%)', 'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²', 'tCO₂/an']
        ws = self.emission_sheet(S_RSCF, 'Renforcement des stocks de carbone forestier — absorptions',
                                 [22, 26, 13, 9, 13, 9, 8, 12, 9, 14, 9, 13, 10, 14], groups, headers)
        first = 6
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r = first + i
            band, top = row.zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            R = lambda c: a1(r, c)
            self.formula(ws, r, 1, '=' + self.act_cell(S_AGAIN, i, 'zone'), cf(None, bold=top), row.zone)
            self.formula(ws, r, 2, '=' + ref(S_AGAIN, self.act[S_AGAIN]['first'] + i, 2), cf(None), row.gain)
            self.formula(ws, r, 3, '=' + self.act_cell(S_AGAIN, i, 'area'), cf(HA0), row.area)
            self.formula(ws, r, 4, '=' + self.act_cell(S_AGAIN, i, 'u'), cf(PCT1), row.u_area)
            k = f'{R(2)}&"|"&{R(1)}'
            self.formula(ws, r, 5, f'=INDEX({self.stk["g_rate"]},MATCH({k},{self.stk["g_key"]},0))', cf('0.000'), row.rate)
            self.formula(ws, r, 6, f'=INDEX({self.stk["g_u"]},MATCH({k},{self.stk["g_key"]},0))*Facteur_U_FE', cf(PCT1),
                         row.u_rate)
            self.formula(ws, r, 7, f'=INDEX({self.stk["g_r"]},MATCH({k},{self.stk["g_key"]},0))', cf('0.000'), row.r)
            self.formula(ws, r, 8, f'=-{R(5)}*(1+{R(7)})*CO2_C', cf(EF2, bold=True), row.g)
            self.formula(ws, r, 9, f'={R(6)}', cf(PCT1), row.u_rate)
            self.formula(ws, r, 10, f'={R(3)}*{R(8)}*Facteur_cohortes', cf(T0, bold=True), row.em)
            self.formula(ws, r, 11, f'=IF({R(3)}>0,SQRT({R(4)}^2+{R(9)}^2),0)', cf(PCT1), row.u_em)
            self.formula(ws, r, 12, f'=ABS({R(10)})*{R(11)}', cf(T0), row.pm)
            self.formula(ws, r, 13, f'=({R(11)}*{R(10)})^2', cf(SCI), row.ux2)
            self.formula(ws, r, 14, f'={R(10)}/Duree_periode', cf(T0, bold=True), row.em_year)
        last = first + len(rows) - 1
        t, u = rows.em.sum(), u_ratio(rows.ux2.sum(), rows.em.sum())
        rt = self.emission_total(ws, first, last, 14, [
            (3, 'sum', HA0, rows.area.sum()), (10, 'sum', T0, t), (13, 'sum', SCI, rows.ux2.sum()),
            (11, ('u', 10, 13), PCT1, u), (12, ('pm', 10, 11), T0, abs(t) * u), (14, 'sum', T0, rows.em_year.sum())])
        self.em[S_RSCF] = dict(first=first, last=last, total=rt, val=10, ux2=13, year=14)
        ws.conditional_format(first, 1, last, 14, {'type': 'formula', 'criteria': f'=$C{first + 1}=0',
                                                   'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.autofilter(5, 1, last, 14)
        ws.freeze_panes(first, 3)
        self.note(ws, rt + 2, 1, 14,
                  'G = − croissance AGB · (1 + R) · 44/12 (valeur négative = absorption). Avec une superficie annuelle de gain '
                  'constante (A/T), les cohortes s\'accumulent : l\'année t absorbe (A/T)·t·G, soit A · G · (T+1)/2 sur la '
                  'période. U = √(U_AD² + U_G²).', 2)

    # ============================================================ Bilan
    def sheet_bilan(self):
        st, m, b = self.st, self.m, self.bil
        ws = self.sheet(S_BIL, 'Bilan des émissions et absorptions', [48, 18, 13, 16, 16, 16, 16, 13, 16, 16, 16, 13, 16])
        bil = m.bilan
        r = b['first'] - 2
        self.h2(ws, r, 1, 7, 'A. Moyennes annuelles ' + self.lab['fem_pl'] + ' par composante')
        self.hdr_row(ws, r + 1, 1, ['Composante', 'tCO₂e/an', 'Incertitude (%)', '± (tCO₂e/an)', 'Part des émissions',
                                    'Total période (tCO₂e)'])
        src = {'def': S_EM, 'soc': S_SOC, 'ch4': 'ch4', 'n2o': 'n2o', 'deg': S_DEG, 'rscf': S_RSCF}   # clés de self.em
        R = b['rows']
        for code, name in COMPONENTS:
            rr = R[code]
            band = rr % 2 == 1
            bold = code in ('tot', 'net')
            cf = lambda nf, **k: st.cell(nf, band=band, bold=bold, top=bold, **k)
            ws.write(rr, 1, name, cf(None))
            v = bil[code]
            if code in src:
                e = self.em[src[code]]
                sh = e.get('sheet', src[code])
                self.formula(ws, rr, 2, '=' + ref(sh, e['total'], e['year']), cf(T0), v['year'])
                self.formula(ws, rr, 3, f'=IF({ref(sh, e["total"], e["val"])}<>0,SQRT({ref(sh, e["total"], e["ux2"])})'
                                        f'/ABS({ref(sh, e["total"], e["val"])}),0)', cf(PCT1), v['u'])
            else:
                parts = EMISSIONS if code == 'tot' else ['tot', 'rscf']
                self.formula(ws, rr, 2, '=' + '+'.join(a1(R[p], 2) for p in parts), cf(T0), v['year'])
                self.formula(ws, rr, 3, f'=IF({a1(rr, 2)}<>0,SQRT(' + '+'.join(f'{a1(R[p], 4)}^2' for p in parts)
                             + f')/ABS({a1(rr, 2)}),0)', cf(PCT1), v['u'])
            self.formula(ws, rr, 4, f'=ABS({a1(rr, 2)})*{a1(rr, 3)}', cf(T0), v['pm'])
            if code in EMISSIONS + ['tot']:
                self.formula(ws, rr, 5, f'=IF({a1(R["tot"], 2)}<>0,{a1(rr, 2)}/{a1(R["tot"], 2)},0)', cf(PCT1),
                             v['year'] / bil['tot']['year'] if bil['tot']['year'] else 0)
            else:
                ws.write_blank(rr, 5, None, cf(None))
            self.formula(ws, rr, 6, f'={a1(rr, 2)}*Duree_periode', cf(T0), v['year'] * m.T)

        # superficies
        r = b['area0'] - 2
        self.h2(ws, r, 1, 7, 'B. Superficies annuelles ' + self.lab['fem_pl'] + ' (estimation directe)')
        self.hdr_row(ws, r + 1, 1, ['Activité', 'ha/an', 'Incertitude (%)', '± (ha/an)', '', 'Total période (ha)'])
        for i, (code, name, sheet) in enumerate([('Def', 'Déforestation', S_ADEF), ('Deg', 'Dégradation', S_ADEG),
                                                 ('Gain', 'Gain (non-forêt → forêt)', S_AGAIN)]):
            rr, band = b['area0'] + i, i % 2 == 1
            A = self.act[sheet]
            nat = m.zone_act[code].attrs['national']
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            ws.write(rr, 1, name, cf(None, bold=True))
            self.formula(ws, rr, 2, '=' + ref(sheet, A['zd_nat'], A['cols']['ay']), cf(HA0), nat['area_year'])
            self.formula(ws, rr, 3, '=' + ref(sheet, A['zd_nat'], A['cols']['u']), cf(PCT1), nat['u_area'])
            self.formula(ws, rr, 4, f'={a1(rr, 2)}*{a1(rr, 3)}', cf(HA0), nat['u_area'] * nat['area_year'])
            ws.write_blank(rr, 5, None, cf(None))
            self.formula(ws, rr, 6, '=' + ref(sheet, A['zd_nat'], A['cols']['area']), cf(HA0), nat['area'])

        # série par année
        r = b['year_h2']
        self.h2(ws, r, 1, 13, 'C. Série par année (tCO₂e/an)')
        hdr = ['Année', 'DEF biomasse et MOM', 'DEF sols minéraux', 'DEF combustion CH₄', 'DEF combustion N₂O',
               'Dégradation', 'Total émissions', 'U (%)', '± (tCO₂e)', 'RSCF (absorptions)', 'Bilan net', 'U (%)',
               '± (tCO₂e)']
        self.hdr_row(ws, r + 1, 1, hdr, height=36)
        ycol = {'def': 2, 'soc': 3, 'ch4': 4, 'n2o': 5, 'deg': 6, 'tot': 7, 'rscf': 10, 'net': 11}
        self.bil['ycol'] = ycol
        cohort = ('soc', 'rscf')        # les cohortes s'accumulent ; les autres composantes sont constantes
        for i, y in m.years.iterrows():
            rr, band = b['year0'] + i, i % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            self.formula(ws, rr, 1, f'=Annee_initiale+{i + 1}', cf('0', bold=True), int(y.year))
            t = f'({a1(rr, 1)}-Annee_initiale)'
            for k in EMISSIONS + ['rscf']:
                f = '=' + xl_rowcol_to_cell(R[k], 2, True, True) + (f'*2*{t}/(Duree_periode+1)' if k in cohort else '')
                self.formula(ws, rr, ycol[k], f, cf(T0), y[k])
            ct = ycol['tot']
            self.formula(ws, rr, ct, f'=SUM({a1(rr, ycol["def"])}:{a1(rr, ycol["deg"])})', cf(T0, bold=True), y['tot'])
            pm = '+'.join(f'(ABS({a1(rr, ycol[k])})*{xl_rowcol_to_cell(R[k], 3, True, True)})^2' for k in EMISSIONS)
            self.formula(ws, rr, ct + 2, f'=SQRT({pm})', cf(T0), y['pm_tot'])
            self.formula(ws, rr, ct + 1, f'=IF({a1(rr, ct)}<>0,{a1(rr, ct + 2)}/ABS({a1(rr, ct)}),0)', cf(PCT1),
                         y['pm_tot'] / abs(y['tot']) if y['tot'] else 0)
            cr, cn = ycol['rscf'], ycol['net']
            self.formula(ws, rr, cn, f'={a1(rr, ct)}+{a1(rr, cr)}', cf(T0, bold=True), y['net'])
            self.formula(ws, rr, cn + 2, f'=SQRT({a1(rr, ct + 2)}^2+(ABS({a1(rr, cr)})*'
                                         f'{xl_rowcol_to_cell(R["rscf"], 3, True, True)})^2)', cf(T0), y['pm_net'])
            self.formula(ws, rr, cn + 1, f'=IF({a1(rr, cn)}<>0,{a1(rr, cn + 2)}/ABS({a1(rr, cn)}),0)', cf(PCT1),
                         y['pm_net'] / abs(y['net']) if y['net'] else 0)
        rr = b['year0'] + m.T
        ws.write(rr, 1, 'Moyenne', st.cell(bold=True, top=True))
        means = {c: k for k, c in ycol.items()}
        for c in range(2, 14):
            if c not in means:
                ws.write_blank(rr, c, None, st.cell(top=True))
            else:
                v = m.years[means[c]].mean()
                self.formula(ws, rr, c, f'=AVERAGE({a1(b["year0"], c)}:{a1(rr - 1, c)})', st.cell(T0, bold=True, top=True), v)
        self.note(ws, rr + 2, 1, 13,
                  'Valeurs positives : émissions ; négatives : absorptions. Composantes supposées indépendantes : '
                  'U = √Σ(U·x)² / |Σx| (GIEC 2006 vol. 1 éq. 3.2). Déforestation (biomasse et combustion CH₄, N₂O) et '
                  'dégradation : émissions immédiates, constantes chaque année. Sols et RSCF : '
                  'les cohortes s\'accumulent, l\'année t vaut moyenne · 2t / (T+1).', 3)

    # ============================================================ Résultats
    def sheet_resultats(self):
        st, m, Z = self.st, self.m, self.Z
        ws = self.sheet(S_RES, 'Résultats par zone et par catégorie',
                        [28] + [14, 10, 13] * 7)
        r = 4
        # A. superficies par zone
        self.h2(ws, r, 1, 10, 'A. Superficies annuelles par zone (ha/an, estimation directe)')
        r += 1
        ws.merge_range(r, 1, r + 1, 1, 'Zone', st.hdr_l)
        for j, (name, sheet) in enumerate([('Déforestation', S_ADEF), ('Dégradation', S_ADEG), ('Gain', S_AGAIN)]):
            c0 = 2 + 3 * j
            ws.merge_range(r, c0, r, c0 + 2, name, st.hdr)
            for k, h in enumerate(['ha/an', 'U (%)', '± (ha/an)']):
                ws.write(r + 1, c0 + k, h, st.hdr2)
        first = r + 2
        for zi, zn in enumerate(m.zones + [self.lab['Adj']]):
            rr, band, nat = first + zi, zi % 2 == 1, zi == Z
            cf = lambda nf, **k: st.cell(nf, **{'band': band and not nat, 'bold': nat, 'top': nat, **k})
            if nat:
                ws.write(rr, 1, self.lab['Adj'], cf(None))
            else:
                self.formula(ws, rr, 1, '=' + self.zone_ref(zi), cf(None, bold=True), zn)
            for j, (code, sheet) in enumerate([('Def', S_ADEF), ('Deg', S_ADEG), ('Gain', S_AGAIN)]):
                c0, A = 2 + 3 * j, self.act[sheet]
                zrow = A['zd_nat'] if nat else A['zd0'] + zi
                v = m.zone_act[code].attrs['national'] if nat else m.zone_act[code].iloc[zi]
                self.formula(ws, rr, c0, '=' + ref(sheet, zrow, A['cols']['ay']), cf(HA0), v['area_year'])
                self.formula(ws, rr, c0 + 1, '=' + ref(sheet, zrow, A['cols']['u']), cf(PCT1), v['u_area'])
                self.formula(ws, rr, c0 + 2, f'={a1(rr, c0)}*{a1(rr, c0 + 1)}', cf(HA0), v['u_area'] * v['area_year'])
        self.res_area_rows = (first, first + Z)
        r = first + Z + 3

        # B. émissions et absorptions par zone
        comps = [('def', 'DEF biomasse et MOM', S_EM), ('soc', 'DEF sols minéraux', S_SOC),
                 ('ch4', 'DEF combustion CH₄', 'ch4'), ('n2o', 'DEF combustion N₂O', 'n2o'), ('deg', 'Dégradation', S_DEG),
                 ('rscf', 'RSCF (absorptions)', S_RSCF), ('net', 'Bilan net', None)]
        self.h2(ws, r, 1, 1 + 3 * len(comps), 'B. Émissions et absorptions annuelles par zone (tCO₂e/an)')
        r += 1
        ws.merge_range(r, 1, r + 1, 1, 'Zone', st.hdr_l)
        for j, (_, name, _) in enumerate(comps):
            c0 = 2 + 3 * j
            ws.merge_range(r, c0, r, c0 + 2, name, st.hdr)
            for k, h in enumerate(['tCO₂e/an', 'U (%)', '± (tCO₂e/an)']):
                ws.write(r + 1, c0 + k, h, st.hdr2)
        self.res_net_col = 2 + 3 * (len(comps) - 1)
        first = r + 2
        for zi, zn in enumerate(m.zones + [self.lab['Adj']]):
            rr, band, nat = first + zi, zi % 2 == 1, zi == Z
            cf = lambda nf, **k: st.cell(nf, **{'band': band and not nat, 'bold': nat, 'top': nat, **k})
            if nat:
                ws.write(rr, 1, self.lab['Adj'], cf(None))
            else:
                self.formula(ws, rr, 1, '=' + self.zone_ref(zi), cf(None, bold=True), zn)
            vals = m.bilan if nat else m.bilan_zone[zn]
            for j, (code, _, sheet) in enumerate(comps):
                c0 = 2 + 3 * j
                v = vals[code]
                if sheet is None:
                    parts = list(range(2, c0, 3))
                    self.formula(ws, rr, c0, '=' + '+'.join(a1(rr, p) for p in parts), cf(T0, bold=True), v['year'])
                    self.formula(ws, rr, c0 + 2, '=SQRT(' + '+'.join(f'{a1(rr, p + 2)}^2' for p in parts) + ')', cf(T0),
                                 v['pm'])
                    self.formula(ws, rr, c0 + 1, f'=IF({a1(rr, c0)}<>0,{a1(rr, c0 + 2)}/ABS({a1(rr, c0)}),0)', cf(PCT1), v['u'])
                    continue
                e = self.em[sheet]
                sh = e.get('sheet', sheet)
                zc = self.col(sh, 1, e['first'], e['last'])
                vc, uc = self.col(sh, e['val'], e['first'], e['last']), self.col(sh, e['ux2'], e['first'], e['last'])
                if nat:
                    self.formula(ws, rr, c0, f'=SUM({vc})/Duree_periode', cf(T0), v['year'])
                    self.formula(ws, rr, c0 + 1, f'=IFERROR(SQRT(SUM({uc}))/ABS(SUM({vc})),0)', cf(PCT1), v['u'])
                else:
                    key = a1(rr, 1)
                    self.formula(ws, rr, c0, f'=SUMIFS({vc},{zc},{key})/Duree_periode', cf(T0), v['year'])
                    self.formula(ws, rr, c0 + 1, f'=IFERROR(SQRT(SUMIFS({uc},{zc},{key}))/ABS(SUMIFS({vc},{zc},{key})),0)',
                                 cf(PCT1), v['u'])
                self.formula(ws, rr, c0 + 2, f'=ABS({a1(rr, c0)})*{a1(rr, c0 + 1)}', cf(T0), v['pm'])
        self.res_em_rows = (first, first + Z)
        r = first + Z + 3

        # C et D. déforestation (biomasse) par catégorie forestière et par utilisation
        e = self.em[S_EM]
        rows = m.def_rows
        col = lambda c: self.col(S_EM, c, e['first'], e['last'])
        hdr = ['', 'Superficie (ha/an)', 'U (%)', '± (ha/an)', 'Émissions brutes (tCO₂/an)', 'U (%)', '± (tCO₂/an)',
               'Émissions nettes (tCO₂/an)', 'U (%)', '± (tCO₂/an)', 'Part des émissions nettes']
        self.res_cat_rows = {}
        for title, by, field, groups in [
                (f'C. Déforestation (biomasse et MOM) par catégorie forestière en {self.p["initial_year"]}', 2, 'forest',
                 m.forest_cats),
                ('D. Déforestation (biomasse et MOM) par utilisation post-déforestation', 3, 'post', m.post_cats)]:
            self.h2(ws, r, 1, 11, title)
            h = list(hdr)
            h[0] = 'Catégorie'
            self.hdr_row(ws, r + 1, 1, h, height=40)
            first = r + 2
            for i, g in enumerate(groups):
                rr, band = first + i, i % 2 == 1
                cf = lambda nf, **k: st.cell(nf, band=band, **k)
                key = a1(rr, 1)
                sub = rows[rows[field] == g]
                ws.write_string(rr, 1, g, cf(None, bold=True))
                va = sub.area_year.sum()
                ua = np.sqrt(((sub.u_area * sub.area) ** 2).sum()) / sub.area.sum() if sub.area.sum() else 0
                self.formula(ws, rr, 2, f'=SUMIFS({col(5)},{col(by)},{key})/Duree_periode', cf(HA0), va)
                self.formula(ws, rr, 3, f'=IFERROR(SQRT(SUMPRODUCT(({col(by)}={key})*({col(6)}*{col(5)})^2))'
                                        f'/SUMIFS({col(5)},{col(by)},{key}),0)', cf(PCT1), ua)
                self.formula(ws, rr, 4, f'={a1(rr, 2)}*{a1(rr, 3)}', cf(HA0), va * ua)
                for c_val, c_em, c_ux2, s_em, s_ux2 in [(5, 13, 16, 'em_gross', 'ux2_gross'), (8, 17, 20, 'em_net', 'ux2_net')]:
                    tot = sub[s_em].sum()
                    uu = u_ratio(sub[s_ux2].sum(), tot)
                    self.formula(ws, rr, c_val, f'=SUMIFS({col(c_em)},{col(by)},{key})/Duree_periode', cf(T0), tot / m.T)
                    self.formula(ws, rr, c_val + 1, f'=IFERROR(SQRT(SUMIFS({col(c_ux2)},{col(by)},{key}))'
                                                    f'/ABS(SUMIFS({col(c_em)},{col(by)},{key})),0)', cf(PCT1), uu)
                    self.formula(ws, rr, c_val + 2, f'=ABS({a1(rr, c_val)})*{a1(rr, c_val + 1)}', cf(T0), abs(tot / m.T) * uu)
                self.formula(ws, rr, 11, f'=IFERROR({a1(rr, 8)}/{ref(S_EM, e["total"], 22)},0)', cf(PCT1),
                             sub.em_net.sum() / rows.em_net.sum() if rows.em_net.sum() else 0)
            self.res_cat_rows[field] = (first, first + len(groups))
            r = first + len(groups) + 2
        self.note(ws, r, 1, 16,
                  'Tableau A : estimation directe de chaque activité par zone ; ' + self.lab['adj'] + ' = somme des zones, ES = √Σ ES_z². '
                  'Tableaux B à D : sommes des lignes des feuilles d\'émissions, U = √Σ(U·x)² / |Σx| (GIEC approche 1). '
                  'Le bilan net combine les six composantes en quadrature.', 2)

    # ============================================================ Contrôles
    def sheet_controles(self):
        st, L, m, H, Z = self.st, self.L, self.m, self.H, self.Z
        ws = self.sheet(S_CTL, 'Contrôles de cohérence', [46, 15, 15, 15, 13, 15, 15, 13, 15, 15, 13, 14, 15, 14, 14])
        ok_fmt = self.wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#166534', 'bold': True})
        ko_fmt = self.wb.add_format({'bg_color': '#F8CBAD', 'font_color': '#B91C1C', 'bold': True})

        def status(r1, r2, c):
            ws.conditional_format(r1, c, r2, c, {'type': 'cell', 'criteria': '==', 'value': '"OK"', 'format': ok_fmt})
            ws.conditional_format(r1, c, r2, c, {'type': 'cell', 'criteria': '==', 'value': '"À VÉRIFIER"', 'format': ko_fmt})

        s = m.samples
        r = 4
        self.h2(ws, r, 1, 5, '1. Échantillons et correspondances')
        self.hdr_row(ws, r + 1, 1, ['Contrôle', 'Valeur', 'Attendu', 'Statut'])
        rr = r + 2
        self.label(ws, rr, 1, 'Échantillons valides', st.cell(bold=True))
        self.formula(ws, rr, 2, f"=COUNTA('{S_ECH}'!A:A)-1", st.cell(INT), len(s))
        ws.write_blank(rr, 3, None, st.cell())
        ws.write_blank(rr, 4, None, st.cell())
        rr += 1
        self.label(ws, rr, 1, 'Échantillons sans zone ou sans strate', st.cell(bold=True))
        v = len(s) - m.n_zh.sum()
        self.formula(ws, rr, 2, f"=COUNTA('{S_ECH}'!A:A)-1-{ref(S_PAR, L['n_tot'], 2 + H)}", st.cell(INT), v)
        ws.write_number(rr, 3, 0, st.cell(INT))
        self.formula(ws, rr, 4, f'=IF({a1(rr, 2)}={a1(rr, 3)},"OK","À VÉRIFIER")', st.cell(align='center'),
                     'OK' if v == 0 else 'À VÉRIFIER')
        for code, sheet in [('Def', S_ADEF), ('Deg', S_ADEG), ('Gain', S_AGAIN)]:
            A = self.act[sheet]
            n_code = int((s.activity == code).sum())
            n_tr = A['rows'].n.sum()
            rr += 1
            self.label(ws, rr, 1, f'Échantillons {code}', st.cell(bold=True))
            self.formula(ws, rr, 2, f"=COUNTIF('{S_ECH}'!D:D,\"{code}\")", st.cell(INT), n_code)
            ws.write_blank(rr, 3, None, st.cell())
            ws.write_blank(rr, 4, None, st.cell())
            rr += 1
            self.label(ws, rr, 1, f'{code} sans transition valide (classe non attribuée ou catégorie inconnue)',
                       st.cell(bold=True))
            self.formula(ws, rr, 2, f"={a1(rr - 1, 2)}-{ref(sheet, A['total'], A['cols']['n'])}", st.cell(INT), n_code - n_tr)
            ws.write_number(rr, 3, 0, st.cell(INT))
            self.formula(ws, rr, 4, f'=IF({a1(rr, 2)}={a1(rr, 3)},"OK","À VÉRIFIER")', st.cell(align='center'),
                         'OK' if n_code == n_tr else 'À VÉRIFIER')
        status(r + 2, rr, 4)
        r = rr + 3

        # 2. somme des transitions vs estimation directe
        self.h2(ws, r, 1, 11, '2. Superficie : somme des transitions vs estimation directe (ha, période)')
        r += 1
        ws.merge_range(r, 1, r + 1, 1, 'Zone', st.hdr_l)
        for j, code in enumerate(['Def', 'Deg', 'Gain']):
            c0 = 2 + 3 * j
            ws.merge_range(r, c0, r, c0 + 2, code, st.hdr)
            for k, h in enumerate(['Σ transitions', 'Directe', 'Écart']):
                ws.write(r + 1, c0 + k, h, st.hdr2)
        ws.merge_range(r, 11, r + 1, 11, 'Statut', st.hdr)
        first = r + 2
        for zi, zn in enumerate(m.zones):
            rr = first + zi
            self.formula(ws, rr, 1, '=' + self.zone_ref(zi), st.cell(bold=True), zn)
            diffs = []
            for j, (code, sheet) in enumerate([('Def', S_ADEF), ('Deg', S_ADEG), ('Gain', S_AGAIN)]):
                c0, A = 2 + 3 * j, self.act[sheet]
                rows = A['rows']
                vs, vd = rows[rows.zone == zn].area.sum(), m.zone_act[code].area.iloc[zi]
                ac = self.col(sheet, A['cols']['area'], A['first'], A['last'])
                zc = self.col(sheet, 1, A['first'], A['last'])
                self.formula(ws, rr, c0, f'=SUMIFS({ac},{zc},{a1(rr, 1)})', st.cell(HA), vs)
                self.formula(ws, rr, c0 + 1, '=' + ref(sheet, A['zd0'] + zi, A['cols']['area']), st.cell(HA), vd)
                self.formula(ws, rr, c0 + 2, f'={a1(rr, c0)}-{a1(rr, c0 + 1)}', st.cell('#,##0.000'), vs - vd)
                diffs.append(abs(vs - vd))
            self.formula(ws, rr, 11, f'=IF(MAX(ABS({a1(rr, 4)}),ABS({a1(rr, 7)}),ABS({a1(rr, 10)}))<0.001,"OK","À VÉRIFIER")',
                         st.cell(align='center'), 'OK' if max(diffs) < 1e-3 else 'À VÉRIFIER')
        status(first, first + Z - 1, 11)
        r = first + Z + 3

        # 3. estimateur direct des émissions nettes DEF
        self.h2(ws, r, 1, 15, '3. Déforestation (biomasse) : estimateur direct des émissions nettes')
        r += 1
        self.note(ws, r, 1, 15, 'Chaque échantillon reçoit y = FE net de sa transition (0 s\'il n\'est pas déforesté). Ê_z = A_z · Σ_h W_hz · ȳ_hz '
                  'et ES² = A_z² · Σ_h W_hz² · s²_hz / n_hz. L\'estimation est identique à la somme des transitions ; son '
                  'incertitude d\'échantillonnage tient compte des covariances entre transitions (ignorées par l\'approche 1), '
                  'mais pas de l\'incertitude des FE.', 2)
        r += 3
        hdr = ['Zone'] + [f'Σy  h={h}' for h in m.strata] + [f'Σy²  h={h}' for h in m.strata] + [
            'Ê directe (tCO₂, période)', 'ES (tCO₂)', self.ci_label('U échantillonnage', ' (%)'),
            'Σ transitions (tCO₂)', 'Écart (tCO₂)', 'U AD seule, approche 1 (%)']
        self.hdr_row(ws, r, 1, hdr, height=44)
        d0 = r + 1
        e = self.em[S_EM]
        emc = lambda c: self.col(S_EM, c, e['first'], e['last'])
        c_e, c_se, c_u, c_sum, c_d, c_uad = 2 + 2 * H, 3 + 2 * H, 4 + 2 * H, 5 + 2 * H, 6 + 2 * H, 7 + 2 * H
        for zi, zn in enumerate(m.zones):
            rr, band = d0 + zi, zi % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            z_ref = self.zone_ref(zi)
            dr = m.direct.iloc[zi]
            self.formula(ws, rr, 1, '=' + z_ref, cf(None, bold=True), zn)
            for hi in range(H):
                h_ref = ref(S_PAR, L['az_hdr'], 2 + hi)
                self.formula(ws, rr, 2 + hi, f"=SUMIFS('{S_ECH}'!$J:$J,'{S_ECH}'!$C:$C,{z_ref},'{S_ECH}'!$B:$B,{h_ref})",
                             cf(T0), dr.sum_y[hi])
                self.formula(ws, rr, 2 + H + hi, f"=SUMIFS('{S_ECH}'!$K:$K,'{S_ECH}'!$C:$C,{z_ref},'{S_ECH}'!$B:$B,{h_ref})",
                             cf(SCI), dr.sum_y2[hi])
            Sy, Sy2 = f'{a1(rr, 2)}:{a1(rr, 1 + H)}', f'{a1(rr, 2 + H)}:{a1(rr, 1 + 2 * H)}'
            Wz = rng(S_PAR, L['w0'] + zi, 2, L['w0'] + zi, 1 + H)
            NH = rng(S_PAR, L['n0'] + zi, 2, L['n0'] + zi, 1 + H)
            Az = ref(S_PAR, L['az0'] + zi, 2 + H)
            self.formula(ws, rr, c_e, f'={Az}*SUMPRODUCT({Wz},{Sy}/{NH})', cf(T0, bold=True), dr.em_net)
            self.formula(ws, rr, c_se, f'={Az}*SQRT(SUMPRODUCT({Wz}^2*({Sy2}-{Sy}^2/{NH})/({NH}-1)/{NH}))', cf(T0), dr.se)
            self.formula(ws, rr, c_u, f'=IF({a1(rr, c_e)}<>0,z_IC*{a1(rr, c_se)}/ABS({a1(rr, c_e)}),0)', cf(PCT1),
                         m.z * dr.se / abs(dr.em_net) if dr.em_net else 0)
            sub = m.def_rows[m.def_rows.zone == zn]
            self.formula(ws, rr, c_sum, f'=SUMIFS({emc(17)},{emc(1)},{a1(rr, 1)})', cf(T0), sub.em_net.sum())
            self.formula(ws, rr, c_d, f'={a1(rr, c_e)}-{a1(rr, c_sum)}', cf('#,##0.000'), dr.em_net - sub.em_net.sum())
            uad = u_ratio(((sub.u_area * sub.em_net) ** 2).sum(), sub.em_net.sum())
            self.formula(ws, rr, c_uad, f'=IFERROR(SQRT(SUMPRODUCT(({emc(1)}={a1(rr, 1)})*({emc(6)}*{emc(17)})^2))'
                                        f'/ABS({a1(rr, c_sum)}),0)', cf(PCT1), uad)
        rr = d0 + Z
        ws.write(rr, 1, self.lab['Adj'], st.cell(bold=True, top=True))
        for c in range(2, 2 + 2 * H):
            ws.write_blank(rr, c, None, st.cell(top=True))
        dn, dse, rn = m.direct.em_net.sum(), np.sqrt((m.direct.se ** 2).sum()), m.def_rows.em_net.sum()
        tf = lambda nf: st.cell(nf, bold=True, top=True)
        self.formula(ws, rr, c_e, f'=SUM({a1(d0, c_e)}:{a1(rr - 1, c_e)})', tf(T0), dn)
        self.formula(ws, rr, c_se, f'=SQRT(SUMSQ({a1(d0, c_se)}:{a1(rr - 1, c_se)}))', tf(T0), dse)
        self.formula(ws, rr, c_u, f'=IF({a1(rr, c_e)}<>0,z_IC*{a1(rr, c_se)}/ABS({a1(rr, c_e)}),0)', tf(PCT1),
                     m.z * dse / abs(dn) if dn else 0)
        self.formula(ws, rr, c_sum, f'=SUM({a1(d0, c_sum)}:{a1(rr - 1, c_sum)})', tf(T0), rn)
        self.formula(ws, rr, c_d, f'={a1(rr, c_e)}-{a1(rr, c_sum)}', tf('#,##0.000'), dn - rn)
        self.formula(ws, rr, c_uad, f'=IFERROR(SQRT(SUMPRODUCT(({emc(6)}*{emc(17)})^2))/ABS({a1(rr, c_sum)}),0)', tf(PCT1),
                     u_ratio(((m.def_rows.u_area * m.def_rows.em_net) ** 2).sum(), rn))
        r = rr + 3

    # ============================================================ Graphiques
    def sheet_graphiques(self):
        st, m = self.st, self.m
        ws = self.sheet(S_GR, 'Graphiques', [14] * 16)
        ws.write(3, 1, 'Graphiques liés aux feuilles « Bilan » et « Résultats » (barres d\'erreur = ± incertitude).', st.note1)

        def err(sheet, r1, r2, c):
            rg = f"='{sheet}'!${xl_col_to_name(c)}${r1 + 1}:${xl_col_to_name(c)}${r2 + 1}"
            return {'type': 'custom', 'plus_values': rg, 'minus_values': rg, 'line': {'color': '#4B5563', 'width': 0.75}}

        def style(ch, title, y_name, legend=False):
            ch.set_y_axis({'name': y_name, 'num_format': '#,##0',
                           'major_gridlines': {'visible': True, 'line': {'color': '#E5E7EB'}}})
            ch.set_title({'name': title, 'name_font': {'size': 11, 'color': GREEN, 'bold': True}})
            ch.set_legend({'position': 'bottom'} if legend else {'none': True})
            ch.set_size({'width': 620, 'height': 330})
            ch.set_chartarea({'border': {'none': True}})
            return ch

        b = self.bil
        r1, r2 = b['first'], b['first'] + len(COMPONENTS) - 1
        ch = self.wb.add_chart({'type': 'column'})
        ch.add_series({'name': 'Bilan', 'categories': [S_BIL, r1, 1, r2, 1], 'values': [S_BIL, r1, 2, r2, 2],
                       'fill': {'color': CHART_COLORS[0]}, 'border': {'none': True}, 'gap': 60,
                       'points': [{'fill': {'color': c}} for c in
                                  [CHART_COLORS[0], CHART_COLORS[1], '#B45309', '#D97706', CHART_COLORS[3], '#374151',
                                   CHART_COLORS[2], '#111827']],
                       'y_error_bars': err(S_BIL, r1, r2, 4)})
        ws.insert_chart(5, 1, style(ch, 'Bilan annuel par composante (tCO₂/an)', 'tCO₂/an'))

        y0, y1 = b['year0'], b['year0'] + m.T - 1
        ch = self.wb.add_chart({'type': 'column', 'subtype': 'stacked'})
        yc = b['ycol']
        for k, color in [('def', CHART_COLORS[0]), ('soc', CHART_COLORS[1]), ('ch4', '#B45309'), ('n2o', '#D97706'),
                         ('deg', CHART_COLORS[3]), ('rscf', CHART_COLORS[2])]:
            c = yc[k]
            ch.add_series({'name': [S_BIL, b['year_h2'] + 1, c], 'categories': [S_BIL, y0, 1, y1, 1],
                           'values': [S_BIL, y0, c, y1, c], 'fill': {'color': color}, 'border': {'none': True}, 'gap': 60})
        line = self.wb.add_chart({'type': 'line'})
        cn = yc['net']
        line.add_series({'name': [S_BIL, b['year_h2'] + 1, cn], 'categories': [S_BIL, y0, 1, y1, 1],
                         'values': [S_BIL, y0, cn, y1, cn], 'line': {'color': '#111827', 'width': 2},
                         'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#111827'}},
                         'y_error_bars': err(S_BIL, y0, y1, cn + 2)})
        ch.combine(line)
        ws.insert_chart(5, 8, style(ch, 'Série par année (tCO₂/an)', 'tCO₂/an', legend=True))

        f, _ = self.res_em_rows
        ch = self.wb.add_chart({'type': 'column'})
        ch.add_series({'name': 'Bilan net', 'categories': [S_RES, f, 1, f + self.Z - 1, 1],
                       'values': [S_RES, f, self.res_net_col, f + self.Z - 1, self.res_net_col],
                       'fill': {'color': CHART_COLORS[0]}, 'border': {'none': True}, 'gap': 70,
                       'y_error_bars': err(S_RES, f, f + self.Z - 1, self.res_net_col + 2)})
        ws.insert_chart(23, 1, style(ch, 'Bilan net annuel par zone (tCO₂/an)', 'tCO₂/an'))

        f, _ = self.res_area_rows
        ch = self.wb.add_chart({'type': 'column'})
        for j, (name, color) in enumerate([('Déforestation', CHART_COLORS[0]), ('Dégradation', CHART_COLORS[3]),
                                           ('Gain', CHART_COLORS[2])]):
            c0 = 2 + 3 * j
            ch.add_series({'name': name, 'categories': [S_RES, f, 1, f + self.Z - 1, 1],
                           'values': [S_RES, f, c0, f + self.Z - 1, c0], 'fill': {'color': color}, 'border': {'none': True},
                           'y_error_bars': err(S_RES, f, f + self.Z - 1, c0 + 2)})
        ws.insert_chart(23, 8, style(ch, 'Superficies annuelles par zone (ha/an)', 'ha/an', legend=True))

    # ============================================================ Équations
    def sheet_equations(self):
        st = self.st
        ws = self.sheet(S_EQ, 'Équations', [5, 72, 50, 52])
        F_EQ = st.fmt(border=1, border_color=GRID, indent=1)
        F_EQ_B = st.fmt(border=1, border_color=GRID, indent=1, bg_color=BAND)
        F_NUM = st.fmt(font_color=GREY, align='center', border=1, border_color=GRID)
        F_DESC = st.fmt(text_wrap=True, border=1, border_color=GRID, font_color='#374151')
        F_XL = st.fmt(text_wrap=True, border=1, border_color=GRID, font_color=GREEN, font_name='Consolas', font_size=9)
        sections = [
            ('1 · Données d\'activité — estimateur stratifié par zone', [
                ('W_{hz} = A_{hz} / A_{z}', 'Poids de la strate h dans la zone z', 'Paramètres : bloc W_hz'),
                ('p̂_{kz} = Σ_{h} W_{hz} · n_{hkz} / n_{hz}', 'Proportion de la transition k (domaine) dans la zone z',
                 'Activité DEF / DEG / RSCF : SOMMEPROD(W_hz ; n_hkz / n_hz)'),
                ('SE(p̂_{kz}) = √[ Σ_{h} W_{hz}^{2} · p̂_{hkz} (1 − p̂_{hkz}) / (n_{hz} − 1) ]', 'Erreur standard', 'Activité'),
                ('Â_{kz} = A_{z} · p̂_{kz} ,   U_{A} = z · SE(Â_{kz}) / Â_{kz} ,   ± = U_{A} · Â_{kz}',
                 'Superficie de la période, incertitude en % et en ha', 'Activité'),
            ]),
            ('2 · Stocks de carbone', [
                ('C_{AGB} = AGB · CF ,   C_{BGB} = BGB · CF ,   C_{BM} = Σ BM_{i} · CF', 'Carbone par réservoir (tC/ha)',
                 'Stocks de carbone, tableau B'),
                ('U_{pool} = √( U_{mesure}^{2} + U_{CF}^{2} ) ,   U_{CF} = |CF_{bas} − CF| / CF', 'Incertitude d\'un réservoir',
                 'Stocks de carbone, tableau B'),
                ('S = (C_{AGB} + C_{BGB} + C_{L} + C_{BM}) · 44/12 ,   U_{S} = √Σ(U_{i} · C_{i})^{2} / C',
                 'Stock total (tCO₂/ha) et incertitude', 'Stocks de carbone, tableaux B et C'),
                ('U(niveau) = U(90 %) · z / z_{FE}', 'Conversion des incertitudes saisies', 'Facteur_U_FE'),
            ]),
            ('3 · Déforestation — biomasse et MOM', [
                ('FE_{net} = S_{forêt} − S_{post} ,   U = √[(U_{f}·S_{f})^{2} + (U_{p}·S_{p})^{2}] / |FE_{net}|',
                 'FE net et incertitude d\'une différence (GIEC 2006 vol. 1 éq. 3.2)', 'Émissions DEF'),
                ('E = Â · FE_{net} ,   U_{E} = √( U_{A}^{2} + U_{FE}^{2} ) ,   E^{an} = E / T',
                 'Émissions de la transition (éq. 3.1), émissions immédiates', 'Émissions DEF'),
            ]),
            ('4 · Déforestation — sols minéraux (GIEC 2006 vol. 4 éq. 2.25)', [
                ('F = F_{LU} · F_{MG} · F_{I} ,   ΔSOC = SOC_{ref} · (F − 1)', 'Variation du stock sur D années (forêt : F = 1)',
                 'Stocks de carbone, tableau D ; Émissions DEF SOC'),
                ('taux = − ΔSOC / D · 44/12', 'Émission annuelle par hectare déforesté (tCO₂/ha/an)', 'Émissions DEF SOC'),
                ('E = Â · taux · (T + 1) / 2 ,   E_{t} = (Â / T) · t · taux', 'Cohortes annuelles cumulées sur la période',
                 'Facteur_cohortes ; Bilan, tableau C'),
                ('U_{ΔSOC} = √[ U_{SOCref}^{2} + (U_{F} · F / |F − 1|)^{2} ] ,   U_{E} = √( U_{A}^{2} + U_{ΔSOC}^{2} )',
                 'Incertitude', 'Émissions DEF SOC'),
            ]),
            ('4 bis · Déforestation — combustion de la biomasse (GIEC 2006 vol. 4 éq. 2.27)', [
                ('MB = AGB + L / CF_{litière} + BM', 'Biomasse combustible (tms/ha), sans les racines', 'Stocks de carbone, '
                 'tableau B'),
                ('L_{gaz} = Â · f_{brûlée} · MB · C_{f} · G_{ef} · 10^{−3} ,   E = L_{gaz} · PRG',
                 'Tonnes de CH₄ ou de N₂O, puis tCO₂e (PRG AR5)', 'Émissions DEF CH4 N2O'),
                ('U_{E} = √( U_{A}^{2} + U_{MB}^{2} + U_{Cf}^{2} + U_{Gef}^{2} + U_{f}^{2} ) ,   E^{an} = E / T',
                 'Incertitude ; émissions immédiates', 'Émissions DEF CH4 N2O'),
            ]),
            ('5 · Dégradation', [
                ('FE_{deg} = S_{forêt 2021} − S_{forêt 2025} ,   U = √[(U_{i}·S_{i})^{2} + (U_{f}·S_{f})^{2}] / |FE_{deg}|',
                 'FE de dégradation (même zone)', 'Émissions DEG'),
                ('E = Â · FE_{deg} ,   U_{E} = √( U_{A}^{2} + U_{FE}^{2} ) ,   E^{an} = E / T', 'Émissions immédiates',
                 'Émissions DEG'),
            ]),
            ('6 · Renforcement des stocks de carbone forestier (gains)', [
                ('G = − G_{AGB} · (1 + R) · 44/12', 'Facteur d\'absorption (tCO₂/ha/an ; R = 0 : croissance aérienne seule)',
                 'Stocks de carbone, tableau E'),
                ('E = Â · G · (T + 1) / 2 ,   E_{t} = (Â / T) · t · G ,   U_{E} = √( U_{A}^{2} + U_{G}^{2} )',
                 'Cohortes annuelles cumulées', 'Absorptions RSCF'),
            ]),
            ('7 · Bilan', [
                ('E_{tot} = E_{DEF} + E_{SOC} + E_{CH4} + E_{N2O} + E_{DEG} ,   Bilan = E_{tot} + E_{RSCF}',
                 'Total des émissions et bilan net',
                 'Bilan, tableau A'),
                ('U = √Σ_{c} (U_{c} · E_{c})^{2} / |Σ_{c} E_{c}|', 'Combinaison de composantes indépendantes (éq. 3.2)',
                 'Bilan ; Résultats'),
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
                self.rich(ws, r, 2, expr, F_EQ_B if i % 2 else F_EQ, font_name='Cambria', font_size=12)
                ws.write(r, 3, desc, F_DESC)
                self.rich(ws, r, 4, xl, F_XL, font_name='Consolas', font_size=9, font_color=GREEN)
                r, no = r + 1, no + 1
            r += 1
        self.note(ws, r, 1, 4, 'Références : GIEC (2006) Lignes directrices, vol. 1 ch. 3 et vol. 4 ch. 2, 4 et 5 ; GIEC (2019) '
                  'Affinement ; Cochran (1977) Sampling Techniques ; Olofsson et al. (2014) Remote Sensing of Environment 148.', 2)

    # ============================================================ Échantillons
    def sheet_echantillons(self):
        st, m = self.st, self.m
        ws = self.wb.add_worksheet(S_ECH)
        s = m.samples.reset_index(drop=True)
        headers = ['Échantillon', 'Strate', 'Zone', 'Activité', 'Classe CEO 2021', 'Classe CEO 2025',
                   'Catégorie forestière 2021', 'Utilisation post-déforestation', 'Clé DEF', 'y = FE net DEF (tCO₂/ha)', 'y²',
                   'Catégorie forestière 2025', 'Clé DEG', 'Catégorie de gain', 'Clé RSCF'] + [h for h, _ in self.extra_cols]
        widths = [12, 8, 22, 9, 26, 26, 26, 26, 46, 14, 12, 26, 46, 22, 34] + [12] * len(self.extra_cols)
        for j, (h, w) in enumerate(zip(headers, widths)):
            ws.write(0, j, h, st.hdr if j else st.hdr_l)
            ws.set_column(j, j, w)
        ws.set_row(0, 30)
        f_in = st.fmt(bg_color=INPUT_BG, font_color=INPUT_INK)
        f_id = st.fmt(font_color=GREY)
        f_calc = st.fmt()
        f_key = st.fmt(font_color=GREY, font_size=8)
        ws.write_column(1, 0, s['plotid'].tolist(), f_id)
        ws.write_column(1, 1, [int(v) for v in s['strata']], f_in)
        ws.write_column(1, 2, s['zone'].tolist(), f_in)
        ws.write_column(1, 3, s['activity'].tolist(), f_in)
        ws.write_column(1, 4, s['class_initial'].fillna('').tolist(), f_in)
        ws.write_column(1, 5, s['class_final'].fillna('').tolist(), f_in)
        U = UNASSIGNED
        for i in range(len(s)):
            r, R = i + 1, i + 2
            row = s.iloc[i]
            ws.write_formula(r, 6, f'=IF(OR($D{R}="Def",$D{R}="Deg"),IFERROR(VLOOKUP($E{R},Corr_foret,2,FALSE),"{U}"),"")',
                             f_calc, row.fcat_i)
            ws.write_formula(r, 7, f'=IF($D{R}="Def",IFERROR(VLOOKUP($F{R},Corr_post,2,FALSE),"{U}"),"")', f_calc, row.pcat)
            ws.write_formula(r, 8, f'=IF($D{R}="Def",$C{R}&"|"&G{R}&"|"&H{R},"")', f_key, row.key_def)
            ws.write_formula(r, 9, f'=IF($I{R}="",0,IFERROR(INDEX(FE_net_transitions,MATCH($I{R},Cle_transitions,0)),0))',
                             f_calc, row.y)
            ws.write_formula(r, 10, f'=J{R}^2', f_calc, row.y2)
            ws.write_formula(r, 11, f'=IF($D{R}="Deg",IFERROR(VLOOKUP($F{R},Corr_foret,2,FALSE),"{U}"),"")', f_calc, row.fcat_f)
            ws.write_formula(r, 12, f'=IF($D{R}="Deg",$C{R}&"|"&G{R}&"|"&L{R},"")', f_key, row.key_deg)
            ws.write_formula(r, 13, f'=IF($D{R}="Gain",IFERROR(VLOOKUP($F{R},Corr_gain,2,FALSE),"{U}"),"")', f_calc, row.gcat)
            ws.write_formula(r, 14, f'=IF($D{R}="Gain",$C{R}&"|"&N{R},"")', f_key, row.key_gain)
        for j, (_, colname) in enumerate(self.extra_cols, start=15):
            ws.write_column(1, j, [_clean(v) for v in s[colname]], f_id)
        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, len(s), len(headers) - 1)
        ws.data_validation(1, 3, len(s) + 2000, 3, {'validate': 'list',
                                                    'source': ['SNF', 'SF', 'SDegF', 'Deg', 'Def', 'Gain']})


def build_emissions_workbook(path, samples, area_hz, cats, maps, forest_stocks, post_stocks, soc_factors, removal_rates,
                             confidence=0.90, ef_confidence=0.90, initial_year=2021, final_year=2025, cf=0.47, cf_low=0.44,
                             socref=47.0, u_socref=0.0, soc_d=20, low_n=10, fire=None, fire_fraction=None, title=None,
                             subtitle=None, sources=None, extra_cols=(), territory=None):
    """Écrit le classeur et renvoie le modèle Python (valeurs identiques aux formules du classeur).

    samples : DataFrame (plotid, strata, zone, activity, class_initial, class_final, + extra_cols), échantillons valides.
    area_hz : DataFrame indexé par zone, une colonne par strate (ha).
    cats / maps : dict avec les clés 'forest', 'post', 'gain' (listes ordonnées de catégories / dict classe CEO -> catégorie).
    forest_stocks, post_stocks, soc_factors, removal_rates : DataFrame (voir le notebook 3).
    fire : dict des paramètres de combustion (cf_litter, comb, u_comb, gef_ch4, u_gef_ch4, gef_n2o, u_gef_n2o, gwp_ch4,
           gwp_n2o, u_fire) ; valeurs du NERF 2024 par défaut.
    fire_fraction : dict (catégorie forestière, utilisation post-déforestation) -> fraction brûlée (0,5 par défaut).
    territory : None (total « national ») ou nom du territoire couvert par area_hz (ex. 'PRE') ; ne change que les libellés.
    """
    params = dict(confidence=confidence, ef_confidence=ef_confidence, initial_year=initial_year, final_year=final_year,
                  cf=cf, cf_low=cf_low, socref=socref, u_socref=u_socref, soc_d=soc_d, low_n=low_n, **(fire or {}))
    model = EmissionModel(samples, area_hz, cats, maps, forest_stocks, post_stocks, soc_factors, removal_rates, params,
                          fire_fraction)
    title = title or "Côte d'Ivoire — Émissions et absorptions forestières"
    subtitle = subtitle or (f'Projet PROMIRE / ONU-REDD · SBAE {initial_year}-{final_year} · déforestation (biomasse, sols, '
                            f'combustion), dégradation, renforcement des stocks · généré le {datetime.date.today():%d/%m/%Y}')
    b = _EmissionBuilder(path, model, title, subtitle, params, sources or {}, extra_cols, territory)
    b.build()
    return model
