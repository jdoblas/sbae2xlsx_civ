"""Live Excel workbook of forest emissions and removals, by phytogeographic zone and at the national
level, from the SBAE samples and the national emission factors (FREL 2024).

Components:

    DEF      deforestation - biomass and dead organic matter (net EF = forest stock − post-conversion stock)
    DEF SOC  deforestation - mineral soil organic carbon (IPCC eq. 2.25, D = 20 years)
    DEF CH4  deforestation - biomass burning, CH₄ and N₂O (IPCC eq. 2.27:
    DEF N2O    L = A_burnt · MB · Cf · Gef · 10⁻³, then × GWP; MB = AGB + litter + deadwood, in t d.m./ha)
    DEG      degradation - EF = stock of the initial forest − stock of the final forest
    GAIN     enhancement of forest carbon stocks - gains (non-forest -> forest), annual growth rate

Calculation chain (all result cells are Excel formulas):

    Samples -> Mappings -> Activity DEF / DEG / GAIN (stratified estimator by zone)
      -> Carbon stocks (stocks, fuel biomass, SOC factors, growth rates)
      -> Emissions DEF, Emissions DEF SOC, Emissions DEF CH4 N2O, Emissions DEG, Removals GAIN
      -> Balance, Results, Checks

Editable inputs are in yellow. The values computed in Python (EmissionModel) are written as cached values of the
formulas and returned for checking.

References: IPCC (2006, 2019) vol. 1 ch. 3, vol. 4 ch. 2, 4 and 5; Cochran (1977); Olofsson et al. (2014).
"""
import datetime
from statistics import NormalDist

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from sbae_excel_en import (BAND, CHART_COLORS, GREEN, GREY, GRID, HA, INPUT_BG, INPUT_INK, INT, WorkbookBase, _clean,
                        _Styles, a1, ref, rng, territory_labels)

# ---------------------------------------------------------------- sheets
S_SYN = 'Summary'
S_PAR = 'Parameters'
S_COR = 'Mappings'
S_STK = 'Carbon stocks'
S_ADEF = 'Activity DEF'
S_ADEG = 'Activity DEG'
S_AGAIN = 'Activity GAIN'
S_EM = 'Emissions DEF'
S_SOC = 'Emissions DEF SOC'
S_FIRE = 'Emissions DEF CH4 N2O'
S_DEG = 'Emissions DEG'
S_ENH = 'Removals GAIN'
S_BIL = 'Balance'
S_RES = 'Results'
S_CTL = 'Checks'
S_GR = 'Charts'
S_EQ = 'Equations'
S_ECH = 'Samples'

SHEETS = [
    (S_PAR, 'General parameters, soil parameters, areas A_hz, weights W_hz and sample counts n_hz'),
    (S_COR, 'Mapping of the CEO classes to the stock, soil and gain categories (editable)'),
    (S_STK, 'Carbon stocks, emission factors, SOC factors and growth rates'),
    (S_ADEF, 'Deforested area by zone and transition (stratified estimator)'),
    (S_ADEG, 'Degraded area by zone and transition'),
    (S_AGAIN, 'Gain area by zone and transition (2021 category → 2025 gain category)'),
    (S_EM, 'Deforestation: emissions from biomass and DOM (gross and net EF)'),
    (S_SOC, 'Deforestation: emissions from mineral soil organic carbon'),
    (S_FIRE, 'Deforestation: CH₄ and N₂O emissions from biomass burning'),
    (S_DEG, 'Degradation: emissions'),
    (S_ENH, 'Enhancement of forest carbon stocks: removals'),
    (S_BIL, 'Balance: annual averages by component and yearly series'),
    (S_RES, 'Results by zone, forest category and post-deforestation land use'),
    (S_CTL, 'Consistency checks and direct estimator'),
    (S_GR, 'Charts (± uncertainty)'),
    (S_EQ, 'Equations and implementation in the workbook'),
    (S_ECH, 'Valid samples: stratum, zone, activity, CEO classes 2021 and 2025 (editable)'),
]

UNASSIGNED = 'UNASSIGNED'
PCT1 = '0.0%;-0.0%;"–"'
HA0 = '#,##0;-#,##0;"–"'
T0 = '#,##0;-#,##0;"–"'
EF2 = '#,##0.00;-#,##0.00;"–"'
F3 = '0.000;-0.000;"–"'
DM2 = '0.00'
DIFF = '0.000000;-0.000000;"0"'
SCI = '0.00E+00;-0.00E+00;"–"'
POOLS = ['agb', 'bgb', 'litter', 'dw']
COMPONENTS = [('def', 'Deforestation — biomass and DOM (net EF)'), ('soc', 'Deforestation — mineral soils'),
              ('ch4', 'Deforestation — biomass burning, CH₄'),
              ('n2o', 'Deforestation — biomass burning, N₂O'),
              ('deg', 'Degradation'), ('tot', 'Total emissions'), ('enh', 'Enhancement of stocks (gains)'),
              ('net', 'Net balance (emissions + removals)')]
EMISSIONS = ['def', 'soc', 'ch4', 'n2o', 'deg']     # components of the total emissions
GASES = [('ch4', 'CH₄'), ('n2o', 'N₂O')]
FIRE_DEFAULT = 0.5      # burnt fraction of a transition missing from the "fire_fraction" table


def z_of(conf):
    """Two-sided normal quantile rounded to 3 decimals (1.645 for 90%, 1.960 for 95%)."""
    return round(NormalDist().inv_cdf(1 - (1 - conf) / 2), 3)


def u_ratio(ux2, total):
    return np.sqrt(ux2) / abs(total) if total else 0.0


# ================================================================ model (cached values of the formulas)
class EmissionModel:
    """Computes in Python all the quantities written as formulas in the workbook."""

    def __init__(self, samples, area_hz, cats, maps, forest_stocks, post_stocks, soc_factors, removal_rates, p,
                 fire_fraction=None):
        # biomass burning: FREL 2024 values by default
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
        # 2021 category of the gains: non-forest land use, or forest category (gain to a higher-ranked forest)
        self.gain_init_cats = self.post_cats + self.forest_cats
        self.gain_init_map = {**self.forest_map, **self.post_map}
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
        s['gcat_i'] = mapped(act == 'Gain', s['class_initial'], self.gain_init_map)
        s['gcat'] = mapped(act == 'Gain', s['class_final'], self.gain_map)
        s['key_def'] = np.where(act == 'Def', s['zone'] + '|' + s['fcat_i'] + '|' + s['pcat'], '')
        s['key_deg'] = np.where(act == 'Deg', s['zone'] + '|' + s['fcat_i'] + '|' + s['fcat_f'], '')
        s['key_gain'] = np.where(act == 'Gain', s['zone'] + '|' + s['gcat_i'] + '|' + s['gcat'], '')
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
        self._balance()

    # ------------------------------------------------------------ stocks and factors
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
            # fuel biomass (t d.m./ha): AGB + litter + deadwood, without roots (IPCC 2006 eq. 2.27)
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

    # ------------------------------------------------------------ stratified estimator by zone
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
        """items: list of dicts (zone, zi, key, + labels) -> DataFrame with areas and uncertainties."""
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

    # ------------------------------------------------------------ components
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
        """Biomass burning during deforestation: L = A · burnt fraction · MB · Cf · Gef · 10⁻³ (t gas)."""
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
        items = [dict(zone=zn, zi=zi, g_i=gi, gain=g, key=f'{zn}|{gi}|{g}')
                 for zi, zn in enumerate(self.zones) for gi in self.gain_init_cats for g in self.gain_cats]
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
        """Direct estimator of the net DEF emissions: stratified mean of the net EF of each sample."""
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

    def _balance(self):
        """Annual averages by component (national and by zone) and yearly series."""
        T = self.T
        src = {'def': (self.def_rows, 'em_net', 'ux2_net'), 'soc': (self.soc_rows, 'em', 'ux2'),
               'ch4': (self.fire_rows, 'em_ch4', 'ux2_ch4'), 'n2o': (self.fire_rows, 'em_n2o', 'ux2_n2o'),
               'deg': (self.deg_rows, 'em', 'ux2'), 'enh': (self.gain_rows, 'em', 'ux2')}

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
        nat['net'] = combine([nat['tot'], nat['enh']])
        self.balance = nat
        zones = {}
        for zn in self.zones:
            d = {c: comp(df[df.zone == zn], col, ux) for c, (df, col, ux) in src.items()}
            d['tot'] = combine([d[c] for c in EMISSIONS])
            d['net'] = combine([d['tot'], d['enh']])
            zones[zn] = d
        self.balance_zone = zones
        years = []
        for t in range(1, T + 1):
            f = 2 * t / (T + 1)
            y = {'def': nat['def']['year'], 'soc': nat['soc']['year'] * f, 'ch4': nat['ch4']['year'],
                 'n2o': nat['n2o']['year'], 'deg': nat['deg']['year'], 'enh': nat['enh']['year'] * f}
            pm = {c: nat[c]['u'] * abs(y[c]) for c in y}
            y['tot'] = sum(y[c] for c in EMISSIONS)
            y['pm_tot'] = np.sqrt(sum(pm[c] ** 2 for c in EMISSIONS))
            y['net'] = y['tot'] + y['enh']
            y['pm_net'] = np.sqrt(y['pm_tot'] ** 2 + pm['enh'] ** 2)
            y['year'] = self.p['initial_year'] + t
            years.append(y)
        self.years = pd.DataFrame(years)


# ================================================================ workbook builder
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
        L = dict(conf=5, z=6, conf_fe=7, z_fe=8, k_fe=9, y0=10, y1=11, years=12, coh=13, low_n=14, cf=15, cf_low=16,
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
        # activity sheets: one row per transition from row 6 (index 5), then the zone table
        self.act = {}
        for name, rows, labels, keycol, code in [
                (S_ADEF, m.def_rows, ['forest', 'post'], 'I', 'Def'),
                (S_ADEG, m.deg_rows, ['f_i', 'f_f'], 'M', 'Deg'),
                (S_AGAIN, m.gain_rows, ['g_i', 'gain'], 'P', 'Gain')]:
            k = len(labels)
            c = dict(zone=1, key=2 + k, s0=3 + k)
            c.update(n=3 + k + self.H, p=4 + k + self.H, area=5 + k + self.H, se=6 + k + self.H, u=7 + k + self.H,
                     pm=8 + k + self.H, ay=9 + k + self.H)
            first, last = 6, 6 + len(rows) - 1
            zd_h2 = last + 4
            self.act[name] = dict(rows=rows, labels=labels, keycol=keycol, code=code, cols=c, first=first, last=last,
                                  total=last + 1, zd_h2=zd_h2, zd_hdr=zd_h2 + 1, zd0=zd_h2 + 2, zd_nat=zd_h2 + 2 + Z)
        # Balance
        self.bil = dict(first=6, rows={c: 6 + i for i, (c, _) in enumerate(COMPONENTS)})
        self.bil['area0'] = self.bil['first'] + len(COMPONENTS) + 3
        self.bil['year_h2'] = self.bil['area0'] + 3 + 3
        self.bil['year0'] = self.bil['year_h2'] + 2

    # ------------------------------------------------------------ utilities
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
        self.wb.set_properties({'title': self.title, 'subject': 'Forest emissions and removals - SBAE',
                                'company': 'FAO'})
        self.wb.set_calc_mode('auto')
        self.wb.set_size(1600, 1000)
        L = self.L
        for name, row in [('Confidence_level', 'conf'), ('z_IC', 'z'), ('Confidence_level_EF', 'conf_fe'),
                          ('z_EF', 'z_fe'), ('U_EF_factor', 'k_fe'), ('Initial_year', 'y0'), ('Final_year', 'y1'),
                          ('Period_length', 'years'), ('Cohort_factor', 'coh'), ('Small_n_threshold', 'low_n'), ('CF', 'cf'),
                          ('CF_low', 'cf_low'), ('U_CF', 'u_cf'), ('CO2_C', 'co2c'), ('SOC_ref', 'socref'),
                          ('U_SOC_ref', 'u_socref'), ('SOC_D', 'soc_d'), ('CF_litter', 'cf_litter'),
                          ('Combustion_factor', 'comb'), ('U_Combustion_factor', 'u_comb'), ('Gef_CH4', 'gef_ch4'),
                          ('U_Gef_CH4', 'u_gef_ch4'), ('Gef_N2O', 'gef_n2o'), ('U_Gef_N2O', 'u_gef_n2o'),
                          ('GWP_CH4', 'gwp_ch4'), ('GWP_N2O', 'gwp_n2o'), ('U_Burnt_fraction', 'u_fire')]:
            self.wb.define_name(name, '=' + ref(S_PAR, L[row], 2))
        self.sheet_summary()
        self.sheet_parameters()
        self.sheet_mappings()
        self.sheet_stocks()
        for name, title in [(S_ADEF, 'Deforested area by transition'),
                            (S_ADEG, 'Degraded area by transition'),
                            (S_AGAIN, 'Gain area by transition')]:
            self.sheet_activity(name, title)
        self.sheet_emissions_def()
        self.sheet_soc()
        self.sheet_fire()
        self.sheet_deg()
        self.sheet_enh()
        self.sheet_balance()
        self.sheet_results()
        self.sheet_checks()
        self.sheet_charts()
        self.sheet_equations()
        self.sheet_samples()
        self.wb.worksheets()[0].activate()
        self.wb.close()

    # ============================================================ Summary
    def sheet_summary(self):
        st, m = self.st, self.m
        ws = self.sheet(S_SYN, 'Forest emissions and removals — dynamic workbook', [46, 18, 14, 16, 14, 40])
        r = 4
        self.h2(ws, r, 1, 6, 'Parameters')
        s = m.samples
        par = [
            ('Reference data (CEO)', self.sources.get('data', ''), None, None),
            ('Period', '=Initial_year&" – "&Final_year&" ("&Period_length&" years)"', None,
             f'{self.p["initial_year"]} – {self.p["final_year"]} ({m.T} years)'),
            ('Valid samples', f"=COUNTA('{S_ECH}'!A:A)-1", INT, len(s)),
            ('Def / Deg / Gain samples',
             f"=COUNTIF('{S_ECH}'!D:D,\"Def\")&\" / \"&COUNTIF('{S_ECH}'!D:D,\"Deg\")&\" / \"&COUNTIF('{S_ECH}'!D:D,\"Gain\")",
             None, f'{(s.activity == "Def").sum()} / {(s.activity == "Deg").sum()} / {(s.activity == "Gain").sum()}'),
            ('Confidence level', '=Confidence_level', '0%', self.p['confidence']),
            ('Components', 'DEF biomass and DOM, DEF mineral soils, DEF burning (CH₄, N₂O), degradation, '
                            'enhancement of stocks', None, None),
            ('Method', 'IPCC 2006/2019 (approach 1); stratified estimator by zone (Cochran 1977)', None, None),
        ]
        for i, (k, v, nf, cached) in enumerate(par, start=r + 1):
            self.label(ws, i, 1, k, st.key)
            ws.merge_range(i, 2, i, 6, '', st.cell(nf, align='left'))
            if isinstance(v, str) and v.startswith('='):
                self.formula(ws, i, 2, v, st.cell(nf, align='left'), cached)
            else:
                ws.write(i, 2, v, st.cell(align='left'))
        r += len(par) + 2

        self.h2(ws, r, 1, 6, self.lab['Adj'] + ' results — annual averages')
        r += 1
        self.hdr_row(ws, r, 1, ['', 'Value', 'Uncertainty (%)', '± (value)', 'Unit'])
        b, bil = self.bil, m.balance
        lines = []
        for name, code in [('Deforested area', 'Def'), ('Degraded area', 'Deg'),
                           ('Gain area', 'Gain')]:
            rr = b['area0'] + ['Def', 'Deg', 'Gain'].index(code)
            nat = m.zone_act[code].attrs['national']
            lines.append((name, ref(S_BIL, rr, 2), nat['area_year'], ref(S_BIL, rr, 3), nat['u_area'], 'ha/yr'))
        for code, name in COMPONENTS:
            rr = b['rows'][code]
            lines.append((name, ref(S_BIL, rr, 2), bil[code]['year'], ref(S_BIL, rr, 3), bil[code]['u'], 'tCO₂/yr'))
        for i, (name, fv, vv, fu, vu, unit) in enumerate(lines, start=r + 1):
            band = (i - r) % 2 == 0
            bold = name.startswith(('Total', 'Net'))
            ws.write(i, 1, name, st.cell(band=band, bold=True))
            self.formula(ws, i, 2, '=' + fv, st.cell(T0, band=band, bold=bold), vv)
            self.formula(ws, i, 3, '=' + fu, st.cell(PCT1, band=band), vu)
            self.formula(ws, i, 4, f'=ABS({a1(i, 2)})*{a1(i, 3)}', st.cell(T0, band=band), abs(vv) * vu)
            ws.write(i, 5, unit, st.cell(band=band))
        r += len(lines) + 1
        self.note(ws, r, 1, 6, 'Removals as negative values. Uncertainties: half-width of the confidence interval, in % '
                  'of the value and in absolute value (±). Components combined with IPCC approach 1 (see Balance).', 2)
        r += 3

        self.h2(ws, r, 1, 6, 'How to use')
        r += 1
        ws.write(r, 1, 'Input cell (editable)', st.cell(inp=True))
        ws.merge_range(r, 2, r, 6, 'Parameters, areas, mappings, stocks, factors, sample data.',
                       st.text)
        r += 1
        ws.write(r, 1, 'Computed cell (formula)', st.cell())
        ws.merge_range(r, 2, r, 6, 'All other cells are formulas and recalculate automatically.', st.text)
        r += 1
        self.note(ws, r, 1, 6,
                  'Scenarios: change the parameters and the areas A_hz (Parameters), the mappings of the CEO classes '
                  '(Mappings), the stocks, SOC factors and growth rates (Carbon stocks) or the classes of the '
                  'samples (Samples). The zone and category names must remain identical in all '
                  'sheets. To return to the original values, rerun the notebook.', 3)
        r += 4
        self.h2(ws, r, 1, 6, 'Workbook contents')
        for i, (name, desc) in enumerate(SHEETS, start=r + 1):
            ws.write_url(i, 1, f"internal:'{name}'!A1", st.link, name)
            ws.merge_range(i, 2, i, 6, '', st.text)
            self.label(ws, i, 2, desc, st.text)

    # ============================================================ Parameters
    def sheet_parameters(self):
        st, L, m, H, Z = self.st, self.L, self.m, self.H, self.Z
        ws = self.sheet(S_PAR, 'Parameters and mapped areas', [44, 16, 16, 16, 16, 50])
        ws.write(3, 1, 'Yellow cells: editable values. The other cells are computed.', st.note1)
        self.h2(ws, 4, 1, 6, 'General parameters')
        p = self.p
        items = [
            ('conf', 'Confidence level of the results', p['confidence'], '0%', 'Two-sided interval'),
            ('z', 'Normal quantile  z', '=ROUND(_xlfn.NORM.S.INV(1-(1-Confidence_level)/2),3)', '0.000', m.z),
            ('conf_fe', 'Confidence level of the input uncertainties', p['ef_confidence'], '0%',
             'Stocks, SOC factors, growth rates (90% CI)'),
            ('z_fe', 'Normal quantile  z_EF', '=ROUND(_xlfn.NORM.S.INV(1-(1-Confidence_level_EF)/2),3)', '0.000', m.z_fe),
            ('k_fe', 'Conversion of the input uncertainties  z / z_EF', '=z_IC/z_EF', '0.0000', m.k_fe),
            ('y0', 'Initial year', p['initial_year'], '0', 'Reference state of the samples'),
            ('y1', 'Final year', p['final_year'], '0', ''),
            ('years', 'Length of the period  T (years)', '=Final_year-Initial_year', '0', m.T),
            ('coh', 'Cohort factor  (T + 1) / 2', '=(Period_length+1)/2', '0.00', m.coh),
            ('low_n', 'Small-sample threshold  n', p['low_n'], INT, 'Flags the transitions with few samples'),
            ('cf', 'Carbon fraction of biomass  CF', m.cf, '0.000', 'IPCC 2006, table 4.3'),
            ('cf_low', 'Lower bound of CF', m.cf_low, '0.000', 'IPCC 2006, table 4.3 (0.44 – 0.49)'),
            ('u_cf', 'Uncertainty of CF  U_CF', '=ABS(CF_low-CF)/CF', PCT1, m.u_cf),
            ('co2c', 'CO₂ / C ratio', '=44/12', '0.0000', 44 / 12),
            ('socref', 'Soil reference stock  SOC_ref (tC/ha)', p['socref'], '0.0',
             'LAC mineral soils, tropical moist climate (IPCC table 2.3)'),
            ('u_socref', 'Uncertainty of SOC_ref', p['u_socref'], PCT1, 'Not available: to be completed'),
            ('soc_d', 'Soil transition period  D (years)', p['soc_d'], '0', 'IPCC: 20 years by default'),
            ('cf_litter', 'Carbon fraction of litter', m.p['cf_litter'], '0.00',
             'IPCC 2006: 0.37; litter tC/ha → t d.m./ha for the fuel biomass'),
            ('comb', 'Combustion factor  Cf', m.p['comb'], '0.00', 'IPCC 2006 table 2.6'),
            ('u_comb', 'Uncertainty of Cf', m.p['u_comb'], PCT1, 'Not available: to be completed'),
            ('gef_ch4', 'Emission factor  Gef CH₄ (g/kg d.m. burnt)', m.p['gef_ch4'], '0.00',
             'IPCC 2006 table 2.5, tropical forest'),
            ('u_gef_ch4', 'Uncertainty of Gef CH₄', m.p['u_gef_ch4'], PCT1, 'Not available: to be completed'),
            ('gef_n2o', 'Emission factor  Gef N₂O (g/kg d.m. burnt)', m.p['gef_n2o'], '0.00',
             'IPCC 2006 table 2.5, tropical forest'),
            ('u_gef_n2o', 'Uncertainty of Gef N₂O', m.p['u_gef_n2o'], PCT1, 'Not available: to be completed'),
            ('gwp_ch4', 'GWP of CH₄ (tCO₂e / t)', m.p['gwp_ch4'], '0', 'IPCC AR5'),
            ('gwp_n2o', 'GWP of N₂O (tCO₂e / t)', m.p['gwp_n2o'], '0', 'IPCC AR5'),
            ('u_fire', 'Uncertainty of the burnt fraction', m.p['u_fire'], PCT1,
             'Fractions by transition: Mappings; uncertainty to be completed'),
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
                            border_color='white', num_format='"Stratum "0')

        def header(r, first, last):
            self.label(ws, r, 1, first, st.hdr_l)
            for hi, h in enumerate(m.strata):
                ws.write_number(r, 2 + hi, h, strata_fmt)
            self.label(ws, r, 2 + H, last, st.hdr2)

        self.h2(ws, L['az_h2'], 1, 2 + H, 'Mapped area by zone and stratum  A_hz (ha)')
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
            ws.write(r + 1, 1, f'Source: {self.sources["areas"]}', st.note1)

        self.h2(ws, L['w_h2'], 1, 2 + H, 'Weights  W_hz = A_hz / A_z')
        header(L['w_hdr'], 'Zone', 'Σ_h W_hz')
        for zi, zn in enumerate(m.zones):
            r, band = L['w0'] + zi, zi % 2 == 1
            self.formula(ws, r, 1, '=' + xl_rowcol_to_cell(L['az0'] + zi, 1, True, True), st.cell(band=band, bold=True), zn)
            for hi in range(H):
                self.formula(ws, r, 2 + hi,
                             f'={a1(L["az0"] + zi, 2 + hi)}/{xl_rowcol_to_cell(L["az0"] + zi, 2 + H, True, True)}',
                             st.cell('0.0000', band=band), m.W_zh[zi, hi])
            self.formula(ws, r, 2 + H, f'=SUM({a1(r, 2)}:{a1(r, 1 + H)})', st.cell('0.0000', band=band, bold=True), 1.0)

        self.h2(ws, L['n_h2'], 1, 2 + H, 'Samples by zone and stratum  n_hz  (COUNTIFS on "Samples")')
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

    # ============================================================ Mappings
    def sheet_mappings(self):
        st, m = self.st, self.m
        ws = self.sheet(S_COR, 'Mapping of the CEO classes (level 2) to the categories',
                        [30, 30, 3, 28, 26, 3, 30, 26, 3, 36])
        ws.write(3, 1, 'The categories are editable (drop-down lists). A class missing from the tables receives '
                       '"UNASSIGNED" (see Checks).', st.note1)
        r0, d0 = 5, 6
        tables = [
            (1, 'Forest CEO class (2021 or 2025)', 'Forest category (stock)', m.forest_map, 'Map_forest'),
            (4, 'Post-deforestation CEO class (2025)', 'Land use (stock and soils)', m.post_map, 'Map_post'),
            (7, 'Forest CEO class of the gains (2025)', 'Gain category (growth)', m.gain_map, 'Map_gain'),
        ]
        lists = [('Forest categories', m.forest_cats), ('Post-deforestation land uses', m.post_cats),
                 ('Gain categories', m.gain_cats)]
        # category lists (sources of the drop-down lists)
        self.hdr_row(ws, r0, 10, ['Available categories'])
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
        # burnt fraction by deforestation transition (biomass burning)
        r = last + 3
        self.h2(ws, r, 1, 5, 'Burnt fraction of the deforested area, by transition (biomass burning)')
        self.hdr_row(ws, r + 1, 1, ['Forest category 2021', 'Post-deforestation land use', '',
                                    'Burnt fraction', 'Key'], left=2)
        f0 = r + 2
        pairs = [(f, u) for f in m.forest_cats for u in m.post_cats]
        for i, (f, u) in enumerate(pairs):
            rr = f0 + i
            ws.write_string(rr, 1, f, st.cell(bold=True))
            ws.write_string(rr, 2, u, st.cell())
            self.inp(ws, rr, 4, m.fire_fraction.get((f, u), FIRE_DEFAULT), '0.00')
            self.formula(ws, rr, 5, f'={a1(rr, 1)}&"|"&{a1(rr, 2)}', self.key_fmt(False, False), f'{f}|{u}')
        f1 = f0 + len(pairs) - 1
        self.wb.define_name('Fire_fraction', '=' + rng(S_COR, f0, 4, f1, 4))
        self.wb.define_name('Fire_key', '=' + rng(S_COR, f0, 5, f1, 5))
        last = f1 + 1
        notes = list(self.sources.get('mapping_notes', [])) + [
            'Burnt fraction: FREL 2024 values — 0.5 of the deforested area, 0 for the '
            'conversions to fallow / thicket or to grassland and for agroforest converted to cocoa or to perennial '
            'crops. A transition missing from the table receives 0.5.']
        for i, t in enumerate(notes):
            self.note(ws, last + 2 + i, 1, 10, t)

    # ============================================================ Carbon stocks
    def sheet_stocks(self):
        st, m = self.st, self.m
        ws = self.sheet(S_STK, 'Carbon stocks, SOC factors and growth rates', [28, 24] + [11] * 20)
        ws.write(3, 1, 'Biomass in tonnes of dry matter (t d.m./ha), carbon in tC/ha. Uncertainties = half-width of '
                       'the CI at the "Confidence level of the input uncertainties" (Parameters).', st.note1)
        fr = list(m.forest_stocks.values())
        self.stk = {}
        # --- A. forest inputs
        r = 5
        self.h2(ws, r, 1, 20, 'A. Forest stocks — inputs (NFI Tier 2 and IPCC Tier 1)')
        r += 1
        hdr = ['Category', 'Zone', 'AGB (t d.m./ha)', 'U AGB', 'BGB (t d.m./ha)', 'U BGB', 'Litter (tC/ha)',
               'Litter upper bound (tC/ha)', 'Lying DW (t d.m./ha)', 'Stumps (t d.m./ha)', 'Standing DW (t d.m./ha)',
               'BGB standing DW (t d.m./ha)', 'U lying DW', 'U stumps', 'U standing DW', 'U BGB standing DW',
               'Litter rule', 'Source']
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
            ws.write_string(rr, 17, {'ipcc': 'IPCC Tier 1', 'scaled': 'Proportional to AGB',
                                     'none': 'None'}[o['litter_rule']], st.cell())
            ws.write_string(rr, 18, o.get('source', ''), st.cell())
        a1_ = a0 + len(fr) - 1
        # --- B. forest calculations
        r = a1_ + 3
        self.h2(ws, r, 1, 20, 'B. Forest stocks — carbon by pool and total stock (gross deforestation EF)')
        r += 1
        hdr = ['Category', 'Zone', 'Key', 'C AGB (tC/ha)', 'U', 'C BGB (tC/ha)', 'U', 'C litter (tC/ha)', 'U',
               'Deadwood (t d.m./ha)', 'U deadwood (measurement)', 'C deadwood (tC/ha)', 'U', 'C total (tC/ha)', 'U C total',
               'Stock (tCO₂/ha)', 'U stock', 'Fuel biomass MB (t d.m./ha)', 'U MB']
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
            # MB = AGB + litter (tC/ha / CF litter) + deadwood, in t d.m./ha
            self.formula(ws, rr, 18, f'={a1(a, 3)}+{a1(rr, 8)}/CF_litter+{a1(rr, 10)}', cf(DM2, bold=True), o['mb'])
            self.formula(ws, rr, 19, f'=IF({a1(rr, 18)}>0,SQRT(({a1(a, 4)}*{a1(a, 3)})^2+({a1(rr, 9)}*{a1(rr, 8)}/CF_litter)^2'
                                     f'+({a1(rr, 11)}*{a1(rr, 10)})^2)/{a1(rr, 18)},0)', cf(PCT1), o['u_mb'])
        b1 = b0 + len(fr) - 1
        self.stk.update(f_key=rng(S_STK, b0, 3, b1, 3), f_ef=rng(S_STK, b0, 16, b1, 16), f_u=rng(S_STK, b0, 17, b1, 17),
                        f_mb=rng(S_STK, b0, 18, b1, 18), f_umb=rng(S_STK, b0, 19, b1, 19),
                        f_rows={(o['category'], o['zone']): b0 + i for i, o in enumerate(fr)})
        # --- C. post-conversion stocks
        pr = list(m.post_stocks.values())
        r = b1 + 3
        self.h2(ws, r, 1, 20, 'C. Stocks of the post-deforestation land use (above- and below-ground biomass)')
        r += 1
        hdr = ['Category', 'Rule', 'AGB (t d.m./ha)', 'U AGB', 'BGB (t d.m./ha)', 'U BGB', 'C AGB (tC/ha)', 'U', 'C BGB (tC/ha)',
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
                ws.write_string(rr, 2, ('Mean of ' if len(srcs) > 1 else 'Same as ') + ' and '.join(o['mean_of']),
                                cf(None))
                for c_v, c_u, v, u in [(3, 4, 'agb', 'u_agb'), (5, 6, 'bgb', 'u_bgb')]:
                    self.formula(ws, rr, c_v, '=AVERAGE(' + ','.join(a1(x, c_v) for x in srcs) + ')', cf(DM2), o[v])
                    self.formula(ws, rr, c_u, '=IFERROR(SQRT(' + '+'.join(f'({a1(x, c_v)}*{a1(x, c_u)})^2' for x in srcs)
                                 + ')/(' + '+'.join(a1(x, c_v) for x in srcs) + '),0)', cf(PCT1), o[u])
            else:
                ws.write_string(rr, 2, 'Input' if o['rule'] == 'input' else 'No biomass (0)', cf(None))
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
        # --- D. SOC factors
        sr = list(m.soc_factors.values())
        r = c1 + 3
        self.h2(ws, r, 1, 20, 'D. Soil carbon change factors (post-deforestation land use; forest = 1)')
        r += 1
        hdr = ['Land use', 'F_LU', 'U F_LU', 'F_MG', 'U F_MG', 'F_I', 'U F_I', 'F = F_LU·F_MG·F_I', 'U F', 'Source']
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
        # --- E. growth rates
        gr = list(m.removal_rates.values())
        r = d1 + 3
        self.h2(ws, r, 1, 20, 'E. Growth rates of the gains (non-forest → forest)')
        r += 1
        hdr = ['Gain category', 'Zone', 'Key', 'AGB growth (tC/ha/yr)', 'U', 'R (root/shoot)',
               'G (tCO₂/ha/yr)', 'U G', 'Source']
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
        self.note(ws, r, 1, 19, 'A-B: C AGB = AGB·CF; C BGB = BGB·CF; litter in tC/ha (IPCC Tier 1, or proportional to '
                  'AGB for secondary forest); C deadwood = Σ components·CF; U of each pool combined with U_CF; '
                  'U of the total = √Σ(U·C)² / C. MB (fuel biomass) = AGB + litter / CF litter + deadwood, in t d.m./ha, '
                  'without roots (IPCC 2006 eq. 2.27); U MB = √Σ(U·x)² / MB. D: F = 1 for forest; '
                  'U F = √(U_LU² + U_MG² + U_I²) (product). E: G = − AGB growth · (1 + R) · 44/12; R = 0: above-ground '
                  'growth only.', 3)
        ws.freeze_panes(0, 3)

    # ============================================================ activity sheets
    def sheet_activity(self, name, title):
        st, L, m, H = self.st, self.L, self.m, self.H
        A = self.act[name]
        rows, labels, c = A['rows'], A['labels'], A['cols']
        lab_hdr = {'forest': 'Forest 2021', 'post': 'Post-deforestation land use', 'f_i': 'Forest 2021',
                   'f_f': 'Forest 2025', 'g_i': 'Category 2021', 'gain': 'Gain category 2025'}
        widths = [24] + [26] * len(labels) + [8] + [8] * H + [8, 11, 14, 13, 11, 13, 14]
        ws = self.sheet(name, title + ' (stratified estimator by zone)', widths)
        r0 = A['first'] - 1
        strata_hdr = st.fmt(bold=True, font_color='white', bg_color=GREEN, align='center', border=1, border_color='white',
                            num_format='"n  h="0')
        tail = ['n', 'Proportion', 'Period area (ha)', 'Standard error (ha)', self.ci_label('Uncertainty', ' (%)'),
                '± (ha)', 'Annual area (ha/yr)']

        def header(r, with_labels):
            ws.set_row(r, 40)
            self.label(ws, r, 1, 'Zone', st.hdr_l)
            for j, lb in enumerate(labels):
                ws.write(r, 2 + j, lab_hdr[lb] if with_labels else '', st.hdr_l)
            ws.write(r, c['key'], 'Key' if with_labels else '', st.hdr)
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
            self.formula(ws, r, c['ay'], f'={a1(r, c["area"])}/Period_length', cf(HA0), v['area_year'])

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
        ws.write(rt, 1, 'Total (all transitions)', st.cell(bold=True, top=True))
        for cc in range(2, c['ay'] + 1):
            ws.write_blank(rt, cc, None, st.cell(top=True))
        for cc, v, nf in [(c['n'], rows.n.sum(), T0), (c['area'], rows.area.sum(), HA0), (c['ay'], rows.area_year.sum(), HA0)]:
            self.formula(ws, rt, cc, f'=SUM({a1(A["first"], cc)}:{a1(last, cc)})', st.cell(nf, bold=True, top=True), v)
        n_col = xl_col_to_name(c['n'])
        ws.conditional_format(A['first'], 1, last, c['ay'], {
            'type': 'formula', 'criteria': f'=${n_col}{A["first"] + 1}=0', 'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.conditional_format(A['first'], c['n'], last, c['n'], {
            'type': 'formula', 'criteria': f'=AND(${n_col}{A["first"] + 1}>0,${n_col}{A["first"] + 1}<Small_n_threshold)',
            'format': self.wb.add_format({'bg_color': '#FDE7C8'})})
        ws.autofilter(r0, 1, last, c['ay'])
        ws.freeze_panes(A['first'], c['key'] + 1)

        # total area by zone: direct estimate of the activity
        code = A['code']
        self.h2(ws, A['zd_h2'], 1, c['ay'], f'Total area by zone — direct estimate of activity "{code}"')
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
        ws.write(r, 1, self.lab['Adj'] + ' (sum of zones)', st.cell(bold=True, top=True))
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
        self.formula(ws, r, c['ay'], f'={a1(r, c["area"])}/Period_length', tf(HA0), nat['area_year'])
        self.note(ws, r + 2, 1, c['ay'], 'Each transition is a domain of the zone: p̂ = Σ_h W_hz · n_hkz / n_hz (Cochran 1977). '
                  'Greyed rows: no sample; in orange: fewer than "Small n threshold" samples. The sum of the transitions '
                  'of a zone equals the direct estimate if all samples are assigned (see Checks).', 2)

    def act_cell(self, name, i, col):
        A = self.act[name]
        return ref(name, A['first'] + i, A['cols'][col])

    # ============================================================ emissions: shared utilities
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
        """specs: list of (col, kind, value); kind = sum | u:(col_val, col_ux2) | pm:(col_val, col_u)."""
        label = label or self.lab['Adj'] + ' total'
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

    # ============================================================ Emissions DEF
    def sheet_emissions_def(self):
        st, m = self.st, self.m
        rows = m.def_rows
        groups = [(1, 4, 'Transition'), (5, 6, 'Activity data'), (7, 8, 'Gross EF (forest stock)'),
                  (9, 10, 'Post-deforestation stock'), (11, 12, 'Net EF'), (13, 16, 'Gross emissions (period)'),
                  (17, 20, 'Net emissions (period)'), (21, 22, 'Annual average')]
        headers = ['Zone', 'Forest 2021', 'Post-deforestation land use', 'Key', 'Area (ha)', 'U (%)',
                   'tCO₂/ha', 'U (%)', 'tCO₂/ha', 'U (%)', 'tCO₂/ha', 'U (%)', 'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²',
                   'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²', 'Gross (tCO₂/yr)', 'Net (tCO₂/yr)']
        ws = self.emission_sheet(S_EM, 'Deforestation — emissions from biomass and DOM',
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
            self.formula(ws, r, 8, f'=INDEX({self.stk["f_u"]},MATCH({fkey},{self.stk["f_key"]},0))*U_EF_factor', cf(PCT1),
                         row.u_ef_gross)
            self.formula(ws, r, 9, f'=INDEX({self.stk["p_ef"]},MATCH({R(3)},{self.stk["p_key"]},0))', cf(EF2), row.stock_post)
            self.formula(ws, r, 10, f'=INDEX({self.stk["p_u"]},MATCH({R(3)},{self.stk["p_key"]},0))*U_EF_factor', cf(PCT1),
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
            self.formula(ws, r, 21, f'={R(13)}/Period_length', cf(T0), row.em_gross_year)
            self.formula(ws, r, 22, f'={R(17)}/Period_length', cf(T0, bold=True), row.em_net_year)
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
        self.wb.define_name('Transition_keys', '=' + rng(S_EM, first, 4, last, 4))
        self.wb.define_name('EF_net_transitions', '=' + rng(S_EM, first, 11, last, 11))
        self.note(ws, rt + 2, 1, 22,
                  'Gross EF = forest stock (category, zone); net EF = gross EF − post-deforestation stock. The stock '
                  'uncertainties are converted to the confidence level of the results (× U_EF_factor). Emissions = area × EF; '
                  'U = √(U_AD² + U_EF²); total: U = √Σ(U·x)² / |Σx| (IPCC 2006 vol. 1 eq. 3.1 and 3.2). Immediate emissions: '
                  'annual average = total / T.', 2)

    # ============================================================ Emissions DEF SOC
    def sheet_soc(self):
        st, m = self.st, self.m
        rows = m.soc_rows
        groups = [(1, 3, 'Transition'), (4, 5, 'Activity data'), (6, 6, 'Soil'), (7, 8, 'Factor F'),
                  (9, 11, 'Soil stock change'), (12, 15, 'Emissions (period)'), (16, 16, 'Annual')]
        headers = ['Zone', 'Forest 2021', 'Post-deforestation land use', 'Area (ha)', 'U (%)',
                   'SOC_ref (tC/ha)', 'F', 'U (%)', 'ΔSOC (tC/ha)', 'U (%)', 'Rate (tCO₂/ha/yr)', 'tCO₂', 'U (%)',
                   '± (tCO₂)', '(U·x)²', 'tCO₂/yr']
        ws = self.emission_sheet(S_SOC, 'Deforestation — emissions from mineral soil organic carbon',
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
            self.formula(ws, r, 8, f'=INDEX({self.stk["s_u"]},MATCH({R(3)},{self.stk["s_key"]},0))*U_EF_factor', cf(PCT1),
                         row.u_f)
            self.formula(ws, r, 9, f'={R(6)}*({R(7)}-1)', cf(EF2), row.dsoc)
            self.formula(ws, r, 10, f'=IF({R(7)}<>1,SQRT((U_SOC_ref*U_EF_factor)^2+({R(8)}*{R(7)}/ABS({R(7)}-1))^2),0)',
                         cf(PCT1), row.u_dsoc)
            self.formula(ws, r, 11, f'=-{R(9)}/SOC_D*CO2_C', cf(EF2), row.rate)
            self.formula(ws, r, 12, f'={R(4)}*{R(11)}*Cohort_factor', cf(T0, bold=True), row.em)
            self.formula(ws, r, 13, f'=IF({R(4)}>0,SQRT({R(5)}^2+{R(10)}^2),0)', cf(PCT1), row.u_em)
            self.formula(ws, r, 14, f'=ABS({R(12)})*{R(13)}', cf(T0), row.pm)
            self.formula(ws, r, 15, f'=({R(13)}*{R(12)})^2', cf(SCI), row.ux2)
            self.formula(ws, r, 16, f'={R(12)}/Period_length', cf(T0, bold=True), row.em_year)
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
                  'IPCC 2006 vol. 4 eq. 2.25: ΔSOC = SOC_ref · (F − 1), with F = F_LU·F_MG·F_I of the final land use and '
                  'F = 1 for forest; the change is spread over D years. Each deforested hectare emits −ΔSOC/D · 44/12 per year. '
                  'With a constant annual area (A/T), the cohorts accumulate: year t emits (A/T)·t·rate, i.e. '
                  'A · rate · (T+1)/2 over the period. U = √(U_AD² + U_ΔSOC²); U_ΔSOC combines U_SOC_ref and the uncertainty of '
                  '(F − 1).', 3)

    # ============================================================ Emissions DEF CH4 N2O (biomass burning)
    def sheet_fire(self):
        st, m = self.st, self.m
        rows = m.fire_rows
        groups = [(1, 3, 'Transition'), (4, 5, 'Activity data'), (6, 7, 'Fire use'),
                  (8, 10, 'Burnt biomass'), (11, 17, 'CH₄'), (18, 24, 'N₂O')]
        gas_hdr = ['Gef (g/kg)', 't gas', 'tCO₂e', 'U (%)', '± (tCO₂e)', '(U·x)²', 'tCO₂e/yr']
        headers = ['Zone', 'Forest 2021', 'Post-deforestation land use', 'Area (ha)', 'U (%)', 'Burnt fraction',
                   'Burnt area (ha)', 'MB (t d.m./ha)', 'U (%)', 'Cf'] + gas_hdr + gas_hdr
        ws = self.emission_sheet(S_FIRE, 'Deforestation — CH₄ and N₂O from biomass burning',
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
            self.formula(ws, r, 6, f'=IFERROR(INDEX(Fire_fraction,MATCH({R(2)}&"|"&{R(3)},Fire_key,0)),{FIRE_DEFAULT})',
                         cf('0.00'), row.frac)
            self.formula(ws, r, 7, f'={R(4)}*{R(6)}', cf(HA0), row.burnt)
            fkey = f'{R(2)}&"|"&{R(1)}'
            self.formula(ws, r, 8, f'=INDEX({self.stk["f_mb"]},MATCH({fkey},{self.stk["f_key"]},0))', cf(DM2), row.mb)
            self.formula(ws, r, 9, f'=INDEX({self.stk["f_umb"]},MATCH({fkey},{self.stk["f_key"]},0))*U_EF_factor',
                         cf(PCT1), row.u_mb)
            self.formula(ws, r, 10, '=Combustion_factor', cf('0.00'), row.comb)
            for c0, (g, _), gef, ugef, gwp in [(11, GASES[0], 'Gef_CH4', 'U_Gef_CH4', 'GWP_CH4'),
                                               (18, GASES[1], 'Gef_N2O', 'U_Gef_N2O', 'GWP_N2O')]:
                self.formula(ws, r, c0, f'={gef}', cf('0.00'), row[f'gef_{g}'])
                self.formula(ws, r, c0 + 1, f'={R(7)}*{R(8)}*{R(10)}*{R(c0)}*0.001', cf(EF2), row[f't_{g}'])
                self.formula(ws, r, c0 + 2, f'={R(c0 + 1)}*{gwp}', cf(T0, bold=True), row[f'em_{g}'])
                self.formula(ws, r, c0 + 3, f'=IF(AND({R(4)}>0,{R(6)}>0),SQRT({R(5)}^2+{R(9)}^2'
                                            f'+(U_Combustion_factor*U_EF_factor)^2+({ugef}*U_EF_factor)^2'
                                            f'+(U_Burnt_fraction*U_EF_factor)^2),0)', cf(PCT1), row[f'u_{g}'])
                self.formula(ws, r, c0 + 4, f'=ABS({R(c0 + 2)})*{R(c0 + 3)}', cf(T0), row[f'pm_{g}'])
                self.formula(ws, r, c0 + 5, f'=({R(c0 + 3)}*{R(c0 + 2)})^2', cf(SCI), row[f'ux2_{g}'])
                self.formula(ws, r, c0 + 6, f'={R(c0 + 2)}/Period_length', cf(T0, bold=True), row[f'em_{g}_year'])
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
                  'IPCC 2006 vol. 4 eq. 2.27: L = A · burnt fraction · MB · '
                  'Cf · Gef · 10⁻³ (t of gas), then × GWP (AR5). MB = AGB + litter + deadwood (t d.m./ha, without roots) for '
                  'all transitions; burnt fractions: Mappings. U = √(U_AD² + U_MB² + U_Cf² + U_Gef² + '
                  'U_fraction²); the uncertainties of Cf, Gef and the burnt fraction are to be completed (Parameters). '
                  'Immediate emissions: annual average = total / T.', 3)

    # ============================================================ Emissions DEG
    def sheet_deg(self):
        st, m = self.st, self.m
        rows = m.deg_rows
        groups = [(1, 3, 'Transition'), (4, 5, 'Activity data'), (6, 7, 'Forest stock 2021'),
                  (8, 9, 'Forest stock 2025'), (10, 11, 'Degradation EF'), (12, 15, 'Emissions (period)'),
                  (16, 16, 'Annual')]
        headers = ['Zone', 'Forest 2021', 'Forest 2025', 'Area (ha)', 'U (%)', 'tCO₂/ha', 'U (%)', 'tCO₂/ha', 'U (%)',
                   'tCO₂/ha', 'U (%)', 'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²', 'tCO₂/yr']
        ws = self.emission_sheet(S_DEG, 'Degradation — emissions',
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
                self.formula(ws, r, c0 + 1, f'=INDEX({self.stk["f_u"]},MATCH({k},{self.stk["f_key"]},0))*U_EF_factor',
                             cf(PCT1), row[u])
            self.formula(ws, r, 10, f'={R(6)}-{R(8)}', cf(EF2, bold=True), row.ef)
            self.formula(ws, r, 11, f'=IF({R(10)}<>0,SQRT(({R(7)}*{R(6)})^2+({R(9)}*{R(8)})^2)/ABS({R(10)}),0)', cf(PCT1),
                         row.u_ef)
            self.formula(ws, r, 12, f'={R(4)}*{R(10)}', cf(T0, bold=True), row.em)
            self.formula(ws, r, 13, f'=IF({R(4)}>0,SQRT({R(5)}^2+{R(11)}^2),0)', cf(PCT1), row.u_em)
            self.formula(ws, r, 14, f'=ABS({R(12)})*{R(13)}', cf(T0), row.pm)
            self.formula(ws, r, 15, f'=({R(13)}*{R(12)})^2', cf(SCI), row.ux2)
            self.formula(ws, r, 16, f'={R(12)}/Period_length', cf(T0, bold=True), row.em_year)
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
                  'Degradation EF = stock (forest 2021, zone) − stock (forest 2025, zone), biomass, litter and deadwood '
                  'pools; U = √[(U_i·S_i)² + (U_f·S_f)²] / |S_i − S_f|. Immediate emissions: '
                  'annual average = total / T. Identical categories in 2021 and 2025 have a zero EF.', 2)

    # ============================================================ Removals GAIN
    def sheet_enh(self):
        st, m = self.st, self.m
        rows = m.gain_rows
        groups = [(1, 3, 'Gain'), (4, 5, 'Activity data'), (6, 10, 'Removal factor'),
                  (11, 14, 'Removals (period)'), (15, 15, 'Annual')]
        headers = ['Zone', 'Category 2021', 'Gain category 2025', 'Area (ha)', 'U (%)', 'AGB growth (tC/ha/yr)', 'U (%)', 'R',
                   'G (tCO₂/ha/yr)', 'U (%)', 'tCO₂', 'U (%)', '± (tCO₂)', '(U·x)²', 'tCO₂/yr']
        ws = self.emission_sheet(S_ENH, 'Enhancement of forest carbon stocks — removals',
                                 [22, 26, 26, 13, 9, 13, 9, 8, 12, 9, 14, 9, 13, 10, 14], groups, headers)
        first = 6
        per_zone = len(rows) // self.Z
        for i, row in rows.iterrows():
            r = first + i
            band, top = row.zi % 2 == 1, i % per_zone == 0
            cf = lambda nf, **k: st.cell(nf, band=band, top=top, **k)
            R = lambda c: a1(r, c)
            self.formula(ws, r, 1, '=' + self.act_cell(S_AGAIN, i, 'zone'), cf(None, bold=top), row.zone)
            self.formula(ws, r, 2, '=' + ref(S_AGAIN, self.act[S_AGAIN]['first'] + i, 2), cf(None), row.g_i)
            self.formula(ws, r, 3, '=' + ref(S_AGAIN, self.act[S_AGAIN]['first'] + i, 3), cf(None), row.gain)
            self.formula(ws, r, 4, '=' + self.act_cell(S_AGAIN, i, 'area'), cf(HA0), row.area)
            self.formula(ws, r, 5, '=' + self.act_cell(S_AGAIN, i, 'u'), cf(PCT1), row.u_area)
            k = f'{R(3)}&"|"&{R(1)}'
            self.formula(ws, r, 6, f'=INDEX({self.stk["g_rate"]},MATCH({k},{self.stk["g_key"]},0))', cf('0.000'), row.rate)
            self.formula(ws, r, 7, f'=INDEX({self.stk["g_u"]},MATCH({k},{self.stk["g_key"]},0))*U_EF_factor', cf(PCT1),
                         row.u_rate)
            self.formula(ws, r, 8, f'=INDEX({self.stk["g_r"]},MATCH({k},{self.stk["g_key"]},0))', cf('0.000'), row.r)
            self.formula(ws, r, 9, f'=-{R(6)}*(1+{R(8)})*CO2_C', cf(EF2, bold=True), row.g)
            self.formula(ws, r, 10, f'={R(7)}', cf(PCT1), row.u_rate)
            self.formula(ws, r, 11, f'={R(4)}*{R(9)}*Cohort_factor', cf(T0, bold=True), row.em)
            self.formula(ws, r, 12, f'=IF({R(4)}>0,SQRT({R(5)}^2+{R(10)}^2),0)', cf(PCT1), row.u_em)
            self.formula(ws, r, 13, f'=ABS({R(11)})*{R(12)}', cf(T0), row.pm)
            self.formula(ws, r, 14, f'=({R(12)}*{R(11)})^2', cf(SCI), row.ux2)
            self.formula(ws, r, 15, f'={R(11)}/Period_length', cf(T0, bold=True), row.em_year)
        last = first + len(rows) - 1
        t, u = rows.em.sum(), u_ratio(rows.ux2.sum(), rows.em.sum())
        rt = self.emission_total(ws, first, last, 15, [
            (4, 'sum', HA0, rows.area.sum()), (11, 'sum', T0, t), (14, 'sum', SCI, rows.ux2.sum()),
            (12, ('u', 11, 14), PCT1, u), (13, ('pm', 11, 12), T0, abs(t) * u), (15, 'sum', T0, rows.em_year.sum())])
        self.em[S_ENH] = dict(first=first, last=last, total=rt, val=11, ux2=14, year=15)
        ws.conditional_format(first, 1, last, 15, {'type': 'formula', 'criteria': f'=$D{first + 1}=0',
                                                   'format': self.wb.add_format({'font_color': '#B0B7B3'})})
        ws.autofilter(5, 1, last, 15)
        ws.freeze_panes(first, 4)
        self.note(ws, rt + 2, 1, 15,
                  'Gains broken down by zone, 2021 category (non-forest land use, or forest category for a gain to a '
                  'higher-ranked forest) and 2025 gain category; the growth rate is that of the 2025 gain category '
                  'in the zone. '
                  'G = − AGB growth · (1 + R) · 44/12 (negative value = removal). With a constant annual gain area '
                  '(A/T), the cohorts accumulate: year t removes (A/T)·t·G, i.e. A · G · (T+1)/2 over the '
                  'period. U = √(U_AD² + U_G²).', 2)

    # ============================================================ Balance
    def sheet_balance(self):
        st, m, b = self.st, self.m, self.bil
        ws = self.sheet(S_BIL, 'Balance of emissions and removals', [48, 18, 13, 16, 16, 16, 16, 13, 16, 16, 16, 13, 16])
        bil = m.balance
        r = b['first'] - 2
        self.h2(ws, r, 1, 7, 'A. ' + self.lab['Adj'] + ' annual averages by component')
        self.hdr_row(ws, r + 1, 1, ['Component', 'tCO₂e/yr', 'Uncertainty (%)', '± (tCO₂e/yr)', 'Share of emissions',
                                    'Period total (tCO₂e)'])
        src = {'def': S_EM, 'soc': S_SOC, 'ch4': 'ch4', 'n2o': 'n2o', 'deg': S_DEG, 'enh': S_ENH}   # keys of self.em
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
                parts = EMISSIONS if code == 'tot' else ['tot', 'enh']
                self.formula(ws, rr, 2, '=' + '+'.join(a1(R[p], 2) for p in parts), cf(T0), v['year'])
                self.formula(ws, rr, 3, f'=IF({a1(rr, 2)}<>0,SQRT(' + '+'.join(f'{a1(R[p], 4)}^2' for p in parts)
                             + f')/ABS({a1(rr, 2)}),0)', cf(PCT1), v['u'])
            self.formula(ws, rr, 4, f'=ABS({a1(rr, 2)})*{a1(rr, 3)}', cf(T0), v['pm'])
            if code in EMISSIONS + ['tot']:
                self.formula(ws, rr, 5, f'=IF({a1(R["tot"], 2)}<>0,{a1(rr, 2)}/{a1(R["tot"], 2)},0)', cf(PCT1),
                             v['year'] / bil['tot']['year'] if bil['tot']['year'] else 0)
            else:
                ws.write_blank(rr, 5, None, cf(None))
            self.formula(ws, rr, 6, f'={a1(rr, 2)}*Period_length', cf(T0), v['year'] * m.T)

        # areas
        r = b['area0'] - 2
        self.h2(ws, r, 1, 7, 'B. ' + self.lab['Adj'] + ' annual areas (direct estimate)')
        self.hdr_row(ws, r + 1, 1, ['Activity', 'ha/yr', 'Uncertainty (%)', '± (ha/yr)', '', 'Period total (ha)'])
        for i, (code, name, sheet) in enumerate([('Def', 'Deforestation', S_ADEF), ('Deg', 'Degradation', S_ADEG),
                                                 ('Gain', 'Gain (non-forest → forest)', S_AGAIN)]):
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

        # yearly series
        r = b['year_h2']
        self.h2(ws, r, 1, 13, 'C. Yearly series (tCO₂e/yr)')
        hdr = ['Year', 'DEF biomass and DOM', 'DEF mineral soils', 'DEF burning CH₄', 'DEF burning N₂O',
               'Degradation', 'Total emissions', 'U (%)', '± (tCO₂e)', 'Enhancement (removals)', 'Net balance', 'U (%)',
               '± (tCO₂e)']
        self.hdr_row(ws, r + 1, 1, hdr, height=36)
        ycol = {'def': 2, 'soc': 3, 'ch4': 4, 'n2o': 5, 'deg': 6, 'tot': 7, 'enh': 10, 'net': 11}
        self.bil['ycol'] = ycol
        cohort = ('soc', 'enh')        # cohorts accumulate; the other components are constant
        for i, y in m.years.iterrows():
            rr, band = b['year0'] + i, i % 2 == 1
            cf = lambda nf, **k: st.cell(nf, band=band, **k)
            self.formula(ws, rr, 1, f'=Initial_year+{i + 1}', cf('0', bold=True), int(y.year))
            t = f'({a1(rr, 1)}-Initial_year)'
            for k in EMISSIONS + ['enh']:
                f = '=' + xl_rowcol_to_cell(R[k], 2, True, True) + (f'*2*{t}/(Period_length+1)' if k in cohort else '')
                self.formula(ws, rr, ycol[k], f, cf(T0), y[k])
            ct = ycol['tot']
            self.formula(ws, rr, ct, f'=SUM({a1(rr, ycol["def"])}:{a1(rr, ycol["deg"])})', cf(T0, bold=True), y['tot'])
            pm = '+'.join(f'(ABS({a1(rr, ycol[k])})*{xl_rowcol_to_cell(R[k], 3, True, True)})^2' for k in EMISSIONS)
            self.formula(ws, rr, ct + 2, f'=SQRT({pm})', cf(T0), y['pm_tot'])
            self.formula(ws, rr, ct + 1, f'=IF({a1(rr, ct)}<>0,{a1(rr, ct + 2)}/ABS({a1(rr, ct)}),0)', cf(PCT1),
                         y['pm_tot'] / abs(y['tot']) if y['tot'] else 0)
            cr, cn = ycol['enh'], ycol['net']
            self.formula(ws, rr, cn, f'={a1(rr, ct)}+{a1(rr, cr)}', cf(T0, bold=True), y['net'])
            self.formula(ws, rr, cn + 2, f'=SQRT({a1(rr, ct + 2)}^2+(ABS({a1(rr, cr)})*'
                                         f'{xl_rowcol_to_cell(R["enh"], 3, True, True)})^2)', cf(T0), y['pm_net'])
            self.formula(ws, rr, cn + 1, f'=IF({a1(rr, cn)}<>0,{a1(rr, cn + 2)}/ABS({a1(rr, cn)}),0)', cf(PCT1),
                         y['pm_net'] / abs(y['net']) if y['net'] else 0)
        rr = b['year0'] + m.T
        ws.write(rr, 1, 'Average', st.cell(bold=True, top=True))
        means = {c: k for k, c in ycol.items()}
        for c in range(2, 14):
            if c not in means:
                ws.write_blank(rr, c, None, st.cell(top=True))
            else:
                v = m.years[means[c]].mean()
                self.formula(ws, rr, c, f'=AVERAGE({a1(b["year0"], c)}:{a1(rr - 1, c)})', st.cell(T0, bold=True, top=True), v)
        self.note(ws, rr + 2, 1, 13,
                  'Positive values: emissions; negative values: removals. Components assumed independent: '
                  'U = √Σ(U·x)² / |Σx| (IPCC 2006 vol. 1 eq. 3.2). Deforestation (biomass and burning CH₄, N₂O) and '
                  'degradation: immediate emissions, constant every year. Soils and enhancement: '
                  'the cohorts accumulate, year t equals average · 2t / (T+1).', 3)

    # ============================================================ Results
    def sheet_results(self):
        st, m, Z = self.st, self.m, self.Z
        ws = self.sheet(S_RES, 'Results by zone and category',
                        [28] + [14, 10, 13] * 7)
        r = 4
        # A. areas by zone
        self.h2(ws, r, 1, 10, 'A. Annual areas by zone (ha/yr, direct estimate)')
        r += 1
        ws.merge_range(r, 1, r + 1, 1, 'Zone', st.hdr_l)
        for j, (name, sheet) in enumerate([('Deforestation', S_ADEF), ('Degradation', S_ADEG), ('Gain', S_AGAIN)]):
            c0 = 2 + 3 * j
            ws.merge_range(r, c0, r, c0 + 2, name, st.hdr)
            for k, h in enumerate(['ha/yr', 'U (%)', '± (ha/yr)']):
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

        # B. emissions and removals by zone
        comps = [('def', 'DEF biomass and DOM', S_EM), ('soc', 'DEF mineral soils', S_SOC),
                 ('ch4', 'DEF burning CH₄', 'ch4'), ('n2o', 'DEF burning N₂O', 'n2o'), ('deg', 'Degradation', S_DEG),
                 ('enh', 'Enhancement (removals)', S_ENH), ('net', 'Net balance', None)]
        self.h2(ws, r, 1, 1 + 3 * len(comps), 'B. Annual emissions and removals by zone (tCO₂e/yr)')
        r += 1
        ws.merge_range(r, 1, r + 1, 1, 'Zone', st.hdr_l)
        for j, (_, name, _) in enumerate(comps):
            c0 = 2 + 3 * j
            ws.merge_range(r, c0, r, c0 + 2, name, st.hdr)
            for k, h in enumerate(['tCO₂e/yr', 'U (%)', '± (tCO₂e/yr)']):
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
            vals = m.balance if nat else m.balance_zone[zn]
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
                    self.formula(ws, rr, c0, f'=SUM({vc})/Period_length', cf(T0), v['year'])
                    self.formula(ws, rr, c0 + 1, f'=IFERROR(SQRT(SUM({uc}))/ABS(SUM({vc})),0)', cf(PCT1), v['u'])
                else:
                    key = a1(rr, 1)
                    self.formula(ws, rr, c0, f'=SUMIFS({vc},{zc},{key})/Period_length', cf(T0), v['year'])
                    self.formula(ws, rr, c0 + 1, f'=IFERROR(SQRT(SUMIFS({uc},{zc},{key}))/ABS(SUMIFS({vc},{zc},{key})),0)',
                                 cf(PCT1), v['u'])
                self.formula(ws, rr, c0 + 2, f'=ABS({a1(rr, c0)})*{a1(rr, c0 + 1)}', cf(T0), v['pm'])
        self.res_em_rows = (first, first + Z)
        r = first + Z + 3

        # C and D. deforestation (biomass) by forest category and by land use
        e = self.em[S_EM]
        rows = m.def_rows
        col = lambda c: self.col(S_EM, c, e['first'], e['last'])
        hdr = ['', 'Area (ha/yr)', 'U (%)', '± (ha/yr)', 'Gross emissions (tCO₂/yr)', 'U (%)', '± (tCO₂/yr)',
               'Net emissions (tCO₂/yr)', 'U (%)', '± (tCO₂/yr)', 'Share of net emissions']
        self.res_cat_rows = {}
        for title, by, field, groups in [
                (f'C. Deforestation (biomass and DOM) by forest category in {self.p["initial_year"]}', 2, 'forest',
                 m.forest_cats),
                ('D. Deforestation (biomass and DOM) by post-deforestation land use', 3, 'post', m.post_cats)]:
            self.h2(ws, r, 1, 11, title)
            h = list(hdr)
            h[0] = 'Category'
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
                self.formula(ws, rr, 2, f'=SUMIFS({col(5)},{col(by)},{key})/Period_length', cf(HA0), va)
                self.formula(ws, rr, 3, f'=IFERROR(SQRT(SUMPRODUCT(({col(by)}={key})*({col(6)}*{col(5)})^2))'
                                        f'/SUMIFS({col(5)},{col(by)},{key}),0)', cf(PCT1), ua)
                self.formula(ws, rr, 4, f'={a1(rr, 2)}*{a1(rr, 3)}', cf(HA0), va * ua)
                for c_val, c_em, c_ux2, s_em, s_ux2 in [(5, 13, 16, 'em_gross', 'ux2_gross'), (8, 17, 20, 'em_net', 'ux2_net')]:
                    tot = sub[s_em].sum()
                    uu = u_ratio(sub[s_ux2].sum(), tot)
                    self.formula(ws, rr, c_val, f'=SUMIFS({col(c_em)},{col(by)},{key})/Period_length', cf(T0), tot / m.T)
                    self.formula(ws, rr, c_val + 1, f'=IFERROR(SQRT(SUMIFS({col(c_ux2)},{col(by)},{key}))'
                                                    f'/ABS(SUMIFS({col(c_em)},{col(by)},{key})),0)', cf(PCT1), uu)
                    self.formula(ws, rr, c_val + 2, f'=ABS({a1(rr, c_val)})*{a1(rr, c_val + 1)}', cf(T0), abs(tot / m.T) * uu)
                self.formula(ws, rr, 11, f'=IFERROR({a1(rr, 8)}/{ref(S_EM, e["total"], 22)},0)', cf(PCT1),
                             sub.em_net.sum() / rows.em_net.sum() if rows.em_net.sum() else 0)
            self.res_cat_rows[field] = (first, first + len(groups))
            r = first + len(groups) + 2
        self.note(ws, r, 1, 16,
                  'Table A: direct estimate of each activity by zone; ' + self.lab['adj'] + ' = sum of the zones, SE = √Σ SE_z². '
                  'Tables B to D: sums of the rows of the emission sheets, U = √Σ(U·x)² / |Σx| (IPCC approach 1). '
                  'The net balance combines the six components in quadrature.', 2)

    # ============================================================ Checks
    def sheet_checks(self):
        st, L, m, H, Z = self.st, self.L, self.m, self.H, self.Z
        ws = self.sheet(S_CTL, 'Consistency checks', [46, 15, 15, 15, 13, 15, 15, 13, 15, 15, 13, 14, 15, 14, 14])
        ok_fmt = self.wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#166534', 'bold': True})
        ko_fmt = self.wb.add_format({'bg_color': '#F8CBAD', 'font_color': '#B91C1C', 'bold': True})

        def status(r1, r2, c):
            ws.conditional_format(r1, c, r2, c, {'type': 'cell', 'criteria': '==', 'value': '"OK"', 'format': ok_fmt})
            ws.conditional_format(r1, c, r2, c, {'type': 'cell', 'criteria': '==', 'value': '"CHECK"', 'format': ko_fmt})

        s = m.samples
        r = 4
        self.h2(ws, r, 1, 5, '1. Samples and mappings')
        self.hdr_row(ws, r + 1, 1, ['Check', 'Value', 'Expected', 'Status'])
        rr = r + 2
        self.label(ws, rr, 1, 'Valid samples', st.cell(bold=True))
        self.formula(ws, rr, 2, f"=COUNTA('{S_ECH}'!A:A)-1", st.cell(INT), len(s))
        ws.write_blank(rr, 3, None, st.cell())
        ws.write_blank(rr, 4, None, st.cell())
        rr += 1
        self.label(ws, rr, 1, 'Samples without zone or stratum', st.cell(bold=True))
        v = len(s) - m.n_zh.sum()
        self.formula(ws, rr, 2, f"=COUNTA('{S_ECH}'!A:A)-1-{ref(S_PAR, L['n_tot'], 2 + H)}", st.cell(INT), v)
        ws.write_number(rr, 3, 0, st.cell(INT))
        self.formula(ws, rr, 4, f'=IF({a1(rr, 2)}={a1(rr, 3)},"OK","CHECK")', st.cell(align='center'),
                     'OK' if v == 0 else 'CHECK')
        for code, sheet in [('Def', S_ADEF), ('Deg', S_ADEG), ('Gain', S_AGAIN)]:
            A = self.act[sheet]
            n_code = int((s.activity == code).sum())
            n_tr = A['rows'].n.sum()
            rr += 1
            self.label(ws, rr, 1, f'{code} samples', st.cell(bold=True))
            self.formula(ws, rr, 2, f"=COUNTIF('{S_ECH}'!D:D,\"{code}\")", st.cell(INT), n_code)
            ws.write_blank(rr, 3, None, st.cell())
            ws.write_blank(rr, 4, None, st.cell())
            rr += 1
            self.label(ws, rr, 1, f'{code} without a valid transition (unassigned class or unknown category)',
                       st.cell(bold=True))
            self.formula(ws, rr, 2, f"={a1(rr - 1, 2)}-{ref(sheet, A['total'], A['cols']['n'])}", st.cell(INT), n_code - n_tr)
            ws.write_number(rr, 3, 0, st.cell(INT))
            self.formula(ws, rr, 4, f'=IF({a1(rr, 2)}={a1(rr, 3)},"OK","CHECK")', st.cell(align='center'),
                         'OK' if n_code == n_tr else 'CHECK')
        status(r + 2, rr, 4)
        r = rr + 3

        # 2. sum of the transitions vs direct estimate
        self.h2(ws, r, 1, 11, '2. Area: sum of the transitions vs direct estimate (ha, period)')
        r += 1
        ws.merge_range(r, 1, r + 1, 1, 'Zone', st.hdr_l)
        for j, code in enumerate(['Def', 'Deg', 'Gain']):
            c0 = 2 + 3 * j
            ws.merge_range(r, c0, r, c0 + 2, code, st.hdr)
            for k, h in enumerate(['Σ transitions', 'Direct', 'Difference']):
                ws.write(r + 1, c0 + k, h, st.hdr2)
        ws.merge_range(r, 11, r + 1, 11, 'Status', st.hdr)
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
            self.formula(ws, rr, 11, f'=IF(MAX(ABS({a1(rr, 4)}),ABS({a1(rr, 7)}),ABS({a1(rr, 10)}))<0.001,"OK","CHECK")',
                         st.cell(align='center'), 'OK' if max(diffs) < 1e-3 else 'CHECK')
        status(first, first + Z - 1, 11)
        r = first + Z + 3

        # 3. direct estimator of the net DEF emissions
        self.h2(ws, r, 1, 15, '3. Deforestation (biomass): direct estimator of the net emissions')
        r += 1
        self.note(ws, r, 1, 15, 'Each sample receives y = net EF of its transition (0 if it is not deforested). Ê_z = A_z · Σ_h W_hz · ȳ_hz '
                  'and SE² = A_z² · Σ_h W_hz² · s²_hz / n_hz. The estimate is identical to the sum of the transitions; its '
                  'sampling uncertainty accounts for the covariances between transitions (ignored by approach 1), '
                  'but not for the uncertainty of the EFs.', 2)
        r += 3
        hdr = ['Zone'] + [f'Σy  h={h}' for h in m.strata] + [f'Σy²  h={h}' for h in m.strata] + [
            'Direct Ê (tCO₂, period)', 'SE (tCO₂)', self.ci_label('Sampling U', ' (%)'),
            'Σ transitions (tCO₂)', 'Difference (tCO₂)', 'U AD only, approach 1 (%)']
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

    # ============================================================ Charts
    def sheet_charts(self):
        st, m = self.st, self.m
        ws = self.sheet(S_GR, 'Charts', [14] * 16)
        ws.write(3, 1, 'Charts linked to the "Balance" and "Results" sheets (error bars = ± uncertainty).', st.note1)

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
        ch.add_series({'name': 'Balance', 'categories': [S_BIL, r1, 1, r2, 1], 'values': [S_BIL, r1, 2, r2, 2],
                       'fill': {'color': CHART_COLORS[0]}, 'border': {'none': True}, 'gap': 60,
                       'points': [{'fill': {'color': c}} for c in
                                  [CHART_COLORS[0], CHART_COLORS[1], '#B45309', '#D97706', CHART_COLORS[3], '#374151',
                                   CHART_COLORS[2], '#111827']],
                       'y_error_bars': err(S_BIL, r1, r2, 4)})
        ws.insert_chart(5, 1, style(ch, 'Annual balance by component (tCO₂/yr)', 'tCO₂/yr'))

        y0, y1 = b['year0'], b['year0'] + m.T - 1
        ch = self.wb.add_chart({'type': 'column', 'subtype': 'stacked'})
        yc = b['ycol']
        for k, color in [('def', CHART_COLORS[0]), ('soc', CHART_COLORS[1]), ('ch4', '#B45309'), ('n2o', '#D97706'),
                         ('deg', CHART_COLORS[3]), ('enh', CHART_COLORS[2])]:
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
        ws.insert_chart(5, 8, style(ch, 'Yearly series (tCO₂/yr)', 'tCO₂/yr', legend=True))

        f, _ = self.res_em_rows
        ch = self.wb.add_chart({'type': 'column'})
        ch.add_series({'name': 'Net balance', 'categories': [S_RES, f, 1, f + self.Z - 1, 1],
                       'values': [S_RES, f, self.res_net_col, f + self.Z - 1, self.res_net_col],
                       'fill': {'color': CHART_COLORS[0]}, 'border': {'none': True}, 'gap': 70,
                       'y_error_bars': err(S_RES, f, f + self.Z - 1, self.res_net_col + 2)})
        ws.insert_chart(23, 1, style(ch, 'Annual net balance by zone (tCO₂/yr)', 'tCO₂/yr'))

        f, _ = self.res_area_rows
        ch = self.wb.add_chart({'type': 'column'})
        for j, (name, color) in enumerate([('Deforestation', CHART_COLORS[0]), ('Degradation', CHART_COLORS[3]),
                                           ('Gain', CHART_COLORS[2])]):
            c0 = 2 + 3 * j
            ch.add_series({'name': name, 'categories': [S_RES, f, 1, f + self.Z - 1, 1],
                           'values': [S_RES, f, c0, f + self.Z - 1, c0], 'fill': {'color': color}, 'border': {'none': True},
                           'y_error_bars': err(S_RES, f, f + self.Z - 1, c0 + 2)})
        ws.insert_chart(23, 8, style(ch, 'Annual areas by zone (ha/yr)', 'ha/yr', legend=True))

    # ============================================================ Equations
    def sheet_equations(self):
        st = self.st
        ws = self.sheet(S_EQ, 'Equations', [5, 72, 50, 52])
        F_EQ = st.fmt(border=1, border_color=GRID, indent=1)
        F_EQ_B = st.fmt(border=1, border_color=GRID, indent=1, bg_color=BAND)
        F_NUM = st.fmt(font_color=GREY, align='center', border=1, border_color=GRID)
        F_DESC = st.fmt(text_wrap=True, border=1, border_color=GRID, font_color='#374151')
        F_XL = st.fmt(text_wrap=True, border=1, border_color=GRID, font_color=GREEN, font_name='Consolas', font_size=9)
        sections = [
            ('1 · Activity data — stratified estimator by zone', [
                ('W_{hz} = A_{hz} / A_{z}', 'Weight of stratum h in zone z', 'Parameters: W_hz block'),
                ('p̂_{kz} = Σ_{h} W_{hz} · n_{hkz} / n_{hz}', 'Proportion of transition k (domain) in zone z',
                 'Activity DEF / DEG / GAIN: SUMPRODUCT(W_hz, n_hkz / n_hz)'),
                ('SE(p̂_{kz}) = √[ Σ_{h} W_{hz}^{2} · p̂_{hkz} (1 − p̂_{hkz}) / (n_{hz} − 1) ]', 'Standard error', 'Activity'),
                ('Â_{kz} = A_{z} · p̂_{kz} ,   U_{A} = z · SE(Â_{kz}) / Â_{kz} ,   ± = U_{A} · Â_{kz}',
                 'Area of the period, uncertainty in % and in ha', 'Activity'),
            ]),
            ('2 · Carbon stocks', [
                ('C_{AGB} = AGB · CF ,   C_{BGB} = BGB · CF ,   C_{BM} = Σ BM_{i} · CF', 'Carbon by pool (tC/ha)',
                 'Carbon stocks, table B'),
                ('U_{pool} = √( U_{meas}^{2} + U_{CF}^{2} ) ,   U_{CF} = |CF_{low} − CF| / CF', 'Uncertainty of a pool',
                 'Carbon stocks, table B'),
                ('S = (C_{AGB} + C_{BGB} + C_{L} + C_{BM}) · 44/12 ,   U_{S} = √Σ(U_{i} · C_{i})^{2} / C',
                 'Total stock (tCO₂/ha) and uncertainty', 'Carbon stocks, tables B and C'),
                ('U(level) = U(90%) · z / z_{EF}', 'Conversion of the input uncertainties', 'U_EF_factor'),
            ]),
            ('3 · Deforestation — biomass and DOM', [
                ('EF_{net} = S_{forest} − S_{post} ,   U = √[(U_{f}·S_{f})^{2} + (U_{p}·S_{p})^{2}] / |EF_{net}|',
                 'Net EF and uncertainty of a difference (IPCC 2006 vol. 1 eq. 3.2)', 'Emissions DEF'),
                ('E = Â · EF_{net} ,   U_{E} = √( U_{A}^{2} + U_{EF}^{2} ) ,   E^{yr} = E / T',
                 'Emissions of the transition (eq. 3.1), immediate emissions', 'Emissions DEF'),
            ]),
            ('4 · Deforestation — mineral soils (IPCC 2006 vol. 4 eq. 2.25)', [
                ('F = F_{LU} · F_{MG} · F_{I} ,   ΔSOC = SOC_{ref} · (F − 1)', 'Stock change over D years (forest: F = 1)',
                 'Carbon stocks, table D; Emissions DEF SOC'),
                ('rate = − ΔSOC / D · 44/12', 'Annual emission per deforested hectare (tCO₂/ha/yr)', 'Emissions DEF SOC'),
                ('E = Â · rate · (T + 1) / 2 ,   E_{t} = (Â / T) · t · rate', 'Annual cohorts accumulated over the period',
                 'Cohort_factor; Balance, table C'),
                ('U_{ΔSOC} = √[ U_{SOCref}^{2} + (U_{F} · F / |F − 1|)^{2} ] ,   U_{E} = √( U_{A}^{2} + U_{ΔSOC}^{2} )',
                 'Uncertainty', 'Emissions DEF SOC'),
            ]),
            ('4b · Deforestation — biomass burning (IPCC 2006 vol. 4 eq. 2.27)', [
                ('MB = AGB + L / CF_{litter} + DW', 'Fuel biomass (t d.m./ha), without roots', 'Carbon stocks, '
                 'table B'),
                ('L_{gas} = Â · f_{burnt} · MB · C_{f} · G_{ef} · 10^{−3} ,   E = L_{gas} · GWP',
                 'Tonnes of CH₄ or N₂O, then tCO₂e (GWP AR5)', 'Emissions DEF CH4 N2O'),
                ('U_{E} = √( U_{A}^{2} + U_{MB}^{2} + U_{Cf}^{2} + U_{Gef}^{2} + U_{f}^{2} ) ,   E^{yr} = E / T',
                 'Uncertainty; immediate emissions', 'Emissions DEF CH4 N2O'),
            ]),
            ('5 · Degradation', [
                ('EF_{deg} = S_{forest 2021} − S_{forest 2025} ,   U = √[(U_{i}·S_{i})^{2} + (U_{f}·S_{f})^{2}] / |EF_{deg}|',
                 'Degradation EF (same zone)', 'Emissions DEG'),
                ('E = Â · EF_{deg} ,   U_{E} = √( U_{A}^{2} + U_{EF}^{2} ) ,   E^{yr} = E / T', 'Immediate emissions',
                 'Emissions DEG'),
            ]),
            ('6 · Enhancement of forest carbon stocks (gains)', [
                ('G = − G_{AGB} · (1 + R) · 44/12', 'Removal factor (tCO₂/ha/yr; R = 0: above-ground growth only)',
                 'Carbon stocks, table E'),
                ('E = Â · G · (T + 1) / 2 ,   E_{t} = (Â / T) · t · G ,   U_{E} = √( U_{A}^{2} + U_{G}^{2} )',
                 'Accumulated annual cohorts', 'Removals GAIN'),
            ]),
            ('7 · Balance', [
                ('E_{tot} = E_{DEF} + E_{SOC} + E_{CH4} + E_{N2O} + E_{DEG} ,   Balance = E_{tot} + E_{GAIN}',
                 'Total emissions and net balance',
                 'Balance, table A'),
                ('U = √Σ_{c} (U_{c} · E_{c})^{2} / |Σ_{c} E_{c}|', 'Combination of independent components (eq. 3.2)',
                 'Balance; Results'),
            ]),
        ]
        r, no = 4, 1
        for title, eqs in sections:
            self.h2(ws, r, 1, 4, title)
            ws.set_row(r, 22)
            r += 1
            for c, h in enumerate(['#', 'Equation', 'Description', 'Implementation in the workbook'], start=1):
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
        self.note(ws, r, 1, 4, 'References: IPCC (2006) Guidelines, vol. 1 ch. 3 and vol. 4 ch. 2, 4 and 5; IPCC (2019) '
                  'Refinement; Cochran (1977) Sampling Techniques; Olofsson et al. (2014) Remote Sensing of Environment 148.', 2)

    # ============================================================ Samples
    def sheet_samples(self):
        st, m = self.st, self.m
        ws = self.wb.add_worksheet(S_ECH)
        s = m.samples.reset_index(drop=True)
        headers = ['Sample', 'Stratum', 'Zone', 'Activity', 'CEO class 2021', 'CEO class 2025',
                   'Forest category 2021', 'Post-deforestation land use', 'Key DEF', 'y = net EF DEF (tCO₂/ha)', 'y²',
                   'Forest category 2025', 'Key DEG', 'Gain category 2021', 'Gain category 2025', 'Key GAIN'] + [h for h, _ in self.extra_cols]
        widths = [12, 8, 22, 9, 26, 26, 26, 26, 46, 14, 12, 26, 46, 26, 22, 46] + [12] * len(self.extra_cols)
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
            ws.write_formula(r, 6, f'=IF(OR($D{R}="Def",$D{R}="Deg"),IFERROR(VLOOKUP($E{R},Map_forest,2,FALSE),"{U}"),"")',
                             f_calc, row.fcat_i)
            ws.write_formula(r, 7, f'=IF($D{R}="Def",IFERROR(VLOOKUP($F{R},Map_post,2,FALSE),"{U}"),"")', f_calc, row.pcat)
            ws.write_formula(r, 8, f'=IF($D{R}="Def",$C{R}&"|"&G{R}&"|"&H{R},"")', f_key, row.key_def)
            ws.write_formula(r, 9, f'=IF($I{R}="",0,IFERROR(INDEX(EF_net_transitions,MATCH($I{R},Transition_keys,0)),0))',
                             f_calc, row.y)
            ws.write_formula(r, 10, f'=J{R}^2', f_calc, row.y2)
            ws.write_formula(r, 11, f'=IF($D{R}="Deg",IFERROR(VLOOKUP($F{R},Map_forest,2,FALSE),"{U}"),"")', f_calc, row.fcat_f)
            ws.write_formula(r, 12, f'=IF($D{R}="Deg",$C{R}&"|"&G{R}&"|"&L{R},"")', f_key, row.key_deg)
            ws.write_formula(r, 13, f'=IF($D{R}="Gain",IFERROR(VLOOKUP($E{R},Map_post,2,FALSE),'
                                    f'IFERROR(VLOOKUP($E{R},Map_forest,2,FALSE),"{U}")),"")', f_calc, row.gcat_i)
            ws.write_formula(r, 14, f'=IF($D{R}="Gain",IFERROR(VLOOKUP($F{R},Map_gain,2,FALSE),"{U}"),"")', f_calc, row.gcat)
            ws.write_formula(r, 15, f'=IF($D{R}="Gain",$C{R}&"|"&N{R}&"|"&O{R},"")', f_key, row.key_gain)
        for j, (_, colname) in enumerate(self.extra_cols, start=16):
            ws.write_column(1, j, [_clean(v) for v in s[colname]], f_id)
        ws.freeze_panes(1, 1)
        ws.autofilter(0, 0, len(s), len(headers) - 1)
        ws.data_validation(1, 3, len(s) + 2000, 3, {'validate': 'list',
                                                    'source': ['SNF', 'SF', 'SDegF', 'Deg', 'Def', 'Gain']})


def build_emissions_workbook(path, samples, area_hz, cats, maps, forest_stocks, post_stocks, soc_factors, removal_rates,
                             confidence=0.90, ef_confidence=0.90, initial_year=2021, final_year=2025, cf=0.47, cf_low=0.44,
                             socref=47.0, u_socref=0.0, soc_d=20, low_n=10, fire=None, fire_fraction=None, title=None,
                             subtitle=None, sources=None, extra_cols=(), territory=None):
    """Writes the workbook and returns the Python model (values identical to the workbook formulas).

    samples : DataFrame (plotid, strata, zone, activity, class_initial, class_final, + extra_cols), valid samples.
    area_hz : DataFrame indexed by zone, one column per stratum (ha).
    cats / maps : dict with the keys 'forest', 'post', 'gain' (ordered lists of categories / dict CEO class -> category).
    forest_stocks, post_stocks, soc_factors, removal_rates : DataFrame (see notebook 3).
    fire : dict of the burning parameters (cf_litter, comb, u_comb, gef_ch4, u_gef_ch4, gef_n2o, u_gef_n2o, gwp_ch4,
           gwp_n2o, u_fire); FREL 2024 values by default.
    fire_fraction : dict (forest category, post-deforestation land use) -> burnt fraction (0.5 by default).
    territory : None ("national" total) or name of the territory covered by area_hz (e.g. 'ERP'); only changes the labels.
    """
    params = dict(confidence=confidence, ef_confidence=ef_confidence, initial_year=initial_year, final_year=final_year,
                  cf=cf, cf_low=cf_low, socref=socref, u_socref=u_socref, soc_d=soc_d, low_n=low_n, **(fire or {}))
    model = EmissionModel(samples, area_hz, cats, maps, forest_stocks, post_stocks, soc_factors, removal_rates, params,
                          fire_fraction)
    title = title or "Côte d'Ivoire — Forest emissions and removals"
    subtitle = subtitle or (f'PROMIRE project / UN-REDD · SBAE {initial_year}-{final_year} · deforestation (biomass, soils, '
                            f'burning), degradation, enhancement of stocks · generated on {datetime.date.today():%Y-%m-%d}')
    b = _EmissionBuilder(path, model, title, subtitle, params, sources or {}, extra_cols, territory)
    b.build()
    return model
