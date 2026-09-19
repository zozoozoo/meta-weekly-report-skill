"""Behavioral regression tests. No source/output spreadsheet mutation."""
import datetime as dt, math, tempfile, unittest
from pathlib import Path
from unittest.mock import patch
from clean_meta import parse_name, products, split_id, id_source, template_id, aggregate, metrics, build, full_period, BASE

def row(cid='p26051452-inspo_2',day='2026-01-05',spend=10,gmv=20,**kw):
    r=parse_name('20260101^cov^cbo','20260101^us^web^eo^int-finance',f'20260101^{cid}^S225BS737_KHAKI+S225BS718_WHITE^sg^img^ugc^adp-shopall')
    r.update({k:0.0 for k in BASE});r.update(spend=spend,gmv=gmv,purchases=2,impressions=100,clicks=8,lpv=4,atc=3,checkout=2,day=dt.date.fromisoformat(day),country='US',currency='USD',account='A',line=2);r.update(kw);return r

class Regression(unittest.TestCase):
    def test_full_period_threshold_boundaries(self):
        rs=[row('p0001',spend=199.99,gmv=0,purchases=0),row('p0002',spend=200,gmv=0,purchases=0),
            row('p0003',spend=1000,gmv=0,purchases=0),row('p0004',spend=1000,gmv=2500,purchases=5),
            row('p0005',spend=1000,gmv=2499.99,purchases=5),row('p0006',spend=999.99,gmv=100,purchases=1)]
        a=full_period(rs);d={m['key']:m for m in a['materials']}
        self.assertEqual([d[f'p000{i}']['flagged'] for i in range(1,7)],[False,True,True,False,True,False])
        self.assertEqual(len(d['p0003']['flag_reason'].split('；')),2)
        self.assertEqual(d['p0006']['status'],'未触发规则')
    def test_full_period_includes_old_rows_and_keeps_versions(self):
        rs=[row('P0001',day='2026-01-01',spend=700,gmv=1000),row('p0001',day='2026-01-25',spend=400,gmv=1000),row('p0001-ac',spend=10,gmv=50)]
        a=full_period(rs);d={m['key']:m for m in a['materials']}
        self.assertEqual(len(d),2);self.assertEqual(d['p0001']['spend'],1100)
        self.assertAlmostEqual(d['p0001']['roas'],2000/1100);self.assertTrue(d['p0001']['flagged'])
        self.assertEqual(a['start'],'2026-01-01');self.assertEqual(a['end'],'2026-01-25')
    def test_cost_metrics_share_denominators(self):
        m=metrics([row(spend=100,gmv=300,purchases=2,impressions=10000,clicks=100,lpv=80)])
        self.assertEqual((m['cpm'],m['cpc'],m['cpa'],m['lpv_cost']),(10,1,50,1.25))
        self.assertAlmostEqual(m['roas'],m['aov']*m['cvr']/m['lpv_cost'])
    def test_station_and_asc_either_position(self):
        for campaign in ['d^asc^cbo','d^cov^asc']:
            r=parse_name(campaign,'d^us^web^eo^lal-1%','d^id^S225BS737^ca^acvid^UGC^adp+best - 广告副本 - 广告副本')
            self.assertTrue(r['asc']);self.assertEqual((r['ad_type'],r['ad_source'],r['landing']),('acvid','ugc','adp-best'))
    def test_catalog(self):
        r=parse_name('d^cov^abo','','d^catalog-sg');self.assertEqual(r['landing'],'catalog-sg');self.assertFalse(r['asc'])
    def test_product_colors_and_boundary(self):
        self.assertEqual(products('S124TY872-CREAM')[0],['S124TY872_CREAM'])
        self.assertEqual(products('S225BS737')[0],['S225BS737_KHAKI'])
        self.assertEqual(products('SW26DS115_LIGHT GRAY')[0],['SW26DS115_LT GRAY'])
        self.assertEqual(products('S326TS265X_ECRU')[0],['S326TS265X_ECRU'])
        self.assertEqual(products('S325JS7864')[0],['Unknown'])
    def test_id_number_and_suffix(self):
        self.assertEqual(split_id('Jasmina-v081601-Augw3')['number'],'081601')
        self.assertEqual(split_id('atgreenwell-v0823')['number'],'0823')
        self.assertEqual(split_id('p26051452-inspo_2')['suffix'],'inspo_2')
        self.assertEqual(split_id('h260205v1-mix')['suffix'],'v1-mix')
        self.assertEqual(split_id('Jasmina-724')['review'],'无数字编号')
        self.assertEqual(id_source('Jasmina-724'),'合创')
    def test_template_letters(self):
        self.assertEqual(template_id('p26061802-inspo_ai8-S126TS136_WHITE+S225BS737'),'inspo_ai8')
        self.assertEqual(template_id('p26082502-M17'),'m17')
    def test_aggregate_ratio_and_multispu(self):
        rs=[row(spend=1,gmv=10,lpv=2,purchases=1),row(spend=9,gmv=0,lpv=8,purchases=1)]
        m=metrics(rs);self.assertEqual(m['roas'],1);self.assertEqual(m['cvr'],.2)
        p=aggregate(rs,'product_material');self.assertEqual(sum(x['spend'] for x in p.values()),20)
    def test_creator_full_source_and_display(self):
        rs=[row('JASMINA-v081601',spend=8),row('JASMINA-v081601',spend=8),row('jasmina-v081601',spend=12),row('p26051452',spend=100)]
        d=aggregate(rs,'idkey');self.assertEqual(d['jasmina-v081601']['label'],'JASMINA-v081601')
        c=aggregate(rs,'creator');self.assertEqual(c['jasmina']['spend'],28);self.assertEqual(len(c),1)
    def test_week_ranges_and_new_disappeared(self):
        rs=[row(day=(dt.date(2026,1,5)+dt.timedelta(days=i)).isoformat(),spend=1,gmv=2) for i in range(14)]
        rs += [row('u26010501',day='2026-01-05',spend=10,gmv=30),row('v26011801',day='2026-01-18',spend=20,gmv=80)]
        qa=dict(currency='USD')
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'dummy.xlsx';path.write_bytes(b'fixture')
            with patch('clean_meta.load',return_value=(rs,qa)):
                a=build(path)
            self.assertEqual(a['period']['current_start'],'2026-01-12');self.assertTrue(a['period']['complete_week'])
            ids={x['key']:x for x in a['comparisons']['idkey']}
            self.assertIsNone(ids['v26011801']['wow']['roas']);self.assertEqual(ids['u26010501']['current']['spend'],0)
            self.assertEqual(a['comparisons']['structure'][8]['current']['spend'],27)
    def test_missing_day_not_zero_and_partial_week(self):
        rs=[row(day=(dt.date(2026,1,5)+dt.timedelta(days=i)).isoformat()) for i in range(10)]
        with tempfile.TemporaryDirectory() as td:
            p=Path(td)/'dummy.xlsx';p.write_bytes(b'x')
            with patch('clean_meta.load',return_value=(rs,dict(currency='USD'))):a=build(p)
            self.assertEqual(a['period']['previous_end'],'2026-01-07');self.assertFalse(a['period']['complete_week'])
            with patch('clean_meta.load',return_value=(rs[1:],dict(currency='USD'))):
                with self.assertRaisesRegex(ValueError,'缺少日期'):build(p)

if __name__=='__main__':unittest.main(verbosity=2)
