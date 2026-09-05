"""Portable XLSX export for hosted Python sessions; cached formulas need no Excel.

Usage: python export_workbook.py analysis.json output.xlsx
"""
import argparse
import json
from pathlib import Path
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as cell


def export(a, output):
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(output), {'strings_to_formulas': False, 'strings_to_urls': False})
    formats = {}
    def fmt(kind='text', bold=False):
        key = kind, bold
        if key not in formats:
            spec = dict(font_name='Arial', font_size=11, valign='vcenter', text_wrap=True)
            if bold:
                spec.update(bold=True, bg_color='#374151', font_color='white')
            if kind in ('money', 'num', 'pct', 'int'):
                spec['num_format'] = {'money': '"$"#,##0.00;[Red]("$"#,##0.00)', 'num': '0.00"x"', 'pct': '0.00%', 'int': '#,##0'}[kind]
            formats[key] = wb.add_format(spec)
        return formats[key]
    def row(s, r, values, bold=False):
        for c, value in enumerate(values):
            s.write(r, c, value, fmt(bold=bold))
        s.set_row(r, 32 if bold else 42)
    def formula(s, r, c, expr, value, kind='pct'):
        s.write_formula(r, c, expr, fmt(kind), value)
    def ratio(s, r, c, nr, nc, dr, dc, value, kind='pct'):
        n, d = cell(nr, nc), cell(dr, dc)
        formula(s, r, c, f'=IF({d}=0,0,{n}/{d})', value, kind)
    def wow(s, r, c, cr, cc, pr, pc, current, previous):
        x, y = cell(cr, cc), cell(pr, pc)
        value = (current-previous)/previous if previous else ('新增' if current else '-')
        formula(s, r, c, f'=IF({y}=0,IF({x}=0,"-","新增"),({x}-{y})/{y})', value)
    def sheet(name):
        s = wb.add_worksheet(name)
        s.hide_gridlines(2)
        s.set_column(0, 0, 42)
        s.set_column(1, 26, 18)
        s.freeze_panes(5, 0)
        s.set_landscape()
        s.fit_to_pages(1, 0)
        s.write(0, 0, name, wb.add_format({'bold': True, 'font_size': 16}))
        s.write(1, 0, f"本期 {a['period']['current_start']} 至 {a['period']['current_end']}；上期 {a['period']['previous_start']} 至 {a['period']['previous_end']}；{a['currency']}")
        return s
    total = next(x for x in a['comparisons']['structure'] if x['key'] == 'SUM-全店汇总')
    s = sheet(a['tables'][0])
    s.set_column(1, 1, 108)
    notes = [
        ['项目', '说明'], ['源文件', a['source']['filename']],
        ['输入指纹', a['source']['sha256']], ['原始明细数', a['source']['row_count']],
        ['比较方式', '两个完整自然周' if a['period']['complete_week'] else '部分周，与上周相同星期范围比较'],
        ['阅读方式', '本期、上期及周度对比；两周维度并集。右侧保留比率计算底数。'],
        ['比率', 'ROAS=GMV/花费；客单价=GMV/购物次数；CTR=点击/展示；加购率=加购/LPV；转化率=购物次数/LPV。先汇总再相除。'],
        ['零分母', '表内取0，正文不解释为真实零效率；上期0的环比显示新增或-。'],
        ['素材来源', '表2与正文类型分组使用广告命名；表6/7/9使用ID前缀来源，不混用。'],
        ['重叠归因', '表2各类交叉，只有SUM是全店总计；多SPU每个产品全额归因，表9不可加总成全店或真实单品销售。'],
        ['达人', '表8涵盖全部来源，无名排除，并非仅合创。'],
        ['非ASC预算类别', '、'.join(a['qa']['nonasc_budget_spend'])],
        ['事件比值', '结账率=结账发起/加购；支付率=购物次数/结账发起。非同用户队列，超过100%不作流失判断。'],
        ['缺失数值', '空值及非数字按0处理；日期缺失不能当作零投放。'],
        ['未识别产品花费', a['qa']['unknown_product_spend']],
        ['未识别产品', '、'.join(a['qa']['unknown_products']) or '无'],
        ['国家剔除花费', a['qa']['excluded_country']['spend']],
        ['核对', '；'.join(a['qa']['checks'])],
        ['历史样例', '仅供写作风格参考；数字以本次源表为准。'],
    ]
    for r, values in enumerate(notes, 4): row(s, r, values, r == 4)
    for p in ('current', 'previous'):
        r = s.dim_rowmax + 2
        row(s, r, [p, '花费', 'GMV', 'ROAS'], True)
        row(s, r+1, [p, total[p]['spend'], total[p]['gmv']])
        ratio(s, r+1, 3, r+1, 2, r+1, 1, total[p]['roas'], 'num')
    dims = ['structure', 'audience_category', 'audience', 'country', 'landing', 'idkey', 'template', 'creator', 'product_material']
    for name, dim in zip(a['tables'][1:], dims):
        s = sheet(name)
        xs = a['comparisons'][dim]
        labels = {'idkey': ['素材ID', '达人名', '类型', '编号', '后缀', '素材来源', '需复核'], 'landing': ['落地页', '形式'], 'product_material': ['SPU+Color', '素材ID', '素材来源']}.get(dim, ['分类'])
        keys = ['spend', 'gmv', 'purchases', 'roas', 'aov', 'ctr', 'atc_rate', 'cvr']
        heads = ['花费', 'GMV', '购物次数', 'ROAS', '客单价', '链接点击率', '加购率', '转化率']
        if dim == 'product_material':
            keys, heads = keys[:5]+['cvr'], heads[:5]+['转化率']
        structure = dim == 'structure'
        if structure:
            headers = ['分类', '花费', '花费环比', 'GMV', 'GMV环比', 'ROAS', 'ROAS环比', '花费占比', 'GMV占比']
        else: headers = labels + heads
        n, j = len(xs), len(labels)
        starts = {'current': 4, 'previous': n+8}
        if not n:
            row(s, 4, headers, True)
            row(s, 5, ['两周均无该维度记录'])
            continue
        for p, h in starts.items():
            s.write(h-1, 0, '本期' if p == 'current' else '上期')
            row(s, h, headers, True)
            hc = len(headers)+2
            if not structure:
                for k, label in enumerate(['链接点击量', '展示次数', 'LPV', 'ATC', '结账发起', '结账率', '支付率']): s.write(h, hc+k, label, fmt(bold=True))
            for i, x in enumerate(xs):
                r, pr = h+1+i, starts['previous']+1+i
                m = x[p]
                meta = m if m.get('rows') else (x['current'] if x['current'].get('rows') else x['previous'])
                if structure:
                    row(s, r, [x['label'], m['spend'], None, m['gmv']])
                    s.write(r, 1, m['spend'], fmt('money')); s.write(r, 3, m['gmv'], fmt('money'))
                    ratio(s, r, 5, r, 3, r, 1, m['roas'], 'num')
                    for c, src, key in [(2,1,'spend'), (4,3,'gmv'), (6,5,'roas')]:
                        if p == 'current': wow(s,r,c,r,src,pr,src,m[key],x['previous'][key])
                        else: s.write(r,c,'-',fmt())
                    sr = h+1+next(k for k,y in enumerate(xs) if y['key']=='SUM-全店汇总')
                    for c, src, key in [(7,1,'spend'),(8,3,'gmv')]: ratio(s,r,c,r,src,sr,src,m[key]/total[p][key] if total[p][key] else 0)
                    continue
                ids = [meta['label']]
                if dim == 'idkey': ids = [meta[k] for k in ['label','creator','type','number','suffix','id_source','review']]
                if dim == 'landing': ids = [x['label'],meta['form']]
                if dim == 'product_material': ids = [x['key'].split(' | ')[0],meta['label'],meta['id_source']]
                row(s,r,ids)
                for k, key in enumerate(keys[:3]): s.write(r,j+k,m[key],fmt('int' if k==2 else 'money'))
                for k,key in enumerate(['clicks','impressions','lpv','atc','checkout']): s.write(r,hc+k,m[key],fmt('int'))
                for c,nc,dc,key,kind in [(j+3,j+1,j,'roas','num'),(j+4,j+1,j+2,'aov','money')]: ratio(s,r,c,r,nc,r,dc,m[key],kind)
                pairs = [(j+5,j+2,hc+2,'cvr')] if dim=='product_material' else [(j+5,hc,hc+1,'ctr'),(j+6,hc+3,hc+2,'atc_rate'),(j+7,j+2,hc+2,'cvr')]
                for c,nc,dc,key in pairs: ratio(s,r,c,r,nc,r,dc,m[key])
                ratio(s,r,hc+5,r,hc+4,r,hc+3,m['checkout']/m['atc'] if m['atc'] else 0)
                ratio(s,r,hc+6,r,j+2,r,hc+4,m['purchases']/m['checkout'] if m['checkout'] else 0)
        if not structure:
            h = 2*n+12
            row(s,h,['维度','本期花费','上期花费','花费环比','本期GMV','上期GMV','GMV环比','本期ROAS','上期ROAS','ROAS环比','GMV增减额'],True)
            for i,x in enumerate(xs):
                r, cr, pr = h+1+i, starts['current']+1+i, starts['previous']+1+i
                row(s,r,[x['key'] if dim=='product_material' else x['label']])
                for c,rr,cc,key,p,kind in [(1,cr,j,'spend','current','money'),(2,pr,j,'spend','previous','money'),(4,cr,j+1,'gmv','current','money'),(5,pr,j+1,'gmv','previous','money'),(7,cr,j+3,'roas','current','num'),(8,pr,j+3,'roas','previous','num')]: formula(s,r,c,'='+cell(rr,cc),x[p][key],kind)
                for c,cc,pc,key in [(3,1,2,'spend'),(6,4,5,'gmv'),(9,7,8,'roas')]: wow(s,r,c,r,cc,r,pc,x['current'][key],x['previous'][key])
                formula(s,r,10,f'={cell(r,4)}-{cell(r,5)}',x['delta']['gmv'],'money')
        if dim in ('idkey','product_material'): s.set_column(1,1,35)
    wb.close()
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('analysis', type=Path)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    export(json.loads(args.analysis.read_text(encoding='utf-8')), args.output)
