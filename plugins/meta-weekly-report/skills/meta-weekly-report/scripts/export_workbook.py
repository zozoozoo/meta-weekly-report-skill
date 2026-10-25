"""Portable XLSX export for hosted Python sessions; cached formulas need no Excel.

Usage: python export_workbook.py analysis.json output.xlsx
"""
import argparse
import json
from pathlib import Path
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as cell
from clean_meta import STRUCT_LABELS


def write_structure(wb, s, a, xs):
    """Overlapping breakdowns; each share uses the single source SUM row."""
    labels = a.get('structure_labels', STRUCT_LABELS)
    s.set_column(0, 0, 27)
    s.set_column(1, 5, 17)
    s.set_column(6, 6, 28)
    base = dict(font_name='Arial', font_size=11, font_color='#000000',
                align='center', valign='vcenter', text_wrap=True, border=1)
    head = wb.add_format(dict(base, bold=True, bg_color='#EDF5C7'))
    styles = {}
    for is_sum in (False, True):
        for kind, nf in [('text',None),('num','#,##0.00'),('roas','0.00'),('pct','0.00%')]:
            spec = dict(base)
            if is_sum: spec.update(bg_color='#FFF24B')
            if nf: spec['num_format'] = nf
            else: spec['bold'] = True
            styles[is_sum, kind] = wb.add_format(spec)
    for p, h in [('current',4), ('previous',len(xs)+8)]:
        date = f"{a['period'][p+'_start']} 至 {a['period'][p+'_end']}"
        s.write(h-1, 0, '本期' if p=='current' else '上期')
        s.merge_range(h-1, 1, h-1, 5, date, head)
        s.merge_range(h-1, 6, h, 6, '备注', head)
        s.set_row(h-1, 25); s.set_row(h, 28)
        for c, label in enumerate(['广告类型','花费','GMV','ROAS','花费占比','GMV占比']):
            s.write(h,c,label,head)
        sr = h+1+next(i for i,x in enumerate(xs) if x['key']=='SUM-全店汇总')
        total = next(x[p] for x in xs if x['key']=='SUM-全店汇总')
        for i,x in enumerate(xs):
            r, m = h+1+i, x[p]
            is_sum = x['key']=='SUM-全店汇总'
            s.set_row(r, 42 if x['key'] in ('ADP','PDP') else 26)
            s.write(r,0,labels.get(x['key'],x['label']),styles[is_sum,'text'])
            s.write(r,1,m['spend'],styles[is_sum,'num'])
            s.write(r,2,m['gmv'],styles[is_sum,'num'])
            for c,nc,dr,dc,value,kind in [
                (3,2,r,1,m['roas'],'roas'),
                (4,1,sr,1,m['spend']/total['spend'] if total['spend'] else 0,'pct'),
                (5,2,sr,2,m['gmv']/total['gmv'] if total['gmv'] else 0,'pct')]:
                numerator, denominator = cell(r,nc), cell(dr,dc,row_abs=True)
                s.write_formula(r,c,f'=IF({denominator}=0,0,{numerator}/{denominator})',styles[is_sum,kind],value)
            s.write_blank(r,6,None,styles[is_sum,'text'])


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
        ['阅读方式', '表2按七列展示本期、上期；环比保留在分析数据供正文使用。其他页保留周度对比与计算底数；两周维度并集。'],
        ['比率', 'ROAS=GMV/花费；客单价=GMV/购物次数；CTR=点击/展示；加购率=加购/LPV；转化率=购物次数/LPV。先汇总再相除。'],
        ['零分母', '表内取0，正文不解释为真实零效率；上期0的环比显示新增或-。'],
        ['素材来源', '表2与正文类型分组使用广告命名；表6/7/9使用ID前缀来源，不混用。'],
        ['重叠归因', 'Slide同时计入video；mpa+img同时计入IMG与mpa。表2各类交叉，只有SUM是全店总计；多SPU全额归因，表9不可加总成全店或真实单品销售。'],
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
        if dim == 'structure':
            write_structure(wb, s, a, xs)
            continue
        labels = {'idkey': ['素材ID', '达人名', '类型', '编号', '后缀', '素材来源', '需复核'], 'landing': ['落地页', '形式'], 'product_material': ['SPU+Color', '素材ID', '素材来源']}.get(dim, ['分类'])
        keys = ['spend', 'gmv', 'purchases', 'roas', 'aov', 'ctr', 'atc_rate', 'cvr']
        heads = ['花费', 'GMV', '购物次数', 'ROAS', '客单价', '链接点击率', '加购率', '转化率']
        if dim == 'product_material':
            keys, heads = keys[:5]+['cvr'], heads[:5]+['转化率']
        headers = labels + heads
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
            for k, label in enumerate(['链接点击量', '展示次数', 'LPV', 'ATC', '结账发起', '结账率', '支付率']): s.write(h, hc+k, label, fmt(bold=True))
            for i, x in enumerate(xs):
                r, pr = h+1+i, starts['previous']+1+i
                m = x[p]
                meta = m if m.get('rows') else (x['current'] if x['current'].get('rows') else x['previous'])
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
