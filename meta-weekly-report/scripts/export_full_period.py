"""Standalone full-source ID report. Python export for hosted sessions.

python export_full_period.py INPUT.xlsx --out OUTPUT.xlsx [--json-only payload.json]
No weekly date filtering or minimum two-week requirement for this report.
"""
import argparse, json
from pathlib import Path
from clean_meta import load, full_period

HEADERS=['素材ID','达人名','类型','编号','后缀','素材来源','需复核','花费','GMV','购物次数','ROAS','客单价','链接点击率','加购率','转化率',
         '达标标记','标记原因','周期内首次记录','周期内最后记录','链接点击量','展示次数','LPV','ATC','结账发起','结账率','支付率']

def payload(path):
    rows,qa=load(path);a=full_period(rows);cfg=a['rules']
    data=[];formulas=[]
    for r,m in enumerate(a['materials'],6):
        values=[m[k] for k in ['label','creator','type','number','suffix','id_source','review','spend','gmv','purchases','roas','aov','ctr','atc_rate','cvr']]
        reason=('规则1：达到花费线且购物次数为0' if m['spend']>=cfg['zero_purchase_spend_usd'] and m['purchases']==0 else '')
        rule2=m['spend']>=cfg['low_roas_spend_usd'] and m['roas']<cfg['minimum_roas']
        if rule2: reason+=('；' if reason else '')+'规则2：达到花费线且ROAS低于门槛'
        values += [m['status'],reason,m['first_day'],m['last_day']]+[m[k] for k in ['clicks','impressions','lpv','atc','checkout','checkout_rate','payment_rate']]
        data.append(values)
        f={}
        for dest,num,den in [('K','I','H'),('L','I','J'),('M','T','U'),('N','W','V'),('O','J','V'),('Y','X','W'),('Z','J','X')]:
            f[f'{dest}{r}']=f'=IF({den}{r}=0,0,{num}{r}/{den}{r})'
        c1=f'AND(H{r}>=$B$3,J{r}=0)';c2=f'AND(H{r}>=$D$3,K{r}<$F$3)'
        f[f'P{r}']=f'=IF(OR({c1},{c2}),"不达标","未触发规则")'
        f[f'Q{r}']=f'=IF({c1},"规则1：达到花费线且购物次数为0","")&IF(AND({c1},{c2}),"；","")&IF({c2},"规则2：达到花费线且ROAS低于门槛","")'
        formulas.append(f)
    return dict(source=path.name,summary=a,qa=qa,headers=HEADERS,data=data,formulas=formulas)

def export(p,out):
    import xlsxwriter
    from xlsxwriter.utility import xl_cell_to_rowcol
    out=Path(out);out.parent.mkdir(parents=True,exist_ok=True)
    wb=xlsxwriter.Workbook(str(out),{'strings_to_formulas':False,'strings_to_urls':False})
    s=wb.add_worksheet('素材ID_全周期');s.hide_gridlines(2);s.freeze_panes(5,1)
    base={'font_name':'Arial','font_size':11,'valign':'vcenter','text_wrap':True}
    text=wb.add_format(base);note=wb.add_format(dict(base,text_wrap=False));head=wb.add_format(dict(base,bold=True,bg_color='#374151',font_color='white'))
    money=wb.add_format(dict(base,num_format='"$"#,##0.00'));num=wb.add_format(dict(base,num_format='0.00'))
    pct=wb.add_format(dict(base,num_format='0.00%'));integer=wb.add_format(dict(base,num_format='#,##0'))
    s.set_column(0,0,44);s.set_column(1,6,18);s.set_column(7,15,16);s.set_column(16,16,55);s.set_column(17,25,18)
    s.write(0,0,'Meta 全周期素材汇总',wb.add_format(dict(base,bold=True,font_size=16)))
    a=p['summary'];cfg=a['rules']
    s.write(1,0,f"{a['start']} 至 {a['end']}；USD；仅本次源文件覆盖周期；来源：{p['source']}")
    for c,v in enumerate(['规则1花费≥',cfg['zero_purchase_spend_usd'],'规则2花费≥',cfg['low_roas_spend_usd'],'规则2 ROAS<',cfg['minimum_roas']]):s.write(2,c,v,text)
    s.write(3,0,'规则1还须购物次数=0；两条规则独立。红色为不达标；未触发规则不等于达标。',note)
    s.set_row(3,36)
    s.write_row(4,0,p['headers'],head);s.set_row(4,32)
    for i,values in enumerate(p['data'],5):
        s.set_row(i,54)
        for j,v in enumerate(values):
            form=money if j in (7,8,11) else num if j==10 else pct if j in (12,13,14,24,25) else integer if j in (9,19,20,21,22,23) else text
            s.write(i,j,v,form)
        for address,expr in p['formulas'][i-5].items():
            r,c=xl_cell_to_rowcol(address)
            form=num if c==10 else money if c==11 else pct if c in (12,13,14,24,25) else text
            s.write_formula(address,expr,form,values[c])
    n=len(p['data'])+4
    s.autofilter(4,0,n,25)
    red=wb.add_format({'bg_color':'#FCE4D6','font_color':'#9C0006'})
    s.conditional_format(5,15,n,16,{'type':'formula','criteria':'=$P6="不达标"','format':red})
    s.write(n+3,0,'口径',head)
    s.write(n+4,0,'同ID跨日期、账户、国家、受众汇总；仅归一大小写和多余空格，不合并版本后缀。',note)
    s.write(n+5,0,'比率先汇总再计算；分母为0表内记0，不能解释为真实零效率。首次/最后记录不是实际上线/停投日期。',note)
    s.write(n+6,0,'缺失及非数字处理记录：'+json.dumps({k:p['qa'][k] for k in ['missing_as_zero','non_numeric_as_zero']},ensure_ascii=False),note)
    for r in (n+4,n+5,n+6):s.set_row(r,45)
    wb.close()

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('input',type=Path);ap.add_argument('--out',type=Path);ap.add_argument('--json-only',type=Path);args=ap.parse_args()
    p=payload(args.input)
    if args.json_only:
        args.json_only.parent.mkdir(parents=True,exist_ok=True);args.json_only.write_text(json.dumps(p,ensure_ascii=False),encoding='utf-8')
    elif args.out:export(p,args.out)
    else:ap.error('--out or --json-only required')
