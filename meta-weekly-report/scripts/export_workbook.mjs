// Run with the bundled Node: export_workbook.mjs analysis.json output.xlsx [node_modules]
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
const [input, output, deps] = process.argv.slice(2);
const api = deps ? await import(pathToFileURL(createRequire(path.join(deps,'_resolve.cjs')).resolve('@oai/artifact-tool')).href) : await import('@oai/artifact-tool');
const {Workbook,SpreadsheetFile}=api;
const a=JSON.parse(await fs.readFile(input,'utf8'));
const wb=Workbook.create();
const outdir=path.dirname(output);await fs.mkdir(outdir,{recursive:true});
const qaDir=path.join(path.dirname(input),'xlsx-preview');await fs.mkdir(qaDir,{recursive:true});
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const money='"$"#,##0.00;[Red]("$"#,##0.00);"$"0.00';
const names=a.tables;
for(const name of names)wb.worksheets.add(name);
const period=p=>p==='current'?`${a.period.current_start} 至 ${a.period.current_end}`:`${a.period.previous_start} 至 ${a.period.previous_end}`;
const ranges={};
function base(s,width=18){s.showGridLines=false;s.getRange('A1:AA3').format.font.name='Arial';s.getRange('A1:AA3').format.font.size=11;s.getRange('A1:AA3').format.columnWidth=width;s.getRange('A:A').format.columnWidth=38;}
function title(s,text,note){s.getRange('A1').values=[[text]];s.getRange('A1').format.font.bold=true;s.getRange('A1').format.font.size=16;s.getRange('A1').format.rowHeight=27;s.getRange('A2').values=[[note]];s.getRange('A2').format.rowHeight=22;}
function header(s,row,heads){const r=s.getRange(`A${row}:${col(heads.length-1)}${row}`);r.values=[heads];r.format.fill='#374151';r.format.font.color='#FFFFFF';r.format.font.bold=true;r.format.wrapText=true;r.format.rowHeight=32;r.format.verticalAlignment='center';}
function tableStyle(s,row,n,width){if(!n)return;const r=s.getRange(`A${row}:${col(width-1)}${row+n-1}`);r.format.font.name='Arial';r.format.font.size=11;r.format.rowHeight=30;r.format.verticalAlignment='center';s.getRange(`A${row}:A${row+n-1}`).format.wrapText=true;}
const sum=a.comparisons.structure.find(x=>x.key==='SUM-全店汇总');
const nonascNote=Object.keys(a.qa.nonasc_budget_spend).every(k=>k==='cbo')?'本次非ASC均为CBO。':'本次非ASC包含 '+Object.keys(a.qa.nonasc_budget_spend).join('、')+'，报告不得统称CBO。';
{
 const s=wb.worksheets.getItemAt(0);base(s);s.getRange('A:A').format.columnWidth=25;s.getRange('B:B').format.columnWidth=105;
 title(s,'Meta 周报数据口径',`源文件 ${a.source.filename}　币种 ${a.currency}`);
 const notes=[['项目','说明'],['本期',period('current')],['上期',period('previous')],['原始覆盖',`${a.source.first_day} 至 ${a.source.last_day}，${a.source.row_count} 条明细`],['比较方式',a.period.complete_week?'最近两个完整自然周，周一至周日':'最近自然周截至最新日期，与上周相同星期范围比较（部分周）'],['本期花费',sum.current.spend],['本期GMV',sum.current.gmv],['本期ROAS',null],['上期花费',sum.previous.spend],['上期GMV',sum.previous.gmv],['上期ROAS',null],['表格阅读','每张数据页依次为本期、上期、周度对比。保留两周维度并集，本期无记录显示0。右侧为计算底数。'],['比率口径','ROAS=GMV/花费；客单价=GMV/购物次数；CTR=链接点击/展示；ATC Rate=加购/LPV；转化率=购物次数/LPV。分母0时表内取0，正文不把缺分母解释为真实零效率。'],['补充漏斗口径','结账率=结账发起/加购；支付率=购物次数/结账发起。仅作聚合事件比值，非同一用户队列；比值超过100%时不作漏斗流失判断。'],['环比','(本期−上期)/上期。上期0、本期非0显示新增；两期均0显示-；使用未四舍五入值计算。'],['复合分类','表2素材类型、来源、受众、落地页存在交叉，不可加总；只有SUM是真实全店总计。'],['ASC','Campaign第2段或第3段=asc；非ASC不默认等同CBO。'+nonascNote],['两种素材来源','表2及报告分组用广告命名的素材类型/来源；表6/7/9来源从素材ID推断。h开头也可能在广告命名中标注ugc，不强行互改。'],['素材ID归并','小写+压缩多余空格作为key；显示名取各期组内总花费最高原始写法；编号文本存储。分隔符差异不自动合并。'],['达人','全来源提取达人名并归一汇总，无名排除；不得把表8达人汇总称为仅合创。'],['产品归因','多SPU广告的每个产品全额计入，不均分；表9不可加总为全店销量、GMV或产品真实销售额。'],['命名适配','按素材类型值定位后续来源与落地页，兼容站点段；移除连续广告副本后缀；保留产品代码末尾X/Y等单字母，不截断。'],['未识别产品',`本期相关花费 $${a.qa.unknown_product_spend.toFixed(2)}；${Object.keys(a.qa.unknown_products).join('、')||'无'}。仅产品字段标Unknown，花费/素材分析保留。`],['国家空值','空或nan国家从表4排除；本期排除花费 $'+a.qa.excluded_country.spend.toFixed(2)],['解析核对',a.qa.checks.join('；')],['缺失数值','空值及非数字按原规格记0；不将缺失日期视为零投放。'],['复核样例','旧周报文字及截图与本次导出可能不同，数字以本源表重算，不回填旧报告数字。'],['输入指纹',a.source.sha256]];
 s.getRange(`A4:B${notes.length+3}`).values=notes;header(s,4,notes[0]);s.getRange(`A5:B${notes.length+3}`).format.rowHeight=43;s.getRange(`A5:B${notes.length+3}`).format.wrapText=true;s.getRange(`A5:B${notes.length+3}`).format.font.size=11;
 s.getRange('B9:B10').setNumberFormat(money);s.getRange('B12:B13').setNumberFormat(money);
 s.getRange('B11').formulas=[['=IF(B9=0,0,B10/B9)']];s.getRange('B14').formulas=[['=IF(B12=0,0,B13/B12)']];s.getRange('B11').setNumberFormat('0.00"x"');s.getRange('B14').setNumberFormat('0.00"x"');
 ranges[names[0]]='A1:B14';
}
const dimensions=['structure','audience_category','audience','country','landing','idkey','template','creator','product_material'];
const standard=['花费','GMV','购物次数','ROAS','客单价','链接点击率','加购率','转化率'];
const helpers=['链接点击量','展示次数','LPV','ATC','结账发起','结账率','支付率'];
const helperKeys=['clicks','impressions','lpv','atc','checkout'];
for(let i=0;i<dimensions.length;i++){
 const dim=dimensions[i],s=wb.worksheets.getItemAt(i+1),xs=a.comparisons[dim];base(s);
 title(s,names[i+1],`本期 ${period('current')}　上期 ${period('previous')}　${a.currency}`);
 const isStruct=dim==='structure';
 let labels=dim==='idkey'?['素材ID','达人名','类型','编号','后缀','素材来源','需复核']:dim==='landing'?['落地页','形式']:dim==='product_material'?['SPU+Color','素材ID','素材来源']:[{'structure':'分类','audience_category':'受众大类','audience':'受众明细','country':'国家','template':'图片模板','creator':'达人名'}[dim]];
 const heads=isStruct?['分类','花费','花费环比','GMV','GMV环比','ROAS','ROAS环比','花费占比','GMV占比']:labels.concat(dim==='product_material'?['花费','GMV','出单数','ROAS','客单价','转化率']:standard);
 if(!xs.length){header(s,5,heads);s.getRange('A6').values=[['两周均无该维度记录']];ranges[names[i+1]]=`A1:${col(heads.length-1)}6`;continue;}
 const w=heads.length,n=xs.length,starts={current:5,previous:n+9};
 for(const p of ['current','previous']){
  const h=starts[p],r0=h+1; s.getRange(`A${h-1}`).values=[[(p==='current'?'本期 ':'上期 ')+period(p)]];s.getRange(`A${h-1}`).format.font.bold=true;header(s,h,heads);
  const body=xs.map(x=>{const m=x[p],meta=m.rows?m:(x.current.rows?x.current:x.previous);if(isStruct)return [x.label,m.spend,null,m.gmv,null,null,null,null,null];let id=dim==='idkey'?[meta.label,meta.creator,meta.type,meta.number,meta.suffix,meta.id_source,meta.review]:dim==='landing'?[x.label,meta.form]:dim==='product_material'?[x.key.split(' | ')[0],meta.label,meta.id_source]:[meta.label];return [...id,m.spend,m.gmv,m.purchases,...Array(w-id.length-3).fill(null)];});
  if(n)s.getRange(`A${r0}:${col(w-1)}${r0+n-1}`).values=body;
  tableStyle(s,r0,n,w);const j=labels.length;
  if(dim==='idkey'){s.getRange(`D${r0}:D${r0+n-1}`).setNumberFormat('@');s.getRange('B:B').format.columnWidth=24;s.getRange('C:D').format.columnWidth=12;s.getRange('E:E').format.columnWidth=30;s.getRange('F:F').format.columnWidth=14;s.getRange('G:G').format.columnWidth=26;s.getRange(`E${r0}:G${r0+n-1}`).format.wrapText=true;}
  if(dim==='product_material'){s.getRange('A:A').format.columnWidth=32;s.getRange('B:B').format.columnWidth=45;s.getRange('C:C').format.columnWidth=15;s.getRange(`B${r0}:B${r0+n-1}`).format.wrapText=true;s.getRange(`A${r0}:${col(w-1)}${r0+n-1}`).format.rowHeight=42;}
  if(dim==='idkey')s.getRange(`A${r0}:${col(w-1)}${r0+n-1}`).format.rowHeight=42;
  if(isStruct){
   const sr=r0+xs.findIndex(x=>x.key==='SUM-全店汇总');
   for(let k=0;k<n;k++){
    const row=r0+k,pr=starts.previous+1+k;
    for(const [c,formula] of [[2,p==='current'?`=IF(B${pr}=0,IF(B${row}=0,"-","新增"),(B${row}-B${pr})/B${pr})`:'="-"'],[4,p==='current'?`=IF(D${pr}=0,IF(D${row}=0,"-","新增"),(D${row}-D${pr})/D${pr})`:'="-"'],[5,`=IF(B${row}=0,0,D${row}/B${row})`],[6,p==='current'?`=IF(F${pr}=0,IF(F${row}=0,"-","新增"),(F${row}-F${pr})/F${pr})`:'="-"'],[7,`=IF(B$${sr}=0,0,B${row}/B$${sr})`],[8,`=IF(D$${sr}=0,0,D${row}/D$${sr})`]])s.getRange(`${col(c)}${row}`).formulas=[[formula]];
   }
   for(const c of ['B','D'])s.getRange(`${c}${r0}:${c}${r0+n-1}`).setNumberFormat(money);
   for(const c of ['C','E','G','H','I'])s.getRange(`${c}${r0}:${c}${r0+n-1}`).setNumberFormat('0.00%');
   s.getRange(`F${r0}:F${r0+n-1}`).setNumberFormat('0.00"x"');s.getRange(`A${sr}:I${sr}`).format.fill='#E5E7EB';s.getRange(`A${sr}:I${sr}`).format.font.bold=true;
  }else{
   const hc=w+2,he=hc+helpers.length-1;const hr=s.getRange(`${col(hc)}${h}:${col(he)}${h}`);hr.values=[helpers];hr.format.fill='#E5E7EB';hr.format.font.bold=true;hr.format.rowHeight=32;
   s.getRange(`${col(hc)}${r0}:${col(hc+4)}${r0+n-1}`).values=xs.map(x=>helperKeys.map(k=>x[p][k]));
   s.getRange(`${col(hc)}${r0}:${col(hc+4)}${r0+n-1}`).setNumberFormat('#,##0');
   for(let k=0;k<n;k++){
    const r=r0+k,sp=col(j)+r,gm=col(j+1)+r,pu=col(j+2)+r,lp=col(hc+2)+r,at=col(hc+3)+r,ch=col(hc+4)+r;
    const pairs=[[j+3,gm,sp],[j+4,gm,pu],...(dim==='product_material'?[[j+5,pu,lp]]:[[j+5,col(hc)+r,col(hc+1)+r],[j+6,at,lp],[j+7,pu,lp]]),[hc+5,ch,at],[hc+6,pu,ch]];
    for(const [c,num,den] of pairs)s.getRange(`${col(c)}${r}`).formulas=[[`=IF(${den}=0,0,${num}/${den})`]];
   }
   s.getRange(`${col(j)}${r0}:${col(j+1)}${r0+n-1}`).setNumberFormat(money);s.getRange(`${col(j+2)}${r0}:${col(j+2)}${r0+n-1}`).setNumberFormat('#,##0');s.getRange(`${col(j+3)}${r0}:${col(j+3)}${r0+n-1}`).setNumberFormat('0.00"x"');s.getRange(`${col(j+4)}${r0}:${col(j+4)}${r0+n-1}`).setNumberFormat(money);s.getRange(`${col(j+5)}${r0}:${col(w-1)}${r0+n-1}`).setNumberFormat('0.00%');s.getRange(`${col(hc+5)}${r0}:${col(hc+6)}${r0+n-1}`).setNumberFormat('0.00%');
  }
 }
 if(!isStruct){
  const h=2*n+13,r0=h+1,j=labels.length; s.getRange(`A${h-1}`).values=[['周度对比　两周维度并集']];
  header(s,h,['维度','本期花费','上期花费','花费环比','本期GMV','上期GMV','GMV环比','本期ROAS','上期ROAS','ROAS环比','GMV增减额']);
  s.getRange(`A${r0}:A${r0+n-1}`).values=xs.map(x=>[dim==='product_material'?x.key:x.label]);
  for(let k=0;k<n;k++){
   const r=r0+k,c=starts.current+1+k,p=starts.previous+1+k;
   const formulas=[`=${col(j)}${c}`,`=${col(j)}${p}`,`=IF(C${r}=0,IF(B${r}=0,"-","新增"),(B${r}-C${r})/C${r})`,`=${col(j+1)}${c}`,`=${col(j+1)}${p}`,`=IF(F${r}=0,IF(E${r}=0,"-","新增"),(E${r}-F${r})/F${r})`,`=${col(j+3)}${c}`,`=${col(j+3)}${p}`,`=IF(I${r}=0,IF(H${r}=0,"-","新增"),(H${r}-I${r})/I${r})`,`=E${r}-F${r}`];
   s.getRange(`B${r}:K${r}`).formulas=[formulas];
  }
  tableStyle(s,r0,n,11);s.getRange(`A${r0}:A${r0+n-1}`).format.rowHeight=dim==='product_material'?64:42;
  for(const c of ['B','C','E','F','K'])s.getRange(`${c}${r0}:${c}${r0+n-1}`).setNumberFormat(money);
  for(const c of ['D','G','J'])s.getRange(`${c}${r0}:${c}${r0+n-1}`).setNumberFormat('0.00%');
  s.getRange(`H${r0}:I${r0+n-1}`).setNumberFormat('0.00"x"');
 }
 s.freezePanes.freezeRows(5);
 ranges[names[i+1]]=`A1:${col(w-1)}${Math.min(n+5,15)}`;
}
console.log((await wb.inspect({kind:'table',range:'表2_花费结构!A6:I14',include:'values,formulas',tableMaxRows:9,tableMaxCols:9,maxChars:2500})).ndjson);
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:50},maxChars:2500});
await fs.writeFile(path.join(qaDir,'formula-scan.json'),errors.ndjson);console.log('Formula scan:',errors.ndjson);
for(let i=0;i<names.length;i++){
 const blob=await wb.render({sheetName:names[i],range:ranges[names[i]],scale:1.5,format:'png'});
 await fs.writeFile(path.join(qaDir,`sheet-${i+1}.png`),new Uint8Array(await blob.arrayBuffer()));
}
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(output);
console.log('Saved',output);
