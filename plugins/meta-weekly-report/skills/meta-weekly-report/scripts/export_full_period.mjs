// Local artifact-tool export of the same payload used by the hosted Python exporter.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
const [input,output,deps]=process.argv.slice(2);
const {Workbook,SpreadsheetFile}=deps ? await import(pathToFileURL(createRequire(path.join(deps,'_resolve.cjs')).resolve('@oai/artifact-tool')).href) : await import('@oai/artifact-tool');
const p=JSON.parse(await fs.readFile(input,'utf8')),a=p.summary,cfg=a.rules;
const wb=Workbook.create(),s=wb.worksheets.add('素材ID_全周期'),n=p.data.length+5;
s.showGridLines=false;
s.getRange(`A1:Z${n}`).format.font.name='Arial';s.getRange(`A1:Z${n}`).format.font.size=11;
s.getRange('A:A').format.columnWidth=44;s.getRange('B:G').format.columnWidth=18;
s.getRange('H:P').format.columnWidth=16;s.getRange('Q:Q').format.columnWidth=55;s.getRange('R:Z').format.columnWidth=18;
s.getRange('A1').values=[['Meta 全周期素材汇总']];s.getRange('A1').format.font.bold=true;s.getRange('A1').format.font.size=16;
s.getRange('A2').values=[[`${a.start} 至 ${a.end}；USD；仅本次源文件覆盖周期；来源：${p.source}`]];
s.getRange('A3:F3').values=[['规则1花费≥',cfg.zero_purchase_spend_usd,'规则2花费≥',cfg.low_roas_spend_usd,'规则2 ROAS<',cfg.minimum_roas]];
s.getRange('A4').values=[['规则1还须购物次数=0；两条规则独立。红色为不达标；未触发规则不等于达标。']];
s.getRange('A5:Z5').values=[p.headers];s.getRange('A5:Z5').format.fill='#374151';s.getRange('A5:Z5').format.font.color='#FFFFFF';s.getRange('A5:Z5').format.font.bold=true;s.getRange('A5:Z5').format.rowHeight=32;
s.getRange(`A6:Z${n}`).values=p.data;s.getRange(`A6:Z${n}`).format.rowHeight=54;s.getRange(`A5:Z${n}`).format.wrapText=true;s.getRange(`A5:Z${n}`).format.verticalAlignment='center';
s.getRange(`D6:D${n}`).setNumberFormat('@');
for(const f of p.formulas)for(const [c,expr] of Object.entries(f))s.getRange(c).formulas=[[expr]];
for(const col of ['H','I','L'])s.getRange(`${col}6:${col}${n}`).setNumberFormat('"$"#,##0.00');
for(const col of ['M','N','O','Y','Z'])s.getRange(`${col}6:${col}${n}`).setNumberFormat('0.00%');
s.getRange(`K6:K${n}`).setNumberFormat('0.00');
s.freezePanes.freezeRows(5);
s.tables.add(`A5:Z${n}`,true,'AllMaterials');
s.getRange(`P6:Q${n}`).conditionalFormats.add('Custom',{formula:'=$P6="不达标"',format:{fill:'#FCE4D6',font:{color:'#9C0006'}}});
s.getRange(`A${n+3}`).values=[['同ID跨日期、账户、国家、受众汇总；不合并版本后缀。比率先汇总再计算，零分母表内为0。']];
s.getRange(`A${n+4}`).values=[['首次/最后记录不是实际上线/停投日期；缺失与非数字按既定口径记0：'+JSON.stringify({missing:p.qa.missing_as_zero,invalid:p.qa.non_numeric_as_zero})]];
await fs.mkdir(path.dirname(output),{recursive:true});
for(const [label,range] of [['summary','A1:P12'],['flags','P5:Z12']]){
 const preview=await wb.render({sheetName:s.name,range,scale:1.3,format:'png'});
 await fs.writeFile(path.join(path.dirname(input),label+'.png'),new Uint8Array(await preview.arrayBuffer()));
}
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:10},maxChars:1000})).ndjson);
await (await SpreadsheetFile.exportXlsx(wb)).save(output);
