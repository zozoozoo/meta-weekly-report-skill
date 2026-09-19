# report.json

报告由当次执行skill的助手阅读证据后写作，脚本负责填数字与排版，不机械生成判断。

```json
{
  "refs": {
    "total": "structure::SUM-全店汇总",
    "chosen": "idkey::当次analysis.json中存在的归一key"
  },
  "blocks": [
    {"role":"title","text":"实际本期日期周例会"},
    {"role":"h1","text":"花费"},
    {"role":"p","text":"本周花费环比{{total.wow.spend|change}}，ROAS从 {{total.previous.roas|num}} 变为 {{total.current.roas|num}}。"},
    {"role":"p","text":"**{{chosen.label}}**本周花费 {{chosen.current.spend|money}}，出单 {{chosen.current.purchases|int}} 单。依据这些数字补上当期判断。"}
  ]
}
```

周度`refs`指向当次`comparisons[维度]`内的key，全周期引用使用下述full_period入口。别名自定；分隔符是第一个`::`。
每个对象含`current`、`previous`、`delta`、`wow`，基础字段有spend/gmv/purchases/impressions/clicks/lpv/atc/checkout，派生有roas/aov/ctr/atc_rate/cvr/checkout_rate/payment_rate/cpm/cpc/cpa/lpv_cost。`label`是显示名。成本字段用money格式，分母0不能解释为实际效率。

全周期引用用 `"cumulative": "full_period::归一素材key"`，直接引用 `{{cumulative.spend|money}}`、`{{cumulative.roas|num}}`、`{{cumulative.status}}`，不加current/previous；范围仅源文件全部日期。正文须说明该数字是累计而非本周。

格式：money美元两位小数；num两位小数；int整数；pct带符号百分比；abs_pct绝对百分比；change返回“增长X%/下降X%/持平/新增”；text直接文本。`current.spend|share`和`current.gmv|gmv_share`以别名total同期为分母，必须定义total。
`wow`为null表示上期0，写作必须依据上下文使用“新增”或无基准，不能让模板自动补出错误因果。

角色：title、h1、h2、h3、p、note、blank。`**…**`加粗；可选`page_break:true`只用于解决实际分页问题。不要嵌入旧报告的静态经营数值。金额、比率、订单数和绝对增减尽量全部用占位符。
`render_report.py`会验证每个引用并生成`report-evidence.json`，但不会替代对“升/降、因果、范围、样本量”的人工核对。
