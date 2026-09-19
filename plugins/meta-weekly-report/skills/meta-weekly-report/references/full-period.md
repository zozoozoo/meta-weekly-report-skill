# 源表全周期素材汇总

每次周报额外交付独立Excel。使用本次源文件全部日期，不仅最近两周，也不是源表以外的素材历史生命周期。原周度10张表保持不变。

同ID跨日期、账户、国家、受众合计，先累加底数再算比率。归一大小写和多余空格，不合并版本后缀；保留全部ID，包括最近两周无记录的老素材。np等无有效ID的占位组保留并提示复核，不能当作独立真实素材。

前15列与周度素材ID表一致：素材ID、达人名、类型、编号、后缀、素材来源、需复核、花费、GMV、购物次数、ROAS、客单价、链接点击率、加购率、转化率。随后增加标记、原因、周期内首次/最后记录；右侧保留底数和结账率、支付率。默认标记行优先，组内按花费降序，可筛选。

规则配置见material-thresholds.json：
- 累计花费≥200美元且购物次数=0。
- 累计花费≥1000美元且累计ROAS<2.5。
- 任一触发即不达标，两条可同时触发。用未四舍五入值比较，ROAS恰好2.5不触发第二条。未触发不等于达标。
- 缺失/非数字沿清洗口径记0并披露；购买指标缺失时标记需要核对，不能断言实际无转化。

网页版执行：
`PYTHON scripts/export_full_period.py INPUT.xlsx --out OUTPUT.xlsx`

有artifact-tool的本地环境执行：
`PYTHON scripts/export_full_period.py INPUT.xlsx --json-only RUN_DIR/full-period.json`
`NODE scripts/export_full_period.mjs RUN_DIR/full-period.json OUTPUT.xlsx NODE_MODULES`

NODE_MODULES使用当次依赖工具返回的目录，不写死用户路径；包在正常模块搜索路径时可省略。

独立脚本不要求至少两周。周报 analysis.json 也包含full_period，包括materials、total、起止日、规则及标记数，可供正文联系本周与累计表现。首次/最后记录不是实际投放起止日期。

核对累计花费、GMV、购物次数与源行全量合计一致，全部ID仅一行，检查200/1000/2.5边界和双命中，编号保留前导零。表内门槛控制标记公式；修改门槛后可重新筛选，默认行序是生成时排序。
