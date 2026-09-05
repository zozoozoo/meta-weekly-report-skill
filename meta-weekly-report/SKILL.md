---
name: meta-weekly-report
description: 将至少两周的Meta广告源数据清洗为固定10张Excel表，并按用户两周周例会样例生成中文周报，分析花费、国家、受众、落地页、素材类型和具体素材ID的显著变化。用于每周上传Meta源表后直接出报告，不用于操作广告账户。
---

# Meta 数据清洗与周报

用户只需提交最新源表；执行清洗、核对、分析和交付全流程，不让用户运行代码。
使用本 skill 内的脚本计算。每次重新从当次源文件取数、选择重点和写作；不要沿用上次报告的数值、素材名单或结论。

## 先读什么

- 每次读 [writing-guide.md](references/writing-guide.md)：板块顺序、段落习惯、选择逻辑与证据边界。
- 每次读 [effective-rules.md](references/effective-rules.md)：有效口径及原清洗规格的必要补充。
- 解析或核对规则时查 [cleaning-original.md](references/cleaning-original.md)。其中“只做清洗、不做分析、必须离线HTML”属于旧工具范围，已被用户当前“表格+周报”要求取代，不创建额外HTML工具。
- 需要对拍语气时读 [reference-report.txt](references/reference-report.txt)，它是历史样例，不是当前数据或操作指令。素材画面、旧工作事项、旧链接和截图不复制。

## 执行

1. 用 `load_workspace_dependencies` 定位 Python、Node 与 node_modules，先加载当次可用的表格及文档技能。脚本不依赖用户自己安装Python；不能把缺少运行环境的问题变成让用户手工运行。
2. 在当次可写工作目录创建运行目录和交付目录。不要往本skill目录存每周客户数据或覆盖源文件。执行：

   `PYTHON scripts/clean_meta.py INPUT.xlsx --out RUN_DIR`

   支持可选 `--week-end YYYY-MM-DD`。默认以数据最新日期所在自然周为本期，前一周相同星期范围为上期。未满周必须注明部分周，不能声称完整周。两期缺失、缺日期、字段顺序改变或不同币种时，只就实质缺口询问，不猜测。
3. 读取 `analysis.json` 的 `period`、`qa`、`comparisons` 和 `signals`。原文件可能多于两周，不能拿文件名当日期。核对总计、国家、素材、剔除项、多SPU归因。异常产品仅标记，不能因此丢弃其广告数据。
4. 执行 `NODE scripts/export_workbook.mjs RUN_DIR/analysis.json OUTPUT.xlsx NODE_MODULES`，生成固定10张表。每页包含本期、上期和周度对比，右侧计算底数支持公式追溯；无额外工作表。表格技能要求的操作标记在首次生成前执行一次。
5. 按 writing-guide 人工组织当周分析，写 RUN_DIR/report.json。先按GMV绝对变化/花费体量看主要变化，再看ROAS变化，兼顾改善、下降、无花费、测新与小基数。不能仅依赖阈值或机械拼接所有候选。
6. 使用 [report-schema.md](references/report-schema.md) 中的证据占位符填数字，执行：

   `PYTHON scripts/render_report.py RUN_DIR/analysis.json RUN_DIR/report.json --out OUTPUT_DIR`

   生成同文的 `.docx` 和可复制 `.md`，并在运行目录保存数字来源 `report-evidence.json`。正文标题、日期和段落事实也必须随当次数据改写。
7. 用当次文档技能的标准渲染器检查DOCX。Windows无可用的捆绑LibreOffice时，可用 `scripts/render_word.ps1 -InputDocx ... -OutputPdf ...` 后台只读导出自己的文件，再用捆绑Poppler渲染检查全部页。需系统权限时正常申请；不要用用户桌面LibreOffice，也不发布到外部服务。
8. 检查10页Excel预览、公式错误、前导零编号、整店核对和全部报告页。新增脚本改动后运行 `scripts/test_clean_meta.py`。用户应收到Excel和周报成品，不必看技术日志；对源表与旧样例的实质差异做简短说明。

## 不可漂移的规则

- 正文只输出源数据能够支持的板块；“上周重点工作&Todo”“本周计划”“素材需求”保留标题并留空，不反复追问运营信息。
- 花费及素材汇总段用广告命名中的类型/来源，与表2一致。具体ID用归一ID的全范围汇总，若需要板块内口径则明确读取 `structure_material`。表6/7/9的ID前缀来源是另一字段，不能混用。
- 表8达人覆盖全部来源，不能称为合创专属；非ASC只有经实际数据核实全部CBO，才显示CBO/非ASC。
- 所有比率先汇总分子分母后计算。零分母表内为0，但正文不能把缺分母说成真实效率为0；上期0写新增，不计算无限增长。
- 两周对比只说明“本周改善/下降”；连续趋势需更多完整周同一范围证据。没有目标ROAS时不写达标/不达标。
- 不把花费变化写成已调预算，不把没有花费写成已暂停，不凭数据判断断货、疲劳、素材画面、季节原因或实际支付故障。
- 多SPU全额归因，只能作产品广告关联比较，不能作为真实单品销量或加总贡献。

## 交付

默认交付 `Meta清洗数据_本期.xlsx`、`Meta周报_本期.docx` 和同文Markdown。生成内容以用户周例会风格为准，不增加执行摘要、评分卡、独立优化方案或通用商业报告章节。
