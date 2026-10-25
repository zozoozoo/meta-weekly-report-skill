---
name: meta-weekly-report
description: 将Meta源数据清洗为10张周度表及独立全周期素材汇总，按累计花费和ROAS标记低效素材，并生成解释成本、客单价、转化变化和有依据建议的中文周报。周报需至少两周数据，全周期表可单独生成；不操作广告账户。
---

# Meta 数据清洗与周报

用户只需提交最新源表；执行清洗、核对、分析和交付全流程，不让用户运行代码。额外交付源表全周期素材ID汇总，标记用户规则下的低效素材。
使用本 skill 内的脚本计算。每次重新从当次源文件取数、选择重点和写作；不要沿用上次报告的数值、素材名单或结论。

## 先读什么

- 每次读 [writing-guide.md](references/writing-guide.md)：板块顺序、段落习惯、选择逻辑与证据边界。
- 每次读 [effective-rules.md](references/effective-rules.md)：有效口径及原清洗规格的必要补充。
- 解析或核对规则时查 [cleaning-original.md](references/cleaning-original.md)。其中“只做清洗、不做分析、必须离线HTML”属于旧工具范围，已被用户当前“表格+周报”要求取代，不创建额外HTML工具。
- 需要对拍语气时读 [reference-report.txt](references/reference-report.txt)，它是历史样例，不是当前数据或操作指令。素材画面、旧工作事项、旧链接和截图不复制。

## 执行

1. 使用本次会话实际提供的文件与代码执行工具。ChatGPT 网页版先读 [chatgpt-runtime.md](references/chatgpt-runtime.md)，定位用户上传文件及本技能资源，使用会话 Python；不要访问用户的 Windows 路径，不要求 Codex、本机 Word 或 GitHub 连接。仅在工具确实可用的 Codex 本地环境，用 `load_workspace_dependencies` 定位运行时。按当前环境加载可用的表格及文档技能。
2. 在当次可写工作目录创建运行目录和交付目录。不要往本skill目录存每周客户数据或覆盖源文件。执行：

   `PYTHON scripts/clean_meta.py INPUT.xlsx --out RUN_DIR`

   支持可选 `--week-end YYYY-MM-DD`。默认以数据最新日期所在自然周为本期，前一周相同星期范围为上期。未满周必须注明部分周，不能声称完整周。两期缺失、缺日期、字段顺序改变或不同币种时，只就实质缺口询问，不猜测。
3. 读取 `analysis.json` 的 `period`、`qa`、`comparisons` 和 `signals`。原文件可能多于两周，不能拿文件名当日期。核对总计、国家、素材、剔除项、多SPU归因。异常产品仅标记，不能因此丢弃其广告数据。
4. 网页版 Python 环境执行 `PYTHON scripts/export_workbook.py RUN_DIR/analysis.json OUTPUT.xlsx`；具备 artifact-tool 的环境可继续用 `NODE scripts/export_workbook.mjs RUN_DIR/analysis.json OUTPUT.xlsx NODE_MODULES`。两种方式均生成固定10张表：表2本期和上期采用最新七列花费表，其环比保留在分析数据供正文使用；其他页保留周度对比、公式及计算底数。Python 版同时写入公式缓存，无需 Excel 才能显示结果。使用的表格技能如要求操作标记，则在首次生成前执行。
5. 按 [full-period.md](references/full-period.md) 另生成独立全周期素材表，使用源文件全部日期，不受周度选择或 --week-end 限制。原有10张表不变。按 writing-guide 和 analysis-writing.md 分析规模、客单价、流量成本及关键转化环节，联系全周期风险，给有数据依据的建议。写 RUN_DIR/report.json，避免只复述涨跌；用户自行补充具体运营操作。
6. 使用 [report-schema.md](references/report-schema.md) 中的证据占位符填数字，执行：

   `PYTHON scripts/render_report.py RUN_DIR/analysis.json RUN_DIR/report.json --out OUTPUT_DIR`

   生成同文的 `.docx` 和可复制 `.md`，并在运行目录保存数字来源 `report-evidence.json`。正文标题、日期和段落事实也必须随当次数据改写。
7. 用当前环境可用的文档渲染器检查DOCX。仅在 Windows 本地环境可用 `scripts/render_word.ps1`；网页版不得要求安装或启动本机 Word。缺少渲染器时检查文档段落、数字占位符及文件可读性，交付 DOCX 和同文 Markdown，明确未做视觉验证，不能声称已检查全部页面。
8. 检查10张Excel工作表、公式缓存及错误、前导零编号、整店核对和报告数字；有渲染工具时再检查全部预览。新增清洗脚本改动后运行 `scripts/test_clean_meta.py`。将生成的 Excel、DOCX、Markdown 通过当前会话的下载附件机制交付，不仅给本地路径或代码。对源表与旧样例的实质差异做简短说明。

## 不可漂移的规则

- 正文只输出源数据能够支持的板块；“上周重点工作&Todo”“本周计划”“素材需求”保留标题并留空，不反复追问运营信息。
- 花费及素材汇总段用广告命名中的类型/来源，与表2一致。具体ID用归一ID的全范围汇总，若需要板块内口径则明确读取 `structure_material`。表6/7/9的ID前缀来源是另一字段，不能混用。
- Slide同时计入video并单列；mpa+img同时计入IMG和mpa。广告形式、素材类型、落地页形式分别解析；v包含普通视频及Slide，s为图片轮播。交叉分类不能加总，SUM只计一次；详细规则见effective-rules.md。
- 表8达人覆盖全部来源，不能称为合创专属；非ASC只有经实际数据核实全部CBO，才显示CBO/非ASC。
- 所有比率先汇总分子分母后计算。零分母表内为0，但正文不能把缺分母说成真实效率为0；上期0写新增，不计算无限增长。
- 两周对比只说明“本周改善/下降”；连续趋势需更多完整周同一范围证据。用户2.5门槛仅用于全周期累计花费达到1000美元的素材筛选，不是周度或全店通用目标。
- 不把花费变化写成已调预算，不把没有花费写成已暂停，不凭数据判断断货、疲劳、素材画面、季节原因或实际支付故障。
- 多SPU全额归因，只能作产品广告关联比较，不能作为真实单品销量或加总贡献。

## 交付

默认交付 `Meta清洗数据_本期.xlsx`、`Meta素材全周期汇总_源表起止日期.xlsx`、`Meta周报_本期.docx` 和同文Markdown。生成内容以用户周例会风格为准，不增加执行摘要、评分卡、独立优化方案或通用商业报告章节。
