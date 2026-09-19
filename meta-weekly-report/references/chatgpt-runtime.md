# ChatGPT 网页版运行

此文件指导已加载本技能的 ChatGPT 使用会话内代码执行工具完成任务，不负责安装插件。

## 输入与环境

- INPUT 是当次上传的真实 XLSX 文件路径，不是历史 Downloads 路径或 GitHub 链接。用文件工具确认实际附件；多份文件仅在无法判定源表时询问。
- SKILL_ROOT 是当前工具实际提供的技能资源目录。读取配套脚本与 references；不要假设工作目录就是技能目录。
- Python 依赖：openpyxl（读源表）、xlsxwriter（导出含缓存公式的 Excel）、python-docx（写 Word）。先检查导入。缺少库时仅在环境支持且允许时安装；不能声称未执行的脚本已成功。
- 不需要 API Key、Meta 账户或 GitHub 连接；每周业务数据不提交回仓库。
- 如果技能只显示文本、配套文件未挂载，先用当前技能资源工具读取文件；若仍不可获取，明确说明资源缺失，不能编造路径或省略清洗规则。如果当前对话没有代码执行能力，明确说明并请用户换到有数据分析/代码执行能力的会话。

## Python 调用方式

在会话可写目录建立本次运行目录。命令应通过参数列表传递，支持空格与中文路径：

```python
import subprocess, sys
from pathlib import Path

# skill_root、input_path、run_dir、out_dir 均从本次真实工具结果中取得。
run_dir.mkdir(parents=True, exist_ok=True)
out_dir.mkdir(parents=True, exist_ok=True)
subprocess.run([sys.executable, str(skill_root/'scripts/clean_meta.py'),
                str(input_path), '--out', str(run_dir)], check=True)
subprocess.run([sys.executable, str(skill_root/'scripts/export_workbook.py'),
                str(run_dir/'analysis.json'), str(out_dir/'Meta清洗数据.xlsx')], check=True)
```

然后读取 analysis.json，按 writing-guide.md 和 report-schema.md 为本次数据撰写 report.json，再运行 render_report.py。不能仅导出数据表就结束，也不能复用上次 report.json。脚本负责准确计算及填数，助手负责阅读变化、选择重点和组织段落。

另外按 full-period.md 运行 export_full_period.py，生成独立的源表全周期素材汇总。每次交付两份Excel，不把累计表代替周度表；阈值只用于累计素材筛选。

数字核对：用 openpyxl 的 data_only=True 读最终 Excel，核对两周全店 spend/GMV/ROAS 及正文引用，以及全周期素材总额与全量源表；另用 data_only=False 确认比率与标记为公式。ID 编号须保持文本前导零。

正文和下载文件均按当次日期命名。提供真实存在、可下载的 Excel、Word 和 Markdown 链接；没有渲染工具时准确说明视觉检查限制。不要把普通附件上传称为技能已安装。
