"""Build a self-contained plugin and standalone skill ZIP, without weekly data."""
import json
import shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent
source = ROOT/'meta-weekly-report'
plugin = ROOT/'plugins/meta-weekly-report'
target = plugin/'skills/meta-weekly-report'
for path in source.rglob('*'):
    if not path.is_file() or '__pycache__' in path.parts or path.suffix == '.pyc':
        continue
    dest = target/path.relative_to(source)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, dest)
manifest = {
    'name': 'meta-weekly-report', 'version': '0.3.0',
    'description': '上传Meta源数据，生成周度分析、全周期素材低效筛选和中文周报。',
    'author': {'name': 'zozoozoo'},
    'repository': 'https://github.com/zozoozoo/meta-weekly-report-skill',
    'skills': './skills/',
    'interface': {
        'displayName': 'Meta 周报', 'shortDescription': '清洗Meta源表并按固定写法分析周度变化',
        'longDescription': '生成10张周度清洗表、独立全周期素材汇总和中文周报。解释成本、客单价与转化变化，按累计花费及ROAS筛选低效素材。需要代码执行能力。',
        'developerName': 'zozoozoo', 'category': 'Productivity', 'capabilities': [],
        'defaultPrompt': ['根据我上传的Meta源表，生成清洗Excel和本周周报。']
    }
}
(plugin/'.codex-plugin').mkdir(parents=True, exist_ok=True)
(plugin/'.codex-plugin/plugin.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
out = ROOT/'outputs/chatgpt'
out.mkdir(parents=True,exist_ok=True)
for folder,name,base in [(plugin,'meta-weekly-report-plugin.zip',plugin), (source,'meta-weekly-report-skill.zip',source.parent)]:
    with ZipFile(out/name,'w',ZIP_DEFLATED) as z:
        for path in sorted(folder.rglob('*')):
            if path.is_file() and '__pycache__' not in path.parts and path.suffix!='.pyc':
                z.write(path,path.relative_to(base).as_posix())
    print(out/name)
