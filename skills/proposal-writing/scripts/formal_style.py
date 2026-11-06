#!/usr/bin/env python3
"""Check direct proposal prose and bracketed filling guidance, including final DOCX text."""
import argparse
import json
from pathlib import Path
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

VERSION = '1.4.0'
PLACEHOLDER = re.compile(r'【([^【】]*)】')
PREFIXES = ('请填写', '请补充', '请提供', '请根据', '请依据', '请核实', '请确认')
RULES = [
    ('参考性套话', r'(?:提供|作为|用作|仅供|只供|可供|以供|供|作).{0,6}参考|(?:仅|只)(?:具有|有|起|承担|提供|作|作为|用于|能作|能提供|可作|可供).{0,8}(?:辅助|示例|借鉴|提示)', '直接说明研究成果的用途、应用环节和作用机制'),
    ('限制性框架', r'(?:仅|只)(?:提供|用于|适用于|针对|覆盖|接收|接受|作为|说明|反映|表明|支持|代表|意味着)|(?:仅|只)限于|(?:作用|用途)限于', '用适用范围、功能清单或输入输出直接说明项目安排'),
    ('否定式指令', r'不得|不将|不把|不能|不应|不宜|不必|无需|不以|不作|不做|不作为|不用于|不依赖|不可以|不予|并非|并不|而非|不是', '按研究对象、实施动作、适用条件和技术效果重新正面组织句子'),
    ('免责与推定套话', r'不(?:构成|代表|等于|等同|意味着|表明|表示|保证|承诺|视为|据此推定)|不可(?:直接|据此|由此|依此|作为|用于|以此)|无法(?:直接|据此|保证|证明)|(?:难以|不足以|尚不足以|仍不足以).{0,10}(?:替代|推定|证明|判断|认定|视为)|以(?:最终|实际|有关部门).{0,12}为准|(?:最终|具体).{0,12}以.{0,18}为准|不承担.{0,8}责任|不负责|概不负责|免责声明', '写明项目采用的判断依据、测评程序、使用条件或具体填写提示'),
    ('自述与提醒', r'作为(?:AI助手|人工智能助手|语言模型)|本回答|本次答复|本次生成|本次写作|以下(?:内容|文字|表述|段落)|本(?:文|稿|段|节|章|部分)(?:旨在|用于|主要说明|主要介绍|将|先|仅|只|说明|介绍|阐述)|需要(?:说明|指出|强调|提醒|注意)的是|需(?:要)?注意|特别提醒|温馨提示|请注意|说明一下|需要明确的是|(?:这里|此处|本段)(?:需要|需|应|可|先|暂)|^(?:说明|提示|注意|备注|填写建议)[：:]', '直接进入项目的研究内容，写作说明留在工作记录'),
    ('括号外填写建议', r'请(?:填写|补充|提供|核实|确认|根据|依据|结合|自行)|建议(?:用户|申报单位|相关方|项目组|在|根据|结合|另行|自行|补充|填写|核实|确认|完善)|(?:建议|需要|应当|应|须|需|后续|另行|待)(?:根据|结合|依据|按|由)?(?:本项目|项目|具体|实际|合作|申报|所提供|现场)?.{0,16}(?:填写|补充|核实|确认|完善|替换)|(?:根据|结合)实际(?:情况|条件)(?:填写|确定|补充|完善)|(?:在|于)(?:实际)?(?:申报|填报|提交|定稿)(?:时|前|后)', '将缺失信息及填写建议完整写入同一组【】中'),
    ('悬置性说明', r'建议(?:值|指标|目标|周期|采用|使用|选用)|暂定|待(?:论证|确认|核实|补充|完善)|名称驱动初稿|仅供参考|不构成正式承诺', '已明确的研究方案直接陈述，实际资料缺口使用具体的【请填写……】提示'),
    ('内部管理标签', r'冻结(?:本项目)?(?:指标|目标|口径|方案|版本|任务|承诺)|(?:指标|目标|口径|方案|版本|任务|承诺)(?:已|先行|进行|予以)?冻结|锁定版本|交付门禁|任务契约', '使用具体的研究目标、指标、任务或实施程序名称'),
    ('抽象字母标签', r'(?<![A-Za-z0-9])(?:P|G|K|O)\d+\s*(?:[/／、]|[：:]|对应|映射|支撑)|(?<![A-Za-z0-9])(?:P|G|K|O)\d+(?:\s+|[/／、]\s*)(?:P|G|K|O)\d+', '使用具体中文名称或正常的课题、指标编号'),
    ('空泛占位', r'\bXXX\b', '使用说明应填事实和口径的【请填写……】提示'),
]
COMPILED = [(label, re.compile(pattern), suggestion) for label, pattern, suggestion in RULES]


def style_findings(text, location='正文'):
    text = str(text)
    findings = []
    spans = list(PLACEHOLDER.finditer(text))
    if text.count('【') != len(spans) or text.count('】') != len(spans):
        findings.append({'location': location, 'rule': '填写提示配对', 'matched': '【】', 'rewrite': '将具体信息和填写建议完整置于同一组中文方括号中'})
    for span in spans:
        content = span.group(1).strip()
        if not content.startswith(PREFIXES) or len(content) < 6:
            findings.append({'location': location, 'rule': '填写提示内容', 'matched': span.group(0), 'rewrite': '在括号内说明需填的具体事实、资料或计算口径'})
        if re.search(r'仅供参考|不构成(?:正式)?承诺|不能保证|免责声明|名称驱动初稿', content):
            findings.append({'location': location, 'rule': '括号承载内容', 'matched': span.group(0), 'rewrite': '括号承载实际资料缺口和填写方法；将该说明改写为具体填写项'})
    # Keep valid filling advice separate from the surrounding formal prose.
    outside = PLACEHOLDER.sub(' ', text)
    compact = re.sub(r'[\s\u200b-\u200f\ufeff]+', '', outside)
    for label, expression, suggestion in COMPILED:
        searched = outside if label in ('抽象字母标签', '空泛占位') else compact
        match = expression.search(searched)
        if match:
            findings.append({'location': location, 'rule': label, 'matched': match.group(0), 'rewrite': suggestion})
    return findings


def scan_docx(path):
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    findings = []
    units = 0
    with ZipFile(path) as archive:
        parts = [p for p in archive.namelist() if re.fullmatch(r'word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml', p)]
        if 'word/document.xml' not in parts:
            raise ValueError('DOCX缺少正文XML')
        for part in parts:
            root = ET.fromstring(archive.read(part))
            for index, paragraph in enumerate(root.findall('.//w:p', namespace), 1):
                text = ''.join(t.text or '' for t in paragraph.findall('.//w:t', namespace))
                if text.strip():
                    units += 1
                    findings.extend(style_findings(text, f'{part} 段落{index}'))
    return {'ruleset': VERSION, 'passed': not findings, 'checked_text_units': units, 'findings': findings}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--docx', help='Read back actual final Word text')
    group.add_argument('--text-file', help='Check a UTF-8 prose draft')
    parser.add_argument('--report', help='Optional working JSON report')
    args = parser.parse_args()
    if args.docx:
        result = scan_docx(args.docx)
    else:
        findings = style_findings(Path(args.text_file).read_text(encoding='utf-8'))
        result = {'ruleset': VERSION, 'passed': not findings, 'findings': findings}
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        report = Path(args.report)
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(output + '\n', encoding='utf-8')
    print(output)
    raise SystemExit(0 if result['passed'] else 1)


if __name__ == '__main__':
    main()
