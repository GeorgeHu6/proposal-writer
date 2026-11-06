#!/usr/bin/env python3
"""Count narrative Chinese characters and check the sample-derived chapter budgets."""
import argparse
from dataclasses import dataclass, field
import json
from pathlib import Path
import re

CJK = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U0002fa1f]')
PLACEHOLDER = re.compile(r'【[^【】]*】')
COURSE = re.compile(r'^(?:\d+[.、]\s*)?课题\s*(?:\d+|[一二三四五六七八九十]+)')
POLICY_PATH = Path(__file__).resolve().parents[1] / 'references' / 'word-limits.json'


def chinese_count(text):
    return len(CJK.findall(PLACEHOLDER.sub('', str(text))))


def name(text):
    return re.sub(r'\s+', ' ', str(text)).strip().rstrip('：:')


@dataclass(eq=False)
class Node:
    title: str
    level: int
    parent: object = None
    children: list = field(default_factory=list)
    paragraphs: list = field(default_factory=list)
    tables: list = field(default_factory=list)

    @property
    def path(self):
        return ((self.parent.path + '/') if self.parent and self.parent.level else '') + self.title

    @property
    def count(self):
        if re.search(r'参考文献|参考资料|references', self.title, re.I):
            return 0
        return sum(chinese_count(t) for t in self.paragraphs) + sum(c.count for c in self.children)

    @property
    def table_count(self):
        return sum(chinese_count(t) for table in self.tables for t in table) + sum(c.table_count for c in self.children)

    @property
    def placeholder_count(self):
        strings = self.paragraphs + [t for table in self.tables for t in table]
        return sum(len(PLACEHOLDER.findall(t)) for t in strings) + sum(c.placeholder_count for c in self.children)


def tree_from(data):
    root = Node('', 0)
    nodes = []
    for section in data.get('sections', []) + data.get('appendices', []):
        part = Node(name(section['heading']), 1, root)
        root.children.append(part)
        nodes.append(part)
        stack = [part]
        for block in section.get('blocks', []):
            kind = block.get('type')
            if kind == 'heading':
                level = block.get('level', 2)
                while len(stack) > 1 and stack[-1].level >= level:
                    stack.pop()
                node = Node(name(block['text']), level, stack[-1])
                stack[-1].children.append(node)
                stack.append(node)
                nodes.append(node)
            elif kind in ('paragraph', 'bullet') and block.get('role') not in ('caption', 'reference'):
                stack[-1].paragraphs.append(block['text'])
            elif kind == 'table':
                stack[-1].tables.append(block.get('headers', []) + [cell for row in block.get('rows', []) for cell in row])
    return root, nodes


def budget(spec, scale=1.0):
    return {key: max(1, round(spec[key] * scale)) for key in ('target', 'min', 'max')}


def valid_override(value):
    if not isinstance(value, dict) or not isinstance(value.get('target'), int) or isinstance(value['target'], bool) or value['target'] < 1:
        raise ValueError('字数覆盖要求须包含正整数target')
    result = {'target': value['target'], 'min': value.get('min', round(value['target'] * .8)), 'max': value.get('max', round(value['target'] * 1.2))}
    if any(not isinstance(v, int) or isinstance(v, bool) or v < 1 for v in result.values()) or not result['min'] <= result['target'] <= result['max']:
        raise ValueError('字数范围须为正整数，且min ≤ target ≤ max')
    return result


def length_report(data):
    policy = json.loads(POLICY_PATH.read_text(encoding='utf-8'))
    root, nodes = tree_from(data)
    by_path = {n.path: n for n in nodes}
    errors = []
    if data.get('profile', 'major-program') != 'major-program':
        return {'enforced': False, 'narrative_chinese_characters': root.count, 'table_chinese_characters': root.table_count, 'placeholder_spans': root.placeholder_count, 'checks': [], 'errors': []}
    rules = []

    def add(path, spec, quota=True):
        node = by_path.get(path)
        if node is None:
            errors.append(f'字数统计缺少栏目：{path}')
            return
        rules.append({'node': node, 'path': path, **budget(spec), 'quota': quota, 'check': True})

    for spec in policy['fixed']:
        add('/'.join(spec['path']), spec)
    course_parent = next((n for n in nodes if n.title == '（二）各课题内容' and n.level == 3), None)
    courses = [n for n in (course_parent.children if course_parent else []) if n.level == 4 and COURSE.match(n.title)]
    if not courses:
        errors.append('各课题内容须设置可识别的四级课题标题，例如“1. 课题1 具体名称”')
    for course in courses:
        add(course.path, policy['per_course'])
        for spec in policy['course_fields']:
            matches = [n for n in course.children if n.level == 5 and re.search(spec['pattern'], n.title)]
            if len(matches) != 1:
                errors.append(f'课题“{course.title}”须有一个明确的“{spec["label"]}”五级小节')
            else:
                add(matches[0].path, spec, quota=False)
    progress = policy['progress']
    progress_target = progress['base'] + progress['per_course'] * len(courses)
    add(progress['path'], {'target': progress_target, 'min': round(progress_target * .8), 'max': round(progress_target * 1.2)})
    default_total = sum(r['target'] for r in rules if r['quota'])
    requirements = data.get('length_requirements', {})
    if not isinstance(requirements, dict):
        raise ValueError('length_requirements须为对象')
    explicit_total = None
    if 'total' in requirements:
        explicit_total = valid_override(requirements['total'])
        factor = explicit_total['target'] / max(default_total, 1)
        for rule in rules:
            rule.update(budget(rule, factor))
    overrides = requirements.get('sections', {})
    if not isinstance(overrides, dict):
        raise ValueError('length_requirements.sections须为栏目路径到字数要求的映射')
    # Parent requirements are apportioned first; a more specific requirement wins.
    for path, spec in sorted(overrides.items(), key=lambda item: item[0].count('/')):
        node = by_path.get(path)
        if node is None:
            errors.append(f'字数覆盖要求对应的栏目不存在：{path}')
            continue
        limits = valid_override(spec)
        descendants = [r for r in rules if r['path'] == path or r['path'].startswith(path + '/')]
        primary = next((r for r in descendants if r['path'] == path), None)
        current_target = primary['target'] if primary else sum(r['target'] for r in descendants if r['quota'])
        if not current_target:
            errors.append(f'栏目没有可分配的正文统计项：{path}')
            continue
        factor = limits['target'] / current_target
        for rule in descendants:
            rule.update(budget(rule, factor))
        if primary:
            primary.update(limits)
        else:
            rules.append({'node': node, 'path': path, **limits, 'quota': False, 'check': True})
        # A narrower explicit field replaces the generic course-total constraint.
        for rule in rules:
            if path.startswith(rule['path'] + '/') and rule['node'] in courses and rule['path'] not in overrides:
                rule['check'] = False
                fields = [r for r in rules if r['path'].startswith(rule['path'] + '/') and not r['quota']]
                for key in ('target', 'min', 'max'):
                    rule[key] = sum(r[key] for r in fields)
    checks = []
    for rule in rules:
        count = rule['node'].count
        passed = not rule['check'] or rule['min'] <= count <= rule['max']
        checks.append({'path': rule['path'], 'target': rule['target'], 'min': rule['min'], 'max': rule['max'], 'actual': count, 'enforced': rule['check'], 'passed': passed, 'quota': rule['quota']})
        if not passed:
            errors.append(f'“{rule["path"]}”有效正文{count}字，要求{rule["min"]}—{rule["max"]}字，默认目标{rule["target"]}字')
    total_target = sum(r['target'] for r in rules if r['quota'])
    total_bounds = explicit_total or {'target': total_target, 'min': sum(r['min'] for r in rules if r['quota']), 'max': sum(r['max'] for r in rules if r['quota'])}
    if explicit_total and not explicit_total['min'] <= root.count <= explicit_total['max']:
        errors.append(f'全篇有效正文{root.count}字，不在用户要求的{explicit_total["min"]}—{explicit_total["max"]}字内')
    summaries = []
    for node in nodes:
        if node.level <= 2:
            related = [r for r in rules if r['quota'] and (r['path'] == node.path or r['path'].startswith(node.path + '/'))]
            summaries.append({'path': node.path, 'actual': node.count, 'target': sum(r['target'] for r in related), 'table_chinese_characters': node.table_count, 'placeholder_spans': node.placeholder_count})
    return {'enforced': True, 'counting_basis': policy['counting_basis'], 'course_count': len(courses), 'default_total_before_overrides': default_total, 'total': {**total_bounds, 'actual': root.count}, 'table_chinese_characters': root.table_count, 'placeholder_spans': root.placeholder_count, 'summaries': summaries, 'checks': checks, 'errors': errors}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True)
    parser.add_argument('--report', help='Optional working JSON report path')
    args = parser.parse_args()
    data = json.loads(Path(args.input).read_text(encoding='utf-8'))
    report = length_report(data)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        out = Path(args.report)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text + '\n', encoding='utf-8')
    print(text)
    raise SystemExit(1 if report['errors'] else 0)


if __name__ == '__main__':
    main()
