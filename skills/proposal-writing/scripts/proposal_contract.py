#!/usr/bin/env python3
"""Check the requested proposal outline, heading levels and explicit wording rules."""
import argparse
import json
import re
from pathlib import Path
from formal_style import style_findings

PARTS = [
    "第一部分 国内外现状及趋势分析",
    "第二部分 研究目标及内容",
    "第四部分 进度安排",
    "第五部分 项目组织实施、保障措施及风险分析",
]
RESEARCH_CHAPTERS = [
    "一、项目目标及考核指标",
    "二、项目研究内容、研究方法及技术路线",
    "三、课题分解方案",
    "四、主要创新点",
    "五、预期经济社会效益",
]
RESEARCH_SECTIONS = {
    RESEARCH_CHAPTERS[0]: [
        "（一）申报项目与所属指南方向的关联关系",
        "（二）项目目标及考核指标、考核方式/方法",
        "（三）项目成果的呈现形式及描述",
    ],
    RESEARCH_CHAPTERS[1]: ["（一）项目的主要研究内容", "（二）项目拟采取的研究方法"],
    RESEARCH_CHAPTERS[2]: ["（一）课题分解情况", "（二）各课题内容"],
}
MANAGEMENT_CHAPTERS = [
    "一、项目组织实施机制",
    "二、保障措施",
    "三、知识产权对策、成果管理及合作权益分配",
    "四、风险分析及对策",
]


def normalized(text):
    return re.sub(r"\s+", " ", str(text)).strip().rstrip("：:")


def wording_errors(text, location):
    return [f"{location}：{item['rule']}“{item['matched']}”；{item['rewrite']}" for item in style_findings(text, location)]


def content_strings(data):
    for key in ("title", "subtitle", "status"):
        if isinstance(data.get(key), str):
            yield data[key], key
    for section_index, section in enumerate(data.get("sections", []) + data.get("appendices", []), 1):
        if not isinstance(section, dict):
            continue
        yield section.get("heading", ""), f"部分{section_index}标题"
        for index, block in enumerate(section.get("blocks", []), 1):
            if not isinstance(block, dict):
                continue
            location = f"{section.get('heading', section_index)} 第{index}项"
            for key in ("text", "caption"):
                if isinstance(block.get(key), str):
                    yield block[key], location
            if block.get("type") == "table":
                for header in block.get("headers", []):
                    if isinstance(header, str):
                        yield header, location
                for row in block.get("rows", []):
                    if isinstance(row, list):
                        for cell in row:
                            if isinstance(cell, str):
                                yield cell, location


def contract_errors(data):
    errors = []
    if not isinstance(data, dict):
        return ["输入须为JSON对象"]
    profile = data.get("profile", "major-program")
    if profile not in ("major-program", "custom"):
        errors.append("profile须为major-program或用户明确要求的custom")
    sections = data.get("sections", [])
    if not isinstance(sections, list):
        return ["sections须为列表"]
    if profile == "major-program":
        titles = [normalized(s.get("heading", "")) if isinstance(s, dict) else "" for s in sections]
        if titles != PARTS:
            errors.append("完整正文顶层须依次为第一、第二、第四、第五部分的指定标题，第三部分省略且不得重编号")
        if data.get("appendices"):
            errors.append("默认正式正文不另加核查附录或其他顶层栏目")
        if data.get("status"):
            errors.append("正式正文不输出草稿性质等状态说明行")
        for part_index, expected in ((1, RESEARCH_CHAPTERS), (3, MANAGEMENT_CHAPTERS)):
            if len(sections) <= part_index or not isinstance(sections[part_index], dict):
                continue
            blocks = sections[part_index].get("blocks", [])
            actual = [normalized(b.get("text", "")) for b in blocks if isinstance(b, dict) and b.get("type") == "heading" and b.get("level", 2) == 2]
            if actual != expected:
                errors.append(f"{PARTS[part_index]}的二级章标题应按指定名称和顺序完整保留")
        if len(sections) > 1 and isinstance(sections[1], dict):
            children = {}
            parent = None
            for block in sections[1].get("blocks", []):
                if not isinstance(block, dict) or block.get("type") != "heading":
                    continue
                level = block.get("level", 2)
                if level == 2:
                    parent = normalized(block.get("text", ""))
                elif level == 3:
                    children.setdefault(parent, []).append(normalized(block.get("text", "")))
            for parent, expected in RESEARCH_SECTIONS.items():
                if children.get(parent, []) != expected:
                    errors.append(f"“{parent}”的三级节标题应完整保留并放在正确父标题下")
    for section in sections:
        if not isinstance(section, dict):
            continue
        previous = 1
        blocks = section.get("blocks", [])
        for index, block in enumerate(blocks):
            if not isinstance(block, dict) or block.get("type") != "heading":
                continue
            level = block.get("level", 2)
            if not isinstance(level, int) or isinstance(level, bool) or not 2 <= level <= 6:
                errors.append(f"标题“{block.get('text', '')}”须使用2—6级标题")
                continue
            if level > previous + 1:
                errors.append(f"标题“{block.get('text', '')}”从{previous}级跳到{level}级，请补足父标题")
            previous = level
            substantive = False
            for child in blocks[index + 1:]:
                if not isinstance(child, dict):
                    continue
                if child.get("type") == "heading" and child.get("level", 2) <= level:
                    break
                if (child.get("type") in ("paragraph", "bullet") and str(child.get("text", "")).strip()) or (child.get("type") == "table" and child.get("rows")):
                    substantive = True
            if not substantive:
                errors.append(f"标题“{block.get('text', '')}”下缺少正文内容")
    for text, location in content_strings(data):
        errors.extend(wording_errors(text, location))
    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True)
    args = parser.parse_args()
    data = json.loads(Path(args.input).read_text(encoding="utf-8"))
    errors = contract_errors(data)
    print(json.dumps({"passed": not errors, "errors": errors}, ensure_ascii=False, indent=2))
    raise SystemExit(1 if errors else 0)


if __name__ == "__main__":
    main()
