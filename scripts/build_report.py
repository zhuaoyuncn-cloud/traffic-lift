"""将已从当周分表核对的 JSON 写入学院周通报模板。"""
import argparse
import json
import re
from collections import defaultdict
from copy import deepcopy
from datetime import date
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


SECTIONS = ("时政学习", "晚自习", "早操", "宿舍卫生", "教室卫生", "宿舍安全")
MAJORS = ("机械", "机工", "智造", "成型", "型职贯", "机职贯")
TEMPLATE = Path(__file__).resolve().parents[1] / "assets" / "weekly.docx"


def chinese_week(n):
    digits = "零一二三四五六七八九"
    if type(n) is not int or not 1 <= n <= 99:
        raise ValueError("周次应为 1 至 99 的整数")
    tens, units = divmod(n, 10)
    return ((digits[tens] if tens > 1 else "") + "十" if tens else "") + (digits[units] if units else "")


def class_parts(value):
    match = re.fullmatch(r"(\D+)(\d+)", re.sub(r"\s", "", value).removesuffix("班"))
    if not match or match[1] not in MAJORS:
        raise ValueError(f"请核对班级 {value!r}")
    return match.groups()


def class_key(value):
    return tuple((MAJORS.index(m), int(n)) for m, n in map(class_parts, value.split("、")))


def class_label(value):
    return "、".join(("　".join(m) if len(m) == 2 else m) + n
                    for m, n in map(class_parts, value.split("、")))


def day(value):
    if not value:
        return ""
    d = date.fromisoformat(value)
    return f"{d.month}月{d.day}日"


def name_token(name):
    if not isinstance(name, str) or not name.strip():
        raise ValueError("姓名不能为空")
    name = name.strip()
    if re.fullmatch(r"[\u3400-\u9fff]{2}", name):
        name = "　".join(name)
    return name


def write_line(p, tokens):
    # 16 磅正文可用约 25.9 个全角字。留出字体替换和标点压缩的余量。
    # 姓名是完整 token，行末放不下时，在姓名前插入 Word 软换行。
    used = 0
    line = []
    def width_of(text):
        return sum(0.5 if c.isascii() and (c == " " or c.isdigit()) else 1 for c in text)

    for i, (s, atomic) in enumerate(tokens):
        for chunk in ([s] if atomic else s):
            width = width_of(chunk)
            if width > 24:
                raise ValueError(f"单个姓名超过正文行宽，请核对：{s}")
            if used + width > 24 and chunk != "；":
                p.add_run("".join(line).rstrip()).add_break()
                line.clear()
                used = 0
            if not used and chunk == " ":
                continue
            line.append(chunk)
            used += width
        if s == "；":
            next_width = 0
            for j in range(i + 1, len(tokens)):
                if tokens[j][0] == "；":
                    break
                next_width += width_of(tokens[j][0])
            if used + next_width > 24:
                p.add_run("".join(line)).add_break()
                line.clear()
                used = 0
    p.add_run("".join(line))


def time_label(r):
    return day(r.get("date")) or r.get("period", "")


def details(rows, prefix="", include_issue=True, include_date=False, include_room=False):
    tokens = [(prefix, False)]
    for i, r in enumerate(rows):
        fields = [time_label(r) if include_date else "", r.get("room", "") if include_room else "",
                  r.get("bed", "")]
        parts = [(v, False) for v in fields if v]
        parts.extend((name_token(n), True) for n in r.get("names", []))
        if include_issue:
            parts.append((r["issue"], False))
        for j, part in enumerate(parts):
            if i or j:
                tokens.append((" ", False))
            tokens.append(part)
    return tokens


def grouped(rows, key):
    groups = defaultdict(list)
    for r in rows:
        groups[r[key]].append(r)
    return groups


def evening_groups(records, dated):
    # 先按人员归集，再合列问题完全相同的人员。地点、床位或检查批次不同要分开。
    issues = defaultdict(list)
    order = []
    for i, r in enumerate(records):
        location = (r.get("room", ""), r.get("bed", ""), r.get("check", ""))
        if not r.get("names"):
            order.append((None, r))
            continue
        for name in r["names"]:
            key = (location, name, None if dated or r.get("check") else i)
            if key not in issues:
                order.append((key, r))
            issues[key].append(r["issue"])

    groups = []
    by_problems = {}
    for key, r in order:
        if key is None:
            groups.append(r)
            continue
        location, name, batch = key
        problems = tuple(issues[key])
        group_key = (location, batch, problems)
        if group_key not in by_problems:
            group = dict(r, names=[], issue="、".join(problems))
            groups.append(group)
            by_problems[group_key] = group
        by_problems[group_key]["names"].append(name)
    return groups


def separated(clauses, prefix=""):
    tokens = [(prefix, False)]
    for i, clause in enumerate(clauses):
        if i:
            tokens.append(("；", False))
        tokens.extend(clause)
    return tokens


def section_lines(section, rows):
    if not rows:
        raise ValueError(f"{section} 空列表含义不明，已确认无问题请填“无”")
    normalized = []
    shared_findings = set()
    for r in rows:
        if not r["issue"].strip() or not isinstance(r.get("names", []), list):
            raise ValueError(f"{section} 的问题或姓名格式有误")
        day(r.get("date"))
        classes = sorted({"".join(class_parts(c)) for c in r["class"].split("、")}, key=class_key)
        if section in ("宿舍卫生", "宿舍安全") and len(classes) > 1:
            finding = (tuple(classes), r["room"], r.get("date"), r.get("period"), r.get("check"),
                       r.get("bed"), tuple(sorted(r.get("names", []))), r["issue"])
            if finding in shared_findings:
                continue
            shared_findings.add(finding)
        normalized.append(dict(r, **{"class": "、".join(classes)}))
    rows = sorted(normalized, key=lambda r: class_key(r["class"]))
    if section == "晚自习":
        batches = defaultdict(list)
        for r in rows:
            dt = r.get("date") or ""
            batches[(dt, "" if dt else r.get("period", ""))].append(r)
        for (dt, period), items in sorted(batches.items(), key=lambda pair: (not pair[0][0], pair[0])):
            heading = day(dt) or period
            if heading or len(batches) > 1:
                yield [((heading or "具体日期未注明") + "：", False)]
            for cls, records in grouped(items, "class").items():
                clauses = (details([g], include_room=True) for g in evening_groups(records, bool(dt)))
                yield separated(clauses, class_label(cls) + "：")
        return
    if section == "时政学习":
        tokens = []
        for cls, items in grouped(rows, "class").items():
            if tokens:
                tokens.append(("；", False))
            tokens.extend(separated((details([r], include_date=True, include_room=True) for r in items),
                                    class_label(cls) + "："))
        yield tokens
        return
    if section in ("宿舍卫生", "宿舍安全", "教室卫生"):
        for cls, items in grouped(rows, "class").items():
            yield [(class_label(cls) + "：", False)]
            for room, records in grouped(items, "room").items():
                if not room.strip():
                    raise ValueError(f"{section} 缺少地点")
                if section == "宿舍卫生":
                    # 每个子句自带床位及责任人，不跨检查合并床位。
                    clauses = (details([r], include_date=True) for r in records)
                    yield separated(clauses, room + "：")
                else:
                    yield [(room + "：", False)]
                    for r in records:
                        yield details([r], include_date=True)
        return
    by_issue = grouped(rows, "issue")
    for issue in sorted(by_issue, key=lambda s: {"旷操": 0, "迟到": 1}.get(s, 2)):
        yield [(issue + "：", False)]
        for cls, records in grouped(by_issue[issue], "class").items():
            label = class_label(cls) + "："
            if all(r.get("names") and not r.get("room") and not r.get("bed") and not time_label(r) for r in records):
                yield details(records, label, include_issue=False)
            else:
                for r in records:
                    yield details([r], label, include_issue=False, include_date=True, include_room=True)


def build(data, folder):
    week = chinese_week(data["week"])
    grade = str(data["grade"])
    if not re.fullmatch(r"\d{2}|\d{4}", grade):
        raise ValueError("年级应为两位或四位阿拉伯数字")
    issued = date.fromisoformat(data["issued"])
    if set(data["sections"]) != set(SECTIONS):
        raise ValueError("六个板块必须齐全；缺表请先核实")
    doc = Document(TEMPLATE)
    _, title, _, _, head, body, *_, signature_date = doc.paragraphs
    title.runs[0].text = f"第{week}周通报"
    signature_date.runs[0].text = f"{issued.year}年{issued.month}月{issued.day}日"

    def insert(prototype, tokens, keep=False):
        element = deepcopy(prototype._p)
        element.attrib.pop(qn("w14:paraId"), None)
        head._p.addprevious(element)
        p = Paragraph(element, head._parent)
        p.clear()
        p.paragraph_format.keep_with_next = keep
        write_line(p, tokens)
        for run in p.runs:
            run._r.insert(0, deepcopy(prototype.runs[0]._r.rPr))

    section_order = data.get("section_order", SECTIONS)
    if len(section_order) != len(SECTIONS) or set(section_order) != set(SECTIONS):
        raise ValueError("版式样表的板块顺序须完整列出六项，每项一次")
    for section in section_order:
        rows = data["sections"][section]
        if rows == "无":
            insert(head, [(section + "：无", False)])
            continue
        insert(head, [(section + "：", False)], True)
        for tokens in section_lines(section, rows):
            insert(body, tokens, len(tokens) == 1 and tokens[0][0].endswith("："))
    for p in (head, body):
        p._p.getparent().remove(p._p)
    output = folder / f"工程机械学院{grade}级第{week}周通报.docx"
    with output.open("xb") as f:
        doc.save(f)
    return output


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("output_directory", type=Path)
    args = parser.parse_args()
    print(build(json.loads(args.input.read_text(encoding="utf-8-sig")), args.output_directory))
