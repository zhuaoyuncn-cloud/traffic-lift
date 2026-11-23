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
    for s, atomic in tokens:
        for chunk in ([s] if atomic else s):
            width = sum(0.5 if c.isascii() and (c == " " or c.isdigit())
                        else 1 for c in chunk)
            if width > 24:
                raise ValueError(f"单个姓名超过正文行宽，请核对：{s}")
            if used + width > 24:
                p.add_run("".join(line).rstrip()).add_break()
                line.clear()
                used = 0
            if not used and chunk == " ":
                continue
            line.append(chunk)
            used += width
    p.add_run("".join(line))


def details(rows, prefix="", include_issue=True, include_date=False):
    tokens = [(prefix, False)]
    for i, r in enumerate(rows):
        fields = [day(r.get("date")) if include_date else "", r.get("bed", "")]
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


def section_lines(section, rows, problem_groups):
    if rows == []:
        raise ValueError(f"{section} 空列表含义不明，已确认无问题请填“无”")
    for r in rows:
        if not r["issue"].strip() or not isinstance(r.get("names", []), list):
            raise ValueError(f"{section} 的问题或姓名格式有误")
        if section == "晚自习" and not r.get("date"):
            raise ValueError("晚自习记录缺少日期")
        day(r.get("date"))
    rows = sorted(rows, key=lambda r: class_key(r["class"]))
    if section == "时政学习":
        tokens = []
        for cls, items in grouped(rows, "class").items():
            if tokens:
                tokens.append(("；", False))
            tokens.extend(details(items, class_label(cls) + "：", include_date=True))
        yield tokens
        return
    if section in ("宿舍卫生", "宿舍安全", "教室卫生"):
        for cls, items in grouped(rows, "class").items():
            for room, records in grouped(items, "room").items():
                if not room.strip():
                    raise ValueError(f"{section} 缺少地点")
                prefix = class_label(cls) + "：" + room
                if section == "宿舍卫生":
                    for r in records:
                        yield details([r], prefix + " ", include_date=True)
                else:
                    yield [(prefix, False)]
                    for r in records:
                        yield details([r], include_date=True)
        return
    batches = sorted(grouped(rows, "date").items()) if section == "晚自习" else [("", rows)]
    for dt, items in batches:
        if dt:
            yield [(day(dt) + "：", False)]
        by_issue = grouped(items, "issue")
        selected = (sorted(by_issue, key=lambda s: {"旷操": 0, "迟到": 1}.get(s, 2))
                    if section == "早操" else [s for s in problem_groups if s in by_issue])
        for issue in selected:
            yield [(issue + "：", False)]
            for cls, records in grouped(by_issue[issue], "class").items():
                label = class_label(cls) + "："
                if section == "晚自习":
                    yield [(label, False)]
                    label = ""
                # 日期、床位等细节属于各自记录；仅纯名单可拼在同一行。
                if all(r.get("names") and not r.get("bed") and (dt or not r.get("date")) for r in records):
                    yield details(records, label, include_issue=False)
                else:
                    for r in records:
                        yield details([r], label, include_issue=False, include_date=not dt)
        for r in items:
            if r["issue"] not in selected:
                yield details([r], class_label(r["class"]) + "：")


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

    for section in SECTIONS:
        rows = data["sections"][section]
        if rows == "无":
            insert(head, [(section + "：无", False)])
            continue
        insert(head, [(section + "：", False)], True)
        for tokens in section_lines(section, rows, data.get("problem_groups", [])):
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
