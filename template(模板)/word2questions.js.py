#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word(.docx) -> questions_data.js 转换脚本
- 读取 Word 文档中的题目
- 解析题型（单选/多选/判断）、选项、答案
- 导出为与前端题库一致的 JS 文件：window.questionsData = [...];

示例输入（Word 段落，题与题之间可以空一行）:
1.（单选）劳动的双重含义决定了从业人员全新的（ ）和职业道德观念。
A. 精神文明
B. 思想境界
C. 劳动态度
D. 整体素质
答案：C

2、(多选) 下列哪几项是……
A. 选项1
B. 选项2
C. 选项3
D. 选项4
答案：A、C、D

3) （判断）关于……的说法正确的是：
答案：对        # 支持：对/错、正确/错误、√/×、T/F/True/False/Yes/No

依赖：python-docx
pip install python-docx
"""
import argparse
import json
import re
import sys
from pathlib import Path
from typing import List, Dict, Any

try:
    from docx import Document  # python-docx
except Exception as e:
    print("缺少依赖 python-docx，请先安装：pip install python-docx", file=sys.stderr)
    raise

# 正则模式
RE_Q_START = re.compile(
    r"""^\s*
        (?P<num>\d+)\s*[\.\、\)]\s*      # 题号
        (?:[（(]?\s*(?P<typedesc>单选|多选|判断)\s*[)）])?\s* # 可选题型(在题号后)
        (?P<qtext>.*)                    # 题干剩余
    $""",
    re.X,
)

RE_Q_TYPE_AT_END = re.compile(r"(?P<body>.*?)[（(]\s*(单选|多选|判断)\s*[)）]\s*$")
RE_OPT = re.compile(r"^\s*([A-Za-z])\s*[\.\、\)]\s*(.*\S)\s*$")
RE_ANS = re.compile(r"^\s*(?:答案|正确答案)\s*[:：]\s*(.+?)\s*$")

TRUE_SET = {"对", "正确", "√", "T", "TRUE", "YES", "Y", "是"}
FALSE_SET = {"错", "错误", "×", "F", "FALSE", "NO", "N", "否"}

def clean(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()

def extract_type_from_text(text: str) -> (str, str):
    """
    如果题型写在题干末尾 (……(单选))，从末尾提取题型并返回 (type, body_without_type)
    否则返回 (None, 原文)
    """
    m = RE_Q_TYPE_AT_END.match(text)
    if m:
        body, typ = m.group("body"), m.group(2)
        return typ, clean(body)
    return None, clean(text)

def normalize_answers(ans_text: str, q_type: str) -> List[str]:
    ans_text = ans_text.strip().upper()
    # 多个答案分隔符：中文逗号、顿号、英文逗号、空格、加号等
    parts = re.split(r"[、，,;；\s\+]+", ans_text)
    parts = [p for p in parts if p]

    if q_type == "判断":
        # 判断题：固定 A=正确 B=错误
        token = parts[0]
        if token in TRUE_SET or token.upper() in TRUE_SET:
            return ["A"]
        if token in FALSE_SET or token.upper() in FALSE_SET:
            return ["B"]
        # 也支持直接写 A/B
        if token in {"A", "B"}:
            return [token]
        # 兜底：无法识别时，返回空
        return []
    else:
        # 单选/多选：只取字母
        letters = []
        for p in parts:
            m = re.match(r"([A-Z])", p)
            if m:
                letters.append(m.group(1))
        # 去重且保持原顺序
        seen = set()
        uniq = []
        for ch in letters:
            if ch not in seen:
                seen.add(ch)
                uniq.append(ch)
        return uniq

def flush_current(q_list: List[Dict[str, Any]], buf: Dict[str, Any]):
    if not buf:
        return
    # 判断题如果用户没有给选项，自动补 “正确/错误”
    if buf.get("type") == "判断" and not buf.get("options"):
        buf["options"] = [
            {"label": "A", "text": "正确"},
            {"label": "B", "text": "错误"},
        ]
    # 选项去重/清洗
    seen = set()
    cleaned_opts = []
    for opt in buf.get("options", []):
        label = opt.get("label", "").upper()
        text = clean(opt.get("text", ""))
        if not label or not text or label in seen:
            continue
        seen.add(label)
        cleaned_opts.append({"label": label, "text": text})
    buf["options"] = cleaned_opts

    # 答案规范为列表（可能为空）
    ans = buf.get("answer", [])
    if isinstance(ans, str):
        ans = [ans] if ans else []
    buf["answer"] = ans

    # 最终检查：题干必须存在
    if clean(buf.get("question", "")):
        q_list.append(buf.copy())

def parse_docx(docx_path: Path, start_number: int = 1, respect_word_number: bool = False) -> List[Dict[str, Any]]:
    doc = Document(str(docx_path))
    lines: List[str] = []
    for p in doc.paragraphs:
        t = p.text.replace("\u3000", " ").strip()  # 去全角空格
        if t:
            lines.append(t)
        else:
            lines.append("")  # 保留空行分段

    questions: List[Dict[str, Any]] = []
    cur = {}
    auto_num = start_number

    def set_number(n: int):
        cur["number"] = n

    for raw in lines + [""]:  # 末尾补空行，便于 flush
        line = raw.strip()

        # 识别“题起始”
        m = RE_Q_START.match(line)
        if m:
            # 如果已有缓冲题，先收录
            flush_current(questions, cur)
            cur = {"options": [], "answer": []}

            num_in_doc = int(m.group("num"))
            typedesc = m.group("typedesc")
            qtext = m.group("qtext").strip()

            # 题型有两种写法：紧跟题号，或写在题干末尾，这里两处都尝试。
            typ2, qtext2 = extract_type_from_text(qtext)
            q_type = typedesc or typ2 or "单选"  # 默认为单选

            cur["type"] = q_type
            cur["question"] = clean(qtext2 if typ2 else qtext)

            if respect_word_number:
                set_number(num_in_doc)
                # 自动编号也跟上，避免后续无题号题目继续编号错位
                auto_num = num_in_doc + 1
            else:
                set_number(auto_num)
                auto_num += 1
            continue

        # 选项
        m2 = RE_OPT.match(line)
        if m2 and cur:
            label = m2.group(1).upper()
            text = m2.group(2)
            cur.setdefault("options", []).append({"label": label, "text": clean(text)})
            continue

        # 答案
        m3 = RE_ANS.match(line)
        if m3 and cur:
            ans_text = m3.group(1)
            cur["answer"] = normalize_answers(ans_text, cur.get("type", "单选"))
            continue

        # 普通文本：当作题干追加（换行拼接）
        if cur and line:
            # 如果行尾带题型再提取一次（兼容“题干最后标注(单选)”）
            typ2, body2 = extract_type_from_text(line)
            if typ2:
                cur["type"] = typ2
                line = body2
            cur["question"] = clean((cur.get("question", "") + " " + line).strip())

        # 空行：认为一个题块结束点之一，但不要强制 flush（答案可能在后面几行）
        # 这里不在空行处 flush，统一等下一题或文档结束时 flush

    # 结束时 flush
    flush_current(questions, cur)
    return questions

def merge_existing(js_path: Path, new_list: List[Dict[str, Any]], renumber_after_merge: bool = False, start_number: int = 1) -> List[Dict[str, Any]]:
    if not js_path.exists():
        return new_list

    text = js_path.read_text(encoding="utf-8")
    m = re.search(r"window\.questionsData\s*=\s*(\[[\s\S]*?\])\s*;", text)
    if not m:
        # 不是预期格式，直接追加
        return new_list

    try:
        old_list = json.loads(m.group(1))
    except Exception:
        old_list = []

    merged = list(old_list) + list(new_list)

    if renumber_after_merge:
        n = start_number
        for item in merged:
            item["number"] = n
            n += 1
    return merged

def write_js(js_path: Path, q_list: List[Dict[str, Any]]):
    js_path.parent.mkdir(parents=True, exist_ok=True)
    js = json.dumps(q_list, ensure_ascii=False, indent=2)
    js_text = "window.questionsData = " + js + ";\n"
    js_path.write_text(js_text, encoding="utf-8")

def main():
    ap = argparse.ArgumentParser(
        description="将 Word(.docx) 题库转换为前端使用的 questions_data.js"
    )
    ap.add_argument("input", help="输入 .docx 文件路径")
    ap.add_argument("-o", "--output", default="questions_data.js", help="输出 .js 文件路径（默认：questions_data.js）")
    ap.add_argument("--start-number", type=int, default=1, help="题号起始值（默认：1）")
    ap.add_argument("--respect-number", action="store_true", help="优先使用 Word 内的题号（默认否）")
    ap.add_argument("--append-to", help="将结果追加合并到现有 JS（如：./questions_data.js）")
    ap.add_argument("--renumber-after-merge", action="store_true", help="合并后按顺序重新编号（从 --start-number 开始）")

    args = ap.parse_args()
    in_path = Path(args.input)
    out_path = Path(args.output)

    if not in_path.exists():
        print(f"未找到输入文件：{in_path}", file=sys.stderr)
        sys.exit(2)

    try:
        q_list = parse_docx(in_path, start_number=args.start_number, respect_word_number=args.respect_number)
    except Exception as e:
        print(f"解析失败：{e}", file=sys.stderr)
        sys.exit(3)

    if args.append_to:
        merged = merge_existing(Path(args.append_to), q_list, renumber_after_merge=args.renumber_after_merge, start_number=args.start_number)
        write_js(out_path, merged)
    else:
        write_js(out_path, q_list)

    print(f"✅ 已生成：{out_path}（共 {len(q_list)} 题）")

if __name__ == "__main__":
    main()
