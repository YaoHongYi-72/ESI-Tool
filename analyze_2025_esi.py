from __future__ import annotations

import argparse
import html
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

import pandas as pd
from docx import Document


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_BASE = PROJECT_ROOT / "input_data"
DEFAULT_OUT = PROJECT_ROOT / "analysis"
DEFAULT_TEMPLATE_2024 = PROJECT_ROOT / "templates" / "我校 2024年度 ESI 高水平论文的产出情况及影响力分析.docx"

BASE = DEFAULT_BASE
OUT = DEFAULT_OUT
SOURCES: list[tuple[int, Path, str]] = []

SUBJECT_MAP = {
    "AGRICULTURAL SCIENCES": "农业科学",
    "BIOLOGY & BIOCHEMISTRY": "生物与生物化学",
    "CHEMISTRY": "化学",
    "CLINICAL MEDICINE": "临床医学",
    "COMPUTER SCIENCE": "计算机科学",
    "ECONOMICS & BUSINESS": "经济学与商学",
    "ENGINEERING": "工程",
    "ENVIRONMENT/ECOLOGY": "环境/生态学",
    "MATERIALS SCIENCE": "材料科学",
    "MATHEMATICS": "数学",
    "PHARMACOLOGY & TOXICOLOGY": "药理学和毒理学",
    "PLANT & ANIMAL SCIENCE": "植物与动物科学",
    "PSYCHIATRY/PSYCHOLOGY": "精神病学/心理学",
    "SOCIAL SCIENCES, GENERAL": "一般社会科学",
}

SUBJECT_ORDER = [
    "农业科学",
    "生物与生物化学",
    "化学",
    "临床医学",
    "计算机科学",
    "经济学与商学",
    "工程",
    "环境/生态学",
    "材料科学",
    "数学",
    "精神病学/心理学",
    "药理学和毒理学",
    "植物与动物科学",
    "一般社会科学",
]

COLLEGE_ALIASES = {
    "计算机与人工智能学院（原信息工程学院）": "计算机与人工智能学院",
    "计算机与人工智能学院(原信息工程学院)": "计算机与人工智能学院",
    "财税与税务学院": "财政与税务学院",
    "财政与税务学院": "财政与税务学院",
    "财政与税务": "财政与税务学院",
    "粮食和物资学院": "粮食和物资学院",
    "粮食经济研究院": "粮食经济研究院",
    "红山学院": "红山学院",
    "公共管理学院": "公共管理学院",
    "国际经贸学院": "国际经贸学院",
    "食品科学与工程学院": "食品科学与工程学院",
    "金融学院": "金融学院",
    "会计学院": "会计学院",
    "应用数学学院": "应用数学学院",
    "经济学院": "经济学院",
    "校部": "校部",
    "科学研究院": "科学研究院",
    "继续教育学院": "继续教育学院",
    "工商管理学院": "工商管理学院",
    "工商管理": "工商管理学院",
}

COLLEGE_ORDER = [
    "经济学院",
    "财政与税务学院",
    "金融学院",
    "国际经贸学院",
    "会计学院",
    "公共管理学院",
    "食品科学与工程学院",
    "计算机与人工智能学院",
    "应用数学学院",
    "红山学院",
    "粮食和物资学院",
    "粮食经济研究院",
    "校部",
    "科学研究院",
    "继续教育学院",
    "工商管理学院",
]

AUTHOR_NAME_STOPWORDS = {
    "食工",
    "已毕业的",
    "博士",
    "学生",
}

AUTHOR_ALIAS_MANUAL = {
    "wupeipei": "吴佩佩",
    "lirongrong": "李荣荣",
    "chenafei": "陈阿飞",
}

MANUAL_SUBJECT_OVERRIDE_RAW = {
    "CHARACTERIZATION AND FUNCTIONAL EVALUATION OF OAT PROTEIN ISOLATE-PLEUROTUS OSTREATUS BETA-GLUCAN CONJUGATES FORMED VIA MAILLARD REACTION": ["农业科学"],
    "EFFECTS OF TOCOPHEROL NANOEMULSION ADDITION ON FISH SAUSAGE PROPERTIES AND FATTY ACID OXIDATION": ["农业科学"],
    "PROTECTIVE EFFECTS OF MICROBIOME-DERIVED INOSINE ON LIPOPOLYSACCHARIDE-INDUCED ACUTE LIVER DAMAGE AND INFLAMMATION IN MICE VIA MEDIATING THE TLR4/NF-KAPPA B PATHWAY": ["药理学和毒理学"],
    "THE IMPACT OF NATURAL RESOURCE ENDOWMENT AND GREEN FINANCE ON GREEN ECONOMIC EFFICIENCY IN THE CONTEXT OF COP26": ["环境/生态学"],
    "CORRELATIONS BETWEEN THE CRUDE OIL MARKET AND CAPITAL MARKETS UNDER THE RUSSIA-UKRAINE CONFLICT: A PERSPECTIVE OF CRUDE OIL IMPORTING AND EXPORTING COUNTRIES": ["经济学与商学"],
    "DOES THE DEVELOPMENT OF DIGITAL INCLUSIVE FINANCE IMPROVE THE ENTHUSIASM AND QUALITY OF CORPORATE GREEN TECHNOLOGY INNOVATION?": ["经济学与商学"],
    "UNVEILING THE POLICY INTERVENTION EFFECTS OF NATURAL RESOURCE REGULATION ON FIRM-LEVEL POLLUTION EMISSIONS: EVIDENCE FROM CHINAS RESTRAINED LAND SUPPLY": ["环境/生态学"],
    "INSTITUTIONAL INVESTOR SHAREHOLDING AND THE QUALITY OF CORPORATE INNOVATION: MODERATING EFFECTS BASED ON INTERNAL AND EXTERNAL ENVIRONMENT": ["经济学与商学"],
}

MANUAL_PENDING_AUTHOR_OVERRIDES_RAW = {
    "GOVERNMENT INVESTMENT, HUMAN CAPITAL FLOW, AND URBAN INNOVATION: EVIDENCE FROM SMART CITY CONSTRUCTION IN CHINA": [
        "Li,Zhihui",
        "Wang,Xinyue",
    ],
    "AN INTELLIGENT THYMOL/ALIZARIN-LOADED POLYCAPROLACTONE/GELATIN/ZEIN NANOFIBROUS FILM WITH PH-RESPONSIVE AND ANTIBACTERIAL PROPERTIES FOR SHRIMP FRESHNESS MONITORING AND PRESERVATION": [
        "Deng,Jing",
    ],
}

GREEK_TITLE_MAP = {
    "Α": "ALPHA",
    "α": "ALPHA",
    "Β": "BETA",
    "β": "BETA",
    "Γ": "GAMMA",
    "γ": "GAMMA",
    "Δ": "DELTA",
    "δ": "DELTA",
    "Κ": "KAPPA",
    "κ": "KAPPA",
    "Ω": "OMEGA",
    "ω": "OMEGA",
}


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text in {"nan", "None"}:
        return ""
    return text


def strip_tags(text: str) -> str:
    text = html.unescape(text)
    return re.sub(r"<[^>]+>", "", text)


def normalize_title(text: str) -> str:
    text = strip_tags(text).upper()
    text = text.replace("∞", "INFINITY")
    text = text.replace("–", "-").replace("—", "-").replace("‐", "-")
    for source, target in GREEK_TITLE_MAP.items():
        text = text.replace(source, target)
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def find_first_match(base_dir: Path, patterns: list[str]) -> Path:
    for pattern in patterns:
        matches = sorted(base_dir.glob(pattern))
        if matches:
            return matches[0]
    raise FileNotFoundError(f"未在 {base_dir} 中找到匹配文件：{patterns}")


def build_sources(base_dir: Path) -> list[tuple[int, Path, str]]:
    issue_45 = find_first_match(base_dir, ["ESI*第4-5期附表*.xlsx", "ESI*第4-５期附表*.xlsx"])
    return [
        (1, find_first_match(base_dir, ["ESI*第1期附表*.xlsx"]), "附表1.本期我校高被引论文和热点论文"),
        (2, find_first_match(base_dir, ["ESI*第2期附表*.xlsx"]), "附表1.本期我校高被引论文和热点论文"),
        (3, find_first_match(base_dir, ["ESI*第3期附表*.xlsx"]), "附表3.本期我校高被引论文和热点论文"),
        (4, issue_45, "附表1.2025年7月高被引论文和热点论文"),
        (5, issue_45, "附表2.2025年9月高被引论文和热点论文"),
        (6, find_first_match(base_dir, ["ESI*第6期附表*.xlsx"]), "附表1.本期我校高被引论文和热点论文"),
    ]


MANUAL_SUBJECT_OVERRIDES = {
    normalize_title(title): subjects for title, subjects in MANUAL_SUBJECT_OVERRIDE_RAW.items()
}

MANUAL_PENDING_AUTHOR_OVERRIDES = {
    normalize_title(title): set(authors) for title, authors in MANUAL_PENDING_AUTHOR_OVERRIDES_RAW.items()
}


def normalize_display_title(text: str) -> str:
    text = strip_tags(text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def split_multi_value(text: str) -> list[str]:
    if not text:
        return []
    text = text.replace("、", "，").replace(",", "，").replace(";", "，").replace("；", "，")
    parts = [COLLEGE_ALIASES.get(part.strip(), part.strip()) for part in text.split("，")]
    return [part for part in parts if part]


def normalize_subject(text: str) -> list[str]:
    if not text:
        return []
    text = strip_tags(text).upper().replace(" AND ", " & ")
    text = text.replace(" / ", "/")
    if text.strip() in SUBJECT_MAP:
        return [SUBJECT_MAP[text.strip()]]
    parts = re.split(r"[;；，、]+", text)
    subjects = []
    for part in parts:
        key = part.strip()
        if not key:
            continue
        subjects.append(SUBJECT_MAP.get(key, key))
    return subjects


def extract_chinese_names(text: str) -> list[str]:
    names: list[str] = []
    for raw in re.findall(r"[\u4e00-\u9fff]{2,4}", text):
        if raw not in {"原信息工程学院"} and raw not in names:
            names.append(raw)
    if names:
        return names
    for match in re.findall(r"\(([^()]+)\)", text):
        candidate = re.sub(r"\s+", " ", match).strip(" .;，；")
        if "," in candidate and all("\u4e00" > ch or ch > "\u9fff" for ch in candidate):
            if candidate not in names:
                names.append(candidate)
    return names


def build_author_seed(template_path: Path) -> dict[str, str]:
    doc = Document(template_path)
    rows = [[cell.text.strip() for cell in row.cells] for row in doc.tables[2].rows[1:]]
    mapping: dict[str, str] = {}
    current = ["", "", ""]
    for row in rows:
        for idx in range(3):
            name = row[idx * 2].replace("\n", "").strip()
            if not name:
                continue
            if name in COLLEGE_ALIASES.values():
                current[idx] = name
            else:
                mapping[name] = current[idx]
    return mapping


AUTHOR_SEED: dict[str, str] = {}


@dataclass
class PaperRecord:
    issue: int
    category: str
    title_raw: str
    title_key: str
    authors_raw: str
    publication: str
    citations: str
    year: str
    first_corr: str
    colleges: list[str]
    subjects: list[str]
    chinese_names: list[str]
    author_entities: list[dict[str, object]]
    unresolved_author_tokens: list[str] = field(default_factory=list)


@dataclass
class PaperAggregate:
    title_key: str
    title: str = ""
    high_cited_issues: set[int] = field(default_factory=set)
    hot_issues: set[int] = field(default_factory=set)
    colleges: set[str] = field(default_factory=set)
    subjects: set[str] = field(default_factory=set)
    chinese_names: list[str] = field(default_factory=list)
    unresolved_author_tokens: list[str] = field(default_factory=list)
    records: list[PaperRecord] = field(default_factory=list)

    def add(self, record: PaperRecord) -> None:
        if record.category == "高被引论文":
            self.high_cited_issues.add(record.issue)
        else:
            self.hot_issues.add(record.issue)
        if len(record.title_raw) > len(self.title):
            self.title = normalize_display_title(record.title_raw)
        self.records.append(record)
        self.colleges.update(record.colleges)
        self.subjects.update(record.subjects)
        for name in record.chinese_names:
            if name not in self.chinese_names:
                self.chinese_names.append(name)
        for token in record.unresolved_author_tokens:
            if token not in self.unresolved_author_tokens:
                self.unresolved_author_tokens.append(token)


def dedupe_preserve(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for item in items:
        if not item or item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return ordered


def normalize_author_alias(text: str) -> str:
    text = strip_tags(text)
    text = re.sub(r"[\u4e00-\u9fff]", " ", text)
    text = re.sub(r"[^A-Za-z]+", " ", text).lower()
    return "".join(text.split())


def normalize_display_pinyin(text: str) -> str:
    text = strip_tags(text)
    text = re.sub(r"[\u4e00-\u9fff]", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" .;，；")
    text = re.sub(r"\s*,\s*", ",", text)
    return text


def extract_display_pinyin(token: str) -> str:
    candidates: list[str] = []
    for inner in re.findall(r"[（(]([^()（）]+)[)）]", token):
        candidate = normalize_display_pinyin(inner)
        if "," in candidate and re.search(r"[A-Za-z]", candidate):
            candidates.append(candidate)
    prefix = normalize_display_pinyin(re.split(r"[（(]", token, maxsplit=1)[0])
    if "," in prefix and re.search(r"[A-Za-z]", prefix):
        candidates.append(prefix)
    if not candidates:
        fallback = normalize_display_pinyin(token)
        if re.search(r"[A-Za-z]", fallback):
            candidates.append(fallback)
    if not candidates:
        return ""
    return max(candidates, key=lambda item: (sum(char.isalpha() for char in item), len(item)))


def extract_author_entities(text: str) -> list[dict[str, object]]:
    entities: list[dict[str, object]] = []
    if not text:
        return entities

    tokens = [token.strip() for token in re.split(r"[;；]+", text) if token.strip()]
    for token in tokens:
        chinese_names = [
            name
            for name in re.findall(r"[\u4e00-\u9fff]{2,4}", token)
            if name not in AUTHOR_NAME_STOPWORDS and name != "原信息工程学院"
        ]
        aliases: list[str] = []
        for inner in re.findall(r"[（(]([^()（）]+)[)）]", token):
            alias = normalize_author_alias(inner)
            if len(alias) >= 5:
                aliases.append(alias)
        prefix = re.split(r"[（(]", token, maxsplit=1)[0]
        prefix_alias = normalize_author_alias(prefix)
        if len(prefix_alias) >= 5:
            aliases.append(prefix_alias)

        entities.append(
            {
                "raw": token,
                "chinese_names": dedupe_preserve(chinese_names),
                "aliases": dedupe_preserve(aliases),
            }
        )
    return entities


def build_author_alias_map(records: Iterable[PaperRecord]) -> dict[str, str]:
    alias_counts: dict[str, Counter[str]] = defaultdict(Counter)
    for record in records:
        for entity in record.author_entities:
            chinese_names = entity["chinese_names"]
            aliases = entity["aliases"]
            if not chinese_names or not aliases:
                continue
            for alias in aliases:
                for chinese_name in chinese_names:
                    alias_counts[alias][chinese_name] += 1

    for alias, chinese_name in AUTHOR_ALIAS_MANUAL.items():
        alias_counts[alias][chinese_name] += 100

    alias_map: dict[str, str] = {}
    for alias, counter in alias_counts.items():
        alias_map[alias] = counter.most_common(1)[0][0]
    return alias_map


def resolve_record_authors(record: PaperRecord, alias_map: dict[str, str]) -> tuple[list[str], list[str]]:
    chinese_names: list[str] = []
    unresolved: list[str] = []
    for entity in record.author_entities:
        direct_names = entity["chinese_names"]
        if direct_names:
            chinese_names.extend(direct_names)
            continue
        resolved = None
        for alias in entity["aliases"]:
            if alias in alias_map:
                resolved = alias_map[alias]
                break
        if resolved:
            chinese_names.append(resolved)
        elif entity["raw"]:
            unresolved.append(str(entity["raw"]))
    return dedupe_preserve(chinese_names), dedupe_preserve(unresolved)


def read_sheet(issue: int, path: Path, sheet_name: str) -> list[PaperRecord]:
    df = pd.read_excel(path, sheet_name=sheet_name, header=None).fillna("")
    records: list[PaperRecord] = []
    section = ""
    headers: list[str] = []
    for _, row in df.iterrows():
        values = [clean_text(v) for v in row.tolist()]
        first = values[0]
        if "高被引论文清单" in first:
            section = "高被引论文"
            headers = []
            continue
        if "热点论文清单" in first:
            section = "热点论文"
            headers = []
            continue
        if first == "序号":
            headers = values
            continue
        if not section or not headers:
            continue
        if not first.isdigit():
            continue
        data = dict(zip(headers, values))
        title_raw = data.get("题名", "")
        authors_raw = data.get("作者", "")
        record = PaperRecord(
            issue=issue,
            category=section,
            title_raw=title_raw,
            title_key=normalize_title(title_raw),
            authors_raw=authors_raw,
            publication=data.get("出版物", ""),
            citations=data.get("被引次数", ""),
            year=data.get("出版时间", ""),
            first_corr=data.get("第一/通讯作者", ""),
            colleges=split_multi_value(data.get("所属学院", "")),
            subjects=normalize_subject(data.get("所属ESI学科", "")),
            chinese_names=[],
            author_entities=extract_author_entities(authors_raw),
        )
        records.append(record)
    return records


def infer_author_mapping(aggregates: Iterable[PaperAggregate]) -> dict[str, str]:
    counts: dict[str, Counter[str]] = defaultdict(Counter)
    for agg in aggregates:
        if len(agg.colleges) != 1:
            continue
        only_college = next(iter(agg.colleges))
        for name in agg.chinese_names:
            counts[name][only_college] += 1
    mapping = dict(AUTHOR_SEED)
    for name, counter in counts.items():
        college, freq = counter.most_common(1)[0]
        if len(counter) == 1 or freq >= 2:
            mapping.setdefault(name, college)
    return mapping


def choose_author_for_college(agg: PaperAggregate, college: str, author_map: dict[str, str]) -> str:
    for name in agg.chinese_names:
        if author_map.get(name) == college:
            return name
    for name in agg.chinese_names:
        if name in AUTHOR_SEED and AUTHOR_SEED[name] == college:
            return name
    if agg.chinese_names:
        return agg.chinese_names[0]
    if agg.unresolved_author_tokens:
        pinyin = extract_display_pinyin(agg.unresolved_author_tokens[0])
        if pinyin:
            return pinyin
        return "待核作者"
    return "未识别作者"


def run_analysis(base_dir: Path, out_dir: Path, template_2024: Path) -> dict[str, object]:
    global BASE, OUT, SOURCES, AUTHOR_SEED

    BASE = base_dir
    OUT = out_dir
    SOURCES = build_sources(base_dir)
    AUTHOR_SEED = build_author_seed(template_2024)

    OUT.mkdir(parents=True, exist_ok=True)

    records: list[PaperRecord] = []
    for issue, path, sheet_name in SOURCES:
        records.extend(read_sheet(issue, path, sheet_name))

    author_alias_map = build_author_alias_map(records)
    for record in records:
        record.chinese_names, record.unresolved_author_tokens = resolve_record_authors(record, author_alias_map)

    aggregates: dict[str, PaperAggregate] = {}
    for record in records:
        agg = aggregates.setdefault(record.title_key, PaperAggregate(title_key=record.title_key))
        agg.add(record)

    # Backfill missing college/subject info from other periods when early sheets omitted them.
    for agg in aggregates.values():
        if not agg.colleges:
            for name in agg.chinese_names:
                if name in AUTHOR_SEED:
                    agg.colleges.add(AUTHOR_SEED[name])
        if not agg.subjects and agg.title_key in MANUAL_SUBJECT_OVERRIDES:
            agg.subjects.update(MANUAL_SUBJECT_OVERRIDES[agg.title_key])

    author_map = infer_author_mapping(aggregates.values())

    union_total = len(aggregates)
    high_cited_total = sum(1 for agg in aggregates.values() if agg.high_cited_issues)
    hot_total = sum(1 for agg in aggregates.values() if agg.hot_issues)

    hc_dist = Counter(len(agg.high_cited_issues) for agg in aggregates.values() if agg.high_cited_issues)
    hot_dist = Counter(len(agg.hot_issues) for agg in aggregates.values() if agg.hot_issues)

    issue_counts = defaultdict(lambda: {"高被引论文": 0, "热点论文": 0})
    for record in records:
        issue_counts[record.issue][record.category] += 1

    missing_meta = []
    for agg in aggregates.values():
        if not agg.colleges or not agg.subjects:
            missing_meta.append(
                {
                    "title": agg.title,
                    "colleges": sorted(agg.colleges),
                    "subjects": sorted(agg.subjects),
                }
            )

    college_subject_counts: dict[str, Counter[str]] = defaultdict(Counter)
    college_unique_counts: Counter[str] = Counter()
    author_counts: dict[str, Counter[str]] = defaultdict(Counter)
    subject_source_colleges: dict[str, set[str]] = defaultdict(set)

    for agg in aggregates.values():
        colleges = sorted(agg.colleges)
        subjects = sorted(agg.subjects)
        for college in colleges:
            college_unique_counts[college] += 1
            author = choose_author_for_college(agg, college, author_map)
            author_counts[college][author] += 1
            for subject in subjects:
                college_subject_counts[college][subject] += 1
                subject_source_colleges[subject].add(college)

    hot_college_counts = Counter()
    hot_multi = []
    hot_became_hc = 0
    for agg in aggregates.values():
        if not agg.hot_issues:
            continue
        if agg.high_cited_issues:
            hot_became_hc += 1
        for college in agg.colleges:
            hot_college_counts[college] += 1
        if len(agg.hot_issues) >= 2:
            primary_college = next(iter(sorted(agg.colleges))) if agg.colleges else ""
            primary_author = choose_author_for_college(agg, primary_college, author_map) if primary_college else ""
            hot_multi.append(
                {
                    "title": agg.title,
                    "issues": sorted(agg.hot_issues),
                    "college": primary_college,
                    "author": primary_author,
                }
            )

    order_map = {name: idx for idx, name in enumerate(SUBJECT_ORDER)}
    unique_subjects = sorted(
        {subject for agg in aggregates.values() for subject in agg.subjects},
        key=lambda subject: (order_map.get(subject, 999), subject),
    )
    unique_colleges = [c for c in COLLEGE_ORDER if c in college_unique_counts]

    unresolved_author_papers = []
    pinyin_pending_papers = []
    for agg in sorted(aggregates.values(), key=lambda item: item.title):
        if not agg.unresolved_author_tokens:
            continue
        all_display_authors = dedupe_preserve(
            [extract_display_pinyin(token) for token in agg.unresolved_author_tokens if extract_display_pinyin(token)]
        )
        if not agg.chinese_names:
            display_authors = all_display_authors
        else:
            allowed = MANUAL_PENDING_AUTHOR_OVERRIDES.get(agg.title_key, set())
            display_authors = [name for name in all_display_authors if name in allowed]
        if not display_authors:
            continue
        item = {
            "title": agg.title,
            "colleges": sorted(agg.colleges),
            "resolved_authors": agg.chinese_names,
            "raw_author_tokens": agg.unresolved_author_tokens,
            "display_authors": display_authors,
            "display_author": display_authors[0] if display_authors else "",
            "fully_unresolved": not bool(agg.chinese_names),
            "issues": {
                "high_cited": sorted(agg.high_cited_issues),
                "hot": sorted(agg.hot_issues),
            },
        }
        pinyin_pending_papers.append(item)
        if not agg.chinese_names:
            unresolved_author_papers.append(item)

    summary = {
        "totals": {
            "high_level": union_total,
            "high_cited": high_cited_total,
            "hot": hot_total,
        },
        "distribution": {
            "high_cited": {str(k): hc_dist.get(k, 0) for k in range(1, 7)},
            "hot": {str(k): hot_dist.get(k, 0) for k in range(1, 7)},
        },
        "issue_counts": issue_counts,
        "hot_college_counts": dict(hot_college_counts.most_common()),
        "hot_multi": hot_multi,
        "hot_also_high_cited": hot_became_hc,
        "unique_subjects": unique_subjects,
        "unique_colleges": unique_colleges,
        "college_unique_counts": dict(college_unique_counts),
        "college_subject_counts": {k: dict(v) for k, v in college_subject_counts.items()},
        "subject_source_colleges": {k: sorted(v) for k, v in subject_source_colleges.items()},
        "author_counts": {k: dict(v) for k, v in author_counts.items()},
        "missing_meta": missing_meta,
        "unresolved_author_papers": unresolved_author_papers,
        "pinyin_pending_papers": pinyin_pending_papers,
        "author_alias_map": author_alias_map,
    }

    papers = []
    for agg in sorted(aggregates.values(), key=lambda x: x.title):
        papers.append(
            {
                "title": agg.title,
                "high_cited_issues": sorted(agg.high_cited_issues),
                "hot_issues": sorted(agg.hot_issues),
                "colleges": sorted(agg.colleges),
                "subjects": sorted(agg.subjects),
                "authors": agg.chinese_names,
            }
        )

    (OUT / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    (OUT / "papers.json").write_text(json.dumps(papers, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="分析 ESI 高水平论文附表并生成汇总 JSON")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_BASE, help="包含 6 期附表的文件夹")
    parser.add_argument("--analysis-dir", type=Path, default=DEFAULT_OUT, help="analysis 输出目录")
    parser.add_argument(
        "--template-2024",
        type=Path,
        default=DEFAULT_TEMPLATE_2024,
        help="用于构建作者种子映射的 2024 年报告 docx",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    summary = run_analysis(args.input_dir, args.analysis_dir, args.template_2024)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
