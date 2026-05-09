from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from database import get_connection
from search_utils import normalize_standard_no, normalize_text


CLAUSE_FILE_PREFIX = "UL_FM"
CLAUSE_FILE_HINT = "分级"
BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"


def locate_clause_table_file() -> Path | None:
    bundled_candidates = [
        path
        for path in ASSETS_DIR.glob("*.xlsx")
        if path.name.startswith(CLAUSE_FILE_PREFIX) and CLAUSE_FILE_HINT in path.name
    ]
    if bundled_candidates:
        bundled_candidates.sort(key=lambda path: ("完整版" not in path.name, path.name))
        return bundled_candidates[0]

    desktop = Path.home() / "Desktop"
    if not desktop.exists():
        return None
    candidates = [
        path
        for path in desktop.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xls"}
        and path.name.startswith(CLAUSE_FILE_PREFIX)
        and CLAUSE_FILE_HINT in path.name
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda path: ("完整版" not in path.name, path.name))
    return candidates[0]


def init_clause_table() -> None:
    with get_connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS clause_table_meta (
                source_path TEXT PRIMARY KEY,
                source_mtime REAL,
                indexed_at TEXT,
                row_count INTEGER
            );

            CREATE TABLE IF NOT EXISTS clause_table_rows (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_path TEXT NOT NULL,
                sheet_name TEXT,
                row_no INTEGER,
                standard_file TEXT,
                original_pdf TEXT,
                level1 TEXT,
                level2 TEXT,
                level3 TEXT,
                content TEXT,
                content_norm TEXT,
                indexed_at TEXT DEFAULT (datetime('now','localtime'))
            );

            CREATE INDEX IF NOT EXISTS idx_clause_table_source ON clause_table_rows(source_path);
            """
        )


def _clean(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    return "" if text.lower() == "nan" else text


def _read_clause_rows(path: Path) -> list[dict[str, Any]]:
    xl = pd.ExcelFile(path)
    sheet_name = "全部汇总" if "全部汇总" in xl.sheet_names else xl.sheet_names[1]
    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str).fillna("")

    # The full workbook uses these six columns. Fallback by position keeps the importer tolerant.
    if {"标准文件", "原PDF文件", "一级标题", "二级标题", "三级标题", "内容"}.issubset(df.columns):
        columns = ["标准文件", "原PDF文件", "一级标题", "二级标题", "三级标题", "内容"]
    else:
        columns = list(df.columns[:6])

    rows: list[dict[str, Any]] = []
    for idx, row in df.iterrows():
        values = [_clean(row.get(column, "")) for column in columns]
        if not any(values):
            continue
        standard_file, original_pdf, level1, level2, level3, content = values
        blob = " ".join(values)
        rows.append(
            {
                "sheet_name": sheet_name,
                "row_no": int(idx) + 2,
                "standard_file": standard_file,
                "original_pdf": original_pdf,
                "level1": level1,
                "level2": level2,
                "level3": level3,
                "content": content,
                "content_norm": normalize_text(blob),
            }
        )
    return rows


def index_clause_table(path: Path | None = None, force: bool = False) -> dict[str, Any]:
    init_clause_table()
    path = path or locate_clause_table_file()
    if not path:
        return {"indexed": False, "reason": "未找到 UL_FM 消防标准分级条款表", "rows": 0}

    stat = path.stat()
    source_path = str(path)
    with get_connection() as conn:
        meta = conn.execute(
            "SELECT source_mtime, row_count FROM clause_table_meta WHERE source_path = ?",
            (source_path,),
        ).fetchone()
        if meta and not force and float(meta["source_mtime"]) == float(stat.st_mtime):
            return {"indexed": False, "reason": "索引已是最新", "rows": int(meta["row_count"] or 0)}

        rows = _read_clause_rows(path)
        conn.execute("DELETE FROM clause_table_rows")
        conn.execute("DELETE FROM clause_table_meta")
        conn.executemany(
            """
            INSERT INTO clause_table_rows
            (source_path, sheet_name, row_no, standard_file, original_pdf, level1, level2, level3, content, content_norm, indexed_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    source_path,
                    item["sheet_name"],
                    item["row_no"],
                    item["standard_file"],
                    item["original_pdf"],
                    item["level1"],
                    item["level2"],
                    item["level3"],
                    item["content"],
                    item["content_norm"],
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                )
                for item in rows
            ],
        )
        conn.execute(
            """
            INSERT INTO clause_table_meta (source_path, source_mtime, indexed_at, row_count)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(source_path) DO UPDATE SET
                source_mtime = excluded.source_mtime,
                indexed_at = excluded.indexed_at,
                row_count = excluded.row_count
            """,
            (source_path, float(stat.st_mtime), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), len(rows)),
        )
    return {"indexed": True, "reason": "索引完成", "rows": len(rows), "path": source_path}


def clause_table_summary() -> dict[str, Any]:
    init_clause_table()
    with get_connection() as conn:
        files = conn.execute("SELECT COUNT(*) FROM clause_table_meta").fetchone()[0]
        rows = conn.execute("SELECT COUNT(*) FROM clause_table_rows").fetchone()[0]
        last_indexed = conn.execute("SELECT MAX(indexed_at) FROM clause_table_meta").fetchone()[0]
    return {"files": files, "rows": rows, "last_indexed": last_indexed or "尚未索引"}


def query_terms(query: str) -> list[str]:
    terms = [query.strip()]
    known_pairs = {
        "水力警铃": ["water motor gong", "water motor alarm", "1055", "FM1055"],
        "water motor gong": ["水力警铃", "1055", "FM1055"],
        "报警止回阀": ["alarm check valve", "FM1041", "1041"],
        "闸阀": ["gate valve", "UL262", "UL 262"],
        "止回阀": ["check valve", "UL312", "UL 312"],
        "报警附件": ["alarm accessories", "UL753", "UL 753"],
        "压力开关": ["pressure switch", "UL753", "UL 753"],
        "消防泵泄压阀": ["fire pump relief valve", "UL1478", "UL 1478"],
        "先导式压力控制阀": ["pilot-operated pressure-control valve", "UL1739", "UL 1739"],
        "切断阀": ["shutoff valve", "UL258", "UL 258"],
    }
    q = normalize_text(query)
    for key, values in known_pairs.items():
        key_norm = normalize_text(key)
        if key_norm and (key_norm in q or q in key_norm):
            terms.extend(values)
    for token in re.findall(r"[\u4e00-\u9fff]{2,}|[a-zA-Z][a-zA-Z0-9/-]{2,}|\d{3,4}", query):
        terms.append(token)
    compact = normalize_standard_no(query)
    match = re.search(r"(ul|fm)(\d{3,4})", compact)
    if match:
        terms.extend([match.group(0), f"{match.group(1).upper()} {match.group(2)}", match.group(2)])

    unique: list[str] = []
    seen = set()
    for term in terms:
        norm = normalize_text(term)
        if norm and norm not in seen:
            seen.add(norm)
            unique.append(term)
    return unique


def search_clause_table(query: str, limit: int = 40) -> list[dict[str, Any]]:
    init_clause_table()
    if not query.strip():
        return []
    terms = query_terms(query)
    q_norm = normalize_text(query)
    has_test_intent = any(term in q_norm for term in ["检验", "检测", "试验", "测试", "test", "examination"])
    rows = []
    with get_connection() as conn:
        db_rows = conn.execute(
            """
            SELECT id, source_path, sheet_name, row_no, standard_file, original_pdf,
                   level1, level2, level3, content, content_norm
            FROM clause_table_rows
            """
        ).fetchall()

    for row in db_rows:
        item = dict(row)
        blob = item["content_norm"]
        file_blob = normalize_text(" ".join([item.get("standard_file") or "", item.get("original_pdf") or ""]))
        title_blob = normalize_text(" ".join([item.get("level1") or "", item.get("level2") or "", item.get("level3") or ""]))
        score = 0
        matched = []
        for term in terms:
            norm = normalize_text(term)
            compact = normalize_standard_no(term)
            if not norm:
                continue
            if norm in blob:
                score += 35 + min(blob.count(norm), 4) * 5
                matched.append(term)
            if norm in title_blob:
                score += 25
                matched.append(term)
            if norm in file_blob:
                score += 30
                matched.append(term)
            if compact and compact in normalize_standard_no(file_blob):
                score += 35
                matched.append(term)

        if score <= 0:
            continue
        level1_norm = normalize_text(item.get("level1"))
        level2_norm = normalize_text(item.get("level2"))
        level3_norm = normalize_text(item.get("level3"))
        if level1_norm in {"文档信息", "标准名称", "目录", "目录（中文）", "前言", "前言与修订说明"}:
            score -= 180
        if not _clean(item.get("content")):
            continue
        if has_test_intent and any(
            term in " ".join([level1_norm, level2_norm, level3_norm, blob])
            for term in ["性能", "性能要求", "测试", "测试/验证", "检验", "试验", "制造和生产试验", "水压", "声压", "耐久"]
        ):
            score += 90
        if has_test_intent and any(term in level1_norm for term in ["性能", "性能要求"]) :
            score += 120
        if has_test_intent and re.match(r"4(\.|$)", _clean(item.get("level2"))):
            score += 80
        if has_test_intent and any(term in " ".join([level2_norm, level3_norm, blob]) for term in ["测试/验证", "检测/验证", "试验方法"]):
            score += 70
        if has_test_intent and re.search(r"\d+(\.\d+)?\s*(psi|kpa|mpa|分贝|db|小时|秒|s|min|gpm|l/min)", blob, re.IGNORECASE):
            score += 45
        if has_test_intent and "附录" in level1_norm:
            score -= 45
        if has_test_intent and any(
            term in " ".join([level1_norm, level2_norm, level3_norm, blob])
            for term in ["质量保证", "认证依据", "持续认证", "规范性引用", "术语", "定义"]
        ):
            score -= 60
        if score <= 0:
            continue
        item["score"] = score
        item["matched_terms"] = " / ".join(dict.fromkeys(matched))
        item["source_location"] = " / ".join(
            part for part in [item.get("original_pdf"), item.get("level1"), item.get("level2"), item.get("level3")] if part
        )
        rows.append(item)

    rows.sort(key=lambda item: item["score"], reverse=True)
    return rows[:limit]
