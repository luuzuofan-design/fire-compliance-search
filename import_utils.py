from __future__ import annotations

from datetime import datetime
from typing import Any

import pandas as pd

from database import PLACEHOLDER_VALUE, get_connection
from template_utils import TEMPLATE_COLUMNS


PENDING = "\u5f85\u8865\u5145"


def read_uploaded_file(uploaded_file: Any) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str).fillna("")
    return pd.read_excel(uploaded_file, dtype=str).fillna("")


def validate_import_dataframe(df: pd.DataFrame) -> tuple[bool, list[str]]:
    errors: list[str] = []
    missing = [c for c in TEMPLATE_COLUMNS if c not in df.columns]
    if missing:
        errors.append("\u7f3a\u5c11\u5b57\u6bb5: " + ", ".join(missing))
    if df.empty:
        errors.append("\u6587\u4ef6\u6ca1\u6709\u53ef\u5bfc\u5165\u7684\u6570\u636e\u3002")

    for idx, row in df.iterrows():
        line = idx + 2
        product_ok = str(row.get("product_cn", "")).strip() or str(row.get("product_en", "")).strip()
        test_ok = str(row.get("test_item_cn", "")).strip() or str(row.get("test_item_en", "")).strip()
        if not product_ok:
            errors.append(f"\u7b2c {line} \u884c: product_cn \u6216 product_en \u81f3\u5c11\u586b\u5199\u4e00\u4e2a\u3002")
        if not str(row.get("standard_no", "")).strip():
            errors.append(f"\u7b2c {line} \u884c: standard_no \u5fc5\u586b\u3002")
        if not test_ok:
            errors.append(f"\u7b2c {line} \u884c: test_item_cn \u6216 test_item_en \u81f3\u5c11\u586b\u5199\u4e00\u4e2a\u3002")

    return not errors, errors


def _clean(value: Any, fallback: str = "") -> str:
    text = "" if value is None else str(value).strip()
    return text if text and text.lower() != "nan" else fallback


def _get_or_create_product(conn, row: pd.Series) -> int:
    product_cn = _clean(row.get("product_cn"))
    product_en = _clean(row.get("product_en"))
    aliases = _clean(row.get("aliases"))
    category = _clean(row.get("product_category"))
    existing = conn.execute(
        """
        SELECT id FROM products
        WHERE lower(COALESCE(product_cn,'')) = lower(?)
           OR lower(COALESCE(product_en,'')) = lower(?)
        LIMIT 1
        """,
        (product_cn, product_en),
    ).fetchone()
    if existing:
        conn.execute(
            """
            UPDATE products
            SET aliases = COALESCE(NULLIF(?, ''), aliases),
                category = COALESCE(NULLIF(?, ''), category),
                updated_at = datetime('now','localtime')
            WHERE id = ?
            """,
            (aliases, category, existing["id"]),
        )
        return int(existing["id"])

    cur = conn.execute(
        """
        INSERT INTO products (product_cn, product_en, aliases, category, description)
        VALUES (?, ?, ?, ?, ?)
        """,
        (product_cn, product_en, aliases, category, "\u7531\u5bfc\u5165\u6570\u636e\u521b\u5efa"),
    )
    return int(cur.lastrowid)


def _get_or_create_standard(conn, row: pd.Series) -> int:
    standard_no = _clean(row.get("standard_no"))
    standard_version = _clean(row.get("standard_version"), PENDING)
    standard_title = _clean(row.get("standard_title"))
    existing = conn.execute(
        """
        SELECT id FROM standards
        WHERE lower(replace(standard_no, ' ', '')) = lower(replace(?, ' ', ''))
          AND lower(COALESCE(standard_version,'')) = lower(?)
        LIMIT 1
        """,
        (standard_no, standard_version),
    ).fetchone()
    if existing:
        conn.execute(
            """
            UPDATE standards
            SET standard_title = COALESCE(NULLIF(?, ''), standard_title),
                updated_at = datetime('now','localtime')
            WHERE id = ?
            """,
            (standard_title, existing["id"]),
        )
        return int(existing["id"])

    cur = conn.execute(
        """
        INSERT INTO standards
        (standard_no, standard_title, standard_version, publisher, scope_summary, copyright_note)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            standard_no,
            standard_title,
            standard_version,
            "UL",
            "\u6765\u81ea\u5bfc\u5165\u6570\u636e\uff0c\u5177\u4f53\u9002\u7528\u8303\u56f4\u8bf7\u6309\u6807\u51c6\u6761\u6b3e\u8865\u5145\u3002",
            "\u4ec5\u4fdd\u5b58\u7ed3\u6784\u5316\u6458\u8981\u548c\u77ed\u6458\u5f55\uff0c\u4e0d\u4fdd\u5b58\u6216\u4f20\u64ad\u6807\u51c6\u5168\u6587\u3002",
        ),
    )
    return int(cur.lastrowid)


def import_requirements(df: pd.DataFrame) -> dict[str, int]:
    valid, errors = validate_import_dataframe(df)
    if not valid:
        raise ValueError("\n".join(errors))

    imported = 0
    products = set()
    standards = set()
    with get_connection() as conn:
        for _, row in df.iterrows():
            product_id = _get_or_create_product(conn, row)
            standard_id = _get_or_create_standard(conn, row)
            products.add(product_id)
            standards.add(standard_id)

            conn.execute(
                """
                INSERT INTO requirements
                (product_id, standard_id, parent_product_cn, parent_product_en,
                 component_cn, component_en, requirement_object, requirement_type,
                 assembly_scope, clause_no, page_no, section_title, requirement_category,
                 test_item_cn, test_item_en, applicable_condition, test_condition,
                 specific_value, pass_criteria, fail_criteria, source_excerpt_short,
                 supplier_documents, sales_explanation_cn, sales_explanation_en,
                 risk_note, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    product_id,
                    standard_id,
                    _clean(row.get("parent_product_cn")),
                    _clean(row.get("parent_product_en")),
                    _clean(row.get("component_cn")),
                    _clean(row.get("component_en")),
                    _clean(row.get("requirement_object")),
                    _clean(row.get("requirement_type")),
                    _clean(row.get("assembly_scope")),
                    _clean(row.get("clause_no"), PENDING),
                    _clean(row.get("page_no"), PENDING),
                    _clean(row.get("section_title"), PENDING),
                    _clean(row.get("requirement_category"), "\u672a\u5206\u7c7b"),
                    _clean(row.get("test_item_cn")),
                    _clean(row.get("test_item_en")),
                    _clean(row.get("applicable_condition"), PENDING),
                    _clean(row.get("test_condition"), PENDING),
                    _clean(row.get("specific_value"), PLACEHOLDER_VALUE),
                    _clean(row.get("pass_criteria"), PLACEHOLDER_VALUE),
                    _clean(row.get("fail_criteria"), PENDING),
                    _clean(row.get("source_excerpt_short")),
                    _clean(row.get("supplier_documents")),
                    _clean(row.get("sales_explanation_cn")),
                    _clean(row.get("sales_explanation_en")),
                    _clean(row.get("risk_note")),
                    _clean(row.get("created_at"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                    _clean(row.get("updated_at"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                ),
            )
            imported += 1

    return {"requirements": imported, "products": len(products), "standards": len(standards)}
