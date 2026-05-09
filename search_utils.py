from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Any

from database import fetch_dataframe


def normalize_text(text: Any) -> str:
    raw = "" if text is None else str(text)
    return re.sub(r"\s+", " ", raw.lower()).strip()


def normalize_standard_no(text: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize_text(text))


def fuzzy_ratio(a: str, b: str) -> float:
    if not a or not b:
        return 0
    return SequenceMatcher(None, a, b).ratio()


def row_relevance(row: dict[str, Any], query: str) -> float:
    q = normalize_text(query)
    q_std = normalize_standard_no(query)
    product_cn = normalize_text(row.get("product_cn"))
    product_en = normalize_text(row.get("product_en"))
    aliases = normalize_text(row.get("aliases"))
    standard_no = normalize_text(row.get("standard_no"))
    standard_compact = normalize_standard_no(row.get("standard_no"))

    score = 0.0
    if q and q == product_cn:
        score += 100
    if q and q in product_cn:
        score += 80
    if q and product_cn and product_cn in q:
        score += 75
    if q and q == product_en:
        score += 70
    if q and q in product_en:
        score += 60
    if q and product_en and product_en in q:
        score += 55
    if q and q in aliases:
        score += 50
    if q and any(part.strip() and part.strip() in q for part in re.split(r"[;,，；/|]", aliases)):
        score += 45
    if q_std and (q_std == standard_compact or q_std in standard_compact):
        score += 40

    lower_blob = normalize_text(
        " ".join(
            str(row.get(k, ""))
            for k in [
                "section_title",
                "parent_product_cn",
                "parent_product_en",
                "component_cn",
                "component_en",
                "requirement_object",
                "requirement_type",
                "assembly_scope",
                "requirement_category",
                "test_item_cn",
                "test_item_en",
                "applicable_condition",
                "test_condition",
                "specific_value",
                "pass_criteria",
                "risk_note",
            ]
        )
    )
    if q and q in lower_blob:
        # Clause body / notes are useful, but they should not outrank product,
        # alias, or standard-number matches. This avoids a support term such as
        # "水力警铃" surfacing "湿式报警阀" only because it is mentioned in a risk note.
        score += 8

    candidates = [product_cn, product_en, aliases, standard_no, lower_blob]
    score += max(fuzzy_ratio(q, candidate) for candidate in candidates) * 25 if q else 0
    return score


def search_requirements(
    query: str = "",
    product_category: str = "",
    standard_no: str = "",
    requirement_category: str = "",
) -> list[dict[str, Any]]:
    rows = fetch_dataframe(
        """
        SELECT
            r.id AS requirement_id,
            p.id AS product_id,
            s.id AS standard_id,
            p.product_cn,
            p.product_en,
            p.aliases,
            p.category AS product_category,
            p.description AS product_description,
            s.standard_no,
            s.standard_title,
            s.standard_version,
            s.scope_summary,
            s.copyright_note,
            r.parent_product_cn,
            r.parent_product_en,
            r.component_cn,
            r.component_en,
            r.requirement_object,
            r.requirement_type,
            r.assembly_scope,
            r.clause_no,
            r.page_no,
            r.section_title,
            r.requirement_category,
            r.test_item_cn,
            r.test_item_en,
            r.applicable_condition,
            r.test_condition,
            r.specific_value,
            r.pass_criteria,
            r.fail_criteria,
            r.source_excerpt_short,
            r.supplier_documents,
            r.sales_explanation_cn,
            r.sales_explanation_en,
            r.risk_note
        FROM requirements r
        JOIN products p ON p.id = r.product_id
        JOIN standards s ON s.id = r.standard_id
        """
    )
    result = [dict(row) for row in rows]

    if product_category:
        result = [r for r in result if r.get("product_category") == product_category]
    if standard_no:
        target = normalize_standard_no(standard_no)
        result = [r for r in result if normalize_standard_no(r.get("standard_no")) == target]
    if requirement_category:
        result = [r for r in result if r.get("requirement_category") == requirement_category]

    if query.strip():
        scored = [(row_relevance(r, query), r) for r in result]
        scored = [(score, r) for score, r in scored if score >= 25]
        scored.sort(key=lambda item: item[0], reverse=True)
        result = [r | {"relevance": round(score, 2)} for score, r in scored]
    else:
        result = [r | {"relevance": 0} for r in result]

    return result


def get_filter_values() -> dict[str, list[str]]:
    rows = fetch_dataframe(
        """
        SELECT DISTINCT p.category AS product_category, s.standard_no, r.requirement_category
        FROM requirements r
        JOIN products p ON p.id = r.product_id
        JOIN standards s ON s.id = r.standard_id
        """
    )
    return {
        "product_categories": sorted({r["product_category"] for r in rows if r["product_category"]}),
        "standard_nos": sorted({r["standard_no"] for r in rows if r["standard_no"]}),
        "requirement_categories": sorted({r["requirement_category"] for r in rows if r["requirement_category"]}),
    }


def get_checklist(product_id: int) -> list[dict[str, Any]]:
    rows = fetch_dataframe(
        """
        SELECT checklist_item_cn, checklist_item_en, explanation, priority
        FROM supplier_checklists
        WHERE product_id = ?
        ORDER BY priority ASC, id ASC
        """,
        (product_id,),
    )
    return [dict(r) for r in rows]
