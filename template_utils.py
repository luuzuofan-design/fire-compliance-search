from __future__ import annotations

from io import BytesIO

import pandas as pd


TEMPLATE_COLUMNS = [
    "id",
    "product_cn",
    "product_en",
    "aliases",
    "product_category",
    "parent_product_cn",
    "parent_product_en",
    "component_cn",
    "component_en",
    "requirement_object",
    "requirement_type",
    "assembly_scope",
    "standard_no",
    "standard_title",
    "standard_version",
    "clause_no",
    "page_no",
    "section_title",
    "requirement_category",
    "test_item_cn",
    "test_item_en",
    "applicable_condition",
    "test_condition",
    "specific_value",
    "pass_criteria",
    "fail_criteria",
    "source_excerpt_short",
    "supplier_documents",
    "sales_explanation_cn",
    "sales_explanation_en",
    "risk_note",
    "created_at",
    "updated_at",
]


REQUIRED_COLUMNS = [
    "product_cn",
    "product_en",
    "standard_no",
    "test_item_cn",
    "test_item_en",
]


def generate_template() -> bytes:
    example = {
        "id": "",
        "product_cn": "\u6e7f\u5f0f\u62a5\u8b66\u9600",
        "product_en": "Wet Alarm Valve",
        "aliases": "\u62a5\u8b66\u9600; alarm valve; wet alarm valve",
        "product_category": "\u62a5\u8b66\u9600\u7ec4",
        "parent_product_cn": "\u6e7f\u5f0f\u62a5\u8b66\u9600\u7ec4",
        "parent_product_en": "Wet Alarm Valve Assembly",
        "component_cn": "\u9600\u4f53",
        "component_en": "Valve Body",
        "requirement_object": "Valve body material / construction",
        "requirement_type": "Construction",
        "assembly_scope": "Assembly / component-level requirement",
        "standard_no": "UL 193",
        "standard_title": "Alarm Valves for Fire-Protection Service",
        "standard_version": "\u5f85\u8865\u5145",
        "clause_no": "\u5f85\u8865\u5145",
        "page_no": "\u5f85\u8865\u5145",
        "section_title": "\u5f85\u8865\u5145",
        "requirement_category": "Construction",
        "test_item_cn": "\u9600\u4f53\u6750\u6599/\u7ed3\u6784\u8981\u6c42",
        "test_item_en": "Valve Body Material / Construction Requirement",
        "applicable_condition": "\u5f85\u8865\u5145",
        "test_condition": "\u5f85\u8865\u5145",
        "specific_value": "\u5f85\u8865\u5145\uff1a\u8bf7\u4ece\u5408\u6cd5\u6301\u6709\u7684\u6807\u51c6\u6761\u6b3e\u5f55\u5165",
        "pass_criteria": "\u5f85\u8865\u5145\uff1a\u8bf7\u4ece\u5408\u6cd5\u6301\u6709\u7684\u6807\u51c6\u6761\u6b3e\u5f55\u5165",
        "fail_criteria": "\u5f85\u8865\u5145",
        "source_excerpt_short": "\u77ed\u6458\u5f55\u6216\u5173\u952e\u8bcd\uff0c\u4e0d\u4fdd\u5b58\u6807\u51c6\u5168\u6587",
        "supplier_documents": "UL Certificate; UL File Number; installation instructions; marking sample",
        "sales_explanation_cn": "\u8be5\u9879\u4e3a\u90e8\u4ef6\u7ea7\u5408\u89c4\u8981\u6c42\uff0c\u9700\u6838\u5bf9\u5177\u4f53\u578b\u53f7\u548c\u7ed3\u6784\u662f\u5426\u88ab Listing \u8986\u76d6\u3002",
        "sales_explanation_en": "This is a component-level compliance requirement. Please verify whether the exact model and construction are covered by the Listing.",
        "risk_note": "\u771f\u5b9e\u6570\u503c\u3001\u6761\u6b3e\u53f7\u548c\u9875\u7801\u5fc5\u987b\u4eba\u5de5\u8865\u5145\u3002",
        "created_at": "",
        "updated_at": "",
    }
    df = pd.DataFrame([example], columns=TEMPLATE_COLUMNS)
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="requirements_template")
    return bio.getvalue()
