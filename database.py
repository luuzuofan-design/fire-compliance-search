from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, Iterable


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "ul_fire_system.db"

PLACEHOLDER_VALUE = "数据库暂无该项具体数值，请补充标准条款数据"
PENDING = "待补充"


def get_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def now_sql() -> str:
    return "datetime('now','localtime')"


def init_db() -> None:
    with get_connection() as conn:
        conn.executescript(
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_cn TEXT,
                product_en TEXT,
                aliases TEXT,
                category TEXT,
                description TEXT,
                created_at TEXT DEFAULT (datetime('now','localtime')),
                updated_at TEXT DEFAULT (datetime('now','localtime'))
            );

            CREATE TABLE IF NOT EXISTS standards (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                standard_no TEXT NOT NULL,
                standard_title TEXT,
                standard_version TEXT,
                publisher TEXT,
                scope_summary TEXT,
                copyright_note TEXT,
                created_at TEXT DEFAULT (datetime('now','localtime')),
                updated_at TEXT DEFAULT (datetime('now','localtime')),
                UNIQUE(standard_no, standard_version)
            );

            CREATE TABLE IF NOT EXISTS requirements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                standard_id INTEGER NOT NULL,
                clause_no TEXT,
                page_no TEXT,
                section_title TEXT,
                requirement_category TEXT,
                test_item_cn TEXT,
                test_item_en TEXT,
                applicable_condition TEXT,
                test_condition TEXT,
                specific_value TEXT,
                pass_criteria TEXT,
                fail_criteria TEXT,
                source_excerpt_short TEXT,
                supplier_documents TEXT,
                sales_explanation_cn TEXT,
                sales_explanation_en TEXT,
                risk_note TEXT,
                created_at TEXT DEFAULT (datetime('now','localtime')),
                updated_at TEXT DEFAULT (datetime('now','localtime')),
                FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE,
                FOREIGN KEY(standard_id) REFERENCES standards(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS supplier_checklists (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                checklist_item_cn TEXT NOT NULL,
                checklist_item_en TEXT,
                explanation TEXT,
                priority INTEGER DEFAULT 3,
                created_at TEXT DEFAULT (datetime('now','localtime')),
                updated_at TEXT DEFAULT (datetime('now','localtime')),
                FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE
            );
            """
        )
        migrate_db(conn)
        seed_database(conn)


def migrate_db(conn: sqlite3.Connection) -> None:
    existing = {row["name"] for row in conn.execute("PRAGMA table_info(requirements)").fetchall()}
    additions = {
        "parent_product_cn": "TEXT",
        "parent_product_en": "TEXT",
        "component_cn": "TEXT",
        "component_en": "TEXT",
        "requirement_object": "TEXT",
        "requirement_type": "TEXT",
        "assembly_scope": "TEXT",
    }
    for column, column_type in additions.items():
        if column not in existing:
            conn.execute(f"ALTER TABLE requirements ADD COLUMN {column} {column_type}")


def seed_database(conn: sqlite3.Connection) -> None:
    count = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    if count:
        return

    products = [
        (
            "蝶阀",
            "Butterfly Valve",
            "消防蝶阀;信号蝶阀;butterfly valve;fire-protection butterfly valve;signal butterfly valve;supervisory butterfly valve",
            "阀门",
            "消防水系统蝶阀产品。示例数据不包含真实 UL 试验数值。",
        ),
        (
            "水流指示器",
            "Waterflow Indicator",
            "水流开关;waterflow indicator;water flow switch;vane type waterflow indicator",
            "报警装置",
            "消防水系统水流指示器。示例数据不包含真实 UL 试验数值。",
        ),
        (
            "湿式报警阀",
            "Wet Alarm Valve",
            "报警阀;alarm valve;wet alarm valve",
            "报警阀组",
            "湿式报警阀。示例数据不包含真实 UL 试验数值。",
        ),
    ]

    conn.executemany(
        """
        INSERT INTO products (product_cn, product_en, aliases, category, description)
        VALUES (?, ?, ?, ?, ?)
        """,
        products,
    )

    standards = [
        (
            "UL 1091",
            "Butterfly Valves for Fire-Protection Service",
            PENDING,
            "UL",
            PENDING,
            "仅保存结构化摘要和短摘录，不保存或传播标准全文。",
        ),
        (
            "UL 346",
            "Waterflow Indicators for Fire Protective Signaling Systems",
            PENDING,
            "UL",
            PENDING,
            "仅保存结构化摘要和短摘录，不保存或传播标准全文。",
        ),
        (
            "UL 193",
            "Alarm Valves for Fire-Protection Service",
            PENDING,
            "UL",
            PENDING,
            "仅保存结构化摘要和短摘录，不保存或传播标准全文。",
        ),
    ]
    conn.executemany(
        """
        INSERT INTO standards
        (standard_no, standard_title, standard_version, publisher, scope_summary, copyright_note)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        standards,
    )

    product_ids = {r["product_cn"]: r["id"] for r in conn.execute("SELECT id, product_cn FROM products")}
    standard_ids = {r["standard_no"]: r["id"] for r in conn.execute("SELECT id, standard_no FROM standards")}

    requirements = [
        (
            product_ids["蝶阀"],
            standard_ids["UL 1091"],
            PENDING,
            PENDING,
            "阀体强度相关条款",
            "强度试验 Strength Test",
            "阀体强度试验",
            "Body Strength Test",
            PENDING,
            PENDING,
            PLACEHOLDER_VALUE,
            PLACEHOLDER_VALUE,
            PENDING,
            "请录入短摘录或关键词，避免保存标准全文。",
            "UL Certificate; UL File Number; Product iQ listing; installation instructions; marking sample",
            "我们的消防蝶阀按 UL 1091 进行评估。关键合规项目包括阀体强度、阀座密封、操作性能、耐久性、结构、标识和安装说明。请提供所需尺寸、压力等级、连接方式和材质，我们可以进一步确认该具体型号是否在 UL Listing 覆盖范围内。",
            "Our fire-protection butterfly valves are evaluated according to UL 1091. Key compliance items include body strength, seat leakage, operating performance, durability, construction, marking, and installation instructions. Please provide the required size, pressure rating, connection type, and material so that we can confirm whether the exact model is covered by the UL Listing.",
            "示例数据仅用于模板演示。真实数值、条款号、页码需由合法持有的标准资料人工录入。",
        ),
        (
            product_ids["蝶阀"],
            standard_ids["UL 1091"],
            PENDING,
            PENDING,
            "密封相关条款",
            "密封试验 Leakage Test",
            "阀座密封试验",
            "Seat Leakage Test",
            PENDING,
            PENDING,
            PLACEHOLDER_VALUE,
            PLACEHOLDER_VALUE,
            PENDING,
            "请录入短摘录或关键词，避免保存标准全文。",
            "test report; marking sample; installation instructions",
            "我们的消防蝶阀按 UL 1091 进行评估。请以具体型号、尺寸、压力等级和连接方式核对 Listing 覆盖范围。",
            "Our fire-protection butterfly valves are evaluated according to UL 1091. Please confirm the exact model, size, pressure rating, and connection type against the Listing coverage.",
            "不要用样本系列证书替代具体型号覆盖确认。",
        ),
        (
            product_ids["水流指示器"],
            standard_ids["UL 346"],
            PENDING,
            PENDING,
            "动作性能相关条款",
            "操作试验 Operation Test",
            "动作性能试验",
            "Operation Performance Test",
            PENDING,
            PENDING,
            PLACEHOLDER_VALUE,
            PLACEHOLDER_VALUE,
            PENDING,
            "请录入短摘录或关键词，避免保存标准全文。",
            "UL Certificate; UL File Number; Product iQ listing; installation instructions; wiring diagram",
            "我们的水流指示器按 UL 346 进行评估。请提供管径范围、安装方式、延迟设置、电气参数和具体型号，以便确认 Listing 覆盖范围。",
            "Our waterflow indicators are evaluated according to UL 346. Please provide pipe size range, installation method, retard setting, electrical rating, and exact model for Listing coverage confirmation.",
            "需确认型号、管径、安装方向和电气额定值是否均被覆盖。",
        ),
        (
            product_ids["湿式报警阀"],
            standard_ids["UL 193"],
            PENDING,
            PENDING,
            "报警阀性能相关条款",
            "操作试验 Operation Test",
            "报警阀动作性能试验",
            "Alarm Valve Operation Test",
            PENDING,
            PENDING,
            PLACEHOLDER_VALUE,
            PLACEHOLDER_VALUE,
            PENDING,
            "请录入短摘录或关键词，避免保存标准全文。",
            "UL Certificate; UL File Number; trim drawing; installation instructions; marking sample",
            "我们的湿式报警阀按对应 UL 标准进行评估。请提供阀门尺寸、压力等级、配套附件和安装配置，以便核对具体 Listing 范围。",
            "Our wet alarm valves are evaluated according to the applicable UL standard. Please provide valve size, pressure rating, trim accessories, and installation configuration for Listing coverage verification.",
            "报警阀与配套附件、延迟器、水力警铃等需分别核对覆盖范围。",
        ),
    ]
    conn.executemany(
        """
        INSERT INTO requirements
        (product_id, standard_id, clause_no, page_no, section_title, requirement_category,
         test_item_cn, test_item_en, applicable_condition, test_condition, specific_value,
         pass_criteria, fail_criteria, source_excerpt_short, supplier_documents,
         sales_explanation_cn, sales_explanation_en, risk_note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        requirements,
    )

    checklist_items = [
        ("是否提供 UL Certificate", "UL Certificate provided", "核认证书与产品型号是否一致。", 1),
        ("是否提供 UL File Number", "UL File Number provided", "用于追溯 UL Listing。", 1),
        ("是否可在 UL Product iQ 查询", "Searchable in UL Product iQ", "确认认证状态和覆盖范围。", 1),
        ("具体型号是否被覆盖", "Exact model covered", "不能只确认系列名。", 1),
        ("尺寸范围是否被覆盖", "Size range covered", "核对客户要求尺寸。", 2),
        ("压力等级是否被覆盖", "Pressure rating covered", "核对设计压力和证书范围。", 2),
        ("连接方式是否被覆盖", "Connection type covered", "如沟槽、法兰、螺纹等。", 2),
        ("材质是否被覆盖", "Material covered", "核对阀体、阀瓣、密封材料。", 2),
        ("是否带 supervisory switch", "Supervisory switch included if required", "如客户需要信号反馈，需确认开关配置。", 2),
        ("证书状态是否仍然有效", "Certificate status active", "报价前建议复核。", 1),
        ("是否有安装说明书", "Installation instructions available", "用于项目提交和安装合规。", 3),
        ("是否有测试报告", "Test report available", "供应商内部或第三方资料。", 3),
        ("是否有铭牌/marking 样张", "Marking sample available", "核对产品标识。", 3),
    ]
    for product_id in product_ids.values():
        conn.executemany(
            """
            INSERT INTO supplier_checklists
            (product_id, checklist_item_cn, checklist_item_en, explanation, priority)
            VALUES (?, ?, ?, ?, ?)
            """,
            [(product_id, *item) for item in checklist_items],
        )


def fetch_dataframe(query: str, params: Iterable[Any] = ()) -> list[sqlite3.Row]:
    with get_connection() as conn:
        return conn.execute(query, tuple(params)).fetchall()


def delete_all_data() -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM supplier_checklists")
        conn.execute("DELETE FROM requirements")
        conn.execute("DELETE FROM standards")
        conn.execute("DELETE FROM products")


def delete_requirement(requirement_id: int) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM requirements WHERE id = ?", (requirement_id,))
