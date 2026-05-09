# 消防水系统 UL 标准合格判定检索工具 MVP

这是一个本地运行的 Streamlit + SQLite MVP，用于把人工整理后的 UL 标准条款数据做成结构化检索工具。

重要原则：

- 不内置、不生成任何真实 UL 标准试验数值。
- 所有具体数值、合格判定、条款号、页码必须来自你导入的 Excel/CSV 数据。
- 数据缺失时，系统显示“待补充”或“数据库暂无该项具体数值，请补充标准条款数据”。
- `source_excerpt_short` 只用于短摘录或关键词，不建议保存标准全文。

## 项目结构

```text
.
├── app.py
├── database.py
├── document_search.py
├── import_utils.py
├── search_utils.py
├── template_utils.py
├── requirements.txt
└── README.md
```

运行后会自动创建本地数据库：

```text
data/ul_fire_system.db
```

## 安装

建议使用 Python 3.10 或以上版本。

```bash
pip install -r requirements.txt
```

## 启动

```bash
streamlit run app.py
```

启动后浏览器会打开本地页面。

## 导入 Excel / CSV

1. 打开左侧菜单“模板下载”。
2. 下载 Excel 模板。
3. 按模板字段录入 UL1091、UL346、UL753 等标准的条款数据。
4. 打开左侧菜单“数据导入”。
5. 上传 `.xlsx`、`.xls` 或 `.csv` 文件。
6. 预览并校验字段。
7. 点击“导入数据库”。

导入逻辑会自动创建或复用产品、标准，再把条款要求写入 `requirements` 表。

## 搜索“蝶阀”

1. 打开左侧菜单“首页搜索”。
2. 在搜索框输入：

```text
蝶阀
```

也可以输入：

```text
butterfly valve
UL1091
DN100 300psi butterfly valve
```

系统会基于产品中文名、英文名、aliases、标准号、条款字段做模糊匹配，并按相关度排序。

## 检索桌面认证标准文件

左侧菜单打开“文件检索”，点击“重新索引认证标准文件夹”。默认读取：

```text
C:\Users\Administrator\Desktop\认证标准
```

支持文件类型：

- PDF
- DOCX
- TXT / Markdown / CSV

索引完成后可以搜索：

```text
蝶阀
butterfly valve
UL753
水力警铃
alarm valve
```

搜索结果会显示文件名、页码或片段位置、短命中片段和本地文件路径。为了避免传播标准全文，页面默认只展示短片段；如需核对上下文，可展开单条短片段查看。

## 标准文件拆条工作台

左侧打开“文件检索”，展开“标准文件拆条工作台”。

这个视图会把每个已索引标准文件按以下结构拆成候选片段：

```text
标准文件 → 部件/对象 → 要求类型 → 候选短片段 → 来源位置
```

示例对象：

- 阀体 / Valve Body
- 阀盖 / Cover
- 法兰 / Flange
- 螺纹 / Thread
- 管路连接点 / Pipe Connection
- 阀瓣 / Clapper
- 阀瓣支撑 / Clapper Support
- 阀瓣挡块 / Clapper Stop
- 阀座 / Seat
- 标识 / Marking
- 安装说明 / Installation Instructions

拆条结果只是“候选依据”，不会自动写入合格判定字段。你可以下载候选 CSV，人工确认条款号、页码、试验条件、具体数值和合格判定后，再按模板导入结构化数据库。

## 扩展到更多标准

后续扩展 UL346、UL753、UL193、UL262、UL312 等标准时，推荐流程：

1. 在模板中新增对应产品和标准，例如水流指示器 / UL 346。
2. 人工从合法持有的标准文件中整理条款号、页码、试验条件、具体数值和合格判定。
3. 每一行代表一个结构化要求。
4. 导入 Excel 后即可检索。
5. 如需供应商审核 checklist，可在导入数据中填写 `supplier_documents`，也可在数据库种子或后续管理功能中扩展。

## 数据字段

模板字段：

```text
id, product_cn, product_en, aliases, product_category, standard_no,
standard_title, standard_version, clause_no, page_no, section_title,
requirement_category, test_item_cn, test_item_en, applicable_condition,
test_condition, specific_value, pass_criteria, fail_criteria,
source_excerpt_short, supplier_documents, sales_explanation_cn,
sales_explanation_en, risk_note, created_at, updated_at
```

必填建议：

- `product_cn` 或 `product_en` 至少一个
- `standard_no`
- `test_item_cn` 或 `test_item_en` 至少一个

## 注意

本工具只适合作为公司内部学习和工作辅助。请不要把 UL 标准全文导入数据库，也不要把它作为公开传播或销售标准内容的平台。
