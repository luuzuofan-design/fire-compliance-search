from __future__ import annotations

import pandas as pd
import streamlit as st

from clause_table import clause_table_summary, index_clause_table, search_clause_table
from database import delete_all_data, delete_requirement, fetch_dataframe, init_db
from document_search import (
    DEFAULT_DOCS_DIR,
    DEFAULT_DOCS_DIRS,
    QUICK_SEARCH_TERMS,
    document_index_summary,
    find_bilingual_counterpart,
    index_default_document_folders,
    index_documents,
    language_of,
    list_indexed_files,
    offline_translate_en_to_cn,
    search_documents,
    split_file_candidates,
    split_standard_candidates,
)
from import_utils import import_requirements, read_uploaded_file, validate_import_dataframe
from search_utils import get_checklist, get_filter_values, search_requirements
from template_utils import generate_template


st.set_page_config(
    page_title="消防水系统 UL 标准合格判定检索工具",
    page_icon="UL",
    layout="wide",
)

init_db()
index_clause_table()


DISPLAY_COLUMNS = {
    "component_cn": "部件/对象",
    "component_en": "Component",
    "requirement_type": "要求类型",
    "test_item_cn": "试验项目",
    "test_item_en": "英文试验名称",
    "clause_no": "条款号",
    "page_no": "页码",
    "test_condition": "试验条件",
    "specific_value": "具体数值",
    "pass_criteria": "合格判定",
    "fail_criteria": "不合格判定",
    "applicable_condition": "适用产品条件",
    "risk_note": "备注/风险点",
}


def safe_text(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return text if text else "待补充"


def rows_to_df(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    return df


def render_document_results(rows: list[dict], title: str = "本地文件匹配", query: str = "") -> None:
    if not rows:
        return
    st.subheader(title)
    st.caption("以下来自桌面标准文件夹的本地索引。中英文对照只使用本地文件片段，不调用翻译 API，也不自动生成标准内容。")
    for row in rows:
        with st.container(border=True):
            top1, top2, top3 = st.columns([3, 1, 1])
            top1.markdown(f"**{row['file_name']}**")
            top2.markdown(f"**位置：** {safe_text(row.get('location_label'))}")
            top3.markdown(f"**相关度：** {row['score']}")
            st.caption(f"{safe_text(row.get('source_folder'))} · {row['file_path']}")
            counterpart = find_bilingual_counterpart(query, row) if query else None
            lang = row.get("language") or language_of(row.get("content", ""))
            cn_row = row if lang == "cn" else counterpart
            en_row = row if lang == "en" else counterpart

            cn_col, en_col = st.columns(2)
            with cn_col:
                st.markdown("**中文片段**")
                if cn_row:
                    st.caption(f"{cn_row['file_name']} · {safe_text(cn_row.get('location_label'))}")
                    st.write(cn_row.get("content") or cn_row.get("snippet"))
                elif en_row:
                    st.caption("本地辅助译文，仅供阅读，以英文原文为准")
                    st.warning(offline_translate_en_to_cn(en_row.get("snippet") or en_row.get("content")))
                else:
                    st.info("本地索引中暂未找到对应中文片段。")
            with en_col:
                st.markdown("**English snippet**")
                if en_row:
                    st.caption(f"{en_row['file_name']} · {safe_text(en_row.get('location_label'))}")
                    st.write(en_row.get("content") or en_row.get("snippet"))
                else:
                    st.info("No extractable English snippet was found in the local English-file index for the same standard.")

            with st.expander("查看当前命中完整文本片段"):
                st.write(row["content"])


def render_clause_table_results(rows: list[dict], title: str = "UL/FM 分级条款表匹配") -> None:
    if not rows:
        return
    st.subheader(title)
    st.caption("以下来自桌面《UL_FM消防标准分级条款表_完整版.xlsx》。用于快速定位分级条款，正式引用仍以原标准文件为准。")
    table = pd.DataFrame(rows)[
        [
            "standard_file",
            "level1",
            "level2",
            "level3",
            "content",
            "source_location",
            "score",
        ]
    ].rename(
        columns={
            "standard_file": "标准文件",
            "level1": "部件/对象/一级标题",
            "level2": "要求类型/二级标题",
            "level3": "三级标题",
            "content": "候选短片段/内容",
            "source_location": "来源位置",
            "score": "相关度",
        }
    )
    st.dataframe(table, use_container_width=True, hide_index=True)

    with st.expander("按层级查看：标准文件 → 部件/对象 → 要求类型 → 候选短片段 → 来源位置", expanded=False):
        df = pd.DataFrame(rows)
        for standard_file, standard_df in df.groupby("standard_file", sort=False):
            st.markdown(f"**标准文件：{safe_text(standard_file)}**")
            for level1, level1_df in standard_df.groupby("level1", sort=False):
                st.markdown(f"- 部件/对象：{safe_text(level1)}")
                for level2, level2_df in level1_df.groupby("level2", sort=False):
                    st.markdown(f"  - 要求类型：{safe_text(level2)}")
                    level_table = level2_df[["level3", "content", "source_location"]].rename(
                        columns={
                            "level3": "三级标题",
                            "content": "候选短片段/内容",
                            "source_location": "来源位置",
                        }
                    )
                    st.dataframe(level_table, use_container_width=True, hide_index=True)


def render_detail(row: dict) -> None:
    with st.expander(f"查看详情：{safe_text(row.get('test_item_cn'))} / {safe_text(row.get('test_item_en'))}"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**产品中文名：** {safe_text(row.get('product_cn'))}")
            st.markdown(f"**产品英文名：** {safe_text(row.get('product_en'))}")
            st.markdown(f"**上级总成：** {safe_text(row.get('parent_product_cn'))} / {safe_text(row.get('parent_product_en'))}")
            st.markdown(f"**部件/对象：** {safe_text(row.get('component_cn'))} / {safe_text(row.get('component_en'))}")
            st.markdown(f"**要求对象：** {safe_text(row.get('requirement_object'))}")
            st.markdown(f"**要求类型：** {safe_text(row.get('requirement_type'))}")
            st.markdown(f"**UL 标准号：** {safe_text(row.get('standard_no'))}")
            st.markdown(f"**标准标题：** {safe_text(row.get('standard_title'))}")
            st.markdown(f"**标准版本：** {safe_text(row.get('standard_version'))}")
            st.markdown(f"**条款号：** {safe_text(row.get('clause_no'))}")
            st.markdown(f"**页码：** {safe_text(row.get('page_no'))}")
            st.markdown(f"**章节标题：** {safe_text(row.get('section_title'))}")
        with col2:
            st.markdown(f"**试验项目：** {safe_text(row.get('test_item_cn'))} / {safe_text(row.get('test_item_en'))}")
            st.markdown(f"**适用条件：** {safe_text(row.get('applicable_condition'))}")
            st.markdown(f"**试验条件：** {safe_text(row.get('test_condition'))}")
            st.markdown(f"**具体数值：** {safe_text(row.get('specific_value'))}")
            st.markdown(f"**合格判定：** {safe_text(row.get('pass_criteria'))}")
            st.markdown(f"**不合格判定：** {safe_text(row.get('fail_criteria'))}")
            st.markdown(f"**供应商需要提供的资料：** {safe_text(row.get('supplier_documents'))}")
            st.markdown(f"**风险提醒：** {safe_text(row.get('risk_note'))}")
        st.markdown("**中文销售解释**")
        st.info(safe_text(row.get("sales_explanation_cn")))
        st.markdown("**英文销售话术**")
        st.info(safe_text(row.get("sales_explanation_en")))

        pending_fields = [
            row.get("clause_no"),
            row.get("page_no"),
            row.get("test_condition"),
            row.get("specific_value"),
            row.get("pass_criteria"),
        ]
        has_pending = any("待补充" in safe_text(value) or "数据库暂无" in safe_text(value) for value in pending_fields)
        if has_pending:
            candidate_query = " ".join(
                safe_text(row.get(key))
                for key in ["standard_no", "product_cn", "product_en", "test_item_cn", "test_item_en", "section_title"]
            )
            candidate_rows = search_documents(candidate_query, limit=3)
            if candidate_rows:
                with st.expander("本地文件候选依据：用于人工补充该条结构化字段"):
                    render_document_results(candidate_rows, "候选片段", query=candidate_query)
            else:
                st.caption("该条仍需人工从标准文件中定位并补充条款数据。")


def page_search() -> None:
    st.title("消防水系统 UL 标准合格判定检索工具")
    st.caption("本地结构化检索工具。所有数值和判定均来自导入数据；缺失项显示待补充。")

    filters = get_filter_values()
    query = st.text_input("搜索关键词", placeholder="蝶阀 / butterfly valve / UL1091 / DN100 300psi butterfly valve")
    c1, c2, c3 = st.columns(3)
    product_category = c1.selectbox("产品类别筛选", [""] + filters["product_categories"], format_func=lambda x: x or "全部")
    standard_no = c2.selectbox("标准号筛选", [""] + filters["standard_nos"], format_func=lambda x: x or "全部")
    requirement_category = c3.selectbox("试验类别筛选", [""] + filters["requirement_categories"], format_func=lambda x: x or "全部")

    rows = search_requirements(query, product_category, standard_no, requirement_category)
    st.divider()
    clause_rows = search_clause_table(query, limit=30) if query.strip() else []
    if clause_rows:
        render_clause_table_results(clause_rows)
        st.divider()

    if not rows:
        st.warning("结构化数据库没有找到匹配结果。正在显示本地标准文件中的匹配片段。")
        doc_rows = search_documents(query, limit=30) if query.strip() else []
        if doc_rows:
            render_document_results(doc_rows, "本地标准文件匹配", query=query)
        else:
            st.info("本地文件索引也没有命中。可以到“文件检索”页面点击“重新索引标准文件夹”后再试。")
        return

    df = rows_to_df(rows)
    grouped = df.groupby(["product_id", "standard_id"], sort=False)
    for (_, _), group in grouped:
        first = group.iloc[0].to_dict()
        with st.container(border=True):
            h1, h2, h3 = st.columns([2, 2, 1])
            h1.subheader(f"{safe_text(first.get('product_cn'))} / {safe_text(first.get('product_en'))}")
            h2.markdown(f"**对应标准：** {safe_text(first.get('standard_no'))}")
            h3.metric("匹配条款", len(group))

            st.markdown(f"**标准名称：** {safe_text(first.get('standard_title'))}")
            st.markdown(f"**标准版本：** {safe_text(first.get('standard_version'))}")
            st.markdown(f"**适用范围：** {safe_text(first.get('scope_summary'))}")
            st.markdown(f"**核心风险提醒：** {safe_text(first.get('risk_note'))}")

            table = group[list(DISPLAY_COLUMNS.keys())].rename(columns=DISPLAY_COLUMNS)
            st.dataframe(table, use_container_width=True, hide_index=True)

            st.markdown("**供应商审核 Checklist**")
            checklist = get_checklist(int(first["product_id"]))
            if checklist:
                checklist_df = pd.DataFrame(checklist).rename(
                    columns={
                        "checklist_item_cn": "检查项",
                        "checklist_item_en": "英文",
                        "explanation": "说明",
                        "priority": "优先级",
                    }
                )
                st.dataframe(checklist_df, use_container_width=True, hide_index=True)
            else:
                st.info("暂无 checklist，请在数据库中补充。")

            st.markdown("**销售话术**")
            st.markdown("中文：")
            st.info(safe_text(first.get("sales_explanation_cn")))
            st.markdown("English:")
            st.info(safe_text(first.get("sales_explanation_en")))

            for row in group.to_dict(orient="records"):
                render_detail(row)

            render_component_requirement_candidates(safe_text(first.get("standard_no")), query or safe_text(first.get("product_en")))

    if query.strip():
        doc_rows = search_documents(query, limit=10)
        render_document_results(doc_rows, "相关本地文件片段", query=query)


def page_import() -> None:
    st.title("数据导入")
    st.caption("上传按模板整理好的 Excel / CSV。导入不会生成任何 UL 标准数值。")
    uploaded = st.file_uploader("上传文件", type=["xlsx", "xls", "csv"])
    if not uploaded:
        return

    try:
        df = read_uploaded_file(uploaded)
    except Exception as exc:
        st.error(f"读取文件失败：{exc}")
        return

    st.subheader("数据预览")
    st.dataframe(df.head(50), use_container_width=True)

    valid, errors = validate_import_dataframe(df)
    if errors:
        st.error("校验失败")
        for error in errors[:30]:
            st.write("- " + error)
        if len(errors) > 30:
            st.write(f"还有 {len(errors) - 30} 条错误未显示。")
        return

    st.success("字段校验通过，可以导入。")
    if st.button("导入数据库", type="primary"):
        try:
            stats = import_requirements(df)
            st.success(
                f"导入成功：要求 {stats['requirements']} 条，涉及产品 {stats['products']} 个，标准 {stats['standards']} 个。"
            )
            st.cache_data.clear()
        except Exception as exc:
            st.error(f"导入失败：{exc}")


def page_document_search() -> None:
    st.title("标准文件检索")
    st.caption("检索桌面标准文件夹中的全部 PDF / DOCX。不是只限蝶阀、水流指示器、湿式报警阀；任何中文或英文关键词都可以搜。")

    summary = document_index_summary()
    clause_summary = clause_table_summary()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("已索引文件", summary["files"])
    c2.metric("文本片段", summary["chunks"])
    c3.metric("最近索引", summary["last_indexed"])
    c4.metric("分级条款表", clause_summary["rows"])

    if DEFAULT_DOCS_DIRS:
        st.markdown("**默认文件夹：**")
        for folder in DEFAULT_DOCS_DIRS:
            st.markdown(f"- `{folder}`")
    else:
        st.markdown(f"**默认文件夹：** `{DEFAULT_DOCS_DIR}`")
    if st.button("重新索引标准文件夹", type="primary"):
        with st.spinner("正在读取 PDF / DOCX 并建立本地索引..."):
            stats = index_default_document_folders()
        st.success(f"索引完成：文件夹 {stats['folders']} 个，文件 {stats['files']} 个，文本片段 {stats['chunks']} 条。")
        if stats["errors"]:
            st.warning("部分文件索引失败：")
            for error in stats["errors"]:
                st.write("- " + error)
    if st.button("重新索引 UL/FM 分级条款表"):
        stats = index_clause_table(force=True)
        st.success(f"{stats['reason']}：{stats['rows']} 行。")

    with st.expander("标准文件拆条工作台", expanded=False):
        render_split_candidates_workbench()

    st.divider()
    quick = st.selectbox(
        "常用快捷搜索",
        [""] + list(QUICK_SEARCH_TERMS.keys()),
        format_func=lambda x: x or "不使用快捷项",
    )
    default_query = QUICK_SEARCH_TERMS.get(quick, "")
    query = st.text_input(
        "文件关键词搜索",
        placeholder="任意关键词：蝶阀 / butterfly valve / UL753 / 水力警铃 / hydrostatic / marking",
    )
    effective_query = query.strip() or default_query
    file_type = st.selectbox("文件类型", ["", ".pdf", ".docx"], format_func=lambda x: x or "全部")
    limit = st.slider("最多显示结果", min_value=10, max_value=200, value=50, step=10)

    if not effective_query:
        st.info("请输入关键词。第一次使用请先点击“重新索引标准文件夹”。")
        return

    rows = search_documents(effective_query, file_type=file_type, limit=limit)
    if not rows:
        st.warning("没有找到匹配片段。可以尝试重新索引，或换用英文/中文别名搜索。")
        return

    st.success(f"找到 {len(rows)} 条相关片段。")
    render_document_results(rows, "文件搜索结果", query=effective_query)


def render_split_candidates_workbench() -> None:
    st.subheader("按文件拆分候选")
    st.caption("把每个已索引标准文件按部件/对象/要求类型拆成候选片段。候选结果只用于人工确认，不自动写入合格判定。")
    files = list_indexed_files()
    if not files:
        st.info("尚未索引文件。请先点击重新索引。")
        return

    options = {"全部文件": None}
    options.update({f"{item['file_name']}": item["id"] for item in files})
    selected = st.selectbox("选择要拆分的文件", list(options.keys()))
    candidates = split_file_candidates(options[selected], limit=500)
    if not candidates:
        st.warning("该文件暂未识别到部件级候选片段。可以后续扩充关键词规则。")
        return

    df = pd.DataFrame(candidates)
    show_cols = [
        "file_name",
        "component_cn",
        "component_en",
        "requirement_type_cn",
        "requirement_type",
        "source_excerpt_short",
        "location_label",
        "source_folder",
    ]
    st.dataframe(
        df[show_cols].rename(
            columns={
                "file_name": "标准文件",
                "component_cn": "部件/对象",
                "component_en": "Component",
                "requirement_type_cn": "要求类型",
                "requirement_type": "Requirement Type",
                "source_excerpt_short": "候选短片段",
                "location_label": "来源位置",
                "source_folder": "来源文件夹",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "下载拆条候选 Excel",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name="standard_split_candidates.csv",
        mime="text/csv",
    )

    with st.expander("按层级查看：标准文件 → 部件/对象 → 要求类型 → 候选短片段 → 来源位置", expanded=True):
        for file_name, file_df in df.groupby("file_name", sort=True):
            st.markdown(f"**标准文件：{file_name}**")
            for (component_cn, component_en), component_df in file_df.groupby(["component_cn", "component_en"], sort=True):
                st.markdown(f"- 部件/对象：{component_cn} / {component_en}")
                for (type_cn, type_en), type_df in component_df.groupby(["requirement_type_cn", "requirement_type"], sort=True):
                    st.markdown(f"  - 要求类型：{type_cn} / {type_en}")
                    level_table = type_df[["source_excerpt_short", "location_label", "source_folder"]].rename(
                        columns={
                            "source_excerpt_short": "候选短片段",
                            "location_label": "来源位置",
                            "source_folder": "来源文件夹",
                        }
                    )
                    st.dataframe(level_table, use_container_width=True, hide_index=True)


def render_component_requirement_candidates(standard_no: str, query: str) -> None:
    candidates = split_standard_candidates(standard_no, limit=250)
    if not candidates:
        return

    df = pd.DataFrame(candidates)
    st.subheader("部件级要求/测试条件候选")
    st.caption("以下是从对应标准文件中自动拆出的候选片段。请人工核对条款号、页码、数值和合格判定后，再导入结构化数据库。")

    component_options = [""] + sorted(df["component_cn"].dropna().unique().tolist())
    type_options = [""] + sorted(df["requirement_type_cn"].dropna().unique().tolist())
    c1, c2 = st.columns(2)
    component_filter = c1.selectbox(
        "部件筛选",
        component_options,
        format_func=lambda x: x or "全部",
        key=f"component_filter_{standard_no}_{query}",
    )
    type_filter = c2.selectbox(
        "要求类型筛选",
        type_options,
        format_func=lambda x: x or "全部",
        key=f"type_filter_{standard_no}_{query}",
    )

    filtered = df
    if component_filter:
        filtered = filtered[filtered["component_cn"] == component_filter]
    if type_filter:
        filtered = filtered[filtered["requirement_type_cn"] == type_filter]

    summary = (
        filtered.groupby(["component_cn", "component_en"], dropna=False)
        .agg(
            候选要求数=("component_cn", "size"),
            要求类型=("requirement_type_cn", lambda values: " / ".join(sorted(set(str(v) for v in values if str(v))))),
            来源文件=("file_name", lambda values: " / ".join(sorted(set(str(v) for v in values if str(v)))[:3])),
        )
        .reset_index()
        .rename(columns={"component_cn": "部件", "component_en": "Component"})
    )
    st.markdown("**部件汇总**")
    st.caption("同一个部件会在不同条款、不同要求类型中重复出现。这里先按部件合并，数量表示该部件下有多少条候选要求。")
    st.dataframe(summary, use_container_width=True, hide_index=True)

    component_labels = [
        f"{row['部件']} / {row['Component']} ({row['候选要求数']} 条)"
        for _, row in summary.iterrows()
    ]
    component_lookup = {
        label: (row["部件"], row["Component"])
        for label, (_, row) in zip(component_labels, summary.iterrows())
    }
    selected_label = st.selectbox(
        "点击/选择部件查看对应文本片段",
        component_labels,
        key=f"selected_component_{standard_no}_{query}",
    )
    selected_cn, selected_en = component_lookup[selected_label]
    detail_df = filtered[(filtered["component_cn"] == selected_cn) & (filtered["component_en"] == selected_en)]

    st.markdown(f"**{selected_cn} / {selected_en} 的全部候选要求**")
    detail_table = detail_df[
        ["source_folder", "file_name", "location_label", "requirement_type_cn", "requirement_type", "source_excerpt_short"]
    ].rename(
        columns={
            "source_folder": "来源文件夹",
            "file_name": "标准文件",
            "location_label": "来源位置",
            "requirement_type_cn": "要求/试验类型",
            "requirement_type": "Type",
            "source_excerpt_short": "候选短片段",
        }
    )
    st.dataframe(detail_table, use_container_width=True, hide_index=True)

    for idx, item in detail_df.iterrows():
        with st.expander(f"{item['requirement_type_cn']} · {item['location_label']} · {item['file_name']}"):
            lang = item.get("language") or language_of(item.get("candidate_snippet", ""))
            text = item.get("candidate_snippet") or item.get("content") or ""
            if lang == "en":
                cn_col, en_col = st.columns(2)
                with cn_col:
                    st.markdown("**中文辅助译文**")
                    st.warning(offline_translate_en_to_cn(text))
                with en_col:
                    st.markdown("**English original**")
                    st.write(text)
            else:
                st.markdown("**中文原文片段**")
                st.write(text)
                counterpart = find_bilingual_counterpart(
                    f"{standard_no} {item['component_en']} {item['requirement_type']}",
                    item.to_dict(),
                )
                if counterpart:
                    st.markdown("**English counterpart candidate**")
                    st.caption(f"{counterpart['file_name']} · {counterpart['location_label']}")
                    st.write(counterpart.get("snippet") or counterpart.get("content"))

    st.download_button(
        "下载该标准部件拆条候选 CSV",
        data=filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{standard_no.replace(' ', '')}_component_candidates.csv",
        mime="text/csv",
        key=f"download_candidates_{standard_no}_{query}",
    )


def page_template() -> None:
    st.title("模板下载")
    st.write("下载标准 Excel 模板后，人工整理条款数据再导入。")
    st.download_button(
        "下载 Excel 模板",
        data=generate_template(),
        file_name="ul_fire_system_requirements_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


def page_manage() -> None:
    st.title("数据管理")
    tab1, tab2, tab3, tab4 = st.tabs(["产品", "标准", "要求", "清空数据"])

    with tab1:
        rows = fetch_dataframe("SELECT * FROM products ORDER BY id")
        st.dataframe(pd.DataFrame([dict(r) for r in rows]), use_container_width=True, hide_index=True)

    with tab2:
        rows = fetch_dataframe("SELECT * FROM standards ORDER BY id")
        st.dataframe(pd.DataFrame([dict(r) for r in rows]), use_container_width=True, hide_index=True)

    with tab3:
        rows = fetch_dataframe(
            """
            SELECT r.id, p.product_cn, p.product_en, s.standard_no, r.clause_no, r.page_no,
                   r.component_cn, r.component_en, r.requirement_type,
                   r.requirement_category, r.test_item_cn, r.test_item_en, r.specific_value,
                   r.pass_criteria, r.risk_note
            FROM requirements r
            JOIN products p ON p.id = r.product_id
            JOIN standards s ON s.id = r.standard_id
            ORDER BY r.id
            """
        )
        df = pd.DataFrame([dict(r) for r in rows])
        st.dataframe(df, use_container_width=True, hide_index=True)
        req_id = st.number_input("删除单条要求 ID", min_value=1, step=1)
        if st.button("删除该要求"):
            delete_requirement(int(req_id))
            st.success("已删除。请刷新页面查看最新数据。")

    with tab4:
        st.warning("清空后会删除全部产品、标准、要求和 checklist。重新打开应用会自动写入示例种子数据。")
        confirm = st.text_input("如需清空，请输入 DELETE")
        if st.button("清空全部数据", type="primary") and confirm == "DELETE":
            delete_all_data()
            st.success("已清空。刷新页面后会重新初始化示例数据。")


def page_about() -> None:
    st.title("关于工具")
    st.markdown(
        """
        本工具用于公司内部学习和工作辅助，将人工整理后的 UL 标准试验要求转为可搜索的结构化数据。

        使用边界：

        - 不联网，不调用外部 API。
        - 不内置真实 UL 标准数值。
        - 不保存或传播标准全文。
        - 每条要求必须能追溯到标准号、版本、条款号和页码。
        - 缺失数据必须人工补充，系统不会自动编造。
        """
    )


PAGES = {
    "首页搜索": page_search,
    "文件检索": page_document_search,
    "数据导入": page_import,
    "模板下载": page_template,
    "数据管理": page_manage,
    "关于工具": page_about,
}

with st.sidebar:
    st.header("导航")
    page_name = st.radio("页面", list(PAGES.keys()), label_visibility="collapsed")
    st.divider()
    st.caption("内部辅助工具，不用于公开传播标准全文。")

PAGES[page_name]()
