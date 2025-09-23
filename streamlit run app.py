# -*- coding: utf-8 -*-
# ============================================================================== 
# MRP BOM Analysis - UI Enhanced & State-Preserving Version (with Child Qty Support)
# Developed by: Reda Roshdy
# ==============================================================================

import streamlit as st
import pandas as pd
from io import BytesIO

# --- إعداد الصفحة ---
st.set_page_config(page_title="MRP BOM Analysis", layout="wide")
st.title("🚀 أداة تحليل BOM مع دعم كميات الأبناء")
st.markdown("---")

# ============================================================================== 
# 🔹 0. تهيئة Session State
# ==============================================================================
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
    st.session_state.summary_df = pd.DataFrame()
    st.session_state.top10_global = pd.DataFrame()
    st.session_state.per_parent_topdev = {}
    st.session_state.all_merged_df = pd.DataFrame()
    st.session_state.output_excel = BytesIO()

# ============================================================================== 
# 🔹 1. الشريط الجانبي للإعدادات
# ==============================================================================
st.sidebar.header("⚙️ 1. إعدادات التحليل")
uploaded_file = st.sidebar.file_uploader("⬆️ ارفع ملف Excel", type=["xlsx"])

if uploaded_file is None:
    st.info("👋 يرجى رفع ملف Excel من الشريط الجانبي لبدء التحليل.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 2. اختر الشيتات")
    default_bom = sheets.index("Bom") if "Bom" in sheets else 0
    bom_sheet = st.sidebar.selectbox("اختر شيت الـ BOM", options=sheets, index=default_bom)

    father_options = ["None"] + sheets
    default_father = 1 + sheets.index("father code") if "father code" in sheets else 0
    father_sheet = st.sidebar.selectbox("اختر شيت الـ Father", options=father_options, index=default_father)

    mrp_options = ["None"] + sheets
    default_mrp = 1 + sheets.index("MRP Control") if "MRP Control" in sheets else 0
    mrp_sheet = st.sidebar.selectbox("اختر شيت MRP Control (اختياري)", options=mrp_options, index=default_mrp)

    # قراءة البيانات
    bom_df = pd.read_excel(uploaded_file, sheet_name=bom_sheet)
    father_df = pd.read_excel(uploaded_file, sheet_name=father_sheet) if father_sheet != "None" else None
    mrp_control_df = pd.read_excel(uploaded_file, sheet_name=mrp_sheet) if mrp_sheet != "None" else None

    # تنظيف الأعمدة
    bom_df.columns = [str(c).strip() for c in bom_df.columns]
    if father_df is not None: father_df.columns = [str(c).strip() for c in father_df.columns]
    if mrp_control_df is not None: mrp_control_df.columns = [str(c).strip() for c in mrp_control_df.columns]

    # اختيار الأعمدة
    def detect_or_choose(df, candidates, label, key_suffix):
        found_cols = [c for c in candidates if c in df.columns]
        index = list(df.columns).index(found_cols[0]) if found_cols else 0
        return st.sidebar.selectbox(f"اختر عمود '{label}'", options=list(df.columns), index=index, key=key_suffix)

    code_col = detect_or_choose(bom_df, ['Code', 'Material', 'Parent', 'Planning Material'], "Parent في BOM", "code_col")
    component_col = detect_or_choose(bom_df, ['Component', 'Item', 'Material Name'], "Component في BOM", "component_col")

    qty_col = None
    qty_candidates = [c for c in ['Qty', 'Quantity', 'Qty_Per', 'Quantity_Per'] if c in bom_df.columns]
    if qty_candidates:
        qty_col = detect_or_choose(bom_df, qty_candidates, "Qty في BOM (اختياري)", "qty_col")

    parent_col, child_col = None, None
    if father_df is not None:
        parent_col = detect_or_choose(father_df, ['Parent', 'Planning Material', 'Parent_Material'], "Parent في شيت Father", "parent_col")
        child_col = detect_or_choose(father_df, ['Material', 'Child', 'Child_Material'], "Child في شيت Father", "child_col")

    mrp_component_col, mrp_controller_col, mrp_order_type_col = None, None, None
    if mrp_control_df is not None:
        mrp_component_col = detect_or_choose(mrp_control_df, ['Component', 'Material'], "Component في MRP", "mrp_comp")
        mrp_controller_col = detect_or_choose(mrp_control_df, ['MRP_Controller', 'MFC'], "MRP Controller", "mrp_ctrl")
        mrp_order_type_col = detect_or_choose(mrp_control_df, ['Order_Type', 'Type'], "Order Type", "mrp_order")

    # فلترة Parents
    parents_available = sorted(father_df[parent_col].dropna().unique().astype(str)) if father_df is not None else []
    selected_parents = st.sidebar.multiselect("اختر Parent(s) للتحليل", options=parents_available, default=parents_available)

    order_type_filter = None
    if mrp_control_df is not None and mrp_order_type_col in mrp_control_df.columns:
        order_types = ["All"] + sorted(mrp_control_df[mrp_order_type_col].dropna().unique().astype(str))
        choice = st.sidebar.selectbox("فلترة حسب Order Type", options=order_types)
        if choice != "All": order_type_filter = choice

    # زر تشغيل التحليل
    st.sidebar.markdown("---")
    if st.sidebar.button("🚀 تشغيل التحليل", type="primary"):
        with st.spinner("⏳ جاري معالجة البيانات..."):
            # تحويل الأعمدة الرئيسية إلى نص
            bom_df[code_col] = bom_df[code_col].astype(str).str.strip()
            bom_df[component_col] = bom_df[component_col].astype(str).str.strip()
            if father_df is not None:
                father_df[parent_col] = father_df[parent_col].astype(str).str.strip()
                father_df[child_col] = father_df[child_col].astype(str).str.strip()
            if mrp_control_df is not None:
                mrp_control_df[mrp_component_col] = mrp_control_df[mrp_component_col].astype(str).str.strip()

            # تجميع BOM
            if qty_col:
                bom_grouped = bom_df.groupby(code_col).apply(
                    lambda g: dict(zip(g[component_col], g[qty_col]))
                ).to_dict()
            else:
                bom_grouped = bom_df.groupby(code_col)[component_col].apply(set).to_dict()

            # تجهيز قاموس MRP
            mrp_dict = {}
            if mrp_control_df is not None:
                mrp_dict = mrp_control_df.drop_duplicates(subset=[mrp_component_col]).set_index(mrp_component_col).to_dict(orient='index')

            summary_list, all_parents_rows, per_parent_topdev = [], [], {}
            output = BytesIO()

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for parent in selected_parents:
                    parent = str(parent).strip()
                    children = father_df[father_df[parent_col]==parent][child_col].dropna().astype(str).unique().tolist() if father_df is not None else []
                    total_children = len(children)
                    parent_components = bom_grouped.get(parent, set())

                    # ==============================
                    # معالجة كل Parent + دمج بيانات الأبناء
                    # ==============================
                    usage_rows = []
                    for comp in parent_components:
                        mrp_info = mrp_dict.get(comp, {})
                        if order_type_filter and str(mrp_info.get(mrp_order_type_col)) != order_type_filter:
                            continue

                        count = 0
                        child_usage = {}
                        for child in children:
                            child_components = bom_grouped.get(child, {})
                            if qty_col and isinstance(child_components, dict):
                                qty_value = child_components.get(comp, 0)
                            else:
                                qty_value = 1 if comp in child_components else 0
                            child_usage[child] = qty_value
                            if qty_value > 0:
                                count += 1

                        usage_pct = round(count / total_children * 100, 2) if total_children > 0 else 0.0
                        row = {
                            component_col: comp,
                            "Num_Children_with_Component": count,
                            "Total_Children": total_children,
                            "Usage_%": usage_pct,
                            "MRP_Controller": mrp_info.get(mrp_controller_col),
                            "Order_Type": mrp_info.get(mrp_order_type_col)
                        }
                        row.update(child_usage)
                        usage_rows.append(row)
#####################################################
                    # ترتيب الأعمدة الخاصة بالأبناء لضمان ظهورها بالشكل المطلوب
                    child_columns = [str(child) for child in children]  # أسماء الأبناء
                    ordered_columns = [
                        component_col,
                        "Num_Children_with_Component",
                        "Total_Children",
                        "Usage_%",
                        "MRP_Controller",
                        "Order_Type"
                    ] + child_columns

                    # إنشاء DataFrame مع ترتيب الأعمدة
                    parent_df = pd.DataFrame(usage_rows, columns=ordered_columns)

                    # كتابة شيت Parent
                    if not parent_df.empty:
                        parent_df.to_excel(writer, sheet_name=str(parent)[:31], index=False)
##########################################################################
                    parent_df = pd.DataFrame(usage_rows)
                    if not parent_df.empty:
                        parent_df.to_excel(writer, sheet_name=str(parent)[:31], index=False)
                        parent_df["Deviation"] = abs(parent_df["Num_Children_with_Component"] - (total_children/2))
                        per_parent_topdev[parent] = parent_df.sort_values("Deviation", ascending=False).head(10)
                        all_parents_rows.append(parent_df.assign(Parent=parent))

                    # ملخص Parent
                    total_comps = len(parent_df)
                    shared_comps = parent_df['Num_Children_with_Component'].gt(0).sum() if total_comps>0 else 0
                    similarity_pct = round(shared_comps / total_comps * 100, 2) if total_comps>0 else 0.0
                    summary_list.append({
                        "Parent_Code": parent,
                        "Num_Children": total_children,
                        "Total_Components": total_comps,
                        "Shared_Components": int(shared_comps),
                        "Shared_Components_%": similarity_pct
                    })

                # حفظ النتائج في Session State
                st.session_state.summary_df = pd.DataFrame(summary_list)
                st.session_state.summary_df.to_excel(writer, sheet_name="Summary_Report", index=False)

                if all_parents_rows:
                    all_merged_df = pd.concat(all_parents_rows, ignore_index=True)
                    st.session_state.all_merged_df = all_merged_df
                    st.session_state.top10_global = all_merged_df.sort_values("Deviation", ascending=False).head(10)
                    #st.session_state.top10_global.to_excel(writer, sheet_name="TopDeviation_Global", index=False)

                st.session_state.per_parent_topdev = per_parent_topdev
                st.session_state.output_excel = output
                st.session_state.analysis_complete = True

        st.success("✅ اكتمل التحليل بنجاح! يمكنك الآن تصفح النتائج.")

    # ============================================================================== 
    # 🔹 3. عرض النتائج
    # ==============================================================================
    if not st.session_state.analysis_complete:
        st.info("ℹ️ اضغط على زر 'تشغيل التحليل' لعرض النتائج.")
    else:
        st.header("📈 نتائج التحليل")
        col1, col2, col3 = st.columns(3)
        col1.metric("👨‍👩‍👧 عدد الـ Parents", len(st.session_state.summary_df))
        avg_similarity = st.session_state.summary_df['Shared_Components_%'].mean()
        col2.metric("🔄 متوسط نسبة التشابه", f"{avg_similarity:.2f}%")
        total_shared = st.session_state.summary_df['Shared_Components'].sum()
        col3.metric("🔗 إجمالي المكونات المشتركة", f"{total_shared}")

        tab1, tab2, tab3 = st.tabs(["📊 الملخص الرئيسي", "🔥 أعلى الانحرافات", "👨‍👩‍👧 تفاصيل كل Parent"])
        with tab1:
            st.subheader("ملخص أداء كل Parent")
            st.dataframe(st.session_state.summary_df)
            st.markdown("---")
            if not st.session_state.all_merged_df.empty:
                low_shared_df = st.session_state.all_merged_df[st.session_state.all_merged_df['Usage_%'] < 100].sort_values('Usage_%')
                st.subheader("📉 المكونات الأقل مشاركة عبر كل الـ Parents")
                st.dataframe(low_shared_df.head(200))
        with tab2:
            st.subheader("أعلى 10 مكونات انحرافًا على المستوى الإجمالي")
            st.dataframe(st.session_state.top10_global)
        with tab3:
            st.subheader("استعراض تفاصيل الانحراف لكل Parent")
            parents_with_dev = list(st.session_state.per_parent_topdev.keys())
            if parents_with_dev:
                chosen_parent = st.selectbox("اختر Parent لعرض تفاصيله", options=parents_with_dev)
                st.dataframe(st.session_state.per_parent_topdev.get(chosen_parent, pd.DataFrame()))
            else:
                st.warning("لا توجد بيانات انحراف لعرضها.")

        st.markdown("---")
        st.download_button(
            label="📥 تحميل التقرير الكامل (Excel)",
            data=st.session_state.output_excel.getvalue(),
            file_name="MRP_BOM_Report_Stateful.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

except Exception as e:
    st.exception(f"❌ حدث خطأ: {e}")
