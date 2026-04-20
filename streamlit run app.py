# -*- coding: utf-8 -*-
# ==============================================================================
# MRP BOM Analysis - UI Enhanced & State-Preserving Version (with Child Qty Support)
# Developed by: Reda Roshdy
# ==============================================================================
import streamlit as st
import pandas as pd
from io import BytesIO

def auto_detect(df, candidates):
    # دالة بسيطة تحاول تختار أول اسم عمود مطابق من قائمة المرشحين
    # لو لقيت أي اسم من candidates في df.columns هترجعه، وإلا ترجع العمود الأول كfallback
    for col in candidates:
        if col in df.columns:
            return col
    # fallback
    return df.columns[0]

# Helper: try to get a column safely بدون fallback (لو مش موجود يرجّع None)
def try_get_col(df, candidates):
    # دالة مساعدة ترجع None لو الداتا None أو مفيش أي من المرشحين
    if df is None:
        return None
    for c in candidates:
        if c in df.columns:
            return c
    return None

# --- إعداد الصفحة ---
st.set_page_config(page_title="MRP BOM Analysis", layout="wide")
st.subheader("🚀 الأبناء مع الاباء BOM أداة تحليل ")
st.markdown("---")

# ==============================================================================
# 🔹 0. تهيئة Session State
# ==============================================================================
# نهيّئ المتغيرات اللي هنخزن فيها نتائج التحليل وملفات الإخراج داخل session_state
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

# لو المستخدم ما رفعش ملف، نوقف التنفيذ ونطلب رفع الملف
if uploaded_file is None:
    st.info("👋 يرجى رفع ملف Excel من الشريط الجانبي لبدء التحليل.")
    st.stop()

try:
    # محاولة قراءة ملف الإكسل ومعرفة أسماء الشيتات المتاحة
    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    st.sidebar.markdown("---")
    st.sidebar.subheader("📄 2. اختر الشيتات")

    # اختيار شيت BOM بشكل افتراضي لو موجود، وإلا أول شيت
    default_bom = sheets.index("Bom") if "Bom" in sheets else 0
    bom_sheet = st.sidebar.selectbox("اختر شيت الـ BOM", options=sheets, index=default_bom)

    # اختيار شيت father code (يمكن "None")
    father_options = ["None"] + sheets
    default_father = 1 + sheets.index("father code") if "father code" in sheets else 0
    father_sheet = st.sidebar.selectbox("اختر شيت الـ Father", options=father_options, index=default_father)

    # اختيار شيت MRP Controller اختيارياً (تم المحافظة على الاسم الافتراضي كما في الكود الأصلي)
    mrp_options = ["None"] + sheets
    default_mrp = 1 + sheets.index("MRP Controller") if "MRP Controller" in sheets else 0
    mrp_sheet = st.sidebar.selectbox("اختر شيت MRP Controller (اختياري)", options=mrp_options, index=default_mrp)

    # قراءة البيانات من الشيتات المختارة
    bom_df = pd.read_excel(uploaded_file, sheet_name=bom_sheet)
    father_df = pd.read_excel(uploaded_file, sheet_name=father_sheet) if father_sheet != "None" else None
    mrp_control_df = pd.read_excel(uploaded_file, sheet_name=mrp_sheet) if mrp_sheet != "None" else None

    # تنظيف أسماء الأعمدة من مسافات زائدة لضمان تطابق أفضل
    bom_df.columns = [str(c).strip() for c in bom_df.columns]
    if father_df is not None:
        father_df.columns = [str(c).strip() for c in father_df.columns]
    if mrp_control_df is not None:
        mrp_control_df.columns = [str(c).strip() for c in mrp_control_df.columns]

    # -------------------------
    # اختيار الأعمدة الرئيسية باستخدام auto_detect أو try_get_col
    # -------------------------
    code_col = auto_detect(bom_df, ['Code', 'Material', 'Parent', 'Planning Material'])
    component_col = auto_detect(bom_df, ['Component', 'Item', 'Material Name'])

    # محاولة تحديد عمود الكمية إن وُجد
    qty_col = None
    qty_candidates = [c for c in ['Qty', 'Quantity', 'Component Quantity', 'Quantity_Per'] if c in bom_df.columns]
    if qty_candidates:
        qty_col = auto_detect(bom_df, qty_candidates)

    # أعمدة الأب والابن في شيت الـ father (لو موجود)
    parent_col, child_col = None, None
    if father_df is not None:
        parent_col = auto_detect(father_df, ['Parent', 'Planning Material', 'Parent_Material'])
        child_col = auto_detect(father_df, ['Material', 'Child', 'Child_Material'])

    # أعمدة من شيت MRP Control (اختيار مرن للتعامل مع أسماء مختلفة)
    mrp_component_col = None
    mrp_controller_col = None
    mrp_order_type_col = None

    if mrp_control_df is not None:
        mrp_component_col = auto_detect(mrp_control_df, ['Component', 'Material'])
        # دعم أسماء مختلفة للـ Controller
        mrp_controller_col = try_get_col(mrp_control_df, [
            'MRP_Controller', 'MRP Controller', 'MRP controller', 'MRPC', 'MFC'
        ]) or auto_detect(mrp_control_df, ['MRP_Controller', 'MFC'])
        # دعم أسماء مختلفة للـ Order Type
        mrp_order_type_col = try_get_col(mrp_control_df, [
            'Order_Type', 'Order Type', 'Order type', 'Type'
        ]) or auto_detect(mrp_control_df, ['Order_Type', 'Type'])

    # 🔸 التقاط عمود الوصف (Component Description) من BOM أو MRP
    desc_candidates = [
        'Component Description', 'Component_Description',
        'Description', 'Material Description', 'Short Text',
        'Item Description', 'Component Name', 'Material Name', 'Name'
    ]
    desc_col_bom = try_get_col(bom_df, desc_candidates)
    desc_col_mrp = try_get_col(mrp_control_df, desc_candidates) if mrp_control_df is not None else None

    # فلترة الـ Parents المتاحة في شيت father
    parents_available = sorted(father_df[parent_col].dropna().unique().astype(str)) if father_df is not None else []
    selected_parents = st.sidebar.multiselect("اختر Parent(s) للتحليل", options=parents_available, default=parents_available)

    # =============== NEW: فلاتر متعددة لـ Order Type و MRP Controller ===============
    selected_order_types = []
    selected_mrp_controllers = []

    if mrp_control_df is not None and mrp_order_type_col in mrp_control_df.columns:
        order_types_options = sorted(mrp_control_df[mrp_order_type_col].dropna().astype(str).unique().tolist())
        selected_order_types = st.sidebar.multiselect(
            "فلترة حسب Order Type (متعدد)",
            options=order_types_options,
            default=order_types_options,
            help="اتركها كما هي لعدم تضييق النتائج؛ اختر قيمًا محددة لتطبيق الفلتر."
        )

    if mrp_control_df is not None and mrp_controller_col in mrp_control_df.columns:
        mrp_ctrl_options = sorted(mrp_control_df[mrp_controller_col].dropna().astype(str).unique().tolist())
        selected_mrp_controllers = st.sidebar.multiselect(
            "فلترة حسب MRP Controller (متعدد)",
            options=mrp_ctrl_options,
            default=mrp_ctrl_options,
            help="اتركها كما هي لعدم تضييق النتائج؛ اختر قيمًا محددة لتطبيق الفلتر."
        )
    # ================================================================================

    # زر تشغيل التحليل
    st.sidebar.markdown("---")
    if st.sidebar.button("🚀 تشغيل التحليل", type="primary"):
        with st.spinner("⏳ جاري معالجة البيانات..."):
            # --- تحويل قيم الأعمدة الأساسية إلى نص وتنظيفها ---
            bom_df[code_col] = bom_df[code_col].astype(str).str.strip()
            bom_df[component_col] = bom_df[component_col].astype(str).str.strip()

            if father_df is not None:
                father_df[parent_col] = father_df[parent_col].astype(str).str.strip()
                father_df[child_col] = father_df[child_col].astype(str).str.strip()

            if mrp_control_df is not None and mrp_component_col:
                mrp_control_df[mrp_component_col] = mrp_control_df[mrp_component_col].astype(str).str.strip()

            # --- تجميع BOM حسب الـ Parent (مع دعم الكميات إن وُجدت) ---
            if qty_col:
                # لو فيه عمود كمية: نحوّل كل parent لقاموس component->qty
                bom_grouped = bom_df.groupby(code_col).apply(
                    lambda g: dict(zip(g[component_col], g[qty_col]))
                ).to_dict()
            else:
                # لو مفيش كمية: مجرد مجموعة من المكونات لكل parent
                bom_grouped = bom_df.groupby(code_col)[component_col].apply(set).to_dict()

            # تجهيز قاموس بيانات MRP لتسهيل الربط لاحقًا
            mrp_dict = {}
            if mrp_control_df is not None and mrp_component_col:
                mrp_dict = (
                    mrp_control_df
                    .drop_duplicates(subset=[mrp_component_col])
                    .set_index(mrp_component_col)
                    .to_dict(orient='index')
                )

            # قاموس الوصف للمكوّن (نملأه من MRP أو BOM إن توفّر)
            desc_lookup = {}
            if mrp_control_df is not None and mrp_component_col and desc_col_mrp:
                desc_lookup.update(
                    mrp_control_df.dropna(subset=[mrp_component_col]).drop_duplicates(subset=[mrp_component_col])
                    .set_index(mrp_component_col)[desc_col_mrp]
                    .to_dict()
                )

            if desc_col_bom:
                bom_desc_map = (
                    bom_df.dropna(subset=[component_col, desc_col_bom])
                    .drop_duplicates(subset=[component_col])
                    .set_index(component_col)[desc_col_bom]
                    .to_dict()
                )
                # أكمل أي فراغات من الـ BOM لو الـ MRP ما غطّاش كله
                for k, v in bom_desc_map.items():
                    if k not in desc_lookup and pd.notna(v):
                        desc_lookup[k] = v

            # تهيئة قوائم مساعدة لتجميع النتائج
            summary_list, all_parents_rows, per_parent_topdev = [], [], {}
            output = BytesIO()

            # نكتب النتائج في ملف Excel مؤقت داخل الذاكرة (output)
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for parent in selected_parents:
                    parent = str(parent).strip()
                    # جلب قائمة الأبناء للـ parent الحالي
                    children = father_df[father_df[parent_col] == parent][child_col].dropna().astype(str).unique().tolist() if father_df is not None else []
                    total_children = len(children)
                    parent_components = bom_grouped.get(parent, set())

                    # ==============================
                    # معالجة كل Parent + دمج بيانات الأبناء
                    # ==============================
                    usage_rows = []
                    for comp in parent_components:
                        mrp_info = mrp_dict.get(comp, {})

                        # =============== NEW: تطبيق فلاتر Order Type + MRP Controller ===============
                        # ملاحظة: لو المستخدم ما اختارش حاجة (القائمة فاضية) => ما فيش فلترة لهذا الحقل.
                        if selected_order_types:
                            if str(mrp_info.get(mrp_order_type_col)) not in set(selected_order_types):
                                continue

                        if selected_mrp_controllers:
                            if str(mrp_info.get(mrp_controller_col)) not in set(selected_mrp_controllers):
                                continue
                        # ============================================================================

                        count = 0
                        child_usage = {}
                        for child in children:
                            # جلب مكونات الابن، وإلا 0
                            child_components = bom_grouped.get(child, {})
                            if qty_col and isinstance(child_components, dict):
                                qty_value = child_components.get(comp, 0)
                            else:
                                qty_value = 1 if comp in child_components else 0

                            child_usage[child] = qty_value
                            if qty_value > 0:
                                count += 1

                        # حساب نسبة الاستخدام (نسبة الأبناء اللي استخدموا المكون)
                        usage_pct = round(count / total_children * 100, 2) if total_children > 0 else 0.0
                        row = {
                            # سنحوّل الاسم لــ 'Component' لاحقًا للتوحيد في العرض
                            component_col: comp,
                            "Component Description": desc_lookup.get(comp, ""),
                            "Total_Children": total_children,
                            "Num_Children_with_Component": count,
                            "Usage_%": usage_pct,
                            "MRP_Controller": mrp_info.get(mrp_controller_col),
                            "Order_Type": mrp_info.get(mrp_order_type_col)
                        }
                        # نضيف أعمدة كل ابن وقيمته (كمية أو صار/لم يصَر)
                        row.update(child_usage)
                        usage_rows.append(row)

                    # إنشاء DataFrame واحتساب الانحراف + ترتيب الأعمدة
                    parent_df = pd.DataFrame(usage_rows)
                    if not parent_df.empty:
                        # توحيد اسم العمود إلى 'Component' للعرض والفرز
                        if component_col != 'Component' and component_col in parent_df.columns:
                            parent_df.rename(columns={component_col: 'Component'}, inplace=True)
                        comp_col_for_display = 'Component' if 'Component' in parent_df.columns else component_col

                        # حساب الانحراف (كم يبتعد العدد عن إجمالي الأبناء)
                        parent_df["Deviation"] = abs(parent_df["Num_Children_with_Component"] - (total_children))

                        # ترتيب الأعمدة لواجهة المستخدم لعرض منطقي
                        child_columns = [str(child) for child in children]
                        first_block = [
                            "Component",
                            "Component Description",
                            "Total_Children",
                            "Num_Children_with_Component",
                            "Usage_%",
                            "Deviation",
                            "parent",
                        ]
                        # باقي الأعمدة (MRP + Order + الأبناء + أي أعمدة تانية)
                        rest_cols = [c for c in ["MRP_Controller", "Order_Type"] if c in parent_df.columns] + child_columns
                        # أضف أي أعمدة أخرى غير مذكورة (مثل Deviation، ننقله لآخر الجدول)
                        others = [c for c in parent_df.columns if c not in first_block + rest_cols]
                        ordered_columns = [c for c in first_block if c in parent_df.columns] + rest_cols + others
                        parent_df = parent_df.reindex(columns=ordered_columns)

                        # كتابة شيت Parent داخل ملف الإكسل المؤقت
                        parent_df.to_excel(writer, sheet_name=str(parent)[:31], index=False)

                        # نحفظ لأعلى الانحرافات لكل Parent (لمحة سريعة)
                        per_parent_topdev[parent] = parent_df.sort_values("Deviation", ascending=False).head(10)

                        # للاستخدام الإجمالي (التجميع لكل Parents لاحقًا)
                        all_parents_rows.append(parent_df.assign(Parent=parent))

                    # ملخص Parent (سطر واحد لكل Parent في ملخص)
                    total_comps = int(len(parent_df)) if 'parent_df' in locals() and not parent_df.empty else 0
                    shared_comps = int(parent_df['Num_Children_with_Component'].gt(0).sum()) if total_comps > 0 else 0
                    similarity_pct = round(shared_comps / total_comps * 100, 2) if total_comps > 0 else 0.0
                    summary_list.append({
                        "Parent_Code": parent,
                        "Num_Children": total_children,
                        "Total_Components": total_comps,
                        "Shared_Components": shared_comps,
                        "Shared_Components_%": similarity_pct
                    })

                # شيت الملخص: نحفظه في session_state وفي ملف الإكسل
                st.session_state.summary_df = pd.DataFrame(summary_list)
                st.session_state.summary_df.to_excel(writer, sheet_name="Summary_Report", index=False)

                # تجميعة الكل + أعلى 10
                if all_parents_rows:
                    all_merged_df = pd.concat(all_parents_rows, ignore_index=True)
                    st.session_state.all_merged_df = all_merged_df

                    # تأكيد وجود عمود Component موحّد قبل الفرز
                    if component_col != 'Component' and component_col in all_merged_df.columns:
                        all_merged_df = all_merged_df.rename(columns={component_col: 'Component'})

                    st.session_state.top10_global = all_merged_df.sort_values("Deviation", ascending=False).head(10)

                # حفظ باقي النتائج في الستيت
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

        # --- بطاقة مؤشرات سريعة بالجزء العلوي ---
        col1, col2, col3 = st.columns(3)
        col1.metric("👨‍👩‍👧 عدد الـ Parents", len(st.session_state.summary_df))
        avg_similarity = st.session_state.summary_df['Shared_Components_%'].mean()
        col2.metric("🔄 متوسط نسبة التشابه", f"{avg_similarity:.2f}%")
        total_shared = st.session_state.summary_df['Shared_Components'].sum()
        col3.metric("🔗 إجمالي المكونات المشتركة", f"{total_shared}")

        # --- تبويبات العرض ---
        tab1, tab2, tab3 = st.tabs(["📊 الملخص الرئيسي", "🔥 أعلى الانحرافات", "👨‍👩‍👧 تفاصيل كل Parent"])

        with tab1:
            st.subheader("ملخص أداء كل Parent")
#            st.dataframe(st.session_state.summary_df)
            # نعمل نسخة محلية من الـ DataFrame عشان نقدر نعدل عليها بدون تغيير الستيت مباشرة
            summary_df = st.session_state.summary_df.copy()
            if not summary_df.empty:
                # تحديد الأعمدة الرقمية (لاستخدامها في حساب المجموع/المتوسط)
                numeric_cols = summary_df.select_dtypes(include='number').columns

                # تحديد الأعمدة اللي تمثل نسب/تشابه (لو أسم العمود فيه % أو "تشابه")
                percent_cols = [c for c in numeric_cols if "%" in c or "تشابه" in c]

                # مجموع الأعمدة الرقمية (سنعتمد على مجموع للأعمدة الطبيعية)
                totals = summary_df[numeric_cols].sum()
                # متوسط للأعمدة النسبية (لأن جمع النسب غير منطقِي)
                averages = summary_df[percent_cols].mean()

                # دمج النتائج (المتوسط يحل محل المجموع لو العمود نسبي)
                for col in percent_cols:
                    totals[col] = averages[col]

                # إنشاء صف يجمع الإجماليات/المتوسطات (صف أفقي باستخدام .T)
                totals_row = pd.DataFrame(totals).T
                # وضع علامة واضحة في عمود Parent_Code للإشارة إلى أن هذا سطر إجمالي
                totals_row["Parent_Code"] = ['🔢 الإجماليات / المتوسطات']
                # دمج صف الإجماليات مع الـ DataFrame (ignore_index=True لضمان إندكس سليم)
                summary_df = pd.concat([summary_df, totals_row], ignore_index=True)


            # عرض الجدول للمستخدم (مع إخفاء الإندكس لأنه يسبب عمود فارغ)
            st.dataframe(summary_df, hide_index=True)   # ← بديل للسطر الأخير

            st.markdown("---")

            # --- قسم عرض المكونات الأقل مشاركة (إذا كانت متوفرة) ---
            if not st.session_state.all_merged_df.empty:
                low_shared_df = st.session_state.all_merged_df[st.session_state.all_merged_df['Usage_%'] < 100].sort_values('Usage_%')
                st.subheader("📉 المكونات الأقل مشاركة عبر كل الـ Parents")
                display_first = ['Parent', 'Component', 'Component Description', 'Parents', 'Total_Children', 'Num_Children_with_Component', 'Usage_%']
                cols = [c for c in display_first if c in low_shared_df.columns] + [c for c in low_shared_df.columns if c not in display_first]
                st.dataframe(low_shared_df[cols].head(200), hide_index=True)   # ← بديل للسطر الأخير)

        with tab2:
            st.subheader("أعلى 10 مكونات انحرافًا على المستوى الإجمالي")
            top10 = st.session_state.top10_global.copy()
            if not top10.empty:
                display_first = ['Parent', 'Component', 'Component Description', 'Total_Children', 'Num_Children_with_Component', 'Usage_%']
                cols = [c for c in display_first if c in top10.columns] + [c for c in top10.columns if c not in display_first]
                st.dataframe(top10[cols], hide_index=True)   # ← بديل للسطر الأخير)
            else:
                st.info("لا توجد بيانات لعرض أعلى الانحرافات.")

        with tab3:
            st.subheader("استعراض تفاصيل الانحراف لكل Parent")
            parents_with_dev = list(st.session_state.per_parent_topdev.keys())
            if parents_with_dev:
                chosen_parent = st.selectbox("اختر Parent لعرض تفاصيله", options=parents_with_dev)
                dfp = st.session_state.per_parent_topdev.get(chosen_parent, pd.DataFrame()).copy()
                if not dfp.empty:
                    display_first = ['Parent', 'Component', 'Component Description', 'Parents', 'Total_Children', 'Num_Children_with_Component', 'Usage_%']
                    cols = [c for c in display_first if c in dfp.columns] + [c for c in dfp.columns if c not in display_first]
                    st.dataframe(dfp[cols], hide_index=True)   # ← بديل للسطر الأخير)
                else:
                    st.info("لا توجد بيانات انحراف لهذا الـ Parent.")
            else:
                st.warning("لا توجد بيانات انحراف لعرضها.")

        st.markdown("---")
        # زر تحميل الإكسل النهائي
        st.download_button(
            label="🗂️  (Excel) تحميل التقرير الكامل  🔥",
            data=st.session_state.output_excel.getvalue(),
            file_name="MRP_BOM_Report_Stateful.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

except Exception as e:
    # في حالة أي خطأ، نعرضه للمستخدم لكي يسهل تتبعه
    st.exception(f"❌ حدث خطأ: {e}")

# ---------------------------
# Footer
# ---------------------------
st.markdown(
    """
    <p style="text-align:center; margin-top:30px;">
        ✨ تم التنفيذ بواسطة <b>م / رضا رشدي</b> – جميع الحقوق محفوظة © 2025 ✨
    </p>
    """,
    unsafe_allow_html=True
)

