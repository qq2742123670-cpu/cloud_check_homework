# -*- coding: utf-8 -*-
import streamlit as st
import os
import re
import pandas as pd
from pathlib import Path
import io
import zipfile
import tempfile
import shutil

# ==============================
# 0. 页面配置与 CSS 样式
# ================================
st.set_page_config(page_title="作业提交检查助手", layout="wide", page_icon="📝")
# 添加自定义 CSS
st.markdown("""
<style>
    section[data-testid="stSidebar"] {
        width: 350px;
    }
    .sub-header {
        font-size: 6rem;
        color: #555;
        margin-top: 1px !important;
        margin-bottom: 20px !important;
        border-bottom: 2px solid #A9A9A9;
    }
    .folder-item {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
        margin-top: 1px;
        margin-bottom: 6px;
        border-left: 4px solid #40E0D0;
    }
    .folder-item small {
        color: #666;
        font-size: 0.8em;
    }

</style>
""", unsafe_allow_html=True)


# =======================================
# 1. 核心逻辑函数
# ========================================
def extract_student_id_from_filename(filename):
    """从文件名中提取前9位数字作为学号"""
    match = re.search(r'\d{9}', filename)
    if match:
        return match.group()
    return None


def process_roster_file(roster_file):
    """处理花名册文件，返回结构化数据"""
    try:
        header_index = 0  # 默认表头为第0行（第一行）
        try:
            # 预读取前6行（header=None表示不指定表头，全作为数据读入）
            df_preview = pd.read_excel(uploaded_file, header=None, nrows=6)

            # 循环检查前5行
            found_header = False
            for i in range(min(5, len(df_preview))):
                # 将该行所有数据转为字符串并拼接，便于搜索
                row_values = df_preview.iloc[i].astype(str).values
                row_str = " ".join(row_values)

                # 如果该行包含关键字
                if '学号' in row_str or '姓名' in row_str:
                    header_index = i
                    print(f"在 Excel 第 {i + 1} 行检测到表头关键字，将以此行作为表头读取。")
                    found_header = True
                    break

            if not found_header:
                print("在前5行未检测到'学号'或'姓名'关键字，将默认使用第1行作为表头。")

        except Exception as pre_e:
            print(f"预扫描表头失败，将尝试默认读取: {pre_e}")
        # 使用确定的 header_index 正式读取数据
        df = pd.read_excel(uploaded_file, header=header_index)

        # 查找学号列
        student_id_col = None
        for col in df.columns:
            if '学号' in str(col):
                student_id_col = col
                break
        if student_id_col is None:
            # 备用策略：找包含9位数字的列
            for col in df.columns:
                sample_values = df[col].dropna().head(5)
                if len(sample_values) > 0:
                    has_9digit = any(re.search(r'\d{9}', str(v)) for v in sample_values)
                    if has_9digit:
                        student_id_col = col
                        break
        if student_id_col is None:
            student_id_col = df.columns[0]
            st.warning(f"未找到明确的'学号'列，使用第一列: {student_id_col}")
        else:
            st.success(f"使用学号列: {student_id_col}")

        # 查找姓名列
        name_col = None
        for col in df.columns:
            if '姓名' in str(col):
                name_col = col
                break
        if name_col is None:
            if student_id_col == df.columns[0] and len(df.columns) > 1:
                name_col = df.columns[1]
            else:
                col_index = list(df.columns).index(student_id_col)
                if col_index + 1 < len(df.columns):
                    name_col = df.columns[col_index + 1]

        if name_col:
            st.success(f"使用姓名列: {name_col}")
        else:
            st.warning("未找到姓名列，将只显示学号")

        student_id_to_name = {}
        student_ids = set()

        for idx, row in df.iterrows():
            id_value = row[student_id_col]
            if pd.isna(id_value):
                continue
            str_value = str(id_value).strip()
            student_id = None
            if str_value.isdigit() and len(str_value) >= 9:
                student_id = str_value[:9]
            else:
                match = re.search(r'\d{9}', str_value)
                if match:
                    student_id = match.group()

            if student_id:
                student_ids.add(student_id)
                name = "未知"
                if name_col and not pd.isna(row[name_col]):
                    name = str(row[name_col]).strip()
                student_id_to_name[student_id] = name

        return {
            'student_ids': student_ids,
            'student_id_to_name': student_id_to_name,
            'total_students': len(student_ids)
        }
    except Exception as e:
        st.error(f"读取花名册时出错: {e}")
        return None


def check_homework_in_folder(folder_path, roster_student_ids, target_extensions=None, check_all_types=False):
    """
    检查指定文件夹中的作业文件，支持自定义后缀筛选
    """
    try:
        # 获取文件夹下所有文件
        all_files = [f for f in Path(folder_path).iterdir() if f.is_file()]

        submitted_ids = set()
        file_type_stats = {}  # 用于统计提交的文件类型：{'.py': 10, '.docx': 2}

        for file_path in all_files:
            file_name = file_path.name
            file_ext = file_path.suffix.lower()  # 获取小写后缀，如 .py

            # 1. 提取学号
            student_id = extract_student_id_from_filename(file_name)

            if student_id:
                # 2. 判断是否符合文件类型要求
                is_valid_type = False
                if check_all_types:
                    is_valid_type = True
                elif target_extensions and file_ext in target_extensions:
                    is_valid_type = True

                # 3. 如果符合要求，计入提交名单并统计类型
                if is_valid_type:
                    submitted_ids.add(student_id)
                    # 统计该类型文件的数量
                    if file_ext in file_type_stats:
                        file_type_stats[file_ext] += 1
                    else:
                        file_type_stats[file_ext] = 1

        missing_ids = roster_student_ids - submitted_ids

        return {
            'submitted_ids': submitted_ids,
            'missing_ids': missing_ids,
            'submitted_count': len(submitted_ids),
            'missing_count': len(missing_ids),
            'file_type_stats': file_type_stats  # 新增：返回类型统计
        }
    except Exception as e:
        st.error(f"检查文件夹 {folder_path} 时出错: {e}")
        return None


# ===========================
# 2. 状态初始化
# =============================
if 'roster_data' not in st.session_state:
    st.session_state.roster_data = None
if 'student_id_to_name' not in st.session_state:
    st.session_state.student_id_to_name = {}
if 'folder_paths' not in st.session_state:
    st.session_state.folder_paths = []
if 'folder_results' not in st.session_state:
    st.session_state.folder_results = {}
if 'check_performed' not in st.session_state:
    st.session_state.check_performed = False

if 'folder_display_names' not in st.session_state:
    st.session_state.folder_display_names = {} # 新增：路径 -> 显示名称的映射

# ==========================
# 3. 侧边栏逻辑
# =============================
with st.sidebar:
    st.markdown('<h1 class="sub-header">🛠 配置选项</h1>', unsafe_allow_html=True)

    # 1 上传花名册文件
    st.subheader("1️⃣ 上传花名册")
    uploaded_file = st.file_uploader("选择花名册Excel文件", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        if st.button("处理花名册", type="primary"):
            with st.spinner("正在处理花名册..."):
                roster_data = process_roster_file(uploaded_file)
                if roster_data:
                    st.session_state.roster_data = roster_data
                    st.session_state.student_id_to_name = roster_data['student_id_to_name']
                    # 重置检查状态，因为数据变了
                    st.session_state.check_performed = False
                    st.success(f"花名册处理完成！共读取 {roster_data['total_students']} 名学生")

    # 2 文件类型配置
    st.subheader("2️⃣ 文件查找配置")
    check_all_types = st.checkbox("查找所有类型文件(无视后缀)🔍", value=False)

    target_exts = []
    if not check_all_types:
        # 默认只查找 .py，用户可以输入多个，用逗号隔开
        ext_input = st.text_input("输入要查找的文件后缀 (英文逗号分隔)", value=".py, .zip, .docx")
        # 处理用户输入：分割、去空格、转小写、确保有点号
        if ext_input:
            raw_exts = ext_input.replace('，', ',').split(',')
            for ext in raw_exts:
                clean_ext = ext.strip().lower()
                if clean_ext:
                    if not clean_ext.startswith('.'):
                        clean_ext = '.' + clean_ext
                    target_exts.append(clean_ext)
        st.caption(f"当前将查找: {', '.join(target_exts)}")
    else:
        st.caption("当前将查找文件夹内包含学号的 **所有** 文件")

    # 3 添加作业文件夹
    st.subheader("3️⃣ 添加作业文件")

    # 使用 Tabs 分开两种添加方式
    tab_local, tab_upload = st.tabs(["📂 本地路径", "📦 上传压缩包"])

    # --- 方式 A: 本地路径 (原逻辑) ---
    with tab_local:
        folder_input = st.text_input("输入文件夹路径（绝对路径）", placeholder="例如: D:\\Teaching\\作业1")
        if st.button("添加路径", use_container_width=True):
            if folder_input and os.path.exists(folder_input):
                abs_path = str(Path(folder_input).absolute())
                if abs_path not in st.session_state.folder_paths:
                    st.session_state.folder_paths.append(abs_path)
                    # 本地路径的显示名就是它自己
                    st.session_state.folder_display_names[abs_path] = os.path.basename(abs_path)
                    st.session_state.check_performed = False
                    st.success(f"已添加: {os.path.basename(abs_path)}")
                    st.rerun()
                else:
                    st.warning("该文件夹已存在")
            else:
                st.error("路径无效")

    # --- 方式 B: 上传压缩包 (新逻辑) ---
    with tab_upload:
        uploaded_zip = st.file_uploader("上传作业ZIP包", type="zip")
        if uploaded_zip and st.button("解压并添加", use_container_width=True):
            try:
                # 1. 创建临时目录
                temp_dir = tempfile.mkdtemp(prefix="homework_check_")

                # 2. 解压文件
                with zipfile.ZipFile(uploaded_zip, 'r') as zf:
                    zf.extractall(temp_dir)

                # 3. 添加到路径列表 (逻辑同上)
                if temp_dir not in st.session_state.folder_paths:
                    st.session_state.folder_paths.append(temp_dir)
                    # 关键：把临时路径映射为上传的文件名，方便显示
                    st.session_state.folder_display_names[temp_dir] = f"📦 {uploaded_zip.name}"
                    st.session_state.check_performed = False
                    st.success(f"已解压并添加: {uploaded_zip.name}")
                    st.rerun()
            except Exception as e:
                st.error(f"解压失败: {e}")

    col_clear = st.columns(1)[0]
    with col_clear:
        if st.button("清空所有来源", use_container_width=True, type="secondary"):
            # 可选：这里可以遍历 folder_paths 删除临时目录，但这步如果不做，操作系统重启也会清理
            st.session_state.folder_paths = []
            st.session_state.folder_display_names = {}  # 清空映射
            st.session_state.folder_results = {}
            st.session_state.check_performed = False
            st.rerun()

    # 显示已添加的列表 (稍微修改显示逻辑)
    if st.session_state.folder_paths:
        st.subheader(f"已添加 ({len(st.session_state.folder_paths)})")
        container = st.container(height=200)
        for i, folder_path in enumerate(st.session_state.folder_paths):
            # 获取显示名称，如果没有映射则显示 basename
            display_name = st.session_state.folder_display_names.get(folder_path, os.path.basename(folder_path))

            container.markdown(f"""
                <div class="folder-item">
                    <strong>{i + 1}. {display_name}</strong><br>
                    <small title="{folder_path}">{folder_path}</small>
                </div>
                """, unsafe_allow_html=True)

# ==========================================
# 4. 主界面逻辑 (可视化与下载)
# ==========================================

st.title("📝 学生作业查收与可视化工具")

if not st.session_state.check_performed:
    st.info("""#### 👈🫡 请在左侧侧边栏上传花名册，进行文件查找配置，并添加作业文件夹，然后点击“开始检查作业”。""")
    # 显示使用指南
    st.markdown("""
    ### 使用指南
    1. **上传花名册**：Excel文件需包含“学号”和“姓名”列。
    2. **文件查找配置**：可以指定要查找的文件类型，或者查找所有类型文件。
    3. **添加文件夹**：复制电脑上的文件夹路径粘贴到输入框中，点击添加。支持添加多个不同位置的文件夹。
    4. **开始检查**：点击按钮，系统将自动比对名单。
    5. **查看结果**：系统将显示提交统计、可视化图表和未交名单.
    6. **下载文件**：可以下载打包文件.zip或者单个文件.xlsx/.txt。

    ### 文件要求：
    - **花名册文件**：Excel格式，需包含9位学号和姓名列。
    - **作业文件**：支持多种格式，但文件名中需包含9位学号。
    - **文件夹路径**：确保有访问权限的本地文件夹路径。
    """)
    # 开始检查按钮
    # 只有当花名册和文件夹都有的时候才显示主按钮
    ready_to_check = st.session_state.roster_data and st.session_state.folder_paths

    # ... (在“开始检查”按钮逻辑中，调用新的 check 函数) ...
    if st.button("开始检查作业✔️", type="primary", use_container_width=True, disabled=not ready_to_check):
        with st.spinner("正在检查作业提交情况..."):
            folder_results = {}
            for folder_path in st.session_state.folder_paths:
                # 优先使用我们记录的名字（如 "📦 作业1.zip"），找不到才用文件夹名
                folder_name = st.session_state.folder_display_names.get(folder_path, os.path.basename(folder_path))
                # !!! 注意这里传入了新的参数 !!!
                result = check_homework_in_folder(
                    folder_path,
                    st.session_state.roster_data['student_ids'],
                    target_extensions=target_exts,
                    check_all_types=check_all_types
                )
                if result:
                    folder_results[folder_name] = result

            st.session_state.folder_results = folder_results
            st.session_state.check_performed = True
            st.success("检查完成！")
            st.rerun()  # 强制刷新主界面显示结果

else:
    # ------------------
    # 4.1 数据准备
    # ------------------
    results = st.session_state.folder_results
    id_map = st.session_state.student_id_to_name

    # 准备图表数据
    chart_data = []
    generated_files_list = []
    total_missing_all = []  # 汇总列表

    for folder_name, res in results.items():
        # 图表数据
        chart_data.append({
            "作业文件夹": folder_name,
            "已提交": res['submitted_count'],
            "未提交": res['missing_count']
        })

        missing_list = sorted(list(res['missing_ids']))

        # 汇总数据收集
        for sid in missing_list:
            total_missing_all.append({
                "文件夹": folder_name,
                "学号": sid,
                "姓名": id_map.get(sid, "未知")
            })

        # 生成下载文件数据
        if missing_list:
            # Excel
            df_out = pd.DataFrame([{"学号": sid, "姓名": id_map.get(sid, "未知")} for sid in missing_list])
            excel_buffer = io.BytesIO()
            df_out.to_excel(excel_buffer, index=False)

            # TXT
            txt_content = f"未交作业名单 - {folder_name}\n" + "=" * 30 + "\n"
            for sid in missing_list:
                txt_content += f"{sid}\t{id_map.get(sid, '未知')}\n"

            generated_files_list.append({
                "filename": f"未交名单_{folder_name}.xlsx",
                "data": excel_buffer.getvalue(),
                "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "folder": folder_name
            })
            generated_files_list.append({
                "filename": f"未交名单_{folder_name}.txt",
                "data": txt_content.encode('utf-8'),
                "mime": "text/plain",
                "folder": folder_name
            })

    # 生成汇总文件
    if total_missing_all:
        df_total = pd.DataFrame(total_missing_all)
        excel_buffer_total = io.BytesIO()
        df_total.to_excel(excel_buffer_total, index=False)
        generated_files_list.insert(0, {
            "filename": "未交作业名单_汇总.xlsx",
            "data": excel_buffer_total.getvalue(),
            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "folder": "汇总数据"
        })

    # ------------------
    # 4.2 可视化展示
    # ------------------
    st.divider()

    # 概览图表
    col_chart, col_stat = st.columns([2, 1])
    with col_chart:
        st.subheader("📊 提交情况概览")
        if chart_data:
            st.bar_chart(pd.DataFrame(chart_data).set_index("作业文件夹")[["已提交", "未提交"]])

    with col_stat:
        st.subheader("📈 统计数据")
        total_submitted = sum(d['已提交'] for d in chart_data)
        total_missing = sum(d['未提交'] for d in chart_data)
        st.metric("总已交作业份数", total_submitted)
        st.metric("总缺交作业人次", total_missing, delta_color="inverse")

    # 详细名单 Tabs
    st.subheader("🫣 详细缺交名单")

    # 动态创建 Tabs
    tab_labels = ["汇总视图"] + list(results.keys())
    tabs = st.tabs(tab_labels)

    # Tab 1: 汇总
    with tabs[0]:
        if total_missing_all:
            st.dataframe(pd.DataFrame(total_missing_all), use_container_width=True)
        else:
            st.success("🎉 所有文件夹作业均已收齐！")

        # ... (在主界面的 Tabs 循环中) ...

        # Tab 2+: 各个文件夹
        for i, (folder_name, res) in enumerate(results.items()):
            with tabs[i + 1]:
                c1, c2 = st.columns([1, 2])

                # --- c1: 统计数据 ---
                with c1:
                    # 1. 显示缺交大数字
                    st.metric(f"{folder_name} - ❌ ", f"😡{res['missing_count']} 人-缺交")

                    # 2. 显示提交文件类型详情 (新增功能)
                    if res['file_type_stats']:
                        all_count = 0
                        for ext, count in res['file_type_stats'].items():
                            all_count += count
                        # 1. 显示提交大数字
                        st.metric(f"{folder_name} - ✅", f"🥰{all_count} 人-已交")
                        # 将字典转换为 DataFrame 以便美观展示
                        stats_data = [
                            {"文件类型": ext, "数量": count}
                            for ext, count in res['file_type_stats'].items()
                        ]
                        df_stats = pd.DataFrame(stats_data).sort_values("数量", ascending=False)

                        # 使用 st.dataframe 展示，隐藏索引，调整高度
                        st.dataframe(
                            df_stats,
                            hide_index=True,
                            use_container_width=True,
                            column_config={
                                "文件类型": st.column_config.TextColumn("类型", width="small"),
                                "数量": st.column_config.ProgressColumn(
                                    "提交数量",
                                    format="%d",
                                    min_value=0,
                                    max_value=max(res['file_type_stats'].values())
                                )
                            }
                        )
                    else:
                        st.caption("没有检测到符合条件的文件。")

                # --- c2: 缺交名单 (保持不变) ---
                with c2:
                    st.markdown("##### 🫵 缺交学生名单")
                    if res['missing_ids']:
                        # 你的原始逻辑...
                        missing_data = [{"学号": sid, "姓名": id_map.get(sid, "未知")} for sid in
                                        sorted(res['missing_ids'])]
                        st.dataframe(pd.DataFrame(missing_data), use_container_width=True, height=400)
                    else:
                        st.success("🎉 全员已交！")

    # ------------------
    # 4.3 下载中心
    # ------------------
    st.markdown("---")
    st.header("👾 下载中心")

    if not generated_files_list:
        st.info("没有生成任何名单文件。")
    else:
        # 方式一：打包下载
        st.subheader("📦- 打包下载所有文件")
        # 生成 ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for file_item in generated_files_list:
                zip_file.writestr(file_item['filename'], file_item['data'])
        st.download_button(
            label="🚀- 下载全部文件 (.zip)",
            data=zip_buffer.getvalue(),
            file_name="作业检查结果_总和.zip",
            mime="application/zip",
            use_container_width=True,
            type="primary"
        )
        # 方式二：单独下载
        st.subheader("📜- 单独下载指定文件")
        cols = st.columns(2)

        # 分离汇总文件和普通文件
        summary_files = [f for f in generated_files_list if "汇总" in f['filename']]
        other_files = [f for f in generated_files_list if "汇总" not in f['filename']]

        # 显示汇总文件
        for i, f in enumerate(summary_files):
            cols[0].download_button(
                label=f"⬇️ {f['filename']}",
                data=f['data'],
                file_name=f['filename'],
                mime=f['mime'],
                key=f"dl_sum_{i}"
            )

        # 显示普通文件
        for i, f in enumerate(other_files):
            col_idx = (i + len(summary_files)) % 2
            cols[col_idx].download_button(
                label=f"⬇️ {f['filename']} ({f['folder']})",
                data=f['data'],
                file_name=f['filename'],
                mime=f['mime'],
                key=f"dl_norm_{i}"
            )