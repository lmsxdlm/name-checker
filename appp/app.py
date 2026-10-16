import streamlit as st
import pandas as pd
import os
import io

# ================= 页面配置 =================
st.set_page_config(page_title="名单对比工具", page_icon="📋", layout="wide")
@st.cache_data
def load_master_data(file_name="master_list.xlsx"):
    # 获取当前脚本(app.py)所在的绝对路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 拼接出 xlsx 文件的绝对路径
    file_path = os.path.join(current_dir, file_name)
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        st.error(f"❌ 找不到文件！在目录 {current_dir} 下没有找到 {file_name}。")
        st.write("🔍 当前目录下实际存在的文件有：", os.listdir(current_dir)) # 这行代码能帮你看到服务器上到底有什么文件
        return None
    
    try:
        df = pd.read_excel(file_path, dtype={'学号': str})
        # 去除列名首尾空格
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"读取总名单失败: {e}")
        return None

# ================= 2. 解析用户上传的文件 =================
def extract_names_from_file(uploaded_file):
    """从用户上传的文件中提取姓名列表，自动寻找表头"""
    names = []
    try:
        # 1. 首先以无表头的形式读取前20行，寻找包含“姓名”的行
        df_raw = pd.read_excel(uploaded_file, header=None, nrows=20)
        
        header_row_idx = -1
        for i, row in df_raw.iterrows():
            row_values = [str(val).strip() for val in row.values if pd.notna(val)]
            # 寻找包含“姓名”或“名字”的单元格
            if any("姓名" in val for val in row_values) or any("名字" in val for val in row_values):
                header_row_idx = i
                break
        
        # 2. 如果找到了表头所在的行，重新读取
        if header_row_idx != -1:
            df_data = pd.read_excel(uploaded_file, header=header_row_idx)
            df_data.columns = df_data.columns.astype(str).str.strip()
            
            # 寻找实际包含“姓名”的列名
            name_col = None
            for col in df_data.columns:
                if "姓名" in col or "名字" in col:
                    name_col = col
                    break
            
            if name_col:
                # 提取姓名，去除空值和空格
                extracted_names = df_data[name_col].dropna().astype(str).str.strip().tolist()
                names.extend(extracted_names)
            else:
                st.warning(f"文件 {uploaded_file.name} 中找到疑似表头，但未找到包含'姓名'的列。")
        else:
            st.warning(f"文件 {uploaded_file.name} 中未找到包含'姓名'的列（已尝试寻找前20行）。")
            
    except Exception as e:
        st.error(f"解析文件 {uploaded_file.name} 时出错: {e}")
        
    return names

# ================= 3. 主界面 =================
st.title("📋 名单对比与差异分析工具")
st.markdown("上传需要对比的名单文件（支持多个），系统会自动识别'姓名'列，并与内置的《总名单.xlsx》进行比对。")

# 加载总名单
master_df = load_master_data("总名单.xlsx")

if master_df is None:
    st.error("❌ 未找到 '总名单.xlsx' 文件！请确保该文件与程序在同一目录下。")
    st.stop()

# 显示总名单概览
with st.expander("🔍 查看内置总名单信息", expanded=False):
    st.write(f"总名单共包含 **{len(master_df)}** 条数据")
    st.dataframe(master_df.head(10), use_container_width=True)

# 用户上传区域
uploaded_files = st.file_uploader(
    "📤 请上传需要对比的 Excel 或 CSV 文件（可多选）", 
    type=["xlsx", "xls", "csv"], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 开始对比", type="primary"):
        with st.spinner("正在解析并对比数据..."):
            # 收集所有上传文件中的姓名
            all_user_names = []
            for file in uploaded_files:
                names = extract_names_from_file(file)
                all_user_names.extend(names)
            
            # 去重（防止多人重复提交）
            user_names_set = set(all_user_names)
            
            if not user_names_set:
                st.warning("未能从上传的文件中提取到任何有效姓名，请检查文件格式。")
                st.stop()
            
            # 提取总名单中的姓名和班级
            master_df.columns = master_df.columns.astype(str).str.strip()
            if "姓名" not in master_df.columns:
                st.error("内置的总名单中未找到'姓名'列，请检查 Excel 格式。")
                st.stop()
                
            master_names_set = set(master_df["姓名"].dropna().astype(str).str.strip())
            
            # ================= 对比逻辑 =================
            # 1. 总名单有，但用户上传的文件里没有（缺失人员）
            missing_names = master_names_set - user_names_set
            df_missing = master_df[master_df["姓名"].isin(missing_names)].copy()
            
            # 2. 用户上传的文件里有，但总名单里没有（多余/错误人员）
            extra_names = user_names_set - master_names_set
            df_extra = pd.DataFrame(list(extra_names), columns=["不在总名单中的姓名"])
            
            # ================= 输出结果 =================
            st.success(f"✅ 对比完成！共识别上传文件姓名数: {len(user_names_set)} 人，总名单人数: {len(master_names_set)} 人。")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader(f"❌ 缺失人员 (总名单有，上传文件中未找到)")
                st.write(f"共 {len(df_missing)} 人")
                if not df_missing.empty:
                    st.dataframe(df_missing, use_container_width=True)
                    # 提供下载按钮
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_missing.to_excel(writer, index=False, sheet_name='缺失人员')
                    st.download_button(
                        label="📥 下载缺失人员名单",
                        data=output.getvalue(),
                        file_name="缺失人员名单.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("🎉 没有缺失人员，所有总名单成员都在上传文件中！")
                    
            with col2:
                st.subheader(f"⚠️ 异常人员 (上传文件有，总名单中没有)")
                st.write(f"共 {len(df_extra)} 人")
                if not df_extra.empty:
                    st.dataframe(df_extra, use_container_width=True)
                    # 提供下载按钮
                    output2 = io.BytesIO()
                    with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                        df_extra.to_excel(writer, index=False, sheet_name='异常人员')
                    st.download_button(
                        label="📥 下载异常人员名单",
                        data=output2.getvalue(),
                        file_name="异常人员名单.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("🎉 没有异常人员！")
                    