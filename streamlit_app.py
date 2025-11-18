import streamlit as st
import pandas as pd
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="九宫格潜力展示系统",
    page_icon="📊",
    layout="wide"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-title {
        font-size: 24px;
        font-weight: bold;
        color: #1f4e79;
        margin-bottom: 20px;
    }
    
    .stats-table {
        border: 1px solid #4472c4;
        border-radius: 5px;
        padding: 10px;
        background-color: #e7f3ff;
    }
    
    .grid-cell {
        border: 2px solid #4472c4;
        border-radius: 8px;
        padding: 15px;
        margin: 8px 4px;
        min-height: 220px;
        max-height: 280px;
        background-color: #f8f9fa;
        display: flex;
        flex-direction: column;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        position: relative;
    }
    
    .cell-title {
        font-weight: bold;
        color: #1f4e79;
        font-size: 14px;
        margin-bottom: 8px;
        border-bottom: 1px solid #d0d0d0;
        padding-bottom: 5px;
        background-color: #f8f9fa;
        position: sticky;
        top: 0;
        z-index: 10;
        flex-shrink: 0;
    }
    
    .cell-content {
        font-size: 12px;
        line-height: 1.4;
        flex-grow: 1;
        overflow-y: auto;
        color: #333;
        max-height: 200px;
        padding-right: 5px;
    }
    
    .cell-content::-webkit-scrollbar {
        width: 4px;
    }
    
    .cell-content::-webkit-scrollbar-track {
        background: #f1f1f1;
        border-radius: 2px;
    }
    
    .cell-content::-webkit-scrollbar-thumb {
        background: #c1c1c1;
        border-radius: 2px;
    }
    
    .cell-content::-webkit-scrollbar-thumb:hover {
        background: #a1a1a1;
    }
    
    .grid-row {
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

def load_and_validate_data(uploaded_file, sheet_name="Sheet1"):
    """
    加载并验证Excel数据
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # 检查必需的列是否存在
        required_columns = ['九宫格', '员工姓名', '部门']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"缺少必需的列: {missing_columns}")
            st.info("请确保Excel文件包含以下列：九宫格、员工姓名、部门")
            return None
            
        # 重命名列以保持代码兼容性
        df = df.rename(columns={'九宫格': '档位', '员工姓名': '姓名'})
            
        # 数据清洗：移除空值
        df_clean = df.dropna(subset=['档位', '姓名', '部门'])
        
        if len(df_clean) < len(df):
            removed_count = len(df) - len(df_clean)
            st.warning(f"已移除 {removed_count} 行包含空值的数据")
        
        return df_clean
        
    except Exception as e:
        st.error(f"读取Excel文件时出错: {str(e)}")
        return None

def get_potential_level(rating):
    """
    根据档位获取潜力级别
    """
    if rating in [7, 8, 9]:
        return "高潜力"
    elif rating in [4, 5, 6]:
        return "中潜力"
    elif rating in [1, 2, 3]:
        return "低潜力"
    else:
        return "未知"

def generate_summary(df, rating):
    """
    生成指定档位的员工汇总信息
    """
    if df is None or df.empty:
        return "无数据", 0
    
    # 筛选指定档位的数据
    rating_data = df[df['档位'] == rating]
    
    if rating_data.empty:
        return "无符合员工", 0
    
    # 按部门分组
    dept_groups = rating_data.groupby('部门')['姓名'].apply(list).to_dict()
    
    # 格式化输出
    summary_lines = []
    total_count = 0
    
    for dept, names in dept_groups.items():
        names_str = "、".join(names)
        summary_lines.append(f"{dept}：{names_str}")
        total_count += len(names)
    
    summary = "\n".join(summary_lines)
    return summary, total_count

def create_stats_table(df):
    """
    创建统计汇总表格
    """
    if df is None or df.empty:
        return pd.DataFrame()
    
    # 计算各档位统计
    stats_data = []
    
    # 高潜力统计 (7,8,9)
    high_potential = df[df['档位'].isin([7, 8, 9])]
    high_count = len(high_potential)
    
    # 中潜力统计 (4,5,6) 
    mid_potential = df[df['档位'].isin([4, 5, 6])]
    mid_count = len(mid_potential)
    
    # 低潜力统计 (1,2,3)
    low_potential = df[df['档位'].isin([1, 2, 3])]
    low_count = len(low_potential)
    
    # 构建统计表格
    stats_df = pd.DataFrame({
        '指标': ['高潜力', '中潜力', '低潜力'],
        '人数': [high_count, mid_count, low_count]
    })
    
    return stats_df

def create_sample_data():
    """
    创建示例数据用于测试
    """
    sample_data = {
        '九宫格': [7, 8, 9, 4, 5, 6, 1, 2, 3, 7, 8, 5, 2, 6, 9],
        '员工姓名': ['张三', '李四', '王五', '赵六', '钱七', '孙八', '周九', '吴十', 
                '郑一', '陈二', '褚三', '卫四', '蒋五', '沈六', '韩七'],
        '部门': ['技术部', '市场部', '人事部', '财务部', '技术部', '市场部', 
                '人事部', '财务部', '技术部', '市场部', '人事部', '财务部',
                '技术部', '市场部', '人事部']
    }
    df = pd.DataFrame(sample_data)
    # 重命名列以保持代码兼容性
    df = df.rename(columns={'九宫格': '档位', '员工姓名': '姓名'})
    return df

def main():
    """
    主应用函数
    """
    # 页面标题
    st.markdown('<div class="main-title">📊 基于九宫格潜力展示系统</div>', 
                unsafe_allow_html=True)
    
    # 侧边栏 - 文件上传
    st.sidebar.header("📁 数据上传")
    
    # 提供示例数据选项
    use_sample = st.sidebar.checkbox("使用示例数据进行测试")
    
    df = None
    
    if use_sample:
        df = create_sample_data()
        st.sidebar.success("已加载示例数据")
    else:
        uploaded_file = st.sidebar.file_uploader(
            "上传Excel文件", 
            type=['xlsx', 'xls'],
            help="请上传包含'九宫格'、'员工姓名'、'部门'列的Excel文件"
        )
        
        if uploaded_file is not None:
            sheet_name = st.sidebar.text_input("工作表名称", value="Sheet1")
            df = load_and_validate_data(uploaded_file, sheet_name)
    
    if df is not None and not df.empty:
        # 数据概览
        total_count = len(df)
        
        # 创建主布局
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown(f'<div class="main-title">2 基于九宫格一览：总计{total_count}人</div>', 
                       unsafe_allow_html=True)
        
        with col2:
            # 统计表格
            stats_df = create_stats_table(df)
            if not stats_df.empty:
                st.markdown('<div class="stats-table">', unsafe_allow_html=True)
                st.dataframe(stats_df, width='stretch', hide_index=True)
                st.markdown('</div>', unsafe_allow_html=True)
        
        # 九宫格布局
        st.markdown("### 📊 九宫格潜力分布")
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 定义九宫格排列 (7,8,9 / 4,5,6 / 1,2,3)
        grid_layout = [
            [7, 8, 9],
            [4, 5, 6], 
            [1, 2, 3]
        ]
        
        # 创建3行布局
        for row_idx, row in enumerate(grid_layout):
            cols = st.columns(3)
            
            for i, rating in enumerate(row):
                with cols[i]:
                    # 获取该档位的汇总信息
                    summary, count = generate_summary(df, rating)
                    potential_level = get_potential_level(rating)
                    
                    # 格子标题 - 显示实际人数
                    title = f"{rating} {potential_level}，低绩效-{count}人"
                    
                    # 使用HTML容器显示内容
                    if summary and summary != "无符合员工":
                        # 确保HTML转义和格式正确
                        content_lines = []
                        for line in summary.split('\n'):
                            if line.strip():
                                # HTML转义特殊字符
                                escaped_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                                content_lines.append(escaped_line)
                        content = "<br>".join(content_lines)
                    else:
                        content = "无符合员工"
                    
                    # 使用完整的HTML容器
                    cell_html = f"""
                    <div class="grid-cell">
                        <div class="cell-title">{title}</div>
                        <div class="cell-content">{content}</div>
                    </div>
                    """
                    st.markdown(cell_html, unsafe_allow_html=True)
            
            # 在每行后添加间距
            st.markdown("<br>", unsafe_allow_html=True)
        
        # 绩效箭头指示
        st.markdown("**绩效 →**")
        
        # 数据详情
        with st.expander("📋 查看原始数据"):
            st.dataframe(df, width='stretch')
            
        # 下载处理后的数据
        if st.button("📥 下载汇总报告"):
            # 创建汇总报告
            report_data = []
            for rating in range(1, 10):
                summary, count = generate_summary(df, rating)
                potential_level = get_potential_level(rating)
                report_data.append({
                    '档位': rating,
                    '潜力级别': potential_level,
                    '人数': count,
                    '员工详情': summary.replace('\n', '; ') if summary != "无符合员工" else "无"
                })
            
            report_df = pd.DataFrame(report_data)
            
            # 转换为Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                report_df.to_excel(writer, sheet_name='九宫格汇总', index=False)
                df.to_excel(writer, sheet_name='原始数据', index=False)
            
            st.download_button(
                label="下载Excel报告",
                data=output.getvalue(),
                file_name="九宫格潜力汇总报告.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    else:
        # 显示使用说明
        st.info("👆 请在左侧上传Excel文件或选择使用示例数据")
        
        st.markdown("""
        ### 📋 使用说明
        
        1. **数据格式要求**：
           - Excel文件需包含三列：`九宫格`、`员工姓名`、`部门`
           - 九宫格列应为1-9的数字
           - 员工姓名和部门列为文本格式
        
        2. **九宫格说明**：
           - 7,8,9 → 高潜力
           - 4,5,6 → 中潜力  
           - 1,2,3 → 低潜力
        
        3. **功能特点**：
           - 自动按部门汇总员工信息
           - 实时统计各档位人数
           - 支持数据下载和报告生成
        """)

if __name__ == "__main__":
    main()