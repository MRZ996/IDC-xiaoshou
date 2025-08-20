import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import openpyxl
import time

# 设置页面 - 优化布局和主题
st.set_page_config(
    page_title="IDC销售健康度分析",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded",
    menu_items={
        'About': "# IDC销售健康度分析工具\n专业评估IDC资源销售健康状况"
    }
)

# 自定义样式 - 现代化UI设计升级
st.markdown("""
<style>
    :root {
        --primary: #165DFF;
        --primary-light: #4080FF;
        --primary-dark: #0E42D2;
        --secondary: #36CFC9;
        --success: #52C41A;
        --warning: #FAAD14;
        --danger: #FF4D4F;
        --light: #F0F2F5;
        --dark: #1D2129;
        --text-primary: #1D2129;
        --text-secondary: #4E5969;
        --text-tertiary: #86909C;
    }
    
    .stApp {
        background-color: #F7F8FA;
    }
    
    .header-text {
        text-align: center;
        margin: 20px 0 30px 0;
        font-size: 2.5rem;
        font-weight: 700;
        background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        position: relative;
    }
    
    .header-text::after {
        content: '';
        display: block;
        width: 120px;
        height: 4px;
        background: var(--primary);
        border-radius: 2px;
        margin: 10px auto 0;
        opacity: 0.7;
    }
    
    .metric-card {
        background: white;
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        transition: all 0.3s ease;
        border: 1px solid rgba(22, 93, 255, 0.08);
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 24px rgba(22, 93, 255, 0.12);
        border-color: rgba(22, 93, 255, 0.2);
    }
    
    .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: var(--text-primary);
        margin: 8px 0;
        line-height: 1.2;
    }
    
    .metric-label {
        font-size: 15px;
        color: var(--text-secondary);
        font-weight: 500;
    }
    
    .section {
        background: white;
        border-radius: 16px;
        padding: 25px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
        margin-bottom: 30px;
        border: 1px solid rgba(22, 93, 255, 0.08);
        transition: border-color 0.3s ease;
    }
    
    .section:hover {
        border-color: rgba(22, 93, 255, 0.15);
    }
    
    .stButton>button {
        background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(22, 93, 255, 0.2);
    }
    
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(22, 93, 255, 0.3);
        background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
    }
    
    .stButton>button:active {
        transform: translateY(0);
        box-shadow: 0 2px 8px rgba(22, 93, 255, 0.2);
    }
    
    .risk-card {
        background: #FFF8E6;
        border-left: 4px solid var(--warning);
        padding: 18px;
        border-radius: 8px;
        margin-bottom: 15px;
        transition: all 0.3s ease;
    }
    
    .risk-card:hover {
        transform: translateX(5px);
        box-shadow: 0 4px 12px rgba(250, 173, 20, 0.1);
    }
    
    .excellent { color: var(--success); }
    .good { color: #1890FF; }
    .fair { color: var(--warning); }
    .danger { color: var(--danger); }
    
    .progress-container {
        height: 8px;
        background: #E5E6EB;
        border-radius: 4px;
        overflow: hidden;
        margin-top: 12px;
    }
    
    .progress-bar {
        height: 100%;
        border-radius: 4px;
        transition: width 1s ease-in-out;
    }
    
    .indicator-badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 600;
        margin-right: 8px;
        margin-bottom: 8px;
    }
    
    .trend-up {
        color: var(--success);
        font-weight: 600;
    }
    
    .trend-down {
        color: var(--danger);
        font-weight: 600;
    }
    
    .stTab > div > div > div > button {
        padding: 12px 24px !important;
        font-weight: 600 !important;
        color: var(--text-secondary) !important;
    }
    
    .stTab > div > div > div > button[aria-selected="true"] {
        color: var(--primary) !important;
        border-bottom: 2px solid var(--primary) !important;
    }
    
    .stDownloadButton > button {
        background: linear-gradient(135deg, var(--primary-dark) 0%, var(--primary) 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(22, 93, 255, 0.2);
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(22, 93, 255, 0.3);
        background: linear-gradient(135deg, var(--primary) 0%, var(--primary-light) 100%);
    }
    
    /* 增强表格样式 */
    .dataframe {
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.05) !important;
    }
    
    /* 优化分隔线 */
    [data-testid="stDivider"] {
        margin: 20px 0 !important;
    }
    
    /* 美化滑块 */
    [data-baseweb="slider"] {
        padding: 15px 0 !important;
    }
    
    /* 美化文件上传器 */
    [data-testid="stFileUploader"] {
        padding: 20px;
        border: 2px dashed rgba(22, 93, 255, 0.2);
        border-radius: 12px;
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: rgba(22, 93, 255, 0.4);
        background-color: rgba(22, 93, 255, 0.02);
    }
    
    /* 优化侧边栏 */
    [data-testid="stSidebar"] {
        background-color: white;
        box-shadow: 2px 0 10px rgba(0, 0, 0, 0.05);
    }
    
    /* 美化标签页 */
    [data-testid="stTabs"] {
        margin-top: 15px;
    }
    
    /* 添加动画效果 */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .fade-in {
        animation: fadeIn 0.6s ease-out;
    }
    
    /* 优化卡片悬停效果 */
    .hover-card {
        transition: all 0.3s ease;
    }
    
    .hover-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 24px rgba(22, 93, 255, 0.12);
    }
    
    /* 美化选择框 */
    .stSelectbox > div > div {
        border-radius: 8px;
        border: 1px solid rgba(22, 93, 255, 0.2);
    }
    
    /* 美化多选框 */
    .stMultiSelect > div > div {
        border-radius: 8px;
        border: 1px solid rgba(22, 93, 255, 0.2);
    }
</style>
""", unsafe_allow_html=True)

# 初始化会话状态
if 'data' not in st.session_state:
    st.session_state.data = None
if 'results' not in st.session_state:
    st.session_state.results = None
if 'analysis_complete' not in st.session_state:
    st.session_state.analysis_complete = False
if 'last_updated' not in st.session_state:
    st.session_state.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

# 全局权重定义
weights = {
    '资源利用': 0.25,
    '客户健康': 0.25,
    '财务健康': 0.20,
    '风险控制': 0.15,
    '增长潜力': 0.15
}

# 页面标题与说明
st.markdown("<h1 class='header-text'>IDC资源销售健康度测评辅助工具</h1>", unsafe_allow_html=True)
st.caption("全面评估数据中心资源销售健康状况，精准识别风险与机遇")

# 使用查询参数示例
params = st.query_params
if 'demo' in params and params['demo'] == 'true' and st.session_state.data is None:
    with st.spinner("检测到演示模式参数，正在加载示例数据..."):
        time.sleep(1)
        # 生成示例数据
        dates = pd.date_range(start='2023-01-01', periods=12, freq='M')
        data = {
            '月份': dates,
            '服务器利用率': [68, 72, 75, 78, 82, 80, 77, 79, 83, 85, 88, 90],
            '带宽利用率': [65, 68, 72, 75, 78, 76, 72, 74, 77, 80, 84, 86],
            '机柜利用率': [75, 78, 82, 85, 88, 86, 83, 84, 87, 89, 91, 93],
            '新客户数量': [8, 10, 12, 14, 16, 13, 11, 15, 18, 20, 23, 25],
            '客户流失率': [5.2, 4.8, 4.5, 4.2, 3.8, 4.1, 4.3, 3.9, 3.6, 3.3, 3.0, 2.7],
            '平均合同期限': [16, 18, 20, 22, 24, 23, 21, 23, 25, 27, 29, 31],
            '月收入(万元)': [120, 140, 160, 180, 200, 190, 175, 195, 210, 225, 240, 255],
            '利润率': [25, 27, 29, 31, 33, 32, 30, 32, 34, 36, 38, 40],
            '应收账款周转天数': [60, 58, 55, 52, 48, 50, 53, 49, 46, 43, 40, 37],
            '高风险客户占比': [15, 14, 13, 12, 11, 12, 13, 11, 10, 9, 8, 7],
            '服务中断次数': [3, 3, 2, 2, 1, 2, 2, 1, 1, 0, 0, 0],
            '市场增长率': [1.2, 1.5, 1.8, 2.0, 2.3, 2.1, 1.9, 2.2, 2.5, 2.7, 3.0, 3.2],
            '销售漏斗数量': [35, 40, 45, 50, 55, 52, 48, 58, 65, 72, 80, 88]
        }
        st.session_state.data = pd.DataFrame(data)
        st.session_state.analysis_complete = False
        st.success("示例数据已加载完成！")

# 侧边栏配置
with st.sidebar:
    # 替换图片为文字标题和图标
    st.markdown("""
    <div style='text-align:center; padding:10px 0 20px 0;'>
        <h1 style='font-size:28px; margin:0; color:#165DFF;'>📊 IDC分析平台</h1>
        <p style='font-size:14px; color:#4E5969; margin:5px 0 0 0;'>数据中心销售健康度分析</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 导航菜单")
    page_nav = st.radio("", ["数据导入", "健康度分析", "风险洞察", "趋势预测", "报告导出"], 
                       format_func=lambda x: f"📌 {x}")
    
    with st.expander("📊 指标说明", expanded=True):
        st.markdown("""
        **健康度分析指标说明：**
        
        - **资源利用指标**  
          服务器、带宽、机柜利用率综合评分
          
        - **客户健康指标**  
          新客户增长、客户流失率、合同稳定性分析
          
        - **财务健康指标**  
          收入趋势、利润率、资金周转效率评估
          
        - **风险控制指标**  
          高风险客户占比、服务稳定性监控
          
        - **增长潜力指标**  
          市场拓展速度、销售机会储备分析
        """)
    
    st.markdown("---")
    st.markdown("**数据质量检查**")
    if st.session_state.data is not None:
        missing_data = st.session_state.data.isnull().sum().sum()
        if missing_data > 0:
            st.error(f"发现{missing_data}处缺失值")
            st.warning("建议清理缺失数据以获得更准确的分析结果")
        else:
            st.success("✅ 数据完整无缺失")
    
    st.markdown("---")
    st.info("**使用提示**")
    st.markdown("""
    1. 先导入数据或使用示例数据
    2. 查看健康度总分和等级评估
    3. 分析各维度表现与趋势
    4. 识别关键风险点与改进方向
    5. 导出分析报告用于决策支持
    """)
    
    st.markdown("---")
    st.caption(f"数据更新时间: {st.session_state.last_updated}")
    st.caption("版本 1.0.0")

# 数据输入区域
if page_nav == "数据导入":
    with st.container():
        st.subheader("数据导入", divider="blue")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.markdown("#### 示例数据")
            st.info("快速生成模拟数据进行分析演示，包含12个月的完整指标")
            
            if st.button("生成示例数据", use_container_width=True, type="primary"):
                with st.spinner("正在生成示例数据..."):
                    # 生成更全面的示例数据
                    dates = pd.date_range(start='2023-01-01', periods=12, freq='M')
                    data = {
                        '月份': dates,
                        '服务器利用率': [68, 72, 75, 78, 82, 80, 77, 79, 83, 85, 88, 90],
                        '带宽利用率': [65, 68, 72, 75, 78, 76, 72, 74, 77, 80, 84, 86],
                        '机柜利用率': [75, 78, 82, 85, 88, 86, 83, 84, 87, 89, 91, 93],
                        '新客户数量': [8, 10, 12, 14, 16, 13, 11, 15, 18, 20, 23, 25],
                        '客户流失率': [5.2, 4.8, 4.5, 4.2, 3.8, 4.1, 4.3, 3.9, 3.6, 3.3, 3.0, 2.7],
                        '平均合同期限': [16, 18, 20, 22, 24, 23, 21, 23, 25, 27, 29, 31],
                        '月收入(万元)': [120, 140, 160, 180, 200, 190, 175, 195, 210, 225, 240, 255],
                        '利润率': [25, 27, 29, 31, 33, 32, 30, 32, 34, 36, 38, 40],
                        '应收账款周转天数': [60, 58, 55, 52, 48, 50, 53, 49, 46, 43, 40, 37],
                        '高风险客户占比': [15, 14, 13, 12, 11, 12, 13, 11, 10, 9, 8, 7],
                        '服务中断次数': [3, 3, 2, 2, 1, 2, 2, 1, 1, 0, 0, 0],
                        '市场增长率': [1.2, 1.5, 1.8, 2.0, 2.3, 2.1, 1.9, 2.2, 2.5, 2.7, 3.0, 3.2],
                        '销售漏斗数量': [35, 40, 45, 50, 55, 52, 48, 58, 65, 72, 80, 88]
                    }
                    st.session_state.data = pd.DataFrame(data)
                    st.session_state.analysis_complete = False
                    st.session_state.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")
                    st.success("示例数据已生成！")
        
        with col2:
            st.markdown("#### 上传数据")
            st.info("支持CSV或Excel格式的数据文件，需包含指定指标列")
            
            uploaded_file = st.file_uploader(
                "上传IDC销售数据文件", 
                type=['csv', 'xlsx'],
                help="请确保数据包含所有必要的指标列，具体要求见指标说明"
            )
            
            if uploaded_file:
                try:
                    # 根据文件类型读取
                    with st.spinner("正在处理文件..."):
                        if uploaded_file.name.endswith('.csv'):
                            df = pd.read_csv(uploaded_file)
                        elif uploaded_file.name.endswith('.xlsx'):
                            df = pd.read_excel(uploaded_file, engine='openpyxl')
                        
                        # 基本列检查
                        required_columns = ['月份', '服务器利用率', '带宽利用率', '机柜利用率', '新客户数量', 
                                           '客户流失率', '平均合同期限', '月收入(万元)', '利润率', 
                                           '应收账款周转天数', '高风险客户占比', '服务中断次数', 
                                           '市场增长率', '销售漏斗数量']
                        
                        missing = [col for col in required_columns if col not in df.columns]
                        if missing:
                            st.error(f"缺少必要列: {', '.join(missing)}")
                            st.info("请检查数据格式是否符合要求，参考示例数据结构")
                        else:
                            st.session_state.data = df
                            st.session_state.analysis_complete = False
                            st.session_state.last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")
                            st.success("数据上传成功！")
                            
                except Exception as e:
                    st.error(f"文件处理错误: {str(e)}")
                    st.info("提示: 请确保Excel文件格式正确且未加密")
    
    if st.session_state.data is not None:
        with st.expander("🔍 数据预览", expanded=True):
            st.dataframe(st.session_state.data.head(10), use_container_width=True)
            
            # 数据摘要统计
            st.markdown("**数据摘要统计**")
            st.dataframe(st.session_state.data.describe(), use_container_width=True)

# 健康度计算函数
def calculate_health_scores(df):
    # 使用全局权重
    global weights
    
    # 计算各项指标得分
    df['资源利用得分'] = (df['服务器利用率']*0.4 + df['带宽利用率']*0.4 + df['机柜利用率']*0.2) * weights['资源利用']
    df['客户健康得分'] = ((100 - df['客户流失率'])*0.4 + (df['平均合同期限']/36*100)*0.4 + (df['新客户数量']/25*100)*0.2) * weights['客户健康']
    df['财务健康得分'] = (df['利润率']*0.5 + (100 - df['应收账款周转天数'])/100*100*0.3 + df['月收入(万元)']/400*100*0.2) * weights['财务健康']
    df['风险控制得分'] = ((100 - df['高风险客户占比'])*0.5 + (10 - df['服务中断次数'])/10*100*0.5) * weights['风险控制']
    df['增长潜力得分'] = (df['市场增长率']/5*100*0.5 + df['销售漏斗数量']/100*100*0.5) * weights['增长潜力']
    
    # 计算总分
    df['健康度总分'] = df[['资源利用得分', '客户健康得分', '财务健康得分', '风险控制得分', '增长潜力得分']].sum(axis=1)
    
    # 添加健康度等级
    conditions = [
        (df['健康度总分'] >= 85),
        (df['健康度总分'] >= 70) & (df['健康度总分'] < 85),
        (df['健康度总分'] >= 50) & (df['健康度总分'] < 70),
        (df['健康度总分'] < 50)
    ]
    choices = ['优秀', '良好', '一般', '危险']
    df['健康度等级'] = np.select(conditions, choices, default='未知')
    
    return df

# 健康度分析页面
if page_nav == "健康度分析" and st.session_state.data is not None:
    # 计算健康度指标
    if not st.session_state.analysis_complete:
        with st.spinner("正在分析数据，请稍候..."):
            st.session_state.results = calculate_health_scores(st.session_state.data.copy())
            st.session_state.analysis_complete = True
            time.sleep(1)
    
    df = st.session_state.results
    latest_data = df.iloc[-1]
    
    st.subheader("健康度概览", divider="blue")
    
    # 创建仪表盘
    col1, col2, col3 = st.columns([1, 1.2, 1])
    
    with col1:
        # 健康度总分卡片
        level_class = ""
        if latest_data['健康度等级'] == '优秀':
            level_class = "excellent"
            level_bg = "rgba(82, 196, 26, 0.1)"
            level_border = "var(--success)"
        elif latest_data['健康度等级'] == '良好':
            level_class = "good"
            level_bg = "rgba(24, 144, 255, 0.1)"
            level_border = "#1890FF"
        elif latest_data['健康度等级'] == '一般':
            level_class = "fair"
            level_bg = "rgba(250, 173, 20, 0.1)"
            level_border = "var(--warning)"
        else:
            level_class = "danger"
            level_bg = "rgba(255, 77, 79, 0.1)"
            level_border = "var(--danger)"
            
        st.markdown(f"""
        <div class='metric-card fade-in'>
            <div class='metric-label'>当前健康度总分</div>
            <div class='metric-value'>{latest_data['健康度总分']:.1f}/100</div>
            <div style='text-align:center;font-size:28px;font-weight:700;padding:10px;border-radius:8px;background:{level_bg};border:1px solid {level_border};' class='{level_class}'>
                {latest_data['健康度等级']}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # 健康度仪表盘
        fig_gauge = go.Figure(go.Indicator(
            mode = "gauge+number+delta",
            value = latest_data['健康度总分'],
            domain = {'x': [0, 1], 'y': [0, 1]},
            title = {'text': "健康度总分", 'font': {'size': 20}},
            delta = {'reference': df.iloc[-2]['健康度总分'] if len(df) > 1 else latest_data['健康度总分'], 
                    'increasing': {'color': "var(--success)"},
                    'decreasing': {'color': "var(--danger)"}},
            gauge = {
                'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "var(--text-tertiary)"},
                'steps': [
                    {'range': [0, 50], 'color': "rgba(255, 77, 79, 0.2)"},
                    {'range': [50, 70], 'color': "rgba(250, 173, 20, 0.2)"},
                    {'range': [70, 85], 'color': "rgba(24, 144, 255, 0.2)"},
                    {'range': [85, 100], 'color': "rgba(82, 196, 26, 0.2)"}
                ],
                'threshold': {
                    'line': {'color': "var(--primary)", 'width': 4},
                    'thickness': 0.8,
                    'value': latest_data['健康度总分']
                },
                'bar': {'color': "var(--primary)"}
            }
        ))
        fig_gauge.update_layout(
            height=300,
            margin=dict(l=20, r=20, t=50, b=20),
            font={'color': "var(--text-secondary)"}
        )
        st.plotly_chart(fig_gauge, use_container_width=True)
    
    with col2:
        st.markdown("<div class='section fade-in'>", unsafe_allow_html=True)
        st.markdown("#### 各维度健康得分")
        
        categories = ['资源利用', '客户健康', '财务健康', '风险控制', '增长潜力']
        scores = [
            latest_data['资源利用得分']/weights['资源利用'],
            latest_data['客户健康得分']/weights['客户健康'],
            latest_data['财务健康得分']/weights['财务健康'],
            latest_data['风险控制得分']/weights['风险控制'],
            latest_data['增长潜力得分']/weights['增长潜力']
        ]
        
        # 使用条形图展示各维度得分
        fig = px.bar(
            x=categories, 
            y=scores,
            labels={'x': '维度', 'y': '得分'},
            text=[f"{s:.1f}" for s in scores],
            color=categories,
            color_discrete_sequence=['#165DFF', '#69b1ff', '#4080FF', '#85ADFF', '#B8D0FF']
        )
        fig.update_layout(
            yaxis_range=[0, 100],
            showlegend=False,
            height=350,
            xaxis_title=None,
            yaxis_title="得分",
            font={'color': "var(--text-secondary)"},
            plot_bgcolor='rgba(0, 0, 0, 0)',
            paper_bgcolor='rgba(0, 0, 0, 0)'
        )
        fig.update_traces(textfont_size=14, textangle=0, textposition="outside")
        st.plotly_chart(fig, use_container_width=True)
        
        # 各维度评分
        st.markdown("**维度评分详情**")
        dim_cols = st.columns(5)
        dim_colors = ['#165DFF', '#69b1ff', '#4080FF', '#85ADFF', '#B8D0FF']
        for i, dim in enumerate(categories):
            with dim_cols[i]:
                st.markdown(f"""
                <div style='background-color:{dim_colors[i]}10; padding:10px; border-radius:8px; border-left:3px solid {dim_colors[i]}'>
                    <div style='font-size:13px; color:var(--text-secondary)'>{dim}</div>
                    <div style='font-size:18px; font-weight:700; color:{dim_colors[i]}'>{scores[i]:.1f}</div>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown("<div class='section fade-in'>", unsafe_allow_html=True)
        st.markdown("#### 关键绩效指标")
        
        # 资源利用指标
        st.markdown("<div class='metric-card hover-card'>", unsafe_allow_html=True)
        st.markdown("<div class='metric-label'>服务器利用率</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>{latest_data['服务器利用率']:.1f}%</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class="progress-container">
            <div class="progress-bar" style="width: {latest_data['服务器利用率']}%; background: linear-gradient(90deg, #85ADFF, #165DFF);"></div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        # 客户健康指标
        st.markdown("<div class='metric-card hover-card'>", unsafe_allow_html=True)
        st.markdown("<div class='metric-label'>客户流失率</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>{latest_data['客户流失率']:.1f}%</div>", unsafe_allow_html=True)
        prev_loss = df.iloc[-2]['客户流失率'] if len(df) > 1 else latest_data['客户流失率']
        trend = "trend-down" if latest_data['客户流失率'] > prev_loss else "trend-up"
        change = abs(latest_data['客户流失率'] - prev_loss)
        st.markdown(f"<div class='{trend}'>{'↑ 改善' if latest_data['客户流失率'] < prev_loss else '↓ 恶化'} {change:.1f}%</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        # 财务健康指标
        st.markdown("<div class='metric-card hover-card'>", unsafe_allow_html=True)
        st.markdown("<div class='metric-label'>月收入</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='metric-value'>{latest_data['月收入(万元)']:.1f} 万元</div>", unsafe_allow_html=True)
        prev_rev = df.iloc[-2]['月收入(万元)'] if len(df) > 1 else latest_data['月收入(万元)']
        trend = "trend-up" if latest_data['月收入(万元)'] > prev_rev else "trend-down"
        change_pct = ((latest_data['月收入(万元)'] - prev_rev) / prev_rev * 100) if prev_rev != 0 else 0
        st.markdown(f"<div class='{trend}'>{'+' if latest_data['月收入(万元)'] > prev_rev else ''}{change_pct:.1f}%</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    # 趋势分析
    st.subheader("📈 健康度趋势分析", divider="blue")
    st.markdown("<div class='section fade-in'>", unsafe_allow_html=True)
    
    # 健康度总分趋势
    tab1, tab2, tab3 = st.tabs(["健康度总分", "各维度趋势", "指标对比"])
    
    with tab1:
        fig_trend = px.line(
            df, 
            x='月份', 
            y='健康度总分', 
            title='健康度总分变化趋势',
            markers=True
        )
        fig_trend.update_traces(
            line=dict(width=4, color='var(--primary)'),
            marker=dict(size=8, color='var(--primary)', line=dict(width=2, color='white'))
        )
        fig_trend.add_hrect(
            y0=85, y1=100, 
            fillcolor="rgba(82, 196, 26, 0.1)", 
            layer="below", 
            annotation_text="优秀", 
            annotation_position="top left"
        )
        fig_trend.add_hrect(
            y0=70, y1=85, 
            fillcolor="rgba(24, 144, 255, 0.1)", 
            layer="below", 
            annotation_text="良好"
        )
        fig_trend.add_hrect(
            y0=50, y1=70, 
            fillcolor="rgba(250, 173, 20, 0.1)", 
            layer="below", 
            annotation_text="一般"
        )
        fig_trend.add_hrect(
            y0=0, y1=50, 
            fillcolor="rgba(255, 77, 79, 0.1)", 
            layer="below", 
            annotation_text="危险"
        )
        fig_trend.update_layout(
            height=450,
            xaxis_title="月份",
            yaxis_title="健康度总分",
            hovermode="x unified",
            font={'color': "var(--text-secondary)"},
            plot_bgcolor='rgba(0, 0, 0, 0)',
            paper_bgcolor='rgba(0, 0, 0, 0)',
            title_font={'size': 18, 'color': "var(--text-primary)"}
        )
        st.plotly_chart(fig_trend, use_container_width=True)
    
    with tab2:
        # 各维度趋势图
        fig_dims = go.Figure()
        dim_colors = ['#165DFF', '#69b1ff', '#4080FF', '#85ADFF', '#B8D0FF']
        dimensions = ['资源利用得分', '客户健康得分', '财务健康得分', '风险控制得分', '增长潜力得分']
        dim_names = ['资源利用', '客户健康', '财务健康', '风险控制', '增长潜力']
        
        for i, dim in enumerate(dimensions):
            fig_dims.add_trace(go.Scatter(
                x=df['月份'], 
                y=df[dim]/weights[dim_names[i]],
                name=dim_names[i],
                line=dict(width=3, color=dim_colors[i]),
                mode='lines+markers',
                marker=dict(size=6, line=dict(width=1, color='white'))
            ))
        
        fig_dims.update_layout(
            title='各维度健康得分趋势',
            height=450,
            xaxis_title="月份",
            yaxis_title="得分",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            hovermode="x unified",
            font={'color': "var(--text-secondary)"},
            plot_bgcolor='rgba(0, 0, 0, 0)',
            paper_bgcolor='rgba(0, 0, 0, 0)',
            title_font={'size': 18, 'color': "var(--text-primary)"}
        )
        st.plotly_chart(fig_dims, use_container_width=True)
    
    with tab3:
        # 选择要分析的指标
        selected_metrics = st.multiselect(
            "选择对比指标", 
            options=['服务器利用率', '带宽利用率', '机柜利用率', '新客户数量', '客户流失率', 
                    '平均合同期限', '月收入(万元)', '利润率', '应收账款周转天数', 
                    '高风险客户占比', '服务中断次数', '市场增长率', '销售漏斗数量'],
            default=['月收入(万元)', '利润率', '客户流失率']
        )
        
        if selected_metrics:
            fig_metrics = go.Figure()
            colors = px.colors.qualitative.Plotly
            
            for i, metric in enumerate(selected_metrics):
                fig_metrics.add_trace(go.Scatter(
                    x=df['月份'], 
                    y=df[metric], 
                    mode='lines+markers',
                    name=metric,
                    line=dict(width=3, color=colors[i % len(colors)]),
                    marker=dict(size=6, line=dict(width=1, color='white'))
                ))
            
            fig_metrics.update_layout(
                title="关键指标趋势对比",
                height=450,
                xaxis_title="月份",
                yaxis_title="指标值",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                ),
                hovermode="x unified",
                font={'color': "var(--text-secondary)"},
                plot_bgcolor='rgba(0, 0, 0, 0)',
                paper_bgcolor='rgba(0, 0, 0, 0)',
                title_font={'size': 18, 'color': "var(--text-primary)"}
            )
            st.plotly_chart(fig_metrics, use_container_width=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# 风险分析页面
if page_nav == "风险洞察" and st.session_state.data is not None:
    if not st.session_state.analysis_complete:
        st.session_state.results = calculate_health_scores(st.session_state.data.copy())
        st.session_state.analysis_complete = True
    
    df = st.session_state.results
    latest_data = df.iloc[-1]
    
    st.subheader("风险分析与优化建议", divider="blue")
    st.markdown("<div class='section fade-in'>", unsafe_allow_html=True)
    
    # 识别主要风险点
    risk_points = []
    if latest_data['客户流失率'] > 3.0:
        risk_points.append(("客户流失率过高", 
                          f"当前流失率 {latest_data['客户流失率']:.1f}%，高于3%的安全阈值",
                          "高"))
    
    if latest_data['高风险客户占比'] > 10.0:
        risk_points.append(("高风险客户过多", 
                          f"高风险客户占比 {latest_data['高风险客户占比']:.1f}%，高于10%的安全阈值",
                          "高"))
    
    if latest_data['应收账款周转天数'] > 45.0:
        risk_points.append(("回款周期过长", 
                          f"应收账款周转天数 {latest_data['应收账款周转天数']}天，高于45天的安全阈值",
                          "中"))
    
    if latest_data['服务中断次数'] > 1.0:
        risk_points.append(("服务稳定性问题", 
                          f"服务中断次数 {latest_data['服务中断次数']}次，影响客户满意度",
                          "高"))
    
    if latest_data['销售漏斗数量'] < 40.0:
        risk_points.append(("销售机会不足", 
                          f"销售漏斗数量仅 {latest_data['销售漏斗数量']}，低于40的安全阈值",
                          "中"))
    
    # 风险概览卡片
    col_sum, col_high, col_medium = st.columns(3)
    with col_sum:
        st.markdown(f"""
        <div class='metric-card hover-card'>
            <div class='metric-label'>风险点总数</div>
            <div class='metric-value'>{len(risk_points)}</div>
        </div>
        """, unsafe_allow_html=True)
    
    high_risk = sum(1 for p in risk_points if p[2] == "高")
    with col_high:
        st.markdown(f"""
        <div class='metric-card hover-card'>
            <div class='metric-label'>高风险点</div>
            <div class='metric-value danger'>{high_risk}</div>
        </div>
        """, unsafe_allow_html=True)
    
    medium_risk = sum(1 for p in risk_points if p[2] == "中")
    with col_medium:
        st.markdown(f"""
        <div class='metric-card hover-card'>
            <div class='metric-label'>中风险点</div>
            <div class='metric-value fair'>{medium_risk}</div>
        </div>
        """, unsafe_allow_html=True)
    
    # 风险详情
    if risk_points:
        st.warning(f"发现{len(risk_points)}个风险点需要关注：")
        
        for i, (risk, detail, severity) in enumerate(risk_points):
            severity_color = "var(--danger)" if severity == "高" else "var(--warning)"
            st.markdown(f"""
            <div class='risk-card fade-in'>
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div>
                        <b>{i+1}. {risk}</b>
                        <div style="font-size:14px; margin-top:5px; color:var(--text-secondary);">{detail}</div>
                    </div>
                    <div style="background:{severity_color}; color:white; padding:4px 12px; border-radius:12px; font-weight:600;">
                        {severity}风险
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.success("✅ 未发现重大风险点，当前销售健康状况良好！")
    
    # 优化建议
    st.markdown("---")
    st.info("💡 综合优化建议：")
    
    rec_cols = st.columns(3)
    
    with rec_cols[0]:
        st.markdown("""
        <div style='background-color:#E8F3FF; padding:15px; border-radius:12px;' class='hover-card'>
            <h4 style='color:#165DFF; margin-top:0;'>资源优化建议</h4>
            <ul style='margin-bottom:0;'>
                <li>对利用率低于80%的资源进行整合</li>
                <li>优化机柜空间分配策略</li>
                <li>实施动态资源调配机制</li>
                <li>淘汰老旧低效设备</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with rec_cols[1]:
        st.markdown("""
        <div style='background-color:#E6F7F0; padding:15px; border-radius:12px;' class='hover-card'>
            <h4 style='color:#52C41A; margin-top:0;'>客户管理建议</h4>
            <ul style='margin-bottom:0;'>
                <li>建立客户健康度评分体系</li>
                <li>对高风险客户进行信用评估</li>
                <li>实施客户挽留计划</li>
                <li>优化客户服务响应流程</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with rec_cols[2]:
        st.markdown("""
        <div style='background-color:#FFF7E8; padding:15px; border-radius:12px;' class='hover-card'>
            <h4 style='color:#FAAD14; margin-top:0;'>财务优化建议</h4>
            <ul style='margin-bottom:0;'>
                <li>优化收款流程，缩短回款周期</li>
                <li>实施阶梯式定价策略</li>
                <li>加强应收账款管理</li>
                <li>优化成本结构</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# 报告导出页面
if page_nav == "报告导出" and st.session_state.data is not None:
    if not st.session_state.analysis_complete:
        st.session_state.results = calculate_health_scores(st.session_state.data.copy())
        st.session_state.analysis_complete = True
    
    df = st.session_state.results
    
    st.subheader("分析报告导出", divider="blue")
    st.markdown("<div class='section fade-in'>", unsafe_allow_html=True)
    
    # 报告配置选项
    st.markdown("### 报告设置")
    report_cols = st.columns(3)
    
    with report_cols[0]:
        report_title = st.text_input("报告标题", "IDC销售健康度分析报告")
    
    with report_cols[1]:
        company_name = st.text_input("公司名称", "ABC数据中心")
    
    with report_cols[2]:
        report_date = st.date_input("报告日期", datetime.today())
    
    # 导出选项
    st.markdown("### 导出选项")
    export_cols = st.columns(3)
    
    with export_cols[0]:
        if st.button("导出CSV报告", use_container_width=True):
            try:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="下载CSV报告",
                    data=csv,
                    file_name=f"{company_name}_IDC健康度报告_{report_date.strftime('%Y%m%d')}.csv",
                    mime='text/csv',
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"CSV导出失败: {str(e)}")
    
    with export_cols[1]:
        if st.button("导出Excel报告", use_container_width=True):
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='健康度数据')
                    
                    # 添加摘要工作表
                    workbook = writer.book
                    summary_sheet = workbook.create_sheet("报告摘要")
                    
                    # 添加摘要内容
                    summary_sheet['A1'] = report_title
                    summary_sheet['A2'] = f"{company_name} | {report_date.strftime('%Y-%m-%d')}"
                    summary_sheet['A4'] = "健康度总分"
                    summary_sheet['B4'] = df.iloc[-1]['健康度总分']
                    summary_sheet['A5'] = "健康度等级"
                    summary_sheet['B5'] = df.iloc[-1]['健康度等级']
                    
                st.download_button(
                    label="下载Excel报告",
                    data=output.getvalue(),
                    file_name=f"{company_name}_IDC健康度报告_{report_date.strftime('%Y%m%d')}.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Excel导出失败: {str(e)}")
    
    with export_cols[2]:
        st.markdown("""
        <div style='background-color:#F0F2F5; padding:20px; border-radius:12px; height:100%; display:flex; flex-direction:column; justify-content:center; align-items:center; text-align:center;' class='hover-card'>
            <div style='font-size:40px; margin-bottom:10px;'>📄</div>
            <div style='font-weight:600; margin-bottom:5px;'>PDF报告导出</div>
            <div style='font-size:13px; color:var(--text-tertiary);'>专业PDF报告生成功能即将上线，敬请期待</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# 趋势预测页面
if page_nav == "趋势预测" and st.session_state.data is not None:
    if not st.session_state.analysis_complete:
        st.session_state.results = calculate_health_scores(st.session_state.data.copy())
        st.session_state.analysis_complete = True
    
    df = st.session_state.results
    
    st.subheader("未来趋势预测", divider="blue")
    st.markdown("<div class='section fade-in'>", unsafe_allow_html=True)
    
    st.info("基于历史数据的线性回归预测，结果仅供参考，实际业务需结合更多因素分析")
    
    # 选择预测指标
    pred_metric = st.selectbox(
        "选择预测指标",
        options=['健康度总分', '服务器利用率', '带宽利用率', '机柜利用率', 
                '客户流失率', '月收入(万元)', '利润率', '销售漏斗数量'],
        index=0
    )
    
    # 预测周期
    periods = st.slider("预测周期（月）", 1, 12, 6)
    
    if st.button("生成预测", type="primary"):
        with st.spinner("正在生成预测..."):
            time.sleep(1)
            
            # 简单线性预测
            last_date = df['月份'].iloc[-1]
            future_dates = [last_date + timedelta(days=30*i) for i in range(1, periods+1)]
            
            # 使用最后6个月数据进行预测
            y = df[pred_metric].values[-6:]
            x = np.arange(len(y))
            
            # 线性回归
            coeff = np.polyfit(x, y, 1)
            future_vals = coeff[0] * np.arange(len(y), len(y)+periods) + coeff[1]
            
            # 确保预测值在合理范围内
            if pred_metric in ['服务器利用率', '带宽利用率', '机柜利用率', '利润率', '高风险客户占比', '客户流失率']:
                future_vals = np.clip(future_vals, 0, 100)
            elif pred_metric == '健康度总分':
                future_vals = np.clip(future_vals, 0, 100)
            elif pred_metric in ['服务中断次数']:
                future_vals = np.clip(future_vals, 0, None)
            
            # 创建预测数据框
            forecast_df = pd.DataFrame({
                '月份': future_dates,
                pred_metric: future_vals,
                '类型': '预测值'
            })
            
            # 历史数据
            history_df = pd.DataFrame({
                '月份': df['月份'],
                pred_metric: df[pred_metric],
                '类型': '历史值'
            })
            
            # 合并数据
            full_df = pd.concat([history_df, forecast_df])
            
            # 绘制预测图
            fig = px.line(
                full_df, 
                x='月份', 
                y=pred_metric,
                color='类型',
                color_discrete_map={'历史值': 'var(--primary)', '预测值': 'var(--danger)'},
                title=f'{pred_metric}趋势预测'
            )
            
            # 添加最后历史点
            fig.add_trace(go.Scatter(
                x=[df['月份'].iloc[-1]], 
                y=[df[pred_metric].iloc[-1]],
                mode='markers',
                marker=dict(size=10, color='var(--primary)', line=dict(width=2, color='white')),
                name='当前值'
            ))
            
            # 添加预测开始点
            fig.add_trace(go.Scatter(
                x=[future_dates[0]], 
                y=[future_vals[0]],
                mode='markers',
                marker=dict(size=10, color='var(--danger)', line=dict(width=2, color='white')),
                name='预测起点'
            ))
            
            # 添加置信区间阴影
            fig.update_layout(
                height=500,
                xaxis_title="月份",
                yaxis_title=pred_metric,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                ),
                font={'color': "var(--text-secondary)"},
                plot_bgcolor='rgba(0, 0, 0, 0)',
                paper_bgcolor='rgba(0, 0, 0, 0)',
                title_font={'size': 18, 'color': "var(--text-primary)"}
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # 显示预测摘要
            change_pct = ((future_vals[-1] - df[pred_metric].iloc[-1]) / df[pred_metric].iloc[-1]) * 100
            st.markdown("<div class='metric-card hover-card'>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-label'>{periods}个月后预测值</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric-value'>{future_vals[-1]:.1f}</div>", unsafe_allow_html=True)
            trend_class = "trend-up" if future_vals[-1] > df[pred_metric].iloc[-1] else "trend-down"
            st.markdown(f"<div class='{trend_class}'>与当前相比: {'+' if future_vals[-1] > df[pred_metric].iloc[-1] else ''}{change_pct:.1f}%</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# 初始页面状态
if st.session_state.data is None and page_nav != "数据导入":
    st.info("👆 请先导入数据或使用示例数据开始分析")
    # 替换图片为文字和图标
    st.markdown("""
    <div style='text-align:center; padding:40px 0; background-color:#F7F8FA; border-radius:12px;'>
        <div style='font-size:60px; margin-bottom:20px;'>📊</div>
        <h3 style='color:#4E5969; margin-bottom:10px;'>上传数据开始分析</h3>
        <p style='color:#86909C;'>请前往"数据导入"页面上传您的IDC销售数据或使用示例数据</p>
    </div>
    """, unsafe_allow_html=True)

# 底部信息
st.markdown("---")
footer_cols = st.columns(3)
with footer_cols[0]:
    st.caption("© 2025 IDC资源销售健康度分析工具")
with footer_cols[1]:
    st.caption("版本 1.0.0 | 数据更新时间: " + st.session_state.last_updated)
with footer_cols[2]:
    st.caption("在岗革新: 云南网络班组")

# 性能优化提示
st.sidebar.markdown("---")
st.sidebar.info("**性能优化提示**")
st.sidebar.markdown("""
- 大型数据集建议使用CSV格式
- 超过10万行数据时启用抽样分析
- 关闭不需要的可视化图表
- 定期清理浏览器缓存
""")