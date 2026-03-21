"""
AIA 友邦储蓄险方案生成器 — Streamlit Web 应用
内部工具，需密码验证后使用。
"""

import streamlit as st
from datetime import datetime
from document_generator import generate_proposal
from benefit_data import format_usd

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="AIA 友邦储蓄险方案生成器",
    page_icon="📋",
    layout="centered",
)

# ============================================================
# 自定义样式
# ============================================================
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #003B73;
        font-size: 2rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        text-align: center;
        color: #666666;
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    .gold-line {
        border: none;
        border-top: 3px solid #C8A951;
        margin: 1rem auto;
        width: 60%;
    }
    .stButton > button {
        background-color: #003B73;
        color: white;
        font-size: 1.1rem;
        padding: 0.6rem 2rem;
        border-radius: 8px;
        width: 100%;
    }
    .stButton > button:hover {
        background-color: #004d99;
        color: white;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 密码验证
# ============================================================
ACCESS_PASSWORD = "888888"


def check_password():
    """密码验证拦截页"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown('<div class="main-header">AIA 友邦储蓄险方案生成器</div>', unsafe_allow_html=True)
    st.markdown('<hr class="gold-line">', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">内部工具 · 请输入访问密码</div>', unsafe_allow_html=True)

    with st.form("login_form"):
        password = st.text_input("访问密码", type="password", placeholder="请输入密码")
        submitted = st.form_submit_button("进入系统", use_container_width=True)

        if submitted:
            if password == ACCESS_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("密码错误，请重试。")

    return False


# ============================================================
# 主应用
# ============================================================
def main():
    # 页头
    st.markdown('<div class="main-header">AIA 友邦储蓄险方案生成器</div>', unsafe_allow_html=True)
    st.markdown('<hr class="gold-line">', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">环宇盈活储蓄保险计划 · 客户方案自动生成</div>', unsafe_allow_html=True)

    # 侧边栏退出
    with st.sidebar:
        st.markdown("### 系统信息")
        st.info(f"当前日期：{datetime.now().strftime('%Y年%m月%d日')}")
        if st.button("退出登录"):
            st.session_state.authenticated = False
            st.rerun()

    # ---- 输入表单 ----
    st.subheader("客户信息录入")

    with st.form("proposal_form"):
        col1, col2 = st.columns(2)

        with col1:
            client_name = st.text_input("客户姓名 *", placeholder="例：王女士")
            client_age = st.number_input("客户年龄 *", min_value=1, max_value=80, value=35)
            client_gender = st.selectbox("性别", ["女", "男"])
            client_occupation = st.text_input("职业", placeholder="例：IT行业高管")

        with col2:
            client_income = st.text_input("年收入", placeholder="例：100万人民币")
            client_family = st.text_input("家庭结构", placeholder="例：已婚，育有一子（3岁）")
            annual_premium = st.number_input(
                "年缴保费（美元） *",
                min_value=5000,
                max_value=10000000,
                value=50000,
                step=5000,
            )
            retirement_age = st.number_input(
                "目标退休年龄",
                min_value=0,
                max_value=100,
                value=0,
                help="填 0 表示不设定退休目标",
            )

        client_needs = st.text_area(
            "核心需求",
            placeholder="例：退休规划、子女教育金、资产隔离、财富传承等",
            height=80,
        )

        custom_notes = st.text_area(
            "自定义备注（可选）",
            placeholder="任何需要体现在方案中的补充信息",
            height=60,
        )

        submitted = st.form_submit_button("生成方案", use_container_width=True)

    # ---- 处理提交 ----
    if submitted:
        # 验证必填项
        if not client_name.strip():
            st.error("请填写客户姓名。")
            return

        # 处理退休年龄
        ret_age = retirement_age if retirement_age > 0 and retirement_age > client_age else None

        with st.spinner("正在生成方案文档，请稍候..."):
            try:
                doc_buffer = generate_proposal(
                    client_name=client_name.strip(),
                    client_age=client_age,
                    client_gender=client_gender,
                    client_occupation=client_occupation.strip(),
                    client_income=client_income.strip(),
                    client_family=client_family.strip(),
                    client_needs=client_needs.strip(),
                    annual_premium=annual_premium,
                    retirement_age=ret_age,
                    custom_notes=custom_notes.strip(),
                )

                # 生成文件名
                date_str = datetime.now().strftime("%Y%m%d")
                filename = f"{client_name}-个人财富管理方案-{date_str}.docx"

                st.success("方案生成成功！")

                # 展示方案概要
                total_premium = annual_premium * 5
                st.markdown("---")
                st.markdown("#### 方案概要")

                col1, col2, col3 = st.columns(3)
                col1.metric("客户", client_name)
                col2.metric("年缴保费", f"${format_usd(annual_premium)}")
                col3.metric("总保费", f"${format_usd(total_premium)}")

                if ret_age:
                    col1, col2, col3 = st.columns(3)
                    ret_year = ret_age - client_age
                    from benefit_data import BASE_TOTAL_SURRENDER
                    scale = annual_premium / 10000
                    ret_value = round(BASE_TOTAL_SURRENDER[ret_year] * scale)
                    withdrawal = round(ret_value * 0.065)
                    col1.metric("退休年龄", f"{ret_age}岁（第{ret_year}年）")
                    col2.metric("退休时预期总价值", f"${format_usd(ret_value)}")
                    col3.metric("年提取金额", f"${format_usd(withdrawal)}")

                # 下载按钮
                st.markdown("---")
                st.download_button(
                    label=f"下载方案文档：{filename}",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"生成方案时出错：{str(e)}")
                st.exception(e)


# ============================================================
# 入口
# ============================================================
if check_password():
    main()
