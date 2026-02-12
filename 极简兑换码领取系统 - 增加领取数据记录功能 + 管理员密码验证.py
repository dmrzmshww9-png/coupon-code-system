#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
极简兑换码领取系统 - 增加领取数据记录功能 + 管理员密码验证
"""

import streamlit as st
import pandas as pd
import os
import re
import time
import datetime
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(
    page_title="兑换码领取",
    page_icon="🎫",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 配置文件
EXCEL_FILE_NAME = "2025调研问卷-手机号清单.xlsx"
RECORD_FILE_NAME = "领取记录.xlsx"
# 管理员密码 - 在实际使用中可以修改这个密码
ADMIN_PASSWORD = "admin123"

# 极简CSS
def minimal_css():
    st.markdown("""
    <style>
    /* 基础样式 */
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }
    
    .main {
        max-width: 100%;
        padding: 1rem;
    }
    
    /* 标题 */
    .simple-title {
        text-align: center;
        font-size: 1.8rem;
        font-weight: 600;
        color: #333;
        margin-bottom: 2rem;
        padding-top: 1rem;
    }
    
    /* 输入框 */
    .simple-input {
        width: 100%;
        padding: 1rem;
        font-size: 1.1rem;
        border: 2px solid #ddd;
        border-radius: 8px;
        margin-bottom: 1rem;
        text-align: center;
        transition: border-color 0.3s;
    }
    
    .simple-input:focus {
        border-color: #4A90E2;
        outline: none;
    }
    
    /* 按钮 */
    .simple-button {
        width: 100%;
        padding: 1rem;
        font-size: 1.1rem;
        font-weight: 600;
        background-color: #4A90E2;
        color: white;
        border: none;
        border-radius: 8px;
        cursor: pointer;
        transition: background-color 0.3s;
        margin-bottom: 1rem;
    }
    
    .simple-button:hover {
        background-color: #357ABD;
    }
    
    .simple-button:disabled {
        background-color: #ccc;
        cursor: not-allowed;
    }
    
    /* 兑换码显示 */
    .coupon-box {
        margin: 1.5rem 0;
        padding: 1.5rem;
        background-color: #f8f9fa;
        border: 2px solid #4A90E2;
        border-radius: 8px;
        text-align: center;
    }
    
    .coupon-code {
        font-family: 'Courier New', monospace;
        font-size: 1.8rem;
        font-weight: 700;
        color: #333;
        letter-spacing: 1px;
        word-break: break-all;
    }
    
    /* 提醒信息 */
    .alert-box {
        margin: 1rem 0;
        padding: 1rem;
        border-radius: 8px;
        font-size: 0.95rem;
        line-height: 1.5;
    }
    
    .alert-success {
        background-color: #d4edda;
        color: #155724;
        border: 1px solid #c3e6cb;
    }
    
    .alert-error {
        background-color: #f8d7da;
        color: #721c24;
        border: 1px solid #f5c6cb;
    }
    
    .alert-info {
        background-color: #d1ecf1;
        color: #0c5460;
        border: 1px solid #bee5eb;
    }
    
    .alert-warning {
        background-color: #fff3cd;
        color: #856404;
        border: 1px solid #ffeaa7;
    }
    
    /* 密码输入框特殊样式 */
    .password-input {
        background-color: #fff8e1;
        border-color: #ffd54f !important;
    }
    
    /* 隐藏streamlit元素 */
    #MainMenu, footer, header, .stDeployButton {
        display: none;
    }
    
    /* 移动端优化 */
    @media (max-width: 768px) {
        .simple-title {
            font-size: 1.5rem;
        }
        
        .simple-input, .simple-button {
            padding: 0.9rem;
            font-size: 1rem;
        }
        
        .coupon-code {
            font-size: 1.5rem;
        }
    }
    </style>
    """, unsafe_allow_html=True)

# 初始化session state
def init_session():
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'record_df' not in st.session_state:
        st.session_state.record_df = None
    if 'phone_input' not in st.session_state:
        st.session_state.phone_input = ''
    if 'last_coupon' not in st.session_state:
        st.session_state.last_coupon = None
    if 'admin_authenticated' not in st.session_state:
        st.session_state.admin_authenticated = False
    if 'password_attempts' not in st.session_state:
        st.session_state.password_attempts = 0

# 兑换码管理器
class CouponManager:
    def __init__(self):
        self.df = None
        self.record_df = None
    
    def clean_phone(self, phone_str):
        """清洗手机号"""
        if not phone_str or pd.isna(phone_str):
            return None
        
        digits = re.sub(r'\D', '', str(phone_str))
        
        if len(digits) == 11 and digits.startswith('1'):
            return digits
        
        return None
    
    def load_excel_data(self):
        """加载主数据文件"""
        try:
            if not os.path.exists(EXCEL_FILE_NAME):
                return False, f"找不到文件: {EXCEL_FILE_NAME}"
            
            # 读取Excel
            df = pd.read_excel(EXCEL_FILE_NAME, dtype=str)
            
            # 检查必要列
            if '手机号' not in df.columns or '兑换码' not in df.columns:
                return False, "Excel缺少'手机号'或'兑换码'列"
            
            # 清理数据
            df = df.copy()
            
            # 清洗手机号列
            df['清洗后手机号'] = df['手机号'].apply(self.clean_phone)
            
            # 添加状态列
            if '状态' not in df.columns:
                df['状态'] = '未发放'
            
            if '领取时间' not in df.columns:
                df['领取时间'] = ''
            
            # 确保兑换码是字符串类型
            df['兑换码'] = df['兑换码'].astype(str).str.strip()
            
            # 修复兑换码重复问题
            df = self.fix_duplicate_coupons(df)
            
            self.df = df
            st.session_state.df = df
            
            return True, f"成功加载 {len(df)} 条记录"
            
        except Exception as e:
            return False, f"加载失败: {str(e)}"
    
    def fix_duplicate_coupons(self, df):
        """修复兑换码重复问题"""
        if df is None or df.empty:
            return df
        
        # 确保每行都有兑换码
        df['兑换码'] = df['兑换码'].fillna('')
        
        # 修复重复模式
        def fix_coupon(coupon):
            if not coupon or len(coupon) < 2:
                return coupon
            
            coupon = str(coupon).strip()
            
            # 检查是否是完全重复模式（如842842）
            if len(coupon) % 2 == 0:
                half_len = len(coupon) // 2
                first_half = coupon[:half_len]
                second_half = coupon[half_len:]
                if first_half == second_half:
                    return first_half
            
            # 检查是否有部分重复（去除多余字符）
            # 这里可以根据实际情况调整重复检测逻辑
            return coupon
        
        df['兑换码'] = df['兑换码'].apply(fix_coupon)
        
        # 去除重复的兑换码记录
        df = df.drop_duplicates(subset=['兑换码', '清洗后手机号'], keep='first')
        
        return df
    
    def load_record_data(self):
        """加载领取记录"""
        try:
            if os.path.exists(RECORD_FILE_NAME):
                record_df = pd.read_excel(RECORD_FILE_NAME, dtype=str)
                self.record_df = record_df
                st.session_state.record_df = record_df
                return True, f"加载 {len(record_df)} 条领取记录"
            else:
                # 创建空的领取记录DataFrame
                record_df = pd.DataFrame(columns=[
                    '手机号', 
                    '兑换码', 
                    '领取时间',
                    'IP地址',
                    '用户代理'
                ])
                self.record_df = record_df
                st.session_state.record_df = record_df
                return True, "创建新的领取记录文件"
        except Exception as e:
            return False, f"加载领取记录失败: {str(e)}"
    
    def save_record_data(self):
        """保存领取记录到文件"""
        try:
            if self.record_df is not None:
                self.record_df.to_excel(RECORD_FILE_NAME, index=False)
                return True, "领取记录保存成功"
            return False, "无领取记录可保存"
        except Exception as e:
            return False, f"保存领取记录失败: {str(e)}"
    
    def add_claim_record(self, phone, coupon):
        """添加领取记录"""
        if self.record_df is None:
            # 初始化记录DataFrame
            self.record_df = pd.DataFrame(columns=[
                '手机号', 
                '兑换码', 
                '领取时间',
                'IP地址',
                '用户代理'
            ])
        
        # 创建新记录
        new_record = pd.DataFrame([{
            '手机号': phone,
            '兑换码': coupon,
            '领取时间': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'IP地址': 'N/A',  # 在实际应用中可以通过request获取
            '用户代理': 'N/A'  # 在实际应用中可以通过request获取
        }])
        
        # 添加到记录DataFrame
        self.record_df = pd.concat([self.record_df, new_record], ignore_index=True)
        
        # 保存到文件
        self.save_record_data()
        
        # 更新session state
        st.session_state.record_df = self.record_df
    
    def find_and_claim(self, phone):
        """查找并领取兑换码"""
        if self.df is None or self.df.empty:
            return False, "数据未加载", None
        
        # 清洗输入的手机号
        clean_phone = self.clean_phone(phone)
        if not clean_phone:
            return False, "请输入11位有效手机号", None
        
        # 查找匹配的手机号
        matches = self.df[self.df['清洗后手机号'] == clean_phone]
        
        if matches.empty:
            return False, "该手机号不在领取名单中", None
        
        # 取第一个匹配记录
        record = matches.iloc[0]
        
        # 检查状态
        if record['状态'] == '已发放':
            return False, "该兑换码已被领取", None
        
        if record['状态'] != '未发放':
            return False, f"兑换码状态不可用", None
        
        # 获取兑换码
        coupon_code = str(record['兑换码']).strip()
        
        # 再次检查并修复重复问题
        if len(coupon_code) % 2 == 0:
            half_len = len(coupon_code) // 2
            first_half = coupon_code[:half_len]
            second_half = coupon_code[half_len:]
            if first_half == second_half:
                coupon_code = first_half
        
        # 更新主数据状态
        idx = record.name
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        self.df.at[idx, '状态'] = '已发放'
        self.df.at[idx, '领取时间'] = current_time
        
        # 保存主数据更新
        st.session_state.df = self.df
        
        # 添加领取记录（不显示给用户）
        self.add_claim_record(phone, coupon_code)
        
        return True, "领取成功", coupon_code
    
    def get_record_excel(self):
        """获取领取记录的Excel数据"""
        if self.record_df is None or self.record_df.empty:
            return None
        
        # 创建Excel文件
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.record_df.to_excel(writer, sheet_name='领取记录', index=False)
        
        return output.getvalue()

def check_admin_password(input_password):
    """检查管理员密码"""
    if input_password == ADMIN_PASSWORD:
        st.session_state.admin_authenticated = True
        st.session_state.password_attempts = 0  # 重置尝试次数
        return True
    else:
        st.session_state.password_attempts += 1
        return False

def admin_login_section():
    """管理员登录区域"""
    st.markdown("### 🔐 管理员登录")
    
    # 警告信息
    if st.session_state.password_attempts > 0:
        st.markdown(f"""
        <div class="alert-box alert-warning">
            ⚠️ 密码错误！已尝试 {st.session_state.password_attempts} 次
        </div>
        """, unsafe_allow_html=True)
    
    # 密码输入框
    password_input = st.text_input(
        "请输入管理员密码",
        type="password",
        placeholder="输入密码...",
        key="admin_password_input"
    )
    
    # 登录按钮
    col1, col2 = st.columns([2, 1])
    with col1:
        if st.button("登录", use_container_width=True, type="primary"):
            if password_input:
                if check_admin_password(password_input):
                    st.success("登录成功！")
                    time.sleep(0.5)
                    st.rerun()
                else:
                    st.error("密码错误！")
            else:
                st.warning("请输入密码")
    
    with col2:
        if st.button("重置", use_container_width=True):
            st.session_state.password_attempts = 0
            st.rerun()

# 页面渲染
def render_header():
    """标题"""
    st.markdown('<div class="simple-title">🎫 兑换码领取</div>', unsafe_allow_html=True)

def render_input_section(manager):
    """输入区域"""
    # 手机号输入
    phone_input = st.text_input(
        "",
        value=st.session_state.phone_input,
        placeholder="请输入11位手机号",
        key="phone_input_field",
        max_chars=11
    )
    
    # 更新session state
    st.session_state.phone_input = phone_input
    
    # 按钮
    col1, col2 = st.columns([3, 1])
    
    with col1:
        claim_clicked = st.button(
            "领取兑换码",
            type="primary",
            disabled=not phone_input,
            use_container_width=True,
            key="claim_button"
        )
    
    with col2:
        if st.button("清空", use_container_width=True):
            st.session_state.phone_input = ''
            st.session_state.last_coupon = None
            st.rerun()
    
    return phone_input, claim_clicked

def render_result(manager, phone, claim_clicked):
    """结果显示"""
    if not claim_clicked or not phone:
        return
    
    with st.spinner("正在处理..."):
        time.sleep(0.5)
        success, message, coupon = manager.find_and_claim(phone)
    
    if success and coupon:
        # 显示成功信息
        st.markdown(f"""
        <div class="alert-box alert-success">
            ✅ {message}
        </div>
        """, unsafe_allow_html=True)
        
        # 显示兑换码
        st.markdown(f"""
        <div class="coupon-box">
            <div class="coupon-code">{coupon}</div>
        </div>
        """, unsafe_allow_html=True)
        
        # 保存最后领取的兑换码
        st.session_state.last_coupon = coupon
        
        # 使用提示
        st.markdown("""
        <div class="alert-box alert-info">
            💡 请立即记录兑换码，每个手机号只能领取一次
        </div>
        """, unsafe_allow_html=True)
        
        # 继续领取按钮
        if st.button("继续领取", use_container_width=True):
            st.session_state.phone_input = ''
            st.session_state.last_coupon = None
            st.rerun()
        
    else:
        # 显示错误信息
        st.markdown(f"""
        <div class="alert-box alert-error">
            ❌ {message}
        </div>
        """, unsafe_allow_html=True)
        
        # 错误提示
        st.markdown("""
        <div class="alert-box alert-info">
            🔍 请检查手机号是否正确或是否已领取
        </div>
        """, unsafe_allow_html=True)

def render_admin_panel(manager):
    """管理员面板（折叠）"""
    with st.expander("管理选项", expanded=False):
        
        # 如果未认证，显示登录界面
        if not st.session_state.admin_authenticated:
            admin_login_section()
            return
        
        # 已认证，显示管理功能
        st.markdown("### ✅ 管理员面板")
        st.markdown(f"登录时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 登出按钮
        if st.button("登出管理员", type="secondary", use_container_width=True):
            st.session_state.admin_authenticated = False
            st.success("已退出管理员模式")
            time.sleep(0.5)
            st.rerun()
        
        st.markdown("---")
        
        # 显示统计信息
        if st.session_state.df is not None:
            total = len(st.session_state.df)
            available = len(st.session_state.df[st.session_state.df['状态'] == '未发放'])
            claimed = len(st.session_state.df[st.session_state.df['状态'] == '已发放'])
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("总记录", total)
            with col2:
                st.metric("可领取", available)
            with col3:
                st.metric("已领取", claimed)
        
        # 领取记录统计
        if st.session_state.record_df is not None:
            record_count = len(st.session_state.record_df)
            st.info(f"领取记录数: {record_count}")
        
        st.markdown("---")
        
        # 操作按钮
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔄 重新加载数据", use_container_width=True):
                success, msg = manager.load_excel_data()
                if success:
                    st.success("主数据加载成功")
                    # 重新加载记录数据
                    record_success, record_msg = manager.load_record_data()
                    if record_success:
                        st.success("领取记录加载成功")
                else:
                    st.error(f"加载失败: {msg}")
                time.sleep(1)
                st.rerun()
        
        with col2:
            # 下载主数据
            if st.session_state.df is not None:
                main_excel_data = BytesIO()
                st.session_state.df.to_excel(main_excel_data, index=False)
                
                st.download_button(
                    label="📥 下载主数据",
                    data=main_excel_data.getvalue(),
                    file_name=f"主数据_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        st.markdown("---")
        
        # 领取记录下载（核心新增功能）
        if st.session_state.record_df is not None and len(st.session_state.record_df) > 0:
            st.markdown("#### 领取记录下载")
            
            # 显示最近5条记录预览（不显示完整信息）
            recent_records = st.session_state.record_df.tail(5)
            st.dataframe(recent_records[['手机号', '兑换码', '领取时间']], use_container_width=True)
            
            # 下载按钮
            excel_data = manager.get_record_excel()
            if excel_data:
                st.download_button(
                    label="📊 下载领取记录",
                    data=excel_data,
                    file_name=f"领取记录_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="此文件包含所有用户的领取记录，用于后台管理",
                    use_container_width=True
                )
        else:
            st.info("暂无领取记录")
            
        st.markdown("---")
        st.markdown("""
        <div class="alert-box alert-warning">
            ⚠️ 注意：管理员功能仅供内部使用，操作后将记录在日志中
        </div>
        """, unsafe_allow_html=True)

# 主函数
def main():
    # 应用CSS
    minimal_css()
    
    # 初始化session
    init_session()
    
    # 创建管理器
    manager = CouponManager()
    
    # 自动加载数据
    if st.session_state.df is None:
        with st.spinner("正在加载数据..."):
            success, message = manager.load_excel_data()
            if not success:
                st.error(message)
    
    # 自动加载领取记录（不显示给普通用户）
    if st.session_state.record_df is None:
        record_success, record_msg = manager.load_record_data()
    
    # 更新管理器数据
    if st.session_state.df is not None:
        manager.df = st.session_state.df
    
    if st.session_state.record_df is not None:
        manager.record_df = st.session_state.record_df
    
    # 渲染页面
    render_header()
    
    # 如果有成功领取的兑换码，直接显示
    if st.session_state.last_coupon:
        st.markdown(f"""
        <div class="alert-box alert-success">
            ✅ 上次领取成功
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="coupon-box">
            <div class="coupon-code">{st.session_state.last_coupon}</div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("领取新的兑换码", type="primary", use_container_width=True):
            st.session_state.phone_input = ''
            st.session_state.last_coupon = None
            st.rerun()
    
    else:
        # 正常输入流程
        phone_input, claim_clicked = render_input_section(manager)
        render_result(manager, phone_input, claim_clicked)
    
    # 管理员面板
    render_admin_panel(manager)

if __name__ == "__main__":
    main()