import streamlit as st
import pandas as pd
import io
# 导入你原来的逻辑类
from main_scoring import CorrectScoringSystem
from calendar import c

# 设置网页标题和图标
st.set_page_config(page_title="专家评分分析系统", page_icon="📊", layout="wide")

st.title("📊 专家评分自动分析系统")
st.markdown("---")

# 1. 初始化引擎
if 'engine' not in st.session_state:
    st.session_state.engine = CorrectScoringSystem()

# 2. 侧边栏：设置评分规则
st.sidebar.header("⚙️ 评分规则设置")
mid_score = st.sidebar.slider("误差在 8-15 分之间给多少分？", 0, 2, 1)
st.sidebar.info("规则说明：\n- 专家标准分最接近平均分：3分\n- 误差 ≤ 8分：2分\n- 误差 > 15分：0分")

# 3. 主界面：文件上传
uploaded_file = st.file_uploader("请上传专家评分 Excel 文件 (支持 .xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    engine = st.session_state.engine

    try:
        # 加载数据
        engine.data = pd.read_excel(uploaded_file, skiprows=2, header=None)
        st.success(f"✅ 成功加载 {len(engine.data)} 行数据")

        with st.expander("👀 查看原始数据预览"):
            st.dataframe(engine.data.head(10))

        # 4. 执行计算按钮
        if st.button("🚀 开始分析专家得分"):
            with st.spinner('计算中...'):
                engine.expert_scores.clear()
                engine.calculate_scores(mid_range_score=mid_score)

                # --- 提取细化后的结果 ---
                results = []
                for name, data in engine.expert_scores.items():
                    if data['review_count'] > 0:
                        avg_score = data['total_score'] / data['review_count']
                        # 获取次数统计字典
                        c = data['counts']

                        results.append({
                            '排名': 0,
                            '专家姓名': name,
                            '总得分': data['total_score'],
                            '评审数': data['review_count'],
                            '3分次数': c[3],
                            '2分次数': c[2],
                            '1分次数': c[1],
                            '0分次数': c[0],
                            '平均分': round(avg_score, 2),
                            '得分效率(%)': round(avg_score / 3 * 100, 1)
                        })

                if not results:
                    st.warning("未能提取到有效数据，请检查 Excel 格式。")
                else:
                    # 排序：按总得分降序
                    results.sort(key=lambda x: x['总得分'], reverse=True)
                    for i, item in enumerate(results, 1):
                        item['排名'] = i

                    df_res = pd.DataFrame(results)

                    # 展示统计看板
                    st.subheader("📊 关键指标统计")
                    # ... (看板代码保持不变) ...

                    # 展示表格 (重点：现在会显示次数列了)
                    st.subheader("🏆 专家评分排名全表 (含分值分布)")
                    st.dataframe(df_res, use_container_width=True, hide_index=True)

                    # 导出 Excel (包含细化列)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, index=False, sheet_name='专家评分细化分析')

                    st.download_button(
                        label="📥 下载详细分析结果 (Excel)",
                        data=output.getvalue(),
                        file_name="专家评分细化统计.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # 5. 展示关键指标统计看板
                    st.subheader("📊 关键指标统计")
                    col1, col2, col3, col4 = st.columns(4)

                    col1.metric("涉及专家总量", f"{len(results)} 位")

                    # 最高平均分专家
                    col2.metric("最高平均分专家", results[0]['专家姓名'], f"{results[0]['平均分']} 分")

                    # 全场平均效率
                    avg_eff = sum(item['得分效率(%)'] for item in results) / len(results)
                    col3.metric("全场平均效率", f"{avg_eff:.1f}%")

                    # 评审量最多的专家
                    most_active = max(results, key=lambda x: x['评审数'])
                    col4.metric("评审量冠军", most_active['专家姓名'], f"{most_active['评审数']} 件")

                    st.markdown("---")

                    # 6. 展示排名全表
                    st.subheader("🏆 专家评分排名全表 (按平均分排序)")
                    # hide_index=True 可以隐藏表格左侧多余的 0,1,2 索引列
                    st.dataframe(df_res, use_container_width=True, hide_index=True)

                    # 7. 导出功能
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, index=False, sheet_name='评分结果分析')

                    st.download_button(
                        label="📥 点击下载分析结果 (Excel)",
                        data=output.getvalue(),
                        file_name="专家评分结果_分析完成.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"分析过程中发生错误: {e}")
        st.info("请确保上传的文件格式与 PyCharm 中的测试文件一致。")