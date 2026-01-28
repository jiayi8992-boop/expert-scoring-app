"""
专家评分计算系统 - 最终正确版
根据原始列名精确匹配
"""

import pandas as pd
import numpy as np
from collections import defaultdict
import os
import warnings
warnings.filterwarnings('ignore')

class CorrectScoringSystem:
    """完全正确的专家评分系统"""

    def __init__(self):
        self.data = None
        self.expert_scores = defaultdict(lambda: {
            'total_score': 0,
            'review_count': 0,
            'counts': {3: 0, 2: 0, 1: 0, 0: 0},  # <--- 新增这行，存放3/2/1/0分的次数
            'details': []
        })

        # 根据你的数据结构定义列映射
        self.column_config = {
            'item_id': '作品编号',
            'avg_standard': '标准分平均分',
            'experts': {
                1: {'name': '专家一', 'raw': 'Unnamed: 6', 'std': 'Unnamed: 7'},
                2: {'name': '专家二', 'raw': 'Unnamed: 9', 'std': 'Unnamed: 10'},
                3: {'name': '专家三', 'raw': 'Unnamed: 12', 'std': 'Unnamed: 13'},
                4: {'name': '专家四', 'raw': 'Unnamed: 15', 'std': 'Unnamed: 16'},
                5: {'name': '专家五', 'raw': 'Unnamed: 18', 'std': 'Unnamed: 19'}
            }
        }

    def load_data(self, filepath):
        """加载数据 - 使用原始列名"""
        print(f"📂 正在加载数据文件: {os.path.basename(filepath)}")

        try:
            # 跳过前2行（第0行空行，第1行表头）
            self.data = pd.read_excel(filepath, skiprows=2, header=None)
            print(f"✅ 成功加载 {len(self.data)} 行数据")

            # 显示数据示例
            print("\n📊 数据示例（前3行）:")
            for i in range(min(3, len(self.data))):
                row = self.data.iloc[i]
                print(f"  行{i+1}: 作品={row[1]}, 平均分={row[2]}, 专家1姓名={row[5]}, 专家1标准分={row[7]}")

            return True

        except Exception as e:
            print(f"❌ 加载失败: {e}")
            return False

    def calculate_scores(self, mid_range_score=1):
        """计算得分 - 使用列索引"""
        print(f"\n🧮 开始计算专家得分...")
        print(f"📐 评分规则:")
        print(f"  • 最接近标准分平均分: 3分")
        print(f"  • 误差 ≤ 8分: 2分")
        print(f"  • 误差 > 15分: 0分")
        print(f"  • 误差 8-15分: {mid_range_score}分")

        total_items = len(self.data)
        processed = 0
        valid_items = 0

        # 使用列索引（从0开始）
        # 根据调试输出：
        # 索引0: 排名, 1: 作品编号, 2: 标准分平均分, 3: 标准分极差, 4: 论文相似度
        # 索引5: 专家一姓名, 6: 专家一原始分, 7: 专家一标准分
        # 索引8: 专家二姓名, 9: 专家二原始分, 10: 专家二标准分
        # 索引11: 专家三姓名, 12: 专家三原始分, 13: 专家三标准分
        # 索引14: 专家四姓名, 15: 专家四原始分, 16: 专家四标准分
        # 索引17: 专家五姓名, 18: 专家五原始分, 19: 专家五标准分

        expert_config = [
            {'name_idx': 5, 'raw_idx': 6, 'std_idx': 7},   # 专家一
            {'name_idx': 8, 'raw_idx': 9, 'std_idx': 10},  # 专家二
            {'name_idx': 11, 'raw_idx': 12, 'std_idx': 13}, # 专家三
            {'name_idx': 14, 'raw_idx': 15, 'std_idx': 16}, # 专家四
            {'name_idx': 17, 'raw_idx': 18, 'std_idx': 19}  # 专家五
        ]


        for idx in range(total_items):
            try:
                row = self.data.iloc[idx]

                # 获取标准分平均分（索引2）
                avg_standard = row[2]
                if pd.isna(avg_standard):
                    continue

                # 收集五位专家的数据
                expert_data = []

                for expert in expert_config:
                    name_idx = expert['name_idx']
                    std_idx = expert['std_idx']

                    if name_idx < len(row) and std_idx < len(row):
                        expert_name = str(row[name_idx]).strip()
                        std_score = row[std_idx]

                        # 跳过无效数据
                        if (pd.isna(expert_name) or pd.isna(std_score) or
                            expert_name == '' or expert_name == 'nan' or
                            expert_name == '姓名'):
                            continue

                        try:
                            expert_data.append({
                                'name': expert_name,
                                'std_score': float(std_score),
                                'avg_standard': float(avg_standard)
                            })
                        except:
                            continue

                # 需要至少2位专家才能比较
                if len(expert_data) < 2:
                    continue

                # 计算误差
                for expert in expert_data:
                    expert['error'] = abs(expert['std_score'] - expert['avg_standard'])

                # 找出最接近的专家（可能有多个）
                errors = [expert['error'] for expert in expert_data]
                min_error = min(errors)
                min_indices = [i for i, err in enumerate(errors) if err == min_error]

                # 计算每位专家得分
                for i, expert in enumerate(expert_data):
                    error = expert['error']

                    if i in min_indices:
                        score = 3  # 最接近
                    elif error <= 8:
                        score = 2
                    elif error > 15:
                        score = 0
                    else:  # 8 < error <= 15
                        score = mid_range_score

                    # 累加得分
                    expert_name = expert['name']
                    self.expert_scores[expert_name]['total_score'] += score
                    self.expert_scores[expert_name]['review_count'] += 1
                    self.expert_scores[expert_name]['counts'][score] += 1
                    self.expert_scores[expert_name]['details'].append({
                        'item_id': row[1] if 1 < len(row) and not pd.isna(row[1]) else f"作品_{idx+1}",
                        'avg_standard': expert['avg_standard'],
                        'std_score': expert['std_score'],
                        'error': expert['error'],
                        'score': score
                    })

                processed += 1
                valid_items += 1

                # 显示进度
                if processed % 1000 == 0:
                    print(f"  已处理 {processed:,}/{total_items:,} 条作品")

            except Exception as e:
                processed += 1
                continue

        print(f"\n✅ 计算完成！")
        print(f"   总作品数: {total_items:,}")
        print(f"   有效作品: {valid_items:,}")
        print(f"   涉及专家: {len(self.expert_scores):,} 位")

        if len(self.expert_scores) == 0:
            print("⚠️  警告：没有找到任何专家数据！")
            print("请检查数据格式，确保专家姓名和原始分列正确")

        return True

    def show_results(self, show_all=True, top_n=100):
        """显示结果"""
        if not self.expert_scores:
            print("❌ 没有计算结果")
            return

        # 转换为列表
        results = []
        for name, data in self.expert_scores.items():
            if data['review_count'] > 0:
                avg_score = data['total_score'] / data['review_count']
                efficiency = avg_score / 3 * 100  # 得分效率
                counts = data['counts']  # 提取次数
                results.append({
                    'name': name,
                    'total': data['total_score'],
                    'count': data['review_count'],
                    'c3': counts[3],  # 3分次数
                    'c2': counts[2],  # 2分次数
                    'c1': counts[1],  # 1分次数
                    'c0': counts[0],  # 0分次数
                    'avg': avg_score,
                    'efficiency': efficiency
                })

        if not results:
            print("❌ 没有有效结果")
            return

        # 按总得分排序
        results.sort(key=lambda x: x['total'], reverse=True)

        display_count = len(results) if show_all else min(top_n, len(results))

        print(f"{'=' * 110}")
        print(
            f"{'排名':<6} {'专家姓名':<12} {'总得分':<8} {'评审数':<8} {'3分':<5} {'2分':<5} {'1分':<5} {'0分':<5} {'平均分':<8} {'效率(%)':<10}")
        print(f"{'-' * 110}")

        for i, expert in enumerate(results[:display_count], 1):
            print(f"{i:<6} {expert['name']:<12} {expert['total']:<8} "
                  f"{expert['count']:<10} {expert['c3']:<6} {expert['c2']:<6} "
                  f"{expert['c1']:<6} {expert['c0']:<6} {expert['avg']:<8.2f} {expert['efficiency']:<9.1f}%")

        # 统计信息
        if results:
            print(f"\n📈 统计摘要:")
            print(f"  🏆 最高分: {results[0]['total']}分 ({results[0]['name']})")
            print(f"  📊 平均总得分: {np.mean([e['total'] for e in results]):.1f}分")
            print(f"  📉 最低分: {results[-1]['total']}分 ({results[-1]['name']})")
            print(f"  👑 评审最多: {max(results, key=lambda x: x['count'])['name']} "
                  f"({max(results, key=lambda x: x['count'])['count']}次)")

            # 前10名
            print(f"\n🏆 TOP 10 专家:")
            for i, expert in enumerate(results[:10], 1):
                print(f"  {i:2d}. {expert['name']:<12} {expert['total']:>6}分 "
                      f"(评审{expert['count']:>4}次, 效率{expert['efficiency']:.1f}%)")

    def export_results(self, filename="专家评分结果_细化统计版.xlsx"):
        """导出结果到Excel - 包含3/2/1/0分细化统计"""
        try:
            results = []
            for name, data in self.expert_scores.items():
                if data['review_count'] > 0:
                    avg_score = data['total_score'] / data['review_count']
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
                        '平均分': round(avg_score, 3),
                        '得分效率(%)': round(avg_score / 3 * 100, 2)
                    })

            results.sort(key=lambda x: x['总得分'], reverse=True)
            for i, item in enumerate(results, 1):
                item['排名'] = i

            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                pd.DataFrame(results).to_excel(writer, sheet_name='专家排名', index=False)

            print(f"\n✅ 细化结果已导出至: {filename}")
            return True
        except Exception as e:
            print(f"❌ 导出失败: {e}")
            return False

def main():
    """主程序"""
    print("=" * 80)
    print("专家评分计算系统 - 最终正确版（使用列索引）")
    print("=" * 80)

    system = CorrectScoringSystem()

    # 文件路径
    default_file = "评审结果.xls"
    filename = input(f"请输入文件名（默认:{default_file}）: ").strip() or default_file
    filepath = os.path.join(os.getcwd(), filename)

    if not os.path.exists(filepath):
        print(f"❌ 文件不存在: {filepath}")
        return

    # 加载数据
    if not system.load_data(filepath):
        input("\n按 Enter 退出...")
        return

    # 设置评分规则
    print("\n⚙️  设置评分规则:")
    mid_score = input("误差在8-15分之间给多少分? (默认1): ").strip()
    try:
        mid_score = int(mid_score) if mid_score else 1
    except:
        mid_score = 1
        print(f"使用默认值: {mid_score}")

    # 计算得分
    if not system.calculate_scores(mid_range_score=mid_score):
        input("\n按 Enter 退出...")
        return

    # 显示选项
    print("\n📊 显示选项:")
    if len(system.expert_scores) > 100:
        choice = input("专家较多，显示所有专家(A)还是显示前N名(T)? (A/T, 默认T): ").strip().upper()
        if choice == 'A':
            system.show_results(show_all=True)
        else:
            top_n = input("显示前多少名? (默认100): ").strip()
            top_n = int(top_n) if top_n else 100
            system.show_results(show_all=False, top_n=top_n)
    else:
        system.show_results(show_all=True)

    # 导出选项
    print("\n💾 导出选项:")
    export = input("是否导出到Excel文件? (Y/N, 默认Y): ").strip().upper()
    if export != 'N':
        system.export_results()

    print(f"\n{'='*80}")
    print("🎉 程序执行完成！")
    print(f"{'='*80}")

    input("\n按 Enter 退出...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"\n❌ 程序执行出错: {e}")
        import traceback
        traceback.print_exc()
        input("\n按 Enter 退出...")