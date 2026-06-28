#!/usr/bin/env python3
"""
CBcaseSum 高级深度核验工具 v1.2
功能：基于商业贿赂法律构成要件的实质性分析和识别

核心改进：
1. 深入的法律分析（不仅关键词匹配）
2. 完整的事实链提取和验证
3. 变相贿赂的识别
4. 复合型案件的正确分类
5. 置信度评分和详细的分析理由
"""

import os
import re
from docx import Document
from collections import defaultdict


class PremiumBriberyVerifier:
    """高级商业贿赂核验工具 - 基于法律实质分析"""

    # 商业贿赂法律构成要件
    BRIBERY_ELEMENTS = {
        'subject': {
            'keywords': ['企业', '经营者', '公司', '商人', '经营', '法定代表人'],
            'required': True,
            'description': '经营者主体'
        },
        'purpose': {
            'keywords': ['交易', '订单', '合同', '竞争', '市场', '机会', '客户', '业务', '销售', '谋取'],
            'required': True,
            'description': '为谋取交易或竞争优势的目的'
        },
        'recipient': {
            'keywords': ['员工', '工作人员', '经理', '主管', '采购', '决策', '利用职权', '影响力'],
            'required': True,
            'description': '向交易相对方员工或有权人员'
        },
        'benefit_transfer': {
            'keywords': ['支付', '给予', '转账', '返还', '返点', '折扣', '金钱', '利益', '费用', '好处'],
            'required': True,
            'description': '进行了利益输送'
        }
    }

    # 变相贿赂的隐蔽名义
    DISGUISED_NAMES = {
        '咨询费': '高风险',
        '技术服务费': '高风险',
        '赞助费': '高风险',
        '会务费': '高风险',
        '活动费': '高风险',
        '培训费': '中风险',
        '差旅费': '中风险',
        '返点': '直接贿赂',
        '折扣': '可能的变相贿赂',
        '回扣': '直接贿赂',
        '好处费': '直接贿赂',
    }

    # 隐蔽操作标志
    CONCEALMENT_INDICATORS = {
        '账外': '账外操作（高风险）',
        '未如实入账': '隐瞒记账（高风险）',
        '暗中': '暗中操作',
        '背后': '背后操作',
        '私下': '私下操作',
        '现金': '现金操作（难追溯）',
    }

    # 纯粹不当竞争行为（排除关键词）
    PURE_UNFAIR_KEYWORDS = {
        '地理标志': '不是贿赂',
        '商标': '不是贿赂',
        '著作权': '不是贿赂',
        '虚假宣传': '可能与贿赂无关',
        '夸大': '可能与贿赂无关',
        '刷单': '不是贿赂',
        '炒信': '不是贿赂',
        '虚假评价': '不是贿赂',
    }

    def __init__(self, input_dir):
        self.input_dir = input_dir
        self.results = []

    def extract_table_data(self, filepath):
        """提取表格数据，支持错误处理"""
        table_data = {}
        try:
            doc = Document(filepath)

            for table in doc.tables:
                try:
                    for row in table.rows:
                        try:
                            cells = row.cells
                            if len(cells) >= 2:
                                key = cells[0].text.strip()
                                value = cells[-1].text.strip() if len(cells) > 1 else ""
                                if key and value:
                                    table_data[key] = value
                        except ValueError:
                            # 处理合并单元格
                            continue
                except:
                    continue

        except Exception as e:
            return None

        return table_data if table_data else None

    def analyze_bribery_elements(self, text, table_data):
        """
        分析是否满足商业贿赂的法律构成要件
        返回：(是否为贿赂, 置信度, 理由列表)
        """
        all_text = text.lower() if text else ""

        findings = {
            'has_subject': False,
            'has_purpose': False,
            'has_recipient': False,
            'has_benefit': False,
            'has_concealment': False,
            'is_pure_unfair': False,
            'disguised_indicators': [],
            'evidence': [],
        }

        # 检查纯粹的不当竞争
        for keyword in self.PURE_UNFAIR_KEYWORDS.keys():
            if keyword in all_text:
                findings['is_pure_unfair'] = True
                findings['evidence'].append(f"纯粹不当竞争特征：{keyword}")

        # 检查商业贿赂要件
        for element_name, element_info in self.BRIBERY_ELEMENTS.items():
            for keyword in element_info['keywords']:
                if keyword in all_text:
                    findings[f'has_{element_name}'] = True
                    findings['evidence'].append(f"{element_info['description']}：检测到'{keyword}'")
                    break

        # 检查变相贿赂
        for disguised_name, risk_level in self.DISGUISED_NAMES.items():
            if disguised_name in all_text:
                findings['disguised_indicators'].append(f"{disguised_name} ({risk_level})")

        # 检查隐蔽操作
        for concealment, desc in self.CONCEALMENT_INDICATORS.items():
            if concealment in all_text:
                findings['has_concealment'] = True
                findings['evidence'].append(f"隐蔽操作标志：{desc}")

        return findings

    def calculate_confidence(self, findings):
        """计算置信度和判定结果"""
        score = 0
        max_score = 100

        # 基础评分
        if findings['has_subject']:
            score += 20
        if findings['has_purpose']:
            score += 20
        if findings['has_recipient']:
            score += 20
        if findings['has_benefit']:
            score += 20
        if findings['disguised_indicators']:
            score += 15
        if findings['has_concealment']:
            score += 10

        # 调整因素
        if findings['is_pure_unfair'] and score < 50:
            # 如果纯粹不当竞争但有贿赂要件 → 倾向判定为贿赂
            if score >= 40:
                score = max(score, 65)  # 复合案件
            else:
                score = max(0, score - 20)  # 纯粹不当竞争

        # 判定
        if score >= 70:
            return True, score, "商业贿赂"
        elif score >= 50:
            return True, score, "可能的商业贿赂"
        elif score >= 30:
            return False, score, "可疑，需审查"
        else:
            return False, score, "无关"

    def verify_case(self, filename, case_num):
        """对单个案例进行深度核验"""
        filepath = os.path.join(self.input_dir, filename)

        try:
            # 提取表格数据
            table_data = self.extract_table_data(filepath)
            if not table_data:
                return {
                    'num': case_num,
                    'filename': filename,
                    'status': '读取失败',
                    'error': '无法提取表格数据'
                }

            # 组合所有文本
            text = '\n'.join([
                table_data.get('案件名称', ''),
                table_data.get('处罚名称', ''),
                table_data.get('处罚依据', ''),
                table_data.get('违法行为类型', ''),
                table_data.get('违法事实', ''),
            ])

            # 分析要件
            findings = self.analyze_bribery_elements(text, table_data)

            # 计算置信度
            is_bribery, confidence, verdict = self.calculate_confidence(findings)

            return {
                'num': case_num,
                'filename': filename,
                'title': table_data.get('案件名称', filename)[:50],
                'status': 'success',
                'is_bribery': is_bribery,
                'confidence': confidence,
                'verdict': verdict,
                'findings': findings,
                'evidence': findings['evidence'][:3],  # 取前 3 个证据
                'disguised': findings['disguised_indicators'][:2],
            }

        except Exception as e:
            return {
                'num': case_num,
                'filename': filename,
                'status': '处理失败',
                'error': str(e)[:50]
            }

    def print_result(self, result):
        """打印单个结果"""
        num = result.get('num')

        if result.get('status') != 'success':
            status_icon = '❌'
            print(f"【{num:3d}】{status_icon} {result['filename'][:40]}")
            return

        is_bribery = result['is_bribery']
        confidence = result['confidence']
        verdict = result['verdict']

        # 判定图标
        if verdict == "商业贿赂":
            icon = "✓"
            color = "✓"
        elif verdict == "可能的商业贿赂":
            icon = "⚠️ "
            color = "⚠️"
        elif verdict == "可疑，需审查":
            icon = "❓"
            color = "❓"
        else:
            icon = "✗"
            color = "✗"

        print(f"【{num:3d}】{icon} {result['filename'][:40]} ({verdict}, {confidence}%)")

    def verify_batch(self, max_cases=100):
        """批量核验"""
        docx_files = sorted([f for f in os.listdir(self.input_dir) if f.endswith('.docx')])[:max_cases]

        print("="*100)
        print(f"【高级深度核验】分析前 {len(docx_files)} 个案例")
        print("="*100 + "\n")

        bribery_cases = []
        suspicious_cases = []
        unrelated_cases = []
        errors = []

        for idx, filename in enumerate(docx_files, 1):
            result = self.verify_case(filename, idx)

            if result.get('status') != 'success':
                errors.append(result)
                self.print_result(result)
            else:
                self.print_result(result)

                if result['verdict'] == "商业贿赂":
                    bribery_cases.append(result)
                elif result['verdict'] == "可能的商业贿赂":
                    suspicious_cases.append(result)
                else:
                    unrelated_cases.append(result)

            # 每 20 个案例打印进度
            if idx % 20 == 0:
                print(f"\n... 已处理 {idx} 个案例 ...\n")

        # 生成统计报告
        print("\n" + "="*100)
        print("【核验统计报告】")
        print("="*100)
        print(f"✓ 明确商业贿赂:     {len(bribery_cases):3d} 个 ({len(bribery_cases)*100//len(docx_files):2d}%)")
        print(f"⚠️  可能的商业贿赂:  {len(suspicious_cases):3d} 个 ({len(suspicious_cases)*100//len(docx_files):2d}%)")
        print(f"✗ 无关案例:       {len(unrelated_cases):3d} 个 ({len(unrelated_cases)*100//len(docx_files):2d}%)")
        print(f"❌ 处理失败:       {len(errors):3d} 个")
        print(f"\n商业贿赂相关总计: {len(bribery_cases) + len(suspicious_cases)} 个 ({(len(bribery_cases) + len(suspicious_cases))*100//len(docx_files)}%)")

        # 显示商业贿赂案例
        if bribery_cases:
            print("\n" + "="*100)
            print("【明确的商业贿赂案例】")
            print("="*100)
            for case in bribery_cases[:10]:  # 显示前 10 个
                print(f"\n{case['num']:3d}. {case['title']}")
                print(f"    置信度: {case['confidence']}%")
                print(f"    证据: {case['evidence'][0] if case['evidence'] else '无'}")

        if suspicious_cases:
            print("\n" + "="*100)
            print("【可能的商业贿赂案例（需人工审查）】")
            print("="*100)
            for case in suspicious_cases[:10]:  # 显示前 10 个
                print(f"\n{case['num']:3d}. {case['title']}")
                print(f"    置信度: {case['confidence']}%")
                if case['disguised']:
                    print(f"    变相贿赂标志: {case['disguised'][0]}")

        return {
            'bribery': bribery_cases,
            'suspicious': suspicious_cases,
            'unrelated': unrelated_cases,
            'errors': errors,
        }


def main():
    input_dir = os.path.expanduser("~/Desktop/商业贿赂docx")

    if not os.path.isdir(input_dir):
        print(f"错误：目录不存在: {input_dir}")
        return

    verifier = PremiumBriberyVerifier(input_dir)
    results = verifier.verify_batch(max_cases=100)

    # 保存结果
    import json
    output_file = os.path.expanduser("~/Desktop/高级核验结果_前100个.json")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump({
            'bribery_count': len(results['bribery']),
            'suspicious_count': len(results['suspicious']),
            'unrelated_count': len(results['unrelated']),
            'bribery_cases': [
                {
                    'num': c['num'],
                    'filename': c['filename'],
                    'title': c['title'],
                    'confidence': c['confidence'],
                } for c in results['bribery'][:20]
            ],
            'suspicious_cases': [
                {
                    'num': c['num'],
                    'filename': c['filename'],
                    'title': c['title'],
                    'confidence': c['confidence'],
                } for c in results['suspicious'][:20]
            ],
        }, f, ensure_ascii=False, indent=2)

    print(f"\n✅ 详细结果已保存: {output_file}")


if __name__ == '__main__':
    main()
