#!/usr/bin/env python3
"""
CBcaseSum 深度核验工具 - 基于构成要件的实质分析
功能：不依赖关键词匹配，而是深入分析案件事实，核验是否满足商业贿赂的实质要件

商业贿赂实质要件：
1. 行为人主体：经营者（企业或自然人）
2. 行为目的：为谋取交易机会或竞争优势
3. 行为对象：特定关系人（交易相对方员工、受托人、有职权影响力的人）
4. 行为方式：利益输送（直接或变相）
5. 因果关系：为了获取交易而输送利益

关键识别点：
- 谁（行为人）
- 以什么名义（支付的理由/费用名称）
- 向谁支付（接收方）
- 支付了什么利益（金额、礼品、折扣等）
- 为了什么目的（获取交易、竞争优势等）
"""

import os
import re
import sys
from docx import Document


class DeepVerifier:
    """商业贿赂深度实质核验工具"""

    # 商业贿赂明确关键词
    CLEAR_BRIBERY_KEYWORDS = ['贿赂', '行贿', '商业贿赂']

    # 变相贿赂的假名义（需要进一步分析）
    DISGUISED_BRIBERY_INDICATORS = {
        '咨询费': '可能的变相贿赂标志',
        '技术服务费': '可能的变相贿赂标志',
        '赞助费': '可能的变相贿赂标志',
        '会务费': '可能的变相贿赂标志',
        '活动费': '可能的变相贿赂标志',
        '协议返点': '典型的变相贿赂',
        '折扣': '可能的变相贿赂（需结合上下文）',
        '账外': '典型的隐藏行为标志',
        '未如实入账': '典型的隐藏行为标志',
        '暗中': '隐蔽性指示词',
        '背后': '隐蔽性指示词',
        '回扣': '变相贿赂标志',
        '提成': '可能的贿赂标志（需结合对象）',
    }

    # 交易相对方关键词（识别受贿对象）
    RECIPIENT_KEYWORDS = [
        '销售员', '业务员', '工作人员',
        '经理', '主管', '决策人',
        '采购', '招标', '采购员',
        '供应商', '代理商', '经销商',
        '监管部门工作人员', '检验员',
    ]

    # 目的/动机关键词（识别为何贿赂）
    MOTIVE_KEYWORDS = [
        '获得', '取得', '谋取',
        '竞争', '市场', '订单',
        '合同', '交易', '机会',
        '通过', '便利', '照顾',
        '优先', '感谢', '补偿',
    ]

    # 非商业贿赂的排除标志（纯粹的不当竞争行为）
    PURE_UNFAIR_COMPETITION = {
        '地理标志': '地理标志侵权',
        '商标': '商标侵权',
        '著作权': '著作权侵权',
        '商业秘密': '侵犯商业秘密',  # 但如果涉及贿赂获取秘密则不排除
        '虚假宣传': '虚假宣传',  # 但如果涉及利益输送则有关
        '夸大': '夸大宣传',
        '刷单': '刷单行为',
        '炒信': '炒信行为',
        '虚假评价': '虚假评价',
        '价格欺诈': '价格欺诈',
    }

    def __init__(self, input_dir):
        self.input_dir = input_dir

    def extract_fact_chain(self, case_text):
        """
        提取事实链：谁、以什么名义、向谁支付了什么利益
        返回：{who, what_name, to_whom, what_benefit, why}
        """
        fact_chain = {
            '行为人': None,
            '支付名义': None,
            '受益人': None,
            '利益内容': None,
            '目的': None,
        }

        # 提取行为人（通常是案件主体）
        person_patterns = [
            r'（当事人）(.+?)（的|）',
            r'当事人(.+?)（为了|为了|向|给予）',
            r'当事人为(.+?)向',
        ]
        for pattern in person_patterns:
            match = re.search(pattern, case_text)
            if match:
                fact_chain['行为人'] = match.group(1)
                break

        # 提取支付名义
        for indicator, desc in self.DISGUISED_BRIBERY_INDICATORS.items():
            if indicator in case_text:
                fact_chain['支付名义'] = indicator
                break

        # 提取受益人（通常是相对方的员工或有权人）
        for keyword in self.RECIPIENT_KEYWORDS:
            if keyword in case_text:
                # 尝试找到具体名字
                context_match = re.search(f'({keyword})(.+?)(?:的|给|向|共|共计)', case_text)
                if context_match:
                    fact_chain['受益人'] = f"{keyword}{context_match.group(2)}"
                else:
                    fact_chain['受益人'] = keyword
                break

        # 提取利益内容（金额、礼品等）
        amount_pattern = r'([0-9,，]+)元|共计(.+?)元'
        amount_match = re.search(amount_pattern, case_text)
        if amount_match:
            fact_chain['利益内容'] = f"{amount_match.group(1) or amount_match.group(2)}元"

        # 提取目的
        for keyword in self.MOTIVE_KEYWORDS:
            if keyword in case_text:
                context_match = re.search(f'(为了|为|以|通过)({keyword}.+?)(?:，|。)', case_text)
                if context_match:
                    fact_chain['目的'] = context_match.group(2)
                else:
                    fact_chain['目的'] = keyword
                break

        return fact_chain

    def analyze_context(self, case_data):
        """
        基于上下文的实质核验
        """
        case_title = case_data.get('案件名称', '')
        law_basis = case_data.get('处罚依据', '')
        violation_type = case_data.get('违法行为类型', '')
        violation_fact = case_data.get('违法事实', '')

        all_text = f"{case_title}\n{law_basis}\n{violation_type}\n{violation_fact}"

        analysis = {
            'is_bribery': False,
            'confidence': 0,  # 0-100
            'reasons': [],
            'evidence': {},
            'fact_chain': {},
            'warnings': [],
        }

        # 第零步：识别2025修订版虚假宣传案件（避免因第七条误判为商业贿赂）
        # 2019修正：第七条=商业贿赂，第八条=虚假宣传，第十九条=商业贿赂处罚
        # 2025修订：第七条=虚假宣传，第八条=商业贿赂，第二十三条=虚假宣传处罚，第二十四条=商业贿赂处罚
        # 若法律依据为"第七条+第二十三条"（2025修订版虚假宣传组合），且无明确贿赂关键词 → 非商业贿赂
        if (re.search(r'第七条', all_text) and
                re.search(r'第二十三条', all_text) and
                not re.search(r'第八条|第二十四条', all_text) and
                not any(kw in all_text for kw in self.CLEAR_BRIBERY_KEYWORDS)):
            analysis['is_bribery'] = False
            analysis['confidence'] = 90
            analysis['reasons'].append('✗ 2025修订版第七条+第二十三条=虚假宣传处罚组合，非商业贿赂')
            return analysis

        # 第一步：检查明确的商业贿赂关键词
        if any(kw in all_text for kw in self.CLEAR_BRIBERY_KEYWORDS):
            analysis['is_bribery'] = True
            analysis['confidence'] = 95
            analysis['reasons'].append('✓ 发现明确的商业贿赂关键词')
            return analysis

        # 第二步：检查是否是纯粹的非商业贿赂行为
        pure_unfair_count = 0
        has_bribery_indicators = False

        for unfair_keyword in self.PURE_UNFAIR_COMPETITION.keys():
            if unfair_keyword in all_text:
                pure_unfair_count += 1
                analysis['evidence'][unfair_keyword] = self.PURE_UNFAIR_COMPETITION[unfair_keyword]

        # 检查是否有贿赂指标
        for bribery_indicator in self.DISGUISED_BRIBERY_INDICATORS.keys():
            if bribery_indicator in all_text:
                has_bribery_indicators = True
                analysis['evidence'][bribery_indicator] = self.DISGUISED_BRIBERY_INDICATORS[bribery_indicator]

        # 第三步：如果既有贿赂指标又有不当竞争指标，深入分析
        if pure_unfair_count > 0 and has_bribery_indicators:
            # 复合型案件
            analysis['reasons'].append('⚠️  复合型案件：同时存在贿赂指标和不当竞争行为')

            # 提取事实链进行深入分析
            fact_chain = self.extract_fact_chain(violation_fact)
            analysis['fact_chain'] = fact_chain

            # 检查是否存在利益输送事实
            if fact_chain.get('受益人') or fact_chain.get('利益内容'):
                analysis['is_bribery'] = True
                analysis['confidence'] = 75
                analysis['reasons'].append('✓ 发现利益输送事实，判定为贿赂相关')
                return analysis
            else:
                analysis['is_bribery'] = False
                analysis['confidence'] = 60
                analysis['reasons'].append('✗ 虽有混合指标，但无明确利益输送事实')
                return analysis

        # 第四步：仅有不当竞争指标，无贿赂指标
        if pure_unfair_count > 0 and not has_bribery_indicators:
            analysis['is_bribery'] = False
            analysis['confidence'] = 85
            analysis['reasons'].append('✗ 纯粹的不当竞争行为，无贿赂指标')
            return analysis

        # 第五步：仅有贿赂指标（变相贿赂）
        if has_bribery_indicators and pure_unfair_count == 0:
            fact_chain = self.extract_fact_chain(violation_fact)
            analysis['fact_chain'] = fact_chain

            # 检查是否满足贿赂要件
            if self._check_bribery_elements(violation_fact, fact_chain):
                analysis['is_bribery'] = True
                analysis['confidence'] = 80
                analysis['reasons'].append('✓ 发现变相贿赂特征，满足贿赂要件')
                return analysis

        # 第六步：都没有明确指标
        analysis['is_bribery'] = False
        analysis['confidence'] = 50
        analysis['reasons'].append('✗ 无明确商业贿赂指标')
        analysis['warnings'].append('⓵ 建议人工审查')

        return analysis

    def _check_bribery_elements(self, violation_fact, fact_chain):
        """
        检查是否满足商业贿赂的实质要件
        """
        # 要件1：存在目的（为谋取交易或竞争优势）
        has_purpose = fact_chain.get('目的') is not None
        or_pattern = r'(为了|为了|以|为|通过).{2,20}(交易|竞争|订单|合同|机会|客户)'
        has_purpose = has_purpose or bool(re.search(or_pattern, violation_fact))

        # 要件2：存在受益人（交易相对方或有权人）
        has_recipient = fact_chain.get('受益人') is not None

        # 要件3：存在利益输送
        has_benefit = fact_chain.get('利益内容') is not None
        or_benefit = bool(re.search(r'(支付|给予|转账|返还|返点).{0,10}([0-9,，]+元|现金|礼品)', violation_fact))
        has_benefit = has_benefit or or_benefit

        # 至少满足 2 个要件就认为可能是贿赂
        elements_met = sum([has_purpose, has_recipient, has_benefit])
        return elements_met >= 2

    def verify_case(self, filename):
        """
        对单个案例进行深度核验
        """
        filepath = os.path.join(self.input_dir, filename)

        try:
            doc = Document(filepath)

            # 收集表格数据
            table_data = {}
            for table in doc.tables:
                for row in table.rows:
                    cells = row.cells
                    if len(cells) >= 2:
                        key = cells[0].text.strip()
                        value = cells[-1].text.strip() if len(cells) > 1 else ""
                        if key and value:
                            table_data[key] = value

            # 进行实质核验
            analysis = self.analyze_context(table_data)

            return {
                'filename': filename,
                'title': table_data.get('案件名称', filename),
                'analysis': analysis,
            }

        except Exception as e:
            return {
                'filename': filename,
                'error': str(e),
            }

    def print_report(self, result):
        """打印核验报告"""
        if 'error' in result:
            print(f"❌ {result['filename']}: {result['error']}")
            return

        filename = result['filename']
        title = result['title']
        analysis = result['analysis']

        # 判定结果
        verdict = '✓ 商业贿赂相关' if analysis['is_bribery'] else '✗ 无关'
        confidence = f"置信度: {analysis['confidence']}%"

        print(f"\n【{filename}】")
        print(f"标题: {title[:50]}")
        print(f"判定: {verdict} ({confidence})")

        # 判定理由
        print(f"理由:")
        for reason in analysis['reasons']:
            print(f"  {reason}")

        # 证据
        if analysis['evidence']:
            print(f"检出特征:")
            for key, desc in analysis['evidence'].items():
                print(f"  • {key}: {desc}")

        # 事实链
        if analysis['fact_chain'] and any(analysis['fact_chain'].values()):
            print(f"事实链:")
            chain = analysis['fact_chain']
            print(f"  谁: {chain.get('行为人', 'N/A')}")
            print(f"  以什么名义: {chain.get('支付名义', 'N/A')}")
            print(f"  向谁: {chain.get('受益人', 'N/A')}")
            print(f"  支付了什么: {chain.get('利益内容', 'N/A')}")
            print(f"  为了什么: {chain.get('目的', 'N/A')}")

        # 警告
        if analysis['warnings']:
            print(f"⚠️  警告:")
            for warning in analysis['warnings']:
                print(f"  {warning}")


def main():
    input_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser("~/Desktop/商业贿赂docx")

    if not os.path.isdir(input_dir):
        print(f"错误：目录不存在 - {input_dir}")
        return

    verifier = DeepVerifier(input_dir)

    print("="*100)
    print("【深度核验】基于构成要件的商业贿赂实质分析")
    print("="*100)

    docx_files = sorted([f for f in os.listdir(input_dir) if f.endswith('.docx')])

    bribery_cases = []
    unrelated_cases = []

    for filename in docx_files:
        result = verifier.verify_case(filename)
        verifier.print_report(result)

        if 'analysis' in result:
            if result['analysis']['is_bribery']:
                bribery_cases.append(result)
            else:
                unrelated_cases.append(result)

    # 总结
    print("\n" + "="*100)
    print("【核验总结】")
    print("="*100)
    print(f"✓ 商业贿赂相关: {len(bribery_cases)} 个")
    print(f"✗ 无关案例: {len(unrelated_cases)} 个")
    print(f"总计: {len(docx_files)} 个\n")


if __name__ == '__main__':
    main()
