#!/usr/bin/env python3
"""
CBcaseSum 三层核查机制
目的：防止误判漏判商业贿赂案例
"""

import os
import re
from docx import Document
import json


class CaseVerifier:
    """案例多层核查工具"""

    # 第一层：确定的商业贿赂关键词
    BRIBERY_KEYWORDS = ['贿赂', '行贿', '商业贿赂']

    # 第二层：需要进一步审查的可疑关键词
    SUSPICIOUS_KEYWORDS = {
        '交易相对方': '可能涉及贿赂交易相对方',
        '利用职权': '可能涉及贿赂利用职权人员',
        '影响力': '可能涉及贿赂有影响力的人',
        '好处费': '可能的贿赂形式',
        '报酬': '可能的贿赂形式',
        '回扣': '可能的贿赂形式',
        '代理': '可能涉及代理人贿赂',
        '业务费': '可能的贿赂形式',
    }

    # 第三层：非商业贿赂的排除关键词
    EXCLUDE_KEYWORDS = {
        '地理标志': '地理标志侵权，非贿赂',
        '商标侵权': '商标侵权，非贿赂',
        '虚假宣传': '虚假宣传，非贿赂',
        '虚假广告': '虚假广告，非贿赂',
        '混淆': '商业混淆行为，非贿赂',
        '价格欺诈': '价格欺诈，非贿赂',
        '不实宣传': '不实宣传，非贿赂',
    }

    def __init__(self, input_dir):
        self.input_dir = input_dir
        self.cases = []

    def verify_all(self):
        """三层验证所有案例"""
        docx_files = sorted([f for f in os.listdir(self.input_dir) if f.endswith('.docx')])

        print("="*100)
        print("【第一层核查】确定的商业贿赂案例（包含明确的贿赂关键词）")
        print("="*100)

        tier1 = []
        tier2 = []
        tier3 = []

        for idx, filename in enumerate(docx_files, 1):
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

                # 获取关键字段
                case_title = table_data.get('案件名称', filename)
                law_basis = table_data.get('处罚依据', '')
                violation_type = table_data.get('违法行为类型', '')
                violation_fact = table_data.get('违法事实', '')

                # 组合所有文本
                all_text = f"{case_title}\n{law_basis}\n{violation_type}\n{violation_fact}"

                # 第一层：确定的商业贿赂
                tier1_keywords = [kw for kw in self.BRIBERY_KEYWORDS if kw in all_text]

                if tier1_keywords:
                    case_info = {
                        'num': idx,
                        'filename': filename,
                        'title': case_title[:60],
                        'keywords': tier1_keywords,
                        'law_basis': law_basis[:100],
                        'tier': 1
                    }
                    tier1.append(case_info)
                    print(f"\n✓ {idx}. {filename[:60]}")
                    print(f"   关键词: {', '.join(tier1_keywords)}")
                    print(f"   标题: {case_title[:60]}")
                    continue

                # 第二层：可疑关键词
                tier2_keywords = {}
                for keyword, desc in self.SUSPICIOUS_KEYWORDS.items():
                    if keyword in all_text:
                        tier2_keywords[keyword] = desc

                if tier2_keywords:
                    # 检查是否被排除
                    is_excluded = any(exc in all_text for exc in self.EXCLUDE_KEYWORDS.keys())

                    if not is_excluded:
                        case_info = {
                            'num': idx,
                            'filename': filename,
                            'title': case_title[:60],
                            'keywords': tier2_keywords,
                            'law_basis': law_basis[:100],
                            'tier': 2
                        }
                        tier2.append(case_info)

                # 第三层：可能需要特别关注的案例
                if '反不正当竞争法' in law_basis and '7条' in law_basis:
                    # 第7条在2019和2025版本都存在，需要仔细判断
                    is_bribery = any(kw in all_text for kw in self.BRIBERY_KEYWORDS)
                    is_excluded = any(exc in all_text for exc in self.EXCLUDE_KEYWORDS.keys())

                    if not is_bribery and not is_excluded:
                        case_info = {
                            'num': idx,
                            'filename': filename,
                            'title': case_title[:60],
                            'reason': '涉及第7条但关键词不明确，需要人工判断',
                            'law_basis': law_basis[:100],
                            'tier': 3
                        }
                        tier3.append(case_info)

            except Exception as e:
                print(f"\n✗ {idx}. {filename[:60]} - 读取失败: {str(e)[:50]}")

        # 打印第二层
        if tier2:
            print("\n" + "="*100)
            print("【第二层核查】可疑案例（包含相关关键词，需要人工审查）")
            print("="*100)
            for case in tier2:
                print(f"\n⚠  {case['num']}. {case['filename'][:60]}")
                print(f"    标题: {case['title']}")
                print(f"    可疑关键词: {', '.join(case['keywords'].keys())}")
                print(f"    依据: {case['law_basis']}")

        # 打印第三层
        if tier3:
            print("\n" + "="*100)
            print("【第三层核查】需要特别关注的案例（第7条但无明确关键词）")
            print("="*100)
            for case in tier3:
                print(f"\n⓵  {case['num']}. {case['filename'][:60]}")
                print(f"    标题: {case['title']}")
                print(f"    依据: {case['law_basis']}")

        # 统计
        print("\n" + "="*100)
        print("【核查统计】")
        print("="*100)
        print(f"✓ 第一层（确定商业贿赂）: {len(tier1)} 个")
        print(f"⚠  第二层（可疑需审查）: {len(tier2)} 个")
        print(f"⓵  第三层（特别关注）: {len(tier3)} 个")
        print(f"ℹ  其他无关案例: {len(docx_files) - len(tier1) - len(tier2) - len(tier3)} 个")
        print(f"\n总计: {len(docx_files)} 个\n")

        return {
            'tier1': tier1,
            'tier2': tier2,
            'tier3': tier3
        }


def main():
    input_dir = os.path.expanduser("~/Desktop/商业贿赂docx")

    if not os.path.isdir(input_dir):
        print(f"错误：目录不存在: {input_dir}")
        return

    verifier = CaseVerifier(input_dir)
    result = verifier.verify_all()

    # 生成待审查清单
    if result['tier2'] or result['tier3']:
        print("\n" + "="*100)
        print("【建议】")
        print("="*100)
        print("请手动审查以下案例，判断是否属于商业贿赂：\n")

        for case in result['tier2']:
            print(f"  - {case['filename']}")
        for case in result['tier3']:
            print(f"  - {case['filename']}")


if __name__ == '__main__':
    main()
