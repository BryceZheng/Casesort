#!/usr/bin/env python3
"""
CBcaseSum 交互式启动脚本
功能：提问用户输入数据库文件夹和输出文件位置，然后执行案例提取

使用方法：
    python3 run.py
"""

import os
import sys
import subprocess
from pathlib import Path


def get_user_input(prompt, default=None):
    """获取用户输入，支持默认值"""
    if default:
        display_prompt = f"{prompt}\n  (按Enter使用默认值: {default}): "
    else:
        display_prompt = f"{prompt}: "

    user_input = input(display_prompt).strip()

    if not user_input and default:
        return default
    elif not user_input:
        print("❌ 输入不能为空，请重新输入")
        return get_user_input(prompt, default)
    else:
        return user_input


def validate_input_dir(input_dir):
    """验证输入目录是否存在且包含 .docx 文件"""
    if not os.path.isdir(input_dir):
        print(f"❌ 错误：目录不存在 - {input_dir}")
        return False

    docx_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
    if not docx_files:
        print(f"⚠️  警告：该目录中没有找到 .docx 文件")
        print(f"   目录：{input_dir}")
        return False

    print(f"✓ 找到 {len(docx_files)} 个 .docx 文件")
    return True


def validate_output_dir(output_file):
    """验证输出文件的目录是否存在"""
    output_dir = os.path.dirname(output_file)

    if not output_dir:
        output_dir = "."

    if not os.path.isdir(output_dir):
        print(f"❌ 错误：输出目录不存在 - {output_dir}")
        return False

    return True


def run_extraction(input_dir, output_file):
    """运行案例提取脚本"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    extract_script = os.path.join(script_dir, 'extract_cases.py')

    print("\n" + "="*80)
    print("开始处理案例...")
    print("="*80 + "\n")

    try:
        result = subprocess.run(
            [sys.executable, extract_script, input_dir, output_file],
            check=True
        )
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 处理失败: {str(e)}")
        return False
    except Exception as e:
        print(f"❌ 执行错误: {str(e)}")
        return False


def run_verification(input_dir):
    """运行三层核查脚本"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    verify_script = os.path.join(script_dir, 'verify_cases.py')

    print("\n" + "="*80)
    print("运行三层核查...")
    print("="*80 + "\n")

    try:
        result = subprocess.run(
            [sys.executable, verify_script],
            check=False
        )
        return True
    except Exception as e:
        print(f"❌ 执行错误: {str(e)}")
        return False


def main():
    print("="*80)
    print("【CBcaseSum v1.1】商业贿赂案例库整理工具")
    print("="*80 + "\n")

    # 获取默认值
    default_input_dir = os.path.expanduser("~/Desktop/商业贿赂docx")
    default_output_file = os.path.expanduser("~/Desktop/商业贿赂案例库2026年版.xlsx")

    # 第一步：选择操作模式
    print("请选择要执行的操作:\n")
    print("1. 提取案例并生成 Excel (推荐)")
    print("2. 先运行三层核查，再提取案例")
    print("3. 仅运行三层核查，不提取案例")
    print("4. 自定义运行")

    mode_choice = input("\n请输入选择 (1-4，默认1): ").strip() or "1"

    if mode_choice not in ["1", "2", "3", "4"]:
        print("❌ 输入错误，请输入 1-4 之间的数字")
        return

    # 第二步：获取输入目录
    print("\n" + "-"*80)
    print("【步骤 1】案例数据库文件夹")
    print("-"*80)

    input_dir = get_user_input(
        "请输入案例数据库文件夹位置",
        default=default_input_dir
    )

    # 验证输入目录
    while not validate_input_dir(input_dir):
        print("\n请重新输入：")
        input_dir = get_user_input(
            "请输入案例数据库文件夹位置",
            default=default_input_dir
        )

    # 第三步：获取输出文件位置
    if mode_choice != "3":
        print("\n" + "-"*80)
        print("【步骤 2】输出 Excel 文件位置")
        print("-"*80)

        output_file = get_user_input(
            "请输入输出 Excel 文件的位置",
            default=default_output_file
        )

        # 验证输出目录
        while not validate_output_dir(output_file):
            print("\n请重新输入：")
            output_file = get_user_input(
                "请输入输出 Excel 文件的位置",
                default=default_output_file
            )

        # 检查是否覆盖现有文件
        if os.path.exists(output_file):
            print(f"\n⚠️  警告：输出文件已存在")
            print(f"   文件：{output_file}")
            overwrite = input("\n是否覆盖？(y/n，默认 y): ").strip().lower() or "y"
            if overwrite != "y":
                print("❌ 取消操作")
                return

    # 第四步：显示确认信息
    print("\n" + "="*80)
    print("【配置确认】")
    print("="*80)
    print(f"输入目录: {input_dir}")

    if mode_choice != "3":
        print(f"输出文件: {output_file}")

    confirm = input("\n是否开始处理？(y/n，默认 y): ").strip().lower() or "y"

    if confirm != "y":
        print("❌ 取消操作")
        return

    # 第五步：执行操作
    if mode_choice == "2":
        # 先运行核查
        print("\n【第一步】运行三层核查...")
        run_verification(input_dir)

        confirm_after_verify = input("\n核查完成，是否继续提取案例？(y/n，默认 y): ").strip().lower() or "y"
        if confirm_after_verify != "y":
            print("✅ 已取消后续操作")
            return

        print("\n【第二步】提取案例...")
        run_extraction(input_dir, output_file)

    elif mode_choice == "3":
        # 仅核查
        run_verification(input_dir)

    else:
        # 直接提取或自定义
        if mode_choice == "4":
            print("\n【自定义模式】")
            print("1. 提取案例")
            print("2. 运行核查")
            print("3. 先核查后提取")

            custom_choice = input("\n请选择 (1-3，默认1): ").strip() or "1"

            if custom_choice == "1":
                run_extraction(input_dir, output_file)
            elif custom_choice == "2":
                run_verification(input_dir)
            elif custom_choice == "3":
                run_verification(input_dir)
                confirm = input("\n继续提取案例？(y/n，默认 y): ").strip().lower() or "y"
                if confirm == "y":
                    run_extraction(input_dir, output_file)
        else:
            # 模式 1：直接提取
            run_extraction(input_dir, output_file)

    print("\n" + "="*80)
    print("✅ 处理完成！")
    print("="*80 + "\n")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ 用户中断操作")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n❌ 发生错误: {str(e)}")
        sys.exit(1)
