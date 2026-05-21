#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Schedule Generator - Main Entry Point
整合AI排班生成和Excel转换功能
"""

import os
import sys
import argparse


def run_ai_schedule():
    """运行AI排班生成"""
    try:
        from call_ai import main as call_ai_main
        print("=== 开始AI排班生成 ===")
        call_ai_main()
        print("=== AI排班生成完成 ===")
        return True
    except Exception as e:
        print(f"AI排班生成失败: {e}")
        return False


def run_json_to_xlsx():
    """运行JSON到Excel转换"""
    try:
        from json_to_xlsx import main as json_to_xlsx_main
        print("=== 开始JSON到Excel转换 ===")
        json_to_xlsx_main()
        print("=== JSON到Excel转换完成 ===")
        return True
    except Exception as e:
        print(f"JSON到Excel转换失败: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description='排班生成工具 - 整合AI生成和Excel转换')
    parser.add_argument('--mode', choices=['ai', 'xlsx', 'both'], default='both',
                       help='运行模式: ai(仅AI生成), xlsx(仅Excel转换), both(完整流程)')

    args = parser.parse_args()

    if args.mode == 'ai':
        success = run_ai_schedule()
        sys.exit(0 if success else 1)
    elif args.mode == 'xlsx':
        success = run_json_to_xlsx()
        sys.exit(0 if success else 1)
    else:  # both
        print("=== 排班生成完整流程 ===")

        # 第一步：AI生成
        ai_success = run_ai_schedule()
        if not ai_success:
            print("❌ AI生成失败，终止流程")
            sys.exit(1)

        # 第二步：Excel转换
        xlsx_success = run_json_to_xlsx()
        if not xlsx_success:
            print("❌ Excel转换失败")
            sys.exit(1)

        print("=== ✅ 完整流程执行成功 ===")
        sys.exit(0)


if __name__ == "__main__":
    main()