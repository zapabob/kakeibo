#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
VSCode Code Runnerテスト用スクリプト
このファイルはCode Runnerの動作確認用です。
Python 3.10-3.12対応
"""

import os
import sys
import locale

# システムのロケール確認 (Python 3.10-3.12対応)
try:
    if sys.version_info >= (3, 11):
        print(f"現在のロケール: {locale.getlocale()}")
        print(f"現在のエンコーディング: {locale.getencoding()}")
    else:
        # 3.10以前 (警告は出るが動作する)
        print(f"現在のロケール: {locale.getdefaultlocale()}")
except Exception as e:
    print(f"ロケール取得エラー: {e}")

print(f"現在の文字コード: {sys.getdefaultencoding()}")
print(f"Pythonバージョン: {sys.version}")
print(f"実行パス: {os.path.abspath(__file__)}")
print(f"カレントディレクトリ: {os.getcwd()}")

print("\n日本語出力テスト：こんにちは世界！")

# 家計簿アプリの起動方法を表示
print("\n家計簿アプリの起動方法:")
print("1. run_kakeibo.bat をダブルクリック")
print("2. コマンドプロンプトで 'py -3 kakeibo.py' を実行")
print("3. VSCodeのCode Runnerで kakeibo.py を実行")

# Python 3.10-3.12の互換性情報
print("\n互換性情報:")
print(f"Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
if sys.version_info >= (3, 10) and sys.version_info < (3, 13):
    print("✓ 対応バージョンです (Python 3.10-3.12)")
else:
    print("! 非対応バージョンです (Python 3.10-3.12が推奨)") 