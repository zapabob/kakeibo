#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
家計簿日報システム (仮称)
========================
このシステムは、以下の機能を実装しているよ：
1. 日々の収支入力 (日付、区分、科目、金額)
2. SQLiteを利用したDBへのデータ保存（パラメータ化クエリでSQLインジェクション対策済み）
3. 月次集計機能（pandasを使った収入・支出・差引計算）
4. GUI (tkinter)による直感的な操作画面
5. リマインダー機能 (scheduleライブラリとスレッドで毎日特定時刻に通知)
6. バックアップ機能 (DBファイルのコピー)
7. エラーハンドリングとログ出力 (loggingモジュール利用)
8. CSVエクスポート機能 (新規追加)
"""

import sqlite3
import logging
import schedule
import threading
import time
import datetime
import pandas as pd
import re
import os
import shutil
import sys
import traceback
import locale
import chardet
from dateutil import parser
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QTableWidget, QTableWidgetItem, QTabWidget,
    QFileDialog, QMessageBox, QSpinBox, QDialog, QListWidget, QDialogButtonBox
)

# システムのロケールを設定（日本語Windows環境向け）
try:
    # Python 3.11以降では推奨される方法
    if sys.version_info >= (3, 11):
        locale.setlocale(locale.LC_ALL, '')
        sys_encoding = locale.getencoding()
        sys_locale = locale.getlocale()
    else:
        # 3.10以前の方法
        locale.setlocale(locale.LC_ALL, 'ja_JP.UTF-8')
except locale.Error:
    try:
        locale.setlocale(locale.LC_ALL, 'Japanese_Japan.932')
    except locale.Error:
        # どのロケールも設定できない場合はデフォルトを使用
        pass

# Python 3.12対応のためにioencoding環境変数を設定
os.environ["PYTHONIOENCODING"] = "utf-8"

# 必要なモジュールの確認とインストール案内
try:
    import pandas as pd
except ImportError:
    print("pandas モジュールがインストールされていません。")
    print("インストール方法: pip install pandas")
    sys.exit(1)

try:
    import schedule
    import threading
except ImportError:
    print("schedule モジュールがインストールされていません。")
    print("インストール方法: pip install schedule")
    sys.exit(1)

try:
    import chardet
    from dateutil import parser
except ImportError:
    print("chardet または python-dateutil モジュールがインストールされていません。")
    print("インストール方法: pip install chardet python-dateutil")
    sys.exit(1)

try:
    from PyQt6.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
        QComboBox, QPushButton, QTableWidget, QTableWidgetItem, QTabWidget,
        QFileDialog, QMessageBox
    )
except ImportError:
    print("PyQt6 モジュールがインストールされていません。")
    print("インストール方法: pip install PyQt6")
    sys.exit(1)

# -------------------------------------------------------------------
# 1. ログ設定
# -------------------------------------------------------------------
try:
    log_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(log_dir, "kakeibo.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filename=log_file,
        filemode="a",
        encoding="utf-8"
    )
    logging.info("家計簿日報システム 起動")
except Exception as e:
    print(f"ログ設定エラー: {e}")
    traceback.print_exc()

# データベースファイル名
if getattr(sys, 'frozen', False):
    # 実行可能ファイルとして実行されている場合
    app_dir = os.path.dirname(sys.executable)
    DATABASE_NAME = os.path.join(app_dir, "kakeibo.db")
else:
    # スクリプトとして実行されている場合
    app_dir = os.path.dirname(os.path.abspath(__file__))
    DATABASE_NAME = os.path.join(app_dir, "kakeibo.db")

# データベースディレクトリの確認と作成
os.makedirs(os.path.dirname(DATABASE_NAME), exist_ok=True)

# -------------------------------------------------------------------
# 2. データベース初期化
# -------------------------------------------------------------------
def initialize_database():
    """
    SQLite DBの初期化。テーブルが存在しなければ新規作成する。
    テーブル: kakeibo
    カラム: id (PK, AUTOINCREMENT), date (TEXT), category (TEXT), subject (TEXT), amount (REAL)
    """
    conn = None
    try:
        # データベースディレクトリの確認
        db_dir = os.path.dirname(DATABASE_NAME)
        if not os.path.exists(db_dir) and db_dir:
            os.makedirs(db_dir, exist_ok=True)
            
        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS kakeibo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                category TEXT NOT NULL,
                subject TEXT NOT NULL,
                amount REAL NOT NULL
            )
        """)
        conn.commit()
        logging.info(f"データベース初期化完了: {DATABASE_NAME}")
    except sqlite3.Error as e:
        logging.error("DB初期化エラー: %s", e)
        QMessageBox.critical(None, "エラー", "データベースの初期化に失敗しました。")
    finally:
        if conn:
            conn.close()

# -------------------------------------------------------------------
# 3. 入力バリデーション
# -------------------------------------------------------------------
def convert_date_format(date_str):
    """
    様々な日付形式をYYYY-MM-DDに変換
    Args:
        date_str (str): 変換対象の日付文字列
    Returns:
        str: YYYY-MM-DD形式の日付文字列
    """
    if not date_str or pd.isna(date_str):
        return None
    
    date_str = str(date_str).strip()
    
    # 既にYYYY-MM-DD形式の場合はそのまま返す
    if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        return date_str
    
    try:
        # 令和形式の変換
        if '令和' in date_str:
            # 令和yy年mm月dd日 → YYYY-MM-DD
            match = re.match(r'令和(\d{1,2})年(\d{1,2})月(\d{1,2})日', date_str)
            if match:
                year = int(match.group(1)) + 2018  # 令和元年は2019年
                month = int(match.group(2))
                day = int(match.group(3))
                return f"{year:04d}-{month:02d}-{day:02d}"
        
        # Ryy/mm/dd形式の変換
        if date_str.startswith('R'):
            match = re.match(r'R(\d{1,2})/(\d{1,2})/(\d{1,2})', date_str)
            if match:
                year = int(match.group(1)) + 2018
                month = int(match.group(2))
                day = int(match.group(3))
                return f"{year:04d}-{month:02d}-{day:02d}"
        
        # yyyy年mm月dd日形式の変換
        if '年' in date_str and '月' in date_str and '日' in date_str:
            match = re.match(r'(\d{4})年(\d{1,2})月(\d{1,2})日', date_str)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
                day = int(match.group(3))
                return f"{year:04d}-{month:02d}-{day:02d}"
        
        # yyyy/mm/dd形式の変換
        if '/' in date_str:
            match = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})', date_str)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
                day = int(match.group(3))
                return f"{year:04d}-{month:02d}-{day:02d}"
        
        # dateutilを使用した汎用的な変換
        parsed_date = parser.parse(date_str, dayfirst=False, yearfirst=True)
        return parsed_date.strftime('%Y-%m-%d')
        
    except Exception as e:
        logging.warning(f"日付形式変換エラー: {date_str}, エラー: {e}")
        return None

def validate_input(date_str, amount_str):
    """
    ユーザー入力のバリデーション
    - 日付はYYYY-MM-DD形式であるか
    - 金額は数値変換できるか
    """
    # 日付形式変換を試行
    converted_date = convert_date_format(date_str)
    if converted_date is None:
        QMessageBox.warning(None, "入力エラー", "日付形式が正しくありません。\n対応形式: YYYY-MM-DD, yyyy/mm/dd, yyyy年mm月dd日, 令和yy年mm月dd日, Ryy/mm/dd")
        return False
    
    try:
        float(amount_str)
    except ValueError:
        QMessageBox.warning(None, "入力エラー", "金額は数値で入力してください。")
        return False
    return True

# -------------------------------------------------------------------
# 4. DBへのデータ登録
# -------------------------------------------------------------------
def insert_record(date, category, subject, amount):
    """
    DBへ新規レコードを挿入する関数。
    SQLインジェクション防止のため、パラメータ化クエリを利用。
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        query = "INSERT INTO kakeibo (date, category, subject, amount) VALUES (?, ?, ?, ?)"
        cursor.execute(query, (date, category, subject, amount))
        conn.commit()
        logging.info("新規レコード追加: 日付=%s, 区分=%s, 科目=%s, 金額=%s", date, category, subject, amount)
        return True
    except sqlite3.Error as e:
        logging.error("DB挿入エラー: %s", e)
        QMessageBox.critical(None, "エラー", "データの保存中にエラーが発生しました。詳細はログを確認してください。")
        return False
    finally:
        if conn:
            conn.close()

# -------------------------------------------------------------------
# 5. 月次集計機能 (pandas利用)
# -------------------------------------------------------------------
def fetch_monthly_summary(month):
    """
    指定された月(YYYY-MM形式)のデータを取得し、収入合計、支出合計、差引残高を計算する。
    pandasを利用してデータ集計を行う。
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        df = pd.read_sql_query("SELECT * FROM kakeibo WHERE substr(date, 1, 7) = ?", conn, params=(month,))
        if df.empty:
            QMessageBox.information(None, "集計結果", "指定された月のデータはありません。")
            return None
        income_total = df[df['category'] == '収入']['amount'].sum()
        expense_total = df[df['category'] == '支出']['amount'].sum()
        balance = income_total - expense_total
        summary_text = (
            f"【{month}の集計結果】\n"
            f"収入合計: {income_total}\n"
            f"支出合計: {expense_total}\n"
            f"差引残高: {balance}"
        )
        # QMessageBox.information(None, "月次集計", summary_text)
        logging.info("月次集計実施: %s", summary_text)
        return summary_text
    except Exception as e:
        logging.error("集計エラー: %s", e)
        QMessageBox.critical(None, "エラー", "集計処理中にエラーが発生しました。詳細はログを確認してください。")
        return None
    finally:
        if conn:
            conn.close()

# -------------------------------------------------------------------
# 6. バックアップ機能
# -------------------------------------------------------------------
def backup_database():
    """
    DBファイルのバックアップを行う。
    現在の日時を付与したファイル名でコピーを作成する。
    """
    backup_filename = f"kakeibo_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    try:
        shutil.copy(DATABASE_NAME, backup_filename)
        logging.info("バックアップ成功: %s", backup_filename)
        QMessageBox.information(None, "バックアップ", f"バックアップが成功しました。\nファイル名: {backup_filename}")
    except Exception as e:
        logging.error("バックアップ失敗: %s", e)
        QMessageBox.critical(None, "エラー", "バックアップ中にエラーが発生しました。")

# -------------------------------------------------------------------
# 7. リマインダー機能 (schedule と threading 利用)
# -------------------------------------------------------------------
def reminder_job():
    """リマインダーのジョブ関数。指定時刻になるとユーザーに家計簿入力を促す。"""
    logging.info("リマインダー実行: 今日の家計簿入力は済んでるかな？")
    QMessageBox.information(None, "リマインダー", "今日の家計簿入力は済んでるかな？")

def run_scheduler():
    """scheduleライブラリを利用して、定期的にジョブを実行するための無限ループ関数。"""
    while True:
        schedule.run_pending()
        time.sleep(1)

def start_scheduler_thread():
    """スレッドを利用して、GUIと並行してリマインダー処理をバックグラウンドで実行する。"""
    schedule.every().day.at("20:00").do(reminder_job)
    scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
    scheduler_thread.start()
    logging.info("リマインダースレッド開始")

# -------------------------------------------------------------------
# 8. CSV機能 (エクスポート・インポート)
# -------------------------------------------------------------------

def detect_encoding(filename):
    """
    CSVファイルのエンコーディングを自動検出
    Args:
        filename (str): ファイルパス
    Returns:
        str: 検出されたエンコーディング
    """
    try:
        with open(filename, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']
            confidence = result['confidence']
            logging.info(f"エンコーディング検出: {encoding} (信頼度: {confidence:.2f})")
            return encoding if confidence > 0.7 else 'utf-8'
    except Exception as e:
        logging.warning(f"エンコーディング検出エラー: {e}")
        return 'utf-8'

def validate_csv_format(df):
    """
    CSVファイルの形式を検証
    Args:
        df (pandas.DataFrame): 検証対象のDataFrame
    Returns:
        dict: 検証結果
    """
    result = {
        'valid': True,
        'errors': [],
        'warnings': []
    }
    
    # 必須列の確認
    required_columns = ['date', 'category', 'subject', 'amount']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        result['valid'] = False
        result['errors'].append(f"必要な列が不足しています: {missing_columns}")
    
    # データ型の確認
    if 'amount' in df.columns:
        try:
            df['amount'] = pd.to_numeric(df['amount'], errors='coerce')
            if df['amount'].isna().any():
                result['warnings'].append("金額列に数値以外のデータが含まれています")
        except Exception as e:
            result['errors'].append(f"金額列の変換エラー: {e}")
    
    # 日付形式の確認
    if 'date' in df.columns:
        invalid_dates = []
        for idx, date_str in enumerate(df['date']):
            if convert_date_format(date_str) is None:
                invalid_dates.append(f"行{idx+1}: {date_str}")
        
        if invalid_dates:
            result['warnings'].append(f"変換できない日付形式: {invalid_dates[:5]}...")
    
    return result

def export_to_csv(filename):
    """
    DBの全データをCSV形式でエクスポートする
    pandasを使用してデータを取得し、CSVファイルに出力
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        df = pd.read_sql_query("SELECT * FROM kakeibo", conn)
        if df.empty:
            QMessageBox.warning(None, "エラー", "エクスポートするデータがありません")
            return
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        QMessageBox.information(None, "成功", f"CSVファイルをエクスポートしました\n{filename}")
        logging.info("CSVエクスポート成功: %s", filename)
    except Exception as e:
        logging.error("CSVエクスポートエラー: %s", e)
        QMessageBox.critical(None, "エラー", "CSVのエクスポート中にエラーが発生しました")
    finally:
        if conn:
            conn.close()

def import_from_csv(filename):
    """
    CSVファイルからデータをインポートする
    Args:
        filename (str): インポートするCSVファイルのパス
    Returns:
        bool: インポート成功時True
    """
    try:
        # 1. エンコーディング検出
        encoding = detect_encoding(filename)
        logging.info(f"CSVインポート開始: {filename}, エンコーディング: {encoding}")
        
        # 2. ファイル読み込み
        df = pd.read_csv(filename, encoding=encoding)
        logging.info(f"CSV読み込み完了: {len(df)}行")
        
        # 3. 形式検証
        validation_result = validate_csv_format(df)
        if not validation_result['valid']:
            error_msg = "\n".join(validation_result['errors'])
            QMessageBox.critical(None, "CSV形式エラー", f"CSVファイルの形式が正しくありません:\n{error_msg}")
            return False
        
        if validation_result['warnings']:
            warning_msg = "\n".join(validation_result['warnings'])
            QMessageBox.warning(None, "CSV形式警告", f"警告:\n{warning_msg}")
        
        # 4. データ変換
        success_count = 0
        error_count = 0
        error_details = []
        
        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        
        for idx, row in df.iterrows():
            try:
                # 日付形式変換
                date_str = convert_date_format(row['date'])
                if date_str is None:
                    error_details.append(f"行{idx+1}: 日付形式エラー - {row['date']}")
                    error_count += 1
                    continue
                
                # 金額変換
                try:
                    amount = float(row['amount'])
                except (ValueError, TypeError):
                    error_details.append(f"行{idx+1}: 金額形式エラー - {row['amount']}")
                    error_count += 1
                    continue
                
                # カテゴリと科目の取得
                category = str(row['category']).strip()
                subject = str(row['subject']).strip()
                
                # DBに挿入
                query = "INSERT INTO kakeibo (date, category, subject, amount) VALUES (?, ?, ?, ?)"
                cursor.execute(query, (date_str, category, subject, amount))
                success_count += 1
                
            except Exception as e:
                error_details.append(f"行{idx+1}: {e}")
                error_count += 1
        
        conn.commit()
        conn.close()
        
        # 5. 結果表示
        result_msg = f"インポート完了\n成功: {success_count}件\nエラー: {error_count}件"
        if error_details:
            result_msg += "\n\nエラー詳細:\n" + "\n".join(error_details[:10])
            if len(error_details) > 10:
                result_msg += f"\n... 他{len(error_details)-10}件"
        
        if error_count == 0:
            QMessageBox.information(None, "インポート成功", result_msg)
        else:
            QMessageBox.warning(None, "インポート完了（一部エラー）", result_msg)
        
        logging.info(f"CSVインポート完了: 成功{success_count}件, エラー{error_count}件")
        return True
        
    except Exception as e:
        logging.error(f"CSVインポートエラー: {e}")
        QMessageBox.critical(None, "エラー", f"CSVのインポート中にエラーが発生しました:\n{e}")
        return False

# -------------------------------------------------------------------
# 月ごとのデータベース管理機能
# -------------------------------------------------------------------

def get_monthly_data(month):
    """
    指定された月のデータを取得する
    Args:
        month (str): YYYY-MM形式の月
    Returns:
        pandas.DataFrame: 月別データ
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        query = "SELECT * FROM kakeibo WHERE substr(date, 1, 7) = ? ORDER BY date"
        df = pd.read_sql_query(query, conn, params=(month,))
        logging.info(f"月別データ取得: {month}, 件数: {len(df)}")
        return df
    except Exception as e:
        logging.error(f"月別データ取得エラー: {e}")
        return pd.DataFrame()
    finally:
        if conn:
            conn.close()

def create_monthly_backup(month):
    """
    指定された月のデータを別ファイルにバックアップする
    Args:
        month (str): YYYY-MM形式の月
    """
    try:
        # 月別データを取得
        monthly_data = get_monthly_data(month)
        if monthly_data.empty:
            QMessageBox.warning(None, "警告", f"{month}のデータがありません")
            return False
        
        # バックアップファイル名
        backup_filename = f"kakeibo_monthly_{month.replace('-', '')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        # CSVとして保存
        monthly_data.to_csv(backup_filename, index=False, encoding='utf-8-sig')
        
        logging.info(f"月別バックアップ成功: {backup_filename}")
        QMessageBox.information(None, "月別バックアップ", f"{month}のデータをバックアップしました\n{backup_filename}")
        return True
    except Exception as e:
        logging.error(f"月別バックアップエラー: {e}")
        QMessageBox.critical(None, "エラー", f"月別バックアップ中にエラーが発生しました: {e}")
        return False

def get_monthly_statistics(month):
    """
    指定された月の詳細統計を取得する
    Args:
        month (str): YYYY-MM形式の月
    Returns:
        dict: 統計情報
    """
    try:
        monthly_data = get_monthly_data(month)
        if monthly_data.empty:
            return None
        
        # 基本統計
        income_data = monthly_data[monthly_data['category'] == '収入']
        expense_data = monthly_data[monthly_data['category'] == '支出']
        
        stats = {
            'month': month,
            'total_records': len(monthly_data),
            'income_count': len(income_data),
            'expense_count': len(expense_data),
            'income_total': income_data['amount'].sum(),
            'expense_total': expense_data['amount'].sum(),
            'balance': income_data['amount'].sum() - expense_data['amount'].sum(),
            'avg_income': income_data['amount'].mean() if len(income_data) > 0 else 0,
            'avg_expense': expense_data['amount'].mean() if len(expense_data) > 0 else 0,
            'max_income': income_data['amount'].max() if len(income_data) > 0 else 0,
            'max_expense': expense_data['amount'].max() if len(expense_data) > 0 else 0,
            'min_income': income_data['amount'].min() if len(income_data) > 0 else 0,
            'min_expense': expense_data['amount'].min() if len(expense_data) > 0 else 0
        }
        
        # 科目別統計
        subject_stats = monthly_data.groupby(['category', 'subject'])['amount'].sum().reset_index()
        stats['subject_breakdown'] = subject_stats.to_dict('records')
        
        logging.info(f"月別統計取得: {month}")
        return stats
    except Exception as e:
        logging.error(f"月別統計取得エラー: {e}")
        return None

def export_monthly_summary_csv(month, filename):
    """
    指定された月の詳細集計をCSVファイルに出力する
    Args:
        month (str): YYYY-MM形式の月
        filename (str): 出力ファイル名
    Returns:
        bool: 出力成功時True
    """
    try:
        # 月別データ取得
        monthly_data = get_monthly_data(month)
        if monthly_data.empty:
            QMessageBox.warning(None, "エラー", f"{month}のデータがありません")
            return False
        
        # 統計情報取得
        stats = get_monthly_statistics(month)
        if stats is None:
            QMessageBox.warning(None, "エラー", f"{month}の統計情報を取得できませんでした")
            return False
        
        # CSV用のデータフレーム作成
        summary_data = []
        
        # 基本統計情報
        summary_data.append({
            '項目': '対象月',
            '値': stats['month'],
            '備考': ''
        })
        summary_data.append({
            '項目': '総レコード数',
            '値': stats['total_records'],
            '備考': '件'
        })
        summary_data.append({
            '項目': '収入総額',
            '値': f"{stats['income_total']:,.0f}",
            '備考': '円'
        })
        summary_data.append({
            '項目': '支出総額',
            '値': f"{stats['expense_total']:,.0f}",
            '備考': '円'
        })
        summary_data.append({
            '項目': '差引残高',
            '値': f"{stats['balance']:,.0f}",
            '備考': '円'
        })
        summary_data.append({
            '項目': '収入件数',
            '値': stats['income_count'],
            '備考': '件'
        })
        summary_data.append({
            '項目': '支出件数',
            '値': stats['expense_count'],
            '備考': '件'
        })
        summary_data.append({
            '項目': '平均収入',
            '値': f"{stats['avg_income']:,.0f}",
            '備考': '円'
        })
        summary_data.append({
            '項目': '平均支出',
            '値': f"{stats['avg_expense']:,.0f}",
            '備考': '円'
        })
        summary_data.append({
            '項目': '最大収入',
            '値': f"{stats['max_income']:,.0f}",
            '備考': '円'
        })
        summary_data.append({
            '項目': '最大支出',
            '値': f"{stats['max_expense']:,.0f}",
            '備考': '円'
        })
        
        # 基本統計をDataFrameに変換
        summary_df = pd.DataFrame(summary_data)
        
        # 科目別集計データ
        subject_df = pd.DataFrame(stats['subject_breakdown'])
        subject_df['amount_formatted'] = subject_df['amount'].apply(lambda x: f"{x:,.0f}")
        subject_df = subject_df.rename(columns={
            'category': '区分',
            'subject': '科目',
            'amount': '金額',
            'amount_formatted': '金額（表示用）'
        })
        
        # ExcelWriterを使用して複数シートのCSVファイルを作成
        # 基本統計用のCSVファイル名
        base_filename = filename.replace('.csv', '')
        summary_filename = f"{base_filename}_基本統計.csv"
        subject_filename = f"{base_filename}_科目別集計.csv"
        detail_filename = f"{base_filename}_詳細データ.csv"
        
        # 基本統計CSV出力
        summary_df.to_csv(summary_filename, index=False, encoding='utf-8-sig')
        
        # 科目別集計CSV出力
        subject_df.to_csv(subject_filename, index=False, encoding='utf-8-sig')
        
        # 詳細データCSV出力（元データ）
        monthly_data.to_csv(detail_filename, index=False, encoding='utf-8-sig')
        
        # 成功メッセージ
        result_msg = "月別集計CSV出力完了:\n"
        result_msg += f"・基本統計: {summary_filename}\n"
        result_msg += f"・科目別集計: {subject_filename}\n"
        result_msg += f"・詳細データ: {detail_filename}"
        
        QMessageBox.information(None, "CSV出力成功", result_msg)
        logging.info(f"月別集計CSV出力成功: {month}")
        return True
        
    except Exception as e:
        logging.error(f"月別集計CSV出力エラー: {e}")
        QMessageBox.critical(None, "エラー", f"月別集計CSV出力中にエラーが発生しました:\n{e}")
        return False

def export_multiple_months_summary_csv(months_list, filename):
    """
    複数月の集計比較をCSVファイルに出力する
    Args:
        months_list (list): YYYY-MM形式の月のリスト
        filename (str): 出力ファイル名
    Returns:
        bool: 出力成功時True
    """
    try:
        comparison_data = []
        
        for month in months_list:
            stats = get_monthly_statistics(month)
            if stats:
                comparison_data.append({
                    '月': month,
                    '収入総額': stats['income_total'],
                    '支出総額': stats['expense_total'],
                    '差引残高': stats['balance'],
                    '収入件数': stats['income_count'],
                    '支出件数': stats['expense_count'],
                    '総レコード数': stats['total_records'],
                    '平均収入': stats['avg_income'],
                    '平均支出': stats['avg_expense'],
                    '最大収入': stats['max_income'],
                    '最大支出': stats['max_expense']
                })
        
        if not comparison_data:
            QMessageBox.warning(None, "エラー", "有効なデータがありません")
            return False
        
        # DataFrameに変換
        comparison_df = pd.DataFrame(comparison_data)
        
        # 金額の表示形式を整形
        for col in ['収入総額', '支出総額', '差引残高', '平均収入', '平均支出', '最大収入', '最大支出']:
            comparison_df[f'{col}（表示用）'] = comparison_df[col].apply(lambda x: f"{x:,.0f}")
        
        # CSV出力
        comparison_df.to_csv(filename, index=False, encoding='utf-8-sig')
        
        QMessageBox.information(None, "CSV出力成功", f"複数月比較CSV出力完了:\n{filename}")
        logging.info(f"複数月比較CSV出力成功: {months_list}")
        return True
        
    except Exception as e:
        logging.error(f"複数月比較CSV出力エラー: {e}")
        QMessageBox.critical(None, "エラー", f"複数月比較CSV出力中にエラーが発生しました:\n{e}")
        return False

def archive_monthly_data(month):
    """
    指定された月のデータをアーカイブテーブルに移動する
    Args:
        month (str): YYYY-MM形式の月
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        
        # アーカイブテーブルの作成
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS kakeibo_archive (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                original_id INTEGER,
                date TEXT NOT NULL,
                category TEXT NOT NULL,
                subject TEXT NOT NULL,
                amount REAL NOT NULL,
                archived_date TEXT NOT NULL
            )
        """)
        
        # 月別データをアーカイブテーブルにコピー
        cursor.execute("""
            INSERT INTO kakeibo_archive (original_id, date, category, subject, amount, archived_date)
            SELECT id, date, category, subject, amount, ?
            FROM kakeibo WHERE substr(date, 1, 7) = ?
        """, (datetime.datetime.now().strftime('%Y-%m-%d'), month))
        
        # 元のテーブルから削除
        cursor.execute("DELETE FROM kakeibo WHERE substr(date, 1, 7) = ?", (month,))
        
        conn.commit()
        archived_count = cursor.rowcount
        logging.info(f"月別アーカイブ成功: {month}, 件数: {archived_count}")
        QMessageBox.information(None, "アーカイブ", f"{month}のデータをアーカイブしました\n件数: {archived_count}")
        return True
    except Exception as e:
        logging.error(f"月別アーカイブエラー: {e}")
        QMessageBox.critical(None, "エラー", f"アーカイブ中にエラーが発生しました: {e}")
        return False
    finally:
        if conn:
            conn.close()

def restore_monthly_data(month):
    """
    アーカイブから指定された月のデータを復元する
    Args:
        month (str): YYYY-MM形式の月
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        
        # アーカイブテーブルからデータを復元
        cursor.execute("""
            INSERT INTO kakeibo (date, category, subject, amount)
            SELECT date, category, subject, amount
            FROM kakeibo_archive WHERE substr(date, 1, 7) = ?
        """, (month,))
        
        # アーカイブテーブルから削除
        cursor.execute("DELETE FROM kakeibo_archive WHERE substr(date, 1, 7) = ?", (month,))
        
        conn.commit()
        restored_count = cursor.rowcount
        logging.info(f"月別復元成功: {month}, 件数: {restored_count}")
        QMessageBox.information(None, "復元", f"{month}のデータを復元しました\n件数: {restored_count}")
        return True
    except Exception as e:
        logging.error(f"月別復元エラー: {e}")
        QMessageBox.critical(None, "エラー", f"復元中にエラーが発生しました: {e}")
        return False
    finally:
        if conn:
            conn.close()

def get_available_months():
    """
    データベースに存在する月の一覧を取得する
    Returns:
        list: 月のリスト（YYYY-MM形式）
    """
    conn = None
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        cursor = conn.cursor()
        
        # メインテーブルから月を取得
        cursor.execute("SELECT DISTINCT substr(date, 1, 7) FROM kakeibo ORDER BY substr(date, 1, 7) DESC")
        main_months = [row[0] for row in cursor.fetchall()]
        
        # アーカイブテーブルから月を取得
        cursor.execute("SELECT DISTINCT substr(date, 1, 7) FROM kakeibo_archive ORDER BY substr(date, 1, 7) DESC")
        archive_months = [row[0] for row in cursor.fetchall()]
        
        all_months = list(set(main_months + archive_months))
        all_months.sort(reverse=True)
        
        logging.info(f"利用可能月取得: {len(all_months)}ヶ月")
        return all_months
    except Exception as e:
        logging.error(f"利用可能月取得エラー: {e}")
        return []
    finally:
        if conn:
            conn.close()

def cleanup_old_data(months_to_keep=12):
    """
    古いデータを自動クリーンアップする
    Args:
        months_to_keep (int): 保持する月数
    """
    try:
        available_months = get_available_months()
        if len(available_months) <= months_to_keep:
            logging.info("クリーンアップ不要: データ数が保持月数以下")
            return True
        
        # 古い月を特定
        months_to_archive = available_months[months_to_keep:]
        
        for month in months_to_archive:
            # バックアップを作成してからアーカイブ
            if create_monthly_backup(month):
                archive_monthly_data(month)
        
        logging.info(f"クリーンアップ完了: {len(months_to_archive)}ヶ月をアーカイブ")
        return True
    except Exception as e:
        logging.error(f"クリーンアップエラー: {e}")
        return False

# -------------------------------------------------------------------
# 9. GUI アプリケーション (PyQt6)
# -------------------------------------------------------------------
class KakeiboApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.load_data()


    def initUI(self):
        self.setWindowTitle("家計簿日報システム (仮称)")
        self.setGeometry(300, 300, 600, 400)

        # タブウィジェット
        self.tabs = QTabWidget()

        # 入力タブ
        self.input_tab = QWidget()
        self.setup_input_tab()

        # 表示タブ
        self.view_tab = QWidget()
        self.setup_view_tab()

        # 月次集計タブ
        self.summary_tab = QWidget()
        self.setup_summary_tab()

        # 月別管理タブ
        self.monthly_management_tab = QWidget()
        self.setup_monthly_management_tab()

        self.tabs.addTab(self.input_tab, "入力")
        self.tabs.addTab(self.view_tab, "表示・編集")
        self.tabs.addTab(self.summary_tab, "月次集計")
        self.tabs.addTab(self.monthly_management_tab, "月別管理")


        # メインレイアウト
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)
        self.setLayout(main_layout)

    def setup_input_tab(self):
        # 入力フォームのレイアウト
        layout = QVBoxLayout()

        # 日付入力
        date_layout = QHBoxLayout()
        date_label = QLabel("日付 (YYYY-MM-DD):")
        self.date_edit = QLineEdit(datetime.date.today().strftime("%Y-%m-%d"))
        date_layout.addWidget(date_label)
        date_layout.addWidget(self.date_edit)
        layout.addLayout(date_layout)

        # 区分 (収入/支出)
        category_layout = QHBoxLayout()
        category_label = QLabel("区分:")
        self.category_combo = QComboBox()
        self.category_combo.addItems(["支出", "収入"])
        category_layout.addWidget(category_label)
        category_layout.addWidget(self.category_combo)
        layout.addLayout(category_layout)

        # 科目入力
        subject_layout = QHBoxLayout()
        subject_label = QLabel("科目:")
        self.subject_edit = QLineEdit()
        subject_layout.addWidget(subject_label)
        subject_layout.addWidget(self.subject_edit)
        layout.addLayout(subject_layout)

        # 金額入力
        amount_layout = QHBoxLayout()
        amount_label = QLabel("金額:")
        self.amount_edit = QLineEdit()
        amount_layout.addWidget(amount_label)
        amount_layout.addWidget(self.amount_edit)
        layout.addLayout(amount_layout)

        # 保存ボタン
        self.save_button = QPushButton("保存")
        self.save_button.clicked.connect(self.save_record)
        layout.addWidget(self.save_button)

        self.input_tab.setLayout(layout)

    def setup_view_tab(self):
        layout = QVBoxLayout()

        # データ表示テーブル
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(5)
        self.table_widget.setHorizontalHeaderLabels(["ID", "日付", "区分", "科目", "金額"])
        self.table_widget.itemChanged.connect(self.update_record)
        layout.addWidget(self.table_widget)

        # 更新ボタン, 削除ボタン
        hbox = QHBoxLayout()
        self.refresh_button = QPushButton("更新")
        self.refresh_button.clicked.connect(self.load_data)
        hbox.addWidget(self.refresh_button)

        self.delete_button = QPushButton("削除")
        self.delete_button.clicked.connect(self.delete_record)
        hbox.addWidget(self.delete_button)

        # CSVエクスポート
        self.export_button = QPushButton("CSVエクスポート")
        self.export_button.clicked.connect(self.export_csv)
        hbox.addWidget(self.export_button)

        # CSVインポート
        self.import_button = QPushButton("CSVインポート")
        self.import_button.clicked.connect(self.import_csv)
        hbox.addWidget(self.import_button)

        # バックアップボタン
        self.backup_button = QPushButton("バックアップ")
        self.backup_button.clicked.connect(backup_database)
        hbox.addWidget(self.backup_button)

        layout.addLayout(hbox)
        self.view_tab.setLayout(layout)

    def setup_summary_tab(self):
        layout = QVBoxLayout()

        # 月選択
        month_layout = QHBoxLayout()
        month_label = QLabel("月 (YYYY-MM):")
        self.month_edit = QLineEdit(datetime.date.today().strftime("%Y-%m"))
        month_layout.addWidget(month_label)
        month_layout.addWidget(self.month_edit)

        # 集計ボタン
        self.summary_button = QPushButton("集計")
        self.summary_button.clicked.connect(self.show_monthly_summary)
        month_layout.addWidget(self.summary_button)

        layout.addLayout(month_layout)

        # 集計結果表示
        self.summary_label = QLabel("")
        layout.addWidget(self.summary_label)

        self.summary_tab.setLayout(layout)

    def setup_monthly_management_tab(self):
        """月別管理タブのセットアップ"""
        layout = QVBoxLayout()

        # 月選択セクション
        month_selection_layout = QHBoxLayout()
        month_label = QLabel("対象月:")
        self.month_management_combo = QComboBox()
        self.refresh_month_list()
        month_selection_layout.addWidget(month_label)
        month_selection_layout.addWidget(self.month_management_combo)
        
        # 月リスト更新ボタン
        refresh_month_button = QPushButton("月リスト更新")
        refresh_month_button.clicked.connect(self.refresh_month_list)
        month_selection_layout.addWidget(refresh_month_button)
        
        layout.addLayout(month_selection_layout)

        # 月別統計表示
        stats_layout = QVBoxLayout()
        stats_label = QLabel("月別統計:")
        stats_layout.addWidget(stats_label)
        
        self.monthly_stats_text = QLabel("月を選択してください")
        self.monthly_stats_text.setStyleSheet("QLabel { background-color: #f0f0f0; padding: 10px; border: 1px solid #ccc; }")
        stats_layout.addWidget(self.monthly_stats_text)
        
        # 統計更新ボタン
        update_stats_button = QPushButton("統計更新")
        update_stats_button.clicked.connect(self.update_monthly_statistics)
        stats_layout.addWidget(update_stats_button)
        
        layout.addLayout(stats_layout)

        # 月別管理ボタン群
        management_buttons_layout = QHBoxLayout()
        
        # 月別バックアップ
        monthly_backup_button = QPushButton("月別バックアップ")
        monthly_backup_button.clicked.connect(self.create_monthly_backup_gui)
        management_buttons_layout.addWidget(monthly_backup_button)
        
        # 月別アーカイブ
        monthly_archive_button = QPushButton("月別アーカイブ")
        monthly_archive_button.clicked.connect(self.archive_monthly_data_gui)
        management_buttons_layout.addWidget(monthly_archive_button)
        
        # 月別復元
        monthly_restore_button = QPushButton("月別復元")
        monthly_restore_button.clicked.connect(self.restore_monthly_data_gui)
        management_buttons_layout.addWidget(monthly_restore_button)
        
        layout.addLayout(management_buttons_layout)

        # CSV出力ボタン群
        csv_export_layout = QHBoxLayout()
        
        # 月別集計CSV出力
        monthly_csv_button = QPushButton("月別集計CSV出力")
        monthly_csv_button.clicked.connect(self.export_monthly_csv_gui)
        csv_export_layout.addWidget(monthly_csv_button)
        
        # 複数月比較CSV出力
        multi_month_csv_button = QPushButton("複数月比較CSV出力")
        multi_month_csv_button.clicked.connect(self.export_multi_month_csv_gui)
        csv_export_layout.addWidget(multi_month_csv_button)
        
        layout.addLayout(csv_export_layout)

        # 自動クリーンアップ
        cleanup_layout = QHBoxLayout()
        cleanup_label = QLabel("保持月数:")
        self.keep_months_spin = QSpinBox()
        self.keep_months_spin.setRange(1, 60)
        self.keep_months_spin.setValue(12)
        cleanup_layout.addWidget(cleanup_label)
        cleanup_layout.addWidget(self.keep_months_spin)
        
        cleanup_button = QPushButton("自動クリーンアップ")
        cleanup_button.clicked.connect(self.cleanup_old_data_gui)
        cleanup_layout.addWidget(cleanup_button)
        
        layout.addLayout(cleanup_layout)

        self.monthly_management_tab.setLayout(layout)

    def refresh_month_list(self):
        """利用可能な月のリストを更新する"""
        try:
            available_months = get_available_months()
            self.month_management_combo.clear()
            if available_months:
                self.month_management_combo.addItems(available_months)
                self.month_management_combo.currentTextChanged.connect(self.update_monthly_statistics)
            else:
                self.month_management_combo.addItem("データがありません")
        except Exception as e:
            logging.error(f"月リスト更新エラー: {e}")
            QMessageBox.critical(self, "エラー", f"月リストの更新中にエラーが発生しました: {e}")

    def update_monthly_statistics(self):
        """選択された月の統計を更新する"""
        try:
            selected_month = self.month_management_combo.currentText()
            if not selected_month or selected_month == "データがありません":
                self.monthly_stats_text.setText("月を選択してください")
                return
            
            stats = get_monthly_statistics(selected_month)
            if stats:
                stats_text = f"""
【{stats['month']}の詳細統計】

基本情報:
• 総レコード数: {stats['total_records']}件
• 収入件数: {stats['income_count']}件
• 支出件数: {stats['expense_count']}件

収支情報:
• 収入合計: ¥{stats['income_total']:,.0f}
• 支出合計: ¥{stats['expense_total']:,.0f}
• 差引残高: ¥{stats['balance']:,.0f}

平均・最大・最小:
• 平均収入: ¥{stats['avg_income']:,.0f}
• 平均支出: ¥{stats['avg_expense']:,.0f}
• 最大収入: ¥{stats['max_income']:,.0f}
• 最大支出: ¥{stats['max_expense']:,.0f}
• 最小収入: ¥{stats['min_income']:,.0f}
• 最小支出: ¥{stats['min_expense']:,.0f}
"""
                self.monthly_stats_text.setText(stats_text)
            else:
                self.monthly_stats_text.setText(f"{selected_month}のデータがありません")
        except Exception as e:
            logging.error(f"月別統計更新エラー: {e}")
            QMessageBox.critical(self, "エラー", f"統計の更新中にエラーが発生しました: {e}")

    def create_monthly_backup_gui(self):
        """月別バックアップのGUI処理"""
        try:
            selected_month = self.month_management_combo.currentText()
            if not selected_month or selected_month == "データがありません":
                QMessageBox.warning(self, "警告", "月を選択してください")
                return
            
            create_monthly_backup(selected_month)
        except Exception as e:
            logging.error(f"月別バックアップGUIエラー: {e}")
            QMessageBox.critical(self, "エラー", f"バックアップ中にエラーが発生しました: {e}")

    def archive_monthly_data_gui(self):
        """月別アーカイブのGUI処理"""
        try:
            selected_month = self.month_management_combo.currentText()
            if not selected_month or selected_month == "データがありません":
                QMessageBox.warning(self, "警告", "月を選択してください")
                return
            
            reply = QMessageBox.question(self, "確認", 
                                       f"{selected_month}のデータをアーカイブしますか？\nこの操作は元に戻せません。",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                       QMessageBox.StandardButton.No)
            
            if reply == QMessageBox.StandardButton.Yes:
                if archive_monthly_data(selected_month):
                    self.refresh_month_list()
                    self.load_data()  # メインテーブルも更新
        except Exception as e:
            logging.error(f"月別アーカイブGUIエラー: {e}")
            QMessageBox.critical(self, "エラー", f"アーカイブ中にエラーが発生しました: {e}")

    def restore_monthly_data_gui(self):
        """月別復元のGUI処理"""
        try:
            selected_month = self.month_management_combo.currentText()
            if not selected_month or selected_month == "データがありません":
                QMessageBox.warning(self, "警告", "月を選択してください")
                return
            
            reply = QMessageBox.question(self, "確認", 
                                       f"{selected_month}のデータを復元しますか？",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                       QMessageBox.StandardButton.No)
            
            if reply == QMessageBox.StandardButton.Yes:
                if restore_monthly_data(selected_month):
                    self.refresh_month_list()
                    self.load_data()  # メインテーブルも更新
        except Exception as e:
            logging.error(f"月別復元GUIエラー: {e}")
            QMessageBox.critical(self, "エラー", f"復元中にエラーが発生しました: {e}")

    def cleanup_old_data_gui(self):
        """自動クリーンアップのGUI処理"""
        try:
            months_to_keep = self.keep_months_spin.value()
            reply = QMessageBox.question(self, "確認", 
                                       f"古いデータを自動クリーンアップしますか？\n保持月数: {months_to_keep}ヶ月",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                       QMessageBox.StandardButton.No)
            
            if reply == QMessageBox.StandardButton.Yes:
                if cleanup_old_data(months_to_keep):
                    self.refresh_month_list()
                    self.load_data()  # メインテーブルも更新
                    QMessageBox.information(self, "完了", "自動クリーンアップが完了しました")
        except Exception as e:
            logging.error(f"自動クリーンアップGUIエラー: {e}")
            QMessageBox.critical(self, "エラー", f"クリーンアップ中にエラーが発生しました: {e}")

    def load_data(self):
        """DBからデータを読み込み、テーブルに表示する"""
        conn = None
        try:
            conn = sqlite3.connect(DATABASE_NAME)
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM kakeibo")
            data = cursor.fetchall()

            self.table_widget.setRowCount(0)  # テーブルをクリア
            for row_number, row_data in enumerate(data):
                self.table_widget.insertRow(row_number)
                for column_number, cell_data in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_data))
                    self.table_widget.setItem(row_number, column_number, item)
        except sqlite3.Error as e:
            logging.error("データ読み込みエラー: %s", e)
            QMessageBox.critical(self, "エラー", "データの読み込み中にエラーが発生しました。")
        finally:
            if conn:
                conn.close()

    def save_record(self):
        """入力データのバリデーションとDB保存"""
        date = self.date_edit.text()
        category = self.category_combo.currentText()
        subject = self.subject_edit.text()
        amount_str = self.amount_edit.text()

        # 日付形式変換
        converted_date = convert_date_format(date)
        if converted_date is None:
            QMessageBox.warning(self, "入力エラー", "日付形式が正しくありません。\n対応形式: YYYY-MM-DD, yyyy/mm/dd, yyyy年mm月dd日, 令和yy年mm月dd日, Ryy/mm/dd")
            return

        if not validate_input(converted_date, amount_str):
            return

        amount = float(amount_str)
        if insert_record(converted_date, category, subject, amount):
            QMessageBox.information(self, "成功", "データが保存されました！")
            self.subject_edit.clear()
            self.amount_edit.clear()
            self.load_data()  # データの再読み込み
        else:
            QMessageBox.critical(self, "エラー", "データの保存に失敗しました")

    def update_record(self, item):
        """テーブルのセルが変更されたときにDBのレコードを更新する"""
        row = item.row()
        column = item.column()
        new_value = item.text()
        record_id = self.table_widget.item(row, 0).text()  # IDを取得

        column_names = ["id", "date", "category", "subject", "amount"]
        column_name = column_names[column]

        conn = None
        try:
            conn = sqlite3.connect(DATABASE_NAME)
            cursor = conn.cursor()
            query = f"UPDATE kakeibo SET {column_name} = ? WHERE id = ?"
            cursor.execute(query, (new_value, record_id))
            conn.commit()
            logging.info(f"レコード更新: ID={record_id}, 列={column_name}, 値={new_value}")

            # 金額が変更された場合、数値チェックを行う
            if column_name == "amount":
                if not validate_input(self.table_widget.item(row,1).text(), new_value):
                    self.load_data() # データ再読み込みで変更をもとに戻す
                    return
        except sqlite3.Error as e:
            logging.error("レコード更新エラー: %s", e)
            QMessageBox.critical(self, "エラー", "データの更新中にエラーが発生しました。")
        finally:
            if conn:
                conn.close()

    def delete_record(self):
        """選択された行のレコードをDBから削除する"""
        selected_row = self.table_widget.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "エラー", "削除する行を選択してください。")
            return

        record_id = self.table_widget.item(selected_row, 0).text()

        reply = QMessageBox.question(self, "確認", "本当にこのレコードを削除しますか？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            conn = None
            try:
                conn = sqlite3.connect(DATABASE_NAME)
                cursor = conn.cursor()
                query = "DELETE FROM kakeibo WHERE id = ?"
                cursor.execute(query, (record_id,))
                conn.commit()
                logging.info("レコード削除: ID=%s", record_id)
                self.load_data()  # データの再読み込み
            except sqlite3.Error as e:
                logging.error("レコード削除エラー: %s", e)
                QMessageBox.critical(self, "エラー", "データの削除中にエラーが発生しました。")
            finally:
                if conn:
                    conn.close()

    def show_monthly_summary(self):
        month = self.month_edit.text()
        if not re.match(r'^\d{4}-\d{2}$', month):
            QMessageBox.warning(None, "入力エラー", "月はYYYY-MM形式で入力してください。")
            return

        summary = fetch_monthly_summary(month)
        if summary:
            self.summary_label.setText(summary)
        else:
            self.summary_label.setText("データがありません")

    def export_csv(self):
        """CSVエクスポート用のファイルダイアログ表示"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "CSVファイル保存", os.path.expanduser("~"),
            "CSVファイル (*.csv);;すべてのファイル (*.*)"
        )
        if file_path:
            export_to_csv(file_path)

    def import_csv(self):
        """CSVインポート用のファイルダイアログ表示"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "CSVファイル選択", os.path.expanduser("~"),
            "CSVファイル (*.csv);;すべてのファイル (*.*)"
        )
        if file_path:
            if import_from_csv(file_path):
                self.load_data()  # データの再読み込み

    def export_monthly_csv_gui(self):
        """月別集計CSV出力用のファイルダイアログ表示"""
        selected_month = self.month_management_combo.currentText()
        if not selected_month or selected_month == "データがありません":
            QMessageBox.warning(self, "エラー", "対象月を選択してください")
            return
        
        # デフォルトファイル名を生成
        default_filename = f"月別集計_{selected_month.replace('-', '')}"
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "月別集計CSV保存", os.path.expanduser(f"~/{default_filename}.csv"),
            "CSVファイル (*.csv);;すべてのファイル (*.*)"
        )
        if file_path:
            export_monthly_summary_csv(selected_month, file_path)

    def export_multi_month_csv_gui(self):
        """複数月比較CSV出力用のダイアログ表示"""
        try:
            available_months = get_available_months()
            if not available_months:
                QMessageBox.warning(self, "エラー", "比較可能な月のデータがありません")
                return
            
            # 複数月選択用のダイアログを作成
            
            dialog = QDialog(self)
            dialog.setWindowTitle("複数月選択")
            dialog.setModal(True)
            dialog.resize(300, 400)
            
            layout = QVBoxLayout()
            
            # 説明ラベル
            label = QLabel("比較したい月を複数選択してください（Ctrlキーを押しながらクリック）")
            layout.addWidget(label)
            
            # 月選択リスト
            month_list = QListWidget()
            month_list.addItems(available_months)
            month_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
            layout.addWidget(month_list)
            
            # ボタン
            button_box = QDialogButtonBox(
                QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
            )
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addWidget(button_box)
            
            dialog.setLayout(layout)
            
            if dialog.exec() == QDialog.DialogCode.Accepted:
                selected_items = month_list.selectedItems()
                if not selected_items:
                    QMessageBox.warning(self, "エラー", "月を選択してください")
                    return
                
                selected_months = [item.text() for item in selected_items]
                
                # デフォルトファイル名を生成
                month_range = f"{selected_months[-1]}_{selected_months[0]}" if len(selected_months) > 1 else selected_months[0]
                default_filename = f"複数月比較_{month_range.replace('-', '')}"
                
                file_path, _ = QFileDialog.getSaveFileName(
                    self, "複数月比較CSV保存", os.path.expanduser(f"~/{default_filename}.csv"),
                    "CSVファイル (*.csv);;すべてのファイル (*.*)"
                )
                if file_path:
                    export_multiple_months_summary_csv(selected_months, file_path)
                    
        except Exception as e:
            logging.error(f"複数月CSV出力ダイアログエラー: {e}")
            QMessageBox.critical(self, "エラー", f"複数月CSV出力ダイアログでエラーが発生しました: {e}")

# -------------------------------------------------------------------
# 10. メイン処理
# -------------------------------------------------------------------
def main():
    """システムの初期化と起動"""
    try:
        # Windows日本語環境のための設定
        os.environ['QT_QPA_PLATFORM'] = 'windows'
        
        # Python 3.10以降のreconfigureメソッドの使用
        if sys.version_info >= (3, 10):
            try:
                sys.stdin.reconfigure(encoding='utf-8')
                sys.stdout.reconfigure(encoding='utf-8')
            except (AttributeError, RuntimeError):
                # 3.12では一部環境でエラーになることがある
                pass
        
        app = QApplication(sys.argv)
        
        # データベース初期化
        initialize_database()
        
        # スケジューラースレッド開始
        start_scheduler_thread()
        
        # アプリケーションウィンドウの表示
        ex = KakeiboApp()
        ex.show()
        
        # アプリケーションの実行 - PyQt6のバージョンに応じた対応
        # Python 3.10-3.12対応
        try:
            # PyQt 6.4.0以降（Python 3.11-3.12推奨）
            sys.exit(app.exec())
        except AttributeError:
            # 古いバージョンのPyQt6（Python 3.10）
            sys.exit(app.exec_())
            
    except Exception as e:
        logging.error(f"実行エラー: {e}")
        print(f"エラーが発生しました: {e}")
        traceback.print_exc()
        
        # GUIが起動していない場合に備えてコンソールにもメッセージを表示
        print("予期せぬエラーが発生しました。詳細はログファイルを確認してください。")
        print("ログファイル: " + os.path.join(os.path.dirname(os.path.abspath(__file__)), "kakeibo.log"))
        
        try:
            QMessageBox.critical(None, "エラー", f"予期せぬエラーが発生しました: {e}")
        except Exception:
            pass
        
        sys.exit(1)

if __name__ == "__main__":
    main()
