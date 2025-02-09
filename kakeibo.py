#!/usr/bin/env python3
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

import tkinter as tk
from tkinter import messagebox, filedialog
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

# -------------------------------------------------------------------
# 1. ログ設定
# -------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filename="kakeibo.log",
    filemode="a"
)
logging.info("家計簿日報システム 起動")

# データベースファイル名
if getattr(sys, 'frozen', False):
    DATABASE_NAME = os.path.join(sys._MEIPASS, "kakeibo.db")
else:
    DATABASE_NAME = "kakeibo.db"

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
        logging.info("データベース初期化完了")
    except sqlite3.Error as e:
        logging.error("DB初期化エラー: %s", e)
        messagebox.showerror("エラー", "データベースの初期化に失敗しました。")
    finally:
        if conn:
            conn.close()

# -------------------------------------------------------------------
# 3. 入力バリデーション
# -------------------------------------------------------------------
def validate_input(date_str, amount_str):
    """
    ユーザー入力のバリデーション
    - 日付はYYYY-MM-DD形式であるか
    - 金額は数値変換できるか
    """
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        messagebox.showwarning("入力エラー", "日付はYYYY-MM-DD形式で入力してください。")
        return False
    try:
        float(amount_str)
    except ValueError:
        messagebox.showwarning("入力エラー", "金額は数値で入力してください。")
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
    except sqlite3.Error as e:
        logging.error("DB挿入エラー: %s", e)
        messagebox.showerror("エラー", "データの保存中にエラーが発生しました。詳細はログを確認してください。")
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
            messagebox.showinfo("集計結果", "指定された月のデータはありません。")
            return
        income_total = df[df['category'] == '収入']['amount'].sum()
        expense_total = df[df['category'] == '支出']['amount'].sum()
        balance = income_total - expense_total
        summary_text = (
            f"【{month}の集計結果】\n"
            f"収入合計: {income_total}\n"
            f"支出合計: {expense_total}\n"
            f"差引残高: {balance}"
        )
        messagebox.showinfo("月次集計", summary_text)
        logging.info("月次集計実施: %s", summary_text)
    except Exception as e:
        logging.error("集計エラー: %s", e)
        messagebox.showerror("エラー", "集計処理中にエラーが発生しました。詳細はログを確認してください。")
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
        messagebox.showinfo("バックアップ", f"バックアップが成功しました。\nファイル名: {backup_filename}")
    except Exception as e:
        logging.error("バックアップ失敗: %s", e)
        messagebox.showerror("エラー", "バックアップ中にエラーが発生しました。")

# -------------------------------------------------------------------
# 7. リマインダー機能 (schedule と threading 利用)
# -------------------------------------------------------------------
def reminder_job():
    """リマインダーのジョブ関数。指定時刻になるとユーザーに家計簿入力を促す。"""
    logging.info("リマインダー実行: 今日の家計簿入力は済んでるかな？")
    messagebox.showinfo("リマインダー", "今日の家計簿入力は済んでるかな？")

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
# 8. CSVエクスポート機能 (新規追加)
# -------------------------------------------------------------------
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
            messagebox.showwarning("エラー", "エクスポートするデータがありません")
            return
        df.to_csv(filename, index=False, encoding='utf-8-sig')
        messagebox.showinfo("成功", f"CSVファイルをエクスポートしました\n{filename}")
        logging.info("CSVエクスポート成功: %s", filename)
    except Exception as e:
        logging.error("CSVエクスポートエラー: %s", e)
        messagebox.showerror("エラー", "CSVのエクスポート中にエラーが発生しました")
    finally:
        if conn:
            conn.close()

def export_csv(self):
    """CSVエクスポート用のファイルダイアログ表示"""
    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
        title="保存先を選択",
        initialdir=os.path.expanduser("~")  # ホームディレクトリを初期パスに設定
    )
    if file_path:
        export_to_csv(file_path)

# -------------------------------------------------------------------
# 9. GUI アプリケーション (tkinter)
# -------------------------------------------------------------------
class KakeiboApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("家計簿日報システム (仮称)")
        self.geometry("500x400")
        self.create_widgets()

    def create_widgets(self):
        # 日付入力
        self.lbl_date = tk.Label(self, text="日付 (YYYY-MM-DD):")
        self.lbl_date.pack(pady=5)
        self.entry_date = tk.Entry(self)
        self.entry_date.pack()
        self.entry_date.insert(0, datetime.date.today().strftime("%Y-%m-%d"))

        # 区分 (収入/支出) の選択
        self.lbl_category = tk.Label(self, text="区分:")
        self.lbl_category.pack(pady=5)
        self.category_var = tk.StringVar(self)
        self.category_var.set("支出")
        self.option_menu = tk.OptionMenu(self, self.category_var, "収入", "支出")
        self.option_menu.pack()

        # 科目入力
        self.lbl_subject = tk.Label(self, text="科目:")
        self.lbl_subject.pack(pady=5)
        self.entry_subject = tk.Entry(self)
        self.entry_subject.pack()

        # 金額入力
        self.lbl_amount = tk.Label(self, text="金額:")
        self.lbl_amount.pack(pady=5)
        self.entry_amount = tk.Entry(self)
        self.entry_amount.pack()

        # 保存ボタン
        self.btn_save = tk.Button(self, text="保存", command=self.save_record)
        self.btn_save.pack(pady=10)

        # 月次集計ボタン
        self.btn_summary = tk.Button(self, text="月次集計", command=self.show_monthly_summary)
        self.btn_summary.pack(pady=5)

        # バックアップボタン
        self.btn_backup = tk.Button(self, text="バックアップ", command=backup_database)
        self.btn_backup.pack(pady=5)

        # CSVエクスポートボタン
        self.btn_csv = tk.Button(self, text="CSVエクスポート", command=self.export_csv)
        self.btn_csv.pack(pady=5)

    def save_record(self):
        """入力データのバリデーションとDB保存"""
        date = self.entry_date.get()
        category = self.category_var.get()
        subject = self.entry_subject.get()
        amount_str = self.entry_amount.get()

        if not validate_input(date, amount_str):
            return
        amount = float(amount_str)
        insert_record(date, category, subject, amount)
        messagebox.showinfo("成功", "データが保存されました！")
        self.entry_subject.delete(0, tk.END)
        self.entry_amount.delete(0, tk.END)

    def show_monthly_summary(self):
        """月次集計の実行"""
        date = self.entry_date.get()
        if not re.match(r'^\d{4}-\d{2}-\d{2}$', date):
            messagebox.showwarning("入力エラー", "日付はYYYY-MM-DD形式で入力してください。")
            return
        month = date[:7]
        fetch_monthly_summary(month)

    def export_csv(self):
        """CSVエクスポート用のファイルダイアログ表示"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
            title="保存先を選択",
            initialdir=os.path.expanduser("~")  # ホームディレクトリを初期パスに設定
        )
        if file_path:
            export_to_csv(file_path)

# -------------------------------------------------------------------
# 10. メイン処理
# -------------------------------------------------------------------
def main():
    """システムの初期化と起動"""
    initialize_database()
    start_scheduler_thread()
    app = KakeiboApp()
    app.mainloop()

if __name__ == "__main__":
    main()