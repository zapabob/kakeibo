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
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QTableWidget, QTableWidgetItem, QTabWidget,
    QFileDialog, QMessageBox
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
def validate_input(date_str, amount_str):
    """
    ユーザー入力のバリデーション
    - 日付はYYYY-MM-DD形式であるか
    - 金額は数値変換できるか
    """
    if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
        QMessageBox.warning(None, "入力エラー", "日付はYYYY-MM-DD形式で入力してください。")
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

        self.tabs.addTab(self.input_tab, "入力")
        self.tabs.addTab(self.view_tab, "表示・編集")
        self.tabs.addTab(self.summary_tab, "月次集計")


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

        if not validate_input(date, amount_str):
            return

        amount = float(amount_str)
        if insert_record(date, category, subject, amount):
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
        except:
            pass
        
        sys.exit(1)

if __name__ == "__main__":
    main()
