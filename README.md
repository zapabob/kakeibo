
# 家計簿日報システム (仮称)

## 概要
このシステムは、日々の収支を記録し、月次集計を行う家計簿アプリケーションです。以下の機能を提供します：

- 日々の収支入力 (日付、区分、科目、金額)
- SQLiteを利用したDBへのデータ保存（パラメータ化クエリでSQLインジェクション対策済み）
- 月次集計機能（pandasを使った収入・支出・差引計算）
- GUI (PyQt6)による直感的な操作画面
- リマインダー機能 (scheduleライブラリとスレッドで毎日特定時刻に通知)
- バックアップ機能 (DBファイルのコピー)
- エラーハンドリングとログ出力 (loggingモジュール利用)
- CSVエクスポート機能 (新規追加)

## 動作環境
- Python 3.x
- SQLite
- pandas
- PyQt6
- schedule
- その他の必要なパッケージは `requirements.txt` を参照してください

## インストール
以下の手順で環境をセットアップします。

1. リポジトリをクローンします：
   ```sh
   git clone https://github.com/zapabob/kakeibo.git
   cd kakeibo
   ```

2. 必要なパッケージをインストールします：
   ```sh
   pip install -r requirements.txt
   ```

## 使い方
1. データベースを初期化します：
   ```sh
   python kakeibo.py --init-db
   ```
   
2. アプリケーションを起動します：
   ```sh
   python kakeibo.py
   ```

### 主な機能
- **データの入力**：日付、区分、科目、金額を入力して保存
- **データの表示・編集**：保存されたデータの表示と編集
- **月次集計**：指定した月の収支集計結果を表示
- **CSVエクスポート**：保存されたデータをCSV形式でエクスポート
- **バックアップ**：データベースファイルのバックアップを作成

## 注意事項
- データベースファイルはアプリケーションのディレクトリに `kakeibo.db` という名前で保存されます。
- エラーハンドリングとログ出力のために `kakeibo.log` ファイルが生成されます。

## 貢献
バグ報告や機能リクエストは、[Issues](https://github.com/zapabob/kakeibo/issues) からお願いします。

## ライセンス
このプロジェクトはMITライセンスのもとで公開されています。詳細は [LICENSE](LICENSE) ファイルをご覧ください。
```

You can create this file in your repository as `README.md`.
