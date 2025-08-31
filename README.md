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
- Python 3.10-3.12（推奨: Python 3.12）
- SQLite
- pandas 2.0.0以上
- PyQt6 6.4.0以上
- schedule 1.2.0以上

## インストール
以下の手順で環境をセットアップします。

1. 必要なパッケージをインストールします：
   ```sh
   pip install -r requirements.txt
   ```

   または個別にインストール：
   ```sh
   pip install PyQt6>=6.4.0 pandas>=2.0.0 schedule>=1.2.0
   ```

## 起動方法

### Windowsの場合
1. `run_kakeibo.bat` ファイルをダブルクリックして実行します。
   - 自動的にPythonを検出して実行します。
   - エラーが発生した場合は、必要なライブラリのインストール方法が表示されます。

### 手動での起動
1. コマンドプロンプトまたはPowerShellを開き、以下のコマンドを実行します：
   ```sh
   python kakeibo.py
   ```
   
   または
   
   ```sh
   py -3 kakeibo.py
   ```

### VSCodeでの実行（Code Runner使用）
1. VSCodeでkakeibo.pyを開きます
2. 右クリックして「Run Code」を選択するか、Ctrl+Alt+Nを押します
3. 出力ウィンドウに結果が表示されます

### 動作確認
テスト用のスクリプト `test_runner.py` を実行して、Python環境と日本語表示が正常かどうかを確認できます。

### 主な機能
- **データの入力**：日付、区分、科目、金額を入力して保存
- **データの表示・編集**：保存されたデータの表示と編集
- **月次集計**：指定した月の収支集計結果を表示
- **CSVエクスポート**：保存されたデータをCSV形式でエクスポート
- **バックアップ**：データベースファイルのバックアップを作成

## 注意事項
- データベースファイルはアプリケーションのディレクトリに `kakeibo.db` という名前で保存されます。
- エラーハンドリングとログ出力のために `kakeibo.log` ファイルが生成されます。
- 初回起動時に必要なファイルが自動的に作成されます。

## トラブルシューティング

### 起動しない場合
1. 必要なライブラリがインストールされているか確認：
   ```sh
   pip list | findstr PyQt6
   pip list | findstr pandas 
   pip list | findstr schedule
   ```

2. 不足しているライブラリをインストール：
   ```sh
   pip install PyQt6 pandas schedule
   ```

3. ログファイル（kakeibo.log）を確認して、エラーの詳細を確認してください。

### Code Runnerのエラー
VSCodeのCode Runnerプラグインで「指定されたパスが見つかりません」エラーが表示される場合：

1. `.vscode/settings.json` ファイルが存在することを確認
2. VSCodeを再起動
3. `test_runner.py` を実行して動作確認

## 貢献
バグ報告や機能リクエストは、[Issues](https://github.com/zapabob/kakeibo/issues) からお願いします。

## ライセンス
このプロジェクトはMITライセンスのもとで公開されています。詳細は [LICENSE](LICENSE) ファイルをご覧ください。
```

You can create this file in your repository as `README.md`.
