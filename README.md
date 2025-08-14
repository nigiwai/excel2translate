# excel2translate

## 説明

ExcelファイルのテキストをAzure OpenAI APIを使用して英語から日本語に翻訳するツールです。指定した列のテキストを翻訳し、隣接する列または新しい列に日本語訳を追加します。

## 前提条件

- Python 3.11以上
- pip (Pythonパッケージインストーラー)
- Azure OpenAIアカウントとAPIキー、エンドポイント

## 準備

1. 仮想環境を作成してアクティブにします:

    ```sh
    python -m venv venv
    .\venv\Scripts\activate.ps1
    ```

2. 必要なパッケージをインストールします:

    python 3.11.9で確認

    ```sh
    pip install -r requirements.txt
    ```

3. .envファイルを作成する

   ```env
   AZURE_OPENAI_API_KEY="xxxxxxxxxx"
   AZURE_OPENAI_ENDPOINT="https://hogehoge-westus-openai.openai.azure.com/"
   AZURE_OPENAI_API_VERSION="2024-05-01-preview"
   AZURE_OPENAI_DEPLOYMENT_NAME="gpt-4o"
   START_ROW=2
   ```

## 使用方法

```sh
python translate.py <Excelファイルパス> <対象列番号> [シート名] [--add] [--force]
```

### パラメータ

- `<Excelファイルパス>`: 翻訳対象のExcelファイルパス（必須）
- `<対象列番号>`: 翻訳する列の番号（必須、1から開始）
- `[シート名]`: 処理対象のシート名（オプション、指定しない場合は全シートを処理）
- `[--add]`: 新しい列に翻訳結果を追加（オプション、指定しない場合は隣接列に出力）
- `[--force]`: 既存の翻訳を上書き（オプション、指定しない場合は空の翻訳列のみ処理）

### 実行例

```sh
# 3列目を翻訳し、隣接列に出力
python translate.py sample.xlsx 3

# 特定のシートの3列目を翻訳し、新しい列に出力
python translate.py sample.xlsx 3 "Sheet1" --add

# 既存の翻訳を強制的に上書き
python translate.py sample.xlsx 3 --force
```

## 機能

- Azure OpenAIを使用した高品質な翻訳
- 日本語テキストの自動スキップ機能
- 複数シートの一括処理
- 既存翻訳の保護（--forceオプションで上書き可能）
- 翻訳開始行の設定可能（環境変数START_ROW）

## ライセンス

MITライセンス
