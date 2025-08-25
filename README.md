# Python Project (Windows)

このリポジトリは Windows 上で Python プロジェクトを実行するための手順を示します。

## 実行手順

```powershell
# 仮想環境作成と有効化
python -m venv venv
.\venv\Scripts\activate

# 依存パッケージインストール
pip install -r requirements.txt

# Playwright ブラウザインストール（Chromium のみでも可）
playwright install

# プロジェクト実行
python main_ui_proxy.py
