# Infill - CSV差し込みテキスト出力ツール

![License](https://img.shields.io/github/license/Ikumyon/Infill)
![Python](https://img.shields.io/badge/python-3.14%2B-blue)
![PySide6](https://img.shields.io/badge/UI-PySide6-green)

**Infill** は、CSVデータとテンプレートファイルを組み合わせて、大量のテキストファイルを一括生成するためのデスクトップアプリケーションです。

メールの下書き、名札のデータ、プログラムのコード生成など、特定のパターンにデータを流し込む作業を直感的な操作で効率化します。

---

## 📥 ダウンロード

Windows向けにビルド済みの実行ファイル（.exe）を公開しています。Pythonのインストールは不要で、解凍してすぐに利用可能です。

[**最新バージョンのダウンロードはこちら**](https://github.com/Ikumyon/Infill/releases/latest)

---

## 🚀 主な機能

- **直感的なドラッグ＆ドロップ**: CSV、ひな形、出力先をドロップするだけで準備完了。
- **柔軟なプレースホルダー**: `{name}` や `[[id]]` など、開始・終了文字列を自由に設定可能。
- **リアルタイムプレビュー**: 置換後の結果をその場で確認。不足している項目は赤くハイライトされます。
- **多様な出力設定**:
  - **個別出力**: CSVの1行を1ファイルとして保存。ファイル名は指定した列から自動生成。
  - **一括出力**: すべてのデータを1つのファイルに結合して保存。
  - **文字コード・改行コード指定**: UTF-8, Shift_JIS, CRLF, LF に対応。
- **設定の自動保存**: 最後に使ったパスや設定を次回起動時に復元します。

---

## 🛠️ 動作環境

- Windows 10 / 11
- Python 3.14 以上 (開発用)
- 依存ライブラリ: `PySide6`

---

## 📖 使い方

### 1. 準備するもの
- **CSVファイル**: 1行目にヘッダーがあるもの。
- **ひな形ファイル**: テキストファイル（`.txt`など）。置換したい箇所を `{name}` のように記述します。

### 2. 操作手順
1. アプリを起動し、CSVファイルとひな形ファイルを画面にドロップします。
2. 出力先フォルダを指定します。
3. 必要に応じて「プレビュー」タブで結果を確認します。
4. 「出力実行」ボタンをクリックします。

---

## 💻 開発者向け

### セットアップ
```bash
git clone https://github.com/Ikumyon/Infill.git
cd Infill
pip install -r requirements.txt
```

### 実行
```bash
python main.py
```

### ビルド (PyInstaller)
```bash
pyinstaller --noconsole --onefile --icon=app_icon.ico --add-data "app_icon.ico;." main.py
```

---

## 📄 ライセンス

このプロジェクトは [MIT License](LICENSE) のもとで公開されています。
© 2026 Ikumyon
