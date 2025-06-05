# PowerPoint統合ツール

このリポジトリには複数のPowerPointファイルを1つにまとめるGUIツールが含まれています。

## 使い方
1. `python -m pip install -r requirements.txt` で依存ライブラリをインストールします。ネットワークからパッケージを取得できる環境が必要です。
2. `python src/merge_ppt_gui.py` を実行します。ウィンドウが表示されるので、統合したい `.pptx` ファイルを追加し、上下ボタンで順番を調整します。
3. "Edit TOC" ボタンで目次を編集できます。デフォルトではファイル名とスライド数が自動で入力されています。
4. "Merge" ボタンを押すと保存先を聞かれ、指定した場所に統合されたPowerPointが出力されます。

## 配布用バイナリの作成例
PyInstaller を使用すると Python の実行環境がなくても動作する単一ファイルの実行形式を作成できます。Windows 環境で下記コマンドを実行してください。

```bash
pip install pyinstaller
pyinstaller --onefile src/merge_ppt_gui.py
```

生成された `dist/merge_ppt_gui.exe` を配布することでインストール不要で動作します。必要に応じて Inno Setup 等でインストーラーを作成してください。

## 注意
- `python-pptx` を用いてスライドをコピーしているため、一部の高度なアニメーションやマクロは保持されない可能性があります。
- Windows 版 PowerPoint がインストールされた環境で COM を利用すると、より正確にスライドを複製できます。必要に応じてスクリプトを書き換えてください。
