# PowerPoint統合ツール

このリポジトリには複数のPowerPointファイルを1つにまとめるGUIツールが含まれています。
Python がインストールされていない環境でも実行できるよう、PyInstaller を用いた配布方法を紹介します。

## 使い方
1. `python -m pip install -r requirements.txt` で依存ライブラリをインストールします。ネットワークからパッケージを取得できる環境が必要です。
2. `python src/merge_ppt_gui.py` を実行します。ウィンドウが表示されるので、統合したい `.pptx` ファイルを追加し、上下ボタンで順番を調整します。
3. "Edit TOC" ボタンで目次を編集できます。デフォルトではファイル名とスライド数が自動で入力されています。
4. "Merge" ボタンを押すと保存先を聞かれ、指定した場所に統合されたPowerPointが出力されます。

## 配布用バイナリの作成例
PyInstaller を使用すると Python の実行環境がなくても動作する単一ファイルの実行形式を作成できます。下記は Windows での例です。
他の OS でも同様に `pyinstaller` を実行することでネイティブ実行形式を生成できます。

```bash
pip install pyinstaller
pyinstaller --onefile src/merge_ppt_gui.py
# 任意: 配布用に ZIP へまとめる
cd dist && zip ../merge_tool_win.zip merge_ppt_gui.exe && cd ..
```

生成された `dist/merge_ppt_gui.exe` または作成した `merge_tool_win.zip` を配布することで、Python が無い環境でも実行できます。さらにインストーラーを作成したい場合は Inno Setup などを利用してください。
Inno Setup 用の簡易スクリプト例を以下に示します。

```
[Setup]
AppName=MergePPT
AppVersion=1.0
DefaultDirName={pf}\MergePPT
OutputBaseFilename=MergePPTSetup

[Files]
Source: "dist\merge_ppt_gui.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Merge PPT"; Filename: "{app}\merge_ppt_gui.exe"
```

## 注意
- `python-pptx` を用いてスライドをコピーしているため、一部の高度なアニメーションやマクロは保持されない可能性があります。
- Windows 版 PowerPoint がインストールされた環境で COM を利用すると、より正確にスライドを複製できます。必要に応じてスクリプトを書き換えてください。
