# 請求書処理完全自動化システム

月次請求書の分割・電子印押印・メール下書き作成を自動化するPythonツールです。

**月額1万円超のサービス費用を0円に。作業時間を30分から3分に短縮。**

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)

---

## 📖 目次

- [概要](#概要)
- [主な機能](#主な機能)
- [導入効果](#導入効果)
- [必要要件](#必要要件)
- [インストール](#インストール)
- [使い方](#使い方)
- [ファイル構成](#ファイル構成)
- [設定方法](#設定方法)
- [トラブルシューティング](#トラブルシューティング)
- [開発の経緯](#開発の経緯)
- [ライセンス](#ライセンス)
- [貢献](#貢献)

---

## 概要

月次で発行される複数社の請求書PDFを、手作業で処理していませんか？

- 会社ごとにPDFを分割
- 管理者印・担当者印・社印を押印
- メールを1通ずつ作成
- ファイル名を手動で変更

このツールは、これらの作業を**完全自動化**します。

### ビフォー・アフター

| 項目 | 手作業 | 自動化後 |
|:-----|-------:|---------:|
| 処理時間 | 約30分 | 約3分 |
| ミス発生 | 月3件程度 | 0件 |
| 属人化 | 特定の担当者のみ | 誰でも実行可能 |
| コスト | - | 年間13万円削減 |

---

## 主な機能

### ✅ PDF自動分割
- 請求書番号を自動認識
- 会社ごとに分割保存
- 控えページ・空白ページを自動スキップ
- 「当月取引なし」ページを自動除外

### ✅ 電子印押印
- 管理者印・担当者印・社印を自動押印
- 座標指定による正確な位置配置
- 1ページ目のみに押印（控えは対象外）

### ✅ 年月フォルダ自動生成
- PDFから締め日を抽出
- `YYYY年MM月` フォルダを自動作成
- 既存フォルダがあれば自動で追加保存

### ✅ メール下書き自動作成
- Outlookメール下書きを自動生成
- CC機能対応
- 会社ごとにカスタマイズ可能なテンプレート

### ✅ ファイル名自動生成
- `YYMMDDCompanyName請求書.pdf` 形式
- アンダーバーなしでスッキリ

---

## 導入効果

### 💰 コスト削減効果

**業者サービスと比較した場合の削減効果:**

| 項目 | 業者サービス | 自作ツール |
|:-----|-------------:|-----------:|
| 初期費用 | 0円 | 0円 |
| 月額 | 10,500円 | 0円 |
| 年間 | 126,000円 | 0円 |
| 5年間 | 630,000円 | 0円 |
| 10年間 | 1,260,000円 | 0円 |

**→ 10年間で約126万円の削減**

---

### ⏱️ ビフォー・アフター

| 項目 | 手作業 | 自動化後 |
|:-----|-------:|---------:|
| 処理時間 | 約30分 | 約3分 |
| ミス発生 | 月3件程度 | 0件 |
| 属人化 | 特定の担当者のみ | 誰でも実行可能 |
| コスト | - | 年間13万円削減 |

---

## 必要要件

### システム要件

- **OS**: Windows 10/11
- **Python**: 3.8以上
- **Microsoft Outlook**: メール下書き作成機能を使う場合

### 必要なファイル

1. **電子印画像（PNG形式）**
   - `管理者.png` - 管理者印（必須）
   - `担当者.png` - 担当者印（必須）
   - `社印.png` - 社印（任意）

2. **会社マスターExcel**
   - 3シート構造（会社マスタ/メール/保存先）

---

## インストール

### 1. リポジトリのクローン

```bash
git clone https://github.com/yourusername/invoice-automation-system.git
cd invoice-automation-system
```

### 2. 依存パッケージのインストール

```bash
pip install -r requirements.txt
```

### 3. EXEファイルの作成（オプション）

```bash
python build_invoice_exe.py
```

---

## 使い方

### 基本的な使い方

1. **EXEファイルを起動**
   ```
   請求書処理システムv5.exe
   ```

2. **請求書PDFを選択**
   - ファイル選択ダイアログが表示されます

3. **自動処理が開始**
   - 会社マスター・電子印フォルダを自動検出
   - 保存先は会社マスターから自動取得
   - 年月フォルダを自動生成

4. **完了**
   - 処理完了ダイアログが表示されます
   - 指定フォルダにPDFが保存されます
   - Outlookに下書きメールが作成されます

### 処理フロー

```
PDFファイルを選択
    ↓
会社マスター・電子印を自動検出
    ↓
保存先を会社マスターから取得
    ↓
PDFから締め日を抽出 → 年月フォルダ生成
    ↓
会社ごとにPDF分割
    ↓
1ページ目に電子印押印
    ↓
Outlookメール下書き作成
    ↓
完了！
```

---

## ファイル構成

```
invoice-automation-system/
├── invoice_automation_system.py　   # メインプログラム
├── README.md                        # このファイル
├── 会社マスター_サンプル.xlsx        # サンプルExcel
└── 電子印サンプル/                   # サンプルフォルダ
    └── README.txt
```

---

## 設定方法

### 会社マスターExcelの設定

#### シート1: 会社マスタ

| 会社名 | メールアドレス | CC |
|--------|---------------|-----|
| サンプル株式会社 | invoice@sample.co.jp | |
| テスト工業株式会社 | accounting@test.jp | manager@test.jp |

#### シート2: メール

```
A1: メールタイトル
A2: 【会社名】YYYY年MM月分 ご請求書送付の件
A3: (空行)
A4: メール本文
A5: A:A
A6: ご担当者様
A7: 
A8: いつもお世話になっております。
A9: ...
```

**プレースホルダー:**
- `A:A` → 会社名に自動置換
- `YYYY年MM月` → 締め日から自動置換（例: 2025年12月）

#### シート3: 保存先

| フォルダ | パス |
|---------|------|
| - | \\server\shared\invoices |

**パス形式:**
- ローカルパス: `C:\請求書`
- ネットワークパス: `\\server\shared\invoices`

### 電子印画像の設定

**推奨サイズ:**
- 管理者印・担当者印: 20x20ピクセル
- 社印: 60x60ピクセル

**形式:**
- PNG形式（透過背景推奨）

**配置座標:**
- 管理者: X=457, Y=655
- 担当者: X=498, Y=655
- 社印: X=500, Y=680

---

## トラブルシューティング

### Q1. 「管理者.png が見つかりません」と表示される

**A1:** 電子印フォルダに `管理者.png` と `担当者.png` があることを確認してください。

### Q2. PDFが分割されない

**A2:** 請求書に「№」（請求書番号）が正しく記載されているか確認してください。

### Q3. Outlookメールが作成されない

**A3:** Outlookがインストールされていない場合は正常動作です（PDF生成のみ実行されます）。

### Q4. 会社マスター読み込みエラー

**A4:** 2シート構造（会社マスタ/メール）になっているか確認してください。サンプルExcelを参考にしてください。

### Q5. 保存先フォルダにPDFが保存されない

**A5:** 
- 会社マスター「保存先」シートのB1セルにパスが入力されているか確認
- ネットワークパスへの書き込み権限があるか確認
- `invoice_automation.log` ファイルでエラー内容を確認

### Q6. EXE起動が遅い

**A6:** PyInstallerのワンファイルEXEは初回起動時に展開が必要です（10-30秒程度）。処理自体は3分程度で完了します。

---

## 開発の経緯

### きっかけ

ある日の雑談で、経理担当者から「PDF分割サービスを導入したいけど、月1万円超えるから保留中」と聞きました。

「それ、作れるんじゃない？」

### 開発時間

**約1時間**

- PDF分割: 20分
- 印鑑押印: 20分
- Excel連携: 10分
- Outlook連携: 10分

### 結果

- **コスト削減**: 年間13万円
- **時間削減**: 月30分 → 3分
- **ミス削減**: 月3件 → 0件

### 学んだこと

小さな自動化が、大きな価値を生む。
技術は、人を楽にするためにある。

---

## 技術スタック

### 主要ライブラリ

- **PDF処理**: `pypdf`, `pdfplumber`
- **印鑑押印**: `reportlab`
- **Excel処理**: `openpyxl`
- **Outlook連携**: `win32com` (pywin32)
- **GUI**: `tkinter`

### アーキテクチャ

```python
CompanyMasterReader  # Excel読み込み
    ↓
SealManager         # 電子印管理
    ↓
InvoicePDFProcessor # PDF分割・押印
    ↓
OutlookMailCreator  # メール下書き作成
    ↓
InvoiceAutomationSystem # 統合処理
```

### 主要機能の実装

#### PDF分割

```python
def split_pdf(self, input_pdf: Path) -> bool:
    with pdfplumber.open(input_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            invoice_number = self._extract_invoice_info(text)
            # 会社ごとに分割
```

#### 電子印押印

```python
def _add_seals_to_pdf(self, input_pdf: Path, output_pdf: Path):
    c = canvas.Canvas(tmp_pdf_path, pagesize=A4)
    c.drawImage(
        str(seal_path),
        x=457, y=655,
        width=20, height=20,
        mask='auto'
    )
```

#### 年月フォルダ生成

```python
def process(self, input_pdf: Path, output_base_dir: Path):
    # 締め日から年月抽出
    year, month, day = close_date.split('-')
    folder_name = f"{year}年{int(month)}月"
    
    # フォルダ作成
    output_dir = output_base_dir / folder_name
    output_dir.mkdir(parents=True, exist_ok=True)
```

---

## ライセンス

MIT License

Copyright (c) 2025

本ソフトウェアは無償で提供され、商用・非商用を問わず自由に使用・改変・配布できます。

---

## 貢献

### バグ報告・機能要望

Issuesでお気軽にご報告ください。

### プルリクエスト

歓迎します！以下の点にご留意ください：

1. コードスタイル: PEP8準拠
2. 型ヒント: 必須
3. docstring: 関数・クラスに必須
4. テストコード: 推奨

### 開発に参加する

```bash
# フォーク → クローン
git clone https://github.com/yourusername/invoice-automation-system.git

# ブランチ作成
git checkout -b feature/new-feature

# コミット
git commit -m "Add: 新機能追加"

# プッシュ
git push origin feature/new-feature

# プルリクエスト作成
```

---

## よくある質問（FAQ）

### Q. 他のPDFにも対応できますか？

A. はい、コードをカスタマイズすることで対応可能です。特に `_extract_invoice_info()` メソッドを修正してください。

### Q. Mac/Linuxでも動きますか？

A. Outlook連携機能以外は動作します。Outlook連携は `win32com` を使用しているためWindows専用です。

### Q. 商用利用できますか？

A. はい、MITライセンスのため商用利用可能です。

### Q. カスタマイズ開発を依頼できますか？

A. Issuesまたは作者のSNSまでご連絡ください。

---

## 作者

**業務自動化エンジニア**

- GitHub: [@yourusername](https://github.com/rancorder)
- note: [noteアカウント](https://note.com/rancorder)
- blog:(https://portfolio-crystal-dreamscape.vercel.app/)
---

## 謝辞

このツールは、現場の声から生まれました。

「月1万円、高いよね...」

その一言が、126万円の価値を生み出しました。

業務改善に悩むすべての方へ。
小さな自動化が、大きな変化を生みます。

---

## スター・フォークをお願いします！⭐

このツールが役に立ったら、ぜひスターをお願いします。
あなたの会社でも、年13万円削減できるかもしれません。

**次はあなたの番です。**

---

最終更新: 2025-12-25
