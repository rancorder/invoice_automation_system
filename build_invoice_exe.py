#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請求書処理自動化システム - EXEビルドスクリプト

このスクリプトは invoice_automation_system_v5.py を
Windows実行可能ファイル(.exe)に変換します。

使い方:
    python build_invoice_exe.py

出力:
    dist/InvoiceAutomationSystem.exe
"""

import os
import sys
import subprocess
from pathlib import Path


def check_pyinstaller():
    """PyInstallerがインストールされているか確認"""
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__} が見つかりました")
        return True
    except ImportError:
        print("✗ PyInstallerがインストールされていません")
        print("\n以下のコマンドでインストールしてください:")
        print("  pip install pyinstaller")
        return False


def build_exe():
    """EXEファイルをビルド"""
    
    # スクリプトファイルの確認
    script_path = Path("invoice_automation_system_v5.py")
    if not script_path.exists():
        print(f"✗ エラー: {script_path} が見つかりません")
        return False
    
    print("="*70)
    print("請求書処理自動化システム - EXEビルド開始")
    print("="*70)
    print(f"ソースファイル: {script_path}")
    print()
    
    # PyInstallerコマンド構築
    command = [
        "pyinstaller",
        "--onefile",                      # 単一実行ファイル
        "--windowed",                     # コンソールウィンドウを非表示
        "--name=InvoiceAutomationSystem", # 出力ファイル名
        "--clean",                        # ビルド前にキャッシュをクリア
        str(script_path)
    ]
    
    # アイコンファイルがあれば追加
    icon_path = Path("icon.ico")
    if icon_path.exists():
        command.extend(["--icon", str(icon_path)])
        print(f"✓ アイコン: {icon_path}")
    else:
        print("ℹ アイコンファイル(icon.ico)が見つかりません（デフォルトアイコンを使用）")
    
    print()
    print("ビルドコマンド:")
    print(" ".join(command))
    print()
    print("ビルド中... (数分かかる場合があります)")
    print("-"*70)
    
    try:
        # PyInstallerを実行
        result = subprocess.run(
            command,
            check=True,
            capture_output=False,
            text=True
        )
        
        print("-"*70)
        print()
        print("="*70)
        print("✓ ビルド成功！")
        print("="*70)
        
        exe_path = Path("dist") / "InvoiceAutomationSystem.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n出力ファイル: {exe_path}")
            print(f"ファイルサイズ: {size_mb:.2f} MB")
            print()
            print("【使い方】")
            print("  1. dist/InvoiceAutomationSystem.exe を任意の場所にコピー")
            print("  2. 同じフォルダに以下を配置:")
            print("     - 会社マスター.xlsx")
            print("     - 電子印フォルダ")
            print("  3. EXEをダブルクリックして実行")
            print()
        
        return True
        
    except subprocess.CalledProcessError as e:
        print("-"*70)
        print()
        print("="*70)
        print("✗ ビルド失敗")
        print("="*70)
        print(f"エラー内容: {e}")
        return False
    except Exception as e:
        print()
        print("="*70)
        print("✗ 予期しないエラー")
        print("="*70)
        print(f"エラー内容: {e}")
        return False


def clean_build_files():
    """ビルド時の中間ファイルを削除"""
    import shutil
    
    dirs_to_remove = ["build", "__pycache__"]
    files_to_remove = ["*.spec"]
    
    print()
    print("中間ファイルのクリーンアップ...")
    
    for dir_name in dirs_to_remove:
        dir_path = Path(dir_name)
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"  削除: {dir_path}/")
    
    for pattern in files_to_remove:
        for file_path in Path(".").glob(pattern):
            file_path.unlink()
            print(f"  削除: {file_path}")
    
    print("✓ クリーンアップ完了")


def main():
    """メイン処理"""
    
    # PyInstallerの確認
    if not check_pyinstaller():
        return 1
    
    print()
    
    # ビルド実行
    success = build_exe()
    
    if success:
        # クリーンアップ
        clean_build_files()
        
        print()
        print("="*70)
        print("すべての処理が完了しました！")
        print("="*70)
        return 0
    else:
        return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nビルドを中断しました")
        sys.exit(1)
    except Exception as e:
        print(f"\n予期しないエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
