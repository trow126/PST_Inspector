# -*- coding: utf-8 -*-

import win32com.client
import hashlib
import os

# --- 1. 設定 ---
# 分析したいPSTファイルの絶対パスを指定してください
PST_FILE_PATH = os.path.abspath("your_email.pst")

# --- 2. 設計: グローバル変数とカウンターの初期化 ---
# メール用
unique_email_hashes = set()
duplicate_email_count = 0
unique_email_count = 0

# 会議用
unique_meeting_hashes = set()
duplicate_meeting_count = 0
unique_meeting_count = 0

# その他
skipped_item_count = 0

def process_folder(folder):
    """
    フォルダを再帰的に処理し、アイテムの種類に応じた重複チェックを行う。
    - Input: Outlook Folder Object
    - Output: None (グローバル変数を更新)
    """
    global unique_email_count, duplicate_email_count
    global unique_meeting_count, duplicate_meeting_count
    global skipped_item_count

    # アイテムを逆順に処理（将来的に削除処理を追加する場合も考慮）
    for item in reversed(folder.Items):
        try:
            # --- 3. 機能仕様: アイテム種別判定と重複判定 ---
            
            # 3.1. 通常のメールの場合 (olMailItem, Class: 43)
            if item.Class == 43:
                # 3.1.1. メールの重複基準に基づいてキーを生成
                key_source = (
                    str(item.SentOn), item.SenderName, item.To,
                    item.Subject, item.Body.strip()
                )
                key_string = "".join(map(str, key_source))
                item_hash = hashlib.sha256(key_string.encode('utf-8', 'ignore')).hexdigest()

                # 3.1.2. 重複チェック
                if item_hash in unique_email_hashes:
                    duplicate_email_count += 1
                else:
                    unique_email_hashes.add(item_hash)
                    unique_email_count += 1
            
            # 3.2. 会議の出席依頼 (53) またはキャンセル通知 (56) の場合
            elif item.Class in [53, 56]:
                # 3.2.1. 会議の重複基準に基づいてキーを生成
                appointment = item.GetAssociatedAppointment(True)
                key_source = (
                    appointment.Subject,
                    str(appointment.Start), str(appointment.End),
                    appointment.Location, appointment.RequiredAttendees,
                    appointment.Body.strip()
                )
                key_string = "".join(map(str, key_source))
                item_hash = hashlib.sha256(key_string.encode('utf-8', 'ignore')).hexdigest()

                # 3.2.2. 重複チェック
                if item_hash in unique_meeting_hashes:
                    duplicate_meeting_count += 1
                else:
                    unique_meeting_hashes.add(item_hash)
                    unique_meeting_count += 1

            # 3.3. 上記以外のアイテムはスキップ
            else:
                skipped_item_count += 1
        
        # 4.3. エラーハンドリング: 個別アイテムのエラーはスキップ
        except Exception as e:
            print(f"アイテム処理中にエラー (件名: {getattr(item, 'Subject', 'N/A')}): {e}")
            skipped_item_count += 1

    # サブフォルダを再帰的に処理
    for subfolder in folder.Folders:
        process_folder(subfolder)

# --- 4. 設計: メイン処理ブロック ---
if __name__ == "__main__":
    pst_store = None
    try:
        # 4.3. エラーハンドリング: Outlook接続とPST読み込み
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        namespace.AddStore(PST_FILE_PATH)
        # ファイルパスでストアを特定
        pst_store = next(s for s in namespace.Stores if s.FilePath == PST_FILE_PATH)
        root_folder = pst_store.GetRootFolder()
        
        print(f"'{root_folder.Name}' フォルダの分析を開始します...")
        process_folder(root_folder)

    except Exception as e:
        print(f"処理の開始に失敗しました。Outlookがインストールされているか、PSTファイルのパスが正しいか確認してください。")
        print(f"エラー詳細: {e}")

    finally:
        # 4.2. 処理フロー: PSTストアの削除
        if pst_store:
            namespace.RemoveStore(pst_store)
            print("\nPSTストアを閉じました。")
        
        # 3.4. 結果出力機能
        print("\n--- ✅ 分析結果 ---")
        print("📧 メール")
        print(f"  - ユニーク: {unique_email_count}件")
        print(f"  - 重複: {duplicate_email_count}件")
        print("\n🗓️ 会議（招待・キャンセル）")
        print(f"  - ユニーク: {unique_meeting_count}件")
        print(f"  - 重複: {duplicate_meeting_count}件")
        print("\n⏭️ その他")
        print(f"  - スキップしたアイテム: {skipped_item_count}件")
