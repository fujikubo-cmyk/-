# 資金繰りメール日次レポート — セットアップ手順

## 必要なもの

| 項目 | 取得先 |
|------|--------|
| Anthropic APIキー | https://console.anthropic.com/ |
| LINE Notifyトークン | https://notify-bot.line.me/my/ |

---

## 手順

### 1. GASプロジェクトを作成

1. https://script.google.com/ を開く
2. 「新しいプロジェクト」をクリック
3. `Code.gs` の内容を貼り付けて保存（Ctrl+S）

---

### 2. Google Drive API を有効化

1. GASエディタ左メニュー「サービス（＋）」をクリック
2. 「Google Drive API」を選択 → バージョン **v3** → 「追加」

---

### 3. スクリプトプロパティを設定

1. GASエディタ上部「プロジェクトの設定（歯車アイコン）」をクリック
2. 「スクリプト プロパティ」セクションで以下を追加：

| プロパティ名 | 値 |
|-------------|-----|
| `CLAUDE_API_KEY` | Anthropic コンソールで発行したAPIキー |
| `LINE_NOTIFY_TOKEN` | LINE Notifyで発行したトークン |
| `REPORT_EMAIL` | レポートを受け取りたいメールアドレス |

---

### 4. 日次トリガーを設定

GASエディタで `setupTrigger` 関数を **1回だけ** 実行します：

1. 関数選択ドロップダウンで `setupTrigger` を選択
2. 「実行」ボタンをクリック
3. 権限の確認ダイアログ → 「権限を確認」→ Googleアカウントでログイン → 「許可」

これで毎朝 **8:00 JST** に自動実行されます。

---

### 5. 動作確認（手動テスト）

1. 関数選択ドロップダウンで `runDailyReport` を選択して実行
2. 「実行ログ」（表示 > ログ）で処理状況を確認
3. 指定メールアドレスにHTMLレポートが届いていることを確認
4. LINEにプッシュ通知が届いていることを確認

---

## メール検索条件

以下のGmailクエリで対象メールを絞り込んでいます：

```
(subject:資金繰り OR subject:経理) has:attachment (filename:xlsx OR filename:xls) newer_than:1d
```

件名に「資金繰り」または「経理」が含まれ、かつExcelファイルが添付されたメールが対象です。
条件を変更したい場合は `Code.gs` の `fetchCashFlowEmails()` 内の `query` 変数を編集してください。

---

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| 「対象メールなし」ログ | Gmail検索クエリを確認。手動でGmailを検索して件名を確認 |
| Claude APIエラー | `CLAUDE_API_KEY` が正しく設定されているか確認 |
| LINE通知が来ない | `LINE_NOTIFY_TOKEN` を確認。LINE Notifyのトークンは再発行も可能 |
| Drive APIエラー | サービスにGoogle Drive API v3 が追加されているか確認 |
| 実行時間超過（6分制限） | 添付ファイルが多い場合は `query` の `newer_than:1d` を `newer_than:12h` に変更 |
