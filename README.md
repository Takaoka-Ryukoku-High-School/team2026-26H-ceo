# Microsoft Teams 連携 Webアプリ（Python）

このプロジェクトは、学校の Microsoft 365 環境で Microsoft Teams のデータを表示・操作するための Web アプリの雛形です。

## 目的

- Teams の所属チーム一覧を表示する
- チーム内のチャネル情報を取得する
- チャネルメッセージを確認する
- 将来的に Teams のデータを自作アプリに表示して業務に活用する

## アーキテクチャ

- アプリ種別: Web アプリ
- 開発言語: Python
- 実行基盤: FastAPI
- 連携先: Microsoft Graph API
- 認証方式: Microsoft Entra ID (Azure AD) のアプリ登録

## 必要な Microsoft 365 構成

1. Microsoft Entra ID の App registrations で新しいアプリを作成
2. API permissions に以下を追加
   - `Group.Read.All` (Application)
   - `Channel.ReadBasic.All` (Application)
   - `ChannelMessage.Read.All` (Application)
3. テナント管理者の同意を付与
4. 「Certificates & secrets」からクライアントシークレットを発行
5. Microsoft Graph の接続情報とアプリのログイン情報を `.env` に設定

## 環境変数

```env
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
APP_USERNAME=trusted-operator
APP_PASSWORD=replace-with-a-long-random-password
```

Graph 接続時はアプリ専用権限でテナント内の Teams 投稿を読み取ります。画面は読み取り専用で、投稿機能にはユーザー委任認証が別途必要です。アプリ専用権限はサインイン中の利用者ごとの Teams 権限には限定されないため、管理者の同意とアプリへのアクセスを厳格に管理してください。利用者単位でアクセスを制御する本番運用では、Microsoft Entra ID の委任認証を実装してください。

## 実行方法

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

ブラウザで `http://localhost:8000` を開くと、Teams の情報を取得する画面が表示されます。

## 代表的な Microsoft Graph 呼び出し

- `GET /me/joinedTeams`
- `GET /teams/{team-id}/channels`
- `GET /teams/{team-id}/channels/{channel-id}/messages`

## Teams への埋め込み

このWebアプリを Teams に載せる場合は、Teams タブまたは Teams アプリとして公開し、Web URL を Teams に登録します。これにより、学校のユーザーが Teams の中からアプリを開いてデータを確認できます。

## 注意点

- Microsoft 365 の管理者による許可が必要になる場合があります。
- 本番運用では、シークレット管理やログ、権限の最小化、エラーハンドリングを追加する必要があります。
- 実際のデータアクセスは、利用権限とユーザー権限に応じて制御してください。

## 開発の進め方

1. Microsoft 365 のアプリ登録を行う
2. OAuth でサインインを実装する
3. Graph API で Teams とチャネルを取得する
4. 画面に一覧表示する
5. 必要に応じて保存・編集・通知機能を追加する
