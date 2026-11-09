import hmac
import os
import uuid
from html.parser import HTMLParser
from typing import Any
from urllib.parse import quote, urlparse

import msal
import requests
from dotenv import load_dotenv
from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import HTMLResponse

load_dotenv()

app = FastAPI(title="Teams Data Viewer", version="0.1.0")
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


def _make_demo_data() -> dict[str, dict[str, Any]]:
    return {
        "team-101": {
            "id": "team-101",
            "name": "学校運営",
            "description": "学校の運営情報と共有事項",
            "channels": [
                {
                    "id": "channel-general",
                    "name": "全体連絡",
                    "messages": [
                        {
                            "id": "msg-1",
                            "author": "山田先生",
                            "content": "今週の校内ミーティングは木曜の 3 時限に実施します。",
                            "timestamp": "2026-10-26T09:00:00Z",
                        },
                        {
                            "id": "msg-2",
                            "author": "佐藤さん",
                            "content": "準備資料の共有をお願いします。",
                            "timestamp": "2026-10-26T09:40:00Z",
                        },
                    ],
                },
                {
                    "id": "channel-admin",
                    "name": "管理",
                    "messages": [
                        {
                            "id": "msg-3",
                            "author": "校務担当",
                            "content": "教材の発注予定を確認しています。",
                            "timestamp": "2026-10-26T10:00:00Z",
                        }
                    ],
                },
            ],
        },
        "team-202": {
            "id": "team-202",
            "name": "進路サポート",
            "description": "進路指導・相談の共有",
            "channels": [
                {
                    "id": "channel-career",
                    "name": "進路相談",
                    "messages": [
                        {
                            "id": "msg-4",
                            "author": "進路担当",
                            "content": "面談の予約枠を追加しました。",
                            "timestamp": "2026-10-26T11:00:00Z",
                        }
                    ],
                },
                {
                    "id": "channel-announce",
                    "name": "お知らせ",
                    "messages": [],
                },
            ],
        },
        "team-grade-1": {
            "id": "team-grade-1",
            "name": "1学年",
            "description": "1学年の連絡・共有",
            "channels": [
                {"id": "channel-grade-1", "name": "学年連絡", "messages": []},
            ],
        },
        "team-grade-2": {
            "id": "team-grade-2",
            "name": "2学年",
            "description": "2学年の連絡・共有",
            "channels": [
                {"id": "channel-grade-2", "name": "学年連絡", "messages": []},
            ],
        },
        "team-grade-3": {
            "id": "team-grade-3",
            "name": "3学年",
            "description": "3学年の連絡・共有",
            "channels": [
                {"id": "channel-grade-3", "name": "学年連絡", "messages": []},
            ],
        },
        "team-clubs": {
            "id": "team-clubs",
            "name": "部活動",
            "description": "部活動に関する連絡・共有",
            "channels": [
                {"id": "channel-clubs", "name": "部活動連絡", "messages": []},
            ],
        },
    }


app.state.sessions: dict[str, dict[str, str]] = {}
app.state.teams = _make_demo_data()
TIMETABLE_DAYS = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日"]
PERIOD_LABELS = ["1限", "2限", "3限", "4限", "5限", "6限"]
app.state.timetable = {
    "days": {
        "月曜日": ["国語", "数学", "英語", "理科", "社会", "体育"],
        "火曜日": ["数学", "英語", "国語", "体育", "理科", "美術"],
        "水曜日": ["英語", "理科", "数学", "社会", "国語", "音楽"],
        "木曜日": ["社会", "国語", "体育", "数学", "英語", "技術"],
        "金曜日": ["理科", "数学", "英語", "国語", "家庭科", "総合"],
    }
}


def get_m365_config() -> dict[str, str]:
    return {
        "tenant_id": os.getenv("TENANT_ID", ""),
        "client_id": os.getenv("CLIENT_ID", ""),
        "client_secret": os.getenv("CLIENT_SECRET", ""),
    }


def get_graph_access_token() -> str:
    config = get_m365_config()
    if not all(config.values()):
        raise RuntimeError(
            "Microsoft 365 の環境変数 TENANT_ID / CLIENT_ID / CLIENT_SECRET を設定してください。"
        )

    msal_client = msal.ConfidentialClientApplication(
        client_id=config["client_id"],
        authority=f"https://login.microsoftonline.com/{config['tenant_id']}",
        client_credential=config["client_secret"],
    )

    result = msal_client.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if not result or "access_token" not in result:
        details = result.get("error_description") if isinstance(result, dict) else "トークン取得に失敗しました。"
        raise RuntimeError(details)
    return result["access_token"]


def graph_is_configured() -> bool:
    return all(get_m365_config().values())


def graph_get_collection(
    path: str,
    access_token: str,
    params: dict[str, str] | None = None,
    extra_headers: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    url = path if path.startswith("https://") else f"{GRAPH_BASE_URL}{path}"
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    headers.update(extra_headers or {})
    items: list[dict[str, Any]] = []

    while url:
        parsed_url = urlparse(url)
        if parsed_url.scheme != "https" or parsed_url.netloc != "graph.microsoft.com":
            raise RuntimeError("Microsoft Graph から不正なページ URL が返されました。")
        try:
            response = requests.get(url, headers=headers, params=params, timeout=30)
        except requests.RequestException as error:
            raise RuntimeError("Microsoft Graph に接続できませんでした。") from error
        if not response.ok:
            raise RuntimeError(f"Microsoft Graph API が HTTP {response.status_code} を返しました。")
        try:
            payload = response.json()
        except ValueError as error:
            raise RuntimeError("Microsoft Graph から不正な JSON が返されました。") from error
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink", "")
        params = None

    return items


def get_graph_teams(access_token: str) -> list[dict[str, Any]]:
    groups = graph_get_collection(
        "/groups",
        access_token,
        params={
            "$filter": "resourceProvisioningOptions/Any(x:x eq 'Team')",
            "$select": "id,displayName,description",
            "$count": "true",
        },
        extra_headers={"ConsistencyLevel": "eventual"},
    )
    teams = []
    for group in groups:
        team_id = str(group.get("id", ""))
        if not team_id:
            continue
        channels = graph_get_collection(
            f"/teams/{quote(team_id, safe='')}/channels", access_token
        )
        teams.append(
            {
                "id": team_id,
                "name": group.get("displayName") or "名称未設定のチーム",
                "description": group.get("description") or "",
                "channels": [
                    {"id": channel.get("id", ""), "name": channel.get("displayName") or "名称未設定"}
                    for channel in channels
                    if channel.get("id")
                ],
            }
        )
    return teams


class _MessageBodyParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def _append_line_break(self) -> None:
        if not self.parts or not self.parts[-1].endswith("\n"):
            self.parts.append("\n")

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"br", "div", "li", "p"}:
            self._append_line_break()

    def handle_endtag(self, tag: str) -> None:
        if tag in {"div", "li", "p"}:
            self._append_line_break()

    def handle_data(self, data: str) -> None:
        self.parts.append(data)


def map_graph_message(
    message: dict[str, Any], team: dict[str, Any], channel: dict[str, Any]
) -> dict[str, str]:
    body = message.get("body") or {}
    parser = _MessageBodyParser()
    parser.feed(str(body.get("content") or ""))
    sender = message.get("from") or {}
    author = sender.get("user", {}).get("displayName") or sender.get("application", {}).get(
        "displayName"
    ) or "Teams"
    return {
        "id": str(message.get("id", "")),
        "author": author,
        "content": "".join(parser.parts).strip(),
        "timestamp": message.get("createdDateTime") or message.get("lastModifiedDateTime") or "",
        "teamId": str(team["id"]),
        "teamName": str(team["name"]),
        "channelId": str(channel["id"]),
        "channelName": str(channel["name"]),
    }


def get_graph_snapshot() -> dict[str, Any]:
    access_token = get_graph_access_token()
    teams = get_graph_teams(access_token)
    messages = []
    for team in teams:
        team_path = quote(team["id"], safe="")
        for channel in team["channels"]:
            channel_path = quote(channel["id"], safe="")
            graph_messages = graph_get_collection(
                f"/teams/{team_path}/channels/{channel_path}/messages",
                access_token,
                params={"$top": "50"},
            )
            messages.extend(map_graph_message(message, team, channel) for message in graph_messages)
    messages.sort(key=lambda message: message["timestamp"], reverse=True)
    return {"teams": teams, "messages": messages, "source": "graph"}


def get_demo_snapshot() -> dict[str, Any]:
    teams = [
        {
            "id": team_id,
            "name": team["name"],
            "description": team["description"],
            "channels": [
                {"id": channel["id"], "name": channel["name"]} for channel in team["channels"]
            ],
        }
        for team_id, team in app.state.teams.items()
    ]
    messages = []
    for team_id, team in app.state.teams.items():
        for channel in team["channels"]:
            messages.extend(
                {
                    **message,
                    "teamId": team_id,
                    "teamName": team["name"],
                    "channelId": channel["id"],
                    "channelName": channel["name"],
                }
                for message in channel["messages"]
            )
    messages.sort(key=lambda message: message["timestamp"], reverse=True)
    return {"teams": teams, "messages": messages, "source": "demo"}


def get_dashboard_snapshot() -> dict[str, Any]:
    if not graph_is_configured():
        return get_demo_snapshot()
    try:
        return get_graph_snapshot()
    except RuntimeError as error:
        raise HTTPException(status_code=502, detail=f"Microsoft Graph 連携に失敗しました: {error}") from error


def require_session(authorization: str | None = Header(default=None, alias="Authorization")) -> str:
    if not authorization:
        raise HTTPException(status_code=401, detail="ログインしてください")
    token = authorization.replace("Bearer ", "").strip()
    if token not in app.state.sessions:
        raise HTTPException(status_code=401, detail="有効なセッションがありません")
    return token


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "service": "teams-data-viewer"}


@app.post("/api/login")
def login(payload: dict[str, str]) -> dict[str, str]:
    username = (payload.get("username") or "").strip()
    password = payload.get("password") or ""
    if graph_is_configured():
        expected_username = os.getenv("APP_USERNAME", "").strip()
        expected_password = os.getenv("APP_PASSWORD", "")
        if not expected_username or not expected_password:
            raise HTTPException(status_code=503, detail="APP_USERNAME と APP_PASSWORD を設定してください")
        authenticated = hmac.compare_digest(username, expected_username) and hmac.compare_digest(
            password, expected_password
        )
    else:
        authenticated = username == "admin" and password == "admin123"
    if authenticated:
        token = uuid.uuid4().hex
        app.state.sessions[token] = {"username": username}
        return {"token": token, "username": username}
    raise HTTPException(status_code=401, detail="ユーザー名またはパスワードが違います")


@app.get("/", response_class=HTMLResponse)
def root() -> str:
    page = """
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Teams Data Viewer</title>
        <style>
            :root {
                --bg: #f3f6f2;
                --surface: #ffffff;
                --line: #dce4dd;
                --primary: #17634e;
                --primary-dark: #104737;
                --text: #1e2c26;
                --muted: #68766e;
                --accent: #e8ad4c;
                --success: #e9f5e9;
                --success-text: #28603d;
            }
            * { box-sizing: border-box; }
            body {
                margin: 0; font-family: "Yu Gothic UI", "Hiragino Kaku Gothic ProN", Meiryo, sans-serif;
                background: radial-gradient(ellipse at 12% 0%, #e8f0e5 0, transparent 34rem), var(--bg); color: var(--text);
            }
            .container {
                max-width: 1440px; margin: 36px auto; padding: 0 28px;
            }
            .login-box {
                max-width: 440px; margin: 8vh auto; padding: 38px 34px; background: var(--surface);
                border: 1px solid var(--line); border-top: 5px solid var(--primary); box-shadow: 0 18px 48px rgba(28, 54, 41, .08);
            }
            .dashboard { display: none; }
            .dashboard.visible { display: block; }
            h1, h2, p { margin-top: 0; }
            h1 { margin-bottom: 8px; font-size: 2rem; letter-spacing: 0; }
            h2 { margin-bottom: 0; font-size: 1.08rem; }
            .subtitle { color: var(--muted); margin-bottom: 20px; line-height: 1.7; }
            label { display: block; margin: 14px 0 7px; font-size: .88rem; font-weight: 700; }
            input, textarea, select, button {
                width: 100%; border: 1px solid var(--line); padding: 11px 12px; font: inherit; font-size: .94rem;
            }
            input, textarea, select { background: #fbfcfa; color: var(--text); }
            button { cursor: pointer; font-weight: 700; }
            button:focus-visible, input:focus-visible, textarea:focus-visible, select:focus-visible {
                outline: 3px solid rgba(23, 99, 78, .2); outline-offset: 2px;
            }
            .demo-note { margin-top: 12px; font-size: 0.86rem; color: var(--muted); }
            .topbar { display: flex; justify-content: space-between; align-items: center; min-height: 64px; border-bottom: 1px solid var(--line); }
            .brand { display: flex; align-items: center; gap: 11px; color: var(--primary-dark); font-size: .93rem; font-weight: 800; }
            .brand-mark { display: grid; place-items: center; width: 30px; height: 30px; background: var(--primary); color: white; font-size: .78rem; }
            .user-area { display: flex; align-items: center; gap: 16px; color: var(--muted); font-size: .88rem; }
            .logout-button { width: auto; padding: 8px 12px; background: transparent; color: var(--primary-dark); }
            .page-heading { display: flex; align-items: end; justify-content: space-between; gap: 24px; padding: 34px 0 26px; }
            .eyebrow { margin-bottom: 8px; color: var(--primary); font-size: .72rem; font-weight: 800; letter-spacing: .08em; }
            .page-heading p:not(.eyebrow) { margin-bottom: 0; color: var(--muted); }
            .sample-label { align-self: start; padding: 6px 9px; border: 1px solid #e8d3aa; color: #79551e; background: #fff8e9; font-size: .76rem; white-space: nowrap; }
            .home-grid { display: grid; grid-template-columns: 220px minmax(0, 1fr) 286px; gap: 26px; align-items: start; }
            .team-panel, .feed-panel, .compose-panel { min-width: 0; }
            .section-heading, .feed-toolbar { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
            .section-heading { padding: 0 0 12px; border-bottom: 1px solid var(--line); }
            .section-count { color: var(--muted); font-size: .8rem; }
            .team-list { display: grid; }
            .team-option { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 12px 10px; border: 0; border-bottom: 1px solid var(--line); background: transparent; color: var(--text); text-align: left; font-weight: 600; }
            .team-option:hover, .team-option.active { background: #e8f0e9; color: var(--primary-dark); }
            .team-option span:last-child { min-width: 24px; color: var(--muted); text-align: right; font-size: .8rem; }
            .feed-toolbar { padding-bottom: 12px; border-bottom: 1px solid var(--line); }
            .feed-toolbar p { margin: 5px 0 0; color: var(--muted); font-size: .82rem; }
            .refresh-button { width: auto; padding: 8px 10px; background: transparent; color: var(--primary); font-size: .84rem; }
            .feed-list { min-height: 250px; }
            .feed-item { display: grid; grid-template-columns: 9px minmax(0, 1fr); gap: 14px; padding: 18px 2px; border-bottom: 1px solid var(--line); animation: arrive .25s ease both; }
            .feed-marker { width: 7px; height: 7px; margin-top: 7px; background: var(--accent); }
            .feed-meta { display: flex; flex-wrap: wrap; align-items: center; gap: 7px; color: var(--muted); font-size: .77rem; }
            .feed-team { color: var(--primary); font-weight: 800; }
            .feed-separator { color: #a4aea7; }
            .feed-time { margin-left: auto; }
            .feed-author { margin-top: 7px; font-size: .87rem; font-weight: 700; }
            .feed-content { margin: 5px 0 0; line-height: 1.7; overflow-wrap: anywhere; }
            .empty-state { padding: 42px 12px; color: var(--muted); text-align: center; }
            .compose-panel { padding-left: 20px; border-left: 1px solid var(--line); }
            .compose-panel h2 { margin-bottom: 8px; }
            .compose-hint { margin-bottom: 14px; color: var(--muted); font-size: .85rem; line-height: 1.6; }
            .timetable-panel { margin-bottom: 22px; padding-bottom: 20px; border-bottom: 1px solid var(--line); }
            .timetable-heading { margin-bottom: 12px; }
            .timetable-edit-button, .schedule-action {
                width: auto; padding: 7px 10px; background: transparent; color: var(--primary); font-size: .82rem;
            }
            .timetable-day { margin-bottom: 12px; }
            .period-list { display: grid; gap: 7px; }
            .period-row { display: grid; grid-template-columns: 42px minmax(0, 1fr); align-items: center; gap: 8px; }
            .period-label { color: var(--muted); font-size: .8rem; font-weight: 700; }
            .period-subject { width: 100%; padding: 8px 9px; font-size: .86rem; }
            .period-subject:disabled { border-color: transparent; background: #f3f6f2; color: var(--text); opacity: 1; }
            .schedule-actions { display: flex; gap: 8px; margin-top: 12px; }
            .schedule-actions[hidden] { display: none; }
            .schedule-actions .schedule-save { border-color: var(--primary); background: var(--primary); color: white; }
            .compose-button, .login-button { margin-top: 14px; border: 0; background: var(--primary); color: white; }
            .compose-button:hover, .login-button:hover { background: var(--primary-dark); }
            .success { display: none; margin-bottom: 16px; padding: 10px 12px; background: var(--success); color: var(--success-text); }
            .success.show { display: block; }
            @keyframes arrive { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: translateY(0); } }
            @media (max-width: 1080px) {
                .home-grid { grid-template-columns: 190px minmax(0, 1fr); gap: 24px; }
                .compose-panel { grid-column: 1 / -1; padding: 20px 0 0; border-left: 0; border-top: 1px solid var(--line); }
                .compose-panel form { display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px; }
                .compose-panel label[for="newMessage"], #newMessage, .compose-button { grid-column: 1 / -1; }
            }
            @media (max-width: 680px) {
                .container { margin: 0 auto; padding: 0 16px 28px; }
                .login-box { margin: 8vh auto; padding: 30px 22px; }
                .topbar { min-height: 58px; }
                .user-area { gap: 8px; font-size: .8rem; }
                .page-heading { align-items: start; padding: 26px 0 20px; }
                h1 { font-size: 1.7rem; }
                .home-grid { grid-template-columns: minmax(0, 1fr); gap: 24px; }
                .team-list { display: flex; overflow-x: auto; }
                .team-option { flex: 0 0 auto; border: 1px solid var(--line); margin-right: 7px; padding: 9px 10px; }
                .feed-time { flex-basis: 100%; margin-left: 0; }
                .compose-panel form { display: block; }
                .compose-panel label[for="newMessage"], #newMessage, .compose-button { grid-column: auto; }
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div id="loginBox" class="login-box">
                <h1>Teams Data Viewer</h1>
                <div class="subtitle">学校の Microsoft 365 環境向けに Teams データを確認するためのログイン画面です。</div>
                <form id="loginForm" method="post" action="/api/login">
                    <label for="username">ユーザー名</label>
                    <input id="username" name="username" autocomplete="username" required />
                    <label for="password">パスワード</label>
                    <input id="password" name="password" type="password" autocomplete="current-password" required />
                    <button type="submit">ログイン</button>
                </form>
                <div id="loginHelp" class="demo-note">__LOGIN_HELP__</div>
            </div>

            <div id="dashboard" class="dashboard">
                <div class="topbar">
                    <div class="brand"><span class="brand-mark">T</span><span>学校連絡</span></div>
                    <div class="user-area"><span><span id="currentUser"></span> でログイン中</span>
                        <button id="logoutButton" class="logout-button" type="button">ログアウト</button>
                    </div>
                </div>

                <section class="page-heading">
                    <div>
                        <p class="eyebrow">TEAM COMMUNICATIONS</p>
                        <h1>連絡ホーム</h1>
                        <p>チームから届いた連絡をまとめて確認できます。</p>
                    </div>
                    <span id="sourceLabel" class="sample-label">現在はサンプルデータを表示</span>
                </section>

                <div id="messageBox" class="success"></div>

                <div class="home-grid">
                    <aside class="team-panel">
                        <div class="section-heading"><h2>チーム</h2><span id="teamCount" class="section-count"></span></div>
                        <nav id="teamList" class="team-list" aria-label="チーム別に連絡を絞り込む"></nav>
                    </aside>

                    <section class="feed-panel" aria-labelledby="feedHeading">
                        <div class="feed-toolbar">
                            <div><h2 id="feedHeading">最新の連絡</h2><p id="feedMeta">すべてのチーム</p></div>
                            <button id="refreshButton" class="refresh-button" type="button">再読み込み</button>
                        </div>
                        <div id="feedList" class="feed-list" aria-live="polite"></div>
                    </section>

                    <aside class="compose-panel">
                        <section class="timetable-panel" aria-labelledby="timetableHeading">
                            <div class="section-heading timetable-heading">
                                <h2 id="timetableHeading">授業の時間割</h2>
                                <button id="scheduleEditButton" class="timetable-edit-button" type="button">編集</button>
                            </div>
                            <label for="scheduleDay">曜日</label>
                            <select id="scheduleDay" class="timetable-day"></select>
                            <div id="periodList" class="period-list"></div>
                            <div id="scheduleActions" class="schedule-actions" hidden>
                                <button id="scheduleSaveButton" class="schedule-action schedule-save" type="button">変更を保存</button>
                                <button id="scheduleCancelButton" class="schedule-action" type="button">キャンセル</button>
                            </div>
                        </section>
                        <p class="eyebrow">QUICK POST</p>
                        <h2>チャネルに投稿</h2>
                        <p class="compose-hint">投稿先を選んでチームに連絡します。</p>
                        <p id="readOnlyNote" class="compose-hint" hidden>Graph 接続中は閲覧専用です。投稿にはユーザー委任認証が必要です。</p>
                        <form id="messageForm">
                            <label for="teamSelect">チーム</label>
                            <select id="teamSelect" required></select>
                            <label for="channelSelect">チャネル</label>
                            <select id="channelSelect" required></select>
                            <label for="newMessage">メッセージ</label>
                            <textarea id="newMessage" rows="5" placeholder="連絡内容を入力" required></textarea>
                            <button class="compose-button" type="submit">投稿する</button>
                        </form>
                    </aside>
                </div>
            </div>
        </div>

        <script>
            const state = {
                token: localStorage.getItem('teamsToken') || '',
                teams: [],
                messages: [],
                source: 'demo',
                selectedTeamId: '',
                selectedChannelId: '',
                selectedTeamFilter: '',
                timetable: null,
                selectedScheduleDay: '月曜日',
                editingSchedule: false,
                user: localStorage.getItem('teamsUser') || ''
            };

            const loginBox = document.getElementById('loginBox');
            const dashboard = document.getElementById('dashboard');
            const teamList = document.getElementById('teamList');
            const feedList = document.getElementById('feedList');
            const feedMeta = document.getElementById('feedMeta');
            const currentUser = document.getElementById('currentUser');
            const messageBox = document.getElementById('messageBox');
            const teamSelect = document.getElementById('teamSelect');
            const channelSelect = document.getElementById('channelSelect');
            const scheduleDay = document.getElementById('scheduleDay');
            const periodList = document.getElementById('periodList');
            const scheduleActions = document.getElementById('scheduleActions');

            function showMessage(text) {
                messageBox.textContent = text;
                messageBox.classList.add('show');
                setTimeout(() => messageBox.classList.remove('show'), 2500);
            }

            function escapeHtml(value) {
                return String(value ?? '').replace(/[&<>"']/g, character => ({
                    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
                })[character]);
            }

            function setAuthUI() {
                const loggedIn = !!state.token;
                loginBox.style.display = loggedIn ? 'none' : 'block';
                dashboard.classList.toggle('visible', loggedIn);
                currentUser.textContent = state.user || '未ログイン';
            }

            async function authFetch(url, options = {}) {
                const headers = { ...(options.headers || {}) };
                if (state.token) {
                    headers['Authorization'] = `Bearer ${state.token}`;
                }
                return fetch(url, { ...options, headers });
            }

            function renderTeams() {
                document.getElementById('teamCount').textContent = `${state.teams.length} チーム`;
                const allCount = state.messages.length;
                const options = [{ id: '', name: 'すべての連絡', count: allCount }, ...state.teams.map(team => ({
                    id: team.id,
                    name: team.name,
                    count: state.messages.filter(message => message.teamId === team.id).length
                }))];
                teamList.innerHTML = options.map(team => `
                    <button class="team-option ${team.id === state.selectedTeamFilter ? 'active' : ''}" type="button" data-team-filter="${escapeHtml(team.id)}">
                        <span>${escapeHtml(team.name)}</span><span>${team.count}</span>
                    </button>
                `).join('');
                teamList.querySelectorAll('[data-team-filter]').forEach(button => {
                    button.addEventListener('click', () => {
                        state.selectedTeamFilter = button.dataset.teamFilter;
                        renderTeams();
                        renderFeed();
                    });
                });
            }

            function renderChannels() {
                teamSelect.innerHTML = state.teams.map(team =>
                    `<option value="${escapeHtml(team.id)}">${escapeHtml(team.name)}</option>`
                ).join('');
                if (!state.teams.some(team => team.id === state.selectedTeamId)) {
                    state.selectedTeamId = state.teams[0]?.id || '';
                }
                teamSelect.value = state.selectedTeamId;
                const selectedTeam = state.teams.find(team => team.id === state.selectedTeamId);
                const channels = selectedTeam?.channels || [];
                if (!channels.some(channel => channel.id === state.selectedChannelId)) {
                    state.selectedChannelId = channels[0]?.id || '';
                }
                channelSelect.innerHTML = channels.map(channel =>
                    `<option value="${escapeHtml(channel.id)}">${escapeHtml(channel.name)}</option>`
                ).join('');
                channelSelect.value = state.selectedChannelId;
                teamSelect.disabled = !state.teams.length;
                channelSelect.disabled = !channels.length;
            }

            function renderFeed() {
                const visibleMessages = state.selectedTeamFilter
                    ? state.messages.filter(message => message.teamId === state.selectedTeamFilter)
                    : state.messages;
                const selectedTeam = state.teams.find(team => team.id === state.selectedTeamFilter);
                feedMeta.textContent = `${selectedTeam ? selectedTeam.name : 'すべてのチーム'} ・ ${visibleMessages.length} 件`;
                if (!visibleMessages.length) {
                    feedList.innerHTML = '<div class="empty-state">表示できる連絡はありません。</div>';
                    return;
                }
                feedList.innerHTML = visibleMessages.map(message => `
                    <article class="feed-item">
                        <span class="feed-marker" aria-hidden="true"></span>
                        <div>
                            <div class="feed-meta">
                                <span class="feed-team">${escapeHtml(message.teamName)}</span>
                                <span class="feed-separator">/</span>
                                <span>${escapeHtml(message.channelName)}</span>
                                <time class="feed-time" datetime="${escapeHtml(message.timestamp)}">${new Date(message.timestamp).toLocaleString('ja-JP')}</time>
                            </div>
                            <div class="feed-author">${escapeHtml(message.author)}</div>
                            <p class="feed-content">${escapeHtml(message.content).replace(/\\n/g, '<br>')}</p>
                        </div>
                    </article>
                `).join('');
            }

            async function loadDashboard() {
                const response = await authFetch('/api/dashboard');
                const data = await response.json();
                if (!response.ok) {
                    throw new Error(data.detail || '連絡ホームの読み込みに失敗しました');
                }
                state.teams = data.teams || [];
                state.messages = data.messages || [];
                state.source = data.source || 'demo';
                const isGraph = state.source === 'graph';
                document.getElementById('sourceLabel').textContent = isGraph
                    ? 'Microsoft Teams から取得'
                    : '現在はサンプルデータを表示';
                document.getElementById('messageForm').hidden = isGraph;
                document.getElementById('readOnlyNote').hidden = !isGraph;
                renderChannels();
                renderTeams();
                renderFeed();
                await loadTimetable();
            }

            async function loadTimetable() {
                const response = await authFetch('/api/timetable');
                const data = await response.json();
                if (!response.ok) {
                    throw new Error(data.detail || '時間割の読み込みに失敗しました');
                }
                state.timetable = data;
                state.editingSchedule = false;
                renderTimetable();
            }

            function renderTimetable() {
                if (!state.timetable) return;
                const days = Object.keys(state.timetable.days);
                if (!days.includes(state.selectedScheduleDay)) {
                    state.selectedScheduleDay = days[0] || '';
                }
                scheduleDay.innerHTML = days.map(day =>
                    `<option value="${escapeHtml(day)}">${escapeHtml(day)}</option>`
                ).join('');
                scheduleDay.value = state.selectedScheduleDay;
                scheduleDay.disabled = state.editingSchedule;
                const subjects = state.timetable.days[state.selectedScheduleDay] || [];
                periodList.innerHTML = subjects.map((subject, index) => `
                    <label class="period-row">
                        <span class="period-label">${index + 1}限</span>
                        <input class="period-subject" type="text" maxlength="40"
                            aria-label="${index + 1}限の授業"
                            value="${escapeHtml(subject)}" ${state.editingSchedule ? '' : 'disabled'} />
                    </label>
                `).join('');
                document.getElementById('scheduleEditButton').hidden = state.editingSchedule;
                scheduleActions.hidden = !state.editingSchedule;
            }

            async function loadFeed() {
                const response = await authFetch('/api/messages');
                const data = await response.json();
                if (!response.ok) {
                    throw new Error(data.detail || '連絡一覧の取得に失敗しました');
                }
                state.messages = data.messages || [];
                state.source = data.source || state.source;
                renderTeams();
                renderFeed();
            }

            document.getElementById('loginForm').addEventListener('submit', async (event) => {
                event.preventDefault();
                const payload = {
                    username: document.getElementById('username').value,
                    password: document.getElementById('password').value
                };
                const response = await fetch('/api/login', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const data = await response.json();
                if (!response.ok) {
                    alert(data.detail || 'ログインに失敗しました');
                    return;
                }
                state.token = data.token;
                state.user = data.username;
                localStorage.setItem('teamsToken', state.token);
                localStorage.setItem('teamsUser', state.user);
                setAuthUI();
                try {
                    await loadDashboard();
                } catch (error) {
                    showMessage(error.message || '連絡ホームを読み込めませんでした');
                }
            });

            document.getElementById('logoutButton').addEventListener('click', () => {
                state.token = '';
                state.user = '';
                localStorage.removeItem('teamsToken');
                localStorage.removeItem('teamsUser');
                setAuthUI();
            });

            teamSelect.addEventListener('change', () => {
                state.selectedTeamId = teamSelect.value;
                state.selectedChannelId = '';
                renderChannels();
            });

            channelSelect.addEventListener('change', () => {
                state.selectedChannelId = channelSelect.value;
            });

            scheduleDay.addEventListener('change', () => {
                state.selectedScheduleDay = scheduleDay.value;
                renderTimetable();
            });

            document.getElementById('scheduleEditButton').addEventListener('click', () => {
                state.editingSchedule = true;
                renderTimetable();
            });

            document.getElementById('scheduleCancelButton').addEventListener('click', () => {
                loadTimetable().catch(error => showMessage(error.message || '時間割を読み込めませんでした'));
            });

            document.getElementById('scheduleSaveButton').addEventListener('click', async () => {
                const subjects = [...periodList.querySelectorAll('.period-subject')].map(input => input.value.trim());
                const days = { ...state.timetable.days, [state.selectedScheduleDay]: subjects };
                const response = await authFetch('/api/timetable', {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ days })
                });
                const data = await response.json();
                if (!response.ok) {
                    showMessage(data.detail || '時間割を保存できませんでした');
                    return;
                }
                state.timetable = data;
                state.editingSchedule = false;
                renderTimetable();
                showMessage('時間割を変更しました');
            });

            document.getElementById('refreshButton').addEventListener('click', () => {
                loadFeed().catch(() => showMessage('連絡一覧を更新できませんでした'));
            });

            document.getElementById('messageForm').addEventListener('submit', async (event) => {
                event.preventDefault();
                const content = document.getElementById('newMessage').value.trim();
                if (!content || !state.selectedTeamId || !state.selectedChannelId) return;
                const response = await authFetch(`/api/teams/${state.selectedTeamId}/channels/${state.selectedChannelId}/messages`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ author: state.user, content })
                });
                const data = await response.json();
                if (!response.ok) {
                    alert(data.detail || '投稿に失敗しました');
                    return;
                }
                document.getElementById('newMessage').value = '';
                showMessage('投稿を追加しました');
                await loadFeed();
            });

            if (state.token) {
                setAuthUI();
                loadDashboard().catch((error) => {
                    showMessage(error.message || '連絡ホームを読み込めませんでした');
                });
            } else {
                setAuthUI();
            }
        </script>
    </body>
    </html>
    """
    login_help = (
        "APP_USERNAME と APP_PASSWORD に設定したアカウントでログインしてください。"
        if graph_is_configured()
        else "デモ用: admin / admin123"
    )
    return page.replace("__LOGIN_HELP__", login_help)


@app.get("/api/teams")
def get_joined_teams(authorization: str | None = Header(default=None, alias="Authorization")) -> dict[str, Any]:
    require_session(authorization)
    snapshot = get_dashboard_snapshot()
    return {"teams": snapshot["teams"], "source": snapshot["source"]}


@app.get("/api/dashboard")
def get_dashboard(authorization: str | None = Header(default=None, alias="Authorization")) -> dict[str, Any]:
    require_session(authorization)
    return get_dashboard_snapshot()


@app.get("/api/messages")
def get_all_messages(authorization: str | None = Header(default=None, alias="Authorization")) -> dict[str, Any]:
    require_session(authorization)
    snapshot = get_dashboard_snapshot()
    return {"messages": snapshot["messages"], "source": snapshot["source"]}


@app.get("/api/timetable")
def get_timetable(
    authorization: str | None = Header(default=None, alias="Authorization"),
) -> dict[str, Any]:
    require_session(authorization)
    return app.state.timetable


@app.put("/api/timetable")
def update_timetable(
    payload: dict[str, Any],
    authorization: str | None = Header(default=None, alias="Authorization"),
) -> dict[str, Any]:
    require_session(authorization)
    days = payload.get("days")
    if not isinstance(days, dict) or set(days) != set(TIMETABLE_DAYS):
        raise HTTPException(status_code=400, detail="月曜日から金曜日までの時間割を指定してください")
    for day in TIMETABLE_DAYS:
        subjects = days[day]
        if (
            not isinstance(subjects, list)
            or len(subjects) != len(PERIOD_LABELS)
            or any(not isinstance(subject, str) or len(subject) > 40 for subject in subjects)
        ):
            raise HTTPException(status_code=400, detail=f"{day}は6限分の授業名（40文字以内）を指定してください")
    app.state.timetable = {"days": {day: list(days[day]) for day in TIMETABLE_DAYS}}
    return app.state.timetable


@app.get("/api/teams/{team_id}/channels")
def get_team_channels(team_id: str, authorization: str | None = Header(default=None, alias="Authorization")) -> dict[str, Any]:
    require_session(authorization)
    if graph_is_configured():
        snapshot = get_dashboard_snapshot()
        team = next((item for item in snapshot["teams"] if item["id"] == team_id), None)
    else:
        team = app.state.teams.get(team_id)
    if team is None:
        raise HTTPException(status_code=404, detail="チームが見つかりません")
    return {
        "teamId": team_id,
        "teamName": team["name"],
        "channels": [{"id": channel["id"], "name": channel["name"]} for channel in team["channels"]],
    }


@app.get("/api/teams/{team_id}/channels/{channel_id}/messages")
def get_channel_messages(
    team_id: str,
    channel_id: str,
    authorization: str | None = Header(default=None, alias="Authorization"),
) -> dict[str, Any]:
    require_session(authorization)
    if graph_is_configured():
        snapshot = get_dashboard_snapshot()
        team = next((item for item in snapshot["teams"] if item["id"] == team_id), None)
        if team is None:
            raise HTTPException(status_code=404, detail="チームが見つかりません")
        channel = next((item for item in team["channels"] if item["id"] == channel_id), None)
        if channel is None:
            raise HTTPException(status_code=404, detail="チャネルが見つかりません")
        messages = [
            message
            for message in snapshot["messages"]
            if message["teamId"] == team_id and message["channelId"] == channel_id
        ]
        return {"teamId": team_id, "channelId": channel_id, "messages": messages}
    team = app.state.teams.get(team_id)
    if team is None:
        raise HTTPException(status_code=404, detail="チームが見つかりません")
    for channel in team["channels"]:
        if channel["id"] == channel_id:
            return {"teamId": team_id, "channelId": channel_id, "messages": channel["messages"]}
    raise HTTPException(status_code=404, detail="チャネルが見つかりません")


@app.post("/api/teams/{team_id}/channels/{channel_id}/messages")
def create_channel_message(
    team_id: str,
    channel_id: str,
    payload: dict[str, str],
    authorization: str | None = Header(default=None, alias="Authorization"),
) -> dict[str, Any]:
    require_session(authorization)
    if graph_is_configured():
        raise HTTPException(status_code=405, detail="Graph 接続中は閲覧専用です")
    team = app.state.teams.get(team_id)
    if team is None:
        raise HTTPException(status_code=404, detail="チームが見つかりません")
    for channel in team["channels"]:
        if channel["id"] == channel_id:
            content = (payload.get("content") or "").strip()
            author = (payload.get("author") or "unknown").strip() or "unknown"
            if not content:
                raise HTTPException(status_code=400, detail="メッセージ内容が空です")
            message = {
                "id": f"msg-{uuid.uuid4().hex[:8]}",
                "author": author,
                "content": content,
                "timestamp": __import__("datetime").datetime.utcnow().isoformat() + "Z",
            }
            channel["messages"].append(message)
            return {"message": message}
    raise HTTPException(status_code=404, detail="チャネルが見つかりません")


@app.get("/api/graph-status")
def graph_status() -> dict[str, Any]:
    try:
        get_graph_access_token()
        return {"status": "configured", "message": "Microsoft Graph の認証設定が有効です"}
    except RuntimeError:
        return {"status": "demo-mode", "message": "環境変数が未設定のため、デモデータを使用しています"}


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
