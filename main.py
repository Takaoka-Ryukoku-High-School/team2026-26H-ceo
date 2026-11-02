import os
import uuid
from typing import Any

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
    }


app.state.sessions: dict[str, dict[str, str]] = {}
app.state.teams = _make_demo_data()


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
    if username == "admin" and password == "admin123":
        token = uuid.uuid4().hex
        app.state.sessions[token] = {"username": username}
        return {"token": token, "username": username}
    raise HTTPException(status_code=401, detail="ユーザー名またはパスワードが違います")


@app.get("/", response_class=HTMLResponse)
def root() -> str:
    return """
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>Teams Data Viewer</title>
        <style>
            :root {
                --bg: #f4f6fb;
                --card: #ffffff;
                --line: #dae1f0;
                --primary: #2b6ef5;
                --text: #1f2937;
                --muted: #6b7280;
                --success: #eaf8ee;
                --success-text: #166534;
            }
            * { box-sizing: border-box; }
            body {
                margin: 0; font-family: "Segoe UI", sans-serif; background: var(--bg); color: var(--text);
            }
            .container {
                max-width: 1120px; margin: 48px auto; padding: 0 20px;
            }
            .login-box, .dashboard {
                background: var(--card); border: 1px solid var(--line); border-radius: 18px; box-shadow: 0 12px 34px rgba(15, 23, 42, 0.08);
            }
            .login-box {
                max-width: 420px; margin: 0 auto; padding: 32px 28px;
            }
            .dashboard { display: none; padding: 24px; }
            .dashboard.visible { display: block; }
            h1 { margin-top: 0; font-size: 2rem; }
            .subtitle { color: var(--muted); margin-bottom: 20px; }
            label { display: block; margin: 16px 0 8px; font-weight: 600; }
            input, textarea, button {
                width: 100%; border-radius: 10px; border: 1px solid var(--line); padding: 12px 14px; font-size: 1rem;
            }
            input, textarea { background: #f9fafb; }
            button {
                margin-top: 18px; background: var(--primary); color: white; border: none; cursor: pointer; font-weight: 600;
            }
            .demo-note { margin-top: 12px; font-size: 0.86rem; color: var(--muted); }
            .topbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 18px; }
            .topbar .meta { color: var(--muted); }
            .layout { display: grid; grid-template-columns: minmax(260px, 1.1fr) minmax(320px, 1.8fr); gap: 20px; }
            table { width: 100%; border-collapse: collapse; background: #fff; }
            th, td { padding: 12px 10px; border-bottom: 1px solid var(--line); text-align: left; }
            tbody tr { cursor: pointer; }
            tbody tr:hover { background: #f7faff; }
            tbody tr.selected { background: #eef4ff; }
            .channel-list { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 8px; }
            .channel-item { background: #eef4ff; color: #1a4ebd; border-radius: 999px; padding: 8px 12px; cursor: pointer; }
            .channel-item.active { background: var(--primary); color: white; }
            .messages { margin-top: 18px; border: 1px solid var(--line); border-radius: 12px; padding: 12px; min-height: 160px; }
            .message { padding: 12px 0; border-bottom: 1px solid var(--line); }
            .message:last-child { border-bottom: none; }
            .message-head { font-size: 0.82rem; color: var(--muted); margin-bottom: 5px; }
            .empty { color: var(--muted); padding: 18px 0; }
            .success { background: var(--success); color: var(--success-text); border-radius: 8px; padding: 10px 12px; margin-bottom: 10px; display: none; }
            .success.show { display: block; }
        </style>
    </head>
    <body>
        <div class="container">
            <div id="loginBox" class="login-box">
                <h1>Teams Data Viewer</h1>
                <div class="subtitle">学校の Microsoft 365 環境向けに Teams データを確認するためのログイン画面です。</div>
                <form id="loginForm">
                    <label for="username">ユーザー名</label>
                    <input id="username" name="username" value="admin" required />
                    <label for="password">パスワード</label>
                    <input id="password" name="password" type="password" value="admin123" required />
                    <button type="submit">ログイン</button>
                </form>
                <div class="demo-note">デモ用: admin / admin123</div>
            </div>

            <div id="dashboard" class="dashboard">
                <div class="topbar">
                    <div>
                        <h1>Teams ダッシュボード</h1>
                        <div class="meta">ログイン中: <span id="currentUser"></span></div>
                    </div>
                    <button id="logoutButton" type="button" style="width:auto; background:#475569;">ログアウト</button>
                </div>

                <div id="messageBox" class="success"></div>

                <div class="layout">
                    <div>
                        <h2>チーム一覧</h2>
                        <table>
                            <thead>
                                <tr><th>チーム名</th><th>説明</th><th>チャネル数</th></tr>
                            </thead>
                            <tbody id="teamTableBody"></tbody>
                        </table>
                    </div>

                    <div>
                        <h2>チャネル</h2>
                        <div id="channelList" class="channel-list"></div>

                        <h2 style="margin-top: 24px;">メッセージ一覧</h2>
                        <div id="messages" class="messages"></div>

                        <form id="messageForm" style="margin-top: 18px;">
                            <label for="newMessage">新規投稿</label>
                            <textarea id="newMessage" rows="4" placeholder="投稿内容を入力"></textarea>
                            <button type="submit">投稿を追加</button>
                        </form>
                    </div>
                </div>
            </div>
        </div>

        <script>
            const state = {
                token: localStorage.getItem('teamsToken') || '',
                teams: [],
                selectedTeamId: null,
                selectedChannelId: null,
                user: localStorage.getItem('teamsUser') || ''
            };

            const loginBox = document.getElementById('loginBox');
            const dashboard = document.getElementById('dashboard');
            const teamTableBody = document.getElementById('teamTableBody');
            const channelList = document.getElementById('channelList');
            const messagesBox = document.getElementById('messages');
            const currentUser = document.getElementById('currentUser');
            const messageBox = document.getElementById('messageBox');

            function showMessage(text) {
                messageBox.textContent = text;
                messageBox.classList.add('show');
                setTimeout(() => messageBox.classList.remove('show'), 2500);
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

            async function loadTeams() {
                const response = await authFetch('/api/teams');
                if (!response.ok) {
                    throw new Error('チーム一覧の取得に失敗しました');
                }
                const data = await response.json();
                state.teams = data.teams || [];
                if (!state.teams.length) {
                    teamTableBody.innerHTML = '<tr><td colspan="3">チームがありません</td></tr>';
                    return;
                }
                if (!state.selectedTeamId) {
                    state.selectedTeamId = state.teams[0].id;
                }
                teamTableBody.innerHTML = state.teams.map(team => {
                    const selected = team.id === state.selectedTeamId ? 'selected' : '';
                    return `<tr data-team-id="${team.id}" class="${selected}">
                        <td>${team.name}</td>
                        <td>${team.description || ''}</td>
                        <td>${(team.channels || []).length}</td>
                    </tr>`;
                }).join('');

                teamTableBody.querySelectorAll('tr[data-team-id]').forEach(row => {
                    row.addEventListener('click', () => {
                        state.selectedTeamId = row.dataset.teamId;
                        loadTeams();
                        loadChannels();
                    });
                });

                await loadChannels();
            }

            async function loadChannels() {
                if (!state.selectedTeamId) return;
                const team = state.teams.find(item => item.id === state.selectedTeamId);
                if (!team) return;
                const channels = team.channels || [];
                if (!channels.length) {
                    channelList.innerHTML = '<span class="empty">チャネルがありません</span>';
                    messagesBox.innerHTML = '<div class="empty">メッセージはまだありません</div>';
                    return;
                }
                if (!state.selectedChannelId || !channels.some(ch => ch.id === state.selectedChannelId)) {
                    state.selectedChannelId = channels[0].id;
                }
                channelList.innerHTML = channels.map(channel => `
                    <button type="button" class="channel-item ${channel.id === state.selectedChannelId ? 'active' : ''}" data-channel-id="${channel.id}">${channel.name}</button>
                `).join('');
                channelList.querySelectorAll('.channel-item').forEach(button => {
                    button.addEventListener('click', () => {
                        state.selectedChannelId = button.dataset.channelId;
                        loadChannels();
                    });
                });
                await loadMessages();
            }

            async function loadMessages() {
                if (!state.selectedTeamId || !state.selectedChannelId) return;
                const response = await authFetch(`/api/teams/${state.selectedTeamId}/channels/${state.selectedChannelId}/messages`);
                if (!response.ok) {
                    messagesBox.innerHTML = '<div class="empty">メッセージを取得できませんでした。</div>';
                    return;
                }
                const data = await response.json();
                const messages = data.messages || [];
                if (!messages.length) {
                    messagesBox.innerHTML = '<div class="empty">メッセージはまだありません。</div>';
                    return;
                }
                messagesBox.innerHTML = messages.map(msg => `
                    <div class="message">
                        <div class="message-head">${msg.author}・${new Date(msg.timestamp).toLocaleString('ja-JP')}</div>
                        <div>${(msg.content || '').replace(/\n/g, '<br>')}</div>
                    </div>
                `).join('');
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
                await loadTeams();
            });

            document.getElementById('logoutButton').addEventListener('click', () => {
                state.token = '';
                state.user = '';
                localStorage.removeItem('teamsToken');
                localStorage.removeItem('teamsUser');
                setAuthUI();
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
                await loadMessages();
            });

            if (state.token) {
                setAuthUI();
                loadTeams().catch(() => {
                    alert('チーム一覧の取得に失敗しました');
                });
            } else {
                setAuthUI();
            }
        </script>
    </body>
    </html>
    """


@app.get("/api/teams")
def get_joined_teams(authorization: str | None = Header(default=None, alias="Authorization")) -> dict[str, Any]:
    require_session(authorization)
    teams = []
    for team_id, team in app.state.teams.items():
        teams.append(
            {
                "id": team_id,
                "name": team["name"],
                "description": team["description"],
                "channels": [
                    {"id": channel["id"], "name": channel["name"]} for channel in team["channels"]
                ],
            }
        )
    return {"teams": teams}


@app.get("/api/teams/{team_id}/channels")
def get_team_channels(team_id: str, authorization: str | None = Header(default=None, alias="Authorization")) -> dict[str, Any]:
    require_session(authorization)
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
