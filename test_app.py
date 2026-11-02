from fastapi.testclient import TestClient
from unittest.mock import patch

import main

client = TestClient(main.app)


def test_root_returns_html():
    response = client.get("/")
    assert response.status_code == 200
    assert "Teams" in response.text
    assert "ログイン" in response.text


def test_health_endpoint():
    response = client.get("/health")
    assert response.status_code == 200
    payload = response.json()
    assert payload["status"] == "ok"


def test_graph_snapshot_maps_team_channel_and_message_data():
    graph_responses = [
        [{"id": "team-graph", "displayName": "Graph Team", "description": "Updates"}],
        [{"id": "channel-graph", "displayName": "General"}],
        [
            {
                "id": "message-graph",
                "from": {"user": {"displayName": "Aki"}},
                "body": {"content": "<p>First &amp; line</p><p>Second</p>"},
                "createdDateTime": "2026-11-02T10:00:00Z",
            }
        ],
    ]
    with patch("main.get_graph_access_token", return_value="test-token"), patch(
        "main.graph_get_collection", side_effect=graph_responses
    ):
        snapshot = main.get_graph_snapshot()

    assert snapshot["source"] == "graph"
    assert snapshot["teams"][0]["channels"][0] == {"id": "channel-graph", "name": "General"}
    message = snapshot["messages"][0]
    assert message["content"] == "First & line\nSecond"
    assert message["author"] == "Aki"
    assert message["teamName"] == "Graph Team"
    assert message["channelName"] == "General"


def test_graph_login_requires_configured_app_credentials():
    with patch("main.graph_is_configured", return_value=True), patch.dict(
        "os.environ", {"APP_USERNAME": "reader", "APP_PASSWORD": "test-secret"}
    ):
        assert client.post("/api/login", json={"username": "admin", "password": "admin123"}).status_code == 401
        response = client.post("/api/login", json={"username": "reader", "password": "test-secret"})

    assert response.status_code == 200


def test_dashboard_endpoint_uses_graph_snapshot_when_configured():
    token = "graph-dashboard-test-session"
    main.app.state.sessions[token] = {"username": "reader"}
    snapshot = {"teams": [{"id": "team-graph"}], "messages": [{"id": "message-graph"}], "source": "graph"}

    with patch("main.graph_is_configured", return_value=True), patch(
        "main.get_graph_snapshot", return_value=snapshot
    ):
        response = client.get("/api/dashboard", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    assert response.json() == snapshot


def test_all_messages_requires_auth_and_returns_team_channel_context():
    assert client.get("/api/messages").status_code == 401

    login_response = client.post("/api/login", json={"username": "admin", "password": "admin123"})
    token = login_response.json()["token"]
    response = client.get("/api/messages", headers={"Authorization": f"Bearer {token}"})

    assert response.status_code == 200
    messages = response.json()["messages"]
    assert len(messages) >= 4
    assert all({"teamId", "teamName", "channelId", "channelName"} <= message.keys() for message in messages)
    timestamps = [message["timestamp"] for message in messages]
    assert timestamps == sorted(timestamps, reverse=True)


def test_login_and_teams_flow():
    login_response = client.post("/api/login", json={"username": "admin", "password": "admin123"})
    assert login_response.status_code == 200
    token = login_response.json()["token"]

    teams_response = client.get("/api/teams", headers={"Authorization": f"Bearer {token}"})
    assert teams_response.status_code == 200
    payload = teams_response.json()
    assert "teams" in payload
    assert len(payload["teams"]) >= 1

    team_id = payload["teams"][0]["id"]
    channel_response = client.get(
        f"/api/teams/{team_id}/channels",
        headers={"Authorization": f"Bearer {token}"},
    )
    assert channel_response.status_code == 200
    channels = channel_response.json()["channels"]
    assert len(channels) >= 1

    channel_id = channels[0]["id"]
    post_response = client.post(
        f"/api/teams/{team_id}/channels/{channel_id}/messages",
        headers={"Authorization": f"Bearer {token}"},
        json={"author": "テスト生", "content": "テスト投稿です"},
    )
    assert post_response.status_code == 200

    messages_response = client.get(
        f"/api/teams/{team_id}/channels/{channel_id}/messages",
        headers={"Authorization": f"Bearer {token}"},
    )
    assert messages_response.status_code == 200
    messages = messages_response.json()["messages"]
    assert any("テスト投稿です" in msg["content"] for msg in messages)
