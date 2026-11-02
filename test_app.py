from fastapi.testclient import TestClient

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
