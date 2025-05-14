import requests
import json
import re
import time
import csv
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

# SlackのBot Token
SLACK_TOKEN = "（ここにSlackのBot Tokenを入力）"

# 対象のSlackチャンネルURLリスト
CHANNEL_URLS = [
    "(ここにSlackのチャンネルURLを入力)",
    "(ここにSlackのチャンネルURLを入力)"
]

MAX_WORKERS = 5  # 並列取得数（Slackに優しく）

def extract_channel_id(url):
    match = re.search(r"/(C[A-Z0-9]+)", url)
    return match.group(1) if match else None

def get_channel_info(channel_id):
    url = "https://slack.com/api/conversations.info"
    headers = {"Authorization": f"Bearer {SLACK_TOKEN}"}
    params = {"channel": channel_id}
    try:
        response = requests.get(url, headers=headers, params=params, timeout=15)
        data = response.json()
        if data.get("ok"):
            return data["channel"]
        else:
            print(f"チャンネル情報取得エラー: {data}")
    except Exception as e:
        print(f"❌ チャンネル情報取得失敗: {e}")
    return None

def fetch_all_messages(channel_id):
    url = "https://slack.com/api/conversations.history"
    headers = {"Authorization": f"Bearer {SLACK_TOKEN}"}
    params = {"channel": channel_id, "limit": 200}
    all_messages = []

    while True:
        try:
            response = requests.get(url, headers=headers, params=params, timeout=15)
            data = response.json()
            if not data.get("ok"):
                print(f"メッセージ取得エラー: {data}")
                break
            all_messages.extend(data["messages"])
            if not data.get("has_more"):
                break
            params["cursor"] = data["response_metadata"]["next_cursor"]
            time.sleep(1)
        except Exception as e:
            print(f"❌ メッセージ取得失敗: {e}")
            break

    return all_messages

def fetch_thread_replies(channel_id, thread_ts):
    url = "https://slack.com/api/conversations.replies"
    headers = {"Authorization": f"Bearer {SLACK_TOKEN}"}
    params = {"channel": channel_id, "ts": thread_ts}

    while True:
        try:
            response = requests.get(url, headers=headers, params=params, timeout=15)
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", 20))
                print(f"⚠️ レート制限。{retry_after}秒待機（スレッドTS: {thread_ts}）")
                time.sleep(retry_after)
                continue
            data = response.json()
            if data.get("ok"):
                return {
                    "thread_ts": thread_ts,
                    "messages": data["messages"]
                }
            else:
                print(f"スレッド取得エラー（{thread_ts}）: {data}")
                return {
                    "thread_ts": thread_ts,
                    "messages": []
                }
        except Exception as e:
            print(f"❌ スレッド取得失敗（{thread_ts}）: {e}")
            return {
                "thread_ts": thread_ts,
                "messages": []
            }

def structure_threads(messages, channel_id):
    thread_ts_list = []
    seen = set()

    for msg in messages:
        if "thread_ts" not in msg or msg["ts"] == msg["thread_ts"]:
            ts = msg["ts"]
            if ts not in seen:
                seen.add(ts)
                thread_ts_list.append(ts)

    threads = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(fetch_thread_replies, channel_id, ts): ts for ts in thread_ts_list}
        for future in tqdm(as_completed(futures), total=len(futures), desc="スレッド取得中", ncols=80):
            try:
                result = future.result()
                threads.append(result)
            except Exception as e:
                print(f"❌ スレッド処理中にエラー: {e}")
    return threads

def save_to_json(data, filename):
    try:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"❌ JSON保存エラー: {e}")

def save_to_csv(threads, filename):
    try:
        with open(filename, mode="w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["thread_ts", "ts", "user", "text"])
            for thread in threads:
                thread_ts = thread.get("thread_ts")
                for msg in thread.get("messages", []):
                    writer.writerow([
                        thread_ts,
                        msg.get("ts", ""),
                        msg.get("user", ""),
                        msg.get("text", "").replace("\n", " ").strip()
                    ])
    except Exception as e:
        print(f"❌ CSV保存エラー: {e}")

# メイン処理
for url in CHANNEL_URLS:
    channel_id = extract_channel_id(url)
    if not channel_id:
        print(f"URLからチャンネルIDが抽出できませんでした: {url}")
        continue

    channel_info = get_channel_info(channel_id)
    if not channel_info:
        continue

    channel_name = channel_info["name"]
    print(f"📥 {channel_name}（ID: {channel_id}）の全メッセージを取得中...")

    all_messages = fetch_all_messages(channel_id)
    structured_threads = structure_threads(all_messages, channel_id)

    filename_json = f"slack_threads_{channel_name}.json"
    filename_csv = f"slack_threads_{channel_name}.csv"
    save_to_json(structured_threads, filename_json)
    save_to_csv(structured_threads, filename_csv)

    print(f"✅ 保存完了！ → {filename_json} / {filename_csv}\n")
