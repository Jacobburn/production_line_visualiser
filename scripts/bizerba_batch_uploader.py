# pip install requests
import json
import os
import time
from datetime import datetime, timezone

import requests

FUNCTION_URL = os.getenv(
    "BIZERBA_FUNCTION_URL",
    "https://<project-ref>.supabase.co/functions/v1/ingest-bizerba",
)
DEVICE_KEY = os.getenv("BIZERBA_DEVICE_KEY", "<device-key>")
BATCH_SIZE = int(os.getenv("BIZERBA_BATCH_SIZE", "100000"))
MAX_RETRIES = int(os.getenv("BIZERBA_MAX_RETRIES", "5"))
REQUEST_TIMEOUT_SECONDS = int(os.getenv("BIZERBA_REQUEST_TIMEOUT_SECONDS", "120"))
FAILED_QUEUE_FILE = os.getenv("BIZERBA_FAILED_QUEUE_FILE", "bizerba_failed_queue.ndjson")


def read_bizerba_batch(limit):
    # Replace this with the vendor's real source reader.
    # Return up to `limit` rows each time this function is called.
    # Example row shape:
    # return [
    #     {
    #         "id": 123456,
    #         "actual_net_weight_value": 0.503,
    #         "device_name": "Bizerba-01",
    #         "article_number": "427502",
    #         "date": "2026-03-13",
    #         "timestamp": "2026-03-13T02:15:00Z",
    #     }
    # ][:limit]
    return []


def build_payload(raw):
    now = datetime.now(timezone.utc)
    raw_date = raw.get("date")
    raw_timestamp = raw.get("timestamp")
    payload = {
        "id": str(raw["id"]),
        "ActualNetWeightValue": float(raw["actual_net_weight_value"]),
        "DeviceName": str(raw["device_name"]),
        "Date": str(raw_date) if raw_date else now.strftime("%Y-%m-%d"),
        "Timestamp": (
            raw_timestamp.astimezone(timezone.utc).isoformat()
            if isinstance(raw_timestamp, datetime)
            else str(raw_timestamp)
            if raw_timestamp
            else now.isoformat()
        ),
        "ArticleNumber": str(raw["article_number"]),
    }

    return payload


def push_payloads(payloads, max_retries=MAX_RETRIES):
    headers = {
        "Content-Type": "application/json",
        "x-device-key": DEVICE_KEY,
    }

    body = {"rows": payloads}
    for attempt in range(max_retries):
        try:
            resp = requests.post(
                FUNCTION_URL,
                headers=headers,
                json=body,
                timeout=REQUEST_TIMEOUT_SECONDS,
            )
            if 200 <= resp.status_code < 300:
                return True
            print(f"Batch push failed: HTTP {resp.status_code} - {resp.text}")
        except requests.RequestException as exc:
            print(f"Batch push error: {exc}")

        time.sleep(2 ** attempt)

    return False


def queue_failed(payloads):
    with open(FAILED_QUEUE_FILE, "a", encoding="utf-8") as file_handle:
        for payload in payloads:
            file_handle.write(json.dumps(payload) + "\n")


def main():
    if "<project-ref>" in FUNCTION_URL or DEVICE_KEY == "<device-key>":
        raise RuntimeError("Set BIZERBA_FUNCTION_URL and BIZERBA_DEVICE_KEY first.")

    if BATCH_SIZE <= 0:
        raise RuntimeError("BIZERBA_BATCH_SIZE must be greater than zero.")

    while True:
        raw_batch = read_bizerba_batch(BATCH_SIZE)
        if not raw_batch:
            print("No more Bizerba rows to send.")
            break

        payloads = [build_payload(raw) for raw in raw_batch]
        if not push_payloads(payloads):
            queue_failed(payloads)
            continue

        print(f"Sent {len(payloads)} Bizerba rows.")


if __name__ == "__main__":
    main()
