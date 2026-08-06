import os, httpx

key = os.environ["OPENROUTER_API_KEY"]
test_models = [
    "google/gemma-4-31b-it:free",
    "qwen/qwen3-next-80b-a3b-instruct:free",
    "openai/gpt-oss-20b:free",
    "deepseek/deepseek-v4-flash:free",
    "nousresearch/hermes-3-llama-3.1-405b:free",
]
msg = "Vertaal naar Nederlands: Zdrave krmivo pro psy. Geef alleen de vertaling."
for model in test_models:
    resp = httpx.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        json={"model": model, "max_tokens": 50, "messages": [{"role": "user", "content": msg}]},
        timeout=30,
    )
    data = resp.json()
    if resp.status_code == 200 and "choices" in data:
        print(f"OK  {model}: {data['choices'][0]['message']['content'][:70]}")
    else:
        print(f"{resp.status_code} {model}: {str(data)[:120]}")
