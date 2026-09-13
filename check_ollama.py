"""
Run this BEFORE starting app.py:  python check_ollama.py
It tells you in plain language whether Ollama is running, which models
are installed, and how long a real generation call takes on your machine.
"""
import requests
import time

BASE = "http://localhost:11434"

print("1) Checking if Ollama is reachable at localhost:11434 ...")
try:
    r = requests.get(f"{BASE}/api/tags", timeout=5)
    r.raise_for_status()
    models = [m["name"] for m in r.json().get("models", [])]
    print("   OK — Ollama is running.")
except Exception as e:
    print(f"   FAILED: {e}")
    print("   Fix: open a terminal and run 'ollama serve', then re-run this script.")
    raise SystemExit(1)

if not models:
    print("2) No models installed.")
    print("   Fix: install at least llama3.2:3b, qwen2.5-coder:1.5b and llava:7b for the prototype.")
    raise SystemExit(1)

print(f"2) Installed models: {models}")

test_model = models[0]
print(f"3) Timing a real generation call on '{test_model}' (this is the actual speed you'll see per page)...")
start = time.time()
try:
    r = requests.post(
        f"{BASE}/api/generate",
        json={"model": test_model, "prompt": "Say 'ready' and nothing else.", "stream": False},
        timeout=180,
    )
    elapsed = time.time() - start
    r.raise_for_status()
    print(f"   OK — response in {elapsed:.1f}s: {r.json().get('response', '').strip()[:100]}")
    if elapsed > 30:
        print("   NOTE: larger vision models can be slow on 4GB VRAM; the prototype prefers llama3.2:3b for text and llava:7b for images.")
        print("   Check 'ollama ps' / nvidia-smi to confirm GPU acceleration.")
except Exception as e:
    print(f"   FAILED: {e}")