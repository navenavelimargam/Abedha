"""Standalone local model router for the prototype."""
import requests

OLLAMA_URL = "http://localhost:11434/api/generate"
TAGS_URL = "http://localhost:11434/api/tags"
TASK_MODEL_MAP = {
    "Finance": ["llama3.2:3b"],
    "Engineering": ["llama3.2:3b"],
    "Coding": ["qwen2.5-coder:1.5b"],
    "General": ["llama3.2:3b"],
}
VISION_MODELS = ["llava:7b", "llava"]

def _available_models():
    try:
        r=requests.get(TAGS_URL,timeout=3); r.raise_for_status()
        return [m["name"] for m in r.json().get("models",[])]
    except Exception: return []

def auto_router(user_prompt, selected_sector="General", vision=False):
    available=_available_models()
    if not available: return "Server Error: Ollama is unreachable or has no models installed."
    if vision:
        model=next((m for m in VISION_MODELS if m in available),None)
        if not model: return "Server Error: no local vision model is installed."
    else:
        q=(user_prompt or '').lower()
        task="Coding" if any(w in q for w in ["code","python","sql","debug","script"]) else selected_sector
        prefs=TASK_MODEL_MAP.get(task,TASK_MODEL_MAP["General"])
        model=next((m for m in prefs if m in available),available[0])
    print(f"[ROUTER] sector={selected_sector} model={model}")
    try:
        r=requests.post(OLLAMA_URL,json={"model":model,"prompt":user_prompt,"stream":False},timeout=180)
        r.raise_for_status(); return r.json().get("response","")
    except Exception as e: return f"Server Error: {e}"
