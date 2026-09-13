"""
Lightweight local knowledge base for grounding answers in the organisation's
own manuals, SOPs and past correspondence.

Uses TF-IDF (scikit-learn) instead of embeddings, so there is no extra
model to download the night before your demo — everything builds and
runs fully offline with one extra pip package.

Setup (do this once):
    1. pip install scikit-learn
    2. mkdir knowledge_base   (or just create the folder in Explorer)
    3. Drop SOP/manual/correspondence files into knowledge_base/
       (.pdf, .docx, .txt, .md all work)
    4. python build_kb.py     -> builds kb_index.pkl

app.py imports search_kb() from this file automatically. If you haven't
built an index yet, search_kb() just returns [] and the app works exactly
as before (no grounding, but no crash either).
"""
import os
import pickle
import re

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

KB_DIR = "knowledge_base"
INDEX_PATH = "kb_index.pkl"
CHUNK_SIZE = 800
CHUNK_OVERLAP = 100

_state = {"vectorizer": None, "matrix": None, "chunks": [], "sources": []}


def _chunk_text(text, source):
    chunks = []
    text = re.sub(r'\s+', ' ', text).strip()
    start = 0
    while start < len(text):
        end = start + CHUNK_SIZE
        chunk = text[start:end]
        if chunk.strip():
            chunks.append((chunk, source))
        start += CHUNK_SIZE - CHUNK_OVERLAP
    return chunks


def build_index():
    """Reads every file in knowledge_base/, chunks it, fits a TF-IDF index,
    and saves it to disk so app.py doesn't rebuild it on every restart."""
    # Imported here (not at module level) to avoid a circular import with
    # app.py, which itself imports search_kb from this file.
    from app import extract_document_text

    if not os.path.isdir(KB_DIR):
        os.makedirs(KB_DIR, exist_ok=True)
        print(f"Created empty '{KB_DIR}/' folder — add SOPs/manuals there and re-run this script.")
        return

    all_chunks = []
    for filename in os.listdir(KB_DIR):
        path = os.path.join(KB_DIR, filename)
        if not os.path.isfile(path):
            continue
        try:
            text, is_vision, _ = extract_document_text(path, filename)
        except Exception as exc:
            print(f"Skipping {filename}: {exc}")
            continue
        if is_vision or not text.strip():
            continue
        all_chunks.extend(_chunk_text(text, filename))

    if not all_chunks:
        print("No readable documents found in knowledge_base/. Index not built.")
        return

    texts = [c[0] for c in all_chunks]
    sources = [c[1] for c in all_chunks]

    vectorizer = TfidfVectorizer(stop_words='english', max_features=20000)
    matrix = vectorizer.fit_transform(texts)

    with open(INDEX_PATH, 'wb') as f:
        pickle.dump({
            "vectorizer": vectorizer,
            "matrix": matrix,
            "chunks": texts,
            "sources": sources,
        }, f)

    print(f"Indexed {len(texts)} chunk(s) from {len(set(sources))} document(s) into {INDEX_PATH}")


def _load_index():
    if _state["vectorizer"] is not None:
        return
    if not os.path.exists(INDEX_PATH):
        return
    with open(INDEX_PATH, 'rb') as f:
        data = pickle.load(f)
    _state.update(data)


def search_kb(query, top_k=3, min_score=0.05, sector=None):
    """Returns [(chunk_text, source_filename, score), ...] for the most
    relevant knowledge-base chunks, or [] if no index exists yet / no match."""
    _load_index()
    if _state["vectorizer"] is None:
        return []

    query_vec = _state["vectorizer"].transform([query])
    scores = cosine_similarity(query_vec, _state["matrix"]).flatten()
    ranked = sorted(range(len(scores)), key=lambda i: scores[i], reverse=True)

    results = []
    for i in ranked:
        if sector and sector not in {"General", "All"}:
            source = str(_state["sources"][i]).lower()
            sector_key = str(sector).lower().replace(" ", "_")
            if not (source.startswith(sector_key + "__") or sector.lower() in source or source.startswith("general__")):
                continue
        if len(results) >= top_k:
            break
        if scores[i] >= min_score:
            results.append((_state["chunks"][i], _state["sources"][i], float(scores[i])))
    return results