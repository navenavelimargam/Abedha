"""
Run this once, and again any time you add new files to knowledge_base/:

    python build_kb.py

It builds kb_index.pkl, which app.py uses automatically to ground answers
in your organization's own SOPs, manuals, and past correspondence.
"""
from kb_index import build_index

if __name__ == "__main__":
    build_index()