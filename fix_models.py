import os

project_dir = "."
target_strings = ["llama3.2-vision:latest", "llama3.2-vision"]

for root, dirs, files in os.walk(project_dir):
    if ".venv" in root or "__pycache__" in root or ".git" in root:
        continue
    for file in files:
        if file.endswith(".py") and file != "fix_models.py":
            filepath = os.path.join(root, file)
            try:
                with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                
                updated = False
                for target in target_strings:
                    if target in content:
                        content = content.replace(target, "llava")
                        updated = True
                
                if updated:
                    with open(filepath, "w", encoding="utf-8") as f:
                        f.write(content)
                    print(f"Updated vision model in: {filepath}")
            except Exception as e:
                print(f"Skipped {filepath} due to error: {e}")

print("Replacement complete!")