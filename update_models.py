import os

project_dir = "."  # or your specific project path
target_strings = ["llava", "llava"]

for root, dirs, files in os.walk(project_dir):
    if ".venv" in root or "__pycache__" in root:
        continue
    for file in files:
        if file.endswith(".py"):
            filepath = os.path.join(root, file)
            with open(filepath, "r", encoding="utf-8") as f:
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

print("Replacement complete!")