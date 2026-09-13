import requests
import docx
from app import resolve_model_for_task
import os

def generate_and_save_doc(user_prompt: str, filename: str = "Approval_Note.docx"):
    print("Generating content via local Qwen model...")
    url = "http://localhost:11434/api/generate"
    model, err = resolve_model_for_task("general")
    if err:
        raise RuntimeError(err)
    payload = {
        "model": model,
        "prompt": user_prompt,
        "stream": False
    }
    
    try:
        response = requests.post(url, json=payload).json()
        ai_text = response.get("response", "Error generating text.")
    except Exception as e:
        ai_text = f"Error connecting to local Ollama: {e}"

    print("Formatting document...")
    doc = docx.Document()
    doc.add_heading("CONFIDENTIAL - CORPORATE PROTOTYPE", level=1)
    doc.add_paragraph(ai_text)
    
    doc.save(filename)
    print(f"Success! Saved to {os.path.abspath(filename)}")

if __name__ == "__main__":
    generate_and_save_doc("Draft a short approval note for the Unit 4 boiler inspection. State that pressure levels are normal.")