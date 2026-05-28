import json
import os

CONFIG_FILE = "config.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                # Ensure new keys exist with defaults
                cfg.setdefault("use_offline", False)
                cfg.setdefault("tesseract_path", "")
                return cfg
        except Exception as e:
            print(f"Error loading config: {e}")
            return {"use_offline": False, "tesseract_path": ""}
    return {"use_offline": False, "tesseract_path": ""}

def save_config(api_key="", api_base_url="", model_name="", image_folder="", excel_file="", thread_count=4, use_offline=False, tesseract_path=""):
    config_data = {
        "api_key": api_key,
        "api_base_url": api_base_url,
        "model_name": model_name,
        "image_folder": image_folder,
        "excel_file": excel_file,
        "thread_count": thread_count,
        "use_offline": use_offline,
        "tesseract_path": tesseract_path
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error saving config: {e}")
