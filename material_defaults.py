import json
import os


APP_DIR = os.path.dirname(os.path.abspath(__file__))
MATERIALS_PATH = os.path.join(APP_DIR, "material_defaults.json")
DEFAULT_MATERIALS = {
    "Cinta Sling": "POLIESTER",
    "Cinta Circular": "POLIESTER",
}


def load_material_defaults():
    if not os.path.exists(MATERIALS_PATH):
        save_material_defaults(DEFAULT_MATERIALS)
        return DEFAULT_MATERIALS.copy()

    try:
        with open(MATERIALS_PATH, "r", encoding="utf-8") as materials_file:
            materials = json.load(materials_file)
    except (OSError, json.JSONDecodeError):
        return DEFAULT_MATERIALS.copy()

    return {
        str(equipment).strip(): str(raw_material).strip()
        for equipment, raw_material in materials.items()
        if str(equipment).strip() and str(raw_material).strip()
    }


def save_material_defaults(materials):
    cleaned_materials = {
        str(equipment).strip(): str(raw_material).strip()
        for equipment, raw_material in materials.items()
        if str(equipment).strip() and str(raw_material).strip()
    }
    with open(MATERIALS_PATH, "w", encoding="utf-8") as materials_file:
        json.dump(cleaned_materials, materials_file, ensure_ascii=False, indent=2)
