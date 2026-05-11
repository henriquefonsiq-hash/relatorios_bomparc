import json
import os


APP_DIR = os.path.dirname(os.path.abspath(__file__))
CRITERIA_PATH = os.path.join(APP_DIR, "acceptance_criteria.json")
DEFAULT_CRITERIA = {
    "Cinta Sling": "ABNT NBR 15637-1",
    "Cinta Circular": "ABNT NBR 15637-2",
}


def load_acceptance_criteria():
    if not os.path.exists(CRITERIA_PATH):
        save_acceptance_criteria(DEFAULT_CRITERIA)
        return DEFAULT_CRITERIA.copy()

    try:
        with open(CRITERIA_PATH, "r", encoding="utf-8") as criteria_file:
            criteria = json.load(criteria_file)
    except (OSError, json.JSONDecodeError):
        return DEFAULT_CRITERIA.copy()

    return {
        str(material).strip(): str(criterion).strip()
        for material, criterion in criteria.items()
        if str(material).strip() and str(criterion).strip()
    }


def save_acceptance_criteria(criteria):
    cleaned_criteria = {
        str(material).strip(): str(criterion).strip()
        for material, criterion in criteria.items()
        if str(material).strip() and str(criterion).strip()
    }
    with open(CRITERIA_PATH, "w", encoding="utf-8") as criteria_file:
        json.dump(cleaned_criteria, criteria_file, ensure_ascii=False, indent=2)
