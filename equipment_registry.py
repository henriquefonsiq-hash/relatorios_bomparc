import json
import os

from acceptance_criteria import load_acceptance_criteria
from material_defaults import load_material_defaults


APP_DIR = os.path.dirname(os.path.abspath(__file__))
EQUIPMENT_REGISTRY_PATH = os.path.join(APP_DIR, "equipment_registry.json")
DEFAULT_EQUIPMENT_REGISTRY = {
    "Cinta Sling": {
        "materia_prima": "POLIESTER",
        "criterio_aceitacao": "ABNT NBR 15637-1",
    },
    "Cinta Circular": {
        "materia_prima": "POLIESTER",
        "criterio_aceitacao": "ABNT NBR 15637-2",
    },
}


def _normalize_registry(registry):
    normalized = {}
    for equipment, config in registry.items():
        equipment = str(equipment).strip()
        if not equipment or not isinstance(config, dict):
            continue

        raw_material = str(config.get("materia_prima", "")).strip()
        acceptance_criterion = str(config.get("criterio_aceitacao", "")).strip()
        if not raw_material or not acceptance_criterion:
            continue

        normalized[equipment] = {
            "materia_prima": raw_material,
            "criterio_aceitacao": acceptance_criterion,
        }
    return normalized


def _legacy_registry():
    criteria = load_acceptance_criteria()
    materials = load_material_defaults()
    registry = {}
    for equipment in sorted(set(criteria) | set(materials)):
        registry[equipment] = {
            "materia_prima": materials.get(equipment, ""),
            "criterio_aceitacao": criteria.get(equipment, ""),
        }
    return _normalize_registry(registry)


def load_equipment_registry():
    if not os.path.exists(EQUIPMENT_REGISTRY_PATH):
        registry = _legacy_registry() or DEFAULT_EQUIPMENT_REGISTRY.copy()
        save_equipment_registry(registry)
        return registry

    try:
        with open(EQUIPMENT_REGISTRY_PATH, "r", encoding="utf-8") as registry_file:
            registry = json.load(registry_file)
    except (OSError, json.JSONDecodeError):
        return DEFAULT_EQUIPMENT_REGISTRY.copy()

    return _normalize_registry(registry)


def save_equipment_registry(registry):
    normalized = _normalize_registry(registry)
    with open(EQUIPMENT_REGISTRY_PATH, "w", encoding="utf-8") as registry_file:
        json.dump(normalized, registry_file, ensure_ascii=False, indent=2)
