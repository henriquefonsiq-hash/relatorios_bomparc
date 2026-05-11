import json
import os


APP_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENT_REGISTRY_PATH = os.path.join(APP_DIR, "client_registry.json")
DEFAULT_CLIENT_REGISTRY = {
    "DOF SUBSEA BRASIL SERVIÇOS LTDA": {
        "SKANDI CHIEFTAIN": "PORTO DO AÇU",
    },
}


def _normalize_registry(registry):
    normalized = {}
    for client, vessels in registry.items():
        client = str(client).strip()
        if not client or not isinstance(vessels, dict):
            continue

        cleaned_vessels = {
            str(vessel).strip(): str(address).strip()
            for vessel, address in vessels.items()
            if str(vessel).strip() and str(address).strip()
        }
        if cleaned_vessels:
            normalized[client] = cleaned_vessels
    return normalized


def load_client_registry():
    if not os.path.exists(CLIENT_REGISTRY_PATH):
        save_client_registry(DEFAULT_CLIENT_REGISTRY)
        return DEFAULT_CLIENT_REGISTRY.copy()

    try:
        with open(CLIENT_REGISTRY_PATH, "r", encoding="utf-8") as registry_file:
            registry = json.load(registry_file)
    except (OSError, json.JSONDecodeError):
        return DEFAULT_CLIENT_REGISTRY.copy()

    return _normalize_registry(registry)


def save_client_registry(registry):
    normalized = _normalize_registry(registry)
    with open(CLIENT_REGISTRY_PATH, "w", encoding="utf-8") as registry_file:
        json.dump(normalized, registry_file, ensure_ascii=False, indent=2)
