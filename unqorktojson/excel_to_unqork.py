import argparse
import json
import uuid
from pathlib import Path
import pandas as pd


def s(x) -> str:
    """Safe string trim."""
    return "" if pd.isna(x) else str(x).strip()


def guid() -> str:
    """UUID for componentId."""
    return str(uuid.uuid4())


def hashkey() -> str:
    """Cosmetic $$hashKey like 'object:12345'."""
    return f"object:{uuid.uuid4().int % 100000}"


# ---------- key normalization helpers ----------

def _replace_underscores(text: str) -> str:
    return text.replace("_", "-")


def _prepend_a_if_digit(text: str) -> str:
    if text and text[0].isdigit():
        return "a" + text
    return text


def normalize_whole_key(text: str) -> str:
    """
    Apply rules to the FULL key string:
      - replace '_' with '-'
      - if resulting string starts with a digit, prepend 'a'
    """
    text = _replace_underscores(text)
    text = _prepend_a_if_digit(text)
    return text


def make_key(*parts: str) -> str:
    """
    Build a key by joining parts with '-' and then normalizing the FULL result.
    This guarantees NO underscores anywhere (including suffix),
    and prepends 'a' if the final key starts with a digit.
    """
    tokens = [ _replace_underscores(p) for p in parts if p ]
    joined = "-".join(tokens)
    return normalize_whole_key(joined)


# ---------- component builder ----------

def make_components(module_path: str,
                    module_id_value: str,
                    component_keys: list[str]) -> list[dict]:
    """
    Build the 3 Unqork components for a module.

    Keys (hidden/button/integrator/clickTrigger/output/linked.outputs & each component key):
      - '_' -> '-'
      - prepend 'a' if *final* key starts with a digit
    Labels remain unprefixed.
    dataInputObject begins with the two constants (transformName/moduleId) using original values.
    """
    # Final keys (entire string normalized, including suffixes)
    key_hidden = make_key(module_path, "transformData")     # e.g., a990119-transformData
    key_button = make_key(module_path, "btnSubmit")         # e.g., a990119-btnSubmit
    key_integrator = make_key(module_path, "plugTransform") # e.g., a990119-plugTransform

    # Output id follows same key rule
    output_target_id = key_hidden

    # 1) hidden
    hidden = {
        "$$hashKey": hashkey(),
        "ancestors": ["root"],
        "componentId": guid(),
        "key": key_hidden,
        "label": "transformData",
        "permissions": {},
        "persistent": False,
        "react": False,
        "type": "hidden",
        "watcherList": [],
    }

    # 2) button
    btn = {
        "$$hashKey": hashkey(),
        "action": "event",
        "ancestors": ["root"],
        "ariaLive": "off",
        "block": False,
        "cancelButton": "Cancel",
        "clickTrigger": key_integrator,  # fully normalized
        "componentId": guid(),
        "confirmationButton": "OK",
        "disableOnInvalid": False,
        "doNotShowValidatedPopup": True,
        "elementType": "module",
        "errorMessage": "Please check your module and fix fields outlined in red",
        "errorTitle": "Oops - {{ errors }} error(s) have been found!",
        "input": True,
        "key": key_button,
        "label": "Submit",
        "leftIcon": "",
        "oneClickOnly": False,
        "permissions": {},
        "persistent": False,
        "react": False,
        "rightIcon": "",
        "size": "md",
        "successMessage": "Success!",
        "successValidateMessage": "Success!",
        "tableView": False,
        "theme": "primary",
        "type": "button",
        "watcherList": [],
    }

    # 3) integrator
    # NOTE: literals use the original values (no key normalization)
    transform_name_literal = f"'{module_path}'" if module_path else "''"
    module_id_literal = f"'{module_id_value}'" if module_id_value else "''"

    # Normalize each component key fully for id and mapping
    normalized_keys = [make_key(k) for k in component_keys if s(k)]

    data_input = [
        {
            "exclude": False, "header": False, "id": transform_name_literal,
            "mapping": "transformName", "optional": False, "required": False, "resolveBase64": False
        },
        {
            "exclude": False, "header": False, "id": module_id_literal,
            "mapping": "moduleId", "optional": False, "required": False, "resolveBase64": False
        },
    ]
    for nk in normalized_keys:
        data_input.append({
            "exclude": False, "header": False, "id": nk,
            "mapping": f"data.{nk}", "optional": False, "required": False, "resolveBase64": False
        })

    integrator = {
        "$$hashKey": hashkey(),
        "allowSingleRecordArray": True,
        "ancestors": ["root"],
        "arrayKeyPromote": "",
        "assignEmptyValues": True,
        "componentId": guid(),
        "doNotLookForDataKeyInResponse": True,
        "input": True,
        "integratorData": {
            "dataFrame": [],
            "dataInputObject": data_input,
            "dataOutputObject": [
                {"header": False, "id": output_target_id, "mapping": "jsonData", "option": "replace"}
            ],
            "dataTestObject": [],
            "integratorUrl": "/fbu/uapi/transformer"
        },
        "key": key_integrator,
        "label": "plugTransform",
        "linked": {
            "inputs": [transform_name_literal, module_id_literal, *normalized_keys, None],
            "outputs": [output_target_id, None]
        },
        "multiple": True,
        "permissions": {},
        "persistent": False,
        "prefix": "",
        "preRequestTransform": None,
        "protected": False,
        "react": False,
        "requestType": "post",
        "service": None,
        "serviceType": "unqork",
        "tableView": True,
        "topLevelArray": False,
        "triggerType": "manual",
        "type": "integrator",
        "unique": False,
        "unqorkApi": "/transformer:post",
        "urlPlaceholders": False,
        "watcherList": [],
    }

    return [hidden, btn, integrator]


def generate_from_excel(excel: Path, out_dir: Path) -> Path:
    """Read Excel, group by module, and write one JSON file per Module_path."""
    df = pd.read_excel(excel)
    df.columns = [c.strip() for c in df.columns]

    needed = {"Module_name", "Module_path", "Component_key"}
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    out_dir.mkdir(parents=True, exist_ok=True)

    grouped = df.groupby(["Module_name", "Module_path"], dropna=False)
    count = 0
    for (module_name, module_path), g in grouped:
        module_path = s(module_path)
        if not module_path:
            print(f"! Skipping group with empty Module_path (Module_name='{s(module_name)}').")
            continue

        # ModuleId: Module_id -> fallback Excel_Component_id -> ""
        if "Module_id" in g.columns and g["Module_id"].notna().any():
            module_id_value = s(g["Module_id"].iloc[0])
        elif "Excel_Component_id" in g.columns and g["Excel_Component_id"].notna().any():
            module_id_value = s(g["Excel_Component_id"].iloc[0])
        else:
            module_id_value = ""

        # Unique component keys for this module (raw)
        keys = [s(x) for x in g["Component_key"].dropna().astype(str).unique().tolist()]
        keys = [k for k in keys if k]

        components = make_components(
            module_path=module_path,
            module_id_value=module_id_value,
            component_keys=keys
        )

        # File named by raw Module_path (safe chars)
        safe_name = module_path.replace("/", "_").replace("\\", "_").strip()
        out_file = out_dir / f"{safe_name}.json"
        out_file.write_text(json.dumps(components, indent=2, ensure_ascii=False), encoding="utf-8")
        count += 1

    print(f"✓ Wrote {count} module JSON files to {out_dir.resolve()}")
    return out_dir


def main():
    ap = argparse.ArgumentParser(description="Create Unqork-style component JSON per module from Excel.")
    ap.add_argument("--excel", type=Path, default=Path("aggregated_components_uniqueId.xlsx"))
    ap.add_argument("--out", type=Path, default=Path("unqork_output"))
    args = ap.parse_args()

    if not args.excel.exists():
        raise SystemExit(f"Excel not found: {args.excel}")

    generate_from_excel(args.excel, args.out)


if __name__ == "__main__":
    main()
