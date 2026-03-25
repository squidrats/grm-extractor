from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, NoReturn

import luadata
import pandas as pd

try:
    import tkinter as tk
    from tkinter import filedialog

    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False
    

class ProgressReporter:
    def __init__(self, total_steps: int) -> None:
        self.total_steps = total_steps
        self.current_step = 0

    def update(self, message: str) -> None:
        self.current_step += 1
        width = 30
        ratio = self.current_step / self.total_steps
        filled = int(width * ratio)
        bar = "#" * filled + "-" * (width - filled)
        percent = int(ratio * 100)
        print(f"\r[{bar}] {percent:3d}% - {message}", end="", flush=True)

    def finish(self) -> None:
        print()


SOURCE_FILE_OVERRIDE = None
INPUT_FILENAME = "Guild_Roster_Manager.lua"
OUTPUT_DIRECTORY_NAME = "grm_output"
NIL_SENTINEL = "__GRM_LUA_NIL__"

WOW_RETAIL_ACCOUNT_ROOTS = [
    Path(r"C:\Program Files (x86)\World of Warcraft\_retail_\WTF\Account"),
    Path(r"C:\Program Files\World of Warcraft\_retail_\WTF\Account"),
]


def exit_with_error(message: str) -> NoReturn:
    print(message, file=sys.stderr)
    raise SystemExit(1)


def get_app_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def normalize_source_path(path_value: str) -> Path:
    windows_drive_match = re.match(r"^([A-Za-z]):\\(.*)$", path_value)
    if windows_drive_match:
        drive_letter = windows_drive_match.group(1).lower()
        rest = windows_drive_match.group(2).replace("\\", "/")
        return Path(f"/mnt/{drive_letter}/{rest}")
    return Path(path_value).expanduser()


def find_wow_savedvariables_candidates() -> list[Path]:
    candidates: list[Path] = []

    for account_root in WOW_RETAIL_ACCOUNT_ROOTS:
        if not account_root.exists() or not account_root.is_dir():
            continue

        for account_dir in sorted(account_root.iterdir()):
            if not account_dir.is_dir():
                continue

            candidate = account_dir / "SavedVariables" / INPUT_FILENAME
            if candidate.exists() and candidate.is_file():
                candidates.append(candidate)

    return candidates


def choose_candidate_interactively(candidates: list[Path]) -> Path | None:
    if not candidates:
        return None

    if len(candidates) == 1:
        print("Found 1 WoW SavedVariables file automatically:")
        print(f"  1. {candidates[0]}")
        print("Using it.")
        return candidates[0]

    print("Found multiple WoW SavedVariables files:")
    for index, candidate in enumerate(candidates, start=1):
        account_name = candidate.parent.parent.name
        print(f"  {index}. Account {account_name}: {candidate}")

    while True:
        response = input(f"Select an account [1-{len(candidates)}] or press Enter to cancel: ").strip()

        if response == "":
            return None

        if response.isdigit():
            choice = int(response)
            if 1 <= choice <= len(candidates):
                return candidates[choice - 1]

        print("Invalid selection. Please try again.")


def prompt_for_source_file() -> Path | None:
    if not TK_AVAILABLE:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    initial_dir = None
    for candidate_root in WOW_RETAIL_ACCOUNT_ROOTS:
        if candidate_root.exists():
            initial_dir = str(candidate_root)
            break

    selected_file = filedialog.askopenfilename(
        title="Select Guild_Roster_Manager.lua",
        filetypes=[
            ("Lua files", "*.lua"),
            ("All files", "*.*"),
        ],
        initialdir=initial_dir,
        initialfile=INPUT_FILENAME,
    )

    root.destroy()

    if not selected_file:
        return None

    return Path(selected_file)


def resolve_source_file() -> Path:
    if SOURCE_FILE_OVERRIDE is not None:
        source_file = normalize_source_path(SOURCE_FILE_OVERRIDE)

        if not source_file.exists():
            exit_with_error(f"Input file not found (override): {source_file}")

        if not source_file.is_file():
            exit_with_error(f"Input path is not a file (override): {source_file}")

        return source_file

    source_file = get_app_directory() / INPUT_FILENAME

    if source_file.exists() and source_file.is_file():
        return source_file

    print(f"{INPUT_FILENAME} was not found next to this program.")
    print()
    print("It is usually located under a path like:")
    print(r"C:\Program Files (x86)\World of Warcraft\_retail_\WTF\Account\<ACCOUNT>\SavedVariables\Guild_Roster_Manager.lua")
    print()

    wow_candidates = find_wow_savedvariables_candidates()
    selected_candidate = choose_candidate_interactively(wow_candidates)

    if selected_candidate is not None:
        return selected_candidate

    selected_file = prompt_for_source_file()

    if selected_file is not None:
        if selected_file.exists() and selected_file.is_file():
            return selected_file
        exit_with_error(f"Selected file is invalid: {selected_file}")

    if TK_AVAILABLE:
        exit_with_error(
            f"{INPUT_FILENAME} was not found automatically and no file was selected."
        )

    exit_with_error(
        f"{INPUT_FILENAME} was not found automatically.\n"
        "No GUI file picker is available in this environment.\n"
        "Place the file next to the program or set SOURCE_FILE_OVERRIDE."
    )


def read_input_file() -> tuple[Path, str]:
    source_file = resolve_source_file()

    try:
        raw_text = source_file.read_text(encoding="utf-8", errors="replace")
    except Exception as exc:
        exit_with_error(f"Unable to read {source_file}: {exc}")

    required_markers = [
        "GRM_GuildMemberHistory_Save",
        "GRM_PlayersThatLeftHistory_Save",
        "GRM_LogReport_Save",
    ]

    if not any(marker in raw_text for marker in required_markers):
        exit_with_error(f"{source_file} does not appear to be a Guild Roster Manager SavedVariables file")

    return source_file, raw_text


def split_top_level_assignments(raw_text: str) -> dict[str, str]:
    variable_pattern = re.compile(r"(?m)^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*")
    matches = list(variable_pattern.finditer(raw_text))

    if not matches:
        exit_with_error("No top-level Lua variable assignments were found")

    assignments: dict[str, str] = {}

    for index, match in enumerate(matches):
        variable_name = match.group(1)
        value_start = match.end()
        value_end = matches[index + 1].start() if index + 1 < len(matches) else len(raw_text)
        lua_value = raw_text[value_start:value_end].strip()

        if lua_value.endswith(","):
            lua_value = lua_value[:-1].rstrip()

        assignments[variable_name] = lua_value

    return assignments


def replace_bare_nil_tokens(lua_value: str) -> str:
    return re.sub(r'(?<![\w"\]])\bnil\b(?!\s*=)', f'"{NIL_SENTINEL}"', lua_value)


def restore_nil_sentinel(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: restore_nil_sentinel(inner_value) for key, inner_value in value.items()}

    if isinstance(value, list):
        return [restore_nil_sentinel(item) for item in value]

    if isinstance(value, tuple):
        return [restore_nil_sentinel(item) for item in value]

    if value == NIL_SENTINEL:
        return None

    return value


def parse_lua_assignments(assignments: dict[str, str]) -> dict[str, Any]:
    parsed: dict[str, Any] = {}

    for variable_name, lua_value in assignments.items():
        normalized_lua_value = replace_bare_nil_tokens(lua_value)

        try:
            parsed_value = luadata.unserialize(normalized_lua_value)
            parsed[variable_name] = restore_nil_sentinel(parsed_value)
        except Exception as exc:
            parsed[variable_name] = {
                "__parse_error__": str(exc),
                "__raw_preview__": lua_value[:5000],
                "__normalized_preview__": normalized_lua_value[:5000],
            }

    return parsed


def convert_to_json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {str(key): convert_to_json_safe(inner_value) for key, inner_value in value.items()}

    if isinstance(value, list):
        return [convert_to_json_safe(item) for item in value]

    if isinstance(value, tuple):
        return [convert_to_json_safe(item) for item in value]

    return value


def serialize_nested_value(value: Any) -> str:
    return json.dumps(convert_to_json_safe(value), ensure_ascii=False)


def flatten_value(prefix: str, value: Any, row: dict[str, Any]) -> None:
    if isinstance(value, dict):
        for key, inner_value in value.items():
            next_prefix = f"{prefix}.{key}" if prefix else str(key)
            flatten_value(next_prefix, inner_value, row)
        return

    if isinstance(value, tuple):
        flatten_value(prefix, list(value), row)
        return

    if isinstance(value, list):
        row[prefix] = serialize_nested_value(value)
        return

    row[prefix] = value


def flatten_record(base_fields: dict[str, Any], record: dict[str, Any]) -> dict[str, Any]:
    row = dict(base_fields)

    for key, value in record.items():
        flatten_value(str(key), value, row)

    return row


def parse_datetime_tuple(value: Any) -> str | None:
    if not isinstance(value, list):
        return None

    if len(value) < 5:
        return None

    first_five = value[:5]

    if not all(isinstance(item, (int, float)) for item in first_five):
        return None

    try:
        day = int(first_five[0])
        month = int(first_five[1])
        year = int(first_five[2])
        hour = int(first_five[3])
        minute = int(first_five[4])
        parsed = datetime(year, month, day, hour, minute)
        return parsed.isoformat(timespec="minutes")
    except Exception:
        return None


def extract_members(dataset: Any, record_type: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_members in dataset.items():
        if not isinstance(guild_members, dict):
            continue

        for player_key, member_record in guild_members.items():
            if not isinstance(member_record, dict):
                continue

            base_fields = {
                "guild_key": guild_key,
                "player_key": player_key,
                "record_type": record_type,
            }

            row = flatten_record(base_fields, member_record)
            rows.append(row)

    return rows


def extract_logs(dataset: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_logs in dataset.items():
        if not isinstance(guild_logs, list):
            continue

        for log_index, log_record in enumerate(guild_logs, start=1):
            if not isinstance(log_record, list):
                rows.append(
                    {
                        "guild_key": guild_key,
                        "log_index": log_index,
                        "raw_record_json": serialize_nested_value(log_record),
                    }
                )
                continue

            row: dict[str, Any] = {
                "guild_key": guild_key,
                "log_index": log_index,
                "event_code": log_record[0] if len(log_record) > 0 else None,
                "message": log_record[1] if len(log_record) > 1 else None,
                "raw_record_json": serialize_nested_value(log_record),
            }

            for field_position, value in enumerate(log_record[2:], start=3):
                if isinstance(value, (dict, list, tuple)):
                    row[f"field_{field_position}"] = serialize_nested_value(value)
                else:
                    row[f"field_{field_position}"] = value

            event_time = None
            for value in log_record:
                possible_time = parse_datetime_tuple(value)
                if possible_time is not None:
                    event_time = possible_time
                    break

            row["event_time"] = event_time
            rows.append(row)

    return rows


def extract_alts(dataset: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_alt_groups in dataset.items():
        if not isinstance(guild_alt_groups, dict):
            continue

        for alt_group_id, alt_group_record in guild_alt_groups.items():
            if not isinstance(alt_group_record, dict):
                continue

            base_fields = {
                "guild_key": guild_key,
                "alt_group_id": alt_group_id,
            }

            row = flatten_record(base_fields, alt_group_record)
            rows.append(row)

    return rows


def extract_alt_flags(dataset: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_flags in dataset.items():
        if not isinstance(guild_flags, dict):
            continue

        for player_key, flag_value in guild_flags.items():
            row: dict[str, Any] = {
                "guild_key": guild_key,
                "player_key": player_key,
                "raw_flag_json": serialize_nested_value(flag_value),
            }

            if isinstance(flag_value, list) and len(flag_value) > 0:
                row["flag_1"] = flag_value[0]
            else:
                row["flag_1"] = None

            rows.append(row)

    return rows


def write_json(output_directory: Path, filename: str, data: Any) -> None:
    output_path = output_directory / filename
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(convert_to_json_safe(data), handle, indent=2, ensure_ascii=False)


def write_csv(output_directory: Path, filename: str, rows: list[dict[str, Any]]) -> None:
    output_path = output_directory / filename
    dataframe = pd.DataFrame(rows)
    dataframe.to_csv(output_path, index=False, encoding="utf-8-sig")


def collect_guild_keys(dataset: Any) -> list[str]:
    if isinstance(dataset, dict):
        return sorted(str(key) for key in dataset.keys())
    return []


def main() -> None:
    progress = ProgressReporter(total_steps=13)

    progress.update("Locating source file")
    source_file = resolve_source_file()

    progress.update("Reading input file")
    try:
        raw_text = source_file.read_text(encoding="utf-8", errors="replace")
    except Exception as exc:
        exit_with_error(f"Unable to read {source_file}: {exc}")

    required_markers = [
        "GRM_GuildMemberHistory_Save",
        "GRM_PlayersThatLeftHistory_Save",
        "GRM_LogReport_Save",
    ]

    if not any(marker in raw_text for marker in required_markers):
        exit_with_error(f"{source_file} does not appear to be a Guild Roster Manager SavedVariables file")

    output_directory = get_app_directory() / OUTPUT_DIRECTORY_NAME
    output_directory.mkdir(parents=True, exist_ok=True)

    progress.update("Splitting top-level Lua assignments")
    assignments = split_top_level_assignments(raw_text)

    progress.update("Parsing Lua data")
    parsed = parse_lua_assignments(assignments)

    members_data = parsed.get("GRM_GuildMemberHistory_Save", {})
    former_members_data = parsed.get("GRM_PlayersThatLeftHistory_Save", {})
    logs_data = parsed.get("GRM_LogReport_Save", {})
    alts_data = parsed.get("GRM_Alts", {})
    alt_flags_data = parsed.get("GRM_PlayerListOfAlts_Save", {})

    progress.update("Extracting current members")
    members_rows = extract_members(members_data, "current_member")

    progress.update("Extracting former members")
    former_members_rows = extract_members(former_members_data, "former_member")

    progress.update("Extracting logs")
    logs_rows = extract_logs(logs_data)

    progress.update("Extracting alt groups")
    alts_rows = extract_alts(alts_data)

    progress.update("Extracting alt flags")
    alt_flags_rows = extract_alt_flags(alt_flags_data)

    progress.update("Writing members.csv")
    write_csv(output_directory, "members.csv", members_rows)

    progress.update("Writing former_members.csv")
    write_csv(output_directory, "former_members.csv", former_members_rows)

    progress.update("Writing logs.csv")
    write_csv(output_directory, "logs.csv", logs_rows)

    progress.update("Writing alts.csv")
    write_csv(output_directory, "alts.csv", alts_rows)

    progress.update("Writing alt_flags.csv")
    write_csv(output_directory, "alt_flags.csv", alt_flags_rows)

    manifest = {
        "source_file": str(source_file),
        "output_directory": str(output_directory),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "top_level_variables_found": sorted(assignments.keys()),
        "guilds_in_members": collect_guild_keys(members_data),
        "guilds_in_former_members": collect_guild_keys(former_members_data),
        "guilds_in_logs": collect_guild_keys(logs_data),
        "guilds_in_alts": collect_guild_keys(alts_data),
        "guilds_in_alt_flags": collect_guild_keys(alt_flags_data),
        "row_counts": {
            "members": len(members_rows),
            "former_members": len(former_members_rows),
            "logs": len(logs_rows),
            "alts": len(alts_rows),
            "alt_flags": len(alt_flags_rows),
        },
        "parse_errors": {
            key: value
            for key, value in parsed.items()
            if isinstance(value, dict) and "__parse_error__" in value
        },
    }

    progress.update("Writing manifest")
    write_json(output_directory, "_manifest.json", manifest)

    progress.finish()

    print()
    print(f"Source: {source_file}")
    print(f"Output: {output_directory}")
    print(json.dumps(manifest["row_counts"], indent=2))
    print()
    input("Done. Press Enter to exit...")


if __name__ == "__main__":
    main()