from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
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


SOURCE_FILE_OVERRIDE = None
INPUT_FILENAME = "Guild_Roster_Manager.lua"
OUTPUT_DIRECTORY_NAME = "grm_output"
NIL_SENTINEL = "__GRM_LUA_NIL__"

WOW_RETAIL_ACCOUNT_ROOTS = [
    Path(r"C:\Program Files (x86)\World of Warcraft\_retail_\WTF\Account"),
    Path(r"C:\Program Files\World of Warcraft\_retail_\WTF\Account"),
]


class ProgressReporter:
    def __init__(self, total_steps: int) -> None:
        self.total_steps = total_steps
        self.current_step = 0

    def update(self, message: str) -> None:
        self.current_step += 1
        width = 36
        ratio = min(self.current_step / self.total_steps, 1)
        filled = int(width * ratio)
        bar = "#" * filled + "-" * (width - filled)
        percent = int(ratio * 100)
        print(f"\r[{bar}] {percent:3d}% - {message}", end="", flush=True)

    def finish(self) -> None:
        print()


def exit_with_error(message: str) -> NoReturn:
    print()
    print(message, file=sys.stderr)
    input("Press Enter to exit...")
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

        root_level_candidate = account_root / "SavedVariables" / INPUT_FILENAME
        if root_level_candidate.exists() and root_level_candidate.is_file():
            candidates.append(root_level_candidate)

    return sorted(set(candidates))


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
        account_name = infer_account_name_from_path(candidate)
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


def infer_account_name_from_path(path: Path) -> str:
    parts = list(path.parts)
    if "Account" in parts:
        account_index = parts.index("Account")
        if account_index + 1 < len(parts):
            account_name = parts[account_index + 1]
            if account_name != "SavedVariables":
                return account_name
    return "manual"


def resolve_source_files() -> list[Path]:
    if SOURCE_FILE_OVERRIDE is not None:
        source_file = normalize_source_path(SOURCE_FILE_OVERRIDE)

        if not source_file.exists():
            exit_with_error(f"Input file not found (override): {source_file}")

        if not source_file.is_file():
            exit_with_error(f"Input path is not a file (override): {source_file}")

        return [source_file]

    same_directory_file = get_app_directory() / INPUT_FILENAME
    if same_directory_file.exists() and same_directory_file.is_file():
        return [same_directory_file]

    print(f"{INPUT_FILENAME} was not found next to this program.")
    print()
    print("It is usually located under paths like:")
    print(r"C:\Program Files (x86)\World of Warcraft\_retail_\WTF\Account\<ACCOUNT>\SavedVariables\Guild_Roster_Manager.lua")
    print()
    print("This build will try to load ALL account files it can find.")
    print()

    wow_candidates = find_wow_savedvariables_candidates()

    if wow_candidates:
        print("Found these account files:")
        for candidate in wow_candidates:
            print(f"  - {infer_account_name_from_path(candidate)}: {candidate}")
        print()
        use_all = input("Use all detected account files? [Y/n]: ").strip().lower()

        if use_all in ("", "y", "yes"):
            return wow_candidates

        selected_candidate = choose_candidate_interactively(wow_candidates)
        if selected_candidate is not None:
            return [selected_candidate]

    selected_file = prompt_for_source_file()

    if selected_file is not None:
        if selected_file.exists() and selected_file.is_file():
            return [selected_file]
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


def parse_epoch_to_iso(value: Any) -> str | None:
    try:
        if value in (None, 0, "0", ""):
            return None
        return datetime.fromtimestamp(int(value)).isoformat(timespec="seconds")
    except Exception:
        return None


def extract_player_id(record: dict[str, Any]) -> str | None:
    guid = record.get("GUID")
    if isinstance(guid, str) and guid.strip():
        return guid
    return None


def strip_color_codes(text: str | None) -> str | None:
    if not isinstance(text, str):
        return text
    text = re.sub(r"\|c[0-9a-fA-F]{8}", "", text)
    text = text.replace("|r", "")
    return text.strip()


def normalize_player_name(value: str | None) -> str | None:
    cleaned = strip_color_codes(value)
    if not isinstance(cleaned, str):
        return cleaned
    cleaned = cleaned.strip()
    return cleaned or None


def classify_event(event_code: Any) -> str:
    mapping = {
        1: "promotion",
        2: "demotion",
        3: "removed",
        4: "added",
        5: "note_change",
        6: "officer_note_change",
        7: "rank_change",
        8: "join",
        9: "leave",
    }
    return mapping.get(event_code, "unknown")


def parse_actor_target(record: list[Any]) -> tuple[str | None, str | None]:
    actor = normalize_player_name(record[3]) if len(record) > 3 else None
    target = normalize_player_name(record[4]) if len(record) > 4 else None
    return actor, target


def safe_get_list_value(value: Any, index: int) -> Any:
    if isinstance(value, list) and len(value) > index:
        return value[index]
    return None


def merge_dict_of_dicts_with_account(parsed_files: list[dict[str, Any]], variable_name: str) -> dict[str, dict[str, dict[str, Any]]]:
    merged: dict[str, dict[str, dict[str, Any]]] = {}

    for file_entry in parsed_files:
        source_account = file_entry["source_account"]
        parsed = file_entry["parsed"]
        dataset = parsed.get(variable_name, {})

        if not isinstance(dataset, dict):
            continue

        for guild_key, guild_members in dataset.items():
            if not isinstance(guild_members, dict):
                continue

            merged.setdefault(guild_key, {})

            for player_key, record in guild_members.items():
                if not isinstance(record, dict):
                    continue

                enriched_record = dict(record)
                enriched_record["_source_account"] = source_account
                merged[guild_key][player_key] = enriched_record

    return merged


def merge_dict_of_lists_with_account(parsed_files: list[dict[str, Any]], variable_name: str) -> dict[str, list[Any]]:
    merged: dict[str, list[Any]] = defaultdict(list)

    for file_entry in parsed_files:
        source_account = file_entry["source_account"]
        parsed = file_entry["parsed"]
        dataset = parsed.get(variable_name, {})

        if not isinstance(dataset, dict):
            continue

        for guild_key, guild_records in dataset.items():
            if not isinstance(guild_records, list):
                continue

            for record in guild_records:
                if isinstance(record, list):
                    merged[guild_key].append(record + [source_account])
                else:
                    merged[guild_key].append(record)

    return dict(merged)


def merge_alt_groups_with_account(parsed_files: list[dict[str, Any]], variable_name: str) -> dict[str, dict[Any, dict[str, Any]]]:
    merged: dict[str, dict[Any, dict[str, Any]]] = defaultdict(dict)

    for file_entry in parsed_files:
        source_account = file_entry["source_account"]
        parsed = file_entry["parsed"]
        dataset = parsed.get(variable_name, {})

        if not isinstance(dataset, dict):
            continue

        for guild_key, guild_groups in dataset.items():
            if not isinstance(guild_groups, dict):
                continue

            for alt_group_id, alt_group_record in guild_groups.items():
                if not isinstance(alt_group_record, dict):
                    continue

                enriched = dict(alt_group_record)
                enriched["_source_account"] = source_account
                merged[guild_key][f"{source_account}:{alt_group_id}"] = enriched

    return dict(merged)


def merge_alt_flags_with_account(parsed_files: list[dict[str, Any]], variable_name: str) -> dict[str, dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = defaultdict(dict)

    for file_entry in parsed_files:
        source_account = file_entry["source_account"]
        parsed = file_entry["parsed"]
        dataset = parsed.get(variable_name, {})

        if not isinstance(dataset, dict):
            continue

        for guild_key, guild_flags in dataset.items():
            if not isinstance(guild_flags, dict):
                continue

            for player_key, flag_value in guild_flags.items():
                merged[guild_key][player_key] = {
                    "_source_account": source_account,
                    "_flag_value": flag_value,
                }

    return dict(merged)


def extract_member_rows(dataset: Any, record_type: str, snapshot_time: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_members in dataset.items():
        if not isinstance(guild_members, dict):
            continue

        for player_key, member_record in guild_members.items():
            if not isinstance(member_record, dict):
                continue

            player_id = extract_player_id(member_record)

            base_fields = {
                "snapshot_time": snapshot_time,
                "guild_key": guild_key,
                "player_key": player_key,
                "player_id": player_id,
                "record_type": record_type,
                "source_account": member_record.get("_source_account"),
            }

            row = flatten_record(base_fields, member_record)
            row["name_clean"] = normalize_player_name(member_record.get("name"))
            row["join_datetime"] = parse_epoch_to_iso(safe_get_list_value(member_record.get("joinDateHist"), 0))
            row["last_online_days"] = member_record.get("lastOnline")
            row["time_entered_zone_iso"] = parse_epoch_to_iso(member_record.get("timeEnteredZone"))
            rows.append(row)

    return rows


def extract_rank_history(dataset: Any, membership_status: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_members in dataset.items():
        if not isinstance(guild_members, dict):
            continue

        for player_key, member_record in guild_members.items():
            if not isinstance(member_record, dict):
                continue

            player_id = extract_player_id(member_record)
            rank_hist = member_record.get("rankHist", [])

            if not isinstance(rank_hist, list):
                continue

            for sequence, entry in enumerate(rank_hist, start=1):
                if not isinstance(entry, list):
                    continue

                rows.append(
                    {
                        "source_account": member_record.get("_source_account"),
                        "guild_key": guild_key,
                        "player_key": player_key,
                        "player_id": player_id,
                        "membership_status": membership_status,
                        "history_sequence": sequence,
                        "rank_name": entry[0] if len(entry) > 0 else None,
                        "day": entry[1] if len(entry) > 1 else None,
                        "month": entry[2] if len(entry) > 2 else None,
                        "year": entry[3] if len(entry) > 3 else None,
                        "date_string": entry[4] if len(entry) > 4 else None,
                        "timestamp_epoch": entry[5] if len(entry) > 5 else None,
                        "timestamp_iso": parse_epoch_to_iso(entry[5] if len(entry) > 5 else None),
                        "unknown_flag": entry[6] if len(entry) > 6 else None,
                        "source_marker_1": entry[7] if len(entry) > 7 else None,
                        "source_marker_2": entry[8] if len(entry) > 8 else None,
                    }
                )

    return rows


def extract_join_history(dataset: Any, membership_status: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_members in dataset.items():
        if not isinstance(guild_members, dict):
            continue

        for player_key, member_record in guild_members.items():
            if not isinstance(member_record, dict):
                continue

            player_id = extract_player_id(member_record)
            join_hist = member_record.get("joinDateHist", [])

            if not isinstance(join_hist, list):
                continue

            for sequence, entry in enumerate(join_hist, start=1):
                if not isinstance(entry, list):
                    continue

                rows.append(
                    {
                        "source_account": member_record.get("_source_account"),
                        "guild_key": guild_key,
                        "player_key": player_key,
                        "player_id": player_id,
                        "membership_status": membership_status,
                        "history_sequence": sequence,
                        "day": entry[0] if len(entry) > 0 else None,
                        "month": entry[1] if len(entry) > 1 else None,
                        "year": entry[2] if len(entry) > 2 else None,
                        "date_string": entry[3] if len(entry) > 3 else None,
                        "timestamp_epoch": entry[4] if len(entry) > 4 else None,
                        "timestamp_iso": parse_epoch_to_iso(entry[4] if len(entry) > 4 else None),
                        "unknown_flag": entry[5] if len(entry) > 5 else None,
                        "source_marker_1": entry[6] if len(entry) > 6 else None,
                    }
                )

    return rows


def extract_fact_events(dataset: Any) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
]:
    events: list[dict[str, Any]] = []
    joins: list[dict[str, Any]] = []
    promotions: list[dict[str, Any]] = []
    demotions: list[dict[str, Any]] = []
    note_changes: list[dict[str, Any]] = []
    officer_note_changes: list[dict[str, Any]] = []
    rank_changes: list[dict[str, Any]] = []
    leaves: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return events, joins, promotions, demotions, note_changes, officer_note_changes, rank_changes, leaves

    for guild_key, guild_logs in dataset.items():
        if not isinstance(guild_logs, list):
            continue

        for event_index, record in enumerate(guild_logs, start=1):
            if not isinstance(record, list):
                continue

            appended_source_account = record[-1] if len(record) > 0 and isinstance(record[-1], str) and "#" in record[-1] else None
            core_record = record[:-1] if appended_source_account is not None else record

            event_code = core_record[0] if len(core_record) > 0 else None
            event_type = classify_event(event_code)
            message = core_record[1] if len(core_record) > 1 else None
            actor, target = parse_actor_target(core_record)

            event_time = None
            for value in core_record:
                possible_time = parse_datetime_tuple(value)
                if possible_time is not None:
                    event_time = possible_time
                    break

            base_event = {
                "source_account": appended_source_account,
                "guild_key": guild_key,
                "event_index": event_index,
                "event_code": event_code,
                "event_type": event_type,
                "actor": actor,
                "target": target,
                "message": message,
                "event_time": event_time,
                "raw_record_json": serialize_nested_value(core_record),
            }

            events.append(base_event)

            if event_type == "join":
                joins.append(
                    {
                        **base_event,
                        "invited_by": actor,
                        "joined_player": target,
                        "join_level": core_record[7] if len(core_record) > 7 else None,
                        "unknown_flag": core_record[6] if len(core_record) > 6 else None,
                    }
                )

            elif event_type == "leave":
                leaves.append(
                    {
                        **base_event,
                        "player_left": target or actor,
                    }
                )

            elif event_type == "promotion":
                promotions.append(
                    {
                        **base_event,
                        "changed_by": actor,
                        "player_promoted": target,
                        "old_rank": core_record[5] if len(core_record) > 5 else None,
                        "new_rank": core_record[6] if len(core_record) > 6 else None,
                    }
                )

            elif event_type == "demotion":
                demotions.append(
                    {
                        **base_event,
                        "changed_by": actor,
                        "player_demoted": target,
                        "old_rank": core_record[5] if len(core_record) > 5 else None,
                        "new_rank": core_record[6] if len(core_record) > 6 else None,
                    }
                )

            elif event_type == "note_change":
                note_changes.append(
                    {
                        **base_event,
                        "player_changed": normalize_player_name(core_record[2]) if len(core_record) > 2 else actor,
                        "old_note": normalize_player_name(core_record[3]) if len(core_record) > 3 else None,
                        "new_note": normalize_player_name(core_record[4]) if len(core_record) > 4 else None,
                    }
                )

            elif event_type == "officer_note_change":
                officer_note_changes.append(
                    {
                        **base_event,
                        "player_changed": normalize_player_name(core_record[2]) if len(core_record) > 2 else actor,
                        "old_officer_note": normalize_player_name(core_record[3]) if len(core_record) > 3 else None,
                        "new_officer_note": normalize_player_name(core_record[4]) if len(core_record) > 4 else None,
                    }
                )

            elif event_type == "rank_change":
                rank_changes.append(
                    {
                        **base_event,
                        "player_changed": target,
                        "old_rank": core_record[5] if len(core_record) > 5 else None,
                        "new_rank": core_record[6] if len(core_record) > 6 else None,
                    }
                )

    return events, joins, promotions, demotions, note_changes, officer_note_changes, rank_changes, leaves


def extract_alt_group_members(dataset: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_alt_groups in dataset.items():
        if not isinstance(guild_alt_groups, dict):
            continue

        for alt_group_id, alt_group_record in guild_alt_groups.items():
            if not isinstance(alt_group_record, dict):
                continue

            main_name = alt_group_record.get("main")
            nickname_details = alt_group_record.get("nicknameDetails")
            birthday_info = alt_group_record.get("birthdayInfo")
            time_modified = alt_group_record.get("timeModified")
            source_account = alt_group_record.get("_source_account")

            numeric_member_entries = sorted(
                (key, value)
                for key, value in alt_group_record.items()
                if isinstance(key, int) and isinstance(value, dict)
            )

            for member_index, member_value in numeric_member_entries:
                rows.append(
                    {
                        "source_account": source_account,
                        "guild_key": guild_key,
                        "alt_group_id": alt_group_id,
                        "member_index": member_index,
                        "main_name": main_name,
                        "alt_name": member_value.get("name"),
                        "alt_name_clean": normalize_player_name(member_value.get("name")),
                        "alt_class": member_value.get("class"),
                        "group_time_modified_iso": parse_epoch_to_iso(time_modified),
                        "nickname_json": serialize_nested_value(nickname_details),
                        "birthday_json": serialize_nested_value(birthday_info),
                    }
                )

    return rows


def extract_alt_flags(dataset: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    if not isinstance(dataset, dict):
        return rows

    for guild_key, guild_flags in dataset.items():
        if not isinstance(guild_flags, dict):
            continue

        for player_key, flag_payload in guild_flags.items():
            source_account = flag_payload.get("_source_account") if isinstance(flag_payload, dict) else None
            flag_value = flag_payload.get("_flag_value") if isinstance(flag_payload, dict) else None

            rows.append(
                {
                    "source_account": source_account,
                    "guild_key": guild_key,
                    "player_key": player_key,
                    "player_name_clean": normalize_player_name(player_key),
                    "is_main_flag": flag_value[0] if isinstance(flag_value, list) and len(flag_value) > 0 else None,
                    "raw_flag_json": serialize_nested_value(flag_value),
                }
            )

    return rows


def build_dim_players(current_members: list[dict[str, Any]], former_members: list[dict[str, Any]]) -> list[dict[str, Any]]:
    players_by_key: dict[str, dict[str, Any]] = {}

    for row in current_members + former_members:
        player_id = row.get("player_id")
        player_key = row.get("player_key")

        unique_key = player_id or f"{row.get('source_account')}::{player_key}"
        if unique_key is None:
            continue

        existing = players_by_key.get(unique_key, {})

        merged = {
            "source_account": row.get("source_account"),
            "player_id": player_id,
            "player_key": player_key,
            "name": row.get("name"),
            "name_clean": row.get("name_clean"),
            "class": row.get("class"),
            "race": row.get("race"),
            "sex": row.get("sex"),
            "faction": row.get("faction"),
            "first_seen_guild": existing.get("first_seen_guild") or row.get("guild_key"),
            "latest_rank_name": row.get("rankName"),
            "latest_rank_index": row.get("rankIndex"),
            "latest_level": row.get("level"),
            "achievement_points": row.get("achievementPoints"),
            "mythic_score": row.get("MythicScore"),
        }

        for key, value in existing.items():
            if merged.get(key) in (None, "", 0):
                merged[key] = value

        players_by_key[unique_key] = merged

    return sorted(players_by_key.values(), key=lambda row: (str(row.get("name_clean") or ""), str(row.get("player_key") or "")))


def build_dim_guilds(*datasets: Any) -> list[dict[str, Any]]:
    guilds: set[str] = set()

    for dataset in datasets:
        if isinstance(dataset, dict):
            guilds.update(str(key) for key in dataset.keys())

    rows: list[dict[str, Any]] = []
    for guild_key in sorted(guilds):
        guild_name, _, realm = guild_key.rpartition("-")
        if not guild_name:
            guild_name = guild_key
            realm = None

        rows.append(
            {
                "guild_key": guild_key,
                "guild_name": guild_name,
                "realm": realm,
            }
        )

    return rows


def build_dim_ranks(current_members: list[dict[str, Any]], former_members: list[dict[str, Any]], rank_history_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: dict[tuple[str | None, Any], dict[str, Any]] = {}

    for row in current_members + former_members:
        rank_name = row.get("rankName")
        rank_index = row.get("rankIndex")
        if rank_name is None and rank_index is None:
            continue
        seen[(rank_name, rank_index)] = {
            "rank_name": rank_name,
            "rank_index": rank_index,
        }

    for row in rank_history_rows:
        rank_name = row.get("rank_name")
        if rank_name is None:
            continue
        seen.setdefault((rank_name, None), {"rank_name": rank_name, "rank_index": None})

    return sorted(seen.values(), key=lambda row: (9999 if row["rank_index"] is None else row["rank_index"], str(row["rank_name"])))


def build_fact_daily_snapshot(current_members: list[dict[str, Any]], snapshot_time: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for row in current_members:
        rows.append(
            {
                "source_account": row.get("source_account"),
                "snapshot_time": snapshot_time,
                "snapshot_date": snapshot_time[:10],
                "guild_key": row.get("guild_key"),
                "player_id": row.get("player_id"),
                "player_key": row.get("player_key"),
                "rank_name": row.get("rankName"),
                "rank_index": row.get("rankIndex"),
                "level": row.get("level"),
                "class": row.get("class"),
                "zone": row.get("zone"),
                "last_online_days": row.get("last_online_days"),
                "is_online": row.get("isOnline"),
                "achievement_points": row.get("achievementPoints"),
                "mythic_score": row.get("MythicScore"),
                "guild_rep": row.get("guildRep"),
            }
        )

    return rows


def build_fact_daily_guild_snapshot(fact_daily_snapshot_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], dict[str, Any]] = defaultdict(
        lambda: {
            "member_count": 0,
            "online_count": 0,
            "max_level_count": 0,
            "inactive_30_plus_count": 0,
            "achievement_points_total": 0,
            "mythic_score_total": 0,
        }
    )

    for row in fact_daily_snapshot_rows:
        guild_key = row.get("guild_key")
        snapshot_date = row.get("snapshot_date")
        if guild_key is None or snapshot_date is None:
            continue

        key = (guild_key, snapshot_date)
        grouped[key]["guild_key"] = guild_key
        grouped[key]["snapshot_date"] = snapshot_date
        grouped[key]["member_count"] += 1

        if row.get("is_online"):
            grouped[key]["online_count"] += 1

        if row.get("level") == 80:
            grouped[key]["max_level_count"] += 1

        last_online_days = row.get("last_online_days")
        if isinstance(last_online_days, (int, float)) and last_online_days >= 30:
            grouped[key]["inactive_30_plus_count"] += 1

        achievement_points = row.get("achievement_points")
        if isinstance(achievement_points, (int, float)):
            grouped[key]["achievement_points_total"] += achievement_points

        mythic_score = row.get("mythic_score")
        if isinstance(mythic_score, (int, float)):
            grouped[key]["mythic_score_total"] += mythic_score

    return sorted(grouped.values(), key=lambda row: (str(row["guild_key"]), str(row["snapshot_date"])))


def build_bridge_player_guild(current_members: list[dict[str, Any]], former_members: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[tuple[Any, Any, Any, Any]] = set()

    for row in current_members + former_members:
        key = (row.get("source_account"), row.get("player_id"), row.get("player_key"), row.get("guild_key"))
        if key in seen:
            continue
        seen.add(key)

        rows.append(
            {
                "source_account": row.get("source_account"),
                "player_id": row.get("player_id"),
                "player_key": row.get("player_key"),
                "guild_key": row.get("guild_key"),
                "record_type": row.get("record_type"),
            }
        )

    return rows


def dedupe_rows(rows: list[dict[str, Any]], key_fields: list[str]) -> list[dict[str, Any]]:
    deduped: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()

    for row in rows:
        key = tuple(row.get(field) for field in key_fields)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)

    return deduped


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
    progress = ProgressReporter(total_steps=34)
    snapshot_time = datetime.now().isoformat(timespec="seconds")

    progress.update("Locating source files")
    source_files = resolve_source_files()

    progress.update("Reading source files")
    parsed_files: list[dict[str, Any]] = []
    parse_errors_by_file: dict[str, dict[str, Any]] = {}

    for source_file in source_files:
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

        assignments = split_top_level_assignments(raw_text)
        parsed = parse_lua_assignments(assignments)
        source_account = infer_account_name_from_path(source_file)

        parsed_files.append(
            {
                "source_file": source_file,
                "source_account": source_account,
                "assignments": assignments,
                "parsed": parsed,
            }
        )

        parse_errors_by_file[str(source_file)] = {
            key: value
            for key, value in parsed.items()
            if isinstance(value, dict) and "__parse_error__" in value
        }

    output_directory = get_app_directory() / OUTPUT_DIRECTORY_NAME
    output_directory.mkdir(parents=True, exist_ok=True)

    progress.update("Merging member history")
    members_data = merge_dict_of_dicts_with_account(parsed_files, "GRM_GuildMemberHistory_Save")

    progress.update("Merging former member history")
    former_members_data = merge_dict_of_dicts_with_account(parsed_files, "GRM_PlayersThatLeftHistory_Save")

    progress.update("Merging logs")
    logs_data = merge_dict_of_lists_with_account(parsed_files, "GRM_LogReport_Save")

    progress.update("Merging alt groups")
    alts_data = merge_alt_groups_with_account(parsed_files, "GRM_Alts")

    progress.update("Merging alt flags")
    alt_flags_data = merge_alt_flags_with_account(parsed_files, "GRM_PlayerListOfAlts_Save")

    progress.update("Extracting current members")
    members_rows = extract_member_rows(members_data, "current_member", snapshot_time)

    progress.update("Extracting former members")
    former_members_rows = extract_member_rows(former_members_data, "former_member", snapshot_time)

    progress.update("Extracting rank history")
    rank_history_rows = extract_rank_history(members_data, "current_member") + extract_rank_history(former_members_data, "former_member")

    progress.update("Extracting join history")
    join_history_rows = extract_join_history(members_data, "current_member") + extract_join_history(former_members_data, "former_member")

    progress.update("Extracting event facts")
    fact_events_rows, fact_joins_rows, fact_promotions_rows, fact_demotions_rows, fact_note_changes_rows, fact_officer_note_changes_rows, fact_rank_changes_rows, fact_leaves_rows = extract_fact_events(logs_data)

    progress.update("Extracting alt group relationships")
    dim_alt_groups_rows = extract_alt_group_members(alts_data)

    progress.update("Extracting alt flags")
    alt_flags_rows = extract_alt_flags(alt_flags_data)

    progress.update("Deduplicating primary entities")
    members_rows = dedupe_rows(members_rows, ["source_account", "guild_key", "player_key"])
    former_members_rows = dedupe_rows(former_members_rows, ["source_account", "guild_key", "player_key"])
    rank_history_rows = dedupe_rows(rank_history_rows, ["source_account", "guild_key", "player_key", "history_sequence", "rank_name", "timestamp_epoch"])
    join_history_rows = dedupe_rows(join_history_rows, ["source_account", "guild_key", "player_key", "history_sequence", "timestamp_epoch"])
    fact_events_rows = dedupe_rows(fact_events_rows, ["source_account", "guild_key", "event_code", "message", "event_time"])
    fact_joins_rows = dedupe_rows(fact_joins_rows, ["source_account", "guild_key", "joined_player", "event_time"])
    fact_promotions_rows = dedupe_rows(fact_promotions_rows, ["source_account", "guild_key", "player_promoted", "event_time", "old_rank", "new_rank"])
    fact_demotions_rows = dedupe_rows(fact_demotions_rows, ["source_account", "guild_key", "player_demoted", "event_time", "old_rank", "new_rank"])
    fact_note_changes_rows = dedupe_rows(fact_note_changes_rows, ["source_account", "guild_key", "player_changed", "event_time", "old_note", "new_note"])
    fact_officer_note_changes_rows = dedupe_rows(fact_officer_note_changes_rows, ["source_account", "guild_key", "player_changed", "event_time", "old_officer_note", "new_officer_note"])
    fact_rank_changes_rows = dedupe_rows(fact_rank_changes_rows, ["source_account", "guild_key", "player_changed", "event_time", "old_rank", "new_rank"])
    fact_leaves_rows = dedupe_rows(fact_leaves_rows, ["source_account", "guild_key", "player_left", "event_time"])
    dim_alt_groups_rows = dedupe_rows(dim_alt_groups_rows, ["source_account", "guild_key", "alt_group_id", "alt_name"])
    alt_flags_rows = dedupe_rows(alt_flags_rows, ["source_account", "guild_key", "player_key"])

    progress.update("Building dim_players")
    dim_players_rows = build_dim_players(members_rows, former_members_rows)

    progress.update("Building dim_guilds")
    dim_guilds_rows = build_dim_guilds(members_data, former_members_data, logs_data, alts_data, alt_flags_data)

    progress.update("Building dim_ranks")
    dim_ranks_rows = build_dim_ranks(members_rows, former_members_rows, rank_history_rows)

    progress.update("Building bridge_player_guild")
    bridge_player_guild_rows = build_bridge_player_guild(members_rows, former_members_rows)

    progress.update("Building fact_daily_snapshot")
    fact_daily_snapshot_rows = build_fact_daily_snapshot(members_rows, snapshot_time)

    progress.update("Building fact_daily_guild_snapshot")
    fact_daily_guild_snapshot_rows = build_fact_daily_guild_snapshot(fact_daily_snapshot_rows)

    progress.update("Writing members.csv")
    write_csv(output_directory, "members.csv", members_rows)

    progress.update("Writing former_members.csv")
    write_csv(output_directory, "former_members.csv", former_members_rows)

    progress.update("Writing rank_history.csv")
    write_csv(output_directory, "rank_history.csv", rank_history_rows)

    progress.update("Writing join_history.csv")
    write_csv(output_directory, "join_history.csv", join_history_rows)

    progress.update("Writing fact_events.csv")
    write_csv(output_directory, "fact_events.csv", fact_events_rows)

    progress.update("Writing fact_joins.csv")
    write_csv(output_directory, "fact_joins.csv", fact_joins_rows)

    progress.update("Writing fact_promotions.csv")
    write_csv(output_directory, "fact_promotions.csv", fact_promotions_rows)

    progress.update("Writing fact_demotions.csv")
    write_csv(output_directory, "fact_demotions.csv", fact_demotions_rows)

    progress.update("Writing fact_note_changes.csv")
    write_csv(output_directory, "fact_note_changes.csv", fact_note_changes_rows)

    progress.update("Writing fact_officer_note_changes.csv")
    write_csv(output_directory, "fact_officer_note_changes.csv", fact_officer_note_changes_rows)

    progress.update("Writing fact_rank_changes.csv")
    write_csv(output_directory, "fact_rank_changes.csv", fact_rank_changes_rows)

    progress.update("Writing fact_leaves.csv")
    write_csv(output_directory, "fact_leaves.csv", fact_leaves_rows)

    progress.update("Writing dim_alt_groups.csv")
    write_csv(output_directory, "dim_alt_groups.csv", dim_alt_groups_rows)

    progress.update("Writing alt_flags.csv")
    write_csv(output_directory, "alt_flags.csv", alt_flags_rows)

    progress.update("Writing dim_players.csv")
    write_csv(output_directory, "dim_players.csv", dim_players_rows)

    progress.update("Writing dim_guilds.csv")
    write_csv(output_directory, "dim_guilds.csv", dim_guilds_rows)

    progress.update("Writing dim_ranks.csv")
    write_csv(output_directory, "dim_ranks.csv", dim_ranks_rows)

    progress.update("Writing bridge_player_guild.csv")
    write_csv(output_directory, "bridge_player_guild.csv", bridge_player_guild_rows)

    progress.update("Writing fact_daily_snapshot.csv")
    write_csv(output_directory, "fact_daily_snapshot.csv", fact_daily_snapshot_rows)

    progress.update("Writing fact_daily_guild_snapshot.csv")
    write_csv(output_directory, "fact_daily_guild_snapshot.csv", fact_daily_guild_snapshot_rows)

    all_top_level_variables = sorted(
        {
            variable_name
            for file_entry in parsed_files
            for variable_name in file_entry["assignments"].keys()
        }
    )

    manifest = {
        "source_files": [str(path) for path in source_files],
        "source_accounts": [infer_account_name_from_path(path) for path in source_files],
        "output_directory": str(output_directory),
        "generated_at": snapshot_time,
        "top_level_variables_found": all_top_level_variables,
        "guilds_in_members": collect_guild_keys(members_data),
        "guilds_in_former_members": collect_guild_keys(former_members_data),
        "guilds_in_logs": collect_guild_keys(logs_data),
        "guilds_in_alts": collect_guild_keys(alts_data),
        "guilds_in_alt_flags": collect_guild_keys(alt_flags_data),
        "row_counts": {
            "members": len(members_rows),
            "former_members": len(former_members_rows),
            "rank_history": len(rank_history_rows),
            "join_history": len(join_history_rows),
            "fact_events": len(fact_events_rows),
            "fact_joins": len(fact_joins_rows),
            "fact_promotions": len(fact_promotions_rows),
            "fact_demotions": len(fact_demotions_rows),
            "fact_note_changes": len(fact_note_changes_rows),
            "fact_officer_note_changes": len(fact_officer_note_changes_rows),
            "fact_rank_changes": len(fact_rank_changes_rows),
            "fact_leaves": len(fact_leaves_rows),
            "dim_alt_groups": len(dim_alt_groups_rows),
            "alt_flags": len(alt_flags_rows),
            "dim_players": len(dim_players_rows),
            "dim_guilds": len(dim_guilds_rows),
            "dim_ranks": len(dim_ranks_rows),
            "bridge_player_guild": len(bridge_player_guild_rows),
            "fact_daily_snapshot": len(fact_daily_snapshot_rows),
            "fact_daily_guild_snapshot": len(fact_daily_guild_snapshot_rows),
        },
        "parse_errors_by_file": parse_errors_by_file,
    }

    write_json(output_directory, "_manifest.json", manifest)

    progress.finish()

    print()
    print("Source files:")
    for path in source_files:
        print(f"  - {path}")
    print()
    print(f"Output: {output_directory}")
    print()
    print(json.dumps(manifest["row_counts"], indent=2))
    print()

    has_any_parse_errors = any(bool(errors) for errors in parse_errors_by_file.values())
    if has_any_parse_errors:
        print("Some top-level GRM variables had parse errors. See _manifest.json for details.")
        print()

    input("Done. Press Enter to exit...")


if __name__ == "__main__":
    main()