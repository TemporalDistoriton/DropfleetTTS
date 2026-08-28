from pathlib import Path
import json
import csv

# ============================================================
# CONFIGURATION
# ============================================================

NEW_FOLDER = Path("Data")
OLD_FOLDER = Path("Data_old")

CSV_LOG = Path("changelog.csv")
TEXT_LOG = Path("changelog.txt")


# ============================================================
# FILE DISCOVERY
# ============================================================

def get_json_files(folder):
    """
    Return:
        relative path -> absolute path

    Only JSON files are included.
    """
    files = {}

    for path in folder.rglob("*.json"):
        if path.is_file():
            relative_path = path.relative_to(folder)
            files[relative_path] = path

    return files


# ============================================================
# JSON LOADING
# ============================================================

def load_json(path):
    """
    Load a JSON file.
    """
    with path.open("r", encoding="utf-8-sig") as file:
        return json.load(file)


# ============================================================
# JSON COMPARISON
# ============================================================

def compare_json(old, new, path=""):
    """
    Recursively compare two JSON structures.

    Returns a list of changes:
        {
            "type": "ADDED" / "REMOVED" / "CHANGED",
            "path": "...",
            "old": ...,
            "new": ...
        }
    """

    changes = []

    # --------------------------------------------------------
    # Dictionaries / JSON objects
    # --------------------------------------------------------

    if isinstance(old, dict) and isinstance(new, dict):

        old_keys = set(old.keys())
        new_keys = set(new.keys())

        # Removed keys
        for key in sorted(old_keys - new_keys):
            current_path = f"{path}.{key}" if path else key

            changes.append({
                "type": "REMOVED",
                "path": current_path,
                "old": old[key],
                "new": None,
            })

        # Added keys
        for key in sorted(new_keys - old_keys):
            current_path = f"{path}.{key}" if path else key

            changes.append({
                "type": "ADDED",
                "path": current_path,
                "old": None,
                "new": new[key],
            })

        # Existing keys
        for key in sorted(old_keys & new_keys):
            current_path = f"{path}.{key}" if path else key

            changes.extend(
                compare_json(
                    old[key],
                    new[key],
                    current_path
                )
            )

        return changes

    # --------------------------------------------------------
    # Arrays / JSON lists
    # --------------------------------------------------------

    if isinstance(old, list) and isinstance(new, list):

        max_length = max(len(old), len(new))

        for index in range(max_length):

            current_path = f"{path}[{index}]"

            # New list item
            if index >= len(old):
                changes.append({
                    "type": "ADDED",
                    "path": current_path,
                    "old": None,
                    "new": new[index],
                })

            # Removed list item
            elif index >= len(new):
                changes.append({
                    "type": "REMOVED",
                    "path": current_path,
                    "old": old[index],
                    "new": None,
                })

            # Compare existing item
            else:
                changes.extend(
                    compare_json(
                        old[index],
                        new[index],
                        current_path
                    )
                )

        return changes

    # --------------------------------------------------------
    # Values
    # --------------------------------------------------------

    if old != new:
        changes.append({
            "type": "CHANGED",
            "path": path,
            "old": old,
            "new": new,
        })

    return changes


# ============================================================
# FORMATTING
# ============================================================

def format_value(value):
    """
    Format JSON values nicely for the text report.
    """

    if value is None:
        return "null"

    if isinstance(value, (dict, list)):
        return json.dumps(
            value,
            indent=2,
            ensure_ascii=False
        )

    return json.dumps(
        value,
        ensure_ascii=False
    )


# ============================================================
# MAIN
# ============================================================

def main():

    if not NEW_FOLDER.exists():
        print(f"ERROR: Folder not found: {NEW_FOLDER}")
        return

    if not OLD_FOLDER.exists():
        print(f"ERROR: Folder not found: {OLD_FOLDER}")
        return

    print("Scanning JSON files...")

    new_files = get_json_files(NEW_FOLDER)
    old_files = get_json_files(OLD_FOLDER)

    all_paths = sorted(
        set(new_files.keys()) | set(old_files.keys()),
        key=lambda x: str(x).lower()
    )

    file_results = []

    summary = {
        "NEW": 0,
        "REMOVED": 0,
        "MODIFIED": 0,
        "UNCHANGED": 0,
        "ERROR": 0,
    }

    # ========================================================
    # COMPARE FILES
    # ========================================================

    for relative_path in all_paths:

        new_path = new_files.get(relative_path)
        old_path = old_files.get(relative_path)

        # ----------------------------------------------------
        # New file
        # ----------------------------------------------------

        if new_path and not old_path:

            file_results.append({
                "status": "NEW",
                "file": str(relative_path),
                "changes": [],
                "error": None,
            })

            summary["NEW"] += 1
            continue

        # ----------------------------------------------------
        # Removed file
        # ----------------------------------------------------

        if old_path and not new_path:

            file_results.append({
                "status": "REMOVED",
                "file": str(relative_path),
                "changes": [],
                "error": None,
            })

            summary["REMOVED"] += 1
            continue

        # ----------------------------------------------------
        # Compare JSON
        # ----------------------------------------------------

        try:
            old_json = load_json(old_path)
            new_json = load_json(new_path)

            changes = compare_json(
                old_json,
                new_json
            )

            if changes:
                status = "MODIFIED"
                summary["MODIFIED"] += 1

            else:
                status = "UNCHANGED"
                summary["UNCHANGED"] += 1

            file_results.append({
                "status": status,
                "file": str(relative_path),
                "changes": changes,
                "error": None,
            })

        except Exception as error:

            file_results.append({
                "status": "ERROR",
                "file": str(relative_path),
                "changes": [],
                "error": str(error),
            })

            summary["ERROR"] += 1

    # ========================================================
    # CSV OUTPUT
    # ========================================================

    with CSV_LOG.open(
        "w",
        newline="",
        encoding="utf-8-sig"
    ) as csv_file:

        writer = csv.writer(csv_file)

        writer.writerow([
            "File Status",
            "File",
            "Change Type",
            "JSON Path",
            "Old Value",
            "New Value",
        ])

        for result in file_results:

            if result["changes"]:

                for change in result["changes"]:

                    writer.writerow([
                        result["status"],
                        result["file"],
                        change["type"],
                        change["path"],
                        format_value(change["old"])
                        if change["type"] != "ADDED"
                        else "",
                        format_value(change["new"])
                        if change["type"] != "REMOVED"
                        else "",
                    ])

            else:

                writer.writerow([
                    result["status"],
                    result["file"],
                    "",
                    "",
                    "",
                    "",
                ])

    # ========================================================
    # TEXT OUTPUT
    # ========================================================

    with TEXT_LOG.open(
        "w",
        encoding="utf-8"
    ) as log:

        log.write("=" * 80 + "\n")
        log.write("JSON DATA CHANGELOG\n")
        log.write("=" * 80 + "\n\n")

        log.write(
            f"Baseline : {OLD_FOLDER.resolve()}\n"
        )

        log.write(
            f"Current  : {NEW_FOLDER.resolve()}\n\n"
        )

        log.write("SUMMARY\n")
        log.write("-" * 80 + "\n")

        for status, count in summary.items():
            log.write(
                f"{status:<12}: {count}\n"
            )

        log.write("\n")

        # ----------------------------------------------------
        # Detailed file changes
        # ----------------------------------------------------

        for result in file_results:

            if result["status"] == "UNCHANGED":
                continue

            log.write("=" * 80 + "\n")
            log.write(
                f"{result['status']}: {result['file']}\n"
            )
            log.write("=" * 80 + "\n\n")

            if result["error"]:
                log.write(
                    f"ERROR: {result['error']}\n\n"
                )
                continue

            if result["status"] == "NEW":
                log.write(
                    "File exists in Data but not Data_old.\n\n"
                )
                continue

            if result["status"] == "REMOVED":
                log.write(
                    "File existed in Data_old but is no longer in Data.\n\n"
                )
                continue

            for change in result["changes"]:

                log.write(
                    f"{change['type']}: "
                    f"{change['path']}\n"
                )

                if change["type"] in {
                    "CHANGED",
                    "REMOVED"
                }:
                    log.write(
                        f"  OLD: "
                        f"{format_value(change['old'])}\n"
                    )

                if change["type"] in {
                    "CHANGED",
                    "ADDED"
                }:
                    log.write(
                        f"  NEW: "
                        f"{format_value(change['new'])}\n"
                    )

                log.write("\n")

    # ========================================================
    # CONSOLE
    # ========================================================

    print()
    print("Comparison complete.")
    print()

    for status, count in summary.items():
        print(f"{status:<12}: {count}")

    print()
    print(f"CSV log  : {CSV_LOG.resolve()}")
    print(f"Text log : {TEXT_LOG.resolve()}")


if __name__ == "__main__":
    main()