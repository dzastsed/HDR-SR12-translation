import argparse
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

try:
    from openpyxl import Workbook
except ImportError as exc:
    raise SystemExit(
        "Missing dependency 'openpyxl'. Install it with: pip install openpyxl"
    ) from exc

TABLE_START = 0x0004
TABLE_END = 0x1ACF
ENTRY_SIZE = 4
ALT_ENTRY_SIZE = 8
EXPECTED_FIRST_STRING_OFFSET = 0x1AD0


def choose_file_with_dialog() -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    selected = filedialog.askopenfilename(
        title="Select binary file to decode",
        filetypes=[("Binary files", "*.*")],
    )
    root.destroy()

    if not selected:
        return None
    return Path(selected)


def decode_null_terminated(data: bytes, start_offset: int) -> Dict[str, Any]:
    if start_offset >= len(data):
        return {
            "text": "",
            "raw_hex": "",
            "end_offset": None,
            "decode_note": "offset out of file range",
        }

    end = start_offset
    while end < len(data) and data[end] != 0:
        end += 1

    raw = data[start_offset:end]

    decode_attempts = ["utf-8", "cp932", "latin-1"]
    text = ""
    decode_note = ""

    for encoding in decode_attempts:
        try:
            text = raw.decode(encoding)
            decode_note = f"decoded as {encoding}"
            break
        except UnicodeDecodeError:
            continue

    if text == "":
        text = raw.decode("latin-1", errors="replace")
        decode_note = "decoded with latin-1 replacement"

    if end >= len(data):
        if decode_note:
            decode_note += "; no null terminator found before EOF"
        else:
            decode_note = "no null terminator found before EOF"

    return {
        "text": text,
        "raw_hex": raw.hex(" ").upper(),
        "end_offset": end if end < len(data) else None,
        "decode_note": decode_note,
    }


def parse_file(data: bytes) -> Dict[str, Any]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []

    if len(data) < TABLE_START + 4:
        raise ValueError("File is too small to contain the pointer table.")

    first_pointer = int.from_bytes(data[TABLE_START:TABLE_START + 4], "little")
    if first_pointer != EXPECTED_FIRST_STRING_OFFSET:
        warnings.append(
            f"First pointer is 0x{first_pointer:08X}, expected about 0x{EXPECTED_FIRST_STRING_OFFSET:08X}."
        )

    computed_table_end = first_pointer - 1 if first_pointer > TABLE_START else TABLE_END
    max_table_end = min(computed_table_end, TABLE_END, len(data) - 1)
    entry_size = ENTRY_SIZE

    if TABLE_START + ALT_ENTRY_SIZE <= len(data):
        candidate_sep = data[TABLE_START + 4:TABLE_START + 8]
        if candidate_sep == b"\x00\x00\x00\x00":
            entry_size = ALT_ENTRY_SIZE

    entry_count = 0

    for pos in range(TABLE_START, max_table_end + 1, entry_size):
        if pos + 4 > len(data):
            break

        pointer = int.from_bytes(data[pos:pos + 4], "little")
        if entry_size == ALT_ENTRY_SIZE:
            if pos + 8 > len(data):
                break
            separator = data[pos + 4:pos + 8]
            if separator != b"\x00\x00\x00\x00":
                warnings.append(
                    f"Entry {entry_count} at 0x{pos:08X} has non-zero separator: {separator.hex(' ').upper()}"
                )

        decoded = decode_null_terminated(data, pointer)

        rows.append(
            {
                "index": entry_count,
                "table_offset": pos,
                "string_offset": pointer,
                "string_end": decoded["end_offset"],
                "text": decoded["text"],
                "raw_hex": decoded["raw_hex"],
                "note": decoded["decode_note"],
            }
        )
        entry_count += 1

    return {
        "rows": rows,
        "warnings": warnings,
        "file_size": len(data),
        "first_pointer": first_pointer,
    }


def write_xlsx(output_path: Path, parsed: Dict[str, Any], source_file: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "DecodedStrings"

    ws.append(["Source file", str(source_file)])
    ws.append(["File size (bytes)", parsed["file_size"]])
    ws.append(["First pointer", f"0x{parsed['first_pointer']:08X}"])
    ws.append([])

    headers = [
        "Index",
        "TableOffsetHex",
        "StringOffsetHex",
        "StringEndHex",
        "DecodedText",
        "RawBytesHex",
        "Note",
    ]
    ws.append(headers)

    for row in parsed["rows"]:
        ws.append(
            [
                row["index"],
                f"0x{row['table_offset']:08X}",
                f"0x{row['string_offset']:08X}",
                f"0x{row['string_end']:08X}" if row["string_end"] is not None else "",
                row["text"],
                row["raw_hex"],
                row["note"],
            ]
        )

    ws2 = wb.create_sheet(title="Warnings")
    ws2.append(["Warnings"])
    if parsed["warnings"]:
        for item in parsed["warnings"]:
            ws2.append([item])
    else:
        ws2.append(["No warnings."])

    wb.save(output_path)


def build_output_path(input_path: Path, output_override: Optional[Path]) -> Path:
    if output_override:
        if output_override.suffix.lower() != ".xlsx":
            return output_override.with_suffix(".xlsx")
        return output_override

    return input_path.with_name(f"{input_path.stem}_decoded.xlsx")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Decode string table from binary file and export to Excel (.xlsx)."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        help="Path to input binary file. If omitted, a file picker dialog opens.",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Optional output .xlsx path. Defaults to <input>_decoded.xlsx",
    )

    args = parser.parse_args()

    if args.input_file:
        input_path = Path(args.input_file)
    else:
        input_path = choose_file_with_dialog()
        if input_path is None:
            print("No input file selected. Aborting.")
            return 1

    if not input_path.exists() or not input_path.is_file():
        print(f"Input file not found: {input_path}")
        return 1

    output_path = build_output_path(input_path, Path(args.output) if args.output else None)

    data = input_path.read_bytes()
    parsed = parse_file(data)
    write_xlsx(output_path, parsed, input_path)

    print(f"Decoded {len(parsed['rows'])} table entries.")
    print(f"Output written to: {output_path}")
    if parsed["warnings"]:
        print(f"Warnings: {len(parsed['warnings'])} (see Warnings sheet)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
