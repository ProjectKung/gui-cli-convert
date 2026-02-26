from __future__ import annotations

import argparse
import io
import random
import re
import textwrap
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape as xml_escape

from PIL import Image, ImageDraw, ImageFont


DEFAULT_FDO = Path(r"c:\Users\proje\OneDrive\Documents\testgui\FDO25040LT2.log")
DEFAULT_APIC = Path(r"c:\Users\proje\OneDrive\Documents\testgui\apic.log")
DEFAULT_IMAGE = Path(r"c:\Users\proje\OneDrive\Documents\testgui\2025-11-25_102020.jpg")
DEFAULT_OUTDIR = Path(r"d:\mycode\gui-convert\output")
A4_PAGE_W = 794
A4_PAGE_H = 1123
PDF_BODY_FONT_SIZE = 11
PDF_BODY_LINE_HEIGHT = 15
PDF_CHAR_WIDTH_ESTIMATE = 6.7
HIGHLIGHT_YELLOW = (255, 244, 130)
TEXT_BLACK = (0, 0, 0)
PAGE_WHITE = (255, 255, 255)
PROMPT_ONLY_PATTERN = re.compile(r"^\s*[^\s#][^#]*#\s*$")
COMMAND_LINE_PATTERN = re.compile(r"^\s*[^\s#][^#]*#\s+\S", re.IGNORECASE)
SHOW_ENV_PATTERN = re.compile(r"#\s*(?:sh|sho|show)\s+(?:environment|env)\b", re.IGNORECASE)
SHOW_VERSION_PATTERN = re.compile(r"#\s*(?:sh|sho|show)\s+(?:version|ver)\b", re.IGNORECASE)
SHOW_RUNNING_CONFIG_PATTERN = re.compile(
    r"#\s*(?:sh|sho|show)\s+(?:running(?:-|\s+)config|run)\b",
    re.IGNORECASE,
)
SHOW_CLOCK_COMMAND_PATTERN = re.compile(r"#\s*(?:sh|sho|show)\s+(?:clock|clo)\b", re.IGNORECASE)
CLEAR_WORD_PATTERN = re.compile(r"\bclear\b", re.IGNORECASE)
TIME_LINE_PATTERN = re.compile(
    r"^(\s*\*?\s*)(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d+))?\s+(\S+)\s+([A-Za-z]{3})\s+([A-Za-z]{3})\s+(\d{1,2})\s+(\d{4})\s*$"
)
INTERFACE_NAME_PATTERN = re.compile(r"^[A-Za-z][A-Za-z0-9/._:-]*$")
INTERFACE_ERRORS_FIELD_PATTERN = re.compile(r"(\s+)(\S+)")
INTERFACE_ERRORS_HEADER_PATTERN = re.compile(r"^\s*port\s+.+$", re.IGNORECASE)
WORD_WRAP_SEPARATOR_PATTERN = re.compile(r"(\s+|,\s*)")
SECTION_SEPARATOR_PATTERN = re.compile(r"^\s*-{3,}(?:\s+-{3,})*\s*$")

MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
MONTH_TO_INDEX = {month: idx for idx, month in enumerate(MONTHS)}
WEEKDAYS = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")


@dataclass(frozen=True)
class FdoClockOptions:
    custom_mode: bool = False
    custom_date: str | None = None
    custom_start_time: str = "08:00:00"
    custom_end_time: str = "18:00:00"


@dataclass(frozen=True)
class FdoPreprocessStats:
    clear_removed: int
    clear_removed_lines: list[tuple[int, str]]
    interface_rows_changed: int
    interface_rows_seen: int
    interface_row_changes: list[str]
    clock_before: list[tuple[int, int, datetime, str]]
    clock_after: list[tuple[int, int, datetime, str]]


def decode_text_with_fallback(raw: bytes) -> str:
    encodings = ("utf-8-sig", "utf-16", "cp874", "cp1252", "latin-1")
    for enc in encodings:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def read_text_with_fallback(path: Path) -> str:
    return decode_text_with_fallback(path.read_bytes())


def _parse_hms_seconds(value: str, fallback_seconds: int) -> int:
    match = re.fullmatch(r"\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*", value or "")
    if not match:
        return fallback_seconds
    hour = int(match.group(1))
    minute = int(match.group(2))
    second = int(match.group(3) or "0")
    if hour > 23 or minute > 59 or second > 59:
        return fallback_seconds
    return (hour * 3600) + (minute * 60) + second


def _parse_clock_time_line(line: str) -> tuple[datetime, int, str, str] | None:
    match = TIME_LINE_PATTERN.match(line)
    if not match:
        return None
    prefix = match.group(1) or ""
    hour = int(match.group(2))
    minute = int(match.group(3))
    second = int(match.group(4))
    fraction_text = match.group(5) or ""
    fraction_digits = len(fraction_text)
    # datetime supports microseconds up to 6 digits.
    # If source has more than 6 digits, use the first 6 for time math and keep full width for formatting.
    microseconds = int((fraction_text[:6].ljust(6, "0"))) if fraction_text else 0
    timezone_token = match.group(6)
    month = match.group(8).title()
    day = int(match.group(9))
    year = int(match.group(10))
    month_index = MONTH_TO_INDEX.get(month)
    if month_index is None:
        return None
    try:
        dt = datetime(year, month_index + 1, day, hour, minute, second, microseconds)
    except ValueError:
        return None
    return dt, fraction_digits, timezone_token, prefix


def _format_clock_time_line(
    dt: datetime,
    fraction_digits: int,
    timezone_token: str,
    prefix: str,
) -> str:
    base = f"{dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}"
    if fraction_digits > 0:
        base_fraction = f"{dt.microsecond:06d}"
        if fraction_digits <= 6:
            fraction = base_fraction[:fraction_digits]
        else:
            fraction = base_fraction + ("0" * (fraction_digits - 6))
        base = f"{base}.{fraction}"
    weekday = WEEKDAYS[dt.weekday()]
    month = MONTHS[dt.month - 1]
    return f"{prefix}{base} {timezone_token} {weekday} {month} {dt.day} {dt.year}"


def _randomize_microsecond_for_precision(fraction_digits: int) -> int:
    if fraction_digits <= 0:
        return 0
    if fraction_digits >= 6:
        return random.randint(0, 999999)
    visible_max = (10 ** fraction_digits) - 1
    step = 10 ** (6 - fraction_digits)
    return random.randint(0, visible_max) * step


def _is_show_clock_command(line: str) -> bool:
    return bool(SHOW_CLOCK_COMMAND_PATTERN.search(line))


def _prompt_command_tokens(line: str) -> list[str]:
    match = re.search(r"#\s*(.+)$", line)
    if not match:
        return []
    raw = match.group(1).strip().lower()
    if not raw:
        return []
    return [token for token in re.split(r"\s+", raw) if token]


def _is_show_interface_counters_errors_command(line: str) -> bool:
    tokens = _prompt_command_tokens(line)
    if not tokens:
        return False

    first = tokens[0]
    if not first.startswith("sh"):
        return False

    search_tokens = tokens[1:]
    found_int = False
    found_count = False
    for token in search_tokens:
        if not found_int:
            if token.startswith("int"):
                found_int = True
            continue
        if not found_count:
            if token.startswith("cou"):
                found_count = True
            continue
        if token.startswith("err"):
            return True
    return False


def _is_interface_errors_header(line: str) -> bool:
    if not INTERFACE_ERRORS_HEADER_PATTERN.match(line):
        return False
    tokens = line.split()
    return len(tokens) >= 3 and tokens[0].lower() == "port"


def _parse_interface_errors_data_row(line: str) -> tuple[str, list[int], str] | None:
    match = re.match(r"^(\s*)(\S+)(.*)$", line)
    if not match:
        return None

    indent = match.group(1)
    interface_name = match.group(2)
    rest = match.group(3)
    if not INTERFACE_NAME_PATTERN.fullmatch(interface_name):
        return None
    # Avoid touching header lines like "Port ..."; data rows always include a digit.
    if not re.search(r"\d", interface_name):
        return None

    fields: list[tuple[str, str]] = []
    cursor = 0
    for field_match in INTERFACE_ERRORS_FIELD_PATTERN.finditer(rest):
        if field_match.start() != cursor:
            return None
        fields.append((field_match.group(1), field_match.group(2)))
        cursor = field_match.end()
    trailing = rest[cursor:]

    if not fields:
        return None
    if trailing.strip():
        return None

    prefix = f"{indent}{interface_name}"
    token_end_positions: list[int] = []
    for field_match in INTERFACE_ERRORS_FIELD_PATTERN.finditer(rest):
        token_end_positions.append(len(prefix) + field_match.end())
    return prefix, token_end_positions, trailing


def _interface_row_contains_non_dash_value(line: str) -> bool:
    match = re.match(r"^(\s*)(\S+)(.*)$", line)
    if not match:
        return False

    interface_name = match.group(2)
    rest = match.group(3)
    if not INTERFACE_NAME_PATTERN.fullmatch(interface_name):
        return False
    if not re.search(r"\d", interface_name):
        return False

    values: list[str] = []
    cursor = 0
    for field_match in INTERFACE_ERRORS_FIELD_PATTERN.finditer(rest):
        if field_match.start() != cursor:
            return False
        values.append(field_match.group(2))
        cursor = field_match.end()
    trailing = rest[cursor:]
    if not values or trailing.strip():
        return False

    return any(value != "--" for value in values)


def _normalize_interface_errors_block(block_lines: list[str]) -> list[str]:
    normalized = block_lines[:]
    group: list[tuple[int, str, list[int], str]] = []

    def flush_group() -> None:
        nonlocal group
        if not group:
            return
        num_cols = len(group[0][2])
        column_end_positions = [0] * num_cols
        for _, _, token_ends, _ in group:
            for idx, end_pos in enumerate(token_ends):
                if end_pos > column_end_positions[idx]:
                    column_end_positions[idx] = end_pos

        for line_idx, prefix, _, trailing in group:
            out = [prefix]
            current_len = len(prefix)
            for end_pos in column_end_positions:
                spaces = max(1, end_pos - current_len - 2)
                out.append(" " * spaces)
                out.append("--")
                current_len += spaces + 2
            out.append(trailing)
            normalized[line_idx] = "".join(out)
        group = []

    for idx, line in enumerate(block_lines):
        parsed = _parse_interface_errors_data_row(line)
        if parsed is None:
            flush_group()
            continue

        prefix, token_ends, trailing = parsed
        if group and len(token_ends) != len(group[0][2]):
            flush_group()
        group.append((idx, prefix, token_ends, trailing))

    flush_group()
    return normalized


def _force_interface_errors_to_dash(lines: list[str]) -> list[str]:
    out: list[str] = []
    in_errors_section = False
    section_buffer: list[str] = []

    def flush_section() -> None:
        nonlocal section_buffer
        if section_buffer:
            out.extend(_normalize_interface_errors_block(section_buffer))
            section_buffer = []

    for line in lines:
        if _is_show_interface_counters_errors_command(line):
            flush_section()
            in_errors_section = True
            out.append(line)
            continue

        if not in_errors_section and _is_interface_errors_header(line):
            in_errors_section = True
            section_buffer.append(line)
            continue

        if in_errors_section and (COMMAND_LINE_PATTERN.match(line) or PROMPT_ONLY_PATTERN.fullmatch(line)):
            flush_section()
            in_errors_section = False
            out.append(line)
            continue

        if in_errors_section:
            section_buffer.append(line)
            continue

        flush_section()
        out.append(line)

    flush_section()
    return out


def _adjust_show_clock_lines(lines: list[str], options: FdoClockOptions | None = None) -> list[str]:
    opts = options or FdoClockOptions()
    blocks: list[tuple[int, datetime, int, str, str]] = []

    for idx, line in enumerate(lines):
        if not _is_show_clock_command(line):
            continue
        value_idx = idx + 1
        while value_idx < len(lines) and lines[value_idx].strip() == "":
            value_idx += 1
        if value_idx >= len(lines):
            continue
        parsed = _parse_clock_time_line(lines[value_idx])
        if not parsed:
            continue
        dt, fraction_digits, timezone_token, prefix = parsed
        blocks.append((value_idx, dt, fraction_digits, timezone_token, prefix))

    if len(blocks) < 3:
        return lines

    line1_idx, dt1_raw, has_ms1, tz1, prefix1 = blocks[0]
    line2_idx, _dt2_raw, has_ms2, tz2, prefix2 = blocks[1]
    line3_idx, _dt3_raw, has_ms3, tz3, prefix3 = blocks[2]

    def apply_fraction_precision(dt: datetime, fraction_digits: int) -> datetime:
        if fraction_digits <= 0:
            return dt.replace(microsecond=0)
        if fraction_digits < 6:
            step = 10 ** (6 - fraction_digits)
            return dt.replace(microsecond=(dt.microsecond // step) * step)
        if fraction_digits > 6:
            return dt.replace(microsecond=_randomize_microsecond_for_precision(fraction_digits))
        return dt

    # Clock #1 behavior:
    # - auto mode (checkbox OFF): keep original clock #1 unchanged
    # - custom mode (checkbox ON): randomize clock #1 in selected date/time window
    if opts.custom_mode:
        if opts.custom_date:
            try:
                chosen_date = datetime.strptime(opts.custom_date, "%Y-%m-%d")
            except ValueError:
                chosen_date = datetime.now()
        else:
            chosen_date = datetime.now()

        start_sec = _parse_hms_seconds(opts.custom_start_time, 8 * 3600)
        end_sec = _parse_hms_seconds(opts.custom_end_time, 18 * 3600)
        if end_sec < start_sec:
            end_sec = start_sec

        dt1_window_start = chosen_date.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(
            seconds=start_sec
        )
        dt1_window_end = chosen_date.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(
            seconds=end_sec
        )

        if dt1_window_end < dt1_window_start:
            dt1_window_end = dt1_window_start

        # Randomize clock #1 inside selected custom range.
        range_us = int((dt1_window_end - dt1_window_start).total_seconds() * 1_000_000)
        offset_us = random.randint(0, max(0, range_us))
        dt1_new = dt1_window_start + timedelta(microseconds=offset_us)
        dt1_new = apply_fraction_precision(dt1_new, has_ms1)
    else:
        # Keep clock #1 as-is from source when custom mode is not selected.
        dt1_new = dt1_raw

    delta12_sec = random.randint(40, 90)
    delta23_sec = random.randint(420, 450)
    dt2_new = dt1_new + timedelta(seconds=delta12_sec)
    dt2_new = apply_fraction_precision(dt2_new, has_ms2)
    dt3_new = dt2_new + timedelta(seconds=delta23_sec)
    dt3_new = apply_fraction_precision(dt3_new, has_ms3)

    out = lines[:]
    out[line1_idx] = _format_clock_time_line(dt1_new, has_ms1, tz1, prefix1)
    out[line2_idx] = _format_clock_time_line(dt2_new, has_ms2, tz2, prefix2)
    out[line3_idx] = _format_clock_time_line(dt3_new, has_ms3, tz3, prefix3)
    return out


def _extract_show_clock_entries(lines: list[str]) -> list[tuple[int, int, datetime, str]]:
    entries: list[tuple[int, int, datetime, str]] = []
    for idx, line in enumerate(lines):
        if not _is_show_clock_command(line):
            continue
        value_idx = idx + 1
        while value_idx < len(lines) and lines[value_idx].strip() == "":
            value_idx += 1
        if value_idx >= len(lines):
            continue
        parsed = _parse_clock_time_line(lines[value_idx])
        if not parsed:
            continue
        dt, _fraction_digits, _tz, _prefix = parsed
        entries.append((idx, value_idx, dt, lines[value_idx]))
    return entries


def _wrapped_seconds_diff(start: datetime, end: datetime) -> int:
    delta = int((end - start).total_seconds())
    if delta < 0:
        delta += 24 * 3600
    return delta


def _count_matches(lines: list[str], pattern: re.Pattern[str]) -> int:
    return sum(1 for line in lines if pattern.search(line))


def _build_fdo_validation_report(
    validated_lines: list[str],
    clear_removed: int,
    clear_removed_lines: list[tuple[int, str]],
    interface_rows_changed: int,
    interface_rows_seen: int,
    interface_row_changes: list[str],
    clock_before: list[tuple[int, int, datetime, str]],
    clock_after: list[tuple[int, int, datetime, str]],
    show_log_title_present: bool = False,
    show_log_image_present: bool = False,
) -> str:
    show_clock_cmd_indices = [idx for idx, line in enumerate(validated_lines) if _is_show_clock_command(line)]
    show_version_indices = [idx for idx, line in enumerate(validated_lines) if SHOW_VERSION_PATTERN.search(line)]
    show_running_indices = [idx for idx, line in enumerate(validated_lines) if SHOW_RUNNING_CONFIG_PATTERN.search(line)]
    show_env_indices = [idx for idx, line in enumerate(validated_lines) if SHOW_ENV_PATTERN.search(line)]
    interface_cmd_indices = [
        idx for idx, line in enumerate(validated_lines) if _is_show_interface_counters_errors_command(line)
    ]
    show_version_count = len(show_version_indices)
    show_running_count = len(show_running_indices)
    show_env_count = len(show_env_indices)

    expected_counts = (
        ("show clock", len(show_clock_cmd_indices), 3),
        ("show version", show_version_count, 1),
        ("show running-config", show_running_count, 1),
        ("show environment", show_env_count, 1),
        ("show interface counters errors", len(interface_cmd_indices), 2),
    )

    def count_status(found: int, required: int) -> str:
        if found == required:
            return "ผ่าน"
        if found < required:
            return f"ขาด {required - found}"
        return f"เกิน {found - required}"

    count_lines = [
        f"- {name}: พบ {found}, ต้องมี {required} [{count_status(found, required)}]"
        for name, found, required in expected_counts
    ]

    order_ok = False
    order_detail = "ข้อมูลไม่พอสำหรับตรวจลำดับคำสั่งหลัก"
    if (
        len(show_clock_cmd_indices) >= 3
        and show_version_count >= 1
        and show_running_count >= 1
        and show_env_count >= 1
        and len(interface_cmd_indices) >= 2
    ):
        c1 = show_clock_cmd_indices[0]
        v1 = show_version_indices[0]
        r1 = show_running_indices[0]
        env1 = show_env_indices[0]
        c2 = show_clock_cmd_indices[1]
        e1 = interface_cmd_indices[0]
        c3 = show_clock_cmd_indices[2]
        e2 = interface_cmd_indices[1]
        order_ok = c1 < v1 < r1 < env1 < c2 < e1 < c3 < e2
        order_detail = (
            f"clock#1@{c1+1}, version@{v1+1}, running-config@{r1+1}, environment@{env1+1}, "
            f"clock#2@{c2+1}, iface#1@{e1+1}, clock#3@{c3+1}, iface#2@{e2+1} "
            f"=> {'ผ่าน' if order_ok else 'ไม่ผ่าน'}"
        )
    else:
        needed = []
        if len(show_clock_cmd_indices) < 3:
            needed.append(f"show clock >=3 (พบ {len(show_clock_cmd_indices)})")
        if show_version_count < 1:
            needed.append(f"show version >=1 (พบ {show_version_count})")
        if show_running_count < 1:
            needed.append(f"show running-config >=1 (พบ {show_running_count})")
        if show_env_count < 1:
            needed.append(f"show environment >=1 (พบ {show_env_count})")
        if len(interface_cmd_indices) < 2:
            needed.append(f"show interface counters errors >=2 (พบ {len(interface_cmd_indices)})")
        if needed:
            order_detail = "ต้องมี " + ", ".join(needed) + "."

    placement_ok = False
    placement_detail = "ข้อมูลไม่พอสำหรับตรวจตำแหน่ง"
    if len(show_clock_cmd_indices) >= 3 and len(interface_cmd_indices) >= 2:
        c2 = show_clock_cmd_indices[1]
        c3 = show_clock_cmd_indices[2]
        e1 = interface_cmd_indices[0]
        e2 = interface_cmd_indices[1]
        placement_ok = (c2 < e1 < c3) and (c3 < e2)
        placement_detail = (
            f"clock#2@{c2+1}, clock#3@{c3+1}, iface#1@{e1+1}, iface#2@{e2+1} "
            f"=> {'ผ่าน' if placement_ok else 'ไม่ผ่าน'}"
        )
    elif len(interface_cmd_indices) != 2:
        placement_detail = f"ต้องมี show interface counters errors 2 คำสั่ง, พบ {len(interface_cmd_indices)}."
    else:
        placement_detail = f"ต้องมี show clock อย่างน้อย 3 คำสั่ง, พบ {len(show_clock_cmd_indices)}."

    clock_delta_12: int | None = None
    clock_delta_23: int | None = None
    range_12_ok = False
    range_23_ok = False
    if len(clock_after) >= 3:
        clock_delta_12 = _wrapped_seconds_diff(clock_after[0][2], clock_after[1][2])
        clock_delta_23 = _wrapped_seconds_diff(clock_after[1][2], clock_after[2][2])
        range_12_ok = 40 <= clock_delta_12 <= 90
        range_23_ok = 420 <= clock_delta_23 <= 450

    clock_changed_lines = 0
    for idx in range(min(len(clock_before), len(clock_after), 3)):
        if clock_before[idx][3] != clock_after[idx][3]:
            clock_changed_lines += 1

    show_log_position_ok = False
    show_log_detail = "ต้องมี show interface counters errors 2 คำสั่งก่อนจึงจะตรวจตำแหน่ง Show log ได้"
    if len(interface_cmd_indices) >= 2:
        iface2_idx = interface_cmd_indices[1]
        next_command_idx = next(
            (idx for idx in range(iface2_idx + 1, len(validated_lines)) if COMMAND_LINE_PATTERN.match(validated_lines[idx])),
            None,
        )
        show_log_position_ok = next_command_idx is None
        if show_log_position_ok:
            show_log_detail = f"iface#2@{iface2_idx+1}, ไม่มี command ถัดจาก iface#2 => ผ่าน"
        else:
            show_log_detail = (
                f"iface#2@{iface2_idx+1}, next command@{next_command_idx+1} "
                f"('{validated_lines[next_command_idx].strip()}') => ไม่ผ่าน"
            )

    show_log_ok = show_log_position_ok and show_log_title_present and show_log_image_present

    counts_ok = all(found == required for _, found, required in expected_counts)
    deltas_ok = range_12_ok and range_23_ok
    overall_ok = counts_ok and order_ok and placement_ok and deltas_ok and show_log_ok

    clear_detail_lines = []
    if clear_removed_lines:
        clear_detail_lines.append("- รายละเอียดบรรทัด clear ที่ถูกลบ:")
        for line_no, line_text in clear_removed_lines:
            clear_detail_lines.append(f"- line {line_no}: {line_text}")

    report_lines = [
        f"ผลการตรวจสอบ: {'ผ่าน' if overall_ok else 'ไม่ผ่าน'}",
        "",
        "การเปลี่ยนแปลงที่ระบบทำ:",
        f"- ลบบรรทัด clear: {clear_removed}",
        *clear_detail_lines,
        f"- แถว show interface counters errors ที่ปรับเป็น '--': {interface_rows_changed}/{interface_rows_seen}",
        f"- จำนวนบรรทัดเวลา show clock ที่เปลี่ยน: {clock_changed_lines}/3",
        "",
        "คำสั่งที่ต้องมี:",
        *count_lines,
        "",
        "ตรวจลำดับ:",
        (
            "- ลำดับที่ต้องเป็น: "
            "clock#1 -> show version -> show running-config -> show environment -> "
            "clock#2 -> show interface counters errors #1 -> clock#3 -> show interface counters errors #2: "
            + ("ผ่าน" if order_ok else "ไม่ผ่าน")
        ),
        f"- รายละเอียด: {order_detail}",
        (
            "- show interface counters errors #1 ต้องอยู่ใต้ clock #2 และ #2 ต้องอยู่ใต้ clock #3: "
            + ("ผ่าน" if placement_ok else "ไม่ผ่าน")
        ),
        f"- รายละเอียด: {placement_detail}",
        "",
        "ตรวจช่วงเวลา show clock:",
    ]

    if clock_delta_12 is None or clock_delta_23 is None:
        report_lines.extend(
            [
                "- #1 -> #2: ไม่มีข้อมูล (ต้องมีเวลา show clock ที่ถูกต้อง 3 จุด)",
                "- #2 -> #3: ไม่มีข้อมูล (ต้องมีเวลา show clock ที่ถูกต้อง 3 จุด)",
            ]
        )
    else:
        report_lines.extend(
            [
                f"- #1 -> #2: {clock_delta_12}s (เป้าหมาย 40-90s) [{'ผ่าน' if range_12_ok else 'ไม่ผ่าน'}]",
                f"- #2 -> #3: {clock_delta_23}s (เป้าหมาย 420-450s) [{'ผ่าน' if range_23_ok else 'ไม่ผ่าน'}]",
            ]
        )

    report_lines.extend(["", "รายละเอียดเวลา clock ที่เปลี่ยน:"])
    if len(clock_before) < 3 or len(clock_after) < 3:
        report_lines.append("- ไม่มีข้อมูล (ต้องมีเวลา show clock ที่ถูกต้องก่อนและหลังอย่างละ 3 จุด)")
    else:
        for idx in range(3):
            before_line = clock_before[idx][3].strip()
            after_line = clock_after[idx][3].strip()
            report_lines.append(f"- clock #{idx+1}: {before_line} -> {after_line}")

    report_lines.extend(
        [
            "",
            "ตรวจส่วน Show log:",
            (
                "- หลัง show interface counters errors #2 ต้องเป็นส่วน Show log ต่อท้าย: "
                + ("ผ่าน" if show_log_ok else "ไม่ผ่าน")
            ),
            f"- รายละเอียดตำแหน่ง: {show_log_detail}",
            f"- พบหัวข้อ Show log: {'ผ่าน' if show_log_title_present else 'ไม่ผ่าน'}",
            f"- พบรูป Show log: {'ผ่าน' if show_log_image_present else 'ไม่ผ่าน'}",
            "",
            f"Interface Row ที่มีการแก้ไข (มีทั้งหมด: {len(interface_row_changes)}):",
        ]
    )

    if not interface_row_changes:
        report_lines.append("- ไม่มี")
    else:
        for before in interface_row_changes:
            report_lines.append(f"- {before}")
        report_lines.append("- ระบบได้ปรับบรรทัดข้างต้นให้เป็น '--' ในผลลัพธ์แล้ว")

    return "\n".join(report_lines)


def _preprocess_fdo_lines_and_stats(
    fdo_text: str,
    options: FdoClockOptions | None = None,
) -> tuple[list[str], FdoPreprocessStats]:
    raw_lines = fdo_text.splitlines()
    lines_no_clear: list[str] = []
    clear_removed_lines: list[tuple[int, str]] = []
    for line_no, line in enumerate(raw_lines, start=1):
        if CLEAR_WORD_PATTERN.search(line):
            clear_removed_lines.append((line_no, line))
            continue
        lines_no_clear.append(line)
    clear_removed = len(clear_removed_lines)

    lines_after_interface = _force_interface_errors_to_dash(lines_no_clear)
    interface_rows_seen = 0
    interface_rows_changed = 0
    interface_row_changes: list[str] = []
    for before, after in zip(lines_no_clear, lines_after_interface):
        if _parse_interface_errors_data_row(before) is not None:
            interface_rows_seen += 1
            if before != after and _interface_row_contains_non_dash_value(before):
                interface_rows_changed += 1
                interface_row_changes.append(before)

    clock_before = _extract_show_clock_entries(lines_after_interface)
    final_lines = _adjust_show_clock_lines(lines_after_interface, options=options)
    clock_after = _extract_show_clock_entries(final_lines)

    stats = FdoPreprocessStats(
        clear_removed=clear_removed,
        clear_removed_lines=clear_removed_lines,
        interface_rows_changed=interface_rows_changed,
        interface_rows_seen=interface_rows_seen,
        interface_row_changes=interface_row_changes,
        clock_before=clock_before,
        clock_after=clock_after,
    )
    return final_lines, stats


def preprocess_fdo_lines_with_report(
    fdo_text: str,
    options: FdoClockOptions | None = None,
) -> tuple[list[str], str]:
    final_lines, stats = _preprocess_fdo_lines_and_stats(fdo_text, options=options)
    report = _build_fdo_validation_report(
        validated_lines=final_lines,
        clear_removed=stats.clear_removed,
        clear_removed_lines=stats.clear_removed_lines,
        interface_rows_changed=stats.interface_rows_changed,
        interface_rows_seen=stats.interface_rows_seen,
        interface_row_changes=stats.interface_row_changes,
        clock_before=stats.clock_before,
        clock_after=stats.clock_after,
    )
    return final_lines, report


def preprocess_fdo_lines(fdo_text: str, options: FdoClockOptions | None = None) -> list[str]:
    lines, _stats = _preprocess_fdo_lines_and_stats(fdo_text, options=options)
    return lines


def find_insert_index(lines: list[str]) -> int:
    show_env_index = next(
        (idx for idx, line in enumerate(lines) if SHOW_ENV_PATTERN.search(line)),
        None,
    )
    if show_env_index is not None:
        return max(0, show_env_index - 1)

    show_version_index = next(
        (idx for idx, line in enumerate(lines) if SHOW_VERSION_PATTERN.search(line)),
        None,
    )
    start = (show_version_index + 1) if show_version_index is not None else 0

    for idx in range(start, len(lines)):
        if PROMPT_ONLY_PATTERN.fullmatch(lines[idx]):
            return idx

    for idx, line in enumerate(lines):
        if PROMPT_ONLY_PATTERN.fullmatch(line):
            return idx

    return max(0, len(lines) - 1)


def build_combined_lines(
    fdo_text: str,
    apic_text: str,
    fdo_clock_options: FdoClockOptions | None = None,
) -> list[str]:
    fdo_lines = preprocess_fdo_lines(fdo_text, options=fdo_clock_options)
    return _combine_fdo_and_apic_lines(fdo_lines, apic_text)


def build_combined_lines_with_report(
    fdo_text: str,
    apic_text: str,
    fdo_clock_options: FdoClockOptions | None = None,
    show_log_title_present: bool = False,
    show_log_image_present: bool = False,
) -> tuple[list[str], str]:
    fdo_lines, stats = _preprocess_fdo_lines_and_stats(fdo_text, options=fdo_clock_options)
    combined_lines = _combine_fdo_and_apic_lines(fdo_lines, apic_text)
    report = _build_fdo_validation_report(
        validated_lines=combined_lines,
        clear_removed=stats.clear_removed,
        clear_removed_lines=stats.clear_removed_lines,
        interface_rows_changed=stats.interface_rows_changed,
        interface_rows_seen=stats.interface_rows_seen,
        interface_row_changes=stats.interface_row_changes,
        clock_before=stats.clock_before,
        clock_after=stats.clock_after,
        show_log_title_present=show_log_title_present,
        show_log_image_present=show_log_image_present,
    )
    return combined_lines, report


def _combine_fdo_and_apic_lines(fdo_lines: list[str], apic_text: str) -> list[str]:
    apic_lines = apic_text.splitlines()
    if not fdo_lines:
        return apic_lines

    insert_at = find_insert_index(fdo_lines)
    first_part = fdo_lines[: insert_at + 1]
    remaining_part = fdo_lines[insert_at + 1 :]
    return first_part + [""] + apic_lines + [""] + remaining_part


def build_combined_text(
    fdo_text: str,
    apic_text: str,
    fdo_clock_options: FdoClockOptions | None = None,
) -> str:
    return "\n".join(build_combined_lines(fdo_text, apic_text, fdo_clock_options=fdo_clock_options)) + "\n"


def load_monospace_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\consola.ttf"),
        Path(r"C:\Windows\Fonts\cour.ttf"),
        Path(r"C:\Windows\Fonts\lucon.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def _wrap_text_segments(line: str, max_chars: int) -> list[str]:
    wrapper = textwrap.TextWrapper(
        width=max_chars,
        expand_tabs=False,
        replace_whitespace=False,
        drop_whitespace=False,
        break_long_words=False,
        break_on_hyphens=True,
    )
    wrapper.wordsep_re = WORD_WRAP_SEPARATOR_PATTERN
    segments = wrapper.wrap(line)
    return segments if segments else [""]


def wrap_lines(lines: list[str], max_chars: int) -> list[tuple[str, bool]]:
    wrapped: list[tuple[str, bool]] = []
    for line in lines:
        line = line.replace("\t", "    ")
        is_command = bool(COMMAND_LINE_PATTERN.match(line))
        if line == "":
            wrapped.append(("", False))
            continue
        segments = _wrap_text_segments(line, max_chars)
        wrapped.extend((segment, is_command) for segment in segments)
    return wrapped


def _wrap_single_line(line: str, max_chars: int) -> list[tuple[str, bool]]:
    line = line.replace("\t", "    ")
    is_command = bool(COMMAND_LINE_PATTERN.match(line))
    if line == "":
        return [("", False)]
    segments = _wrap_text_segments(line, max_chars)
    return [(segment, is_command) for segment in segments]


def _is_section_separator(line: str) -> bool:
    return bool(SECTION_SEPARATOR_PATTERN.match(line))


def _separator_block_segments(
    lines: list[str],
    start_index: int,
    max_chars: int,
) -> tuple[list[tuple[str, bool]], int] | None:
    if not _is_section_separator(lines[start_index]):
        return None

    end_index = None
    max_scan_lines = 40
    scan_limit = min(len(lines), start_index + max_scan_lines + 1)
    for idx in range(start_index + 1, scan_limit):
        if _is_section_separator(lines[idx]):
            end_index = idx
            break

    if end_index is None:
        return None

    block_segments: list[tuple[str, bool]] = []
    for idx in range(start_index, end_index + 1):
        block_segments.extend(_wrap_single_line(lines[idx], max_chars))
    return block_segments, (end_index + 1)


def _line_with_trailing_separator_segments(
    lines: list[str],
    start_index: int,
    max_chars: int,
) -> tuple[list[tuple[str, bool]], int] | None:
    next_index = start_index + 1
    if next_index >= len(lines):
        return None
    if _is_section_separator(lines[start_index]):
        return None
    if not _is_section_separator(lines[next_index]):
        return None

    pair_segments = _wrap_single_line(lines[start_index], max_chars)
    pair_segments.extend(_wrap_single_line(lines[next_index], max_chars))
    return pair_segments, (next_index + 1)


def _line_separator_header_separator_segments(
    lines: list[str],
    start_index: int,
    max_chars: int,
) -> tuple[list[tuple[str, bool]], int] | None:
    sep1_index = start_index + 1
    header_index = start_index + 2
    sep2_index = start_index + 3
    if sep2_index >= len(lines):
        return None
    if _is_section_separator(lines[start_index]):
        return None
    if not _is_section_separator(lines[sep1_index]):
        return None
    if _is_section_separator(lines[header_index]):
        return None
    if not _is_section_separator(lines[sep2_index]):
        return None

    block_segments = _wrap_single_line(lines[start_index], max_chars)
    block_segments.extend(_wrap_single_line(lines[sep1_index], max_chars))
    block_segments.extend(_wrap_single_line(lines[header_index], max_chars))
    block_segments.extend(_wrap_single_line(lines[sep2_index], max_chars))
    next_index = sep2_index + 1
    if next_index < len(lines):
        next_line = lines[next_index]
        if (
            next_line.strip() != ""
            and not _is_section_separator(next_line)
            and not COMMAND_LINE_PATTERN.match(next_line)
        ):
            block_segments.extend(_wrap_single_line(next_line, max_chars))
            return block_segments, (next_index + 1)
    return block_segments, next_index


def _header_with_trailing_separator_segments(
    lines: list[str],
    start_index: int,
    max_chars: int,
) -> tuple[list[tuple[str, bool]], int] | None:
    if _is_section_separator(lines[start_index]) or lines[start_index].strip() == "":
        return None

    max_header_lines = 4
    scan_limit = min(len(lines), start_index + max_header_lines + 1)
    for sep_index in range(start_index + 1, scan_limit):
        candidate = lines[sep_index]
        if _is_section_separator(candidate):
            body_lines = lines[start_index:sep_index]
            if any(line.strip() == "" for line in body_lines):
                return None
            if any(COMMAND_LINE_PATTERN.match(line) for line in body_lines):
                return None

            block_segments: list[tuple[str, bool]] = []
            for idx in range(start_index, sep_index + 1):
                block_segments.extend(_wrap_single_line(lines[idx], max_chars))
            next_index = sep_index + 1
            if next_index < len(lines):
                next_line = lines[next_index]
                if (
                    next_line.strip() != ""
                    and not _is_section_separator(next_line)
                    and not COMMAND_LINE_PATTERN.match(next_line)
                ):
                    block_segments.extend(_wrap_single_line(next_line, max_chars))
                    return block_segments, (next_index + 1)
            return block_segments, next_index

        if candidate.strip() == "":
            return None
        if COMMAND_LINE_PATTERN.match(candidate):
            return None

    return None


def _show_clock_block_segments(
    lines: list[str],
    start_index: int,
    max_chars: int,
) -> tuple[list[tuple[str, bool]], int] | None:
    if not _is_show_clock_command(lines[start_index]):
        return None

    next_index = start_index + 1
    if next_index >= len(lines):
        return None

    block_segments = _wrap_single_line(lines[start_index], max_chars)
    next_line = lines[next_index]

    # Some logs contain a blank line between "show clock" and its time output.
    if next_line.strip() == "" and (next_index + 1) < len(lines):
        time_line = lines[next_index + 1]
        if _parse_clock_time_line(time_line) is not None:
            block_segments.extend(_wrap_single_line(next_line, max_chars))
            block_segments.extend(_wrap_single_line(time_line, max_chars))
            return block_segments, (next_index + 2)

    if _parse_clock_time_line(next_line) is not None:
        block_segments.extend(_wrap_single_line(next_line, max_chars))
        return block_segments, (next_index + 1)

    # Fallback to previous behavior: keep command with immediate next line.
    block_segments.extend(_wrap_single_line(next_line, max_chars))
    return block_segments, (next_index + 1)


def _text_layout_params() -> tuple[int, int, int, int, int, int]:
    page_w, page_h = A4_PAGE_W, A4_PAGE_H
    margin_x = 24
    margin_top = 32
    margin_bottom = 32

    usable_w = page_w - (margin_x * 2)
    usable_h = page_h - margin_top - margin_bottom
    chars_per_line = max(20, int(usable_w // PDF_CHAR_WIDTH_ESTIMATE))
    lines_per_page = max(20, int(usable_h // PDF_BODY_LINE_HEIGHT))
    return (page_w, page_h, margin_x, margin_top, PDF_BODY_LINE_HEIGHT, lines_per_page, chars_per_line)


def _paginate_wrapped_lines(lines: list[str]) -> tuple[list[list[tuple[str, bool]]], int, int, int, int, int]:
    page_w, page_h, margin_x, margin_top, line_h, lines_per_page, max_chars = _text_layout_params()
    pages: list[list[tuple[str, bool]]] = []
    current_page: list[tuple[str, bool]] = []

    def flush_page() -> None:
        nonlocal current_page
        if current_page:
            pages.append(current_page)
            current_page = []

    def append_segments(
        segments: list[tuple[str, bool]],
        keep_together: bool = False,
    ) -> None:
        if not segments:
            return
        if keep_together and len(segments) <= lines_per_page:
            remaining_slots = lines_per_page - len(current_page)
            if remaining_slots < len(segments) and current_page:
                flush_page()

        pending = list(segments)
        while pending:
            if len(current_page) == lines_per_page:
                flush_page()
            free_slots = lines_per_page - len(current_page)
            take = pending[:free_slots]
            current_page.extend(take)
            pending = pending[free_slots:]
            if len(current_page) == lines_per_page:
                flush_page()

    i = 0
    while i < len(lines):
        line = lines[i]

        separator_block = _separator_block_segments(lines, i, max_chars)
        if separator_block is not None:
            block_segments, next_index = separator_block
            if len(block_segments) <= lines_per_page:
                append_segments(block_segments, keep_together=True)
                i = next_index
                continue

        line_separator_header_separator = _line_separator_header_separator_segments(lines, i, max_chars)
        if line_separator_header_separator is not None:
            block_segments, next_index = line_separator_header_separator
            if len(block_segments) <= lines_per_page:
                append_segments(block_segments, keep_together=True)
                i = next_index
                continue

        header_with_separator = _header_with_trailing_separator_segments(lines, i, max_chars)
        if header_with_separator is not None:
            block_segments, next_index = header_with_separator
            if len(block_segments) <= lines_per_page:
                append_segments(block_segments, keep_together=True)
                i = next_index
                continue

        line_with_separator = _line_with_trailing_separator_segments(lines, i, max_chars)
        if line_with_separator is not None:
            pair_segments, next_index = line_with_separator
            if len(pair_segments) <= lines_per_page:
                append_segments(pair_segments, keep_together=True)
                i = next_index
                continue

        show_clock_block = _show_clock_block_segments(lines, i, max_chars)
        if show_clock_block is not None:
            block_segments, next_index = show_clock_block
            if len(block_segments) <= lines_per_page:
                append_segments(block_segments, keep_together=True)
                i = next_index
                continue

        line_segments = _wrap_single_line(line, max_chars)
        append_segments(line_segments)

        i += 1

    flush_page()
    if not pages:
        pages = [[("", False)]]
    return pages, page_w, page_h, margin_x, margin_top, line_h


def render_text_pages(lines: list[str]) -> list[Image.Image]:
    font = load_monospace_font(size=PDF_BODY_FONT_SIZE)
    pages_data, page_w, page_h, margin_x, margin_top, line_h = _paginate_wrapped_lines(lines)

    pages: list[Image.Image] = []
    for page_lines in pages_data:
        page = Image.new("RGB", (page_w, page_h), color=PAGE_WHITE)
        draw = ImageDraw.Draw(page)
        y = margin_top
        for text_line, is_command in page_lines:
            if is_command and text_line.strip():
                bbox = draw.textbbox((margin_x, y), text_line, font=font)
                x1 = margin_x - 2
                y1 = y + 1
                x2 = min(page_w - margin_x, bbox[2] + 3)
                y2 = y + line_h - 2
                draw.rectangle((x1, y1, x2, y2), fill=HIGHLIGHT_YELLOW)
            draw.text((margin_x, y), text_line, font=font, fill=TEXT_BLACK)
            y += line_h
        pages.append(page)
    return pages


def _load_image_rgb(image_input: Path | bytes) -> Image.Image:
    if isinstance(image_input, Path):
        with Image.open(image_input) as src:
            return src.convert("RGB")
    with Image.open(io.BytesIO(image_input)) as src:
        return src.convert("RGB")


def load_heading_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\segoeuib.ttf"),
        Path(r"C:\Windows\Fonts\arialbd.ttf"),
        Path(r"C:\Windows\Fonts\segoeui.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return load_monospace_font(size=size)


def render_image_page(image_input: Path | bytes) -> Image.Image:
    # Keep image at original pixel size to avoid blur from downscaling.
    margin_x = 12
    margin_y = 6
    title_text = "Show log"
    title_font = load_monospace_font(size=PDF_BODY_FONT_SIZE)
    img = _load_image_rgb(image_input)

    probe = ImageDraw.Draw(Image.new("RGB", (10, 10), color=PAGE_WHITE))
    title_box = probe.textbbox((0, 0), title_text, font=title_font)
    title_w = title_box[2] - title_box[0]
    title_h = title_box[3] - title_box[1]

    title_pad_x = 2
    title_pad_y = 1
    header_h = (margin_y * 2) + title_h + (title_pad_y * 2)

    page_w = img.width
    page_h = img.height + header_h
    page = Image.new("RGB", (page_w, page_h), color=PAGE_WHITE)
    draw = ImageDraw.Draw(page)

    title_x = margin_x
    title_y = margin_y + title_pad_y
    draw.rectangle(
        (
            title_x - title_pad_x,
            title_y - title_pad_y,
            title_x + title_w + title_pad_x,
            title_y + title_h + title_pad_y,
        ),
        fill=HIGHLIGHT_YELLOW,
    )
    draw.text((title_x, title_y), title_text, font=title_font, fill=TEXT_BLACK)

    page.paste(img, (0, header_h))
    return page


def build_pdf_pages(combined_lines: list[str], image_input: Path | bytes) -> list[Image.Image]:
    pages = render_text_pages(combined_lines)
    pages.append(render_image_page(image_input))
    return pages


def _sanitize_xml_text(text: str) -> str:
    # Keep printable characters only to avoid invalid XML in .docx.
    sanitized = []
    for ch in text:
        code = ord(ch)
        if ch in ("\t", "\n", "\r") or code >= 32:
            sanitized.append(ch)
        else:
            sanitized.append(" ")
    return "".join(sanitized)


def _docx_image_part(image_input: Path | bytes) -> tuple[str, str, bytes, tuple[int, int]]:
    raw = image_input.read_bytes() if isinstance(image_input, Path) else image_input
    with Image.open(io.BytesIO(raw)) as src:
        size = src.size
        fmt = (src.format or "").upper()
        if fmt in ("JPG", "JPEG"):
            return "jpg", "image/jpeg", raw, size
        if fmt == "PNG":
            return "png", "image/png", raw, size

        out = io.BytesIO()
        src.convert("RGB").save(out, format="PNG")
        return "png", "image/png", out.getvalue(), size


def _docx_run_xml(text: str, highlight: bool = False, bold: bool = False) -> str:
    escaped = xml_escape(_sanitize_xml_text(text))
    rpr_parts = [
        '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas"/>',
        '<w:sz w:val="18"/>',
        '<w:szCs w:val="18"/>',
    ]
    if bold:
        rpr_parts.append("<w:b/>")
    if highlight:
        rpr_parts.append('<w:highlight w:val="yellow"/>')
    rpr = "".join(rpr_parts)
    return f"<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space=\"preserve\">{escaped}</w:t></w:r>"


def _docx_image_inline_xml(rel_id: str, width_px: int, height_px: int) -> str:
    emu_per_px = 9525
    max_width_emu = int(6.7 * 914400)

    cx = max(1, int(width_px * emu_per_px))
    cy = max(1, int(height_px * emu_per_px))
    if cx > max_width_emu:
        scale = max_width_emu / cx
        cx = max_width_emu
        cy = max(1, int(cy * scale))

    return f"""
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:docPr id="1" name="ShowLogImage"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks noChangeAspect="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
        </wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="showlog-image"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="{rel_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="{cx}" cy="{cy}"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
""".strip()


def build_docx_bytes(combined_lines: list[str], image_input: Path | bytes) -> bytes:
    img_ext, img_content_type, img_bytes, (img_w, img_h) = _docx_image_part(image_input)

    paragraph_xml: list[str] = []
    for line in combined_lines:
        if line == "":
            paragraph_xml.append("<w:p/>")
            continue
        is_command = bool(COMMAND_LINE_PATTERN.match(line))
        paragraph_xml.append(f"<w:p>{_docx_run_xml(line, highlight=is_command)}</w:p>")

    paragraph_xml.append("<w:p/>")
    paragraph_xml.append(f"<w:p>{_docx_run_xml('Show log', highlight=True)}</w:p>")
    paragraph_xml.append(_docx_image_inline_xml("rId1", img_w, img_h))

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 wp14">
  <w:body>
    {''.join(paragraph_xml)}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

    document_rels_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.{img_ext}"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>
"""

    content_types_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="{img_ext}" ContentType="{img_content_type}"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

    rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

    output = io.BytesIO()
    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/_rels/document.xml.rels", document_rels_xml)
        zf.writestr(f"word/media/image1.{img_ext}", img_bytes)
    return output.getvalue()


def _jpeg_bytes_from_image(image: Image.Image, quality: int = 95) -> bytes:
    rgb = image.convert("RGB")
    out = io.BytesIO()
    rgb.save(out, format="JPEG", quality=quality, subsampling=0, optimize=False)
    return out.getvalue()


def _jpeg_bytes_and_size_from_input(image_input: Path | bytes) -> tuple[bytes, tuple[int, int]]:
    raw_bytes: bytes | None = None
    if isinstance(image_input, Path):
        raw_bytes = image_input.read_bytes()
    else:
        raw_bytes = image_input

    with Image.open(io.BytesIO(raw_bytes)) as src:
        size = src.size
        if src.format == "JPEG" and src.mode == "RGB":
            return raw_bytes, size
        return _jpeg_bytes_from_image(src, quality=100), size


def _stream_object(stream: bytes, header: str) -> bytes:
    return f"{header}\nstream\n".encode("ascii") + stream + b"\nendstream"


def _escape_pdf_text(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _pdf_text_literal(text: str) -> str:
    safe = text.replace("\t", "    ")
    safe = "".join(ch if ord(ch) <= 255 else "?" for ch in safe)
    return _escape_pdf_text(safe)


def _serialize_pdf(objects: list[bytes], root_obj_num: int) -> bytes:
    out = bytearray()
    out.extend(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")

    offsets = [0]
    for idx, obj in enumerate(objects, start=1):
        offsets.append(len(out))
        out.extend(f"{idx} 0 obj\n".encode("ascii"))
        out.extend(obj)
        out.extend(b"\nendobj\n")

    xref_pos = len(out)
    out.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    out.extend(b"0000000000 65535 f \n")
    for off in offsets[1:]:
        out.extend(f"{off:010d} 00000 n \n".encode("ascii"))

    trailer = (
        f"trailer\n<< /Size {len(objects) + 1} /Root {root_obj_num} 0 R >>\n"
        f"startxref\n{xref_pos}\n%%EOF\n"
    )
    out.extend(trailer.encode("ascii"))
    return bytes(out)


def _build_pdf_with_native_image_page(
    combined_lines: list[str],
    image_input: Path | bytes,
) -> bytes:
    objects: list[bytes] = []

    def add_obj(data: bytes) -> int:
        objects.append(data)
        return len(objects)

    font_body_obj_num = add_obj(b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>")
    font_title_obj_num = add_obj(b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>")
    pages_obj_num = add_obj(b"<< /Type /Pages /Count 0 /Kids [] >>")
    catalog_obj_num = add_obj(b"<< /Type /Catalog /Pages 3 0 R >>")

    kids: list[int] = []

    text_pages, page_w, page_h, margin_x, margin_top, line_h = _paginate_wrapped_lines(combined_lines)
    body_font_size = PDF_BODY_FONT_SIZE
    for page_lines in text_pages:
        y_start = page_h - margin_top - body_font_size
        content_ops: list[str] = []
        for idx, (text_line, is_command) in enumerate(page_lines):
            y = y_start - (idx * line_h)
            if y < 0:
                break
            if is_command and text_line.strip():
                rect_x = margin_x - 2
                rect_y = y - 2
                rect_h = max(10, line_h - 2)
                est_w = int((len(text_line) * body_font_size * 0.60) + 6)
                max_w = page_w - margin_x - rect_x
                rect_w = max(10, min(max_w, est_w))
                content_ops.append("1.0 0.9569 0.5098 rg")
                content_ops.append(f"{rect_x} {rect_y} {rect_w} {rect_h} re f")

            content_ops.append("0 0 0 rg")
            content_ops.append(
                f"BT /F1 {body_font_size} Tf {margin_x} {y} Td ({_pdf_text_literal(text_line)}) Tj ET"
            )

        content = ("\n".join(content_ops) + "\n").encode("latin-1", "replace")
        content_header = f"<< /Length {len(content)} >>"
        content_obj_num = add_obj(_stream_object(content, content_header))

        page_obj = (
            f"<< /Type /Page /Parent {pages_obj_num} 0 R "
            f"/MediaBox [0 0 {page_w} {page_h}] "
            f"/Resources << /ProcSet [/PDF /Text] "
            f"/Font << /F1 {font_body_obj_num} 0 R /F2 {font_title_obj_num} 0 R >> >> "
            f"/Contents {content_obj_num} 0 R >>"
        ).encode("ascii")
        kids.append(add_obj(page_obj))

    image_jpeg, (img_w, img_h) = _jpeg_bytes_and_size_from_input(image_input)
    page_w, page_h = (A4_PAGE_W, A4_PAGE_H)

    # A4 portrait layout like text pages: title near top-left, image right below.
    top_margin = 60
    side_margin = 24
    bottom_margin = 24
    title_gap = 18
    image_header = (
        f"<< /Type /XObject /Subtype /Image /Width {img_w} /Height {img_h} "
        f"/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode "
        f"/Length {len(image_jpeg)} >>"
    )
    image_obj_num = add_obj(_stream_object(image_jpeg, image_header))

    title_text = "Show log"
    font_size = PDF_BODY_FONT_SIZE
    title_tab_indent = 36
    title_x = side_margin + title_tab_indent
    title_baseline_from_top = top_margin + font_size
    title_y = page_h - title_baseline_from_top
    estimated_w = int((len(title_text) * font_size * 0.60) + 6)
    rect_x = title_x - 2
    rect_y = title_y - 2
    rect_w = estimated_w
    rect_h = max(10, PDF_BODY_LINE_HEIGHT - 2)
    escaped_title = _escape_pdf_text(title_text)

    image_max_w = page_w - (side_margin * 2)
    image_max_h = page_h - (top_margin + rect_h + title_gap + bottom_margin)
    image_scale = min(image_max_w / img_w, image_max_h / img_h)
    image_scale = max(0.01, image_scale)
    draw_w = int(img_w * image_scale)
    draw_h = int(img_h * image_scale)
    draw_x = (page_w - draw_w) // 2
    image_top_y = page_h - (top_margin + rect_h + title_gap)
    draw_y = image_top_y - draw_h

    content_lines = [
        "1.0 0.9569 0.5098 rg",
        f"{rect_x} {rect_y} {rect_w} {rect_h} re f",
        "0 0 0 rg",
        f"BT /F2 {font_size} Tf {title_x} {title_y} Td ({escaped_title}) Tj ET",
        "q",
        f"{draw_w} 0 0 {draw_h} {draw_x} {draw_y} cm",
        "/Im0 Do",
        "Q",
        "",
    ]
    image_content = "\n".join(content_lines).encode("latin-1", "replace")
    image_content_header = f"<< /Length {len(image_content)} >>"
    image_content_obj_num = add_obj(_stream_object(image_content, image_content_header))

    image_page_obj = (
        f"<< /Type /Page /Parent {pages_obj_num} 0 R "
        f"/MediaBox [0 0 {page_w} {page_h}] "
        f"/Resources << /ProcSet [/PDF /Text /ImageC] "
        f"/Font << /F1 {font_body_obj_num} 0 R /F2 {font_title_obj_num} 0 R >> "
        f"/XObject << /Im0 {image_obj_num} 0 R >> >> "
        f"/Contents {image_content_obj_num} 0 R >>"
    ).encode("ascii")
    image_page_obj_num = add_obj(image_page_obj)
    kids.append(image_page_obj_num)

    kids_refs = " ".join(f"{num} 0 R" for num in kids)
    objects[pages_obj_num - 1] = (
        f"<< /Type /Pages /Count {len(kids)} /Kids [{kids_refs}] >>".encode("ascii")
    )
    objects[catalog_obj_num - 1] = (
        f"<< /Type /Catalog /Pages {pages_obj_num} 0 R >>".encode("ascii")
    )

    return _serialize_pdf(objects, root_obj_num=catalog_obj_num)


def pages_to_pdf_bytes(
    pages: Iterable[Image.Image],
    image_input: Path | bytes | None = None,
    combined_lines: list[str] | None = None,
) -> bytes:
    page_list = list(pages)
    if not page_list:
        raise ValueError("No pages to convert.")

    if image_input is not None and combined_lines is not None:
        return _build_pdf_with_native_image_page(combined_lines, image_input)

    first_page, other_pages = page_list[0], page_list[1:]
    buffer = io.BytesIO()
    first_page.save(
        buffer,
        "PDF",
        save_all=True,
        append_images=other_pages,
        resolution=72.0,
        quality=100,
        subsampling=0,
    )
    return buffer.getvalue()


def build_pdf_bytes(combined_lines: list[str], image_input: Path | bytes) -> bytes:
    pages = build_pdf_pages(combined_lines, image_input)
    return pages_to_pdf_bytes(pages, image_input=image_input, combined_lines=combined_lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge FDO + APIC logs and append screenshot into a single output file."
    )
    parser.add_argument("--fdo", type=Path, default=DEFAULT_FDO)
    parser.add_argument("--apic", type=Path, default=DEFAULT_APIC)
    parser.add_argument("--image", type=Path, default=DEFAULT_IMAGE)
    parser.add_argument("--outdir", type=Path, default=DEFAULT_OUTDIR)
    parser.add_argument("--format", choices=("pdf", "docx"), default="pdf")
    parser.add_argument("--pdf-name")
    parser.add_argument("--docx-name")
    parser.add_argument("--text-name")
    args = parser.parse_args()

    for path in (args.fdo, args.apic, args.image):
        if not path.exists():
            raise FileNotFoundError(f"Missing input file: {path}")

    args.outdir.mkdir(parents=True, exist_ok=True)
    base_name = args.fdo.stem
    out_pdf = args.outdir / (args.pdf_name or f"{base_name}.pdf")
    out_docx = args.outdir / (args.docx_name or f"{base_name}.docx")
    out_text = args.outdir / (args.text_name or f"{base_name}.txt")

    fdo_text = read_text_with_fallback(args.fdo)
    apic_text = read_text_with_fallback(args.apic)
    combined_lines = build_combined_lines(fdo_text, apic_text)
    out_text.write_text("\n".join(combined_lines) + "\n", encoding="utf-8")

    print(f"Created text file: {out_text}")
    if args.format == "pdf":
        pdf_bytes = build_pdf_bytes(combined_lines, args.image)
        out_pdf.write_bytes(pdf_bytes)
        text_pages, *_ = _paginate_wrapped_lines(combined_lines)
        print(f"Created PDF file:  {out_pdf}")
        print(f"Total pages:       {len(text_pages) + 1}")
    else:
        out_docx.write_bytes(build_docx_bytes(combined_lines, args.image))
        print(f"Created DOCX file: {out_docx}")


if __name__ == "__main__":
    main()
