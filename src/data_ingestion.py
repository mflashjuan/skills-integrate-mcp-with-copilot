import json
import re
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
from docx import Document


def _normalize_header(value: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(value).strip().lower())


def _parse_participants(value: Any) -> List[str]:
    if value is None:
        return []

    if isinstance(value, list):
        values = value
    else:
        values = [value]

    participants: List[str] = []
    for item in values:
        if item is None:
            continue
        text = str(item).strip()
        if not text:
            continue

        # Split common delimiters
        parts = re.split(r'[;,\n]+', text)
        participants.extend(part.strip() for part in parts if part.strip())

    return participants


def _parse_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _activity_record_from_row(row: Dict[str, Any], header_map: Dict[str, str], fallback_name: str) -> Dict[str, Any]:
    activity_name = row.get(header_map.get('activityname', ''), fallback_name)
    description = row.get(header_map.get('description', ''), '') or ''
    schedule = row.get(header_map.get('schedule', ''), '') or ''
    max_participants = _parse_int(row.get(header_map.get('maxparticipants', ''), 0), 0)
    participants = _parse_participants(row.get(header_map.get('participants', ''), ''))

    if not activity_name:
        activity_name = fallback_name

    return {
        'activity_name': str(activity_name),
        'description': str(description),
        'schedule': str(schedule),
        'max_participants': max_participants,
        'participants': participants,
    }


def _header_map(columns: List[str]) -> Dict[str, str]:
    return {_normalize_header(col): col for col in columns}


def _generate_unique_name(name: str, existing: Dict[str, Any]) -> str:
    candidate = name
    index = 2
    while candidate in existing:
        candidate = f"{name} ({index})"
        index += 1
    return candidate


def parse_excel_file(file_path: Path) -> Dict[str, Any]:
    workbook = pd.ExcelFile(file_path)
    activities: Dict[str, Any] = {}

    for sheet_name in workbook.sheet_names:
        sheet_name_safe = str(sheet_name or 'Untitled Activity')
        df = workbook.parse(sheet_name)
        if df.empty:
            continue

        headers = [str(c) for c in df.columns]
        header_map = _header_map(headers)

        # When a sheet contains a single activity with participant rows
        if any(key in header_map for key in ('email', 'participant', 'studentemail')):
            participants = []
            rows = df.to_dict(orient='records')
            for row in rows:
                for key in ('email', 'participant', 'studentemail'):
                    if key in header_map:
                        participants.extend(_parse_participants(row.get(header_map[key], '')))
                        break

            metadata = rows[0] if rows else {}
            activity_name = sheet_name_safe
            unique_name = _generate_unique_name(activity_name, activities)
            activities[unique_name] = {
                'description': str(metadata.get(header_map.get('description', ''), '')) if 'description' in header_map else sheet_name_safe,
                'schedule': str(metadata.get(header_map.get('schedule', ''), '')) if 'schedule' in header_map else '',
                'max_participants': _parse_int(metadata.get(header_map.get('maxparticipants', ''), 0), 0) if 'maxparticipants' in header_map else 0,
                'participants': participants,
            }
            continue

        # When a sheet contains multiple activity entries
        for row in df.to_dict(orient='records'):
            entry = _activity_record_from_row(row, header_map, sheet_name_safe)
            activity_name = _generate_unique_name(entry['activity_name'], activities)
            activities[activity_name] = {
                'description': entry['description'],
                'schedule': entry['schedule'],
                'max_participants': entry['max_participants'],
                'participants': entry['participants'],
            }

    return activities


def parse_docx_file(file_path: Path) -> Dict[str, Any]:
    document = Document(file_path)
    activities: Dict[str, Any] = {}
    fallback_name = file_path.stem.replace('_', ' ')

    for table in document.tables:
        rows = [list(cell.text for cell in row.cells) for row in table.rows]
        if len(rows) < 2:
            continue

        headers = rows[0]
        header_map = _header_map(headers)

        if any(key in header_map for key in ('activityname', 'description', 'schedule', 'maxparticipants', 'participants')):
            for row_cells in rows[1:]:
                row_data = {headers[idx]: value for idx, value in enumerate(row_cells) if idx < len(headers)}
                entry = _activity_record_from_row(row_data, header_map, fallback_name)
                activity_name = _generate_unique_name(entry['activity_name'], activities)
                activities[activity_name] = {
                    'description': entry['description'],
                    'schedule': entry['schedule'],
                    'max_participants': entry['max_participants'],
                    'participants': entry['participants'],
                }
        elif any(key in header_map for key in ('participant', 'email', 'studentemail')):
            participants: List[str] = []
            for row_cells in rows[1:]:
                row_data = {headers[idx]: value for idx, value in enumerate(row_cells) if idx < len(headers)}
                for key in ('participant', 'email', 'studentemail'):
                    if key in header_map:
                        participants.extend(_parse_participants(row_data.get(header_map[key], '')))
                        break

            activity_name = _generate_unique_name(fallback_name, activities)
            activities[activity_name] = {
                'description': '',
                'schedule': '',
                'max_participants': 0,
                'participants': participants,
            }

    return activities


def build_activity_data(source_path: Path, output_path: Path) -> Dict[str, Any]:
    source_path = source_path.expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"Source path not found: {source_path}")

    activities: Dict[str, Any] = {}
    candidates = []

    if source_path.is_dir():
        candidates = sorted([path for path in source_path.rglob('*') if path.suffix.lower() in {'.xlsx', '.xls', '.docx'}])
    elif source_path.suffix.lower() in {'.xlsx', '.xls', '.docx'}:
        candidates = [source_path]
    else:
        raise ValueError("Input path must be an Excel file, DOCX file, or folder containing them.")

    for source_file in candidates:
        if source_file.suffix.lower() in {'.xlsx', '.xls'}:
            parsed = parse_excel_file(source_file)
        else:
            parsed = parse_docx_file(source_file)

        for key, value in parsed.items():
            unique_key = _generate_unique_name(key, activities)
            activities[unique_key] = value

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open('w', encoding='utf-8') as handle:
        json.dump(activities, handle, ensure_ascii=False, indent=2)

    return activities
