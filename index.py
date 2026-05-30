import os
import re
import zipfile
import xml.etree.ElementTree as ET

from flask import Flask, jsonify, render_template_string, request, send_from_directory

app = Flask(__name__)

SPREADSHEET_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def column_to_number(column):
    value = 0
    for char in column:
        value = value * 26 + ord(char) - 64
    return value


def split_cell_ref(ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    return column_to_number(match.group(1)), int(match.group(2))


def excel_time_to_seconds(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    if 0 < number < 1:
        return f"{number * 86400:.2f}".rstrip("0").rstrip(".")
    return f"{number:g}"


def clean_athlete_name(value):
    value = str(value or "").strip()
    if not value or value in {"Tempo", "HR", "Average", "BestAve", "AllAve", "達成率", "標準偏差"}:
        return ""
    if value.startswith("※") or value.startswith("("):
        return ""
    if value in {"-", "#VALUE!", "#DIV/0!"}:
        return ""
    if value in {"Fr", "Fly", "Ba", "Br", "IM", "Choice"}:
        return ""
    if "/" in value:
        return ""
    try:
        float(value)
        return ""
    except ValueError:
        pass
    return value


def read_xlsx_grid(file_bytes):
    with zipfile.ZipFile(file_bytes) as workbook:
        shared_strings = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", SPREADSHEET_NS):
                shared_strings.append("".join(t.text or "" for t in item.findall(".//a:t", SPREADSHEET_NS)))

        sheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        grid = {}
        for cell in sheet.findall(".//a:c", SPREADSHEET_NS):
            value_node = cell.find("a:v", SPREADSHEET_NS)
            if value_node is None:
                continue
            value = value_node.text or ""
            if cell.attrib.get("t") == "s":
                value = shared_strings[int(value)]
            column, row = split_cell_ref(cell.attrib["r"])
            grid[(row, column)] = value
        return grid


def looks_like_repetition(value):
    return str(value).strip() in {"1", "2", "3", "4", "5", "6", "7", "8"}


def parse_manager_workbook(file_storage):
    file_storage.stream.seek(0)
    grid = read_xlsx_grid(file_storage.stream)
    rows = [row for row, _ in grid]
    cols = [col for _, col in grid]
    if not rows or not cols:
        return {"headers": [], "rows": []}

    min_row, max_row = min(rows), max(rows)
    min_col, max_col = min(cols), max(cols)
    group_markers = []
    for row in range(min_row, max_row + 1):
        value = str(grid.get((row, min_col), "")).strip()
        if value in {"(A)", "(BC)", "(D)"}:
            group_markers.append((row, value.strip("()")))
    group_markers.append((max_row + 1, "END"))

    athletes = {}
    for index, (start_row, group) in enumerate(group_markers[:-1]):
        end_row = group_markers[index + 1][0]
        header_rows = []
        for header_row in range(start_row + 1, end_row):
            header_values = [clean_athlete_name(grid.get((header_row, col), "")) for col in range(min_col + 1, max_col + 1)]
            if not any(header_values):
                continue
            candidate_columns = [min_col + 1 + offset for offset, name in enumerate(header_values) if name]
            if not any(find_label_column(grid, column, header_row + 1, end_row) for column in candidate_columns):
                continue
            header_rows.append((header_row, header_values))

        for header_index, (header_row, header_values) in enumerate(header_rows):
            block_end_row = header_rows[header_index + 1][0] if header_index + 1 < len(header_rows) else end_row
            athlete_columns = [(min_col + 1 + offset, name) for offset, name in enumerate(header_values) if name]
            if not athlete_columns:
                continue
            for column, name in athlete_columns:
                athlete = athletes.setdefault(name, {"name": name, "group": group, "values": {}})
                if group == "D":
                    collect_d_athlete_column(grid, athlete, column, header_row + 1, block_end_row)
                else:
                    label_column = find_label_column(grid, column, header_row + 1, block_end_row)
                    if label_column:
                        collect_athlete_column(grid, athlete, label_column, column, header_row + 1, block_end_row)

    headers = []
    for athlete in athletes.values():
        headers.extend(athlete["values"].keys())
    headers = sorted(set(headers), key=lambda item: (item.split()[0], item))

    return {
        "headers": headers,
        "rows": [
            {
                "name": athlete["name"],
                "group": athlete["group"],
                "values": [{"label": label, "value": athlete["values"].get(label, "")} for label in headers],
            }
            for athlete in athletes.values()
        ],
    }


def find_label_column(grid, data_column, start_row, end_row):
    for label_column in range(data_column - 1, max(0, data_column - 5), -1):
        labels = [str(grid.get((row, label_column), "")).strip() for row in range(start_row, min(end_row, start_row + 10))]
        rep_count = sum(1 for label in labels if looks_like_repetition(label))
        has_summary = any(label in {"HR", "Average", "BestAve"} for label in labels)
        if rep_count >= 2 or (rep_count >= 1 and has_summary):
            return label_column
    return None


def collect_athlete_column(grid, athlete, label_column, data_column, start_row, end_row):
    current_set = 1
    seen_repetition = False
    for row in range(start_row, end_row):
        label = str(grid.get((row, label_column), "")).strip()
        value = grid.get((row, data_column), "")
        if value in {"", None}:
            continue
        if label == "HR":
            athlete["values"][f"S{current_set} HR"] = excel_time_to_seconds(value)
        elif label == "BestAve":
            athlete["values"]["BestAve"] = excel_time_to_seconds(value)
        elif label == "Average":
            if seen_repetition:
                current_set += 1
                seen_repetition = False
        elif looks_like_repetition(label):
            athlete["values"][f"S{current_set} {label}"] = excel_time_to_seconds(value)
            seen_repetition = True


def collect_d_athlete_column(grid, athlete, name_column, start_row, end_row):
    if any(looks_like_repetition(grid.get((row, name_column), "")) for row in range(start_row, min(start_row + 8, end_row))):
        label_column = name_column
        data_column = name_column + 1
    else:
        label_column = max(1, name_column - 1)
        data_column = name_column

    current_set = 1
    for row in range(start_row, end_row):
        label = str(grid.get((row, label_column), "")).strip()
        value = grid.get((row, data_column), "")
        if label == "HR" and value not in {"", None}:
            athlete["values"][f"S{current_set} HR"] = excel_time_to_seconds(value)
        elif label == "BestAve" and value not in {"", None}:
            athlete["values"]["BestAve"] = excel_time_to_seconds(value)
        elif label == "Average":
            current_set += 1
        elif looks_like_repetition(label):
            if value not in {"", None}:
                athlete["values"][f"S{current_set} {label} 25m"] = excel_time_to_seconds(value)
            seventy_five = grid.get((row + 1, data_column), "")
            if seventy_five not in {"", None}:
                athlete["values"][f"S{current_set} {label} 75m"] = excel_time_to_seconds(seventy_five)


PAGE = r"""
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>水泳データ集計</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --line: #d8dde6;
      --line-strong: #aab4c2;
      --text: #18212f;
      --muted: #657386;
      --accent: #0077b6;
      --accent-dark: #005f91;
      --ok: #087f5b;
      --warn: #b35c00;
      --danger: #b42318;
      --shadow: 0 1px 2px rgba(16, 24, 40, 0.08);
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: var(--bg);
      color: var(--text);
      overflow-x: hidden;
    }

    button, input, select, textarea {
      font: inherit;
    }

    button {
      border: 1px solid var(--line-strong);
      background: #fff;
      color: var(--text);
      border-radius: 6px;
      min-height: 34px;
      padding: 0 12px;
      cursor: pointer;
      white-space: nowrap;
    }

    button.primary {
      border-color: var(--accent);
      background: var(--accent);
      color: #fff;
    }

    button.primary:hover { background: var(--accent-dark); }
    button:hover { border-color: var(--accent); }
    button.icon { width: 34px; padding: 0; }
    button.danger { color: var(--danger); border-color: #f0b6af; }

    input, select, textarea {
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--text);
    }

    input, select {
      height: 34px;
      padding: 0 9px;
    }

    textarea {
      padding: 10px;
      resize: vertical;
      min-height: 120px;
      width: 100%;
      line-height: 1.5;
    }

    .app-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
      min-height: 60px;
      padding: 0 20px;
      background: #0f2538;
      color: #fff;
      border-bottom: 1px solid #0a1c2c;
    }

    .app-header h1 {
      margin: 0;
      font-size: 20px;
      letter-spacing: 0;
    }

    .header-actions {
      display: flex;
      gap: 8px;
      align-items: center;
    }

    .header-actions button {
      background: #17344d;
      color: #fff;
      border-color: #31516d;
    }

    .layout {
      display: grid;
      grid-template-columns: 260px minmax(0, 1fr);
      min-height: calc(100vh - 60px);
    }

    aside {
      background: #fff;
      border-right: 1px solid var(--line);
      padding: 16px;
      overflow: auto;
    }

    main {
      padding: 18px;
      overflow-y: auto;
      overflow-x: hidden;
      min-width: 0;
    }

    .section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      margin-bottom: 16px;
      min-width: 0;
    }

    .section-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 12px 14px;
      border-bottom: 1px solid var(--line);
    }

    .section-header h2 {
      margin: 0;
      font-size: 16px;
      letter-spacing: 0;
    }

    .section-body {
      padding: 14px;
      min-width: 0;
    }

    .view {
      min-width: 0;
    }

    .tabs {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }

    .tab {
      justify-content: flex-start;
      text-align: left;
      width: 100%;
      border-color: transparent;
      background: transparent;
    }

    .tab.active {
      background: #e8f3fb;
      border-color: #b7d9ee;
      color: #064d74;
      font-weight: 700;
    }

    .field-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(140px, 1fr));
      gap: 10px;
      align-items: end;
    }

    .settings-panel {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      margin-bottom: 12px;
      background: #fbfcfe;
    }

    .settings-panel h3 {
      margin: 0 0 8px;
      font-size: 13px;
    }

    .summary-setting-group {
      border-top: 1px solid var(--line);
      padding-top: 10px;
      margin-top: 10px;
    }

    .summary-setting-group:first-child {
      border-top: 0;
      padding-top: 0;
      margin-top: 0;
    }

    .summary-setting-group h4 {
      margin: 0 0 8px;
      font-size: 12px;
      color: var(--text);
    }

    .check-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 8px 12px;
    }

    .check-grid label {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: var(--text);
      font-weight: 600;
    }

    .check-grid input {
      width: auto;
    }

    label {
      display: grid;
      gap: 5px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }

    label span {
      overflow-wrap: anywhere;
    }

    .toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }

    .quick-guide {
      display: grid;
      grid-template-columns: repeat(3, minmax(160px, 1fr));
      gap: 10px;
      margin-bottom: 16px;
    }

    .guide-item {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      padding: 12px;
      box-shadow: var(--shadow);
    }

    .guide-item strong {
      display: block;
      margin-bottom: 4px;
      font-size: 14px;
    }

    .group-switch {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 6px;
    }

    .group-switch button {
      min-height: 40px;
      font-weight: 700;
    }

    .group-switch button.active {
      background: var(--accent);
      border-color: var(--accent);
      color: #fff;
    }

    .measure-mode-switch {
      display: grid;
      gap: 8px;
    }

    .measure-mode-switch button {
      display: flex;
      align-items: center;
      justify-content: space-between;
      min-height: 40px;
      width: 100%;
      border-color: var(--line);
      background: #fff;
      font-weight: 700;
    }

    .measure-mode-switch button::after {
      content: "";
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: transparent;
      border: 1px solid var(--line-strong);
    }

    .measure-mode-switch button.active {
      background: #e8f3fb;
      border-color: #8ec7e7;
      color: #064d74;
    }

    .measure-mode-switch button.active::after {
      background: var(--accent);
      border-color: var(--accent);
    }

    .tempo-options {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }

    .tempo-options label {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 3px 6px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: #fff;
      color: var(--text);
      font-size: 12px;
      font-weight: 600;
    }

    .athlete-groups {
      display: grid;
      grid-template-columns: repeat(3, minmax(220px, 1fr));
      gap: 12px;
    }

    .group-panel {
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: #fbfcfe;
    }

    .group-panel.active {
      border-color: var(--accent);
      box-shadow: 0 0 0 2px rgba(0, 119, 182, 0.12);
    }

    .group-panel-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      padding: 10px;
      background: #eef2f6;
      border-bottom: 1px solid var(--line);
    }

    .group-panel-head h3 {
      margin: 0;
      font-size: 14px;
    }

    .group-panel-body {
      padding: 10px;
    }

    .dive-settings-grid {
      display: grid;
      grid-template-columns: 1fr;
      gap: 12px;
      align-items: start;
    }

    .dive-detail {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fbfcfe;
      padding: 10px;
      min-width: 0;
    }

    .dive-detail-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      margin-bottom: 10px;
    }

    .dive-detail-head h3 {
      margin: 0;
      font-size: 14px;
    }

    .data-tables {
      display: grid;
      grid-template-columns: 1fr;
      gap: 16px;
      max-width: 100%;
      min-width: 0;
      overflow: hidden;
    }

    .measure-block {
      min-width: 0;
      max-width: 100%;
      overflow: hidden;
    }

    .subsection-title {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 8px;
    }

    .subsection-title h3 {
      margin: 0;
      font-size: 15px;
    }

    .pill {
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      padding: 0 8px;
      border-radius: 999px;
      background: #e8f3fb;
      color: #064d74;
      font-size: 12px;
      font-weight: 700;
    }

    .table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
    }

    .measure-scroll {
      width: 100%;
      min-width: 0;
      max-width: 100%;
      max-height: 420px;
      overflow: auto;
      overscroll-behavior-x: contain;
      overscroll-behavior-y: contain;
    }

    .paste-grid-wrap {
      max-width: 100%;
      max-height: 460px;
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      overscroll-behavior: contain;
    }

    .paste-grid {
      min-width: max-content;
      table-layout: fixed;
    }

    .measure-input-grid {
      width: max-content;
      min-width: max-content;
      table-layout: fixed;
    }

    .paste-grid th,
    .paste-grid td {
      min-width: 128px;
      width: 128px;
      height: 34px;
      padding: 4px 8px;
    }

    .paste-grid th:first-child,
    .paste-grid td:first-child {
      min-width: 160px;
      width: 160px;
      max-width: 160px;
      text-align: left;
      background: #f8fafc;
      font-weight: 700;
    }

    .paste-grid .paste-cell {
      padding: 0;
      background: #fff;
      cursor: text;
    }

    .paste-grid .paste-cell > div,
    .paste-grid .paste-label-cell > div {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 100%;
      min-height: 34px;
      padding: 0 8px;
      outline: none;
      white-space: pre;
    }

    .paste-grid .paste-cell > div:focus,
    .paste-grid .paste-label-cell > div:focus {
      box-shadow: inset 0 0 0 2px var(--accent);
      background: #fff;
    }

    .paste-grid .paste-selected > div {
      box-shadow: inset 0 0 0 2px var(--accent);
      background: #fff;
    }

    .measure-input-grid .paste-cell {
      padding: 0;
    }

    .measure-input-grid .paste-cell > div {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 100%;
      min-height: 34px;
      padding: 0 8px;
      outline: none;
      white-space: pre;
    }

    .measure-input-grid .paste-cell > div:focus {
      box-shadow: inset 0 0 0 2px var(--accent);
      background: #fff;
    }

    .measure-input-grid th,
    .measure-input-grid td {
      min-width: 96px;
      width: 96px;
      max-width: 96px;
      height: 34px;
      padding: 4px 8px;
    }

    .measure-input-grid .measure-set-col {
      min-width: 64px;
      width: 64px;
      max-width: 64px;
      text-align: center;
    }

    .measure-input-grid .measure-rep-col {
      min-width: 54px;
      width: 54px;
      max-width: 54px;
      text-align: center;
      font-weight: 700;
    }

    .measure-input-grid .measure-point-col {
      min-width: 78px;
      width: 78px;
      max-width: 78px;
      text-align: center;
      color: var(--muted);
    }

    .measure-input-grid .measure-set-col,
    .measure-input-grid .measure-rep-col,
    .measure-input-grid .measure-point-col {
      position: sticky;
      background: #f8fafc;
      z-index: 2;
    }

    .measure-input-grid th.measure-set-col,
    .measure-input-grid th.measure-rep-col,
    .measure-input-grid th.measure-point-col {
      z-index: 4;
      background: #e8edf4;
    }

    .measure-input-grid .measure-set-col { left: 0; }
    .measure-input-grid .measure-rep-col { left: 64px; }
    .measure-input-grid .measure-point-col { left: 118px; }

    .measure-input-grid .athlete-name {
      min-width: 104px;
      width: 104px;
      max-width: 104px;
      text-align: center;
    }

    .measure-input-grid .measure-set-row th,
    .measure-input-grid .measure-set-row td {
      position: sticky;
      left: 0;
      z-index: 3;
      height: 28px;
      min-height: 28px;
      background: #e8f3fb;
      color: #064d74;
      font-weight: 700;
      text-align: left;
    }

    .measure-input-grid .measure-set-row td {
      border-right: 0;
    }

    .measure-input-grid .measure-summary-row td {
      background: #fff7ed;
    }

    .measure-input-grid .measure-summary-row .measure-point-col {
      color: #9a3412;
      font-weight: 700;
    }

    .measure-input-grid .rep-start {
      border-top: 2px solid #b8c6d4;
    }

    .dive-input-grid th.measure-point-col,
    .dive-input-grid td.measure-point-col {
      left: 0;
      min-width: 78px;
      width: 78px;
      max-width: 78px;
      text-align: center;
      color: var(--text);
      font-weight: 700;
    }

    .dive-input-grid th.measure-rep-col,
    .dive-input-grid td.measure-rep-col {
      left: 78px;
      min-width: 76px;
      width: 76px;
      max-width: 76px;
      text-align: center;
      color: var(--muted);
      font-weight: 700;
    }

    .dive-input-grid th.athlete-name {
      min-width: 112px;
      width: 112px;
      max-width: 112px;
    }

    .dive-input-grid .dive-distance-start td,
    .dive-input-grid .dive-distance-start th {
      border-top: 2px solid #b8c6d4;
    }

    .dive-input-grid .dive-empty-cell {
      min-width: 320px;
      width: 320px;
      max-width: none;
      color: var(--muted);
      text-align: left;
      background: #fbfcfe;
    }

    .paste-grid .ignored-row td,
    .paste-grid .ignored-row th,
    .paste-grid td.ignored-row {
      background: #f4f6f8;
      color: #8a96a5;
    }

    .paste-grid .set-gap-row td,
    .paste-grid .set-gap-row th {
      height: 10px;
      min-height: 10px;
      padding: 0;
      background: var(--bg);
      border-right: 0;
    }

    .paste-grid .summary-row th {
      color: #9a3412;
    }

    .paste-grid .label-blue th { color: #0070c0; }
    .paste-grid .label-red th { color: #ff0000; }
    .paste-grid .label-magenta th { color: #ff00ff; }
    .paste-grid .label-purple th { color: #7030a0; }

    .paste-grid .athlete-group-head {
      background: #eef2f6;
      text-align: center;
      font-weight: 700;
    }

    .paste-grid .d-subhead {
      background: #f6f8fb;
      color: #48576b;
      font-size: 12px;
      white-space: nowrap;
    }

    .paste-grid .d-merge-slot {
      background: #fbfcfe;
      color: #738196;
    }

    table {
      border-collapse: collapse;
      width: 100%;
      min-width: 900px;
      font-size: 13px;
    }

    th, td {
      border-bottom: 1px solid var(--line);
      border-right: 1px solid var(--line);
      padding: 6px;
      text-align: center;
      vertical-align: middle;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: #eef2f6;
      color: #354457;
      font-size: 12px;
      font-weight: 700;
    }

    .measure-table {
      min-width: max-content;
      table-layout: fixed;
    }

    .measure-table th,
    .measure-table td {
      min-width: 82px;
      width: 82px;
      height: 30px;
      padding: 3px 5px;
    }

    .measure-table .athlete-name {
      min-width: 130px;
      width: 130px;
    }

    .measure-table .measure-item {
      min-width: 150px;
      width: 150px;
    }

    .measure-table .rep-start {
      border-top: 2px solid #b8c6d4;
    }

    .measure-table .set-gap {
      height: 10px;
      min-height: 10px;
      padding: 0;
      border-top: 0;
      border-bottom: 0;
      background: var(--bg);
    }

    .measure-table th {
      white-space: nowrap;
      min-width: 82px;
      width: 82px;
    }

    .measure-table tbody tr:nth-child(even) td {
      background: #fcfdff;
    }

    .measure-table tbody td:not(:first-child) {
      position: relative;
      padding: 0;
    }

    td:first-child, th:first-child {
      position: sticky;
      left: 0;
      z-index: 2;
      background: #f8fafc;
      text-align: left;
      min-width: 160px;
      max-width: 220px;
    }

    th:first-child { z-index: 3; background: #e8edf4; }

    .measure-table td:first-child,
    .measure-table th:first-child {
      min-width: 150px;
      width: 150px;
      max-width: 150px;
    }

    td input {
      width: 76px;
      text-align: center;
      padding: 0 5px;
    }

    .measure-table td input {
      display: block;
      width: 100%;
      height: 100%;
      min-height: 30px;
      margin: 0;
      padding: 0 5px;
      border: 1px solid transparent;
      background: transparent;
      border-radius: 0;
      box-shadow: inset 0 0 0 1px transparent;
      -webkit-appearance: none;
      appearance: none;
    }

    .measure-table td input:focus {
      background: #fff;
      border-color: transparent;
      outline: none;
      box-shadow: inset 0 0 0 2px var(--accent);
    }

    .final-table {
      min-width: max-content;
      width: auto;
      border-collapse: collapse;
      table-layout: fixed;
      background: #fff;
      font-size: 12px;
    }

    .final-table th,
    .final-table td {
      position: static;
      min-width: 78px;
      width: 78px;
      height: 24px;
      padding: 2px 5px;
      border: 1px solid #111;
      background: #fff;
      color: #111;
      text-align: center;
      font-weight: 400;
      white-space: nowrap;
    }

    .final-table .rep-col {
      min-width: 70px;
      width: 70px;
      max-width: 70px;
    }

    .final-table .name-head {
      font-weight: 400;
      background: #fff;
      overflow: hidden;
      text-overflow: clip;
    }

    .final-table .tempo-head {
      font-weight: 400;
      background: #fff;
      overflow: hidden;
      text-overflow: clip;
    }

    .final-table .rep-col,
    .final-table .merged-summary {
      overflow: hidden;
      text-overflow: clip;
    }

    .final-table .yellow {
      background: #fff;
    }

    .final-table .thick-top > * { border-top-width: 3px; }
    .final-table .thick-bottom > * { border-bottom-width: 3px; }
    .final-table tbody tr:last-child > * { border-bottom-width: 3px; }
    .final-table .thick-left { border-left-width: 3px; }
    .final-table .thick-right { border-right-width: 3px; }
    .final-table .blue { color: #0070c0; }
    .final-table .red { color: #ff0000; }
    .final-table .magenta { color: #ff00ff; }
    .final-table .purple { color: #7030a0; }
    .final-table .gold { color: #f2a900; }
    .final-table .merged-summary {
      text-align: center;
      vertical-align: middle;
    }

    .final-table .d-merged-set {
      min-width: 234px;
      width: 234px;
      max-width: 234px;
    }

    .final-table .d-merged-50 {
      min-width: 234px;
      width: 234px;
      max-width: 234px;
    }
    .final-table .clear-cell {
      border: 0;
      background: transparent;
    }

    .final-grid {
      display: grid;
      grid-template-columns: 1fr;
      align-items: flex-start;
      gap: 24px;
      width: max-content;
    }

    .final-pair {
      display: flex;
      align-items: flex-start;
      gap: 14px;
    }

    .final-page-pair {
      display: grid;
      grid-template-columns: 1fr;
      gap: 14px;
      break-inside: avoid;
      page-break-inside: avoid;
    }

    .final-pair.group-D {
      flex-direction: column;
    }

    .final-block {
      width: auto;
    }

    @media print {
      @page {
        size: A4 landscape;
        margin: 4mm;
      }

      body {
        background: #fff;
        overflow: visible;
      }

      .app-header,
      aside,
      #memoView,
      #managerView,
      #mergeView .section-header,
      #mergeView .field-grid,
      #mergeView .small,
      #mergeStatus {
        display: none !important;
      }

      .layout,
      main,
      .section,
      .section-body,
      .table-wrap {
        display: block;
        padding: 0;
        margin: 0;
        border: 0;
        box-shadow: none;
        overflow: visible;
        background: #fff;
      }

      #mergeView {
        display: block !important;
      }

      #mergeTable,
      #mergeTable > tbody,
      #mergeTable > tbody > tr,
      #mergeTable > tbody > tr > td {
        display: block;
        width: 100%;
      }

      .final-grid {
        display: block;
        width: 100%;
      }

      .final-pair {
        display: flex;
        align-items: flex-start;
        gap: 4mm;
        width: 100%;
        overflow: visible;
        page-break-after: always;
        page-break-inside: avoid;
        break-after: page;
        break-inside: avoid;
        margin: 0;
      }

      .final-pair.group-D {
        flex-direction: column;
        gap: 3mm;
        width: 100%;
      }

      .final-block {
        page-break-after: auto;
        page-break-inside: avoid;
        break-after: auto;
        break-inside: avoid;
        margin: 0;
        transform-origin: top left;
      }

      .final-pair:last-child {
        page-break-after: auto;
        break-after: auto;
      }


      .final-table {
        font-size: 10.5px;
      }

      .final-table th,
      .final-table td {
        min-width: 36px;
        width: 36px;
        height: 10.8px;
        padding: 0 2px;
      }

      .final-pair:not(.group-D) .final-table {
        font-size: 11px;
      }

      .final-pair:not(.group-D) .final-table th,
      .final-pair:not(.group-D) .final-table td {
        min-width: 22mm;
        width: 22mm;
        height: 4.2mm;
        padding: 0 2px;
      }

      .final-table .rep-col {
        min-width: 30px;
        width: 30px;
        max-width: 30px;
        font-size: 0.68em;
      }

      .final-table .name-head,
      .final-table .tempo-head {
        font-size: 0.82em;
      }

      .final-pair:not(.group-D) .final-table .rep-col {
        min-width: 12mm;
        width: 12mm;
        max-width: 12mm;
        font-size: 0.9em;
      }

      .final-table .d-merged-set,
      .final-table .d-merged-50 {
        min-width: 174px;
        width: 174px;
        max-width: 174px;
      }
    }

    .compact-table {
      min-width: 0;
    }

    .compact-table th,
    .compact-table td {
      padding: 5px;
    }

    .compact-table td:first-child,
    .compact-table th:first-child {
      min-width: 42px;
      width: 42px;
      text-align: center;
    }

    .compact-table tr.selected-row td {
      background: #eef7ff;
    }

    td.name-cell input {
      width: 100%;
      text-align: left;
    }

    .small {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.5;
    }

    .summary-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(130px, 1fr));
      gap: 10px;
    }

    .metric {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 10px;
      background: #fbfcfe;
    }

    .metric strong {
      display: block;
      font-size: 20px;
      margin-top: 4px;
    }

    .status {
      min-height: 22px;
      color: var(--muted);
      font-size: 13px;
    }

    .status.ok { color: var(--ok); }
    .status.warn { color: var(--warn); }

    .hidden { display: none; }

    @media (max-width: 900px) {
      .layout { grid-template-columns: 1fr; }
      aside {
        border-right: 0;
        border-bottom: 1px solid var(--line);
      }
      .tabs { flex-direction: row; overflow-x: auto; }
      .tab { width: auto; }
      .field-grid, .summary-grid { grid-template-columns: 1fr 1fr; }
      .quick-guide, .athlete-groups { grid-template-columns: 1fr; }
      main { padding: 12px; }
    }

    @media (max-width: 560px) {
      .app-header {
        align-items: flex-start;
        flex-direction: column;
        padding: 12px;
      }
      .field-grid, .summary-grid { grid-template-columns: 1fr; }
      .section-header {
        align-items: flex-start;
        flex-direction: column;
      }
    }
  </style>
</head>
<body>
  <header class="app-header">
    <h1>水泳データ集計</h1>
    <div class="header-actions">
      <button id="backupExportBtn" title="現在の入力内容をJSONファイルに保存">バックアップ保存</button>
      <button id="backupImportBtn" title="保存したJSONファイルから復元">バックアップ読込</button>
      <input id="backupImportFile" class="hidden" type="file" accept="application/json,.json">
      <button id="clearBtn" class="danger" title="入力内容を初期化">初期化</button>
    </div>
  </header>

  <div class="layout">
    <aside>
      <nav class="tabs" aria-label="画面切り替え">
        <button class="tab active" data-view="settings">選手設定</button>
        <button class="tab" data-view="memo">測定入力</button>
        <button class="tab" data-view="merge">表</button>
        <button class="tab" data-view="output">出力</button>
      </nav>
      <div class="section" style="margin-top:16px;">
        <div class="section-header"><h2>種別</h2></div>
        <div class="section-body">
          <div class="measure-mode-switch" aria-label="測定種別">
            <button class="active" data-measure-mode="pushOff">プッシュオフ</button>
            <button data-measure-mode="dive">ダイブ</button>
          </div>
        </div>
      </div>
      <div class="section" style="margin-top:16px;">
        <div class="section-header"><h2>入力状況</h2></div>
        <div class="section-body summary-grid" style="grid-template-columns:1fr;">
          <div class="metric"><span class="small">選手</span><strong id="athleteCount">0</strong></div>
          <div class="metric"><span class="small">測定入力</span><strong id="memoCount">0</strong></div>
        </div>
      </div>
    </aside>

    <main>
      <section id="settingsView" class="view">
        <div class="quick-guide">
          <div class="guide-item"><strong>1. 班を選ぶ</strong><span class="small">A / BC / Dごとに選手順と入力表を切り替えます。</span></div>
          <div class="guide-item"><strong>2. 選手を並べる</strong><span class="small">班ごとの表に名前を入れ、上から測定順にします。</span></div>
          <div class="guide-item"><strong>3. 表を設定</strong><span class="small">選手ごとに完成表へ出す集計行を切り替えます。</span></div>
        </div>

        <div class="section">
          <div class="section-header">
            <h2>測定する班とメニュー</h2>
            <span id="menuBadge" class="pill">A 50x4x4</span>
          </div>
          <div class="section-body">
            <div class="field-grid" style="grid-template-columns:minmax(160px, 260px);">
              <label><span>班</span>
                <select id="groupSelect">
                  <option value="A">A</option>
                  <option value="BC">BC</option>
                  <option value="D">D</option>
                </select>
              </label>
            </div>
            <div class="group-switch" style="margin-top:12px;">
              <button data-select-group="A" class="active">A班</button>
              <button data-select-group="BC">BC班</button>
              <button data-select-group="D">D班</button>
            </div>
          </div>
        </div>

        <div id="pushOffSettingsPanel">
        <div class="section">
          <div class="section-header">
            <h2>班ごとの選手順</h2>
            <div class="toolbar">
              <button id="removeBlankNamesBtn">空行削除</button>
            </div>
          </div>
          <div class="section-body">
            <div class="athlete-groups">
              <div class="group-panel active" data-group-panel="A">
                <div class="group-panel-head">
                  <h3>A班</h3>
                  <div class="toolbar">
                    <button data-add-group="A">追加</button>
                    <button data-paste-group="A">貼付</button>
                  </div>
                </div>
                <div class="group-panel-body">
                  <div class="table-wrap">
                    <table class="compact-table">
                      <thead><tr><th>順</th><th>選手</th><th>操作</th></tr></thead>
                      <tbody id="athleteBodyA"></tbody>
                    </table>
                  </div>
                </div>
              </div>
              <div class="group-panel" data-group-panel="BC">
                <div class="group-panel-head">
                  <h3>BC班</h3>
                  <div class="toolbar">
                    <button data-add-group="BC">追加</button>
                    <button data-paste-group="BC">貼付</button>
                  </div>
                </div>
                <div class="group-panel-body">
                  <div class="table-wrap">
                    <table class="compact-table">
                      <thead><tr><th>順</th><th>選手</th><th>操作</th></tr></thead>
                      <tbody id="athleteBodyBC"></tbody>
                    </table>
                  </div>
                </div>
              </div>
              <div class="group-panel" data-group-panel="D">
                <div class="group-panel-head">
                  <h3>D班</h3>
                  <div class="toolbar">
                    <button data-add-group="D">追加</button>
                    <button data-paste-group="D">貼付</button>
                  </div>
                </div>
                <div class="group-panel-body">
                  <div class="table-wrap">
                    <table class="compact-table">
                      <thead><tr><th>順</th><th>選手</th><th>操作</th></tr></thead>
                      <tbody id="athleteBodyD"></tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div class="section">
          <div class="section-header">
            <h2>完成表設定</h2>
          </div>
          <div class="section-body">
            <div class="field-grid" style="grid-template-columns:minmax(180px, 280px); margin-bottom:12px;">
              <label><span>設定する選手</span>
                <select id="chartAthlete"></select>
              </label>
            </div>
            <div class="settings-panel summary-config-panel">
              <h3>完成表に出す集計行</h3>
              <div id="summarySettings"></div>
            </div>
          </div>
        </div>
        </div>

        <div id="diveSettingsPanel" class="hidden">
          <div class="section">
            <div class="section-header">
              <h2>ダイブ選手順</h2>
              <div class="toolbar">
                <button id="removeBlankDiveNamesBtn">空行削除</button>
                <button id="syncDiveSettingsBtn" class="primary">プッシュオフに同期</button>
              </div>
            </div>
            <div class="section-body">
              <div class="dive-settings-grid">
                <div class="athlete-groups">
                  <div class="group-panel active" data-dive-group-panel="A">
                    <div class="group-panel-head">
                      <h3>A班</h3>
                      <div class="toolbar">
                        <button data-add-dive-athlete="A">追加</button>
                        <button data-paste-dive-athletes="A">貼付</button>
                      </div>
                    </div>
                    <div class="group-panel-body">
                      <div class="table-wrap">
                        <table class="compact-table">
                          <thead><tr><th>順</th><th>選手</th><th>操作</th></tr></thead>
                          <tbody id="diveAthleteBodyA"></tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                  <div class="group-panel" data-dive-group-panel="BC">
                    <div class="group-panel-head">
                      <h3>BC班</h3>
                      <div class="toolbar">
                        <button data-add-dive-athlete="BC">追加</button>
                        <button data-paste-dive-athletes="BC">貼付</button>
                      </div>
                    </div>
                    <div class="group-panel-body">
                      <div class="table-wrap">
                        <table class="compact-table">
                          <thead><tr><th>順</th><th>選手</th><th>操作</th></tr></thead>
                          <tbody id="diveAthleteBodyBC"></tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                  <div class="group-panel" data-dive-group-panel="D">
                    <div class="group-panel-head">
                      <h3>D班</h3>
                      <div class="toolbar">
                        <button data-add-dive-athlete="D">追加</button>
                        <button data-paste-dive-athletes="D">貼付</button>
                      </div>
                    </div>
                    <div class="group-panel-body">
                      <div class="table-wrap">
                        <table class="compact-table">
                          <thead><tr><th>順</th><th>選手</th><th>操作</th></tr></thead>
                          <tbody id="diveAthleteBodyD"></tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <p class="small">プッシュオフに同期すると、選択中の班の選手名をダイブ側へコピーします。</p>
            </div>
          </div>

          <div class="section">
            <div class="section-header">
              <h2>ダイブ表設定</h2>
            </div>
            <div class="section-body">
              <div class="field-grid" style="grid-template-columns:minmax(180px, 280px); margin-bottom:12px;">
                <label><span>設定する選手</span>
                  <select id="diveConfigAthlete"></select>
                </label>
              </div>
              <div id="diveConfigPanel" class="settings-panel"></div>
            </div>
          </div>
        </div>
      </section>

      <section id="memoView" class="view hidden">
        <div class="section">
          <div class="section-header">
            <h2>測定データ</h2>
            <span id="measureModeBadge" class="pill">プッシュオフ</span>
          </div>
          <div class="section-body">
            <div class="field-grid" style="grid-template-columns:minmax(160px, 260px); margin-bottom:12px;">
              <label><span>チーム</span>
                <select id="memoGroup">
                  <option value="A">A</option>
                  <option value="BC">BC</option>
                  <option value="D">D</option>
                </select>
              </label>
            </div>
            <div id="memoStatus" class="status"></div>
            <div id="pushOffPanel" class="data-tables">
              <div class="measure-block">
                <div class="subsection-title">
                  <h3>プッシュオフ タイム</h3>
                  <div class="toolbar">
                    <span id="timeCount" class="pill">0項目</span>
                    <button id="copyTimeBtn">コピー</button>
                  </div>
                </div>
                <div class="table-wrap measure-scroll">
                  <table id="timeTable"></table>
                </div>
              </div>
              <div class="measure-block">
                <div class="subsection-title">
                  <h3>プッシュオフ テンポ</h3>
                  <div class="toolbar">
                    <span id="tempoCount" class="pill">0項目</span>
                    <button id="copyTempoBtn">コピー</button>
                  </div>
                </div>
                <div class="table-wrap measure-scroll">
                  <table id="tempoTable"></table>
                </div>
              </div>
            </div>
            <div id="divePanel" class="data-tables hidden">
              <div class="measure-block">
                <div class="subsection-title">
                  <h3>ダイブ タイム</h3>
                  <div class="toolbar">
                    <span id="diveTimeCount" class="pill">0距離</span>
                    <button id="copyDiveBtn">コピー</button>
                  </div>
                </div>
                <div id="diveStatus" class="status"></div>
                <div class="table-wrap measure-scroll">
                  <table id="diveTimeTable"></table>
                </div>
              </div>
              <div class="measure-block">
                <div class="subsection-title">
                  <h3>ダイブ テンポ</h3>
                  <div class="toolbar">
                    <span id="diveTempoCount" class="pill">0項目</span>
                    <button id="copyDiveTempoBtn">コピー</button>
                  </div>
                </div>
                <div class="table-wrap measure-scroll">
                  <table id="diveTempoTable"></table>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <section id="mergeView" class="view hidden">
        <div class="section">
          <div class="section-header">
            <h2>表</h2>
            <div class="toolbar">
              <button id="copyMergeBtn">TSVコピー</button>
            </div>
          </div>
          <div class="section-body">
            <div class="field-grid" style="grid-template-columns:minmax(160px, 260px);">
              <label><span>チーム</span>
                <select id="mergeGroup">
                  <option value="all">全班</option>
                  <option value="A">A</option>
                  <option value="BC">BC</option>
                  <option value="D">D</option>
                </select>
              </label>
            </div>
            <p class="small">選択中の種別に応じて表を表示します。</p>
            <div id="mergeStatus" class="status"></div>
            <div class="table-wrap">
              <table id="mergeTable"></table>
            </div>
          </div>
        </div>
      </section>

      <section id="outputView" class="view hidden">
        <div class="section">
          <div class="section-header">
            <h2>出力</h2>
            <div class="toolbar">
              <button id="pdfMergeBtn" class="primary">PDF出力</button>
            </div>
          </div>
          <div class="section-body">
            <div class="field-grid" style="grid-template-columns:minmax(180px, 260px) minmax(180px, 260px);">
              <label><span>チーム</span>
                <select id="outputGroup">
                  <option value="all">全班</option>
                  <option value="A">A</option>
                  <option value="BC">BC</option>
                  <option value="D">D</option>
                </select>
              </label>
              <label><span>出力内容</span>
                <select id="outputMode">
                  <option value="sync">同期</option>
                  <option value="pushOff">プッシュオフだけ</option>
                  <option value="dive">ダイブだけ</option>
                </select>
              </label>
            </div>
            <p class="small">同期では、同じ班・同じ名前のプッシュオフとダイブを同じページに表示します。</p>
            <div id="outputStatus" class="status"></div>
          </div>
        </div>
      </section>
    </main>
  </div>

  <script>
    const menuDefs = {
      A: { label: "A 50x4x4", blocks: 4, reps: 4, distance: 50, fields: ["50m time", "15m tempo", "35m tempo"] },
      BC: {
        label: "BC 50x6x3 + 50x4",
        segments: [
          { label: "50x6x3", blocks: 3, reps: 6, distance: 50, fields: ["50m time", "15m tempo", "35m tempo"] },
          { label: "50x4", blocks: 1, reps: 4, distance: 50, fields: ["50m time", "15m tempo", "35m tempo"] }
        ]
      },
      D: {
        label: "D 100x6x3 + 50x4",
        segments: [
          { label: "100x6x3", blocks: 3, reps: 6, distance: 100, fields: ["50m time", "100m time", "15m tempo", "35m tempo", "65m tempo", "85m tempo"] },
          { label: "50x4", blocks: 1, reps: 4, distance: 50, fields: ["50m time", "15m tempo", "35m tempo"] }
        ]
      }
    };

    const state = {
      athletes: [
        { name: "", group: "A" },
        { name: "", group: "A" },
        { name: "", group: "A" },
        { name: "", group: "A" }
      ],
      session: { group: "A", date: "", stroke: "", note: "", measureMode: "pushOff" },
      memo: {},
      dive: { distances: { A: 50, BC: 50, D: 50 }, athletes: { A: [], BC: [], D: [] }, settings: { A: [], BC: [], D: [] } },
      managerRows: [],
      managerHeaders: [],
      managerPaste: {},
      summaryConfig: {}
    };

    const $ = (id) => document.getElementById(id);
    const STORAGE_KEY = "swimDataApp_finalInput";
    let selectedPasteCell = null;
    let liveUpdateTimer = null;
    let athleteNameUpdateTimer = null;
    const defaultSummaryConfig = {
      setTempo: false,
      setAverage: true,
      setHr: true,
      setBest: false,
      setStdDev: false,
      set4Best: true,
      set4StdDev: true,
      allRange: "1-3",
      allUntil: 3,
      allHr: false,
      allAve: true,
      allStdDev: true,
      allBest: true,
      dSetFrontBack: true,
      dSetBest: false,
      dSetStdDev: false,
      dSet4Total: false,
      dAllFrontBack: false,
      dAllHr: false
    };

    function today() {
      const d = new Date();
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    }

    function fieldKind(field) {
      return field.includes("tempo") ? "tempo" : "time";
    }

    function fieldLabel(field) {
      return field
        .replace(" time", "")
        .replace(" tempo", "")
        .replace("m", "m");
    }

    function flattenMenu(group, kind = "all") {
      const def = menuDefs[group];
      const segments = def.segments || [{ label: def.label, blocks: def.blocks, reps: def.reps, distance: def.distance, fields: def.fields }];
      const columns = [];
      segments.forEach((segment) => {
        for (let block = 1; block <= segment.blocks; block += 1) {
          for (let rep = 1; rep <= segment.reps; rep += 1) {
            segment.fields.forEach((field) => {
              const type = fieldKind(field);
              if (kind !== "all" && type !== kind) return;
              columns.push({
                key: `${segment.distance}-${segment.label}-B${block}-R${rep}-${field}`,
                label: `${segment.label} B${block}-${rep} ${fieldLabel(field)}`,
                type,
                segment: segment.label,
                distance: segment.distance,
                block,
                rep,
                position: fieldLabel(field),
                displayLabel: type === "time" ? String(rep) : fieldLabel(field)
              });
            });
          }
        }
      });
      return columns;
    }

    function athleteKey(index, name) {
      return name.trim() || `row_${index + 1}`;
    }

    function ensureMemoRow(index, name) {
      const key = athleteKey(index, name);
      if (!state.memo[key]) state.memo[key] = {};
      return state.memo[key];
    }

    function ensureStateShape() {
      if (!state.session) state.session = { group: "A", date: "", stroke: "", note: "", measureMode: "pushOff" };
      if (!state.session.measureMode) state.session.measureMode = "pushOff";
      if (!state.session.diveConfigGroup) state.session.diveConfigGroup = state.session.group || "A";
      if (!Number.isInteger(Number(state.session.diveConfigIndex))) state.session.diveConfigIndex = 0;
      if (!state.memo) state.memo = {};
      if (!state.dive) state.dive = {};
      if (!state.dive.distances) state.dive.distances = {};
      if (!state.dive.athletes) state.dive.athletes = {};
      if (!state.dive.settings) state.dive.settings = {};
      ["A", "BC", "D"].forEach((group) => {
        if (!Number.isFinite(Number(state.dive.distances[group]))) state.dive.distances[group] = 50;
        if (!Array.isArray(state.dive.athletes[group])) state.dive.athletes[group] = [];
        if (!Array.isArray(state.dive.settings[group])) state.dive.settings[group] = [];
        state.dive.settings[group] = state.dive.athletes[group].map((name, index) => {
          const current = state.dive.settings[group][index] || {};
          const distance = Math.max(1, Math.floor(Number(current.distance ?? state.dive.distances[group]) || 50));
          return { distance, tempo: current.tempo && typeof current.tempo === "object" ? current.tempo : {} };
        });
      });
      if (!state.managerPaste) state.managerPaste = {};
      if (!state.summaryConfig) state.summaryConfig = {};
    }

    function diveAthleteKey(group, index, name) {
      return `dive::${group}::${name.trim() || `row_${index + 1}`}`;
    }

    function ensureDiveMemoRow(group, index, name) {
      const key = diveAthleteKey(group, index, name);
      if (!state.memo[key]) state.memo[key] = {};
      return state.memo[key];
    }

    function diveSetting(group, index) {
      ensureStateShape();
      return state.dive.settings[group][index] || { distance: 50, tempo: {} };
    }

    function diveDistanceActive(group, athleteIndex, distance) {
      return diveDistances(diveSetting(group, athleteIndex).distance).includes(distance);
    }

    function diveTempoActive(group, athleteIndex, distance) {
      const setting = diveSetting(group, athleteIndex);
      return diveDistanceActive(group, athleteIndex, distance) && setting.tempo?.[distance] === true;
    }

    function selectedDiveConfigAthlete() {
      ensureStateShape();
      const group = state.session.diveConfigGroup || state.session.group || "A";
      const index = Number(state.session.diveConfigIndex) || 0;
      const name = state.dive.athletes[group]?.[index];
      if (name === undefined) return null;
      return { group, index, name };
    }

    function setDiveConfigSelection(group, index) {
      state.session.diveConfigGroup = group || "A";
      state.session.diveConfigIndex = Math.max(0, Number(index) || 0);
    }

    function summaryScopeKey(group, name = "") {
      return name ? `${group}::${name}` : group;
    }

    function selectedSettingsGroup() {
      return state.session.group || $("groupSelect")?.value || "A";
    }

    function selectedChartAthlete() {
      const raw = $("chartAthlete")?.value;
      if (raw === undefined || raw === "") return null;
      const index = Number(raw);
      const athlete = state.athletes[index];
      if (!Number.isInteger(index) || !athlete || athlete.group !== selectedSettingsGroup()) return null;
      return { athlete, index };
    }

    function selectedSummaryName(group = selectedSettingsGroup()) {
      const select = $("summaryAthlete");
      if (select?.value) return select.value;
      const selected = selectedChartAthlete();
      if (selected?.athlete?.group === group) return athleteDisplayName(selected.index);
      return state.athletes.find((athlete) => athlete.group === group)?.name.trim() || "";
    }

    function summaryConfig(group = $("pasteGroup")?.value || state.session.group || "A", name = "") {
      if (!state.summaryConfig) state.summaryConfig = {};
      const key = summaryScopeKey(group, name);
      const groupConfig = state.summaryConfig[group] || {};
      if (!state.summaryConfig[key]) state.summaryConfig[key] = {};
      state.summaryConfig[key] = { ...state.summaryConfig[key] };
      return { ...defaultSummaryConfig, ...groupConfig, ...state.summaryConfig[key] };
    }

    function setSummaryConfigValue(group, name, key, value) {
      if (!state.summaryConfig) state.summaryConfig = {};
      const scope = summaryScopeKey(group, name);
      state.summaryConfig[scope] = { ...state.summaryConfig[scope], [key]: value };
    }

    function aggregateSummaryConfig(group) {
      const athletes = state.athletes.filter((athlete) => athlete.group === group);
      const configs = athletes.length
        ? athletes.map((athlete, index) => summaryConfig(group, athlete.name.trim() || `${group}班 ${index + 1}人目`))
        : [summaryConfig(group)];
      const aggregate = { ...defaultSummaryConfig };
      Object.keys(aggregate).forEach((key) => {
        if (["allRange", "allFrom", "allUntil"].includes(key)) return;
        aggregate[key] = configs.some((config) => config[key]);
      });
      const ranges = configs.map((config) => summaryAllRange(config));
      const widest = ranges.reduce((best, item) => {
        if (!item.until) return best;
        if (!best.until) return item;
        return (item.until - item.from) > (best.until - best.from) ? item : best;
      }, { from: 0, until: 0 });
      aggregate.allRange = widest.until ? `${widest.from}-${widest.until}` : "0";
      aggregate.allUntil = widest.until;
      [1, 2, 3, 4].forEach((setNo) => {
        ["Tempo", "Average", "Hr", "Best", "StdDev", "FrontBack", "Total"].forEach((suffix) => {
          const key = `set${setNo}${suffix}`;
          aggregate[key] = configs.some((config) => setOption(config, setNo, suffix));
        });
      });
      return aggregate;
    }

    function summaryAllUntil(config) {
      return summaryAllRange(config).until;
    }

    function summaryAllRange(config) {
      const rawRange = config?.allRange;
      if (rawRange === 0 || rawRange === "0" || rawRange === "" || rawRange === "none") return { from: 0, until: 0 };
      if (typeof rawRange === "string" && rawRange.includes("-")) {
        const [fromRaw, untilRaw] = rawRange.split("-");
        const from = Math.min(4, Math.max(1, Number(fromRaw)));
        const until = Math.min(4, Math.max(1, Number(untilRaw)));
        if (Number.isFinite(from) && Number.isFinite(until) && from < until) return { from, until };
      }
      const rawFrom = config?.allFrom;
      const raw = config?.allUntil;
      if (raw === 0 || raw === "0" || raw === "" || raw === "none") return { from: 0, until: 0 };
      const fromValue = Number(rawFrom ?? 1);
      const value = Number(raw ?? defaultSummaryConfig.allUntil);
      if (!Number.isFinite(value)) return summaryAllRange({ allRange: defaultSummaryConfig.allRange });
      const from = Number.isFinite(fromValue) ? Math.min(4, Math.max(1, fromValue)) : 1;
      const until = Math.min(4, Math.max(0, value));
      if (until <= 0) return { from: 0, until: 0 };
      if (from >= until) return from === 1 && until > 1 ? { from: 1, until } : { from: Math.max(1, until - 1), until };
      return { from, until };
    }

    function renderSummarySettings() {
      const container = $("summarySettings");
      const select = $("summaryAthlete");
      if (!container) return;
      const selected = selectedChartAthlete();
      const group = selected?.athlete?.group || selectedSettingsGroup();
      const athletes = state.athletes.filter((athlete) => athlete.group === group);
      if (select) {
        const current = select.value;
        select.innerHTML = athletes.map((athlete, index) => {
          const name = athlete.name.trim() || `${group}班 ${index + 1}人目`;
          return `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`;
        }).join("");
        if (current && [...select.options].some((option) => option.value === current)) select.value = current;
      }
      const name = selected ? athleteDisplayName(selected.index) : selectedSummaryName(group);
      const config = summaryConfig(group, name);
      const setBlocks = [1, 2, 3, 4].map((setNo) => {
        const dMainSet = group === "D" && setNo <= 3;
        const items = [
          ["Tempo", "Tempo"],
          ["Average", "Average"],
          ["Hr", "HR"],
          ["Best", "BestAve/達成率"],
          ["StdDev", "標準偏差"]
        ];
        if (dMainSet) items.splice(2, 0, ["FrontBack", "前半/後半Ave"]);
        if (group === "D" && setNo === 4) items.splice(2, 0, ["Total", "Total / BestTotal/達成率"]);
        return `<div class="summary-setting-group"><h4>${setNo}set</h4><div class="check-grid">${items.map(([suffix, label]) => {
          const key = `set${setNo}${suffix}`;
          return `<label><input type="checkbox" data-summary-setting="${key}" ${setOption(config, setNo, suffix) ? "checked" : ""}>${escapeHtml(label)}</label>`;
        }).join("")}</div></div>`;
      }).join("");
      const allItems = [
        ["allHr", "HR"],
        ["allAve", "AllAve"],
        ["allBest", "BestAve/達成率"],
        ["allStdDev", "標準偏差"]
      ];
      if (group === "D") allItems.splice(1, 0, ["dAllFrontBack", "前半/後半Ave"]);
      const allRange = summaryAllRange(config);
      const allRangeValue = allRange.until ? `${allRange.from}-${allRange.until}` : "0";
      const allRangeOptions = [
        ["0", "なし"],
        ["1-2", "1〜2set"],
        ["1-3", "1〜3set"],
        ["1-4", "1〜4set"],
        ["2-3", "2〜3set"],
        ["2-4", "2〜4set"],
        ["3-4", "3〜4set"]
      ];
      const allBlock = `<div class="summary-setting-group"><h4>All系</h4><div class="check-grid"><label>対象 <select data-summary-setting="allRange">${allRangeOptions.map(([value, label]) => `<option value="${value}" ${allRangeValue === value ? "selected" : ""}>${label}</option>`).join("")}</select></label>${allItems.map(([key, label]) => `<label><input type="checkbox" data-summary-setting="${key}" ${config[key] || (key === "allHr" && config.dAllHr) ? "checked" : ""}>${escapeHtml(label)}</label>`).join("")}</div></div>`;
      container.innerHTML = setBlocks + allBlock;
    }

    function setOption(config, setNo, suffix) {
      const direct = config[`set${setNo}${suffix}`];
      if (direct !== undefined) return !!direct;
      if (suffix === "Tempo") return config.setTempo !== false;
      if (suffix === "Average") return !!config.setAverage;
      if (suffix === "Hr") return !!config.setHr;
      if (suffix === "Best") return setNo === 4 ? !!config.set4Best : !!(config.setBest || config.dSetBest);
      if (suffix === "StdDev") return setNo === 4 ? !!config.set4StdDev : !!(config.setStdDev || config.dSetStdDev);
      if (suffix === "FrontBack") return !!config.dSetFrontBack;
      if (suffix === "Total") return !!config.dSet4Total;
      return false;
    }

    function groupIndexes(group) {
      return state.athletes
        .map((athlete, index) => ({ athlete, index }))
        .filter((item) => item.athlete.group === group);
    }

    function renderAthletes() {
      ["A", "BC", "D"].forEach((group) => {
        const tbody = $(`athleteBody${group}`);
        tbody.innerHTML = "";
        groupIndexes(group).forEach(({ athlete, index }, position) => {
          const tr = document.createElement("tr");
          tr.innerHTML = `
            <td>${position + 1}</td>
            <td class="name-cell"><input data-athlete-index="${index}" data-field="name" value="${escapeHtml(athlete.name)}" placeholder="${group}班 ${position + 1}人目"></td>
            <td>
              <button class="icon" title="上へ" data-move-up="${index}">↑</button>
              <button class="icon" title="下へ" data-move-down="${index}">↓</button>
              <button class="icon danger" title="削除" data-delete="${index}">×</button>
            </td>
          `;
          tbody.appendChild(tr);
        });
      });
      document.querySelectorAll("[data-group-panel]").forEach((panel) => {
        panel.classList.toggle("active", panel.dataset.groupPanel === state.session.group);
      });
      document.querySelectorAll("[data-select-group]").forEach((button) => {
        button.classList.toggle("active", button.dataset.selectGroup === state.session.group);
      });
      updateMetrics();
    }

    function setActiveGroup(group) {
      state.session.group = group;
      if ($("groupSelect")) $("groupSelect").value = group;
      if ($("memoGroup")) $("memoGroup").value = group;
      renderAll();
    }

    function renderMemo() {
      ensureStateShape();
      syncSession();
      const group = state.session.group;
      const timeColumns = flattenMenu(group, "time");
      const tempoColumns = flattenMenu(group, "tempo");
      const filteredAthletes = state.athletes.filter((a) => a.group === group);
      renderMemoTable("timeTable", filteredAthletes, timeColumns);
      renderMemoTable("tempoTable", filteredAthletes, tempoColumns);
      $("menuBadge").textContent = menuDefs[group].label;
      $("timeCount").textContent = `${timeColumns.length}項目`;
      $("tempoCount").textContent = `${tempoColumns.length}項目`;
      renderDive();
      renderMeasureMode();
      $("memoStatus").textContent = state.session.measureMode === "dive"
        ? `${group}班 / ダイブ / ${state.dive.athletes[group].length}名`
        : `${group}班 / ${menuDefs[group].label} / ${filteredAthletes.length}名`;
      updateMetrics();
    }

    function renderMeasureMode() {
      const mode = state.session.measureMode || "pushOff";
      $("pushOffPanel")?.classList.toggle("hidden", mode !== "pushOff");
      $("divePanel")?.classList.toggle("hidden", mode !== "dive");
      $("pushOffSettingsPanel")?.classList.toggle("hidden", mode !== "pushOff");
      $("diveSettingsPanel")?.classList.toggle("hidden", mode !== "dive");
      if ($("measureModeBadge")) $("measureModeBadge").textContent = mode === "dive" ? "ダイブ" : "プッシュオフ";
      document.querySelectorAll("[data-measure-mode]").forEach((button) => {
        button.classList.toggle("active", button.dataset.measureMode === mode);
      });
    }

    function diveDistances(maxDistance) {
      const max = Math.floor(Number(maxDistance));
      if (!Number.isFinite(max) || max < 1) return [];
      const values = new Set();
      for (let distance = 15; distance <= max; distance += 10) {
        if (distance % 50 !== 5) values.add(distance);
      }
      for (let distance = 50; distance <= max; distance += 50) values.add(distance);
      return [...values].sort((a, b) => a - b);
    }

    function renderDiveSettings() {
      ensureStateShape();
      const selection = selectedDiveConfigAthlete();
      if (!selection) {
        const firstGroup = ["A", "BC", "D"].find((group) => state.dive.athletes[group].length);
        if (firstGroup) setDiveConfigSelection(firstGroup, 0);
      }
      ["A", "BC", "D"].forEach((group) => {
        const tbody = $(`diveAthleteBody${group}`);
        if (!tbody) return;
        const athletes = state.dive.athletes[group];
        tbody.innerHTML = athletes.length
          ? athletes.map((name, index) => {
            const active = state.session.diveConfigGroup === group && Number(state.session.diveConfigIndex) === index;
            return `<tr class="${active ? "selected-row" : ""}"><td>${index + 1}</td><td class="name-cell"><input data-dive-name-group="${escapeAttr(group)}" data-dive-name-index="${index}" value="${escapeAttr(name)}" placeholder="${group}班 ${index + 1}人目"></td><td><button class="icon" title="設定" data-dive-select="${index}" data-dive-action-group="${escapeAttr(group)}">設定</button><button class="icon" title="上へ" data-dive-move-up="${index}" data-dive-action-group="${escapeAttr(group)}">↑</button><button class="icon" title="下へ" data-dive-move-down="${index}" data-dive-action-group="${escapeAttr(group)}">↓</button><button class="icon danger" title="削除" data-dive-delete="${index}" data-dive-action-group="${escapeAttr(group)}">×</button></td></tr>`;
          }).join("")
          : `<tr><td colspan="3">選手未設定</td></tr>`;
      });
      document.querySelectorAll("[data-dive-group-panel]").forEach((panel) => {
        panel.classList.toggle("active", panel.dataset.diveGroupPanel === state.session.group);
      });
      renderDiveConfigPanel();
    }

    function renderDiveConfigPanel() {
      const select = $("diveConfigAthlete");
      const panel = $("diveConfigPanel");
      if (!select || !panel) return;
      const options = ["A", "BC", "D"].flatMap((group) => state.dive.athletes[group].map((name, index) => {
        const label = `${group}班 ${name.trim() || `${index + 1}人目`}`;
        return { value: `${group}:${index}`, label };
      }));
      select.innerHTML = options.length
        ? options.map((option) => `<option value="${escapeAttr(option.value)}">${escapeHtml(option.label)}</option>`).join("")
        : `<option value="">選手未設定</option>`;
      const selected = selectedDiveConfigAthlete();
      if (!selected) {
        panel.innerHTML = `<p class="small">ダイブ選手を追加すると、距離とTempoを設定できます。</p>`;
        return;
      }
      const selectedValue = `${selected.group}:${selected.index}`;
      if ([...select.options].some((option) => option.value === selectedValue)) select.value = selectedValue;
      const setting = diveSetting(selected.group, selected.index);
      const distances = diveDistances(setting.distance);
      panel.innerHTML = `<h3>${escapeHtml(selected.group)}班 ${escapeHtml(selected.name.trim() || `${selected.index + 1}人目`)}</h3><div class="field-grid" style="grid-template-columns:minmax(130px, 170px); margin-bottom:12px;"><label><span>距離 x(m)</span><input data-dive-distance-group="${escapeAttr(selected.group)}" data-dive-distance-index="${selected.index}" type="number" min="1" step="1" value="${escapeAttr(setting.distance)}"></label></div><div class="summary-setting-group"><h4>Tempo</h4><div class="tempo-options">${distances.map((distance) => `<label><input type="checkbox" data-dive-tempo-group="${escapeAttr(selected.group)}" data-dive-tempo-index="${selected.index}" data-dive-tempo-distance="${distance}" ${setting.tempo?.[distance] === true ? "checked" : ""}>${distance}m</label>`).join("") || `<span class="small">距離を入力してください</span>`}</div></div>`;
    }

    function renderDive() {
      ensureStateShape();
      const group = state.session.group;
      const athletes = state.dive.athletes[group];
      const distances = [...new Set(state.dive.settings[group].flatMap((setting) => diveDistances(setting.distance)))].sort((a, b) => a - b);
      const tempoDistances = distances.filter((distance) => athletes.some((_, index) => diveTempoActive(group, index, distance)));
      $("diveTimeCount").textContent = `${distances.length}距離`;
      $("diveTempoCount").textContent = `${tempoDistances.length}項目`;
      $("diveStatus").textContent = athletes.length
        ? `${group}班 / ${athletes.length}名`
        : `${group}班 / プッシュオフに同期してください`;
      $("diveStatus").className = athletes.length ? "status" : "status warn";
      renderDiveInputTable("diveTimeTable", athletes, distances, "time", (index, distance) => diveDistanceActive(group, index, distance));
      renderDiveInputTable("diveTempoTable", athletes, tempoDistances, "tempo", (index, distance) => diveTempoActive(group, index, distance));
    }

    function renderDiveInputTable(tableId, athletes, distances, kind, activeFor) {
      const group = state.session.group;
      const table = $(tableId);
      table.className = "paste-grid measure-input-grid dive-input-grid";
      table.innerHTML = `<thead><tr><th class="measure-point-col">距離</th>${athletes.map((name, index) => `<th class="athlete-name">${escapeHtml(name || `${group}班 ${index + 1}人目`)}</th>`).join("")}</tr></thead><tbody></tbody>`;
      const tbody = table.querySelector("tbody");
      if (!athletes.length) {
        tbody.innerHTML = `<tr><td class="measure-point-col"></td><td class="dive-empty-cell">選手設定で「プッシュオフに同期」を押すと、ダイブ入力表が作られます。</td></tr>`;
        return;
      }
      if (!distances.length) {
        tbody.innerHTML = `<tr><td class="measure-point-col"></td><td class="dive-empty-cell">${kind === "tempo" ? "選手設定でTempoを有効にしてください。" : "選手設定で距離 x(m) を設定してください。"}</td></tr>`;
        return;
      }
      distances.forEach((distance, rowIndex) => {
        const tr = document.createElement("tr");
        tr.className = "dive-distance-start";
        tr.innerHTML = `<td class="measure-point-col">${distance}</td>` + athletes.map((name, colIndex) => {
          const active = activeFor(colIndex, distance);
          const row = ensureDiveMemoRow(group, colIndex, name);
          const key = `dive-${distance}m-${kind}`;
          const value = row[key] || "";
          return `<td class="paste-cell ${active ? "" : "ignored-row"}" data-memo-grid="1" data-active="${active ? "1" : "0"}" data-paste-row="${rowIndex}" data-paste-col="${colIndex}" data-memo-athlete="${escapeAttr(diveAthleteKey(group, colIndex, name))}" data-memo-key="${escapeAttr(key)}"><div contenteditable="${active ? "true" : "false"}">${escapeHtml(active ? value : "-")}</div></td>`;
        }).join("");
        tbody.appendChild(tr);
      });
    }

    function renderMemoTable(tableId, filteredAthletes, columns) {
      const table = $(tableId);
      table.className = "paste-grid measure-input-grid";
      const displayColumns = columnsWithSetGaps(columns);
      table.innerHTML = `<thead><tr><th class="measure-set-col">set</th><th class="measure-rep-col">本</th><th class="measure-point-col">項目</th>${filteredAthletes.map((athlete) => {
        const originalIndex = state.athletes.indexOf(athlete);
        const name = athlete.name.trim() || `${originalIndex + 1}人目`;
        return `<th class="athlete-name">${escapeHtml(name)}</th>`;
      }).join("")}</tr></thead><tbody></tbody>`;
      const tbody = table.querySelector("tbody");
      let pasteRow = 0;
      displayColumns.forEach((col) => {
        const tr = document.createElement("tr");
        if (col.gap) {
          tr.className = "measure-set-row";
          tr.innerHTML = `<th colspan="${filteredAthletes.length + 3}">${escapeHtml(measureSetTitle(col.next))}</th>`;
          tbody.appendChild(tr);
          return;
        }
        tr.innerHTML = `<td class="measure-set-col ${measureCellClass(col, columns)}">${escapeHtml(String(logicalSetNumber(state.session.group, col)))}</td><td class="measure-rep-col ${measureCellClass(col, columns)}">${escapeHtml(measureRowLabel(col))}</td><td class="measure-point-col ${measureCellClass(col, columns)}">${escapeHtml(measurePointLabel(col))}</td>` + filteredAthletes.map((athlete) => {
          const originalIndex = state.athletes.indexOf(athlete);
          const row = ensureMemoRow(originalIndex, athlete.name);
          const active = memoCellActive(tableId, athlete, originalIndex, col);
          const value = row[col.key] || "";
          return `<td class="paste-cell ${active ? "" : "ignored-row"}" data-memo-grid="1" data-active="${active ? "1" : "0"}" data-paste-row="${pasteRow}" data-paste-col="${filteredAthletes.indexOf(athlete)}" data-memo-athlete="${escapeAttr(athleteKey(originalIndex, athlete.name))}" data-memo-key="${escapeAttr(col.key)}"><div contenteditable="${active ? "true" : "false"}">${escapeHtml(active ? value : "-")}</div></td>`;
        }).join("");
        tbody.appendChild(tr);
        pasteRow += 1;
        if (tableId === "timeTable" && isLastColumnInSet(col, columns)) {
          const setNo = logicalSetNumber(state.session.group, col);
          [
            { label: "HR", key: summaryInputKey(setNo, "HR"), option: "Hr" },
            { label: "BestAve", key: summaryInputKey(setNo, "BestAve"), option: "Best" }
          ].forEach((summary) => {
            if (!filteredAthletes.some((athlete) => {
              const originalIndex = state.athletes.indexOf(athlete);
              return summaryInputActive(athlete, originalIndex, setNo, summary.option);
            })) return;
            const rowEl = document.createElement("tr");
            rowEl.className = "measure-summary-row";
            rowEl.innerHTML = `<td class="measure-set-col">${escapeHtml(String(setNo))}</td><td class="measure-rep-col"></td><td class="measure-point-col">${escapeHtml(summary.label)}</td>` + filteredAthletes.map((athlete) => {
              const originalIndex = state.athletes.indexOf(athlete);
              const row = ensureMemoRow(originalIndex, athlete.name);
              const active = summaryInputActive(athlete, originalIndex, setNo, summary.option);
              const value = row[summary.key] || "";
              return `<td class="paste-cell ${active ? "" : "ignored-row"}" data-memo-grid="1" data-active="${active ? "1" : "0"}" data-paste-row="${pasteRow}" data-paste-col="${filteredAthletes.indexOf(athlete)}" data-memo-athlete="${escapeAttr(athleteKey(originalIndex, athlete.name))}" data-memo-key="${escapeAttr(summary.key)}"><div contenteditable="${active ? "true" : "false"}">${escapeHtml(active ? value : "-")}</div></td>`;
            }).join("");
            tbody.appendChild(rowEl);
            pasteRow += 1;
          });
          [
            { label: "All HR", key: summaryInputKey("All", "HR"), option: "allHr" },
            { label: "All BestAve", key: summaryInputKey("All", "BestAve"), option: "allBest" }
          ].forEach((summary) => {
            if (!filteredAthletes.some((athlete) => {
              const originalIndex = state.athletes.indexOf(athlete);
              return allSummaryInputActive(athlete, originalIndex, setNo, summary.option);
            })) return;
            const rowEl = document.createElement("tr");
            rowEl.className = "measure-summary-row";
            rowEl.innerHTML = `<td class="measure-set-col">All</td><td class="measure-rep-col"></td><td class="measure-point-col">${escapeHtml(summary.label)}</td>` + filteredAthletes.map((athlete) => {
              const originalIndex = state.athletes.indexOf(athlete);
              const row = ensureMemoRow(originalIndex, athlete.name);
              const active = allSummaryInputActive(athlete, originalIndex, setNo, summary.option);
              const value = row[summary.key] || "";
              return `<td class="paste-cell ${active ? "" : "ignored-row"}" data-memo-grid="1" data-active="${active ? "1" : "0"}" data-paste-row="${pasteRow}" data-paste-col="${filteredAthletes.indexOf(athlete)}" data-memo-athlete="${escapeAttr(athleteKey(originalIndex, athlete.name))}" data-memo-key="${escapeAttr(summary.key)}"><div contenteditable="${active ? "true" : "false"}">${escapeHtml(active ? value : "-")}</div></td>`;
            }).join("");
            tbody.appendChild(rowEl);
            pasteRow += 1;
          });
        }
      });
    }

    function isLastColumnInSet(col, columns) {
      const index = columns.indexOf(col);
      const next = columns[index + 1];
      return !next || `${next.segment}-${next.block}` !== `${col.segment}-${col.block}`;
    }

    function summaryInputKey(setNo, kind) {
      return `summary-S${setNo}-${kind}`;
    }

    function summaryInputActive(athlete, athleteIndex, setNo, option) {
      const config = summaryConfig(athlete.group, athleteDisplayName(athleteIndex));
      return setOption(config, setNo, option);
    }

    function allSummaryInputActive(athlete, athleteIndex, setNo, option) {
      const config = summaryConfig(athlete.group, athleteDisplayName(athleteIndex));
      const range = summaryAllRange(config);
      if (range.until !== setNo) return false;
      if (option === "allHr") return !!(config.allHr || config.dAllHr);
      return !!config[option];
    }

    function memoCellActive(tableId, athlete, athleteIndex, col) {
      if (tableId !== "tempoTable") return true;
      const setNo = logicalSetNumber(athlete.group, col);
      const config = summaryConfig(athlete.group, athleteDisplayName(athleteIndex));
      return setOption(config, setNo, "Tempo");
    }

    function measureRowLabel(col) {
      return String(col.rep);
    }

    function measurePointLabel(col) {
      return col.position;
    }

    function measureSetTitle(col) {
      if (!col) return "";
      return `${logicalSetNumber(state.session.group, col)} set / ${col.segment}`;
    }

    function columnsWithSetGaps(columns) {
      const result = [];
      columns.forEach((col, index) => {
        const prev = columns[index - 1];
        if (!prev || `${prev.segment}-${prev.block}` !== `${col.segment}-${col.block}`) {
          result.push({ gap: true, next: col });
        }
        result.push(col);
      });
      return result;
    }

    function measureCellClass(col, columns) {
      const index = columns.indexOf(col);
      const prev = columns[index - 1];
      const blockKey = `${col.segment}-${col.block}`;
      const repKey = `${col.segment}-${col.block}-${col.rep}`;
      const classes = [];
      if (!prev || `${prev.segment}-${prev.block}` !== blockKey) {
        classes.push("rep-start");
      } else if (`${prev.segment}-${prev.block}-${prev.rep}` !== repKey) {
        classes.push("rep-start");
      }
      return classes.join(" ");
    }

    function parseTable(text) {
      return text
        .split(/\r?\n/)
        .map((line) => line.trimEnd())
        .filter((line) => line.length > 0)
        .map((line) => line.split("\t").length > 1 ? line.split("\t") : line.split(","));
    }

    function parseManager() {
      refreshManagerFromPaste(true);
    }

    function refreshManagerFromPaste(showStatus = false) {
      if (!$("managerPasteGrid") || !$("pasteGroup")) return;
      const group = $("pasteGroup").value;
      captureManagerPasteGrid(group);
      const parsed = parseManagerGrid(group);
      replaceManagerGroup(group, parsed.rows);
      if (showStatus) {
        $("managerStatus").textContent = `${parsed.rows.length}名分を貼り付けから取り込みました`;
        $("managerStatus").className = "status ok";
      }
      renderManager();
      renderMerge();
      updateMetrics();
      persistSilently();
    }

    function captureManagerPasteGrid(group = $("pasteGroup").value) {
      if (!$("managerPasteGrid")) return;
      if (!state.managerPaste) state.managerPaste = {};
      const values = {};
      $("managerPasteGrid").querySelectorAll(".paste-cell").forEach((cell) => {
        if (cell.dataset.active !== "1") return;
        const key = cell.dataset.stableKey || `${cell.dataset.pasteRow}:${cell.dataset.pasteCol}`;
        const value = cell.textContent.trim();
        if (value) values[key] = value;
      });
      state.managerPaste[group] = values;
    }

    function replaceManagerGroup(group, rows) {
      state.managerRows = [
        ...state.managerRows.filter((row) => row.group !== group),
        ...rows
      ];
      state.managerHeaders = unique(state.managerRows.flatMap((row) => row.values.map((value) => value.label)));
    }

    function managerPasteRows(group, options = {}) {
      const rows = [];
      const legacyAllOrder = !!options.legacyAllOrder;
      const config = aggregateSummaryConfig(group);
      const athletes = state.athletes.filter((athlete) => athlete.group === group);
      const configs = athletes.length
        ? athletes.map((athlete, index) => summaryConfig(group, athlete.name.trim() || `${group}班 ${index + 1}人目`))
        : [summaryConfig(group)];
      const allTargets = new Set(configs.map((item) => summaryAllUntil(item)).filter((setNo) => setNo > 0));
      const hasAllTarget = (setNo) => allTargets.has(setNo);
      const addStats = (setNo) => {
        if (setOption(config, setNo, "Hr")) rows.push({ label: "HR", type: "hr", set: setNo, import: true });
        if (setOption(config, setNo, "Average")) rows.push({ label: "Average", type: "calc", set: setNo, option: "Average", import: false });
        if (setOption(config, setNo, "Best")) {
          rows.push({ label: "BestAve", type: "setBest", set: setNo, import: true });
          rows.push({ label: "達成率", type: "calc", set: setNo, option: "Best", import: false });
        }
        if (setOption(config, setNo, "StdDev")) rows.push({ label: "標準偏差", type: "calc", set: setNo, option: "StdDev", import: false });
      };
      const addAllStats = (setNo) => {
        if (config.allHr) rows.push({ label: "HR", type: "allHr", set: setNo, import: true });
        if (config.allAve) rows.push({ label: "AllAve", type: "calcAll", set: setNo, option: "allAve", import: false });
        if (legacyAllOrder) {
          if (config.allStdDev) rows.push({ label: "標準偏差", type: "calcAll", set: setNo, option: "allStdDev", import: false });
          if (config.allBest) {
            rows.push({ label: "BestAve", type: "best", set: setNo, import: true });
            rows.push({ label: "達成率", type: "calcAll", set: setNo, option: "allBest", import: false });
          }
        } else {
          if (config.allBest) {
            rows.push({ label: "BestAve", type: "best", set: setNo, import: true });
            rows.push({ label: "達成率", type: "calcAll", set: setNo, option: "allBest", import: false });
          }
          if (config.allStdDev) rows.push({ label: "標準偏差", type: "calcAll", set: setNo, option: "allStdDev", import: false });
        }
      };

      if (group === "A") {
        [1, 2, 3, 4].forEach((setNo) => {
          for (let rep = 1; rep <= 4; rep += 1) rows.push({ label: String(rep), type: "time", set: setNo, rep, import: true });
          addStats(setNo);
          if (hasAllTarget(setNo)) addAllStats(setNo);
        });
      } else if (group === "BC") {
        [1, 2, 3, 4].forEach((setNo) => {
          const reps = setNo === 4 ? 4 : 6;
          for (let rep = 1; rep <= reps; rep += 1) rows.push({ label: String(rep), type: "time", set: setNo, rep, import: true });
          addStats(setNo);
          if (hasAllTarget(setNo)) addAllStats(setNo);
        });
      } else {
        const addDAllStats = (setNo) => {
          if (config.dAllFrontBack) {
            rows.push({ label: "前半Ave", type: "calcAll", set: setNo, option: "dAllFrontBack", import: false });
            rows.push({ label: "後半Ave", type: "calcAll", set: setNo, option: "dAllFrontBack", import: false });
          }
          if (config.allAve) rows.push({ label: "AllAve", type: "calcAll", set: setNo, option: "allAve", import: false });
          if (config.allHr || config.dAllHr) rows.push({ label: "全set HR", type: "allHr", set: setNo, import: true });
          if (config.allBest) {
            rows.push({ label: "BestAve", type: "best", set: setNo, import: true });
            rows.push({ label: "達成率", type: "calcAll", set: setNo, option: "allBest", import: false });
          }
          if (config.allStdDev) rows.push({ label: "標準偏差", type: "calcAll", set: setNo, option: "allStdDev", import: false });
        };
        [1, 2, 3].forEach((setNo) => {
          for (let rep = 1; rep <= 6; rep += 1) {
            rows.push({ label: `${rep}:50m`, type: "time100", set: setNo, rep, position: "50m", import: true });
            rows.push({ label: `${rep}:100m`, type: "time100", set: setNo, rep, position: "100m", import: true });
          }
          if (setOption(config, setNo, "Hr")) rows.push({ label: `${setNo}set HR`, type: "hr", set: setNo, import: true });
          if (setOption(config, setNo, "FrontBack")) {
            rows.push({ label: `${setNo}set 前半Ave`, type: "calc", set: setNo, option: "FrontBack", import: false });
            rows.push({ label: `${setNo}set 後半Ave`, type: "calc", set: setNo, option: "FrontBack", import: false });
          }
          if (setOption(config, setNo, "Average")) rows.push({ label: `${setNo}set Average`, type: "calc", set: setNo, option: "Average", import: false });
          if (setOption(config, setNo, "Best")) {
            rows.push({ label: `${setNo}set BestAve`, type: "setBest", set: setNo, import: true });
            rows.push({ label: `${setNo}set 達成率`, type: "calc", set: setNo, option: "Best", import: false });
          }
          if (setOption(config, setNo, "StdDev")) rows.push({ label: `${setNo}set 標準偏差`, type: "calc", set: setNo, option: "StdDev", import: false });
          if (hasAllTarget(setNo)) addDAllStats(setNo);
        });
        for (let rep = 1; rep <= 4; rep += 1) {
          rows.push({ label: `${rep}:50m`, type: "time", set: 4, rep, import: true });
        }
        if (setOption(config, 4, "Hr")) rows.push({ label: "4set HR", type: "hr", set: 4, import: true });
        if (setOption(config, 4, "Total")) {
          rows.push({ label: "4set Total", type: "calc", set: 4, option: "Total", import: false });
          rows.push({ label: "4set BestTotal", type: "bestTotal", set: 4, import: true });
          rows.push({ label: "4set 達成率", type: "calc", set: 4, option: "Total", import: false });
        }
        if (setOption(config, 4, "Average")) rows.push({ label: "4set Average", type: "calc", set: 4, option: "Average", import: false });
        if (setOption(config, 4, "Best")) {
          rows.push({ label: "4set BestAve", type: "setBest", set: 4, import: true });
          rows.push({ label: "4set 達成率", type: "calc", set: 4, option: "Best", import: false });
        }
        if (setOption(config, 4, "StdDev")) rows.push({ label: "4set 標準偏差", type: "calc", set: 4, option: "StdDev", import: false });
        if (hasAllTarget(4)) addDAllStats(4);
      }
      return rows;
    }

    function managerPasteStableKey(row, colIndex, name) {
      return [
        name,
        row.type || "",
        row.set || "",
        row.rep || "",
        row.position || "",
        row.option || "",
        row.label || "",
        colIndex
      ].join("::");
    }

    function legacyManagerPasteValuesByStableKey(group, athletes) {
      const saved = state.managerPaste?.[group] || {};
      const legacyRows = managerPasteRows(group, { legacyAllOrder: true });
      const values = {};
      legacyRows.forEach((row, rowIndex) => {
        athletes.forEach((athlete, colIndex) => {
          const name = athlete.name.trim() || `${group}班 ${colIndex + 1}人目`;
          const legacyValue = saved[`${rowIndex}:${colIndex}`];
          if (!legacyValue) return;
          values[managerPasteStableKey(row, colIndex, name)] = legacyValue;
        });
      });
      return values;
    }

    function renderManagerPasteGrid() {
      if (!$("managerPasteGrid")) return;
      const group = $("pasteGroup")?.value || "A";
      renderSummarySettings();
      const athletes = state.athletes.filter((athlete) => athlete.group === group);
      const rows = managerPasteRows(group);
      const savedPaste = state.managerPaste?.[group] || {};
      const legacyPaste = legacyManagerPasteValuesByStableKey(group, athletes);
      const table = $("managerPasteGrid");
      table.innerHTML = `<thead><tr><th class="paste-corner paste-label-cell">項目</th>${athletes.map((athlete, index) => {
        const name = athlete.name.trim() || `${group}班 ${index + 1}人目`;
        return `<th>${escapeHtml(name)}</th>`;
      }).join("")}</tr></thead><tbody></tbody>`;
      const tbody = table.querySelector("tbody");
      rows.forEach((row, rowIndex) => {
        const tr = document.createElement("tr");
        if (row.type === "gap") tr.className = "set-gap-row";
        if (row.type === "ignored") tr.className = "ignored-row";
        if (row.type === "hr" || row.type === "best") tr.classList.add("summary-row");
        const colorClass = pasteRowColorClass(row);
        if (colorClass) tr.classList.add(colorClass);
        tr.innerHTML = `<th class="paste-label-cell">${escapeHtml(row.label)}</th>` + athletes.map((athlete, colIndex) => {
          const name = athlete.name.trim() || `${group}班 ${colIndex + 1}人目`;
          const active = managerRowAppliesToAthlete(group, name, row);
          const importsValue = active && row.import;
          const stableKey = managerPasteStableKey(row, colIndex, name);
          const savedValue = savedPaste[stableKey] || legacyPaste[stableKey] || savedPaste[`${rowIndex}:${colIndex}`] || "";
          const value = active ? savedValue || (importsValue ? managerPasteFallbackValue(group, name, row) : "") : "";
          return `<td class="paste-cell ${active ? "" : "ignored-row"}" data-paste-row="${rowIndex}" data-paste-col="${colIndex}" data-stable-key="${escapeAttr(stableKey)}" data-name="${escapeAttr(name)}" data-type="${escapeAttr(row.type)}" data-set="${row.set || ""}" data-rep="${row.rep || ""}" data-position="${row.position || ""}" data-option="${row.option || ""}" data-active="${active ? "1" : "0"}" data-import="${importsValue ? "1" : "0"}"><div contenteditable="${row.type !== "gap" && active ? "true" : "false"}">${escapeHtml(value)}</div></td>`;
        }).join("");
        tbody.appendChild(tr);
      });
    }

    function pasteRowColorClass(row) {
      if (!row.import) return "";
      if (row.type === "hr" || row.type === "allHr") return "label-blue";
      if (row.type === "best" || row.type === "setBest" || row.type === "bestTotal") return "label-magenta";
      return "";
    }

    function managerRowAppliesToAthlete(group, name, row) {
      if (row.type === "gap") return false;
      if (["time", "time100"].includes(row.type)) return true;
      const config = summaryConfig(group, name);
      if (row.type === "calc") return setOption(config, Number(row.set || 0), row.option || "");
      if (row.type === "calcAll") {
        const until = summaryAllUntil(config);
        if (row.set && Number(row.set) !== until) return false;
        return !!config[row.option];
      }
      if (row.type === "hr") return setOption(config, Number(row.set || 0), "Hr");
      if (row.type === "allHr") {
        const until = summaryAllUntil(config);
        if (row.set && Number(row.set) !== until) return false;
        return !!(config.allHr || config.dAllHr);
      }
      if (row.type === "best") {
        const until = summaryAllUntil(config);
        if (row.set && Number(row.set) !== until) return false;
        return !!config.allBest;
      }
      if (row.type === "setBest") return setOption(config, Number(row.set || 0), "Best");
      if (row.type === "bestTotal") return group === "D" && setOption(config, 4, "Total");
      return true;
    }

    function managerPasteFallbackValue(group, name, row) {
      if (!row.import) return "";
      const label = managerImportLabel({
        type: row.type,
        set: row.set || "",
        rep: row.rep || "",
        position: row.position || ""
      });
      if (!label) return "";
      const manager = state.managerRows.find((item) => item.group === group && item.name === name);
      const value = manager?.values?.find((item) => item.label === label || item.key === label);
      return value?.value || "";
    }

    function parseManagerGrid(group) {
      const rows = managerPasteRows(group);
      const athletes = state.athletes.filter((athlete) => athlete.group === group);
      const byName = new Map(athletes.map((athlete, index) => {
        const name = athlete.name.trim() || `${group}班 ${index + 1}人目`;
        return [name, { name, group, values: new Map() }];
      }));
      $("managerPasteGrid").querySelectorAll(".paste-cell[data-active='1']").forEach((cell) => {
        const value = cleanManagerValue(cell.textContent);
        if (!value) return;
        const name = cell.dataset.name;
        const athlete = byName.get(name);
        if (!athlete) return;
        const label = managerImportLabel(cell.dataset);
        if (label) athlete.values.set(label, value);
      });
      const headers = unique([...byName.values()].flatMap((athlete) => [...athlete.values.keys()]));
      return {
        headers,
        rows: [...byName.values()].map((athlete) => ({
          name: athlete.name,
          group: athlete.group,
          values: headers.map((label) => ({ key: label, label, value: athlete.values.get(label) || "" }))
        }))
      };
    }

    function managerImportLabel(dataset) {
      if (dataset.type === "hr") return `S${dataset.set} HR`;
      if (dataset.type === "allHr") return "All HR";
      if (dataset.type === "setBest") return `S${dataset.set} BestAve`;
      if (dataset.type === "bestTotal") return "S4 BestTotal";
      if (dataset.type === "best") return "BestAve";
      if (dataset.type === "calc" && dataset.option === "Best") return `S${dataset.set} 達成率`;
      if (dataset.type === "calc" && dataset.option === "Total") return "S4 Total 達成率";
      if (dataset.type === "calcAll" && dataset.option === "allBest") return "達成率";
      if (dataset.type === "time100") return `S${dataset.set} ${dataset.rep} ${dataset.position}`;
      if (dataset.type === "time") return `S${dataset.set} ${dataset.rep}`;
      return "";
    }

    function cleanManagerValue(value) {
      const text = String(value ?? "").trim().replace(/^\((.*)\)$/, "$1");
      if (!text || text === "#VALUE!" || text === "#DIV/0!") return "";
      return text;
    }

    function isPasteWritableCell(cell) {
      return !!cell && cell.dataset.active === "1";
    }

    function clearNonWritablePasteCells() {
      $("managerPasteGrid").querySelectorAll(".paste-cell").forEach((cell) => {
        if (!isPasteWritableCell(cell)) cell.querySelector("div").textContent = "";
      });
    }

    function handleManagerGridPaste(event) {
      const targetCell = event.target.closest(".paste-cell") || selectedPasteCell;
      if (!targetCell) return;
      event.preventDefault();
      const text = event.clipboardData.getData("text/plain");
      pasteTextIntoManagerGrid(text, targetCell);
    }

    function pasteTextIntoManagerGrid(text, targetCell = selectedPasteCell) {
      if (!targetCell) return;
      let rows = normalizePastedGrid(text);
      const startRow = Number(targetCell.dataset.pasteRow);
      const startCol = Number(targetCell.dataset.pasteCol);
      rows = trimPastedHeaders(rows);
      const hasLabelColumn = rows.some((row) => isPastedRowLabel(row[0]) && row.length > 1);
      if (hasLabelColumn) {
        rows = rows
          .filter((row) => isPastedRowLabel(row[0]))
          .map((row) => row.slice(1));
      } else if (rows.some((row) => shouldSkipPastedLabel(row[0]))) {
        rows = rows.map((row) => row.slice(1));
      }
      rows = trimEmptyPasteEdges(rows);
      pasteIntoWritableCells(rows, startRow, startCol);
      clearNonWritablePasteCells();
      scheduleLiveUpdate({ manager: true });
    }

    function pasteIntoWritableCells(rows, startRow, startCol) {
      if (!rows.length) return;
      const width = Math.max(...rows.map((row) => row.length));
      for (let colOffset = 0; colOffset < width; colOffset += 1) {
        const col = startCol + colOffset;
        if (col < 0) continue;
        const targets = writablePasteCellsInColumn(col, startRow);
        let targetIndex = 0;
        rows.forEach((row) => {
          if (colOffset >= row.length) return;
          const target = targets[targetIndex];
          if (!target) return;
          target.querySelector("div").textContent = row[colOffset] || "";
          targetIndex += 1;
        });
      }
    }

    function writablePasteCellsInColumn(col, startRow) {
      return [...$("managerPasteGrid").querySelectorAll(`.paste-cell[data-paste-col="${col}"]`)]
        .filter((cell) => Number(cell.dataset.pasteRow) >= startRow)
        .filter(isPasteWritableCell)
        .sort((a, b) => Number(a.dataset.pasteRow) - Number(b.dataset.pasteRow));
    }

    function normalizePastedGrid(text) {
      const rows = text
        .split(/\r?\n/)
        .filter((line) => line.length > 0)
        .map((line) => line.split("\t").map((cell) => String(cell ?? "").trim()));
      return trimEmptyPasteEdges(rows);
    }

    function trimEmptyPasteEdges(rows) {
      let result = rows.map((row) => [...row]);
      while (result.length && result[0].every((cell) => !String(cell || "").trim())) result.shift();
      while (result.length && result[result.length - 1].every((cell) => !String(cell || "").trim())) result.pop();
      if (!result.length) return result;
      let lastCol = Math.max(...result.map((row) => row.length)) - 1;
      while (lastCol >= 0 && result.every((row) => !String(row[lastCol] || "").trim())) lastCol -= 1;
      result = result.map((row) => row.slice(0, lastCol + 1));
      return result;
    }

    function trimPastedHeaders(rows) {
      let result = [...rows];
      while (result.length > 1 && isPastedHeaderRow(result[0])) result.shift();
      return result;
    }

    function isPastedHeaderRow(row) {
      if (!row || !row.length) return true;
      if (looksLikePastedHeader(row)) return true;
      const first = String(row[0] || "").trim();
      if (first === "(A)" || first === "(BC)" || first === "(D)") return true;
      if (shouldSkipPastedLabel(first)) return false;
      const meaningful = row.filter((cell) => String(cell || "").trim());
      if (!meaningful.length) return true;
      const dataLike = meaningful.filter((cell) => isPastedDataCell(cell)).length;
      const nameLike = meaningful.filter((cell) => isNameCellForPaste(cell)).length;
      return nameLike > dataLike;
    }

    function isPastedDataCell(value) {
      const text = String(value || "").trim().replace(/^\((.*)\)$/, "$1");
      if (!text || text === "-" || text.startsWith("#")) return false;
      return /^(\d+:\d+(?:\.\d+)?|\d+(?:\.\d+)?%?)$/.test(text);
    }

    function selectPasteCell(cell) {
      if (!isPasteWritableCell(cell)) return;
      if (selectedPasteCell) selectedPasteCell.classList.remove("paste-selected");
      selectedPasteCell = cell;
      selectedPasteCell.classList.add("paste-selected");
      const editable = selectedPasteCell.querySelector("div[contenteditable='true']");
      if (editable) {
        editable.focus();
        const range = document.createRange();
        range.selectNodeContents(editable);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }

    function handleManagerGridPointer(event) {
      const cell = event.target.closest(".paste-cell");
      if (cell) selectPasteCell(cell);
    }

    function handleDocumentPaste(event) {
      if (!$("managerPasteGrid")) return;
      const activeInGrid = document.activeElement?.closest?.("#managerPasteGrid");
      if (!activeInGrid && selectedPasteCell && !document.querySelector("#managerView.hidden")) {
        handleManagerGridPaste(event);
      }
    }

    async function handleControlPaste(event) {
      if (!$("managerPasteGrid")) return;
      if (!event.ctrlKey || event.metaKey || event.altKey || event.shiftKey || event.key.toLowerCase() !== "v") return;
      if (!selectedPasteCell || document.querySelector("#managerView.hidden")) return;
      event.preventDefault();
      try {
        const text = await navigator.clipboard.readText();
        pasteTextIntoManagerGrid(text, selectedPasteCell);
      } catch (error) {
        $("managerStatus").textContent = "Control+Vで貼り付けできませんでした。Safariのクリップボード許可を確認してください。";
        $("managerStatus").className = "status warn";
      }
    }

    function looksLikePastedHeader(row) {
      if (!row || row.length < 2) return false;
      return row.slice(1).some((cell) => isNameCellForPaste(cell));
    }

    function shouldSkipPastedLabel(value) {
      const label = String(value || "").trim();
      return label === "" || /^[1-9][0-9]?$/.test(label) || /^[1-9][0-9]?:\s*(50m|100m)$/.test(label) || /^\d+set\s+\d+本目\s+(50m|100m)$/.test(label) || /^\d+set\s+(HR|Average|BestAve|BestTotal|Total|達成率|標準偏差|前半Ave|後半Ave)$/.test(label) || ["HR", "Average", "BestAve", "BestTotal", "Total", "前半Ave", "後半Ave", "AllAve", "達成率", "標準偏差"].includes(label);
    }

    function isPastedRowLabel(value) {
      const label = String(value || "").trim();
      return shouldSkipPastedLabel(label) || ["BestAverage", "BestTotal", "Total", "全set HR", "前半Average", "後半Average", "BestAverage/達成率"].includes(label);
    }

    function isImportablePastedLabel(value) {
      const label = String(value || "").trim();
      if (/^[1-9][0-9]?$/.test(label)) return true;
      if (/^[1-9][0-9]?:\s*(50m|100m)$/.test(label)) return true;
      if (/^\d+set\s+\d+本目\s+(50m|100m)$/.test(label)) return true;
      if (/^\d+set\s+HR$/.test(label)) return true;
      return ["HR", "全set HR", "BestAve", "BestAverage", "BestTotal"].includes(label);
    }

    function isNameCellForPaste(value) {
      const text = String(value || "").trim();
      if (!text || ["Tempo", "HR", "Average", "BestAve", "AllAve", "達成率", "標準偏差"].includes(text)) return false;
      if (/^-?$/.test(text) || /^[0-9.:%/]+$/.test(text)) return false;
      return true;
    }

    function renderManager() {
      const table = $("managerTable");
      if (!table) return;
      const headers = ["選手", ...state.managerHeaders];
      table.innerHTML = `<thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody></tbody>`;
      const tbody = table.querySelector("tbody");
      state.managerRows.forEach((row) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${escapeHtml(row.name)}</td>${row.values.map((v) => `<td>${escapeHtml(v.value)}</td>`).join("")}`;
        tbody.appendChild(tr);
      });
    }

    function memoRowsForExport(groupFilter = "all") {
      const rows = [];
      state.athletes.forEach((athlete, index) => {
        if (groupFilter !== "all" && athlete.group !== groupFilter) return;
        const group = athlete.group;
        const cols = flattenMenu(group);
        const row = ensureMemoRow(index, athlete.name);
        rows.push({
          name: athlete.name.trim() || `${index + 1}人目`,
          group,
          values: cols.map((col) => ({ label: col.label, value: row[col.key] || "" }))
        });
      });
      return rows;
    }

    function renderMerge() {
      ensureStateShape();
      const group = $("mergeGroup").value;
      if ((state.session.measureMode || "pushOff") === "dive") {
        renderDiveMergeTable($("mergeTable"), group);
        updateMetrics();
        return;
      }
      const mode = $("mergeMode")?.value || "wide";
      const rows = memoRowsForExport(group);
      const managerMap = buildSummaryManagerMap();
      const table = $("mergeTable");
      if (mode === "wide") {
        renderFinalFormatTable(table, group, managerMap);
      } else if (mode === "long") {
        const headers = ["班", "選手", "種類", "項目", "値"];
        table.innerHTML = `<thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody></tbody>`;
        const tbody = table.querySelector("tbody");
        rows.forEach((row) => {
          row.values.forEach((item) => tbody.appendChild(rowEl([row.group, row.name, "測定", item.label, item.value])));
          const manager = managerMap.get(row.name);
          if (manager) manager.values.forEach((item) => tbody.appendChild(rowEl([row.group, row.name, "表", item.label, item.value])));
        });
      } else {
        const allMemoLabels = unique(rows.flatMap((row) => row.values.map((item) => item.label)));
        const headers = ["班", "選手", ...allMemoLabels, ...state.managerHeaders];
        table.innerHTML = `<thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead><tbody></tbody>`;
        const tbody = table.querySelector("tbody");
        rows.forEach((row) => {
          const memoMap = new Map(row.values.map((item) => [item.label, item.value]));
          const manager = managerMap.get(row.name);
          const managerMapValues = new Map(manager ? manager.values.map((item) => [item.label, item.value]) : []);
          tbody.appendChild(rowEl([
            row.group,
            row.name,
            ...allMemoLabels.map((label) => memoMap.get(label) || ""),
            ...state.managerHeaders.map((label) => managerMapValues.get(label) || "")
          ]));
        });
      }
      $("mergeStatus").textContent = `${rows.length}名を表示中`;
      updateMetrics();
    }

    function renderDiveMergeTable(table, groupFilter) {
      table.className = "final-table";
      const groups = groupFilter === "all" ? ["A", "BC", "D"] : [groupFilter];
      const blocks = [];
      groups.forEach((group) => {
        state.dive.athletes[group].forEach((name, index) => {
          blocks.push(diveFinalBlockHtml(group, index, name));
        });
      });
      if (!blocks.length) {
        table.innerHTML = `<tbody><tr><td>ダイブ選手が未同期です</td></tr></tbody>`;
      } else {
        table.innerHTML = `<tbody><tr><td style="border:0;padding:0;background:transparent;"><div class="final-grid">${blocks.join("")}</div></td></tr></tbody>`;
      }
      $("mergeStatus").textContent = `${blocks.length}名を表示中`;
    }

    function diveFinalBlockHtml(group, index, name) {
      const displayName = name || `${group}班 ${index + 1}人目`;
      const distances = diveDistances(diveSetting(group, index).distance);
      const row = ensureDiveMemoRow(group, index, displayName);
      const show50Diff = distances.some((distance) => distance >= 100);
      const headers = ["距離", "Time", "25m差", ...(show50Diff ? ["50m差"] : []), "Tempo"];
      const rows = distances.map((distance) => {
        const time = row[`dive-${distance}m-time`] || "";
        const diff25 = diveDiffValue(row, distance, 25);
        const diff50 = show50Diff ? diveDiffValue(row, distance, 50) : null;
        const tempo = diveTempoActive(group, index, distance) ? row[`dive-${distance}m-tempo`] || "" : "-";
        return `<tr><td class="rep-col thick-left">${distance}</td><td>${escapeHtml(time)}</td><td>${escapeHtml(diff25)}</td>${show50Diff ? `<td>${escapeHtml(diff50)}</td>` : ""}<td class="thick-right">${escapeHtml(tempo)}</td></tr>`;
      }).join("");
      const body = `<thead><tr class="thick-top">${headers.map((header, headerIndex) => `<th class="${headerIndex === 0 ? "rep-col thick-left" : headerIndex === headers.length - 1 ? "tempo-head thick-right" : "name-head"}">${escapeHtml(header)}</th>`).join("")}</tr><tr><th class="rep-col thick-left"></th><th colspan="${headers.length - 1}" class="name-head thick-right">${escapeHtml(displayName)}</th></tr></thead><tbody>${rows}</tbody>`;
      return `<div class="final-pair group-${group}"><table class="final-table final-block">${body}</table></div>`;
    }

    function diveDiffValue(row, distance, span) {
      const current = parseTimeSeconds(row[`dive-${distance}m-time`]);
      const previous = parseTimeSeconds(row[`dive-${distance - span}m-time`]);
      if (current === null || previous === null || current < previous) return "";
      return formatTimeSeconds(current - previous, 2);
    }

    function renderFinalFormatTable(table, groupFilter, managerMap) {
      table.className = "final-table";
      const groups = groupFilter === "all" ? ["A", "BC", "D"] : [groupFilter];
      const blocks = [];
      groups.forEach((group) => {
        state.athletes.forEach((athlete, index) => {
          if (athlete.group !== group) return;
          const name = athlete.name.trim() || `${group}班 ${index + 1}人目`;
          const manager = managerMap.get(name);
          blocks.push(finalBlockHtml(group, name, index, manager));
        });
      });
      if (!blocks.length) {
        table.innerHTML = `<tbody><tr><td>選手が未設定です</td></tr></tbody>`;
        return;
      }
      table.innerHTML = `<tbody><tr><td style="border:0;padding:0;background:transparent;"><div class="final-grid">${blocks.join("")}</div></td></tr></tbody>`;
    }

    function renderOutputTable(table, groupFilter, outputMode) {
      if (outputMode === "pushOff") {
        renderFinalFormatTable(table, groupFilter, buildSummaryManagerMap());
        return;
      }
      if (outputMode === "dive") {
        renderDiveMergeTable(table, groupFilter);
        return;
      }
      renderSyncedOutputTable(table, groupFilter);
    }

    function renderSyncedOutputTable(table, groupFilter) {
      table.className = "final-table";
      const groups = groupFilter === "all" ? ["A", "BC", "D"] : [groupFilter];
      const managerMap = buildSummaryManagerMap();
      const blocks = [];
      groups.forEach((group) => {
        const usedDive = new Set();
        state.athletes.forEach((athlete, index) => {
          if (athlete.group !== group) return;
          const name = athleteDisplayName(index);
          const diveIndex = state.dive.athletes[group].findIndex((diveName) => diveName === name);
          if (diveIndex >= 0) usedDive.add(diveIndex);
          const pushOff = finalBlockHtml(group, name, index, managerMap.get(name));
          const dive = diveIndex >= 0 ? diveFinalBlockHtml(group, diveIndex, name) : "";
          blocks.push(`<div class="final-page-pair">${pushOff}${dive}</div>`);
        });
        state.dive.athletes[group].forEach((name, index) => {
          if (usedDive.has(index)) return;
          blocks.push(`<div class="final-page-pair">${diveFinalBlockHtml(group, index, name)}</div>`);
        });
      });
      table.innerHTML = blocks.length
        ? `<tbody><tr><td style="border:0;padding:0;background:transparent;"><div class="final-grid">${blocks.join("")}</div></td></tr></tbody>`
        : `<tbody><tr><td>出力するデータがありません</td></tr></tbody>`;
    }

    function buildSummaryManagerMap() {
      const result = new Map();
      state.athletes.forEach((athlete, index) => {
        const name = athleteDisplayName(index);
        const row = ensureMemoRow(index, athlete.name);
        const values = [];
        [1, 2, 3, 4].forEach((setNo) => {
          const hr = row[summaryInputKey(setNo, "HR")] || "";
          const bestAve = row[summaryInputKey(setNo, "BestAve")] || "";
          if (hr) values.push({ key: `S${setNo} HR`, label: `S${setNo} HR`, value: hr });
          if (bestAve) values.push({ key: `S${setNo} BestAve`, label: `S${setNo} BestAve`, value: bestAve });
        });
        const allHr = row[summaryInputKey("All", "HR")] || "";
        const allBestAve = row[summaryInputKey("All", "BestAve")] || "";
        if (allHr) values.push({ key: "All HR", label: "All HR", value: allHr });
        if (allBestAve) values.push({ key: "BestAve", label: "BestAve", value: allBestAve });
        result.set(name, { name, group: athlete.group, values });
      });
      return result;
    }

    function finalBlockHtml(group, name, athleteIndex, manager) {
      const body = group === "D"
        ? finalRowsD(athleteIndex, name, manager)
        : finalRows50(group, athleteIndex, name, manager);
      return `<div class="final-pair group-${group}"><table class="final-table final-block">${body}</table></div>`;
    }

    function finalRows50(group, athleteIndex, name, manager) {
      const setReps = group === "A" ? [4, 4, 4, 4] : [6, 6, 6, 4];
      const config = summaryConfig(group, name);
      const allUntil = summaryAllUntil(config);
      const allSets = allSetNumbers(config);
      const averages = {};
      let html = `<thead><tr class="thick-top"><th class="rep-col thick-left"></th><th class="name-head">${escapeHtml(name)}</th><th class="tempo-head thick-right">Tempo</th></tr></thead><tbody>`;
      setReps.forEach((reps, setIndex) => {
        const setNo = setIndex + 1;
        averages[setNo] = averageSeconds(finalSetTimes50(athleteIndex, group, setNo, reps));
        const mainTimes = allSets.flatMap((targetSet) => finalSetTimes50(athleteIndex, group, targetSet, setReps[targetSet - 1]));
        const mainAllAve = averageSeconds(mainTimes);
        const allBestAve = finalAllBestSeconds(manager);
        const setBestAve = finalBestSeconds(manager, setNo);
        const standardTarget = finalSetTimes50(athleteIndex, group, setNo, reps);
        for (let rep = 1; rep <= reps; rep += 1) {
          const firstClass = rep === 1 ? "thick-top" : "";
          html += `<tr class="${firstClass}"><td class="rep-col thick-left">${rep}</td><td>${finalMemoValue(athleteIndex, group, setNo, rep, "50m", "time")}</td><td class="thick-right">${finalTempoPairValue(athleteIndex, group, setNo, rep, "15m", "35m", config)}</td></tr>`;
        }
        if (setOption(config, setNo, "Hr")) html += finalSummaryRow("HR", finalManagerValue(manager, `S${setNo} HR`), "blue", 3);
        if (setOption(config, setNo, "Average")) html += finalSummaryRow("Average", formatAverageSeconds(averages[setNo]), "red", 3);
        if (setOption(config, setNo, "Best")) {
          html += finalSummaryRow("BestAve", finalBestValue(manager, setNo), "magenta", 3);
          html += finalSummaryRow("達成率", finalPercentValue(manager, setBestAve, averages[setNo], `S${setNo} 達成率`), "magenta", 3);
        }
        if (setOption(config, setNo, "StdDev")) html += finalSummaryRow("標準偏差", formatStdDev(standardDeviationSeconds(standardTarget)), "purple", 3);
        if (allUntil > 0 && setNo === allUntil) {
          if (config.allHr) html += finalSummaryRow("HR", finalManagerValue(manager, "All HR"), "blue", 3);
          if (config.allAve) html += finalSummaryRow("AllAve", formatAverageSeconds(mainAllAve), "red", 3);
          if (config.allStdDev) html += finalSummaryRow("標準偏差", formatStdDev(standardDeviationSeconds(mainTimes)), "purple", 3);
          if (config.allBest) {
            html += finalSummaryRow("BestAve", escapeHtml(finalAllBestValueRaw(manager)), "magenta", 3);
            html += finalSummaryRow("達成率", finalPercentValue(manager, allBestAve, mainAllAve, "達成率"), "magenta", 3);
          }
        }
      });
      return `${html}</tbody>`;
    }

    function finalRowsD(athleteIndex, name, manager) {
      const config = summaryConfig("D", name);
      const allSets = allSetNumbers(config);
      const dAverages = {
        1: averageSeconds(finalSetTimesD(athleteIndex, 1, "100")),
        2: averageSeconds(finalSetTimesD(athleteIndex, 2, "100")),
        3: averageSeconds(finalSetTimesD(athleteIndex, 3, "100")),
        4: averageSeconds(finalSetTimesD(athleteIndex, 4, "50"))
      };
      const frontAverages = {
        1: averageSeconds(finalSetTimesD(athleteIndex, 1, "50")),
        2: averageSeconds(finalSetTimesD(athleteIndex, 2, "50")),
        3: averageSeconds(finalSetTimesD(athleteIndex, 3, "50"))
      };
      const backAverages = {
        1: averageSeconds(finalBackHalfTimesD(athleteIndex, 1)),
        2: averageSeconds(finalBackHalfTimesD(athleteIndex, 2)),
        3: averageSeconds(finalBackHalfTimesD(athleteIndex, 3))
      };
      const mainAllAve = averageSeconds(allSets.flatMap((setNo) => finalSetTimesD(athleteIndex, setNo, setNo === 4 ? "50" : "100")));
      const frontSets = allSets.filter((setNo) => setNo <= 3);
      const mainFrontAve = averageSeconds(frontSets.flatMap((setNo) => finalSetTimesD(athleteIndex, setNo, "50")));
      const mainBackAve = averageSeconds(frontSets.flatMap((setNo) => finalBackHalfTimesD(athleteIndex, setNo)));
      const allBestAve = finalAllBestSeconds(manager);
      const mainStdDev = standardDeviationSeconds(allSets.flatMap((setNo) => finalSetTimesD(athleteIndex, setNo, setNo === 4 ? "50" : "100")));
      const set4StdDev = standardDeviationSeconds(finalSetTimesD(athleteIndex, 4, "50"));
      const set4Total = sumSeconds(finalSetTimesD(athleteIndex, 4, "50"));
      const bestTotal = parseTimeSeconds(finalManagerValueRaw(manager, "S4 BestTotal"));
      const allRange = summaryAllRange(config);
      const allUntil = allRange.until;
      const allSummaryRows = () => {
        let rows = "";
        if (config.dAllFrontBack) {
          rows += finalDTargetSummaryRow("前半Ave", allRange, formatAverageSeconds(mainFrontAve), "gold");
          rows += finalDTargetSummaryRow("後半Ave", allRange, formatAverageSeconds(mainBackAve), "gold");
        }
        if (config.allAve) rows += finalDTargetSummaryRow("AllAve", allRange, formatAverageSeconds(mainAllAve), "red");
        if (config.allHr || config.dAllHr) rows += finalDTargetSummaryRow("HR", allRange, finalManagerValue(manager, "All HR"), "blue");
        if (config.allBest) {
          rows += finalDTargetSummaryRow("BestAve", allRange, escapeHtml(finalAllBestValueRaw(manager)), "magenta");
          rows += finalDTargetSummaryRow("達成率", allRange, finalPercentValue(manager, allBestAve, mainAllAve, "達成率"), "magenta");
        }
        if (config.allStdDev) rows += finalDTargetSummaryRow("標準偏差", allRange, formatStdDev(mainStdDev), "purple");
        return rows;
      };
      let html = `<thead><tr class="thick-top"><th class="rep-col thick-left"></th>${[1, 2, 3].map((setNo) => `<th></th><th class="name-head">${setNo === 2 ? escapeHtml(name) : ""}</th><th class="tempo-head thick-right">Tempo</th>`).join("")}</tr></thead><tbody>`;
      for (let rep = 1; rep <= 6; rep += 1) {
        const firstClass = rep === 1 ? "thick-top" : "";
        html += `<tr class="${firstClass}"><td class="rep-col thick-left">${rep}</td>${[1, 2, 3].map((setNo) => finalDSetCells(athleteIndex, setNo, rep, "50", config)).join("")}</tr>`;
        html += `<tr><td class="rep-col thick-left"></td>${[1, 2, 3].map((setNo) => finalDSetCells(athleteIndex, setNo, rep, "100", config)).join("")}</tr>`;
      }
      if ([1, 2, 3].some((setNo) => setOption(config, setNo, "Hr"))) html += finalDSetSummaryRow("HR", [1, 2, 3].map((setNo) => setOption(config, setNo, "Hr") ? finalManagerValue(manager, `S${setNo} HR`) : ""), "blue");
      if ([1, 2, 3].some((setNo) => setOption(config, setNo, "FrontBack"))) {
        html += finalDSetSummaryRow("前半Ave", [1, 2, 3].map((setNo) => setOption(config, setNo, "FrontBack") ? formatAverageSeconds(frontAverages[setNo]) : ""), "gold");
        html += finalDSetSummaryRow("後半Ave", [1, 2, 3].map((setNo) => setOption(config, setNo, "FrontBack") ? formatAverageSeconds(backAverages[setNo]) : ""), "gold");
      }
      if ([1, 2, 3].some((setNo) => setOption(config, setNo, "Average"))) html += finalDSetSummaryRow("Average", [1, 2, 3].map((setNo) => setOption(config, setNo, "Average") ? formatAverageSeconds(dAverages[setNo]) : ""), "red");
      if ([1, 2, 3].some((setNo) => setOption(config, setNo, "Best"))) {
        html += finalDSetSummaryRow("BestAve", [1, 2, 3].map((setNo) => setOption(config, setNo, "Best") ? finalBestValue(manager, setNo) : ""), "magenta");
        html += finalDSetSummaryRow("達成率", [1, 2, 3].map((setNo) => setOption(config, setNo, "Best") ? finalPercentValue(manager, finalBestSeconds(manager, setNo), dAverages[setNo], `S${setNo} 達成率`) : ""), "magenta");
      }
      if ([1, 2, 3].some((setNo) => setOption(config, setNo, "StdDev"))) html += finalDSetSummaryRow("標準偏差", [1, 2, 3].map((setNo) => setOption(config, setNo, "StdDev") ? formatStdDev(standardDeviationSeconds(finalSetTimesD(athleteIndex, setNo, "100"))) : ""), "purple");
      if (allUntil > 0 && allUntil !== 4) html += allSummaryRows();
      for (let rep = 1; rep <= 4; rep += 1) {
        const firstClass = rep === 1 ? "thick-top" : "";
        html += `<tr class="${firstClass}"><td class="rep-col thick-left">${rep}</td><td>${finalMemoValue(athleteIndex, "D", 4, rep, "50m", "time")}</td><td class="thick-right">${finalTempoPairValue(athleteIndex, "D", 4, rep, "15m", "35m", config)}</td><td colspan="7" class="clear-cell"></td></tr>`;
      }
      if (setOption(config, 4, "Hr")) html += finalD50SummaryRow("HR", finalManagerValue(manager, "S4 HR"), "blue");
      if (setOption(config, 4, "Total")) {
        html += finalD50SummaryRow("Total", formatAverageSeconds(set4Total), "red");
        html += finalD50SummaryRow("BestTotal", finalManagerValue(manager, "S4 BestTotal"), "magenta");
        html += finalD50SummaryRow("達成率", finalPercentValue(manager, bestTotal, set4Total, "S4 Total 達成率"), "magenta");
      }
      if (setOption(config, 4, "Average")) html += finalD50SummaryRow("Average", formatAverageSeconds(dAverages[4]), "red");
      if (setOption(config, 4, "Best")) {
        html += finalD50SummaryRow("BestAve", finalBestValue(manager, 4), "magenta");
        html += finalD50SummaryRow("達成率", finalPercentValue(manager, finalBestSeconds(manager, 4), dAverages[4], "S4 達成率"), "magenta");
      }
      if (setOption(config, 4, "StdDev")) html += finalD50SummaryRow("標準偏差", formatStdDev(set4StdDev), "purple");
      if (allUntil === 4) html += allSummaryRows();
      return `${html}</tbody>`;
    }

    function finalDSetCells(athleteIndex, setNo, rep, point, config) {
      const full50 = finalMemoValueRaw(athleteIndex, "D", setNo, rep, "50m", "time");
      const full100 = finalMemoValueRaw(athleteIndex, "D", setNo, rep, "100m", "time");
      if (point === "50") {
        return `<td>${escapeHtml(full50)}</td><td></td><td class="thick-right">${finalTempoPairValue(athleteIndex, "D", setNo, rep, "15m", "35m", config)}</td>`;
      }
      return `<td>${escapeHtml(full100)}</td><td>${finalSplitValue(full100, full50)}</td><td class="thick-right">${finalTempoPairValue(athleteIndex, "D", setNo, rep, "65m", "85m", config)}</td>`;
    }

    function finalDSetSummaryRow(label, values, colorClass) {
      return `<tr class="thick-top"><td class="rep-col thick-left ${colorClass}">${escapeHtml(label)}</td>${values.map((value) => `<td colspan="3" class="${colorClass} thick-right">${value}</td>`).join("")}</tr>`;
    }

    function finalDWholeSummaryRow(label, value, colorClass) {
      return `<tr class="thick-top"><td class="rep-col thick-left ${colorClass}">${escapeHtml(label)}</td><td colspan="9" class="${colorClass} thick-right">${value}</td></tr>`;
    }

    function finalDTargetSummaryRow(label, range, value, colorClass) {
      const from = range?.from || 1;
      const targetSet = range?.until || 0;
      if (targetSet === 4) return finalDAll50SummaryRow(label, value, colorClass);
      const leadingCols = (Math.min(3, Math.max(1, Number(from))) - 1) * 3;
      const activeCols = (Math.min(3, Math.max(1, Number(targetSet))) - Math.min(3, Math.max(1, Number(from))) + 1) * 3;
      const restCols = 9 - leadingCols - activeCols;
      return `<tr class="thick-top"><td class="rep-col thick-left ${colorClass}">${escapeHtml(label)}</td>${leadingCols ? `<td colspan="${leadingCols}" class="clear-cell"></td>` : ""}<td colspan="${activeCols}" class="${colorClass} thick-right merged-summary">${value}</td>${restCols ? `<td colspan="${restCols}" class="clear-cell"></td>` : ""}</tr>`;
    }

    function finalD50SummaryRow(label, value, colorClass) {
      return `<tr class="thick-top"><td class="rep-col thick-left ${colorClass}">${escapeHtml(label)}</td><td colspan="2" class="${colorClass} thick-right merged-summary">${value}</td><td colspan="7" class="clear-cell"></td></tr>`;
    }

    function finalDAll50SummaryRow(label, value, colorClass) {
      return `<tr class="thick-top"><td class="rep-col thick-left ${colorClass}">${escapeHtml(label)}</td><td colspan="3" class="${colorClass} thick-right merged-summary d-merged-50">${value}</td><td colspan="6" class="clear-cell"></td></tr>`;
    }

    function finalSummaryRow(label, value, colorClass, colspan) {
      return `<tr class="thick-top"><td class="rep-col thick-left ${colorClass}">${escapeHtml(label)}</td><td colspan="${colspan - 1}" class="${colorClass} thick-right">${escapeHtml(value)}</td></tr>`;
    }

    function finalMemoValue(athleteIndex, group, setNo, rep, position, type) {
      return escapeHtml(finalMemoValueRaw(athleteIndex, group, setNo, rep, position, type));
    }

    function finalTempoValue(athleteIndex, group, setNo, rep, position, config) {
      if (!setOption(config, setNo, "Tempo")) return "-";
      return finalMemoValue(athleteIndex, group, setNo, rep, position, "tempo");
    }

    function finalTempoPairValue(athleteIndex, group, setNo, rep, firstPosition, secondPosition, config) {
      if (!setOption(config, setNo, "Tempo")) return "-";
      const first = finalMemoValueRaw(athleteIndex, group, setNo, rep, firstPosition, "tempo");
      const second = finalMemoValueRaw(athleteIndex, group, setNo, rep, secondPosition, "tempo");
      return `${escapeHtml(first)}/${escapeHtml(second)}`;
    }

    function finalMemoValueRaw(athleteIndex, group, setNo, rep, position, type) {
      const athlete = state.athletes[athleteIndex];
      if (!athlete) return "";
      const row = ensureMemoRow(athleteIndex, athlete.name);
      const col = flattenMenu(group, type).find((item) => logicalSetNumber(group, item) === setNo && item.rep === rep && item.position === position);
      return col ? row[col.key] || "" : "";
    }

    function logicalSetNumber(group, item) {
      if ((group === "BC" || group === "D") && item.segment === "50x4") return 4;
      return item.block;
    }

    function finalManagerValue(manager, label) {
      return escapeHtml(finalManagerValueRaw(manager, label));
    }

    function finalManagerValueRaw(manager, label) {
      if (!manager) return "";
      const item = manager.values.find((value) => value.label === label || value.key === label);
      return item?.value || "";
    }

    function allSetNumbers(config) {
      const { from, until } = summaryAllRange(config);
      if (until <= 0) return [];
      return Array.from({ length: until - from + 1 }, (_, index) => from + index);
    }

    function finalBestValue(manager, setNo) {
      return escapeHtml(finalBestValueRaw(manager, setNo));
    }

    function finalBestValueRaw(manager, setNo) {
      return finalManagerValueRaw(manager, `S${setNo} BestAve`)
        || finalManagerValueRaw(manager, `S${setNo} BestAverage`)
        || finalManagerValueRaw(manager, `S${setNo} Best`)
        || finalManagerValueRaw(manager, "BestAve")
        || finalManagerValueRaw(manager, "BestAverage");
    }

    function finalBestSeconds(manager, setNo) {
      return parseTimeSeconds(finalBestValueRaw(manager, setNo));
    }

    function finalAllBestValueRaw(manager) {
      return finalManagerValueRaw(manager, "BestAve")
        || finalManagerValueRaw(manager, "BestAverage")
        || finalManagerValueRaw(manager, "Best");
    }

    function finalAllBestSeconds(manager) {
      return parseTimeSeconds(finalAllBestValueRaw(manager));
    }

    function finalPercentValue(manager, numerator, denominator, fallbackLabel) {
      const calculated = formatPercent(numerator, denominator);
      if (calculated) return calculated;
      const fallback = finalManagerValueRaw(manager, fallbackLabel);
      if (!fallback) return "";
      const seconds = parseTimeSeconds(fallback);
      if (Number.isFinite(seconds) && seconds > 0 && seconds <= 2) return `${(seconds * 100).toFixed(2)}%`;
      return escapeHtml(fallback);
    }

    function finalSetTimes50(athleteIndex, group, setNo, reps) {
      return Array.from({ length: reps }, (_, index) => parseTimeSeconds(finalMemoValueRaw(athleteIndex, group, setNo, index + 1, "50m", "time")))
        .filter((value) => value !== null);
    }

    function finalSetTimesD(athleteIndex, setNo, distance) {
      const reps = setNo === 4 ? 4 : 6;
      return Array.from({ length: reps }, (_, index) => {
        const rep = index + 1;
        return parseTimeSeconds(finalMemoValueRaw(athleteIndex, "D", setNo, rep, `${distance}m`, "time"));
      }).filter((value) => value !== null);
    }

    function finalBackHalfTimesD(athleteIndex, setNo) {
      const values = [];
      for (let rep = 1; rep <= 6; rep += 1) {
        const full100 = parseTimeSeconds(finalMemoValueRaw(athleteIndex, "D", setNo, rep, "100m", "time"));
        const full50 = parseTimeSeconds(finalMemoValueRaw(athleteIndex, "D", setNo, rep, "50m", "time"));
        if (full100 !== null && full50 !== null && full100 >= full50) values.push(full100 - full50);
      }
      return values;
    }

    function averageSeconds(values) {
      const valid = values.filter((value) => Number.isFinite(value));
      if (!valid.length) return null;
      return valid.reduce((sum, value) => sum + value, 0) / valid.length;
    }

    function sumSeconds(values) {
      const valid = values.filter((value) => Number.isFinite(value));
      if (!valid.length) return null;
      return valid.reduce((sum, value) => sum + value, 0);
    }

    function standardDeviationSeconds(values) {
      const valid = values.filter((value) => Number.isFinite(value));
      if (valid.length < 2) return null;
      const average = averageSeconds(valid);
      const variance = valid.reduce((sum, value) => sum + (value - average) ** 2, 0) / valid.length;
      return Math.sqrt(variance);
    }

    function formatAverageSeconds(value) {
      if (!Number.isFinite(value)) return "";
      return formatTimeSeconds(value, 2);
    }

    function formatStdDev(value) {
      if (!Number.isFinite(value)) return "";
      return value.toFixed(3);
    }

    function formatPercent(numerator, denominator) {
      if (!Number.isFinite(numerator) || !Number.isFinite(denominator) || denominator === 0) return "";
      return `${((numerator / denominator) * 100).toFixed(2)}%`;
    }

    function parseTimeSeconds(value) {
      const text = String(value || "").trim().replace(/[()]/g, "");
      if (!text || text === "-" || text.startsWith("#")) return null;
      if (/^\d+:\d+(?:\.\d+)?$/.test(text)) {
        const [minutes, seconds] = text.split(":");
        return Number(minutes) * 60 + Number(seconds);
      }
      const number = Number(text);
      return Number.isFinite(number) ? number : null;
    }

    function formatTimeSeconds(value, decimals = 1) {
      if (!Number.isFinite(value)) return "";
      const factor = 10 ** decimals;
      const rounded = Math.round(value * factor) / factor;
      if (rounded >= 60) {
        const minutes = Math.floor(rounded / 60);
        const seconds = (rounded - minutes * 60).toFixed(decimals).padStart(decimals + 3, "0");
        return `${minutes}:${seconds}`;
      }
      return rounded.toFixed(decimals);
    }

    function finalSplitValue(total, previous) {
      const totalSeconds = parseTimeSeconds(total);
      const previousSeconds = parseTimeSeconds(previous);
      if (totalSeconds === null || previousSeconds === null) return "";
      const split = totalSeconds - previousSeconds;
      if (split < 0) return "";
      return escapeHtml(formatTimeSeconds(split, 1));
    }

    function athleteDisplayName(index) {
      const athlete = state.athletes[index];
      return athlete?.name?.trim() || `${athlete?.group || ""}班 ${index + 1}人目`;
    }

    function renderChartAthleteOptions() {
      const select = $("chartAthlete");
      if (!select) return;
      const group = selectedSettingsGroup();
      const current = select.value;
      const options = state.athletes
        .map((athlete, index) => ({ athlete, index }))
        .filter((item) => item.athlete.group === group);
      select.innerHTML = options.length
        ? options.map(({ index }) => `<option value="${index}">${escapeHtml(athleteDisplayName(index))}</option>`).join("")
        : `<option value="">選手未設定</option>`;
      if (current && [...select.options].some((option) => option.value === current)) {
        select.value = current;
      } else if (options.length) {
        select.value = String(options[0].index);
      }
      renderSummarySettings();
    }

    function rowEl(values) {
      const tr = document.createElement("tr");
      tr.innerHTML = values.map((value) => `<td>${escapeHtml(value)}</td>`).join("");
      return tr;
    }

    function unique(values) {
      return [...new Set(values)];
    }

    function tableToTsv(table) {
      return [...table.querySelectorAll("tr")]
        .map((tr) => [...tr.children].map((cell) => {
          const input = cell.querySelector("input,select");
          return input ? input.value : cell.textContent;
        }).join("\t"))
        .join("\n");
    }

    async function copyTable(tableId, statusId) {
      const text = tableToTsv($(tableId));
      await navigator.clipboard.writeText(text);
      if (statusId) {
        $(statusId).textContent = "コピーしました";
        $(statusId).className = "status ok";
      }
    }

    function exportMergePdf() {
      document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === "merge"));
      document.querySelectorAll(".view").forEach((view) => view.classList.add("hidden"));
      $("mergeView").classList.remove("hidden");
      $("mergeGroup").value = $("outputGroup")?.value || "all";
      renderOutputTable($("mergeTable"), $("mergeGroup").value, $("outputMode")?.value || "sync");
      $("mergeStatus").textContent = "印刷画面で「PDFとして保存」を選んでください";
      $("mergeStatus").className = "status ok";
      $("outputStatus").textContent = "印刷画面を開きます";
      $("outputStatus").className = "status ok";
      setTimeout(() => window.print(), 100);
    }

    function syncSession() {
      state.session.group = $("groupSelect").value;
      state.session.date = $("sessionDate")?.value || "";
      state.session.stroke = $("stroke")?.value || "";
      state.session.note = $("sessionNote")?.value || "";
    }

    function updateMetrics() {
      $("athleteCount").textContent = state.athletes.filter((a) => a.name.trim()).length;
      $("memoCount").textContent = Object.values(state.memo).reduce((sum, row) => sum + Object.values(row).filter(Boolean).length, 0);
      if ($("managerCount")) $("managerCount").textContent = state.managerRows.length;
    }

    function save() {
      syncSession();
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ savedAt: Date.now(), state }));
    }

    function persistSilently() {
      syncSession();
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ savedAt: Date.now(), state }));
    }

    function scheduleLiveUpdate(options = {}) {
      clearTimeout(liveUpdateTimer);
      liveUpdateTimer = setTimeout(() => {
        if (options.manager) {
          refreshManagerFromPaste(false);
        } else {
          renderManager();
          renderMerge();
          updateMetrics();
          persistSilently();
        }
      }, 80);
    }

    function load() {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return false;
      const stored = JSON.parse(raw);
      const loaded = stored.state || stored;
      if (!loaded || !Array.isArray(loaded.athletes) || typeof loaded.memo !== "object") {
        return false;
      }
      Object.assign(state, loaded);
      ensureStateShape();
      $("groupSelect").value = state.session.group || "A";
      if ($("memoGroup")) $("memoGroup").value = state.session.group || "A";
      if ($("sessionDate")) $("sessionDate").value = state.session.date || today();
      if ($("stroke")) $("stroke").value = state.session.stroke || "";
      if ($("sessionNote")) $("sessionNote").value = state.session.note || "";
      renderAll();
      persistSilently();
      return true;
    }

    function currentStoragePayload() {
      syncSession();
      return { savedAt: Date.now(), state };
    }

    function exportBackup() {
      const payload = currentStoragePayload();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
      const stamp = new Date().toISOString().replace(/[:.]/g, "-");
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `swimDataApp-backup-${stamp}.json`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    }

    function importBackupFile(file) {
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const imported = JSON.parse(reader.result);
          const importedState = imported.state || imported;
          if (!importedState || !Array.isArray(importedState.athletes) || typeof importedState.memo !== "object") {
            throw new Error("バックアップファイルの形式が違います");
          }
          Object.assign(state, importedState);
          ensureStateShape();
          localStorage.setItem(STORAGE_KEY, JSON.stringify({ savedAt: Date.now(), state }));
          renderAll();
          alert("バックアップを読み込みました");
        } catch (error) {
          alert(`バックアップを読み込めませんでした: ${error.message}`);
        }
      };
      reader.readAsText(file);
    }

    function renderAll() {
      ensureStateShape();
      if ($("groupSelect")) $("groupSelect").value = state.session.group || "A";
      if ($("memoGroup")) $("memoGroup").value = state.session.group || "A";
      renderAthletes();
      renderMemo();
      renderManagerPasteGrid();
      renderManager();
      renderMerge();
      renderChartAthleteOptions();
      renderDiveSettings();
      renderMeasureMode();
    }

    function syncDiveFromPushOff() {
      ensureStateShape();
      const group = state.session.group;
      state.dive.athletes[group] = state.athletes
        .filter((athlete) => athlete.group === group)
        .map((athlete, index) => athlete.name.trim() || `${group}班 ${index + 1}人目`);
      state.dive.settings[group] = state.dive.athletes[group].map((_, index) => {
        const current = state.dive.settings[group]?.[index] || {};
        return {
          distance: Math.max(1, Math.floor(Number(current.distance ?? state.dive.distances[group]) || 50)),
          tempo: current.tempo && typeof current.tempo === "object" ? current.tempo : {}
        };
      });
      setDiveConfigSelection(group, 0);
      state.session.measureMode = "dive";
      renderAll();
      persistSilently();
    }

    function addDiveAthlete(name = "", group = state.session.group) {
      ensureStateShape();
      state.dive.athletes[group].push(name);
      state.dive.settings[group].push({ distance: state.dive.distances[group] || 50, tempo: {} });
      setDiveConfigSelection(group, state.dive.athletes[group].length - 1);
      state.session.measureMode = "dive";
      state.session.group = group;
      renderAll();
      persistSilently();
    }

    function moveDiveAthlete(index, direction, group = state.session.group) {
      ensureStateShape();
      const next = index + direction;
      if (index < 0 || next < 0 || next >= state.dive.athletes[group].length) return;
      [state.dive.athletes[group][index], state.dive.athletes[group][next]] = [state.dive.athletes[group][next], state.dive.athletes[group][index]];
      [state.dive.settings[group][index], state.dive.settings[group][next]] = [state.dive.settings[group][next], state.dive.settings[group][index]];
      setDiveConfigSelection(group, next);
      renderAll();
      persistSilently();
    }

    function deleteDiveAthlete(index, group = state.session.group) {
      ensureStateShape();
      state.dive.athletes[group].splice(index, 1);
      state.dive.settings[group].splice(index, 1);
      setDiveConfigSelection(group, Math.min(index, Math.max(0, state.dive.athletes[group].length - 1)));
      renderAll();
      persistSilently();
    }

    function scheduleAthleteNameUpdate(index) {
      const select = $("chartAthlete");
      const option = select?.querySelector(`option[value="${index}"]`);
      if (option) option.textContent = athleteDisplayName(index);
      clearTimeout(athleteNameUpdateTimer);
      athleteNameUpdateTimer = setTimeout(() => {
        renderMemo();
        renderManagerPasteGrid();
        renderManager();
        renderMerge();
        renderChartAthleteOptions();
        persistSilently();
      }, 220);
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
    }

    function escapeAttr(value) {
      return escapeHtml(value).replace(/`/g, "&#96;");
    }

    document.querySelectorAll(".tab").forEach((tab) => {
      tab.addEventListener("click", () => {
        document.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
        document.querySelectorAll(".view").forEach((view) => view.classList.add("hidden"));
        tab.classList.add("active");
        $(`${tab.dataset.view}View`).classList.remove("hidden");
        if (tab.dataset.view === "settings") renderSummarySettings();
        if (tab.dataset.view === "merge") renderMerge();
      });
    });

    document.addEventListener("input", (event) => {
      const index = Number(event.target.dataset.athleteIndex);
      const field = event.target.dataset.field;
      if (!Number.isInteger(index) || !field) return;
      const oldKey = athleteKey(index, state.athletes[index].name);
      state.athletes[index][field] = event.target.value;
      const newKey = athleteKey(index, state.athletes[index].name);
      if (field === "name" && oldKey !== newKey && state.memo[oldKey] && !state.memo[newKey]) {
        state.memo[newKey] = state.memo[oldKey];
        delete state.memo[oldKey];
      }
      if (field === "name") {
        updateMetrics();
        scheduleAthleteNameUpdate(index);
        persistSilently();
        return;
      }
      renderAll();
      persistSilently();
    });

    document.addEventListener("click", async (event) => {
      const selectGroup = event.target.dataset.selectGroup;
      const addGroup = event.target.dataset.addGroup;
      const pasteGroup = event.target.dataset.pasteGroup;
      const measureMode = event.target.dataset.measureMode;
      const up = event.target.dataset.moveUp;
      const down = event.target.dataset.moveDown;
      const del = event.target.dataset.delete;
      const addDive = event.target.dataset.addDiveAthlete;
      const pasteDive = event.target.dataset.pasteDiveAthletes;
      const diveUp = event.target.dataset.diveMoveUp;
      const diveDown = event.target.dataset.diveMoveDown;
      const diveDel = event.target.dataset.diveDelete;
      const diveSelect = event.target.dataset.diveSelect;

      if (measureMode) {
        state.session.measureMode = measureMode;
        renderAll();
        persistSilently();
      } else if (addDive) {
        addDiveAthlete("", addDive);
      } else if (pasteDive) {
        const text = await navigator.clipboard.readText();
        const names = text.split(/\r?\n/).map((line) => line.split("\t")[0].trim()).filter(Boolean);
        names.forEach((name) => addDiveAthlete(name, pasteDive));
      } else if (diveSelect !== undefined) {
        const group = event.target.dataset.diveActionGroup || state.session.group;
        setDiveConfigSelection(group, Number(diveSelect));
        renderDiveSettings();
        persistSilently();
      } else if (diveUp !== undefined) {
        moveDiveAthlete(Number(diveUp), -1, event.target.dataset.diveActionGroup);
      } else if (diveDown !== undefined) {
        moveDiveAthlete(Number(diveDown), 1, event.target.dataset.diveActionGroup);
      } else if (diveDel !== undefined) {
        deleteDiveAthlete(Number(diveDel), event.target.dataset.diveActionGroup);
      } else if (selectGroup) {
        setActiveGroup(selectGroup);
      } else if (addGroup) {
        state.athletes.push({ name: "", group: addGroup });
        setActiveGroup(addGroup);
      } else if (pasteGroup) {
        const text = await navigator.clipboard.readText();
        const names = text.split(/\r?\n/).map((line) => line.split("\t")[0].trim()).filter(Boolean);
        state.athletes = state.athletes.filter((athlete) => athlete.group !== pasteGroup);
        state.athletes.push(...names.map((name) => ({ name, group: pasteGroup })));
        setActiveGroup(pasteGroup);
      } else if (up !== undefined) {
        moveAthleteInGroup(Number(up), -1);
        renderAll();
      } else if (down !== undefined) {
        moveAthleteInGroup(Number(down), 1);
        renderAll();
      } else if (del !== undefined) {
        state.athletes.splice(Number(del), 1);
        renderAll();
      } else {
        return;
      }
    });

    function moveAthleteInGroup(index, direction) {
      const group = state.athletes[index]?.group;
      if (!group) return;
      const indexes = groupIndexes(group).map((item) => item.index);
      const position = indexes.indexOf(index);
      const nextPosition = position + direction;
      if (position < 0 || nextPosition < 0 || nextPosition >= indexes.length) return;
      const otherIndex = indexes[nextPosition];
      [state.athletes[index], state.athletes[otherIndex]] = [state.athletes[otherIndex], state.athletes[index]];
    }

    function handleMemoInput(event) {
      const cell = event.target.closest("[data-memo-key]");
      if (!cell || cell.dataset.active !== "1") return;
      updateMemoCellValue(cell, event.target.textContent.trim());
      updateMetrics();
      scheduleLiveUpdate();
    }

    function moveNextOnEnter(event) {
      if (event.key !== "Enter" || !event.target.closest("[data-memo-key]")) return;
      event.preventDefault();
      const cell = event.target.closest("td");
      const row = cell?.closest("tr");
      const table = cell?.closest("table");
      if (!cell || !row || !table) return;
      const columnIndex = [...row.children].indexOf(cell);
      const rows = [...table.querySelectorAll("tbody tr")];
      const rowIndex = rows.indexOf(row);
      const editableSelector = ".paste-cell[data-active='1'] div[contenteditable='true']";
      let nextInput = rows.slice(rowIndex + 1)
        .map((item) => item.children[columnIndex]?.querySelector(editableSelector))
        .find(Boolean);
      if (!nextInput) {
        const columnCount = Math.max(...rows.map((item) => item.children.length));
        for (let col = columnIndex + 1; col < columnCount; col += 1) {
          nextInput = rows
            .map((item) => item.children[col]?.querySelector(editableSelector))
            .find(Boolean);
          if (nextInput) break;
        }
      }
      if (!nextInput) nextInput = rows.find((item) => item.querySelector(editableSelector))?.querySelector(editableSelector);
      if (nextInput) {
        nextInput.focus();
        selectEditableContents(nextInput);
      }
    }

    function selectEditableContents(editable) {
      const range = document.createRange();
      range.selectNodeContents(editable);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    }

    function handleMemoGridPaste(event) {
      const targetCell = event.target.closest(".paste-cell[data-memo-grid='1']");
      if (!targetCell || targetCell.dataset.active !== "1") return;
      event.preventDefault();
      const text = event.clipboardData.getData("text/plain");
      const rows = trimEmptyPasteEdges(normalizePastedGrid(text));
      if (!rows.length) return;
      const table = targetCell.closest("table");
      const startRow = Number(targetCell.dataset.pasteRow);
      const startCol = Number(targetCell.dataset.pasteCol);
      rows.forEach((row, rowOffset) => {
        row.forEach((value, colOffset) => {
          const cell = table.querySelector(`.paste-cell[data-memo-grid='1'][data-paste-row="${startRow + rowOffset}"][data-paste-col="${startCol + colOffset}"]`);
          if (!cell || cell.dataset.active !== "1") return;
          const editable = cell.querySelector("div[contenteditable='true']");
          if (!editable) return;
          editable.textContent = value || "";
          updateMemoCellValue(cell, editable.textContent.trim());
        });
      });
      updateMetrics();
      scheduleLiveUpdate();
    }

    function updateMemoCellValue(cell, value) {
      const athlete = cell?.dataset.memoAthlete;
      const key = cell?.dataset.memoKey;
      if (!athlete || !key) return;
      if (!state.memo[athlete]) state.memo[athlete] = {};
      state.memo[athlete][key] = value;
    }

    $("timeTable").addEventListener("input", handleMemoInput);
    $("tempoTable").addEventListener("input", handleMemoInput);
    $("diveTimeTable").addEventListener("input", handleMemoInput);
    $("diveTempoTable").addEventListener("input", handleMemoInput);
    $("timeTable").addEventListener("keydown", moveNextOnEnter);
    $("tempoTable").addEventListener("keydown", moveNextOnEnter);
    $("diveTimeTable").addEventListener("keydown", moveNextOnEnter);
    $("diveTempoTable").addEventListener("keydown", moveNextOnEnter);
    $("timeTable").addEventListener("paste", handleMemoGridPaste);
    $("tempoTable").addEventListener("paste", handleMemoGridPaste);
    $("diveTimeTable").addEventListener("paste", handleMemoGridPaste);
    $("diveTempoTable").addEventListener("paste", handleMemoGridPaste);

    $("groupSelect").addEventListener("change", () => {
      setActiveGroup($("groupSelect").value);
      persistSilently();
    });
    $("memoGroup").addEventListener("change", () => {
      setActiveGroup($("memoGroup").value);
      persistSilently();
    });
    document.addEventListener("input", (event) => {
      ensureStateShape();
      const nameIndex = Number(event.target.dataset.diveNameIndex);
      if (Number.isInteger(nameIndex)) {
        const group = event.target.dataset.diveNameGroup || state.session.group;
        if (!state.dive.athletes[group]) return;
        const oldName = state.dive.athletes[group][nameIndex] || "";
        const oldKey = diveAthleteKey(group, nameIndex, oldName);
        state.dive.athletes[group][nameIndex] = event.target.value;
        const newKey = diveAthleteKey(group, nameIndex, event.target.value);
        if (oldKey !== newKey && state.memo[oldKey] && !state.memo[newKey]) {
          state.memo[newKey] = state.memo[oldKey];
          delete state.memo[oldKey];
        }
        renderDive();
        renderMerge();
        updateMetrics();
        persistSilently();
        return;
      }
      const index = Number(event.target.dataset.diveDistanceIndex);
      if (!Number.isInteger(index)) return;
      const group = event.target.dataset.diveDistanceGroup || state.session.group;
      if (!state.dive.settings[group]?.[index]) return;
      state.dive.settings[group][index].distance = Math.max(1, Math.floor(Number(event.target.value) || 1));
      renderDive();
      renderMerge();
      updateMetrics();
      persistSilently();
    });
    document.addEventListener("change", (event) => {
      ensureStateShape();
      const distanceIndex = Number(event.target.dataset.diveDistanceIndex);
      if (Number.isInteger(distanceIndex)) {
        const group = event.target.dataset.diveDistanceGroup || state.session.group;
        setDiveConfigSelection(group, distanceIndex);
        renderDiveSettings();
        renderDive();
        renderMerge();
        persistSilently();
        return;
      }
      const index = Number(event.target.dataset.diveTempoIndex);
      const distance = Number(event.target.dataset.diveTempoDistance);
      if (!Number.isInteger(index) || !distance) return;
      const group = event.target.dataset.diveTempoGroup || state.session.group;
      if (!state.dive.settings[group]?.[index]) return;
      state.dive.settings[group][index].tempo[distance] = event.target.checked;
      setDiveConfigSelection(group, index);
      renderDiveSettings();
      renderDive();
      renderMerge();
      persistSilently();
    });
    $("diveConfigAthlete")?.addEventListener("change", () => {
      const [group, rawIndex] = $("diveConfigAthlete").value.split(":");
      if (!group) return;
      setDiveConfigSelection(group, Number(rawIndex));
      renderDiveSettings();
      persistSilently();
    });
    $("syncDiveSettingsBtn").addEventListener("click", syncDiveFromPushOff);
    $("removeBlankDiveNamesBtn").addEventListener("click", () => {
      ensureStateShape();
      const group = state.session.group;
      const keepIndexes = state.dive.athletes[group]
        .map((name, index) => ({ name, index }))
        .filter((item) => item.name.trim())
        .map((item) => item.index);
      state.dive.athletes[group] = keepIndexes.map((index) => state.dive.athletes[group][index]);
      state.dive.settings[group] = keepIndexes.map((index) => state.dive.settings[group][index]);
      setDiveConfigSelection(group, 0);
      renderAll();
      persistSilently();
    });
    $("sessionDate")?.addEventListener("input", syncSession);
    $("stroke")?.addEventListener("input", syncSession);
    $("sessionNote")?.addEventListener("input", syncSession);
    $("mergeGroup").addEventListener("change", renderMerge);
    $("outputGroup").addEventListener("change", () => {
      $("outputStatus").textContent = "PDF出力時に反映されます";
      $("outputStatus").className = "status";
    });
    $("outputMode").addEventListener("change", () => {
      $("outputStatus").textContent = "PDF出力時に反映されます";
      $("outputStatus").className = "status";
    });
    $("mergeMode")?.addEventListener("change", renderMerge);
    $("pasteGroup")?.addEventListener("change", renderManagerPasteGrid);
    $("summaryAthlete")?.addEventListener("change", renderSummarySettings);
    $("chartAthlete").addEventListener("change", () => {
      renderSummarySettings();
    });
    $("summarySettings")?.addEventListener("change", (event) => {
      const key = event.target.dataset.summarySetting;
      if (!key) return;
      const selected = selectedChartAthlete();
      const group = selected?.athlete?.group || selectedSettingsGroup();
      const name = selected ? athleteDisplayName(selected.index) : selectedSummaryName(group);
      const value = event.target.type === "checkbox"
        ? event.target.checked
        : key === "allRange" ? event.target.value : Number(event.target.value);
      setSummaryConfigValue(group, name, key, value);
      renderSummarySettings();
      renderMemo();
      renderMerge();
      persistSilently();
    });
    $("pasteFormat")?.addEventListener("change", renderManagerPasteGrid);

    $("removeBlankNamesBtn").addEventListener("click", () => {
      state.athletes = state.athletes.filter((a) => a.name.trim());
      renderAll();
    });

    $("parseManagerBtn")?.addEventListener("click", parseManager);
    $("managerPasteGrid")?.addEventListener("mousedown", handleManagerGridPointer);
    $("managerPasteGrid")?.addEventListener("focusin", handleManagerGridPointer);
    $("managerPasteGrid")?.addEventListener("input", () => scheduleLiveUpdate({ manager: true }));
    $("managerPasteGrid")?.addEventListener("paste", handleManagerGridPaste);
    document.addEventListener("paste", handleDocumentPaste);
    document.addEventListener("keydown", handleControlPaste);
    $("clearManagerBtn")?.addEventListener("click", () => {
      $("managerPasteGrid").querySelectorAll(".paste-cell div").forEach((cell) => { cell.textContent = ""; });
      if (state.managerPaste) state.managerPaste[$("pasteGroup").value] = {};
      state.managerRows = [];
      state.managerHeaders = [];
      renderManager();
      renderMerge();
      updateMetrics();
    });
    $("buildMergeBtn")?.addEventListener("click", renderMerge);
    $("copyTimeBtn").addEventListener("click", () => copyTable("timeTable", "memoStatus"));
    $("copyTempoBtn").addEventListener("click", () => copyTable("tempoTable", "memoStatus"));
    $("copyDiveBtn").addEventListener("click", () => copyTable("diveTimeTable", "diveStatus"));
    $("copyDiveTempoBtn").addEventListener("click", () => copyTable("diveTempoTable", "diveStatus"));
    $("copyManagerBtn")?.addEventListener("click", () => copyTable("managerTable", "managerStatus"));
    $("copyMergeBtn").addEventListener("click", () => copyTable("mergeTable", "mergeStatus"));
    $("pdfMergeBtn").addEventListener("click", exportMergePdf);
    $("backupExportBtn").addEventListener("click", exportBackup);
    $("backupImportBtn").addEventListener("click", () => $("backupImportFile").click());
    $("backupImportFile").addEventListener("change", (event) => {
      importBackupFile(event.target.files?.[0]);
      event.target.value = "";
    });
    $("saveBtn")?.addEventListener("click", save);
    $("loadBtn")?.addEventListener("click", load);
    $("clearBtn").addEventListener("click", () => {
      if (!confirm("入力内容を初期化しますか？")) return;
      localStorage.removeItem(STORAGE_KEY);
      location.reload();
    });

    if ($("sessionDate")) {
      $("sessionDate").value = today();
      state.session.date = $("sessionDate").value;
    }
    if (!load()) renderAll();
  </script>
</body>
</html>
"""


@app.route("/")
def home():
    return render_template_string(PAGE)


@app.route("/vendor/<path:filename>")
def vendor_file(filename):
    return send_from_directory(os.path.join(os.path.dirname(__file__), "vendor"), filename)


@app.route("/health")
def health():
    return "swim-data-app-final-input-ok"


@app.route("/import-manager-excel", methods=["POST"])
def import_manager_excel():
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"error": "Excelファイルがありません"}), 400
    try:
        return jsonify(parse_manager_workbook(uploaded))
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


if __name__ == "__main__":
    app.run(debug=True, port=5155)
