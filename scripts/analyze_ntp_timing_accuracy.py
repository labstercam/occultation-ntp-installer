#!/usr/bin/env python3
"""Analyze NTP loopstats/peerstats logs and report timing accuracy interpretations.

This script implements interpretations A, B, C, and D from docs/ntp_traceability.md.
It is designed for Windows installs used by this repository but will run on any OS.
"""

import argparse
import csv
import json
import math
import os
import re
import statistics
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple


MJD_EPOCH = date(1858, 11, 17)
SQRT3 = math.sqrt(3.0)
UI_WIDTH = 80


class LoopRecord:
    def __init__(self, mjd: int, offset: float, freq: float, jitter: float) -> None:
        self.mjd = mjd
        self.offset = offset
        self.freq = freq
        self.jitter = jitter


class PeerRecord:
    def __init__(self, mjd: int, status: str, delay: float, dispersion: float, jitter: float) -> None:
        self.mjd = mjd
        self.status = status
        self.delay = delay
        self.dispersion = dispersion
        self.jitter = jitter


class DayOption:
    def __init__(
        self,
        key: str,
        label: str,
        loop_path: Path,
        peer_path: Path,
        target_mjd: Optional[int],
        recency_score: float,
    ) -> None:
        self.key = key
        self.label = label
        self.loop_path = loop_path
        self.peer_path = peer_path
        self.target_mjd = target_mjd
        self.recency_score = recency_score


class AnalysisResult:
    def __init__(
        self,
        option: DayOption,
        mjds: List[int],
        loop_rows_used: int,
        peer_rows_total: int,
        peer_rows_used: int,
        peer_selection_note: str,
        metrics: Dict[str, float],
    ) -> None:
        self.option = option
        self.mjds = mjds
        self.loop_rows_used = loop_rows_used
        self.peer_rows_total = peer_rows_total
        self.peer_rows_used = peer_rows_used
        self.peer_selection_note = peer_selection_note
        self.metrics = metrics


def mjd_to_date_string(mjd: int) -> str:
    return (MJD_EPOCH + timedelta(days=mjd)).isoformat()


def ui_rule(char: str = "=") -> None:
    print(char * UI_WIDTH)


def ui_header(title: str, subtitle: Optional[str] = None) -> None:
    ui_rule("=")
    print(title)
    if subtitle:
        print(subtitle)
    ui_rule("=")


def ui_section(title: str) -> None:
    print()
    ui_rule("-")


def has_interactive_stdin() -> bool:
    stream = getattr(sys, "stdin", None)
    if stream is None:
        return False
    if not hasattr(stream, "readline"):
        return False
    isatty = getattr(stream, "isatty", None)
    if callable(isatty):
        try:
            return bool(isatty())
        except Exception:
            return True
    return True


def safe_input(prompt: str) -> str:
    # In some embedded hosts, input(prompt) can throw ValueError from host-side formatting.
    try:
        return input(prompt)
    except ValueError:
        print(prompt, end="")
        try:
            return input("")
        except Exception:
            stream = getattr(sys, "stdin", None) or getattr(sys, "__stdin__", None)
            if stream is None or not hasattr(stream, "readline"):
                raise RuntimeError(
                    "Interactive input is not available in this host. "
                    "Provide --log-folder (and optionally --day/--export-dir)."
                )
            line = stream.readline()
            if line is None:
                return ""
            return line.rstrip("\r\n")
    except EOFError:
        raise RuntimeError(
            "No stdin available for interactive prompts. "
            "Provide --log-folder (and optionally --day/--export-dir)."
        )
    print(title)
    ui_rule("-")


def prompt_with_default(prompt_text: str, default: Optional[str]) -> str:
    suffix = ""
    if default:
        suffix = " [{0}]".format(default)
    return safe_input("{0}{1}: ".format(prompt_text, suffix)).strip()


def prompt_numbered_choice(prompt_text: str, item_count: int, default_index: int) -> int:
    while True:
        entry = prompt_with_default(prompt_text, str(default_index))
        if not entry:
            return default_index
        if entry.isdigit():
            picked = int(entry)
            if 1 <= picked <= item_count:
                return picked
        print("Please enter a number from 1 to {0}.".format(item_count))


def prompt_yes_no(prompt_text: str, default_yes: bool = False) -> bool:
    default_hint = "Y/n" if default_yes else "y/N"
    while True:
        entry = safe_input("{0} [{1}]: ".format(prompt_text, default_hint)).strip().lower()
        if not entry:
            return default_yes
        if entry in ("y", "yes"):
            return True
        if entry in ("n", "no"):
            return False
        print("Please answer yes or no.")


def discover_candidate_dirs() -> List[Path]:
    env = os.environ
    roots: List[Path] = []

    for key in ("ProgramData", "ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA", "APPDATA", "USERPROFILE"):
        value = env.get(key)
        if value:
            roots.append(Path(value))

    known_relatives = [
        Path("NTP") / "logs",
        Path("NTP") / "etc",
        Path("Meinberg") / "NTP" / "logs",
        Path("Meinberg") / "NTP" / "etc",
    ]

    candidates: List[Path] = []
    seen: Set[str] = set()

    def add_if_log_dir(path: Path) -> None:
        resolved = str(path.resolve()) if path.exists() else str(path)
        if resolved in seen:
            return
        if path.is_dir() and has_stats_files(path):
            seen.add(resolved)
            candidates.append(path)

    for root in roots:
        for relative in known_relatives:
            add_if_log_dir(root / relative)

    explicit = env.get("NTP_LOG_DIR")
    if explicit:
        add_if_log_dir(Path(explicit))

    return candidates


def has_stats_files(folder: Path) -> bool:
    try:
        for child in folder.iterdir():
            if not child.is_file():
                continue
            name = child.name.lower()
            if name.startswith("loopstats") or name.startswith("peerstats"):
                return True
    except OSError:
        return False
    return False


def prompt_for_folder(preferred: Optional[Path], discovered: Iterable[Path]) -> Path:
    discovered_list = list(discovered)
    ui_section("NTP Log Folder Selection")
    print("Choose a detected folder number, or type a full path.")
    print("Press Enter to accept the default.")
    print()

    if discovered_list:
        print("Detected possible NTP log folders:")
        for index, path in enumerate(discovered_list, start=1):
            print("  {0}. {1}".format(index, path))
    else:
        print("No common NTP log folders were auto-detected.")

    default = preferred if preferred else (discovered_list[0] if discovered_list else None)

    while True:
        default_text = str(default) if default else None
        value = prompt_with_default("Enter NTP log folder", default_text).strip('"')

        if not value and default:
            chosen = default
        elif value.isdigit() and discovered_list:
            pick = int(value)
            if 1 <= pick <= len(discovered_list):
                chosen = discovered_list[pick - 1]
            else:
                print("Selection out of range.")
                continue
        else:
            chosen = Path(value)

        if chosen.is_dir():
            return chosen
        print(f"Folder not found: {chosen}")


def extract_tag(name: str, prefix: str) -> str:
    if not name.lower().startswith(prefix):
        return ""

    remainder = name[len(prefix):]
    if remainder.startswith("."):
        remainder = remainder[1:]

    if not remainder:
        return ""

    match = re.search(r"(\d{8})", remainder)
    if match:
        return match.group(1)

    return remainder


def build_day_options(folder: Path) -> List[DayOption]:
    loop_by_tag: Dict[str, Path] = {}
    peer_by_tag: Dict[str, Path] = {}

    for child in folder.iterdir():
        if not child.is_file():
            continue

        lower = child.name.lower()
        if lower.startswith("loopstats"):
            loop_by_tag[extract_tag(lower, "loopstats")] = child
        elif lower.startswith("peerstats"):
            peer_by_tag[extract_tag(lower, "peerstats")] = child

    options: List[DayOption] = []
    common_tags = sorted(set(loop_by_tag).intersection(peer_by_tag))

    for tag in common_tags:
        loop_path = loop_by_tag[tag]
        peer_path = peer_by_tag[tag]

        label_tag = tag if tag else "(unsuffixed files)"
        if tag.isdigit() and len(tag) == 8:
            label_tag = f"{tag[:4]}-{tag[4:6]}-{tag[6:]}"

        score = max(loop_path.stat().st_mtime, peer_path.stat().st_mtime)
        options.append(
            DayOption(
                key=f"tag:{tag}",
                label=f"{label_tag}  [{loop_path.name}, {peer_path.name}]",
                loop_path=loop_path,
                peer_path=peer_path,
                target_mjd=None,
                recency_score=score,
            )
        )

    if "" in loop_by_tag and "" in peer_by_tag:
        loop_path = loop_by_tag[""]
        peer_path = peer_by_tag[""]
        loop_mjds = read_available_mjds(loop_path)
        peer_mjds = read_available_mjds(peer_path)
        common_mjds = sorted(loop_mjds.intersection(peer_mjds))

        for mjd in common_mjds:
            score = max(loop_path.stat().st_mtime, peer_path.stat().st_mtime) + (mjd / 1000000000.0)
            options.append(
                DayOption(
                    key=f"mjd:{mjd}",
                    label=f"{mjd_to_date_string(mjd)} (MJD {mjd})  [{loop_path.name}, {peer_path.name}]",
                    loop_path=loop_path,
                    peer_path=peer_path,
                    target_mjd=mjd,
                    recency_score=score,
                )
            )

    options.sort(key=lambda item: item.recency_score, reverse=True)
    return options


def read_available_mjds(path: Path) -> Set[int]:
    values: Set[int] = set()
    try:
        with path.open("r", encoding="utf-8", errors="ignore") as handle:
            for line in handle:
                text = line.strip()
                if not text or text.startswith("#"):
                    continue
                parts = text.split()
                if not parts:
                    continue
                try:
                    values.add(int(parts[0]))
                except ValueError:
                    continue
    except OSError:
        return set()
    return values


def select_day_option(options: List[DayOption], day_override: Optional[str]) -> DayOption:
    if not options:
        raise RuntimeError("No day options found. Ensure both loopstats and peerstats files exist in the folder.")

    if day_override:
        needle = day_override.strip().lower()
        for option in options:
            if needle in option.key.lower() or needle in option.label.lower():
                return option
        raise RuntimeError(f"Requested day '{day_override}' did not match available options.")

    if not has_interactive_stdin():
        return options[0]

    ui_section("Day Dataset Selection")
    print("Available day datasets (newest first):")
    for index, option in enumerate(options, start=1):
        print("  {0}. {1}".format(index, option.label))

    default_index = 1
    selected = prompt_numbered_choice("Select day dataset", len(options), default_index)
    return options[selected - 1]


def parse_loopstats(path: Path, target_mjd: Optional[int]) -> List[LoopRecord]:
    rows: List[LoopRecord] = []
    with path.open("r", encoding="utf-8", errors="ignore") as handle:
        for line in handle:
            text = line.strip()
            if not text or text.startswith("#"):
                continue

            parts = text.split()
            if len(parts) < 5:
                continue

            try:
                mjd = int(parts[0])
                offset = float(parts[2])
                freq = float(parts[3])
                jitter = float(parts[4])
            except ValueError:
                continue

            if target_mjd is not None and mjd != target_mjd:
                continue

            rows.append(LoopRecord(mjd=mjd, offset=offset, freq=freq, jitter=jitter))

    return rows


def parse_peerstats(path: Path, target_mjd: Optional[int]) -> List[PeerRecord]:
    rows: List[PeerRecord] = []
    with path.open("r", encoding="utf-8", errors="ignore") as handle:
        for line in handle:
            text = line.strip()
            if not text or text.startswith("#"):
                continue

            parts = text.split()
            if len(parts) < 8:
                continue

            try:
                mjd = int(parts[0])
                delay = float(parts[5])
                dispersion = float(parts[6])
                jitter = float(parts[7])
            except ValueError:
                continue

            if target_mjd is not None and mjd != target_mjd:
                continue

            rows.append(
                PeerRecord(
                    mjd=mjd,
                    status=parts[3],
                    delay=delay,
                    dispersion=dispersion,
                    jitter=jitter,
                )
            )

    return rows


def is_selected_status(status_text: str) -> bool:
    token = status_text.strip().lower()
    candidates: List[int] = []

    try:
        candidates.append(int(token, 16))
    except ValueError:
        pass

    if token.isdigit():
        try:
            candidates.append(int(token, 10))
        except ValueError:
            pass

    return any((value & 0x0700) != 0 for value in candidates)


def mean(values: List[float]) -> float:
    if not values:
        raise RuntimeError("Cannot compute mean of empty data.")
    return float(sum(values)) / float(len(values))


def stdev(values: List[float]) -> float:
    if len(values) < 2:
        return 0.0
    return statistics.stdev(values)


def format_ms(seconds: float) -> str:
    return f"{seconds * 1000.0:.6f} ms"


def format_us(seconds: float) -> str:
    return f"{seconds * 1000000.0:.3f} us"


def analyze(option: DayOption, loop_rows: List[LoopRecord], peer_rows: List[PeerRecord]) -> AnalysisResult:
    if not loop_rows:
        raise RuntimeError("No usable loopstats rows were found for the selected day.")
    if not peer_rows:
        raise RuntimeError("No usable peerstats rows were found for the selected day.")

    selected_peers = [row for row in peer_rows if is_selected_status(row.status)]
    peer_subset = selected_peers if selected_peers else peer_rows
    peer_selection_note = (
        f"Selected peers by status flags: {len(selected_peers)} of {len(peer_rows)} rows"
        if selected_peers
        else "No selected-peer status rows found; using all peerstats rows"
    )

    offsets = [row.offset for row in loop_rows]
    loop_jitter = [row.jitter for row in loop_rows]
    delays = [row.delay for row in peer_subset]
    dispersions = [row.dispersion for row in peer_subset]
    peer_jitter = [row.jitter for row in peer_subset]

    mean_offset_signed = mean(offsets)
    mean_offset_abs = mean([abs(value) for value in offsets])
    mean_delay = mean(delays)
    mean_half_delay = mean_delay / 2.0
    mean_loop_jitter = mean(loop_jitter)
    mean_dispersion = mean(dispersions)
    mean_peer_jitter = mean(peer_jitter)

    # Interpretation A
    a_uncertainty = math.sqrt(mean_offset_signed * mean_offset_signed + mean_delay * mean_delay)

    # Interpretation B
    b_u_offset = mean_offset_abs / SQRT3
    b_u_delay = mean_half_delay / SQRT3
    b_u_server = 3e-6
    b_u_combined = math.sqrt(b_u_offset * b_u_offset + b_u_delay * b_u_delay + b_u_server * b_u_server)
    b_u_expanded = 2.0 * b_u_combined

    # Interpretation C
    c_bias = mean_offset_signed
    c_u_wander = stdev(offsets)
    c_u_asymmetry = mean_half_delay / SQRT3
    c_u_delay_variation = stdev(delays) / 2.0
    c_u_server = 3e-6 / SQRT3
    c_u_combined = math.sqrt(
        c_u_wander * c_u_wander
        + c_u_asymmetry * c_u_asymmetry
        + c_u_delay_variation * c_u_delay_variation
        + c_u_server * c_u_server
    )
    c_u_expanded = 2.0 * c_u_combined

    # Interpretation D
    d_u_jitter = mean_loop_jitter
    d_u_dispersion = mean_dispersion
    d_u_asymmetry = mean_half_delay / SQRT3
    d_u_server = 3e-6 / SQRT3
    d_u_combined = math.sqrt(
        d_u_jitter * d_u_jitter
        + d_u_dispersion * d_u_dispersion
        + d_u_asymmetry * d_u_asymmetry
        + d_u_server * d_u_server
    )
    d_u_expanded = 2.0 * d_u_combined

    mjds = sorted({row.mjd for row in loop_rows}.union({row.mjd for row in peer_rows}))

    metrics = {
        "mean_offset_signed": mean_offset_signed,
        "mean_delay": mean_delay,
        "a_uncertainty": a_uncertainty,
        "b_u_offset": b_u_offset,
        "b_u_delay": b_u_delay,
        "b_u_server": b_u_server,
        "b_u_combined": b_u_combined,
        "b_u_expanded": b_u_expanded,
        "c_bias": c_bias,
        "c_u_wander": c_u_wander,
        "c_u_asymmetry": c_u_asymmetry,
        "c_u_delay_variation": c_u_delay_variation,
        "c_u_server": c_u_server,
        "c_u_combined": c_u_combined,
        "c_u_expanded": c_u_expanded,
        "d_u_jitter": d_u_jitter,
        "d_u_dispersion": d_u_dispersion,
        "d_u_peer_jitter": mean_peer_jitter,
        "d_u_asymmetry": d_u_asymmetry,
        "d_u_server": d_u_server,
        "d_u_combined": d_u_combined,
        "d_u_expanded": d_u_expanded,
    }

    return AnalysisResult(
        option=option,
        mjds=mjds,
        loop_rows_used=len(loop_rows),
        peer_rows_total=len(peer_rows),
        peer_rows_used=len(peer_subset),
        peer_selection_note=peer_selection_note,
        metrics=metrics,
    )


def generate_report(result: AnalysisResult) -> str:
    date_text = ", ".join(f"{mjd_to_date_string(mjd)} (MJD {mjd})" for mjd in result.mjds)
    option = result.option
    m = result.metrics

    lines = [
        "NTP Timing Accuracy Report",
        "=" * 80,
        f"Dataset: {option.label}",
        f"Loopstats file: {option.loop_path}",
        f"Peerstats file: {option.peer_path}",
        f"Day(s) present in selected data: {date_text}",
        f"Loopstats rows used: {result.loop_rows_used}",
        f"Peerstats rows used: {result.peer_rows_used} ({result.peer_selection_note})",
        "",
        "Interpretation A (literal reading)",
        "-" * 80,
        f"Mean signed loop offset: {format_ms(m['mean_offset_signed'])}",
        f"Mean peer delay (RTT): {format_ms(m['mean_delay'])}",
        f"Quadrature sum: {format_ms(m['a_uncertainty'])}",
        "",
        "Interpretation B (metrological)",
        "-" * 80,
        f"u_offset = mean(abs(offset))/sqrt(3): {format_ms(m['b_u_offset'])}",
        f"u_delay = (mean(delay)/2)/sqrt(3): {format_ms(m['b_u_delay'])}",
        f"u_server (fixed): {format_us(m['b_u_server'])}",
        f"u_combined (k=1): {format_ms(m['b_u_combined'])}",
        f"U_expanded (k=2): +/- {format_ms(m['b_u_expanded'])}",
        "",
        "Interpretation C (statistical)",
        "-" * 80,
        f"Systematic bias (mean signed offset): {format_ms(m['c_bias'])}",
        f"u_wander = stdev(offset): {format_ms(m['c_u_wander'])}",
        f"u_asymmetry = (mean(delay)/2)/sqrt(3): {format_ms(m['c_u_asymmetry'])}",
        f"u_delay_variation = stdev(delay)/2: {format_ms(m['c_u_delay_variation'])}",
        f"u_server (rectangular): {format_us(m['c_u_server'])}",
        f"u_combined (k=1): {format_ms(m['c_u_combined'])}",
        f"U_expanded (k=2): +/- {format_ms(m['c_u_expanded'])}",
        "",
        "Interpretation D (NTP native statistics)",
        "-" * 80,
        f"u_jitter (loopstats mean jitter): {format_ms(m['d_u_jitter'])}",
        f"u_dispersion (peerstats mean dispersion): {format_ms(m['d_u_dispersion'])}",
        f"u_peer_jitter (informational): {format_ms(m['d_u_peer_jitter'])}",
        f"u_asymmetry = (mean(delay)/2)/sqrt(3): {format_ms(m['d_u_asymmetry'])}",
        f"u_server (rectangular): {format_us(m['d_u_server'])}",
        f"u_combined (k=1): {format_ms(m['d_u_combined'])}",
        f"U_expanded (k=2): +/- {format_ms(m['d_u_expanded'])}",
    ]

    return "\n".join(lines)


def sanitize_filename_part(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_") or "dataset"


def resolve_export_paths(export_dir: Path, result: AnalysisResult) -> Tuple[Path, Path]:
    export_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if result.mjds:
        day_part = sanitize_filename_part(mjd_to_date_string(max(result.mjds)))
    else:
        day_part = "unknown_day"
    base = export_dir / f"ntp_accuracy_{day_part}_{stamp}"
    return base.with_suffix(".json"), base.with_suffix(".csv")


def export_json(path: Path, result: AnalysisResult) -> None:
    generated_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    payload = {
        "generated_at": generated_at,
        "dataset_label": result.option.label,
        "loopstats_file": str(result.option.loop_path),
        "peerstats_file": str(result.option.peer_path),
        "mjd_days": result.mjds,
        "iso_days": [mjd_to_date_string(mjd) for mjd in result.mjds],
        "loop_rows_used": result.loop_rows_used,
        "peer_rows_total": result.peer_rows_total,
        "peer_rows_used": result.peer_rows_used,
        "peer_selection_note": result.peer_selection_note,
        "metrics_seconds": result.metrics,
    }
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)


def export_csv(path: Path, result: AnalysisResult) -> None:
    generated_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    columns = [
        "generated_at",
        "dataset_label",
        "loopstats_file",
        "peerstats_file",
        "mjd_days",
        "iso_days",
        "loop_rows_used",
        "peer_rows_total",
        "peer_rows_used",
        "peer_selection_note",
    ] + sorted(result.metrics.keys())

    row = {
        "generated_at": generated_at,
        "dataset_label": result.option.label,
        "loopstats_file": str(result.option.loop_path),
        "peerstats_file": str(result.option.peer_path),
        "mjd_days": ";".join(str(mjd) for mjd in result.mjds),
        "iso_days": ";".join(mjd_to_date_string(mjd) for mjd in result.mjds),
        "loop_rows_used": str(result.loop_rows_used),
        "peer_rows_total": str(result.peer_rows_total),
        "peer_rows_used": str(result.peer_rows_used),
        "peer_selection_note": result.peer_selection_note,
    }
    for key, value in result.metrics.items():
        row[key] = f"{value:.12f}"

    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        writer.writerow(row)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyze NTP timing accuracy for one day of loopstats/peerstats logs.")
    parser.add_argument("--log-folder", dest="log_folder", help="Path to folder containing loopstats/peerstats files.")
    parser.add_argument(
        "--day",
        dest="day",
        help="Optional day selector (e.g. 20260319, MJD value, or text matching an option label).",
    )
    parser.add_argument(
        "--export-dir",
        dest="export_dir",
        help="If set, save a JSON and CSV audit record in this folder.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    ui_header("NTP Timing Accuracy Analyzer", "Interpretations A, B, C, and D")

    discovered = discover_candidate_dirs()
    preferred = Path(args.log_folder) if args.log_folder else None
    if preferred:
        if not preferred.is_dir():
            raise RuntimeError("Specified --log-folder does not exist: {0}".format(preferred))
        log_folder = preferred
    elif has_interactive_stdin():
        log_folder = prompt_for_folder(preferred, discovered)
    else:
        raise RuntimeError(
            "No interactive stdin available to ask for log folder. "
            "Please run with --log-folder <path>."
        )

    options = build_day_options(log_folder)
    if not options:
        raise RuntimeError(
            f"No matching loopstats/peerstats datasets found in {log_folder}. "
            "Expected files such as loopstats.20260319 and peerstats.20260319 or unsuffixed loopstats/peerstats files."
        )

    selected = select_day_option(options, args.day)
    loop_rows = parse_loopstats(selected.loop_path, selected.target_mjd)
    peer_rows = parse_peerstats(selected.peer_path, selected.target_mjd)
    result = analyze(selected, loop_rows, peer_rows)

    print()
    print(generate_report(result))

    export_dir = args.export_dir
    if not export_dir and has_interactive_stdin() and prompt_yes_no("Save JSON and CSV audit files?", default_yes=False):
        default_export = str(log_folder / "reports")
        entered = prompt_with_default("Export folder", default_export).strip().strip('"')
        export_dir = entered if entered else default_export

    if export_dir:
        export_root = Path(export_dir)
        json_path, csv_path = resolve_export_paths(export_root, result)
        export_json(json_path, result)
        export_csv(csv_path, result)
        print()
        ui_section("Export Complete")
        print(f"Saved JSON: {json_path}")
        print(f"Saved CSV:  {csv_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled by user.")
        raise SystemExit(130)
    except RuntimeError as error:
        print(f"Error: {error}")
        raise SystemExit(1)