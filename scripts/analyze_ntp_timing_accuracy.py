#!/usr/bin/env ipy
"""IronPython 3.4 Windows Forms tool to analyze NTP loopstats/peerstats timing accuracy.

This implements interpretations A, B, C, and D from docs/ntp_traceability.md.
"""

import csv
import ipaddress
import json
import math
import os
import re
import socket
import sys
import urllib.request
from datetime import date, datetime, timedelta

try:
    import clr  # type: ignore
except ImportError:
    clr = None

if clr is not None:
    clr.AddReference("System")
    clr.AddReference("System.Drawing")
    clr.AddReference("System.Windows.Forms")

    drawing_module = __import__(
        "System.Drawing",
        fromlist=["Color", "Font", "FontStyle", "Pen", "Point", "Rectangle", "Size", "SolidBrush"],
    )
    forms_module = __import__(
        "System.Windows.Forms",
        fromlist=[
            "AnchorStyles",
            "Application",
            "Button",
            "BorderStyle",
            "CheckBox",
            "ColumnStyle",
            "ComboBox",
            "ComboBoxStyle",
            "DialogResult",
            "DockStyle",
            "FixedPanel",
            "FolderBrowserDialog",
            "Form",
            "FormWindowState",
            "FormStartPosition",
            "Label",
            "MessageBox",
            "MessageBoxButtons",
            "MessageBoxIcon",
            "PictureBox",
            "RowStyle",
            "ScrollBars",
            "SizeType",
            "SplitContainer",
            "TableLayoutPanel",
            "TextBox",
        ],
    )

    Color = drawing_module.Color
    Font = drawing_module.Font
    FontStyle = drawing_module.FontStyle
    Pen = drawing_module.Pen
    Point = drawing_module.Point
    Rectangle = drawing_module.Rectangle
    Size = drawing_module.Size
    SolidBrush = drawing_module.SolidBrush

    AnchorStyles = forms_module.AnchorStyles
    Application = forms_module.Application
    Button = forms_module.Button
    BorderStyle = forms_module.BorderStyle
    CheckBox = forms_module.CheckBox
    ComboBox = forms_module.ComboBox
    ComboBoxStyle = forms_module.ComboBoxStyle
    DialogResult = forms_module.DialogResult
    FolderBrowserDialog = forms_module.FolderBrowserDialog
    Form = forms_module.Form
    FormWindowState = forms_module.FormWindowState
    FormStartPosition = forms_module.FormStartPosition
    Label = forms_module.Label
    MessageBox = forms_module.MessageBox
    MessageBoxButtons = forms_module.MessageBoxButtons
    MessageBoxIcon = forms_module.MessageBoxIcon
    PictureBox = forms_module.PictureBox
    RowStyle = forms_module.RowStyle
    ScrollBars = forms_module.ScrollBars
    SizeType = forms_module.SizeType
    SplitContainer = forms_module.SplitContainer
    TableLayoutPanel = forms_module.TableLayoutPanel
    TextBox = forms_module.TextBox
    ColumnStyle = forms_module.ColumnStyle
    DockStyle = forms_module.DockStyle
    FixedPanel = forms_module.FixedPanel
else:
    Color = None
    Font = None
    FontStyle = None
    Pen = None
    Point = None
    Rectangle = None
    Size = None
    SolidBrush = None
    ColumnStyle = None
    DockStyle = None
    FixedPanel = None
    RowStyle = None
    SizeType = None
    SplitContainer = None
    TableLayoutPanel = None
    AnchorStyles = None
    Application = None
    Button = None
    BorderStyle = None
    CheckBox = None
    ComboBox = None
    ComboBoxStyle = None
    DialogResult = None
    FolderBrowserDialog = None
    Form = object
    FormWindowState = None
    FormStartPosition = None
    Label = None
    MessageBox = None
    MessageBoxButtons = None
    MessageBoxIcon = None
    PictureBox = None
    ScrollBars = None
    TextBox = None


MJD_EPOCH = date(1858, 11, 17)
SQRT3 = math.sqrt(3.0)

# Speed of light in single-mode fibre (km/s): c / refractive_index (n ≈ 1.467).
# Used to compute the minimum one-way propagation delay from geographic distance.
_FIBRE_SPEED_KMS = 299792.458 / 1.467   # ≈ 204,354 km/s

# Color palette for server addresses (cycling through distinct colors)
SERVER_COLORS = [
    Color.FromArgb(31, 119, 180),    # blue
    Color.FromArgb(255, 127, 14),    # orange
    Color.FromArgb(44, 160, 44),     # green
    Color.FromArgb(214, 39, 40),     # red
    Color.FromArgb(148, 103, 189),   # purple
    Color.FromArgb(140, 86, 75),     # brown
    Color.FromArgb(227, 119, 194),   # pink
    Color.FromArgb(127, 127, 127),   # gray
]

def get_server_color(server_address, server_to_color):
    """Get or assign a color for a server address."""
    if server_address not in server_to_color:
        idx = len(server_to_color) % len(SERVER_COLORS)
        server_to_color[server_address] = SERVER_COLORS[idx]
    return server_to_color[server_address]


class LoopRecord(object):
    def __init__(self, mjd, sec_of_day, offset, freq, jitter):
        self.mjd = mjd
        self.sec_of_day = sec_of_day
        self.offset = offset
        self.freq = freq
        self.jitter = jitter


class PeerRecord(object):
    def __init__(self, mjd, sec_of_day, status, delay, dispersion, jitter, server_address="", offset=0.0):
        self.mjd = mjd
        self.sec_of_day = sec_of_day
        self.status = status
        self.delay = delay
        self.dispersion = dispersion
        self.jitter = jitter
        self.server_address = server_address
        self.offset = offset


class DayOption(object):
    def __init__(self, key, label, loop_path, peer_path, target_mjd, recency_score):
        self.key = key
        self.label = label
        self.loop_path = loop_path
        self.peer_path = peer_path
        self.target_mjd = target_mjd
        self.recency_score = recency_score


class AnalysisResult(object):
    def __init__(
        self,
        option,
        mjds,
        loop_rows_used,
        peer_rows_total,
        peer_rows_used,
        peer_selection_note,
        metrics,
        diagnostics,
        pit_result=None,
    ):
        self.option = option
        self.mjds = mjds
        self.loop_rows_used = loop_rows_used
        self.peer_rows_total = peer_rows_total
        self.peer_rows_used = peer_rows_used
        self.peer_selection_note = peer_selection_note
        self.metrics = metrics
        self.diagnostics = diagnostics
        self.pit_result = pit_result


def mjd_to_date_string(mjd):
    return (MJD_EPOCH + timedelta(days=mjd)).isoformat()


def has_stats_files(folder):
    try:
        children = os.listdir(folder)
    except OSError:
        return False

    for child in children:
        child_path = os.path.join(folder, child)
        if not os.path.isfile(child_path):
            continue
        lower = child.lower()
        if lower.startswith("loopstats") or lower.startswith("peerstats"):
            return True
    return False


def discover_candidate_dirs():
    env = os.environ
    roots = []
    for key in (
        "PROGRAMDATA",
        "ProgramData",
        "ProgramFiles",
        "ProgramFiles(x86)",
        "LOCALAPPDATA",
        "APPDATA",
        "USERPROFILE",
    ):
        value = env.get(key)
        if value:
            roots.append(value)

    known_relatives = [
        os.path.join("NTP", "logs"),
        os.path.join("NTP", "etc"),
        os.path.join("Meinberg", "NTP", "logs"),
        os.path.join("Meinberg", "NTP", "etc"),
    ]

    candidates = []
    seen = set()

    def add_if_log_dir(path):
        normalized = os.path.normcase(os.path.abspath(path))
        if normalized in seen:
            return
        if os.path.isdir(path) and has_stats_files(path):
            seen.add(normalized)
            candidates.append(path)

    for root in roots:
        for relative in known_relatives:
            add_if_log_dir(os.path.join(root, relative))

    explicit = env.get("NTP_LOG_DIR")
    if explicit:
        add_if_log_dir(explicit)

    return candidates


def get_settings_file_path():
    base = (
        os.environ.get("APPDATA")
        or os.environ.get("LOCALAPPDATA")
        or os.environ.get("USERPROFILE")
        or os.path.expanduser("~")
    )
    if not base:
        return os.path.join(os.getcwd(), "ntp_analyzer_settings.json")
    settings_dir = os.path.join(base, "occultation-ntp-installer")
    return os.path.join(settings_dir, "ntp_analyzer_settings.json")


def load_folder_settings():
    path = get_settings_file_path()
    try:
        handle = open(path, "r", encoding="utf-8")
    except IOError:
        return {}

    try:
        data = json.load(handle)
    except Exception:
        return {}
    finally:
        handle.close()

    if not isinstance(data, dict):
        return {}
    return data


def save_folder_settings(log_folder, export_folder, observer_lat="", observer_lon=""):
    path = get_settings_file_path()
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)

    payload = {
        "log_folder": log_folder,
        "export_folder": export_folder,
        "observer_lat": observer_lat,
        "observer_lon": observer_lon,
    }
    handle = open(path, "w", encoding="utf-8")
    try:
        json.dump(payload, handle, indent=2, sort_keys=True)
    finally:
        handle.close()


def extract_tag(name, prefix):
    if not name.lower().startswith(prefix):
        return ""

    remainder = name[len(prefix) :]
    if remainder.startswith("."):
        remainder = remainder[1:]

    if not remainder:
        return ""

    match = re.search(r"(\d{8})", remainder)
    if match:
        return match.group(1)

    return remainder


def read_available_mjds(path):
    values = set()
    try:
        handle = open(path, "r")
    except IOError:
        return values

    try:
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
    finally:
        handle.close()

    return values


def build_day_options(folder):
    loop_by_tag = {}
    peer_by_tag = {}

    for child in os.listdir(folder):
        child_path = os.path.join(folder, child)
        if not os.path.isfile(child_path):
            continue

        lower = child.lower()
        if lower.startswith("loopstats"):
            loop_by_tag[extract_tag(lower, "loopstats")] = child_path
        elif lower.startswith("peerstats"):
            peer_by_tag[extract_tag(lower, "peerstats")] = child_path

    options = []
    common_tags = sorted(set(loop_by_tag.keys()).intersection(set(peer_by_tag.keys())))

    for tag in common_tags:
        loop_path = loop_by_tag[tag]
        peer_path = peer_by_tag[tag]

        label_tag = tag if tag else "(unsuffixed files)"
        if tag.isdigit() and len(tag) == 8:
            label_tag = "%s-%s-%s" % (tag[:4], tag[4:6], tag[6:])

        score = max(os.path.getmtime(loop_path), os.path.getmtime(peer_path))
        options.append(
            DayOption(
                key="tag:%s" % tag,
                label="%s  [%s, %s]" % (label_tag, os.path.basename(loop_path), os.path.basename(peer_path)),
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

        base_score = max(os.path.getmtime(loop_path), os.path.getmtime(peer_path))
        for mjd in common_mjds:
            score = base_score + (mjd / 1000000000.0)
            options.append(
                DayOption(
                    key="mjd:%d" % mjd,
                    label="%s (MJD %d)  [%s, %s]"
                    % (mjd_to_date_string(mjd), mjd, os.path.basename(loop_path), os.path.basename(peer_path)),
                    loop_path=loop_path,
                    peer_path=peer_path,
                    target_mjd=mjd,
                    recency_score=score,
                )
            )

    options.sort(key=lambda item: item.recency_score, reverse=True)
    return options


def parse_loopstats(path, target_mjd):
    rows = []
    handle = open(path, "r")
    try:
        for line in handle:
            text = line.strip()
            if not text or text.startswith("#"):
                continue

            parts = text.split()
            if len(parts) < 5:
                continue

            try:
                mjd = int(parts[0])
                sec_of_day = float(parts[1])
                offset = float(parts[2])
                freq = float(parts[3])
                jitter = float(parts[4])
            except ValueError:
                continue

            if target_mjd is not None and mjd != target_mjd:
                continue

            rows.append(LoopRecord(mjd=mjd, sec_of_day=sec_of_day, offset=offset, freq=freq, jitter=jitter))
    finally:
        handle.close()

    return rows


def parse_peerstats(path, target_mjd):
    rows = []
    handle = open(path, "r")
    try:
        for line in handle:
            text = line.strip()
            if not text or text.startswith("#"):
                continue

            parts = text.split()
            if len(parts) < 8:
                continue

            try:
                mjd = int(parts[0])
                sec_of_day = float(parts[1])
                server_address = parts[2]
                peer_offset = float(parts[4])
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
                    sec_of_day=sec_of_day,
                    status=parts[3],
                    delay=delay,
                    dispersion=dispersion,
                    jitter=jitter,
                    server_address=server_address,
                    offset=peer_offset,
                )
            )
    finally:
        handle.close()

    return rows


def parse_status_word(status_text):
    """Parse a peer status word.

    NTP peerstats status is represented as a hex status word in this project
    (for example: "9614"). Treating plain digits as decimal introduces
    false classifications, so parse as hexadecimal by default.
    """
    token = status_text.strip().lower()
    if not token:
        return None

    try:
        if token.startswith("0x"):
            return int(token, 16)
        return int(token, 16)
    except ValueError:
        return None


def is_selected_status(status_text):
    """Return True only for peers currently selected for synchronization.

    Selection code is bits 8..10 of the peer status word:
    - 6: sys.peer
    - 7: pps.peer
    """
    value = parse_status_word(status_text)
    if value is None:
        return False
    select_code = (value & 0x0700) >> 8
    return select_code in (6, 7)


def is_candidate_or_better(status_text):
    """Return True for peers that passed NTP's sanity checks (select code >= 4).

    This includes candidate (4), backup (5), sys.peer (6) and pps.peer (7).
    Codes 0-3 (reject, falseticker, excess, outlier) are excluded.
    These peers carry valid independent offset measurements in peerstats.
    """
    value = parse_status_word(status_text)
    if value is None:
        return False
    select_code = (value & 0x0700) >> 8
    return select_code >= 4


def get_select_code(status_text):
    value = parse_status_word(status_text)
    if value is None:
        return -1
    return (value & 0x0700) >> 8


def select_peer_subset(peer_rows):
    selected_peers = [row for row in peer_rows if is_selected_status(row.status)]
    if selected_peers:
        note = "Selected peers by status flags: %d of %d rows" % (len(selected_peers), len(peer_rows))
        return selected_peers, note
    note = "No selected-peer status rows found; using all peerstats rows"
    return peer_rows, note


def reduce_to_active_timeline(peer_rows):
    """Collapse rows to one active selected peer per timestamp for charting.

    If multiple selected rows exist at the same rounded second, prefer
    pps.peer (7) over sys.peer (6), then lower delay as a stable tie-breaker.
    """
    by_second = {}
    for row in peer_rows:
        key = (row.mjd, int(round(row.sec_of_day)))
        code = get_select_code(row.status)
        priority = 1 if code == 7 else 0
        delay_value = row.delay
        sort_key = (priority, -delay_value)

        current = by_second.get(key)
        if current is None:
            by_second[key] = (sort_key, row)
        else:
            if sort_key > current[0]:
                by_second[key] = (sort_key, row)

    chosen = [entry[1] for key, entry in sorted(by_second.items(), key=lambda item: item[0])]
    return chosen


def compute_peer_diagnostics(peer_subset):
    code_counts = {}
    for row in peer_subset:
        code = get_select_code(row.status)
        code_counts[code] = code_counts.get(code, 0) + 1

    active_rows = reduce_to_active_timeline(peer_subset)
    transitions = 0
    prev_server = None
    for row in active_rows:
        server = row.server_address if hasattr(row, "server_address") else ""
        if prev_server is not None and server != prev_server:
            transitions += 1
        prev_server = server

    return {
        "select_code_counts": code_counts,
        "active_rows": len(active_rows),
        "active_transitions": transitions,
        "active_unique_servers": len(set([(r.server_address if hasattr(r, "server_address") else "") for r in active_rows])),
    }


def to_utc_datetime(mjd, sec_of_day):
    day = MJD_EPOCH + timedelta(days=mjd)
    day_start = datetime(day.year, day.month, day.day)
    return day_start + timedelta(seconds=float(sec_of_day))


def aggregate_peer_timeseries(peer_rows):
    """Aggregate peer measurements by second and server.
    
    Returns a list of aggregated records grouped by (mjd, second, server).
    This allows per-server visualization when using aggregated data.
    """
    by_second_server = {}
    for row in peer_rows:
        second_key = int(round(row.sec_of_day))
        server = row.server_address if hasattr(row, 'server_address') else ""
        key = (row.mjd, second_key, server)
        if key not in by_second_server:
            by_second_server[key] = [0.0, 0.0, 0.0, 0]
        by_second_server[key][0] += row.jitter
        by_second_server[key][1] += row.dispersion
        by_second_server[key][2] += row.delay
        by_second_server[key][3] += 1

    points = []
    for key in sorted(by_second_server.keys()):
        stats = by_second_server[key]
        count = stats[3]
        points.append(
            {
                "mjd": key[0],
                "sec_of_day": key[1],
                "server_address": key[2],
                "jitter": stats[0] / float(count),
                "dispersion": stats[1] / float(count),
                "delay": stats[2] / float(count),
            }
        )
    return points


def compute_axis_day_bounds(loop_rows, peer_rows):
    if not loop_rows and not peer_rows:
        return None, None

    all_mjds = [row.mjd for row in loop_rows] + [row.mjd for row in peer_rows]
    min_mjd = min(all_mjds)
    max_mjd = max(all_mjds)

    start_day = MJD_EPOCH + timedelta(days=min_mjd)
    end_day = MJD_EPOCH + timedelta(days=max_mjd + 1)
    start = datetime(start_day.year, start_day.month, start_day.day, 0, 0, 0)
    end = datetime(end_day.year, end_day.month, end_day.day, 0, 0, 0)
    return start, end


def mean(values):
    if not values:
        raise RuntimeError("Cannot compute mean of empty data.")
    return float(sum(values)) / float(len(values))


def stdev(values):
    if len(values) < 2:
        return 0.0
    avg = mean(values)
    variance = sum((value - avg) * (value - avg) for value in values) / float(len(values) - 1)
    return math.sqrt(variance)


def format_ms(seconds):
    return "%.6f ms" % (seconds * 1000.0)


def format_us(seconds):
    return "%.3f us" % (seconds * 1000000.0)


def _choose_y_step_ms(span_ms):
    """Return a nice round Y-axis tick interval (in ms) for the given data span."""
    if span_ms <= 0:
        return 1.0
    # Candidate nice intervals in ms
    nice = [
        0.001, 0.002, 0.005,
        0.01, 0.02, 0.05,
        0.1, 0.2, 0.5,
        1.0, 2.0, 5.0,
        10.0, 20.0, 50.0,
        100.0, 200.0, 500.0, 1000.0,
    ]
    for step in nice:
        if span_ms / step <= 7:
            return step
    return nice[-1]


def _format_y_label_ms(value_ms, step):
    """Format a Y-axis tick value (ms) with precision matched to the tick step."""
    if step >= 100.0:
        return "%d ms" % int(round(value_ms))
    elif step >= 10.0:
        return "%d ms" % int(round(value_ms))
    elif step >= 1.0:
        return "%d ms" % int(round(value_ms))
    elif step >= 0.1:
        return "%.1f ms" % value_ms
    elif step >= 0.01:
        return "%.2f ms" % value_ms
    elif step >= 0.001:
        return "%.3f ms" % value_ms
    else:
        return "%.4f ms" % value_ms


def _parse_pit_time_sec(hms_str):
    """Parse an HH:MM:SS string to seconds past midnight UTC. Raises ValueError on bad input."""
    parts = hms_str.strip().split(":")
    if len(parts) != 3:
        raise ValueError("Time must be in HH:MM:SS format (e.g. 14:32:05).")
    try:
        h, m, s = int(parts[0]), int(parts[1]), float(parts[2])
    except ValueError:
        raise ValueError("Time must be in HH:MM:SS format (e.g. 14:32:05).")
    if not (0 <= h < 24 and 0 <= m < 60 and 0.0 <= s < 60.0):
        raise ValueError("Time values out of range (H:0-23, M:0-59, S:0-59).")
    return h * 3600.0 + m * 60.0 + s


def format_pit_section(pit):
    """Format the point-in-time offset accuracy estimate as a list of report lines."""
    query_dt = (MJD_EPOCH + timedelta(days=pit["query_mjd"])).isoformat()
    query_hms = str(timedelta(seconds=int(pit["query_sec"])))
    if pit["gap_after_s"] is not None:
        gap_after_text = "%.1f s  (offset %s)" % (pit["gap_after_s"], format_ms(pit["offset_after"]))
    else:
        gap_after_text = "N/A (end of data)"
    interpolated = pit["gap_after_s"] is not None
    method_text = "Linear interpolation between surrounding loopstats records." if interpolated else "Last loopstats offset used (no later record available)."
    server_text = pit["active_server_at_T"] if pit["active_server_at_T"] else "unknown"
    lines = [
        "Point-in-Time Offset Accuracy Estimate",
        "=" * 80,
        "Query time: %s %s UTC" % (query_dt, query_hms),
        "Active reference server:            %s" % server_text,
        "",
        "Method: %s" % method_text,
        "  Asymmetry (rectangular 95%) and statistical terms (Gaussian k=2) combined in quadrature.",
        "",
        "Best-estimate offset at T:          %s" % format_ms(pit["best_offset"]),
        "  (loopstats freq at T:              %.6f ppm)" % pit["freq_ppm"],
        "  (last loopstats sync:  -%s  offset %s)" % ("%.1f s" % pit["gap_before_s"], format_ms(pit["offset_before"])),
        "  (next loopstats sync:  +%s)" % gap_after_text,
        "",
        "Uncertainty components:",
        "  u_drift (residual between sync records): %s" % format_ms(pit["u_drift"]),
        "  u_asymmetry (b_asym/sqrt(3)):            %s" % format_ms(pit["u_asymmetry"]),
        "    (mean RTT near T: %s, N=%d peer records)" % (format_ms(pit["mean_delay_near_T"]), pit["n_peer_near_T"]),
        "    b_asym (RTT/2): %s%s" % (
            format_ms(pit["b_asym_raw"]),
            ("  →  tightened to %s  [%s]" % (format_ms(pit["b_asym"]), pit["server_location_note"]))
            if pit.get("b_asym") is not None and pit["b_asym"] < pit["b_asym_raw"]
            else ("  [%s]" % pit["server_location_note"]) if pit.get("server_location_note") else "",
        ),
        "  u_scatter (stdev of loopstats offsets near T): %s" % format_ms(pit["u_scatter"]),
        "",
        "  u_combined (k=1):  %s" % format_ms(pit["u_combined"]),
        "  U_expanded (~95%%): +/- %s  [sqrt((0.95*RTT/2)^2 + (2*u_stat)^2)]" % format_ms(pit["u_expanded"]),
        "",
        "Corrected UTC time = PC time - (%s)" % format_ms(pit["best_offset"]),
        "Accuracy of corrected time: +/- %s (~95%%)" % format_ms(pit["u_expanded"]),
    ]

    # --- Alternative candidate-peer estimate ---
    lines.append("")
    lines.append("Alternative estimate from best candidate peer (min sqrt(u_asym^2 + u_jitter^2)):")
    n_cand = pit.get("n_candidate_peers_near_T", 0)
    alt_exp = pit.get("alt_u_expanded")
    if n_cand == 0:
        lines.append("  No candidate-or-better peers (select code >= 4) found near T.")
    elif alt_exp is None:
        lines.append("  No alternative candidate peers found near T (sys.peer is the only candidate).")
    else:
        is_better = alt_exp < pit["u_expanded"]
        verdict = ("IMPROVEMENT vs sys.peer estimate" if is_better
                   else "no improvement (sys.peer already has lowest uncertainty)")
        lines += [
            "  Server: %-40s  RTT: %s  (record gap from T: %.1f s)" % (
                pit["alt_server"], format_ms(pit["alt_delay"]), pit["alt_gap_s"]),
            "  Offset at poll near T (peerstats):   %s" % format_ms(pit["alt_best_offset"]),
            "  u_asymmetry (alt RTT/2/sqrt(3)):     %s" % format_ms(pit["alt_u_asymmetry"]),
            "  u_scatter   (alt peer jitter):       %s" % format_ms(pit["alt_u_scatter"]),
            "  U_expanded  (~95%%):  +/- %s  -- %s" % (format_ms(alt_exp), verdict),
        ]
        if is_better:
            lines.append(
                "  *** Lower-uncertainty estimate:  offset = %s  +/- %s (~95%%) ***" % (
                    format_ms(pit["alt_best_offset"]), format_ms(alt_exp))
            )
    return lines


def _haversine_km(lat1, lon1, lat2, lon2):
    """Return the great-circle distance in km between two WGS-84 lat/lon points (degrees)."""
    R = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = (math.sin(dlat / 2.0) ** 2
         + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
         * math.sin(dlon / 2.0) ** 2)
    return R * 2.0 * math.asin(math.sqrt(max(0.0, min(1.0, a))))


def load_known_servers(json_path):
    """Load a national_utc_ntp_servers.json file and return a flat list of server-location dicts.

    Each dict has keys:
        hostnames       list[str]   canonical NTP hostnames for this authority
        ips             list[str]   pre-resolved IPs (empty until populated at runtime)
        lat             float       latitude of server site (degrees)
        lon             float       longitude of server site (degrees)
        location_note   str         human-readable description, e.g. "PTB, Braunschweig (curated)"

    Entries without a "server_location" key are silently skipped.
    Entries with no "ntp_servers" hostnames are also skipped.
    """
    try:
        handle = open(json_path, "r", encoding="utf-8")
        try:
            data = json.load(handle)
        finally:
            handle.close()
    except Exception:
        return []

    result = []
    for entry in data.get("entries", []):
        loc = entry.get("server_location")
        if not loc:
            continue
        lat = loc.get("lat")
        lon = loc.get("lon")
        if lat is None or lon is None:
            continue
        hostnames = [h.lower() for h in entry.get("ntp_servers", []) if h]
        if not hostnames:
            continue
        authority = entry.get("authority", "")
        city = loc.get("city", "")
        note = "%s%s (curated)" % (
            ("%s, " % authority) if authority else "",
            city if city else "known server",
        )
        result.append({
            "hostnames": hostnames,
            "ips": [],
            "lat": float(lat),
            "lon": float(lon),
            "location_note": note,
        })
    return result


# ---------------------------------------------------------------------------
# Persistent IP-location cache  (resources/ip_location_cache.json)
# ---------------------------------------------------------------------------
# Entries survive across runs and are refreshed after _IP_CACHE_MAX_DAYS days.
# Structure:  { "ip_str": { lat, lon, location_note, source, cached_at } }
# Only the server coordinates are stored; geo_km is observer-dependent and
# recalculated each time from the cached lat/lon.
# ---------------------------------------------------------------------------
_ip_location_cache = {}
_ip_cache_file_path = None
_IP_CACHE_MAX_DAYS = 90


def _load_ip_location_cache(path):
    """Read ip_location_cache.json into memory.  Silent on any error."""
    global _ip_location_cache, _ip_cache_file_path
    _ip_cache_file_path = path
    try:
        handle = open(path, "r", encoding="utf-8")
        try:
            _ip_location_cache = json.load(handle)
        finally:
            handle.close()
    except Exception:
        _ip_location_cache = {}


def _save_ip_location_cache():
    """Write the in-memory cache to disk.  Silent on any error."""
    if not _ip_cache_file_path:
        return
    try:
        handle = open(_ip_cache_file_path, "w", encoding="utf-8")
        try:
            json.dump(_ip_location_cache, handle, indent=2)
        finally:
            handle.close()
    except Exception:
        pass


def _cache_ip_location(ip, lat, lon, location_note, source):
    """Store a resolved IP location in memory and flush to disk."""
    _ip_location_cache[ip] = {
        "lat": lat,
        "lon": lon,
        "location_note": location_note,
        "source": source,
        "cached_at": datetime.utcnow().strftime("%Y-%m-%d"),
    }
    _save_ip_location_cache()


# In-process GeoIP cache so a single analysis run doesn't re-query the same IP.
_geoip_cache = {}


def geolocate_ip(ip):
    """Look up the geographic coordinates of an IP address using ip-api.com.

    Results are cached in-process.  Returns a dict with keys lat, lon, city,
    country, or None on failure (network error, rate-limit, private IP, etc.).
    Timeout is 5 seconds; failure is always silent.
    """
    if ip in _geoip_cache:
        return _geoip_cache[ip]
    result = None
    try:
        url = "http://ip-api.com/json/%s?fields=status,lat,lon,city,country" % ip
        response = urllib.request.urlopen(url, timeout=5)
        data = json.loads(response.read().decode("utf-8"))
        if data.get("status") == "success":
            result = {
                "lat": float(data["lat"]),
                "lon": float(data["lon"]),
                "city": data.get("city", ""),
                "country": data.get("country", ""),
            }
    except Exception:
        pass
    _geoip_cache[ip] = result
    return result


def resolve_server_location(server_ip, known_servers, observer_lat, observer_lon):
    """Attempt to determine d_min (seconds) for an NTP server based on its IP address.

    Lookup order:
      1. Private / loopback / link-local IP  → d_min = 0 (local network)
      2. reverse-DNS → match known_servers by hostname suffix
      3. Direct IP match against known_servers pre-resolved IPs
      4. No match → d_min = None (caller uses conservative RTT/2)

    Parameters
    ----------
    server_ip : str
        The IP address (or hostname) from peerstats column 3.
    known_servers : list of dict or None
        Output of load_known_servers().  None disables lookup.
    observer_lat : float or None
        Observer latitude in degrees.
    observer_lon : float or None
        Observer longitude in degrees.

    Returns
    -------
    dict with keys:
        d_min_s        float or None   minimum one-way propagation delay (seconds)
        location_note  str             human-readable description of what matched
        geo_km         float or None   geographic distance used (km)
    """
    _no_match = {"d_min_s": None, "location_note": "server location unknown", "geo_km": None}

    if not server_ip:
        return _no_match

    # Step 1: private / loopback / link-local → local, no propagation uncertainty worth modelling.
    try:
        addr = ipaddress.ip_address(server_ip)
        if addr.is_private or addr.is_loopback or addr.is_link_local:
            return {"d_min_s": 0.0, "location_note": "local network (private IP)", "geo_km": 0.0}
    except ValueError:
        # Not a valid IP literal — treat as a hostname; fall through.
        pass

    # Step 2: persistent IP location cache (valid for _IP_CACHE_MAX_DAYS days).
    cached = _ip_location_cache.get(server_ip)
    if cached:
        try:
            cached_date = datetime.strptime(cached["cached_at"], "%Y-%m-%d")
            age_days = (datetime.utcnow() - cached_date).days
        except Exception:
            age_days = _IP_CACHE_MAX_DAYS + 1  # treat corrupt date as expired
        if age_days <= _IP_CACHE_MAX_DAYS:
            if observer_lat is not None and observer_lon is not None:
                geo_km = _haversine_km(observer_lat, observer_lon, cached["lat"], cached["lon"])
                d_min_s = geo_km / _FIBRE_SPEED_KMS
                note = "%s, %.0f km" % (cached["location_note"], geo_km)
            else:
                geo_km = None
                d_min_s = None
                note = cached["location_note"]
            return {"d_min_s": d_min_s, "location_note": note, "geo_km": geo_km}

    if observer_lat is None or observer_lon is None:
        return _no_match

    # Step 3: reverse DNS then match against known_servers by hostname/domain.
    rdns_hostname = None
    try:
        rdns_hostname = socket.gethostbyaddr(server_ip)[0].lower()
    except Exception:
        pass

    def _registered_domain(hostname):
        """Return the registered domain (e.g. 'nist.gov') from a hostname string."""
        parts = hostname.rstrip(".").split(".")
        return ".".join(parts[-2:]) if len(parts) >= 2 else hostname

    rdns_domain = _registered_domain(rdns_hostname) if rdns_hostname else None

    # Registered-domain comparison handles sibling hostnames returned by rDNS
    # (e.g. time-a-g.nist.gov reverse-resolves for an entry listing time.nist.gov).
    for entry in (known_servers or []):
        matched = False
        if rdns_hostname:
            for h in entry["hostnames"]:
                if rdns_hostname == h or rdns_hostname.endswith("." + h) or h.endswith("." + rdns_hostname):
                    matched = True
                    break
                if rdns_domain and _registered_domain(h) == rdns_domain:
                    matched = True
                    break
        if matched:
            geo_km = _haversine_km(observer_lat, observer_lon, entry["lat"], entry["lon"])
            d_min_s = geo_km / _FIBRE_SPEED_KMS
            note = entry["location_note"]
            _cache_ip_location(server_ip, entry["lat"], entry["lon"], note, "curated")
            return {
                "d_min_s": d_min_s,
                "location_note": "%s, %.0f km" % (note, geo_km),
                "geo_km": geo_km,
            }

    # Step 4: GeoIP fallback via ip-api.com (in-process cached, silent on failure).
    geo = geolocate_ip(server_ip)
    if geo is not None:
        geo_km = _haversine_km(observer_lat, observer_lon, geo["lat"], geo["lon"])
        d_min_s = geo_km / _FIBRE_SPEED_KMS
        parts = [p for p in [geo.get("city", ""), geo.get("country", "")] if p]
        note = "%s (GeoIP)" % (", ".join(parts) if parts else server_ip)
        _cache_ip_location(server_ip, geo["lat"], geo["lon"], note, "geoip")
        return {"d_min_s": d_min_s, "location_note": "%s, %.0f km" % (note, geo_km), "geo_km": geo_km}

    return _no_match


def estimate_offset_at_time(query_mjd, query_sec, loop_rows, peer_rows, window_seconds=3600.0,
                            known_servers=None, observer_lat=None, observer_lon=None):
    """Estimate the NTP offset and its uncertainty at a specific point in time.

    Parameters
    ----------
    query_mjd : int
        Modified Julian Day of the query time.
    query_sec : float
        Seconds past midnight (UTC) of the query time.
    loop_rows : list of LoopRecord
        All loopstats records available (need not be filtered to one day).
    peer_rows : list of PeerRecord
        All peerstats records available.
    window_seconds : float
        Half-width of the time window used to gather nearby peer samples for
        the network uncertainty estimate (default ±1 hour).
    known_servers : list of dict or None
        Output of load_known_servers().  When provided together with
        observer_lat/lon, enables geographic tightening of b_asym via
        resolve_server_location().  Defaults to None (conservative RTT/2 used).
    observer_lat : float or None
        Observer latitude in decimal degrees.  Required for geographic tightening.
    observer_lon : float or None
        Observer longitude in decimal degrees.  Required for geographic tightening.

    Returns
    -------
    dict with keys:
        query_mjd, query_sec          -- echo of inputs
        best_offset                    -- estimated offset at T (seconds)
        u_drift                        -- uncertainty from clock walk since last sync (s)
        u_asymmetry                    -- network path asymmetry uncertainty (s)
        u_scatter                      -- NTP measurement scatter near T (s)
        u_combined                     -- combined standard uncertainty k=1 (s)
        u_expanded                     -- expanded uncertainty ~95% (s)
        gap_before_s                   -- seconds between T and the last loopstats record
        gap_after_s                    -- seconds between T and the next loopstats record (or None)
        freq_ppm                       -- freq correction in use at T (ppm)
        mean_delay_near_T              -- mean RTT of nearby peer records (s)
        n_peer_near_T                  -- count of peer records used for network estimate
        active_server_at_T             -- server address active at T (str or None)
        b_asym_raw                     -- RTT/2 before any geographic tightening (s)
        b_asym                         -- asymmetry bound used in expansion (s)
        server_location_note           -- description of how server location was resolved
        note                           -- human-readable summary string
    """
    # Convert every loopstats row to a single comparable float: seconds since MJD epoch
    def loop_abs_sec(row):
        return row.mjd * 86400.0 + row.sec_of_day

    def peer_abs_sec(row):
        return row.mjd * 86400.0 + row.sec_of_day

    query_abs = query_mjd * 86400.0 + query_sec

    sorted_loop = sorted(loop_rows, key=loop_abs_sec)
    if not sorted_loop:
        raise RuntimeError("No loopstats rows available for point-in-time estimate.")

    # Find the last loopstats record at or before T ("before"), and the first after T ("after").
    before = None
    after = None
    for row in sorted_loop:
        t = loop_abs_sec(row)
        if t <= query_abs:
            before = row
        elif after is None:
            after = row

    if before is None:
        # T is before all data — use the earliest record, extrapolating backward
        before = sorted_loop[0]

    # Time gaps
    gap_before_s = query_abs - loop_abs_sec(before)
    gap_after_s = (loop_abs_sec(after) - query_abs) if after is not None else None

    freq_ppm = before.freq
    offset_before = before.offset
    offset_after = after.offset if after is not None else None

    # Best-estimate offset at T:
    #   loopstats freq is the correction NTP is *already applying* to the kernel clock,
    #   not an uncompensated drift.  Projecting it forward double-counts the correction.
    #   When the record on both sides of T is available, linear interpolation between
    #   the two measured offsets is the most accurate estimate.  When only the record
    #   before T is available, use that offset directly (best known value).
    if after is not None:
        total_gap = gap_before_s + gap_after_s
        fraction = gap_before_s / total_gap if total_gap > 0 else 0.0
        best_offset = before.offset + fraction * (after.offset - before.offset)
    else:
        best_offset = before.offset

    # Uncertainty from residual clock drift between the last sync and T.
    #   When interpolating: the jitter at the surrounding records bounds how much the
    #   true offset deviates from a straight line between them.
    #   When extrapolating (no after record): the freq value tells us the rate at which
    #   ntpd was steering the clock; the residual uncertainty grows with the gap.
    if after is not None:
        u_drift = max(before.jitter, after.jitter)
    else:
        freq_drift_bound = abs(freq_ppm * 1e-6 * gap_before_s)
        u_drift = max(freq_drift_bound / SQRT3, before.jitter)

    # Gather peer records within ±window_seconds of T for the network estimate.
    selected_peer_rows, _note = select_peer_subset(peer_rows)
    near_peers = [
        row for row in selected_peer_rows
        if abs(peer_abs_sec(row) - query_abs) <= window_seconds
    ]
    # Fall back to the full selected subset if the window is empty.
    if not near_peers:
        near_peers = selected_peer_rows

    # Determine which server was active at T using the same reduce_to_active_timeline logic.
    active_server_at_T = None
    active_timeline = reduce_to_active_timeline(selected_peer_rows)
    last_active = None
    for row in active_timeline:
        if peer_abs_sec(row) <= query_abs:
            last_active = row
        else:
            break
    if last_active is not None:
        active_server_at_T = last_active.server_address if hasattr(last_active, "server_address") else None

    # Network path asymmetry uncertainty from nearby peers.
    # Use only the active server's records near T if available;
    # fall back to all selected near-peers.
    active_near_peers = [
        row for row in near_peers
        if (row.server_address if hasattr(row, "server_address") else None) == active_server_at_T
    ] if active_server_at_T else []
    network_peers = active_near_peers if active_near_peers else near_peers

    delays_near = [row.delay for row in network_peers]
    mean_delay_near = mean(delays_near) if delays_near else 0.0
    # Hard rectangular bound on one-way path asymmetry: RTT/2 in either direction.
    # Applying k=2 to the rectangular standard uncertainty (RTT/2/sqrt(3)) alone produces
    # an expanded interval ~15% wider than the 100% physical ceiling (RTT/2).  Instead,
    # the rectangular 95% coverage (0.95 * b_asym) and the Gaussian statistical components
    # (drift + scatter, k=2) are combined in quadrature so U_expanded stays within RTT/2.
    b_asym_raw = mean_delay_near / 2.0
    b_asym = b_asym_raw

    # Geographic tightening: if the server location is known, the minimum one-way
    # propagation delay (d_min = geo_distance / v_fibre) is symmetric and cannot
    # be asymmetric.  The true asymmetry bound shrinks to max(RTT/2 - d_min, 0).
    server_loc = resolve_server_location(
        active_server_at_T or "",
        known_servers,
        observer_lat,
        observer_lon,
    )
    d_min_s = server_loc["d_min_s"]
    if d_min_s is not None:
        b_asym = max(b_asym_raw - d_min_s, 0.0)
    server_location_note = server_loc["location_note"]

    u_asymmetry = b_asym / SQRT3               # rectangular standard uncertainty (for u_combined)

    # Scatter: standard deviation of offsets among nearby loopstats records.
    # Use a matching window around T.
    near_loop = [
        row for row in sorted_loop
        if abs(loop_abs_sec(row) - query_abs) <= window_seconds
    ]
    near_offsets = [row.offset for row in near_loop] if near_loop else [before.offset]
    u_scatter = stdev(near_offsets) if len(near_offsets) >= 2 else 0.0

    u_combined = math.sqrt(u_drift * u_drift + u_asymmetry * u_asymmetry + u_scatter * u_scatter)
    # Expanded uncertainty: rectangular asymmetry at 95% of its hard bound, Gaussian
    # statistical terms at k=2, combined in quadrature.
    u_stat = math.sqrt(u_drift * u_drift + u_scatter * u_scatter)
    u_expanded = math.sqrt((0.95 * b_asym) ** 2 + (2.0 * u_stat) ** 2)

    # --- Alternative low-uncertainty estimate from candidate peers ---
    # The loopstats offset is disciplined by the sys.peer only, so u_asymmetry above
    # reflects that peer's RTT.  Other candidate-or-better peers (select codes 4-7)
    # carry their own independently-measured UTC offsets in peerstats column [4].
    # A peer with a smaller combined sqrt(u_asym^2 + u_jitter^2) may give a
    # lower-uncertainty independent estimate.  We use its offset AND delay together
    # (never mixing another peer's delay with the loopstats offset).
    candidate_near = [
        row for row in peer_rows
        if abs(peer_abs_sec(row) - query_abs) <= window_seconds
        and is_candidate_or_better(row.status)
    ]
    alt_candidate_near = [
        row for row in candidate_near
        if (row.server_address if hasattr(row, "server_address") else "") != active_server_at_T
    ] if active_server_at_T else candidate_near

    alt_best_offset = None
    alt_u_asymmetry = None
    alt_u_scatter = None
    alt_u_combined = None
    alt_u_expanded = None
    alt_server = None
    alt_delay = None
    alt_gap_s = None
    n_candidate_peers_near_T = len(set(
        row.server_address for row in candidate_near
        if hasattr(row, "server_address")
    ))

    if alt_candidate_near:
        # For each distinct server keep only the record nearest in time to T.
        by_server = {}
        for row in alt_candidate_near:
            addr = row.server_address if hasattr(row, "server_address") else ""
            gap = abs(peer_abs_sec(row) - query_abs)
            current = by_server.get(addr)
            if current is None or gap < current[0]:
                by_server[addr] = (gap, row)
        # Score each server by combined uncertainty; pick the minimum.
        best_score = None
        for addr, (gap, row) in by_server.items():
            b_a = row.delay / 2.0
            u_a = b_a / SQRT3
            u_s = row.jitter
            score = math.sqrt(u_a * u_a + u_s * u_s)
            if best_score is None or score < best_score:
                best_score = score
                alt_server = addr
                alt_delay = row.delay
                alt_gap_s = gap
                alt_u_asymmetry = u_a
                alt_u_scatter = u_s
                alt_best_offset = row.offset
                alt_u_combined = score
                # Same expansion as sys.peer: 95% rectangular asymmetry + k=2 Gaussian jitter.
                alt_u_expanded = math.sqrt((0.95 * b_a) ** 2 + (2.0 * u_s) ** 2)

    note_parts = [
        "gap_before=%.1fs" % gap_before_s,
        "freq=%.3fppm" % freq_ppm,
        "u_drift=%s" % format_ms(u_drift),
        "u_asymmetry=%s" % format_ms(u_asymmetry),
        "u_scatter=%s" % format_ms(u_scatter),
        "U_expanded(~95%%)=+/-%s" % format_ms(u_expanded),
    ]
    if active_server_at_T:
        note_parts.insert(0, "server=%s" % active_server_at_T)
    if d_min_s is not None and d_min_s > 0.0:
        note_parts.append("b_asym_tightened=%s->%s" % (format_ms(b_asym_raw), format_ms(b_asym)))

    return {
        "query_mjd": query_mjd,
        "query_sec": query_sec,
        "best_offset": best_offset,
        "offset_before": offset_before,
        "offset_after": offset_after,
        "u_drift": u_drift,
        "u_asymmetry": u_asymmetry,
        "u_scatter": u_scatter,
        "u_combined": u_combined,
        "u_expanded": u_expanded,
        "gap_before_s": gap_before_s,
        "gap_after_s": gap_after_s,
        "freq_ppm": freq_ppm,
        "mean_delay_near_T": mean_delay_near,
        "n_peer_near_T": len(network_peers),
        "active_server_at_T": active_server_at_T,
        "b_asym_raw": b_asym_raw,
        "b_asym": b_asym,
        "server_location_note": server_location_note,
        "n_candidate_peers_near_T": n_candidate_peers_near_T,
        "alt_best_offset": alt_best_offset,
        "alt_u_asymmetry": alt_u_asymmetry,
        "alt_u_scatter": alt_u_scatter,
        "alt_u_combined": alt_u_combined,
        "alt_u_expanded": alt_u_expanded,
        "alt_server": alt_server,
        "alt_delay": alt_delay,
        "alt_gap_s": alt_gap_s,
        "note": "; ".join(note_parts),
    }


def analyze(option, loop_rows, peer_rows, known_servers=None, observer_lat=None, observer_lon=None):
    if not loop_rows:
        raise RuntimeError("No usable loopstats rows were found for the selected day.")
    if not peer_rows:
        raise RuntimeError("No usable peerstats rows were found for the selected day.")

    peer_subset, peer_selection_note = select_peer_subset(peer_rows)
    diagnostics = compute_peer_diagnostics(peer_subset)

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

    a_uncertainty = math.sqrt(mean_offset_signed * mean_offset_signed + mean_delay * mean_delay)

    b_u_offset = mean_offset_abs / SQRT3
    b_u_delay = mean_half_delay / SQRT3
    b_u_server = 3e-6
    b_u_combined = math.sqrt(b_u_offset * b_u_offset + b_u_delay * b_u_delay + b_u_server * b_u_server)
    b_u_expanded = 2.0 * b_u_combined

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

    # Offset accuracy to UTC (3 practical variants)
    # Variant E: Minimal (network asymmetry + NTP measurement variation)
    e_u_asymmetry = mean_half_delay / SQRT3
    e_u_measurement = stdev(offsets)
    e_u_combined = math.sqrt(e_u_asymmetry * e_u_asymmetry + e_u_measurement * e_u_measurement)
    e_u_expanded = 2.0 * e_u_combined

    # Variant F: Using NTP's dispersion directly
    f_u_offset = mean_dispersion
    f_u_expanded = 2.0 * f_u_offset

    # Variant G: Conservative (worst-case delay)
    max_delay = max(delays) if delays else mean_delay
    g_u_asymmetry = (max_delay / 2.0) / SQRT3
    g_u_measurement = stdev(offsets)
    g_u_combined = math.sqrt(g_u_asymmetry * g_u_asymmetry + g_u_measurement * g_u_measurement)
    g_u_expanded = 2.0 * g_u_combined

    # Point-in-time estimate: use the last loopstats record as the representative query time.
    # Callers can also call estimate_offset_at_time() directly with any (mjd, sec) pair.
    sorted_loop_for_pit = sorted(loop_rows, key=lambda r: (r.mjd, r.sec_of_day))
    last_loop = sorted_loop_for_pit[-1]
    pit_result = estimate_offset_at_time(last_loop.mjd, last_loop.sec_of_day, loop_rows, peer_rows,
                                         known_servers=known_servers,
                                         observer_lat=observer_lat, observer_lon=observer_lon)

    mjds = sorted(set([row.mjd for row in loop_rows] + [row.mjd for row in peer_rows]))

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
        "e_u_asymmetry": e_u_asymmetry,
        "e_u_measurement": e_u_measurement,
        "e_u_combined": e_u_combined,
        "e_u_expanded": e_u_expanded,
        "f_u_offset": f_u_offset,
        "f_u_expanded": f_u_expanded,
        "g_u_asymmetry": g_u_asymmetry,
        "g_u_measurement": g_u_measurement,
        "g_u_combined": g_u_combined,
        "g_u_expanded": g_u_expanded,
    }

    return AnalysisResult(
        option=option,
        mjds=mjds,
        loop_rows_used=len(loop_rows),
        peer_rows_total=len(peer_rows),
        peer_rows_used=len(peer_subset),
        peer_selection_note=peer_selection_note,
        metrics=metrics,
        diagnostics=diagnostics,
        pit_result=pit_result,
    )


def generate_report(result):
    date_text = ", ".join(["%s (MJD %d)" % (mjd_to_date_string(mjd), mjd) for mjd in result.mjds])
    option = result.option
    m = result.metrics
    d = result.diagnostics

    code_counts = d.get("select_code_counts", {})
    code_6 = code_counts.get(6, 0)
    code_7 = code_counts.get(7, 0)
    code_other = 0
    for code_value, count in code_counts.items():
        if code_value not in (6, 7):
            code_other += count

    lines = [
        "NTP Timing Accuracy Report",
        "=" * 80,
        "Dataset: %s" % option.label,
        "Loopstats file: %s" % option.loop_path,
        "Peerstats file: %s" % option.peer_path,
        "Day(s) present in selected data: %s" % date_text,
        "Loopstats rows used: %d" % result.loop_rows_used,
        "Peerstats rows used: %d (%s)" % (result.peer_rows_used, result.peer_selection_note),
        "",
        "Selected Peer Diagnostics",
        "-" * 80,
        "Select code 6 (sys.peer) rows: %d" % code_6,
        "Select code 7 (pps.peer) rows: %d" % code_7,
        "Other select codes in subset: %d" % code_other,
        "Active timeline rows (1 peer per second): %d" % d.get("active_rows", 0),
        "Active source transitions: %d" % d.get("active_transitions", 0),
        "Unique active servers: %d" % d.get("active_unique_servers", 0),
        "",
        "Interpretation A (literal reading)",
        "-" * 80,
        "Mean signed loop offset: %s" % format_ms(m["mean_offset_signed"]),
        "Mean peer delay (RTT): %s" % format_ms(m["mean_delay"]),
        "Quadrature sum: %s" % format_ms(m["a_uncertainty"]),
        "",
        "Interpretation B (metrological)",
        "-" * 80,
        "u_offset = mean(abs(offset))/sqrt(3): %s" % format_ms(m["b_u_offset"]),
        "u_delay = (mean(delay)/2)/sqrt(3): %s" % format_ms(m["b_u_delay"]),
        "u_server (fixed): %s" % format_us(m["b_u_server"]),
        "u_combined (k=1): %s" % format_ms(m["b_u_combined"]),
        "U_expanded (k=2): +/- %s" % format_ms(m["b_u_expanded"]),
        "",
        "Interpretation C (statistical)",
        "-" * 80,
        "Systematic bias (mean signed offset): %s" % format_ms(m["c_bias"]),
        "u_wander = stdev(offset): %s" % format_ms(m["c_u_wander"]),
        "u_asymmetry = (mean(delay)/2)/sqrt(3): %s" % format_ms(m["c_u_asymmetry"]),
        "u_delay_variation = stdev(delay)/2: %s" % format_ms(m["c_u_delay_variation"]),
        "u_server (rectangular): %s" % format_us(m["c_u_server"]),
        "u_combined (k=1): %s" % format_ms(m["c_u_combined"]),
        "U_expanded (k=2): +/- %s" % format_ms(m["c_u_expanded"]),
        "",
        "Interpretation D (NTP native statistics)",
        "-" * 80,
        "u_jitter (loopstats mean jitter): %s" % format_ms(m["d_u_jitter"]),
        "u_dispersion (peerstats mean dispersion): %s" % format_ms(m["d_u_dispersion"]),
        "u_peer_jitter (informational): %s" % format_ms(m["d_u_peer_jitter"]),
        "u_asymmetry = (mean(delay)/2)/sqrt(3): %s" % format_ms(m["d_u_asymmetry"]),
        "u_server (rectangular): %s" % format_us(m["d_u_server"]),
        "u_combined (k=1): %s" % format_ms(m["d_u_combined"]),
        "U_expanded (k=2): +/- %s" % format_ms(m["d_u_expanded"]),
        "",
        "Offset Accuracy to UTC (practical variants)",
        "=" * 80,
        "",
        "Variant E (minimal: network + measurement)",
        "-" * 80,
        "u_asymmetry = (mean(delay)/2)/sqrt(3): %s" % format_ms(m["e_u_asymmetry"]),
        "u_measurement = stdev(offset): %s" % format_ms(m["e_u_measurement"]),
        "u_combined (k=1): %s" % format_ms(m["e_u_combined"]),
        "U_expanded (k=2): +/- %s" % format_ms(m["e_u_expanded"]),
        "",
        "Variant F (using NTP's dispersion directly)",
        "-" * 80,
        "u_offset = mean(dispersion): %s" % format_ms(m["f_u_offset"]),
        "U_expanded (k=2): +/- %s" % format_ms(m["f_u_expanded"]),
        "",
        "Variant G (conservative: worst-case delay)",
        "-" * 80,
        "u_asymmetry = (max(delay)/2)/sqrt(3): %s" % format_ms(m["g_u_asymmetry"]),
        "u_measurement = stdev(offset): %s" % format_ms(m["g_u_measurement"]),
        "u_combined (k=1): %s" % format_ms(m["g_u_combined"]),
        "U_expanded (k=2): +/- %s" % format_ms(m["g_u_expanded"]),
    ]

    return "\r\n".join(lines)


def sanitize_filename_part(value):
    sanitized = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_")
    return sanitized or "dataset"


def resolve_export_paths(export_dir, result):
    if not os.path.isdir(export_dir):
        os.makedirs(export_dir)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if result.mjds:
        day_part = sanitize_filename_part(mjd_to_date_string(max(result.mjds)))
    else:
        day_part = "unknown_day"

    base = os.path.join(export_dir, "ntp_accuracy_%s_%s" % (day_part, stamp))
    return base + ".json", base + ".csv"


def export_json(path, result):
    generated_at = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    payload = {
        "generated_at": generated_at,
        "dataset_label": result.option.label,
        "loopstats_file": result.option.loop_path,
        "peerstats_file": result.option.peer_path,
        "mjd_days": result.mjds,
        "iso_days": [mjd_to_date_string(mjd) for mjd in result.mjds],
        "loop_rows_used": result.loop_rows_used,
        "peer_rows_total": result.peer_rows_total,
        "peer_rows_used": result.peer_rows_used,
        "peer_selection_note": result.peer_selection_note,
        "metrics_seconds": result.metrics,
        "diagnostics": result.diagnostics,
    }

    handle = open(path, "w", encoding="utf-8")
    try:
        json.dump(payload, handle, indent=2, sort_keys=True)
    finally:
        handle.close()


def export_csv(path, result):
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
        "diag_select_code_6",
        "diag_select_code_7",
        "diag_select_code_other",
        "diag_active_rows",
        "diag_active_transitions",
        "diag_active_unique_servers",
    ] + sorted(result.metrics.keys())

    row = {
        "generated_at": generated_at,
        "dataset_label": result.option.label,
        "loopstats_file": result.option.loop_path,
        "peerstats_file": result.option.peer_path,
        "mjd_days": ";".join([str(mjd) for mjd in result.mjds]),
        "iso_days": ";".join([mjd_to_date_string(mjd) for mjd in result.mjds]),
        "loop_rows_used": str(result.loop_rows_used),
        "peer_rows_total": str(result.peer_rows_total),
        "peer_rows_used": str(result.peer_rows_used),
        "peer_selection_note": result.peer_selection_note,
    }

    diag = result.diagnostics or {}
    code_counts = diag.get("select_code_counts", {})
    code_other = 0
    for code_value, count in code_counts.items():
        if code_value not in (6, 7):
            code_other += count

    row["diag_select_code_6"] = str(code_counts.get(6, 0))
    row["diag_select_code_7"] = str(code_counts.get(7, 0))
    row["diag_select_code_other"] = str(code_other)
    row["diag_active_rows"] = str(diag.get("active_rows", 0))
    row["diag_active_transitions"] = str(diag.get("active_transitions", 0))
    row["diag_active_unique_servers"] = str(diag.get("active_unique_servers", 0))

    for key, value in result.metrics.items():
        row[key] = "%.12f" % value

    handle = open(path, "w", newline="", encoding="utf-8")
    try:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        writer.writerow(row)
    finally:
        handle.close()


class AnalyzerForm(Form):
    def __init__(self):
        self.Text = "NTP Timing Accuracy Analyzer"
        self.Size = Size(1600, 960)
        self.MinimumSize = Size(1100, 700)
        self.StartPosition = FormStartPosition.CenterScreen
        self.WindowState = FormWindowState.Maximized

        self._options_by_label = {}
        self._plot_data = {}
        self._last_loop_rows = []
        self._last_peer_rows = []
        self._last_result = None
        self._last_aggregate_report = ""
        _known_servers_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "..", "resources", "national_utc_ntp_servers.json",
        )
        self._known_servers = load_known_servers(os.path.normpath(_known_servers_path))
        _load_ip_location_cache(os.path.normpath(os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "..", "resources", "ip_location_cache.json",
        )))

        default_font = Font("Segoe UI", 9)
        bold_font = Font("Segoe UI", 9, FontStyle.Bold)

        split = SplitContainer()
        split.Dock = DockStyle.Fill
        split.FixedPanel = FixedPanel.Panel1
        split.SplitterWidth = 6
        self.Controls.Add(split)
        self._main_split = split
        self.Shown += self.on_form_shown
        split.Panel1.Resize += self.on_left_panel_resize

        lp = split.Panel1

        self.lbl_title = Label()
        self.lbl_title.Text = "NTP Timing Accuracy - Interpretations A, B, C, D"
        self.lbl_title.Font = Font("Segoe UI", 11, FontStyle.Bold)
        self.lbl_title.Location = Point(8, 8)
        self.lbl_title.Size = Size(440, 26)
        self.lbl_title.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.lbl_title)

        self.lbl_log = Label()
        self.lbl_log.Text = "NTP log folder:"
        self.lbl_log.Font = bold_font
        self.lbl_log.Location = Point(8, 44)
        self.lbl_log.Size = Size(200, 20)
        lp.Controls.Add(self.lbl_log)

        self.txt_log_folder = TextBox()
        self.txt_log_folder.Font = default_font
        self.txt_log_folder.Location = Point(8, 66)
        self.txt_log_folder.Size = Size(446, 24)
        self.txt_log_folder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.txt_log_folder)

        self.btn_browse_log = Button()
        self.btn_browse_log.Text = "Browse..."
        self.btn_browse_log.Location = Point(8, 96)
        self.btn_browse_log.Size = Size(100, 28)
        self.btn_browse_log.Click += self.on_browse_log
        lp.Controls.Add(self.btn_browse_log)

        self.btn_scan = Button()
        self.btn_scan.Text = "Scan Datasets"
        self.btn_scan.Location = Point(114, 96)
        self.btn_scan.Size = Size(120, 28)
        self.btn_scan.Click += self.on_scan
        lp.Controls.Add(self.btn_scan)

        self.lbl_filter = Label()
        self.lbl_filter.Text = "Day filter (optional text / MJD / YYYYMMDD):"
        self.lbl_filter.Location = Point(8, 136)
        self.lbl_filter.Size = Size(440, 20)
        lp.Controls.Add(self.lbl_filter)

        self.txt_day_filter = TextBox()
        self.txt_day_filter.Location = Point(8, 158)
        self.txt_day_filter.Size = Size(328, 24)
        self.txt_day_filter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.txt_day_filter)

        self.btn_apply_filter = Button()
        self.btn_apply_filter.Text = "Apply Filter"
        self.btn_apply_filter.Location = Point(342, 156)
        self.btn_apply_filter.Size = Size(112, 28)
        self.btn_apply_filter.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btn_apply_filter.Click += self.on_scan
        lp.Controls.Add(self.btn_apply_filter)

        self.lbl_dataset = Label()
        self.lbl_dataset.Text = "Dataset:"
        self.lbl_dataset.Font = bold_font
        self.lbl_dataset.Location = Point(8, 198)
        self.lbl_dataset.Size = Size(200, 20)
        lp.Controls.Add(self.lbl_dataset)

        self.cmb_dataset = ComboBox()
        self.cmb_dataset.DropDownStyle = ComboBoxStyle.DropDownList
        self.cmb_dataset.Location = Point(8, 220)
        self.cmb_dataset.Size = Size(446, 24)
        self.cmb_dataset.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.cmb_dataset)

        self.chk_export = CheckBox()
        self.chk_export.Text = "Export JSON + CSV"
        self.chk_export.Location = Point(8, 258)
        self.chk_export.Size = Size(160, 24)
        self.chk_export.Checked = True
        self.chk_export.CheckedChanged += self.on_export_toggle
        lp.Controls.Add(self.chk_export)

        self.txt_export_folder = TextBox()
        self.txt_export_folder.Location = Point(8, 284)
        self.txt_export_folder.Size = Size(328, 24)
        self.txt_export_folder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.txt_export_folder)

        self.btn_browse_export = Button()
        self.btn_browse_export.Text = "Browse..."
        self.btn_browse_export.Location = Point(342, 282)
        self.btn_browse_export.Size = Size(112, 28)
        self.btn_browse_export.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btn_browse_export.Click += self.on_browse_export
        lp.Controls.Add(self.btn_browse_export)

        self.chk_raw_peer_points = CheckBox()
        self.chk_raw_peer_points.Text = "Charts: raw peer points"
        self.chk_raw_peer_points.Location = Point(8, 320)
        self.chk_raw_peer_points.Size = Size(228, 24)
        self.chk_raw_peer_points.Checked = False
        lp.Controls.Add(self.chk_raw_peer_points)

        self.lbl_pit_time = Label()
        self.lbl_pit_time.Text = "Point-in-time (HH:MM:SS, dataset day):"
        self.lbl_pit_time.Location = Point(8, 348)
        self.lbl_pit_time.Size = Size(280, 20)
        lp.Controls.Add(self.lbl_pit_time)

        self.txt_pit_time = TextBox()
        self.txt_pit_time.Text = ""
        self.txt_pit_time.Location = Point(8, 370)
        self.txt_pit_time.Size = Size(160, 24)
        lp.Controls.Add(self.txt_pit_time)

        self.btn_pit = Button()
        self.btn_pit.Text = "Calculate PIT"
        self.btn_pit.Location = Point(174, 368)
        self.btn_pit.Size = Size(112, 28)
        self.btn_pit.Click += self.on_pit_calculate
        lp.Controls.Add(self.btn_pit)

        self.lbl_pit_result = Label()
        self.lbl_pit_result.Text = "Estimated Offset and Error to use:"
        self.lbl_pit_result.Font = bold_font
        self.lbl_pit_result.Location = Point(8, 402)
        self.lbl_pit_result.Size = Size(280, 20)
        lp.Controls.Add(self.lbl_pit_result)

        self.txt_pit_result = TextBox()
        self.txt_pit_result.Text = ""
        self.txt_pit_result.ReadOnly = True
        self.txt_pit_result.Font = bold_font
        self.txt_pit_result.Location = Point(8, 424)
        self.txt_pit_result.Size = Size(280, 24)
        lp.Controls.Add(self.txt_pit_result)

        self.lbl_pit_note = Label()
        self.lbl_pit_note.Text = "Actual error via fibre likely 2-5x smaller, but no less than the jitter"
        self.lbl_pit_note.Location = Point(8, 450)
        self.lbl_pit_note.Size = Size(280, 34)
        lp.Controls.Add(self.lbl_pit_note)

        self.lbl_observer = Label()
        self.lbl_observer.Text = "Observer location (decimal degrees):"
        self.lbl_observer.Location = Point(8, 490)
        self.lbl_observer.Size = Size(280, 20)
        lp.Controls.Add(self.lbl_observer)

        self.lbl_observer_lat = Label()
        self.lbl_observer_lat.Text = "Lat:"
        self.lbl_observer_lat.Location = Point(8, 515)
        self.lbl_observer_lat.Size = Size(28, 20)
        lp.Controls.Add(self.lbl_observer_lat)

        self.txt_observer_lat = TextBox()
        self.txt_observer_lat.Text = ""
        self.txt_observer_lat.Location = Point(36, 512)
        self.txt_observer_lat.Size = Size(80, 24)
        self.txt_observer_lat.Font = default_font
        lp.Controls.Add(self.txt_observer_lat)

        self.lbl_observer_comma = Label()
        self.lbl_observer_comma.Text = ""
        self.lbl_observer_comma.Location = Point(119, 515)
        self.lbl_observer_comma.Size = Size(4, 20)
        lp.Controls.Add(self.lbl_observer_comma)

        self.lbl_observer_lon = Label()
        self.lbl_observer_lon.Text = "Lon:"
        self.lbl_observer_lon.Location = Point(124, 515)
        self.lbl_observer_lon.Size = Size(28, 20)
        lp.Controls.Add(self.lbl_observer_lon)

        self.txt_observer_lon = TextBox()
        self.txt_observer_lon.Text = ""
        self.txt_observer_lon.Location = Point(152, 512)
        self.txt_observer_lon.Size = Size(80, 24)
        self.txt_observer_lon.Font = default_font
        lp.Controls.Add(self.txt_observer_lon)

        self.lbl_observer_note = Label()
        self.lbl_observer_note.Text = "Used to tighten asymmetry bound for known servers"
        self.lbl_observer_note.Location = Point(8, 540)
        self.lbl_observer_note.Size = Size(280, 20)
        lp.Controls.Add(self.lbl_observer_note)

        self.btn_analyze = Button()
        self.btn_analyze.Text = "Analyze"
        self.btn_analyze.Font = bold_font
        self.btn_analyze.Location = Point(342, 316)
        self.btn_analyze.Size = Size(112, 32)
        self.btn_analyze.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.btn_analyze.Click += self.on_analyze
        lp.Controls.Add(self.btn_analyze)

        self.txt_output = TextBox()
        self.txt_output.Multiline = True
        self.txt_output.ScrollBars = ScrollBars.Both
        self.txt_output.ReadOnly = True
        self.txt_output.Font = Font("Consolas", 9)
        self.txt_output.Location = Point(8, 358)
        self.txt_output.Size = Size(446, 510)
        self.txt_output.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.txt_output)

        self.lbl_status = Label()
        self.lbl_status.Text = "Ready."
        self.lbl_status.Location = Point(8, 880)
        self.lbl_status.Size = Size(446, 22)
        self.lbl_status.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        lp.Controls.Add(self.lbl_status)

        tbl = TableLayoutPanel()
        tbl.Dock = DockStyle.Fill
        tbl.RowCount = 4
        tbl.ColumnCount = 1
        tbl.RowStyles.Clear()
        for _ in range(4):
            tbl.RowStyles.Add(RowStyle(SizeType.Percent, 25.0))
        tbl.ColumnStyles.Clear()
        tbl.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 100.0))
        split.Panel2.Controls.Add(tbl)

        self.chart_delay = self.create_plot_panel("Delay (Peerstats selected server, loop timeline)", Point(0, 0), Size(100, 100), "delay")
        self.chart_delay.Dock = DockStyle.Fill
        tbl.Controls.Add(self.chart_delay, 0, 0)

        self.chart_offset = self.create_plot_panel("Offset (Loopstats)", Point(0, 0), Size(100, 100), "offset")
        self.chart_offset.Dock = DockStyle.Fill
        tbl.Controls.Add(self.chart_offset, 0, 1)

        self.chart_jitter = self.create_plot_panel("Jitter (Loopstats / Peerstats)", Point(0, 0), Size(100, 100), "jitter")
        self.chart_jitter.Dock = DockStyle.Fill
        tbl.Controls.Add(self.chart_jitter, 0, 2)

        self.chart_dispersion = self.create_plot_panel("Dispersion (Peerstats)", Point(0, 0), Size(100, 100), "dispersion")
        self.chart_dispersion.Dock = DockStyle.Fill
        tbl.Controls.Add(self.chart_dispersion, 0, 3)

        self.prefill_defaults()

    def set_status(self, text):
        self.lbl_status.Text = text

    def on_form_shown(self, sender, event):
        # SplitterDistance is only valid after the control has a real width.
        split = self._main_split
        available_width = split.ClientSize.Width
        if available_width <= 0:
            return

        # Compute how much width the left panel actually needs after DPI scaling.
        required_right = 0
        for ctrl in split.Panel1.Controls:
            if not ctrl.Visible:
                continue
            required_right = max(required_right, int(ctrl.Right))

        # Add breathing room so right-anchored controls (Analyze/Browse) stay visible.
        preferred = max(640, required_right + 18)

        desired_left_min = 420
        desired_right_min = 500

        # If the current width cannot satisfy both desired mins, scale them down.
        if available_width < (desired_left_min + desired_right_min):
            left_min = max(120, int(available_width * 0.35))
            right_min = max(120, available_width - left_min - 1)
            if right_min < 120:
                right_min = 120
                left_min = max(120, available_width - right_min - 1)
        else:
            left_min = desired_left_min
            right_min = desired_right_min

        split.Panel1MinSize = left_min
        split.Panel2MinSize = right_min

        min_left = split.Panel1MinSize
        max_left = available_width - split.Panel2MinSize

        # Clamp to the valid runtime range to avoid WinForms SystemError.
        if max_left < min_left:
            left_width = min_left
        else:
            left_width = max(min_left, min(preferred, max_left))

        try:
            split.SplitterDistance = left_width
        except Exception:
            # Ignore one-time layout race conditions on some WinForms runtimes.
            pass

        self.adjust_left_panel_layout()

    def on_left_panel_resize(self, sender, event):
        self.adjust_left_panel_layout()

    def adjust_left_panel_layout(self):
        panel = self._main_split.Panel1
        panel_width = panel.ClientSize.Width
        if panel_width <= 0:
            return

        margin = 8
        inter = 7  # ~20% more vertical separation than the previous 6 px spacing
        label_h = 22
        text_h = 24
        button_h = 34  # ~20% taller than the previous 28 px buttons
        analyze_h = 38
        full_w = max(120, panel_width - (margin * 2))

        y = 8

        self.lbl_title.Location = Point(margin, y)
        self.lbl_title.Size = Size(full_w, 28)
        y += 28 + inter

        self.lbl_log.Location = Point(margin, y)
        self.lbl_log.Size = Size(full_w, label_h)
        y += label_h

        self.txt_log_folder.Location = Point(margin, y)
        self.txt_log_folder.Size = Size(full_w, text_h)
        y += text_h + inter

        btn_gap = 6
        browse_w = 100
        scan_w = 120
        self.btn_browse_log.Location = Point(margin, y)
        self.btn_browse_log.Size = Size(browse_w, button_h)
        self.btn_scan.Location = Point(margin + browse_w + btn_gap, y)
        self.btn_scan.Size = Size(scan_w, button_h)
        y += button_h + inter

        self.lbl_filter.Location = Point(margin, y)
        self.lbl_filter.Size = Size(full_w, label_h)
        y += label_h

        apply_w = 112
        filter_gap = 6
        if full_w >= (apply_w + 150 + filter_gap):
            filter_w = max(120, full_w - apply_w - filter_gap)
            self.txt_day_filter.Location = Point(margin, y)
            self.txt_day_filter.Size = Size(filter_w, text_h)
            self.btn_apply_filter.Location = Point(margin + filter_w + filter_gap, y - 1)
            self.btn_apply_filter.Size = Size(apply_w, button_h)
            y += max(text_h, button_h) + inter
        else:
            self.txt_day_filter.Location = Point(margin, y)
            self.txt_day_filter.Size = Size(full_w, text_h)
            y += text_h + 4
            self.btn_apply_filter.Location = Point(margin, y)
            self.btn_apply_filter.Size = Size(apply_w, button_h)
            y += button_h + inter

        self.lbl_dataset.Location = Point(margin, y)
        self.lbl_dataset.Size = Size(full_w, label_h)
        y += label_h

        self.cmb_dataset.Location = Point(margin, y)
        self.cmb_dataset.Size = Size(full_w, text_h)
        y += text_h + inter

        self.chk_export.Location = Point(margin, y)
        self.chk_export.Size = Size(min(220, full_w), text_h)
        y += text_h

        analyze_w = 112
        browse_export_w = 112
        controls_gap = 6
        one_row_min = browse_export_w + analyze_w + 160 + controls_gap * 2
        two_row_min = browse_export_w + analyze_w + controls_gap

        if full_w >= one_row_min:
            export_w = max(120, full_w - browse_export_w - analyze_w - controls_gap * 2)
            self.txt_export_folder.Location = Point(margin, y)
            self.txt_export_folder.Size = Size(export_w, text_h)

            bx = margin + export_w + controls_gap
            self.btn_browse_export.Location = Point(bx, y - 1)
            self.btn_browse_export.Size = Size(browse_export_w, button_h)

            ax = bx + browse_export_w + controls_gap
            self.btn_analyze.Location = Point(ax, y - 3)
            self.btn_analyze.Size = Size(analyze_w, analyze_h)
            y += max(text_h, analyze_h) + inter
        elif full_w >= two_row_min:
            self.txt_export_folder.Location = Point(margin, y)
            self.txt_export_folder.Size = Size(full_w, text_h)
            y += text_h + 4

            self.btn_browse_export.Location = Point(margin, y)
            self.btn_browse_export.Size = Size(browse_export_w, button_h)
            self.btn_analyze.Location = Point(margin + browse_export_w + controls_gap, y - 2)
            self.btn_analyze.Size = Size(analyze_w, analyze_h)
            y += max(button_h, analyze_h) + inter
        else:
            self.txt_export_folder.Location = Point(margin, y)
            self.txt_export_folder.Size = Size(full_w, text_h)
            y += text_h + 4

            self.btn_browse_export.Location = Point(margin, y)
            self.btn_browse_export.Size = Size(browse_export_w, button_h)
            y += button_h + 4

            self.btn_analyze.Location = Point(margin, y)
            self.btn_analyze.Size = Size(analyze_w, analyze_h)
            y += analyze_h + inter

        self.chk_raw_peer_points.Location = Point(margin, y)
        self.chk_raw_peer_points.Size = Size(min(260, full_w), text_h)
        y += text_h + inter

        self.lbl_pit_time.Location = Point(margin, y)
        self.lbl_pit_time.Size = Size(full_w, label_h)
        y += label_h

        pit_btn_w = 120
        pit_gap = 6
        pit_txt_w = max(80, full_w - pit_btn_w - pit_gap)
        self.txt_pit_time.Location = Point(margin, y)
        self.txt_pit_time.Size = Size(pit_txt_w, text_h)
        self.btn_pit.Location = Point(margin + pit_txt_w + pit_gap, y - 1)
        self.btn_pit.Size = Size(pit_btn_w, button_h)
        y += max(text_h, button_h) + inter

        self.lbl_pit_result.Location = Point(margin, y)
        self.lbl_pit_result.Size = Size(full_w, label_h)
        y += label_h

        self.txt_pit_result.Location = Point(margin, y)
        self.txt_pit_result.Size = Size(full_w, text_h)
        y += text_h + 3

        self.lbl_pit_note.Location = Point(margin, y)
        self.lbl_pit_note.Size = Size(full_w, label_h * 2)
        y += label_h * 2 + inter

        self.lbl_observer.Location = Point(margin, y)
        self.lbl_observer.Size = Size(full_w, label_h)
        y += label_h

        lat_lbl_w = 28
        lat_w = 80
        gap = 8
        lon_lbl_w = 28
        lon_w = 80
        x = margin
        self.lbl_observer_lat.Location = Point(x, y + 3)
        self.lbl_observer_lat.Size = Size(lat_lbl_w, label_h)
        x += lat_lbl_w
        self.txt_observer_lat.Location = Point(x, y)
        self.txt_observer_lat.Size = Size(lat_w, text_h)
        x += lat_w + gap
        self.lbl_observer_comma.Location = Point(x, y + 3)
        self.lbl_observer_comma.Size = Size(4, label_h)
        self.lbl_observer_lon.Location = Point(x, y + 3)
        self.lbl_observer_lon.Size = Size(lon_lbl_w, label_h)
        x += lon_lbl_w
        self.txt_observer_lon.Location = Point(x, y)
        self.txt_observer_lon.Size = Size(lon_w, text_h)
        y += text_h + 2

        self.lbl_observer_note.Location = Point(margin, y)
        self.lbl_observer_note.Size = Size(full_w, label_h)
        y += label_h + inter

        content_top = y
        status_height = 24
        output_height = max(120, panel.ClientSize.Height - content_top - status_height)
        self.txt_output.Location = Point(margin, content_top)
        self.txt_output.Size = Size(full_w, output_height)
        self.lbl_status.Location = Point(margin, panel.ClientSize.Height - status_height)
        self.lbl_status.Size = Size(full_w, status_height)

    def create_plot_panel(self, title, location, size, plot_key):
        container = Label()
        container.Text = title
        container.Font = Font("Segoe UI", 9, FontStyle.Bold)
        container.Location = location
        container.Size = size
        container.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

        header_h = 30

        if plot_key == "jitter":
            legend_y = 4
            legend_font = Font("Segoe UI", 8)
            start_x = 300
            center_y = int(header_h / 2)

            loop_swatch = Label()
            loop_swatch.BackColor = Color.FromArgb(255, 127, 14)
            loop_swatch.BorderStyle = BorderStyle.FixedSingle
            loop_swatch.Location = Point(start_x, center_y - 2)
            loop_swatch.Size = Size(18, 5)
            container.Controls.Add(loop_swatch)

            loop_label = Label()
            loop_label.Text = "Loop (Local) Jitter"
            loop_label.Font = legend_font
            loop_label.Location = Point(start_x + 22, legend_y - 1)
            loop_label.Size = Size(150, 20)
            container.Controls.Add(loop_label)

            peer_x = start_x + 210
            peer_swatch = Label()
            peer_swatch.BackColor = Color.FromArgb(44, 160, 44)
            peer_swatch.BorderStyle = BorderStyle.FixedSingle
            peer_swatch.Location = Point(peer_x, center_y - 1)
            peer_swatch.Size = Size(14, 3)
            container.Controls.Add(peer_swatch)

            peer_label = Label()
            peer_label.Text = "Peer (Network) Jitter"
            peer_label.Font = legend_font
            peer_label.Location = Point(peer_x + 18, legend_y - 1)
            peer_label.Size = Size(180, 20)
            container.Controls.Add(peer_label)
        elif plot_key == "delay":
            # Server legend for delay chart is rendered in the header container.
            self._delay_legend_controls = []

        plot_box = PictureBox()
        plot_box.Location = Point(0, header_h)
        plot_box.Size = Size(size.Width, size.Height - header_h)
        plot_box.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        plot_box.BorderStyle = BorderStyle.FixedSingle
        plot_box.BackColor = Color.White
        plot_box.Tag = plot_key
        plot_box.Paint += self.on_plot_paint
        container.Controls.Add(plot_box)
        return container

    def update_delay_header_legend(self, server_to_color, unique_servers, server_to_km=None):
        if not hasattr(self, "chart_delay") or self.chart_delay is None:
            return

        # Remove prior dynamic delay legend controls.
        old_controls = getattr(self, "_delay_legend_controls", [])
        for ctrl in old_controls:
            try:
                self.chart_delay.Controls.Remove(ctrl)
                ctrl.Dispose()
            except Exception:
                pass
        self._delay_legend_controls = []

        if not unique_servers:
            return

        legend_font = Font("Segoe UI", 8)
        legend_y = 4
        start_x = 460
        center_y = 15
        x_pos = start_x

        for server in unique_servers:
            if x_pos > self.chart_delay.ClientSize.Width - 178:
                break

            swatch = Label()
            swatch.BackColor = server_to_color.get(server, Color.Gray)
            swatch.BorderStyle = BorderStyle.FixedSingle
            swatch.Location = Point(x_pos, center_y - 2)
            swatch.Size = Size(18, 5)
            self.chart_delay.Controls.Add(swatch)
            self._delay_legend_controls.append(swatch)

            label = Label()
            server_label_text = server if server else "Unknown"
            if server_to_km and server in server_to_km:
                server_label_text = "%s (%d km)" % (server_label_text, int(round(server_to_km[server])))
            label.Text = server_label_text
            label.Font = legend_font
            label.Location = Point(x_pos + 22, legend_y - 1)
            label.Size = Size(150, 20)
            self.chart_delay.Controls.Add(label)
            self._delay_legend_controls.append(label)

            x_pos += 178

    def _get_plot_box(self, container):
        for ctrl in container.Controls:
            if isinstance(ctrl, PictureBox):
                return ctrl
        return None

    def invalidate_plots(self):
        for container in (self.chart_delay, self.chart_offset, self.chart_jitter, self.chart_dispersion):
            plot_box = self._get_plot_box(container)
            if plot_box is not None:
                plot_box.Invalidate()

    def on_plot_paint(self, sender, event):
        plot_key = sender.Tag
        chart_data = self._plot_data.get(plot_key)
        if chart_data is None:
            self.draw_empty_plot(event.Graphics, sender.ClientRectangle)
            return
        chart_data["plot_key"] = plot_key  # Pass plot_key to draw_plot
        self.draw_plot(event.Graphics, sender.ClientRectangle, chart_data)

    def draw_empty_plot(self, graphics, bounds):
        graphics.Clear(Color.White)
        brush = SolidBrush(Color.Gray)
        try:
            graphics.DrawString("Run Analyze to draw data.", Font("Segoe UI", 9), brush, 8, 8)
        finally:
            brush.Dispose()

    def draw_plot(self, graphics, bounds, chart_data):
        graphics.Clear(Color.White)

        left = 62
        top = 8
        right = 6
        bottom = 22
        width = max(10, bounds.Width - left - right)
        height = max(10, bounds.Height - top - bottom)

        plot_rect = Rectangle(left, top, width, height)

        x_start = chart_data["x_start"]
        x_end = chart_data["x_end"]
        y_min = chart_data["y_min"]   # ms
        y_max = chart_data["y_max"]   # ms
        y_step = chart_data["y_step"] # ms
        
        plot_key = chart_data.get("plot_key", "")
        
        # For delay chart, build series from delay points with server coloring
        if plot_key == "delay":
            delay_points = chart_data.get("points", [])
            server_to_color = chart_data.get("server_to_color", {})
            unique_servers = chart_data.get("unique_servers", [])
            
            # Build series with server-colored segments
            series = [
                {
                    "name": "Delay (Server color-coded)",
                    "points": delay_points,  # (timestamp, delay, server) tuples
                    "server_to_color": server_to_color,
                    "colored": True,
                    "unique_servers": unique_servers,
                }
            ]
        else:
            series = chart_data.get("series", [])

        one_hour = timedelta(hours=1)
        h_grid_pen = Pen(Color.FromArgb(220, 220, 220))   # faint horizontal gridlines
        v_grid_pen = Pen(Color.FromArgb(228, 228, 228))   # faint vertical gridlines
        zero_pen = Pen(Color.FromArgb(150, 150, 150))     # zero reference
        axis_pen = Pen(Color.FromArgb(100, 100, 100))     # axis border
        label_brush = SolidBrush(Color.FromArgb(80, 80, 80))
        label_font = Font("Segoe UI", 7)
        try:
            # --- Vertical x-gridlines (hourly) ---
            hour = x_start
            while hour <= x_end:
                x = self.map_x(hour, x_start, x_end, plot_rect)
                graphics.DrawLine(v_grid_pen, x, plot_rect.Top, x, plot_rect.Bottom)
                if hour.hour % 2 == 0:
                    graphics.DrawString(hour.strftime("%H:%M"), label_font, label_brush, x - 14, plot_rect.Bottom + 2)
                hour = hour + one_hour

            # --- Horizontal y-gridlines and tick labels ---
            num_ticks = int(round((y_max - y_min) / y_step)) + 1
            for i in range(num_ticks):
                y_val = y_min + i * y_step
                py = self.map_y(y_val, y_min, y_max, plot_rect)
                if plot_rect.Top <= py <= plot_rect.Bottom:
                    is_zero = abs(y_val) < y_step * 1e-4
                    if is_zero:
                        graphics.DrawLine(zero_pen, plot_rect.Left, py, plot_rect.Right, py)
                    else:
                        graphics.DrawLine(h_grid_pen, plot_rect.Left, py, plot_rect.Right, py)
                    lbl = _format_y_label_ms(y_val, y_step)
                    lbl_y = max(plot_rect.Top, min(plot_rect.Bottom - 10, py - 6))
                    graphics.DrawString(lbl, label_font, label_brush, 2, lbl_y)

            # --- Axis border ---
            graphics.DrawRectangle(axis_pen, plot_rect)
            graphics.DrawString("UTC", label_font, label_brush, plot_rect.Right - 24, plot_rect.Bottom + 2)

            # --- Data series ---
            for item in series:
                points = item["points"]
                if len(points) < 1:
                    continue

                # Check if this is a server-colored series (for delay chart)
                if item.get("colored", False):
                    # Draw delay line with server-based color segments
                    server_to_color = item.get("server_to_color", {})
                    prev_xy = None
                    prev_server = None
                    
                    for point_data in points:
                        dt_value = point_data[0]
                        y_value = point_data[1]
                        server = point_data[2] if len(point_data) > 2 else ""
                        
                        x = self.map_x(dt_value, x_start, x_end, plot_rect)
                        y_ms = y_value * 1000.0
                        y = self.map_y(y_ms, y_min, y_max, plot_rect)
                        
                        if prev_xy is not None:
                            # Color each segment by the active source server at the
                            # previous point; color changes where source changes.
                            segment_color = server_to_color.get(prev_server, Color.Gray)
                            segment_pen = Pen(segment_color, 2)
                            try:
                                graphics.DrawLine(segment_pen, prev_xy[0], prev_xy[1], x, y)
                            finally:
                                segment_pen.Dispose()
                        prev_xy = (x, y)
                        prev_server = server
                else:
                    # Normal single-color series
                    line_pen = Pen(item["color"], item.get("width", 2))
                    try:
                        prev_xy = None
                        for dt_value, y_value in points:
                            x = self.map_x(dt_value, x_start, x_end, plot_rect)
                            y_ms = y_value * 1000.0
                            y = self.map_y(y_ms, y_min, y_max, plot_rect)
                            if prev_xy is not None:
                                graphics.DrawLine(line_pen, prev_xy[0], prev_xy[1], x, y)
                            prev_xy = (x, y)
                    finally:
                        line_pen.Dispose()

        finally:
            h_grid_pen.Dispose()
            v_grid_pen.Dispose()
            zero_pen.Dispose()
            axis_pen.Dispose()
            label_brush.Dispose()
            label_font.Dispose()

    def map_x(self, dt_value, x_start, x_end, rect):
        total = (x_end - x_start).total_seconds()
        if total <= 0:
            return rect.Left
        offset = (dt_value - x_start).total_seconds()
        return int(rect.Left + (float(offset) / float(total)) * rect.Width)

    def map_y(self, value, y_min, y_max, rect):
        span = y_max - y_min
        if span <= 0:
            return rect.Top + int(rect.Height / 2)
        ratio = (float(value) - float(y_min)) / float(span)
        return int(rect.Bottom - ratio * rect.Height)

    def update_charts(self, loop_rows, peer_rows, use_raw_peer_points):
        x_start, x_end = compute_axis_day_bounds(loop_rows, peer_rows)
        if x_start is None or x_end is None:
            self._plot_data = {}
            self.update_delay_header_legend({}, [])
            self.invalidate_plots()
            return

        selected_peer_rows, _note = select_peer_subset(peer_rows)

        offset_points = []
        loop_jitter_points = []
        for row in sorted(loop_rows, key=lambda value: (value.mjd, value.sec_of_day)):
            stamp = to_utc_datetime(row.mjd, row.sec_of_day)
            offset_points.append((stamp, row.offset))
            loop_jitter_points.append((stamp, row.jitter))

        peer_jitter_points = []
        dispersion_points = []
        server_to_color = {}  # Server address -> Color object

        # Build selected-peer timeline from peerstats for server coloring.
        # Reduce to one active selected server per second.
        selected_timeline_rows = reduce_to_active_timeline(selected_peer_rows)
        peer_timeline = []
        for row in selected_timeline_rows:
            stamp = to_utc_datetime(row.mjd, row.sec_of_day)
            server = row.server_address if hasattr(row, "server_address") else ""
            get_server_color(server, server_to_color)
            peer_timeline.append((stamp, server))

        # Dispersion is densified later on the loop timeline (same approach
        # as selected-server delay/jitter), so do not append sparse points here.

        # Delay chart is sampled on the loopstats timeline, but value is the
        # true peerstats delay of the active selected server.
        delay_points = []  # List of (timestamp, selected_server_delay, server) tuples

        # Build per-server raw delay/jitter/dispersion streams from full peer rows.
        delays_by_server = {}
        jitters_by_server = {}
        dispersions_by_server = {}
        for row in sorted(peer_rows, key=lambda value: (value.mjd, value.sec_of_day)):
            stamp = to_utc_datetime(row.mjd, row.sec_of_day)
            server = row.server_address if hasattr(row, "server_address") else ""
            if server not in delays_by_server:
                delays_by_server[server] = []
            if server not in jitters_by_server:
                jitters_by_server[server] = []
            if server not in dispersions_by_server:
                dispersions_by_server[server] = []
            delays_by_server[server].append((stamp, row.delay))
            jitters_by_server[server].append((stamp, row.jitter))
            dispersions_by_server[server].append((stamp, row.dispersion))

        delay_index_by_server = {}
        jitter_index_by_server = {}
        dispersion_index_by_server = {}
        for server in delays_by_server.keys():
            delay_index_by_server[server] = -1
        for server in jitters_by_server.keys():
            jitter_index_by_server[server] = -1
        for server in dispersions_by_server.keys():
            dispersion_index_by_server[server] = -1

        timeline_index = 0
        active_server = ""
        if peer_timeline:
            active_server = peer_timeline[0][1]

        for row in sorted(loop_rows, key=lambda value: (value.mjd, value.sec_of_day)):
            stamp = to_utc_datetime(row.mjd, row.sec_of_day)
            while timeline_index + 1 < len(peer_timeline) and peer_timeline[timeline_index + 1][0] <= stamp:
                timeline_index += 1
                active_server = peer_timeline[timeline_index][1]

            # Carry-forward latest delay from the active selected server.
            server_delays = delays_by_server.get(active_server, [])
            if server_delays:
                idx = delay_index_by_server.get(active_server, -1)
                while idx + 1 < len(server_delays) and server_delays[idx + 1][0] <= stamp:
                    idx += 1
                delay_index_by_server[active_server] = idx
                if idx >= 0:
                    delay_points.append((stamp, server_delays[idx][1], active_server))

            # Carry-forward latest jitter from the active selected server.
            server_jitters = jitters_by_server.get(active_server, [])
            if server_jitters:
                j_idx = jitter_index_by_server.get(active_server, -1)
                while j_idx + 1 < len(server_jitters) and server_jitters[j_idx + 1][0] <= stamp:
                    j_idx += 1
                jitter_index_by_server[active_server] = j_idx
                if j_idx >= 0:
                    peer_jitter_points.append((stamp, server_jitters[j_idx][1]))

            # Carry-forward latest dispersion from the active selected server.
            server_dispersions = dispersions_by_server.get(active_server, [])
            if server_dispersions:
                d_idx = dispersion_index_by_server.get(active_server, -1)
                while d_idx + 1 < len(server_dispersions) and server_dispersions[d_idx + 1][0] <= stamp:
                    d_idx += 1
                dispersion_index_by_server[active_server] = d_idx
                if d_idx >= 0:
                    dispersion_points.append((stamp, server_dispersions[d_idx][1]))

        if not server_to_color:
            get_server_color("", server_to_color)

        def y_limits_ms(series_list):
            """Compute y_min, y_max, y_step in ms for a list of point series (values in seconds).
            Always spans zero; snapped to a nice tick interval."""
            values_ms = []
            for points in series_list:
                values_ms.extend([v * 1000.0 for _, v in points])
            if not values_ms:
                return -1.0, 1.0, 1.0
            raw_min = min(values_ms)
            raw_max = max(values_ms)
            # Always bracket zero
            lo = min(raw_min, 0.0)
            hi = max(raw_max, 0.0)
            span = hi - lo
            if span == 0.0:
                lo -= 0.5
                hi += 0.5
                span = 1.0
            step = _choose_y_step_ms(span)
            # Snap outward to step boundaries
            y_min = math.floor(lo / step) * step
            y_max = math.ceil(hi / step) * step
            # Ensure at least 2 ticks of range
            if y_max - y_min < step * 2:
                y_max = y_min + step * 2
            return y_min, y_max, step

        offset_min, offset_max, offset_step = y_limits_ms([offset_points])
        jitter_min, jitter_max, jitter_step = y_limits_ms([loop_jitter_points, peer_jitter_points])
        disp_min, disp_max, disp_step = y_limits_ms([dispersion_points])
        # For delay points, extract just the values (second element of tuple)
        delay_values_only = [(dt, val) for dt, val, srv in delay_points]
        delay_min, delay_max, delay_step = y_limits_ms([delay_values_only])

        # Get unique selected servers for legend
        unique_servers = sorted(set([srv for _stamp, srv in peer_timeline]))
        if not unique_servers and delay_points:
            unique_servers = sorted(set([srv for _stamp, _value, srv in delay_points]))
        obs_lat, obs_lon = self._get_observer_coords()
        print("[legend debug] obs_lat=%r obs_lon=%r" % (obs_lat, obs_lon))
        print("[legend debug] known_servers count=%d" % len(self._known_servers))
        print("[legend debug] server_to_color keys=%r" % list(server_to_color.keys()))
        server_to_km = {}
        for _srv in list(server_to_color.keys()):
            if _srv:
                _loc = resolve_server_location(_srv, self._known_servers, obs_lat, obs_lon)
                print("[legend debug] server=%r -> d_min_s=%r geo_km=%r note=%r" % (
                    _srv, _loc["d_min_s"], _loc["geo_km"], _loc["location_note"]))
                if _loc["geo_km"] is not None:
                    server_to_km[_srv] = _loc["geo_km"]
        print("[legend debug] server_to_km=%r" % server_to_km)
        self.update_delay_header_legend(server_to_color, unique_servers, server_to_km)

        self._plot_data = {
            "delay": {
                "x_start": x_start,
                "x_end": x_end,
                "y_min": delay_min,
                "y_max": delay_max,
                "y_step": delay_step,
                "points": delay_points,  # List of (timestamp, delay, server) tuples
                "server_to_color": server_to_color,
                "unique_servers": unique_servers,
            },
            "offset": {
                "x_start": x_start,
                "x_end": x_end,
                "y_min": offset_min,
                "y_max": offset_max,
                "y_step": offset_step,
                "series": [
                    {"name": "Offset", "color": Color.FromArgb(31, 119, 180), "points": offset_points},
                ],
            },
            "jitter": {
                "x_start": x_start,
                "x_end": x_end,
                "y_min": jitter_min,
                "y_max": jitter_max,
                "y_step": jitter_step,
                "series": [
                    {
                        "name": "Loop Jitter",
                        "color": Color.FromArgb(255, 127, 14),
                        "width": 3,
                        "points": loop_jitter_points,
                    },
                    {
                        "name": "Peer Jitter",
                        "color": Color.FromArgb(44, 160, 44),
                        "width": 2,
                        "points": peer_jitter_points,
                    },
                ],
            },
            "dispersion": {
                "x_start": x_start,
                "x_end": x_end,
                "y_min": disp_min,
                "y_max": disp_max,
                "y_step": disp_step,
                "series": [
                    {"name": "Dispersion", "color": Color.FromArgb(214, 39, 40), "points": dispersion_points},
                ],
            },
        }

        self.invalidate_plots()

    def _get_observer_coords(self):
        """Parse observer lat/lon text boxes. Returns (lat, lon) as floats, or (None, None) if invalid."""
        try:
            lat = float(self.txt_observer_lat.Text.strip())
            lon = float(self.txt_observer_lon.Text.strip())
        except (ValueError, AttributeError):
            return None, None
        # Detect swapped entry: latitude must be -90..90, longitude -180..180.
        lat_valid = -90.0 <= lat <= 90.0
        lon_valid = -180.0 <= lon <= 180.0
        if not lat_valid and lon_valid and -90.0 <= lon <= 90.0 and -180.0 <= lat <= 180.0:
            print("[observer] WARNING: lat=%r lon=%r looks swapped — auto-correcting to lat=%r lon=%r" % (lat, lon, lon, lat))
            lat, lon = lon, lat
        elif not lat_valid or not lon_valid:
            print("[observer] WARNING: lat=%r lon=%r out of valid range — ignored" % (lat, lon))
            return None, None
        return lat, lon

    def prefill_defaults(self):
        saved = load_folder_settings()
        saved_log = saved.get("log_folder", "").strip()
        saved_export = saved.get("export_folder", "").strip()

        if saved_log and os.path.isdir(saved_log):
            self.txt_log_folder.Text = saved_log
            if saved_export:
                self.txt_export_folder.Text = saved_export
            else:
                self.txt_export_folder.Text = os.path.join(saved_log, "reports")
        else:
            candidates = discover_candidate_dirs()
            if candidates:
                self.txt_log_folder.Text = candidates[0]
                self.txt_export_folder.Text = os.path.join(candidates[0], "reports")

        if not self.txt_export_folder.Text.strip() and self.txt_log_folder.Text.strip():
            self.txt_export_folder.Text = os.path.join(self.txt_log_folder.Text.strip(), "reports")
        self.txt_observer_lat.Text = saved.get("observer_lat", "").strip()
        self.txt_observer_lon.Text = saved.get("observer_lon", "").strip()
        self.on_export_toggle(None, None)
        self.scan_options()

    def show_error(self, message):
        MessageBox.Show(self, message, "NTP Analyzer", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def choose_folder(self, current_path):
        dialog = FolderBrowserDialog()
        if current_path and os.path.isdir(current_path):
            dialog.SelectedPath = current_path
        if dialog.ShowDialog(self) == DialogResult.OK:
            return dialog.SelectedPath
        return None

    def on_browse_log(self, sender, event):
        chosen = self.choose_folder(self.txt_log_folder.Text.strip())
        if chosen:
            self.txt_log_folder.Text = chosen
            if not self.txt_export_folder.Text.strip():
                self.txt_export_folder.Text = os.path.join(chosen, "reports")
            save_folder_settings(self.txt_log_folder.Text.strip(), self.txt_export_folder.Text.strip(),
                                  self.txt_observer_lat.Text.strip(), self.txt_observer_lon.Text.strip())
            self.scan_options()

    def on_browse_export(self, sender, event):
        chosen = self.choose_folder(self.txt_export_folder.Text.strip())
        if chosen:
            self.txt_export_folder.Text = chosen
            save_folder_settings(self.txt_log_folder.Text.strip(), self.txt_export_folder.Text.strip(),
                                  self.txt_observer_lat.Text.strip(), self.txt_observer_lon.Text.strip())

    def on_export_toggle(self, sender, event):
        enabled = self.chk_export.Checked
        self.txt_export_folder.Enabled = enabled
        self.btn_browse_export.Enabled = enabled

    def on_scan(self, sender, event):
        save_folder_settings(self.txt_log_folder.Text.strip(), self.txt_export_folder.Text.strip(),
                              self.txt_observer_lat.Text.strip(), self.txt_observer_lon.Text.strip())
        self.scan_options()

    def scan_options(self):
        log_folder = self.txt_log_folder.Text.strip().strip('"')
        self.cmb_dataset.Items.Clear()
        self._options_by_label = {}

        if not log_folder:
            self.set_status("Set an NTP log folder, then scan datasets.")
            return

        if not os.path.isdir(log_folder):
            self.set_status("Log folder does not exist.")
            return

        try:
            options = build_day_options(log_folder)
        except Exception as error:
            self.show_error("Failed to scan datasets:\n%s" % str(error))
            self.set_status("Scan failed.")
            return

        filter_text = self.txt_day_filter.Text.strip().lower()
        if filter_text:
            options = [o for o in options if filter_text in o.key.lower() or filter_text in o.label.lower()]

        if not options:
            self.set_status("No matching loopstats/peerstats datasets found.")
            return

        for option in options:
            self.cmb_dataset.Items.Add(option.label)
            self._options_by_label[option.label] = option

        self.cmb_dataset.SelectedIndex = 0
        self.set_status("Loaded %d dataset option(s)." % len(options))

    def get_selected_option(self):
        selected_label = self.cmb_dataset.Text
        if not selected_label:
            raise RuntimeError("No dataset selected.")
        option = self._options_by_label.get(selected_label)
        if option is None:
            raise RuntimeError("Dataset selection is invalid. Please rescan datasets.")
        return option

    def on_analyze(self, sender, event):
        try:
            option = self.get_selected_option()
            loop_rows = parse_loopstats(option.loop_path, option.target_mjd)
            peer_rows = parse_peerstats(option.peer_path, option.target_mjd)
            obs_lat, obs_lon = self._get_observer_coords()
            result = analyze(option, loop_rows, peer_rows,
                             known_servers=self._known_servers,
                             observer_lat=obs_lat, observer_lon=obs_lon)
            self._last_loop_rows = loop_rows
            self._last_peer_rows = peer_rows
            self._last_result = result
            self._last_aggregate_report = generate_report(result)
            self.update_charts(loop_rows, peer_rows, self.chk_raw_peer_points.Checked)

            pit = self._compute_pit_for_display(loop_rows, peer_rows, result)
            self._show_combined_output(pit)

            save_folder_settings(self.txt_log_folder.Text.strip(), self.txt_export_folder.Text.strip(),
                                  self.txt_observer_lat.Text.strip(), self.txt_observer_lon.Text.strip())

            if self.chk_export.Checked:
                export_folder = self.txt_export_folder.Text.strip().strip('"')
                if not export_folder:
                    raise RuntimeError("Export folder is empty. Set an export folder or uncheck export.")
                json_path, csv_path = resolve_export_paths(export_folder, result)
                export_json(json_path, result)
                export_csv(csv_path, result)
                self.set_status("Analysis complete. Saved JSON: %s | CSV: %s" % (json_path, csv_path))
            else:
                self.set_status("Analysis complete.")

        except Exception as error:
            self.show_error(str(error))
            self.set_status("Analysis failed.")

    def _compute_pit_for_display(self, loop_rows, peer_rows, result):
        """Return pit_result using the user-entered HH:MM:SS time if valid, else the last loopstats record."""
        hms_text = self.txt_pit_time.Text.strip()
        if hms_text:
            try:
                query_sec = _parse_pit_time_sec(hms_text)
                query_mjd = max(result.mjds) if result.mjds else max(r.mjd for r in loop_rows)
                obs_lat, obs_lon = self._get_observer_coords()
                return estimate_offset_at_time(query_mjd, query_sec, loop_rows, peer_rows,
                                               known_servers=self._known_servers,
                                               observer_lat=obs_lat, observer_lon=obs_lon)
            except Exception as err:
                self.show_error("Invalid point-in-time entry: %s\r\nUsing last loopstats record." % str(err))
        return result.pit_result

    def _update_pit_result_display(self, pit):
        """Update the read-only summary box with the best-estimate offset and error."""
        if pit is None:
            self.txt_pit_result.Text = ""
            return
        alt_exp = pit.get("alt_u_expanded")
        primary_exp = pit["u_expanded"]
        if alt_exp is not None and alt_exp < primary_exp:
            offset_ms = pit["alt_best_offset"] * 1000.0
            error_ms = alt_exp * 1000.0
        else:
            offset_ms = pit["best_offset"] * 1000.0
            error_ms = primary_exp * 1000.0
        self.txt_pit_result.Text = "Offset: %.1f ms; Error: %.1f ms" % (offset_ms, error_ms)

    def _show_combined_output(self, pit):
        """Compose and display PIT section (top) then the aggregate report (below)."""
        if pit is not None:
            pit_text = "\r\n".join(format_pit_section(pit))
        else:
            pit_text = "(No point-in-time estimate available.)"
        separator = "\r\n" + ("=" * 80) + "\r\n\r\n"
        self.txt_output.Text = pit_text + separator + self._last_aggregate_report
        self._update_pit_result_display(pit)

    def on_pit_calculate(self, sender, event):
        try:
            if not self._last_loop_rows:
                self.show_error("Run Analyze first to load a dataset.")
                return
            hms_text = self.txt_pit_time.Text.strip()
            if not hms_text:
                self.show_error("Enter a time in HH:MM:SS format.")
                return
            query_sec = _parse_pit_time_sec(hms_text)
            query_mjd = max(self._last_result.mjds) if self._last_result and self._last_result.mjds else max(r.mjd for r in self._last_loop_rows)
            obs_lat, obs_lon = self._get_observer_coords()
            pit = estimate_offset_at_time(query_mjd, query_sec, self._last_loop_rows, self._last_peer_rows,
                                          known_servers=self._known_servers,
                                          observer_lat=obs_lat, observer_lon=obs_lon)
            self._show_combined_output(pit)
            self._update_pit_result_display(pit)
            self.set_status("Point-in-time estimate calculated for %s." % hms_text)
        except Exception as error:
            self.show_error(str(error))
            self.set_status("Point-in-time calculation failed.")


def main():
    if clr is None:
        sys.stderr.write(
            "This script requires IronPython 3.4 on Windows (clr/System.Windows.Forms not available).\n"
        )
        return 1

    if sys.version_info[0] != 3 or sys.version_info[1] < 4:
        sys.stderr.write(
            "This script targets IronPython 3.4+. Current Python: %d.%d\n"
            % (sys.version_info[0], sys.version_info[1])
        )
        return 1

    Application.EnableVisualStyles()
    Application.Run(AnalyzerForm())
    return 0


if __name__ == "__main__":
    main()
