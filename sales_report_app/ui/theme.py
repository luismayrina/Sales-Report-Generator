"""
ui/theme.py — Theme manager for light and dark mode.
"""
import json
import os

# ── Preference storage ────────────────────────────────────────────────────────
_SETTINGS_PATH = os.path.join(os.path.dirname(__file__), "..", "app_settings.json")

def load_theme_preference() -> str:
    """Return 'light' or 'dark'. Defaults to 'light'."""
    try:
        with open(_SETTINGS_PATH, "r") as f:
            data = json.load(f)
            return data.get("theme", "light")
    except (FileNotFoundError, json.JSONDecodeError):
        return "light"

def save_theme_preference(mode: str):
    """Persist theme preference ('light' or 'dark') to disk."""
    data = {}
    try:
        with open(_SETTINGS_PATH, "r") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pass
    data["theme"] = mode
    with open(_SETTINGS_PATH, "w") as f:
        json.dump(data, f, indent=2)


# ── Stylesheets ───────────────────────────────────────────────────────────────
def get_light_stylesheet() -> str:
    return """
/* ── Global ─────────────────────────────────────────────────────────────── */
QMainWindow, QDialog {
    background-color: #F6F7F9;
}
QWidget {
    color: #111827;
    font-size: 13px;
}
QScrollArea {
    border: none;
    background-color: transparent;
}
QScrollArea > QWidget > QWidget {
    background-color: transparent;
}

/* ── Header ──────────────────────────────────────────────────────────────── */
#header {
    background-color: #FFFFFF;
    border-bottom: 1px solid #E5E7EB;
}
#headerTitle {
    font-size: 24px;
    font-weight: bold;
    color: #111827;
}
#headerSubtitle {
    font-size: 13px;
    color: #6B7280;
}

/* ── Cards ───────────────────────────────────────────────────────────────── */
#card {
    background-color: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 10px;
}
#cardTitle {
    font-size: 15px;
    font-weight: bold;
    color: #111827;
}
#cardSubtitle {
    font-size: 12px;
    color: #6B7280;
}

/* ── Upload Slots ────────────────────────────────────────────────────────── */
#uploadSlot {
    background-color: #F9FAFB;
    border: 1.5px dashed #D1D5DB;
    border-radius: 8px;
}
#uploadSlot:hover {
    background-color: #EFF6FF;
    border-color: #3B82F6;
}
#uploadSlotLabel {
    color: #374151;
    font-weight: 600;
}
#uploadSlotHint {
    color: #9CA3AF;
    font-size: 11px;
}

/* ── Status badges ───────────────────────────────────────────────────────── */
#statusReady {
    color: #16A34A;
    font-weight: bold;
}
#statusMissing {
    color: #DC2626;
    font-weight: bold;
}
#statusWarning {
    color: #D97706;
    font-weight: bold;
}

/* ── Progress panel / log ────────────────────────────────────────────────── */
#progressPanel {
    background-color: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 10px;
}
#progressLog {
    background-color: #F9FAFB;
    color: #374151;
    border: 1px solid #E5E7EB;
    border-radius: 6px;
}

/* ── Inputs ──────────────────────────────────────────────────────────────── */
QLineEdit, QTextEdit {
    background-color: #FFFFFF;
    color: #111827;
    border: 1px solid #D1D5DB;
    border-radius: 6px;
    padding: 4px 8px;
}
QLineEdit:focus, QTextEdit:focus {
    border-color: #2563EB;
}

/* ── Checkboxes ──────────────────────────────────────────────────────────── */
QCheckBox {
    color: #374151;
}
QCheckBox::indicator:checked {
    background-color: #2563EB;
    border: 1px solid #2563EB;
    border-radius: 3px;
}

/* ── Buttons ─────────────────────────────────────────────────────────────── */
#themeToggle {
    background-color: #F3F4F6;
    color: #374151;
    border: 1px solid #D1D5DB;
    border-radius: 20px;
    padding: 5px 14px;
    font-size: 12px;
    font-weight: 600;
}
#themeToggle:hover {
    background-color: #E5E7EB;
}

/* ── Progress bar ────────────────────────────────────────────────────────── */
QProgressBar {
    background-color: #E5E7EB;
    border-radius: 4px;
    height: 6px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #2563EB;
    border-radius: 4px;
}
"""


def get_dark_stylesheet() -> str:
    return """
/* ── Global ─────────────────────────────────────────────────────────────── */
QMainWindow, QDialog {
    background-color: #0F172A;
}
QWidget {
    color: #F9FAFB;
    font-size: 13px;
}
QScrollArea {
    border: none;
    background-color: transparent;
}
QScrollArea > QWidget > QWidget {
    background-color: transparent;
}

/* ── Header ──────────────────────────────────────────────────────────────── */
#header {
    background-color: #111827;
    border-bottom: 1px solid #374151;
}
#headerTitle {
    font-size: 24px;
    font-weight: bold;
    color: #F9FAFB;
}
#headerSubtitle {
    font-size: 13px;
    color: #9CA3AF;
}

/* ── Cards ───────────────────────────────────────────────────────────────── */
#card {
    background-color: #111827;
    border: 1px solid #374151;
    border-radius: 10px;
}
#cardTitle {
    font-size: 15px;
    font-weight: bold;
    color: #F9FAFB;
}
#cardSubtitle {
    font-size: 12px;
    color: #9CA3AF;
}

/* ── Upload Slots ────────────────────────────────────────────────────────── */
#uploadSlot {
    background-color: #1F2937;
    border: 1.5px dashed #4B5563;
    border-radius: 8px;
}
#uploadSlot:hover {
    background-color: #1E3A5F;
    border-color: #3B82F6;
}
#uploadSlotLabel {
    color: #E5E7EB;
    font-weight: 600;
}
#uploadSlotHint {
    color: #6B7280;
    font-size: 11px;
}

/* ── Status badges ───────────────────────────────────────────────────────── */
#statusReady {
    color: #22C55E;
    font-weight: bold;
}
#statusMissing {
    color: #EF4444;
    font-weight: bold;
}
#statusWarning {
    color: #F59E0B;
    font-weight: bold;
}

/* ── Progress panel / log ────────────────────────────────────────────────── */
#progressPanel {
    background-color: #111827;
    border: 1px solid #374151;
    border-radius: 10px;
}
#progressLog {
    background-color: #0F172A;
    color: #D1D5DB;
    border: 1px solid #374151;
    border-radius: 6px;
}

/* ── Inputs ──────────────────────────────────────────────────────────────── */
QLineEdit, QTextEdit {
    background-color: #1F2937;
    color: #F9FAFB;
    border: 1px solid #4B5563;
    border-radius: 6px;
    padding: 4px 8px;
}
QLineEdit:focus, QTextEdit:focus {
    border-color: #3B82F6;
}

/* ── Checkboxes ──────────────────────────────────────────────────────────── */
QCheckBox {
    color: #D1D5DB;
}
QCheckBox::indicator {
    background-color: #1F2937;
    border: 1px solid #4B5563;
    border-radius: 3px;
}
QCheckBox::indicator:checked {
    background-color: #3B82F6;
    border: 1px solid #3B82F6;
    border-radius: 3px;
}

/* ── Buttons ─────────────────────────────────────────────────────────────── */
#themeToggle {
    background-color: #1F2937;
    color: #D1D5DB;
    border: 1px solid #4B5563;
    border-radius: 20px;
    padding: 5px 14px;
    font-size: 12px;
    font-weight: 600;
}
#themeToggle:hover {
    background-color: #374151;
}

/* ── Progress bar ────────────────────────────────────────────────────────── */
QProgressBar {
    background-color: #374151;
    border-radius: 4px;
    height: 6px;
    text-align: center;
}
QProgressBar::chunk {
    background-color: #3B82F6;
    border-radius: 4px;
}
"""


def get_stylesheet(mode: str) -> str:
    """Return the full stylesheet for the given mode ('light' or 'dark')."""
    if mode == "dark":
        return get_dark_stylesheet()
    return get_light_stylesheet()
