import base64
import codecs
import csv
import hashlib
import io
import re
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


st.set_page_config(page_title="FileSort Cleaner", layout="wide", page_icon="🔥")

THEME_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --bg: #000000;
    --surface: #0c0c10;
    --surface-2: #14141a;
    --surface-3: #1d1d26;
    --line: rgba(255, 255, 255, 0.08);
    --line-strong: rgba(255, 255, 255, 0.14);
    --text: #f5f5f7;
    --muted: #9b9ba8;
    --accent: #ff7a1a;
    --accent-2: #ffb347;
    --accent-soft: rgba(255, 122, 26, 0.14);
    --good: #3ddc97;
    --warn: #ffb347;
    --radius: 16px;
}

[data-testid="stToolbar"], [data-testid="stDecoration"], header { display: none !important; }
#MainMenu, footer { visibility: hidden; }

.stApp {
    background:
        radial-gradient(900px 520px at 0% 0%, rgba(255, 122, 26, 0.09), transparent 60%),
        var(--bg);
    color: var(--text);
    font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif;
}
.block-container { position: relative; z-index: 2; padding-top: 1.4rem !important; max-width: 1180px; }

/* ---------- Embers ---------- */
.ember-field { position: fixed; inset: 0; pointer-events: none; z-index: 1; overflow: hidden; }
.ember {
    position: absolute;
    bottom: -20px;
    border-radius: 50%;
    background: radial-gradient(circle, #fff3d6 0%, var(--accent) 50%, transparent 72%);
    box-shadow: 0 0 10px 2px rgba(255, 122, 26, 0.45);
    opacity: 0;
    animation: emberRise linear infinite;
}
@keyframes emberRise {
    0%   { transform: translate3d(0, 0, 0) scale(0.6); opacity: 0; }
    15%  { opacity: 0.7; }
    80%  { opacity: 0.35; }
    100% { transform: translate3d(var(--drift, 30px), -100vh, 0) scale(1); opacity: 0; }
}

/* ---------- Typography ---------- */
p, label, span, li { color: #d8d8e0; }
h1, h2, h3, h4 {
    font-family: 'Inter', sans-serif !important;
    color: var(--text) !important;
    letter-spacing: -0.01em;
    font-weight: 700 !important;
}
[data-testid="stCaptionContainer"] p { color: var(--muted) !important; }

@keyframes fadeUp {
    from { opacity: 0; transform: translateY(10px); }
    to   { opacity: 1; transform: none; }
}

/* ---------- Workspace header ---------- */
.app-bar {
    display: flex;
    align-items: flex-end;
    justify-content: space-between;
    gap: 16px;
    padding: 4px 0 16px;
    border-bottom: 1px solid var(--line);
    margin-bottom: 18px;
    animation: fadeUp 0.5s ease both;
}
.app-kicker {
    font-size: 0.7rem;
    letter-spacing: 0.3em;
    text-transform: uppercase;
    color: var(--muted);
    font-weight: 600;
    display: flex;
    align-items: center;
    gap: 8px;
}
.app-kicker i {
    width: 7px; height: 7px; border-radius: 50%;
    background: var(--accent);
    box-shadow: 0 0 10px var(--accent);
}
.app-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.6rem;
    letter-spacing: 0.05em;
    line-height: 1;
    color: var(--text);
    margin-top: 6px;
}
.app-title span, .landing-title span {
    background: linear-gradient(180deg, #fff3d6 0%, var(--accent-2) 45%, var(--accent) 100%);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
}
.app-files { color: var(--muted); font-size: 0.85rem; text-align: right; }
.app-files b { color: var(--text); font-weight: 600; }

/* ---------- Stat cards (count-up) ---------- */
@property --n { syntax: '<integer>'; initial-value: 0; inherits: false; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin: 0 0 18px; }
.stat {
    position: relative;
    overflow: hidden;
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: var(--radius);
    padding: 14px 18px 12px;
    transition: transform 0.2s ease, border-color 0.2s ease;
    animation: fadeUp 0.5s ease both;
}
.stat:nth-child(2) { animation-delay: 0.06s; }
.stat:nth-child(3) { animation-delay: 0.12s; }
.stat:nth-child(4) { animation-delay: 0.18s; }
.stat:hover { transform: translateY(-2px); border-color: rgba(255, 122, 26, 0.4); }
.stat::after {
    content: "";
    position: absolute;
    right: -40px; bottom: -60px;
    width: 140px; height: 140px;
    background: radial-gradient(circle, rgba(255, 122, 26, 0.16), transparent 70%);
    pointer-events: none;
}
.stat-label { font-size: 0.7rem; letter-spacing: 0.14em; text-transform: uppercase; color: var(--muted); font-weight: 600; }
.stat-num {
    font-family: 'Bebas Neue', sans-serif;
    font-size: 2.4rem;
    line-height: 1.05;
    color: var(--text);
    counter-reset: n var(--n);
    animation: countUp 1.1s cubic-bezier(0.2, 0.8, 0.2, 1) forwards;
}
.stat-num::after { content: counter(n); }
@keyframes countUp { from { --n: 0; } to { --n: var(--to); } }

/* ---------- Step tabs ---------- */
.stTabs [data-baseweb="tab-list"] {
    counter-reset: step;
    gap: 8px;
    padding: 8px 0 12px;
    border-bottom: 1px solid var(--line);
    background: transparent;
}
.stTabs [data-baseweb="tab"] {
    counter-increment: step;
    height: auto;
    padding: 7px 16px 7px 7px !important;
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: 999px;
    transition: border-color 0.2s ease, background 0.2s ease, box-shadow 0.2s ease;
}
.stTabs [data-baseweb="tab"]::before {
    content: counter(step);
    display: inline-grid;
    place-items: center;
    width: 24px; height: 24px;
    margin-right: 10px;
    border-radius: 50%;
    background: var(--surface-3);
    color: var(--muted);
    font-weight: 700;
    font-size: 0.78rem;
    transition: background 0.2s ease, color 0.2s ease;
}
.stTabs [data-baseweb="tab"] p { font-weight: 600; color: var(--muted); font-size: 0.9rem; }
.stTabs [data-baseweb="tab"]:hover { border-color: var(--line-strong); }
.stTabs [data-baseweb="tab"][aria-selected="true"] {
    border-color: rgba(255, 122, 26, 0.55);
    background: linear-gradient(135deg, rgba(255, 122, 26, 0.16), rgba(255, 122, 26, 0.03));
    box-shadow: 0 0 24px rgba(255, 122, 26, 0.14);
}
.stTabs [data-baseweb="tab"][aria-selected="true"]::before {
    background: linear-gradient(135deg, var(--accent), var(--accent-2));
    color: #1a0a00;
}
.stTabs [data-baseweb="tab"][aria-selected="true"] p { color: var(--text); }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }
.stTabs [data-baseweb="tab-panel"] { padding-top: 18px; animation: fadeUp 0.4s ease both; }

/* Nested tabs (Map values): segmented switch, no step numbers. */
.stTabs .stTabs [data-baseweb="tab-list"] {
    display: inline-flex;
    gap: 4px;
    padding: 4px;
    border: 1px solid var(--line);
    border-radius: 14px;
    background: var(--surface);
}
.stTabs .stTabs [data-baseweb="tab"] {
    padding: 8px 22px !important;
    border: none;
    border-radius: 10px;
    background: transparent;
}
.stTabs .stTabs [data-baseweb="tab"]::before { display: none; }
.stTabs .stTabs [data-baseweb="tab"] p { font-size: 0.95rem; }
.stTabs .stTabs [data-baseweb="tab"][aria-selected="true"] {
    background: linear-gradient(135deg, rgba(255, 122, 26, 0.22), rgba(255, 179, 71, 0.10));
    box-shadow: inset 0 0 0 1px rgba(255, 122, 26, 0.5);
}
.stTabs .stTabs [data-baseweb="tab-panel"] {
    margin-top: 14px;
    padding: 18px 20px 8px;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    background: var(--surface);
}

.val-summary { color: var(--muted); font-size: 0.92rem; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.val-summary b { color: var(--text); }
.val-renamed {
    margin-left: 10px;
    padding: 3px 10px;
    border-radius: 999px;
    background: var(--surface-3);
    color: var(--muted);
    font-size: 0.78rem;
    font-weight: 600;
}
.val-renamed.on { background: rgba(255, 122, 26, 0.16); color: var(--accent-2); }


/* ---------- Buttons ---------- */
.stButton > button, .stDownloadButton > button {
    position: relative;
    overflow: hidden;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    color: var(--text) !important;
    background: var(--surface-2) !important;
    border: 1px solid var(--line-strong) !important;
    border-radius: 12px !important;
    padding: 0.55rem 1.1rem !important;
    transition: transform 0.15s ease, border-color 0.2s ease, box-shadow 0.2s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    border-color: rgba(255, 122, 26, 0.6) !important;
    box-shadow: 0 8px 24px rgba(255, 122, 26, 0.12);
}
.stButton > button::after, .stDownloadButton > button::after {
    content: "";
    position: absolute;
    top: 0; left: -120%;
    width: 50%; height: 100%;
    background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.14), transparent);
    transform: skewX(-20deg);
    transition: left 0.6s ease;
}
.stButton > button:hover::after, .stDownloadButton > button:hover::after { left: 140%; }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--accent), var(--accent-2)) !important;
    border: none !important;
    color: #1a0a00 !important;
    font-weight: 700 !important;
    box-shadow: 0 8px 28px rgba(255, 122, 26, 0.28);
}
.stButton > button[kind="primary"] p, .stDownloadButton > button[kind="primary"] p { color: #1a0a00 !important; }
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 12px 36px rgba(255, 122, 26, 0.42);
}

/* ---------- Uploader (workspace: compact file strip) ---------- */
[data-testid="stFileUploader"] {
    background: var(--surface);
    border: 1px dashed var(--line-strong);
    border-radius: var(--radius);
    padding: 10px 14px;
    transition: border-color 0.2s ease;
}
[data-testid="stFileUploader"]:hover { border-color: rgba(255, 122, 26, 0.55); }
[data-testid="stFileUploaderDropzone"] { background: transparent !important; }
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stFileUploaderDropzoneInstructions"] div { color: var(--muted) !important; }
[data-testid="stFileUploaderDropzoneInstructions"] svg { fill: var(--accent) !important; }
[data-testid="stFileUploader"] button {
    background: linear-gradient(135deg, var(--accent), var(--accent-2)) !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    box-shadow: 0 6px 20px rgba(255, 122, 26, 0.3);
}
[data-testid="stFileUploader"] button p,
[data-testid="stFileUploader"] button span,
[data-testid="stFileUploader"] button div { color: #1a0a00 !important; }
[data-testid="stFileUploaderFile"] [data-testid="stFileUploaderDeleteBtn"] button {
    background: transparent !important;
    box-shadow: none;
}

/* ---------- Panels ---------- */
[data-testid="stAlert"] {
    border-radius: 12px;
    background: var(--surface) !important;
    border: 1px solid var(--line);
    border-left: 3px solid var(--accent);
    color: var(--text) !important;
}
[data-testid="stAlert"] p { color: var(--text) !important; }

[data-testid="stExpander"] {
    border: 1px solid var(--line) !important;
    border-radius: 14px !important;
    background: var(--surface) !important;
}
[data-testid="stVerticalBlockBorderWrapper"], [data-testid="stContainer"] { border-radius: 14px; }
div[data-testid="stVerticalBlockBorderWrapper"]:has(> div > [data-testid="stVerticalBlock"]) { border-color: var(--line) !important; }

[data-testid="stDataFrame"], [data-testid="stTable"] {
    border: 1px solid var(--line);
    border-radius: 14px;
    overflow: hidden;
}
[data-testid="stDataFrame"] {
    --gdg-bg-cell: #0c0c10;
    --gdg-bg-cell-medium: #111117;
    --gdg-bg-header: #14141a;
    --gdg-bg-header-hovered: #1d1d26;
    --gdg-bg-header-has-focus: #1d1d26;
    --gdg-text-dark: #f5f5f7;
    --gdg-text-medium: #b8b8c4;
    --gdg-text-light: #8a8a98;
    --gdg-text-header: #ffcc99;
    --gdg-text-header-selected: #ffffff;
    --gdg-border-color: rgba(255, 255, 255, 0.06);
    --gdg-horizontal-border-color: rgba(255, 255, 255, 0.05);
    --gdg-accent-color: var(--accent);
    --gdg-accent-light: rgba(255, 122, 26, 0.14);
    --gdg-accent-fg: #1a0a00;
    --gdg-bg-bubble: #14141a;
    --gdg-bg-bubble-selected: #1d1d26;
    --gdg-bg-search-result: rgba(255, 179, 71, 0.3);
}

[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: 14px;
    padding: 12px 16px;
}

/* ---------- Inputs ---------- */
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div,
.stTextInput input, .stDateInput input {
    background: var(--surface-2) !important;
    border: 1px solid var(--line) !important;
    border-radius: 10px !important;
    color: var(--text) !important;
    transition: border-color 0.2s ease, box-shadow 0.2s ease;
}
.stSelectbox div[data-baseweb="select"] > div:hover,
.stTextInput input:hover { border-color: var(--line-strong) !important; }
.stTextInput input:focus, .stDateInput input:focus {
    border-color: rgba(255, 122, 26, 0.7) !important;
    box-shadow: 0 0 0 3px rgba(255, 122, 26, 0.15) !important;
}
[data-testid="stProgress"] > div > div > div > div {
    background: linear-gradient(90deg, var(--accent), var(--accent-2)) !important;
}

/* ---------- Mapping health ---------- */
.health {
    display: flex;
    align-items: center;
    gap: 14px;
    margin: 2px 0 14px;
}
.health-bar { flex: 1; height: 6px; border-radius: 6px; background: var(--surface-3); overflow: hidden; }
.health-bar i {
    display: block;
    height: 100%;
    border-radius: 6px;
    background: linear-gradient(90deg, var(--accent), var(--accent-2));
    transform-origin: left;
    animation: growX 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
@keyframes growX { from { transform: scaleX(0); } to { transform: scaleX(1); } }
.health-text { color: var(--text); font-weight: 600; font-size: 0.9rem; white-space: nowrap; }

/* ---------- Value mapping ---------- */
.val-src {
    display: inline-block;
    margin: 10px 0 4px;
    padding: 4px 12px;
    border-radius: 999px;
    background: var(--surface-3);
    color: var(--text);
    font-size: 0.78rem;
    font-weight: 600;
}

/* ---------- Error legend / results ---------- */
.err-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; margin: 6px 0 16px; }
.err-card {
    display: flex;
    align-items: center;
    gap: 12px;
    background: var(--surface);
    border: 1px solid var(--line);
    border-radius: 12px;
    padding: 10px 14px;
    animation: fadeUp 0.45s ease both;
}
.err-dot { width: 12px; height: 12px; border-radius: 4px; flex: none; }
.err-name { flex: 1; color: var(--muted); font-size: 0.85rem; font-weight: 600; }
.err-count { font-family: 'Bebas Neue', sans-serif; font-size: 1.6rem; line-height: 1; color: var(--text); }

/* ---------- Upload transition flash ---------- */
.ki-flash {
    position: fixed;
    inset: 0;
    z-index: 999;
    pointer-events: none;
    background: radial-gradient(circle at 18% 60%, rgba(255, 243, 214, 0.95), rgba(255, 179, 71, 0.5) 25%, rgba(255, 122, 26, 0.15) 50%, transparent 75%);
    animation: kiFlash 0.9s ease-out forwards;
}
@keyframes kiFlash {
    0%   { opacity: 0; }
    15%  { opacity: 1; }
    100% { opacity: 0; visibility: hidden; }
}

::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: #050507; }
::-webkit-scrollbar-thumb { background: #2a2a34; border-radius: 8px; border: 2px solid #050507; }
::-webkit-scrollbar-thumb:hover { background: var(--accent); }

@media (max-width: 900px) {
    .stats { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .err-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after { animation: none !important; transition: none !important; }
    .stat-num { --n: var(--to); }
    .ki-flash { display: none; }
}
</style>
"""

_EMBERS = "".join(
    f'<span class="ember" style="left:{left}%; --drift:{drift}px; '
    f'width:{size}px; height:{size}px; '
    f'animation-duration:{duration}s; animation-delay:{delay}s;"></span>'
    for left, drift, size, duration, delay in [
        (4, 60, 5, 13, 0.0), (11, -40, 8, 17, 2.4), (19, 30, 4, 11, 5.1),
        (27, 70, 7, 19, 1.2), (35, -55, 5, 14, 6.8), (43, 25, 9, 21, 3.6),
        (51, -30, 4, 12, 8.0), (58, 65, 6, 16, 0.8), (66, -45, 8, 20, 4.4),
        (73, 35, 5, 13, 7.2), (81, -60, 7, 18, 2.0), (88, 40, 4, 15, 5.6),
        (95, -25, 6, 22, 9.0),
    ]
)

st.markdown(THEME_CSS, unsafe_allow_html=True)
st.markdown(
    f'<div class="ember-field">{_EMBERS}</div><div class="fire-floor"></div>',
    unsafe_allow_html=True,
)


CANONICAL_FIELDS = [
    "full_name",
    "first_name",
    "last_name",
    "gender",
    "phone",
    "dob",
    "category",
    "tshirt_size",
    "blood_group",
    "address",
    "city",
    "state",
    "country",
    "emergency_name",
    "emergency_phone",
    "emergency_relation",
]

FIELD_LABELS = {
    "full_name": "Full Name",
    "first_name": "First Name",
    "last_name": "Last Name",
    "gender": "Gender",
    "phone": "Phone",
    "dob": "Date of Birth",
    "category": "Category",
    "tshirt_size": "T-Shirt Size",
    "blood_group": "Blood Group",
    "address": "Address",
    "city": "City",
    "state": "State",
    "country": "Country",
    "emergency_name": "Emergency Contact Name",
    "emergency_phone": "Emergency Contact Phone",
    "emergency_relation": "Emergency Contact Relation",
}

ESSENTIAL_MAPPING_FIELDS = [
    "full_name",
    "first_name",
    "last_name",
    "gender",
    "phone",
    "dob",
    "category",
    "tshirt_size",
]

ALIASES = {
    "full_name": ["attendee name", "runner name", "full name", "participant name", "name"],
    "first_name": ["first name", "first_name", "firstname", "fname"],
    "last_name": ["last name", "last_name", "lastname", "lname", "surname"],
    "gender": ["gender", "sex"],
    "phone": [
        "contact number",
        "mobile",
        "mobile no",
        "phone",
        "telephone",
        "mobile number",
        "contact no",
    ],
    "dob": [
        "date of birth",
        "dob",
        "birth date",
        "birthdate",
        "birth_date",
        "date_of_birth",
    ],
    "category": [
        "ticket_name",
        "race category",
        "racecategory",
        "event category",
        "distance",
        "registration category",
        "category",
    ],
    "tshirt_size": [
        "t shirt size",
        "t-shirt size",
        "tshirt size",
        "t-shirt",
        "tshirt",
        "shirt size",
    ],
    "blood_group": ["blood group", "bloodgroup"],
    "address": ["address"],
    "city": ["city"],
    "state": ["state", "state (india)"],
    "country": ["country"],
    "emergency_name": [
        "emergency contact name",
        "emergency name",
        "emergency contact person",
    ],
    "emergency_phone": [
        "emergency contact number",
        "emergency phone",
        "emergency number",
        "emergency no",
        "emergency contact no",
    ],
    "emergency_relation": ["emergency contact relation", "emergency relation", "relationship"],
}

KNOWN_TITLES = {
    "mr",
    "mr.",
    "mrs",
    "mrs.",
    "ms",
    "ms.",
    "dr",
    "dr.",
    "col",
    "col.",
    "colonel",
    "lt",
    "lt.",
    "lt col",
    "capt",
    "capt.",
    "prof",
    "prof.",
}

TSHIRT_MAP = {
    "xxs": "XXS",
    "xs": "XS",
    "s": "S",
    "small": "S",
    "m": "M",
    "medium": "M",
    "l": "L",
    "large": "L",
    "xl": "XL",
    "x-large": "XL",
    "xxl": "XXL",
    "xx-large": "XXL",
    "xxxl": "XXXL",
}

NUM_TO_SIZE = {
    "34": "XXS",
    "36": "XS",
    "38": "S",
    "40": "M",
    "42": "L",
    "44": "XL",
    "46": "XXL",
    "48": "XXXL",
}

VALID_BLOOD_GROUPS = {"A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-"}

COUNTRY_MAP = {
    "india": "India",
    "ind": "India",
    "in": "India",
    "usa": "United States",
    "us": "United States",
    "u.s.a.": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "uae": "United Arab Emirates",
}

EMERGENCY_RELATION_MAP = {
    "wife": "Wife",
    "husband": "Husband",
    "father": "Father",
    "mother": "Mother",
    "brother": "Brother",
    "sister": "Sister",
    "son": "Son",
    "daughter": "Daughter",
    "spouse": "Spouse",
    "friend": "Friend",
    "colleague": "Colleague",
    "guardian": "Guardian",
    "mentor": "Mentor",
    "uncle": "Uncle",
    "aunt": "Aunt",
}

ERROR_LEGEND = {
    "phone": {"hex": "FF9800", "label": "Phone Error"},
    "dob": {"hex": "FF5252", "label": "Date of Birth Error"},
    "name": {"hex": "448AFF", "label": "Name Issue"},
    "gender": {"hex": "E040FB", "label": "Gender Error"},
    "relation": {"hex": "00E676", "label": "Relation Error"},
    "country": {"hex": "FFEB3B", "label": "Country Error"},
}

COLUMN_FLAG_TO_ERROR = {
    "Phone": {"yellow": "phone"},
    "Emergency Phone": {"yellow": "phone"},
    "Gender": {"yellow": "gender"},
    "Date of Birth": {"red": "dob", "yellow": "dob"},
    "Country": {"yellow": "country"},
    "Emergency Relation": {"yellow": "relation"},
}

EMPTY_FIELD_ERROR = {
    "Phone": "phone",
    "Emergency Phone": "phone",
    "Gender": "gender",
    "Date of Birth": "dob",
    "Emergency Relation": "relation",
}

NAME_OUTPUT_COLUMNS = ["Full Name", "First Name", "Last Name"]
EMERGENCY_NAME_OUTPUT_COLUMNS = [
    "Emergency Full Name",
    "Emergency First Name",
    "Emergency Last Name",
]


_INVISIBLE_CHARS = re.compile(
    r"[\ufeff\u200b\u200c\u200d\u2060\u00ad"
    r"\u200e\u200f\u202a-\u202e"
    r"\u2061-\u2064\u2066-\u2069"
    r"\ufff9-\ufffb]+"
)
_UNICODE_WHITESPACE = re.compile(r"[\s\xa0\u1680\u180e\u2000-\u200a\u202f\u205f\u3000\v\f]+")


def to_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def clean_spaces(value) -> str:
    text = to_text(value)
    text = _INVISIBLE_CHARS.sub("", text)
    text = _UNICODE_WHITESPACE.sub(" ", text)
    return text.strip()


def strip_formula(value) -> str:
    text = to_text(value)
    match = re.match(r'^="(.*)"$', text)
    if match:
        text = match.group(1)
    return clean_spaces(text)


def is_blank(value) -> bool:
    raw = strip_formula(value).lower()
    return raw in {"", "na", "n/a", "none", "null", "nan", "-", "0"}


def title_token(token: str) -> str:
    if not token:
        return token
    if "." in token:
        parts = token.split(".")
        fixed = []
        for p in parts:
            if p:
                fixed.append(p[0].upper() + p[1:].lower())
            else:
                fixed.append("")
        return ".".join(fixed)
    return token[0].upper() + token[1:].lower()


def proper_case_text(value: str) -> str:
    value = clean_spaces(value)
    tokens = value.split(" ")
    return " ".join(title_token(token) for token in tokens if token)


def normalize_name_text(raw: str) -> str:
    value = clean_spaces(strip_formula(raw))
    value = re.sub(r"(?i)\b(dr)\.(\S)", r"\1. \2", value)
    return proper_case_text(value)


def split_name(full_name_raw: str) -> Tuple[str, str, str]:
    if is_blank(full_name_raw):
        return ".", ".", "."

    full = normalize_name_text(full_name_raw)
    words = full.split()
    if not words:
        return ".", ".", "."

    first = words[0]
    last_words = words[1:]

    if words[0].lower() in KNOWN_TITLES and len(words) >= 2:
        first = f"{words[0]} {words[1]}"
        last_words = words[2:]

    last = " ".join(last_words) if last_words else "."
    return first, last, full


def merge_name_fields(full_name_raw: str, first_name_raw: str, last_name_raw: str) -> Tuple[str, str, str]:
    """
    Build attendee name using all mapped sources per row.
    - Start from Full Name when available.
    - Override with non-empty First/Last parts when provided.
    - Rebuild Full Name from resolved First + Last.
    """
    full_first, full_last, _ = split_name(full_name_raw)

    first_part = normalize_name_text(first_name_raw) if not is_blank(first_name_raw) else "."
    last_part = normalize_name_text(last_name_raw) if not is_blank(last_name_raw) else "."

    if full_first == "." and first_part == "." and last_part == ".":
        return ".", ".", "."

    first = first_part if first_part != "." else full_first
    last = last_part if last_part != "." else full_last

    # If only one side is present, keep it in first name per single-name rule.
    if first == "." and last != ".":
        first, last = last, "."

    if first == ".":
        return ".", ".", "."

    full = first if last == "." else clean_spaces(f"{first} {last}")
    return first, last, full


def clean_gender(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return "", None

    value = clean_spaces(strip_formula(raw))
    lower = value.lower()
    lower_plain = re.sub(r"[^a-z]", "", lower)

    male_set = {"male", "m", "man", "maile"}
    female_set = {"female", "f", "woman", "femail", "feamle", "femaile"}
    others_set = {"others", "other", "nonbinary", "nonbinary", "transgender", "prefernottosay"}

    if lower_plain in male_set:
        return "Male", None
    if lower_plain in female_set:
        return "Female", None
    if lower_plain in others_set:
        return "Others", None
    if lower_plain.startswith("m"):
        return "Male", None
    if lower_plain.startswith("f"):
        return "Female", None

    return value, "yellow"


def clean_phone(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return "", None

    value = clean_spaces(strip_formula(raw))
    digits = re.sub(r"\D", "", value)

    normalized = None
    if len(digits) == 10:
        normalized = digits
    elif len(digits) == 12 and digits.startswith("91"):
        normalized = digits[-10:]
    elif len(digits) == 11 and digits.startswith("0"):
        normalized = digits[-10:]

    if normalized and len(normalized) == 10:
        return normalized, None
    return value, "yellow"


def format_dob_output(dt: pd.Timestamp) -> str:
    """Output as DD-MM-YYYY with leading zeros."""
    return dt.strftime("%d-%m-%Y")


def resolve_numeric_date_parts(raw: str, month_first: bool = False) -> Optional[Tuple[int, int, int]]:
    """
    Decide (day, month, year) from a numeric date string, without checking that
    the calendar date actually exists.
    - Default numeric day/month input to DD-MM-YYYY (MM-DD-YYYY when month_first).
    - If second part is clearly a day (>12), treat input as MM-DD-YYYY.
    - YYYY-MM-DD (ISO) is also supported when first part is clearly a year.
    """
    raw = raw.strip().replace(".", "-").replace("/", "-")
    match = re.match(r"^(\d{1,4})-(\d{1,2})-(\d{1,2}|\d{4})$", raw)
    if not match:
        return None

    p1, p2, p3 = int(match.group(1)), int(match.group(2)), int(match.group(3))

    if p1 > 31:
        year, month, day = p1, p2, p3
    elif p2 > 12 or (month_first and p1 <= 12):
        # MM-DD-YYYY input (e.g. 5/15/1991) -> convert to DD-MM-YYYY.
        month, day, year = p1, p2, p3
    else:
        day, month, year = p1, p2, p3

    if year < 100:
        year += 2000 if year < 30 else 1900

    return day, month, year


def detect_month_first(values) -> bool:
    """
    True when a date column is consistently MM-DD: some values only make sense
    month-first (second part > 12) and none only make sense day-first.
    """
    month_first_hits = 0
    for raw in values:
        match = re.match(r"^\s*(\d{1,2})[-/.](\d{1,2})[-/.]\d{2,4}\s*$", str(raw))
        if not match:
            continue
        p1, p2 = int(match.group(1)), int(match.group(2))
        if p1 > 12:
            return False
        if p2 > 12:
            month_first_hits += 1
    return month_first_hits > 0


def parse_numeric_date(raw: str, month_first: bool = False) -> Optional[pd.Timestamp]:
    """Parse numeric date parts into a timestamp (NaT if the date is impossible)."""
    parts = resolve_numeric_date_parts(raw, month_first)
    if parts is None:
        return None
    day, month, year = parts
    try:
        return pd.Timestamp(year=year, month=month, day=day)
    except ValueError:
        return pd.NaT


def clean_dob(raw: str, month_first: bool = False) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return "", None

    value = clean_spaces(strip_formula(raw))
    value = value.strip("'\"")

    ordinal_fixed = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", value, flags=re.IGNORECASE)
    # "1990-06-02 00:00:00" (datetime exported as text) must stay year-month-day.
    ordinal_fixed = re.sub(
        r"^(\d{4}-\d{1,2}-\d{1,2})[ T]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?$", r"\1", ordinal_fixed
    )

    # Handle pure numeric Excel serial date.
    if re.fullmatch(r"\d{5}(\.\d+)?", ordinal_fixed):
        serial = float(ordinal_fixed)
        dt = pd.to_datetime(serial, origin="1899-12-30", unit="D", errors="coerce")
    else:
        dt = parse_numeric_date(ordinal_fixed, month_first)
        if dt is None:
            dt = pd.to_datetime(ordinal_fixed, dayfirst=True, errors="coerce")

    if pd.isna(dt):
        # Impossible calendar date: still reorder to DD-MM-YYYY so an obvious
        # MM-DD-YYYY input (e.g. 2-30-2025) is shown as 30-02-2025, flagged red.
        parts = resolve_numeric_date_parts(ordinal_fixed, month_first)
        if parts is not None:
            day, month, year = parts
            return f"{day:02d}-{month:02d}-{year:04d}", "red"
        if re.match(r"^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$", value):
            return value, "red"
        return value, "yellow"

    if dt.date() > date.today() or dt.year < 1900:
        return value, "red"

    return format_dob_output(dt), None


def clean_tshirt(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return ".", None

    value = clean_spaces(strip_formula(raw))
    key = value.lower().replace(" ", "")
    key = key.replace("_", "-")

    for word, out in TSHIRT_MAP.items():
        if key == word.replace("-", "") or key == word:
            return out, None

    # Combined formats like XL-44, M40, S-38
    match = re.match(r"^(xxxl|xxl|xl|xxs|xs|s|m|l)[- ]?(\d{2})?$", key)
    if match:
        letter = match.group(1).upper()
        return letter, None

    digits = re.sub(r"\D", "", key)
    if digits in NUM_TO_SIZE:
        return NUM_TO_SIZE[digits], None

    return value, "yellow"


def normalize_blood_group(raw: str) -> Optional[str]:
    """Normalize common blood group text variants to A+/A-/B+/etc."""
    value = clean_spaces(strip_formula(raw)).upper()

    if re.search(r"[/\\_,;|]", value):
        return None

    value = re.sub(r"POSITIVE", "+", value)
    value = re.sub(r"NEGATIVE", "-", value)
    value = re.sub(r"\bPOS\b", "+", value)
    value = re.sub(r"\bNEG\b", "-", value)
    value = value.replace("+VE", "+").replace("-VE", "-")
    value = re.sub(r"\bRH\b", "", value)

    if re.search(r"\d", value):
        return None

    compact = re.sub(r"[^ABO+\-]", "", value)
    compact = re.sub(r"(AB|A|B|O)-\+", r"\1+", compact)
    compact = re.sub(r"(AB|A|B|O)--", r"\1-", compact)

    if compact in VALID_BLOOD_GROUPS:
        return compact
    return None


def clean_blood_group(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return ".", None

    original = clean_spaces(strip_formula(raw))
    normalized = normalize_blood_group(original)
    if normalized:
        return normalized, None

    return original, None


def clean_address_like(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return ".", None
    value = proper_case_text(strip_formula(raw))
    return value, None


def clean_country(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return "India", None
    value = proper_case_text(strip_formula(raw))
    key = value.lower()
    if key in COUNTRY_MAP:
        return COUNTRY_MAP[key], None
    if re.search(r"[a-zA-Z]", value):
        return value, None
    return value, "yellow"


def clean_emergency_relation(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return ".", None
    value = proper_case_text(strip_formula(raw))
    key = value.lower()
    if key in EMERGENCY_RELATION_MAP:
        return EMERGENCY_RELATION_MAP[key], None
    if key == "bhai":
        return "Brother", None
    return value, "yellow"


def is_dob_column(col_name: str) -> bool:
    norm = normalize_header(col_name)
    return (
        norm.startswith("date of birth")
        or norm in {"dob", "birth date", "birthdate"}
        or "birth date" in norm
    )


def find_dob_source_columns(df: pd.DataFrame, primary: Optional[str]) -> List[str]:
    """Primary mapped DOB column plus any duplicate DOB columns in the file."""
    if not primary:
        return []
    cols = []
    for col in df.columns:
        if col == primary or is_dob_column(col):
            cols.append(col)
    if primary in cols:
        cols.remove(primary)
        return [primary] + cols
    return [primary]


def parse_dob_for_excel(value: str) -> Optional[datetime]:
    """Convert cleaned DD-MM-YYYY text to datetime for Excel sorting."""
    if is_blank(value) or value == ".":
        return None
    match = re.match(r"^(\d{2})-(\d{2})-(\d{4})$", str(value).strip())
    if not match:
        return None
    day, month, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
    try:
        return datetime(year, month, day)
    except ValueError:
        return None


def calculate_age_from_dob(dob_value: str, age_as_on: date) -> str:
    """Calculate completed age (years) from cleaned DOB using a manual as-on date."""
    parsed = parse_dob_for_excel(dob_value)
    if not parsed:
        return ""
    dob_date = parsed.date()
    years = age_as_on.year - dob_date.year
    if (age_as_on.month, age_as_on.day) < (dob_date.month, dob_date.day):
        years -= 1
    return str(max(0, years))


def normalize_header(value: object) -> str:
    text = str(value)
    # Split camelCase / PascalCase so firstName -> first Name.
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def guess_mapping(columns) -> Dict[str, Optional[str]]:
    guessed = {field: None for field in CANONICAL_FIELDS}
    lowered = {col: normalize_header(col) for col in columns}
    used_cols = set()

    def claim(field: str, col: object) -> bool:
        if guessed[field] is not None or col in used_cols:
            return False
        guessed[field] = col
        used_cols.add(col)
        return True

    # Pass 1: exact header == alias
    for field, names in ALIASES.items():
        alias_norms = {normalize_header(alias) for alias in names}
        for col, normalized in lowered.items():
            if normalized in alias_norms and claim(field, col):
                break

    # Pass 2: alias is a full token in the header (avoids "name" matching "first name")
    for field, names in ALIASES.items():
        if guessed[field] is not None:
            continue
        for col, normalized in lowered.items():
            tokens = set(normalized.split())
            for alias in names:
                alias_norm = normalize_header(alias)
                alias_tokens = alias_norm.split()
                if not alias_tokens:
                    continue
                if len(alias_tokens) == 1:
                    token = alias_tokens[0]
                    # Bare "name" is too generic for substring/token matching.
                    if token == "name":
                        if normalized == "name" and claim(field, col):
                            break
                    elif token in tokens and claim(field, col):
                        break
                elif all(t in tokens for t in alias_tokens) and claim(field, col):
                    break
            if guessed[field] is not None:
                break

    return guessed


def format_excel_date_display(value, number_format: str) -> str:
    """
    Rebuild the string a user actually sees in an Excel date cell.

    Excel stores dates as a datetime, so an ambiguous entry like ``02-06-1976``
    can be read back as ``1976-02-06`` and lose its displayed field order. We
    reconstruct the visible order (dd-mm vs mm-dd) from the cell's number format
    so the downstream DD-MM default in clean_dob interprets it as written.
    """
    day = f"{value.day:02d}"
    month = f"{value.month:02d}"
    year = f"{value.year:04d}"

    fmt = (number_format or "").lower()
    d_idx = fmt.find("d")
    m_idx = fmt.find("m")

    if d_idx == -1 and m_idx == -1:
        return f"{day}-{month}-{year}"
    if d_idx == -1:
        d_idx = 10 ** 6
    if m_idx == -1:
        m_idx = 10 ** 6

    if m_idx < d_idx:
        return f"{month}-{day}-{year}"
    return f"{day}-{month}-{year}"


def cell_to_text(cell) -> str:
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (datetime, date)):
        return format_excel_date_display(value, cell.number_format)
    return str(value)


def dedupe_headers(headers: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    result: List[str] = []
    for idx, name in enumerate(headers):
        clean = name.strip() if isinstance(name, str) else ""
        if not clean:
            clean = f"Unnamed: {idx}"
        if clean in seen:
            seen[clean] += 1
            clean = f"{clean}.{seen[clean]}"
        else:
            seen[clean] = 0
        result.append(clean)
    return result


SUMMARY_SHEET_HINTS = (
    "count",
    "counts",
    "summary",
    "total",
    "totals",
    "pivot",
    "dashboard",
    "stats",
    "statistic",
    "report",
)

PERSON_COLUMN_HINTS = (
    "name",
    "first name",
    "lastname",
    "last name",
    "mobile",
    "phone",
    "dob",
    "birth",
    "gender",
    "email",
    "bib",
    "category",
    "tshirt",
    "t shirt",
)


HEADER_SCAN_ROWS = 15
# Formatting can extend a sheet to a million empty rows; stop after this many in a row.
MAX_TRAILING_EMPTY_ROWS = 1000


def header_hint_hits(row: List[str]) -> int:
    cells = [normalize_header(v) for v in row if str(v).strip()]
    return sum(1 for c in cells if any(h == c or h in c for h in PERSON_COLUMN_HINTS))


def detect_header_row(rows: List[List[str]]) -> int:
    """Index of the header row: the first row, unless a later row is clearly the header
    (e.g. a report title or banner sits above the real column names)."""
    first_hits = header_hint_hits(rows[0])
    best_idx, best_hits = 0, first_hits
    for idx, row in enumerate(rows[1:HEADER_SCAN_ROWS], start=1):
        hits = header_hint_hits(row)
        if hits >= 2 and hits > best_hits:
            best_idx, best_hits = idx, hits
    return best_idx


def rows_to_dataframe(rows: List[List[str]]) -> pd.DataFrame:
    rows = [["" if v is None else str(v) for v in row] for row in rows]
    rows = [row for row in rows if any(v.strip() for v in row)]
    if not rows:
        return pd.DataFrame()

    header_idx = detect_header_row(rows)
    headers, data = rows[header_idx], rows[header_idx + 1 :]
    width = max([len(headers)] + [len(r) for r in data])
    headers = headers + [""] * (width - len(headers))
    data = [(r + [""] * (width - len(r)))[:width] for r in data]

    used = [
        i for i in range(width)
        if headers[i].strip() or any(r[i].strip() for r in data)
    ]
    headers = [headers[i] for i in used]
    data = [[r[i] for i in used] for r in data]
    return pd.DataFrame(data, columns=dedupe_headers(headers)).fillna("")


def worksheet_to_dataframe(worksheet) -> pd.DataFrame:
    # Some exporters write a wrong sheet size; without this, read-only mode truncates rows.
    if hasattr(worksheet, "reset_dimensions"):
        worksheet.reset_dimensions()
    rows: List[List[str]] = []
    empty_run = 0
    for row in worksheet.iter_rows():
        cells = [cell_to_text(cell) for cell in row]
        if any(v.strip() for v in cells):
            empty_run = 0
            rows.append(cells)
        else:
            empty_run += 1
            if empty_run > MAX_TRAILING_EMPTY_ROWS:
                break
    return rows_to_dataframe(rows)


def read_xlsx_all_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    sheets: Dict[str, pd.DataFrame] = {}
    for sheet_name in workbook.sheetnames:
        sheets[sheet_name] = worksheet_to_dataframe(workbook[sheet_name])
    workbook.close()
    return sheets


def decode_text_bytes(file_bytes: bytes) -> str:
    if file_bytes.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return file_bytes.decode("utf-16")
    for encoding in ("utf-8-sig", "cp1252"):
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    return file_bytes.decode("latin-1")


def read_delimited_text(file_bytes: bytes) -> pd.DataFrame:
    text = decode_text_bytes(file_bytes)
    sample = "\n".join(text.splitlines()[:50])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    csv.field_size_limit(10 * 1024 * 1024)
    return rows_to_dataframe(list(csv.reader(io.StringIO(text), dialect)))


def legacy_cell_to_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return f"{value.day:02d}-{value.month:02d}-{value.year:04d}"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_legacy_xls_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes), engine="xlrd")
    sheets: Dict[str, pd.DataFrame] = {}
    for sheet_name in excel.sheet_names:
        raw = pd.read_excel(excel, sheet_name=sheet_name, header=None)
        rows = [[legacy_cell_to_text(v) for v in row] for row in raw.itertuples(index=False)]
        sheets[sheet_name] = rows_to_dataframe(rows)
    return sheets


def read_html_table(file_bytes: bytes) -> pd.DataFrame:
    tables = pd.read_html(io.StringIO(decode_text_bytes(file_bytes)), header=None)
    biggest = max(tables, key=len)
    rows = biggest.astype(object).where(biggest.notna(), "").astype(str).values.tolist()
    # <th> cells become the column labels even with header=None; put them back as a row.
    if not isinstance(biggest.columns, pd.RangeIndex):
        header = [
            " ".join(str(p) for p in col if not str(p).startswith("Unnamed")) if isinstance(col, tuple) else str(col)
            for col in biggest.columns
        ]
        rows = [header] + rows
    return rows_to_dataframe(rows)


def read_xlsx_with_display_dates(file_bytes: bytes) -> pd.DataFrame:
    sheets = read_xlsx_all_sheets(file_bytes)
    if not sheets:
        return pd.DataFrame()
    return next(iter(sheets.values()))


def looks_like_summary_sheet(sheet_name: str, df: pd.DataFrame) -> bool:
    name = re.sub(r"[^a-z0-9]+", " ", str(sheet_name).lower()).strip()
    tokens = set(name.split())
    if name in SUMMARY_SHEET_HINTS or tokens.intersection(SUMMARY_SHEET_HINTS):
        return True
    if any(hint in name for hint in SUMMARY_SHEET_HINTS):
        return True

    cols_norm = [
        re.sub(r"[^a-z0-9]+", " ", str(c).lower()).strip() for c in df.columns
    ]
    has_person_col = any(
        any(hint == col or hint in col for hint in PERSON_COLUMN_HINTS)
        for col in cols_norm
    )
    if not has_person_col and len(df) <= 50:
        return True
    return False


def make_source_id(file_id: str, sheet_name: Optional[str]) -> str:
    return f"{file_id}::{sheet_name or '__file__'}"


def map_key(source_id: str, field: str) -> str:
    return f"map::{source_id}::{field}"


def make_source_entry(
    file_id: str,
    filename: str,
    sheet_name: str,
    df: pd.DataFrame,
    default_keep: bool,
) -> Dict[str, object]:
    return {
        "source_id": make_source_id(file_id, sheet_name or None),
        "file_id": file_id,
        "file_name": filename,
        "sheet_name": sheet_name,
        "default_keep": default_keep,
        "rows": len(df),
        # Fast keep/ignore signal: full row count for runner sheets, 0 for summary-like.
        "runner_count": 0 if not default_keep else len(df),
        "columns": list(df.columns),
        "df": df,
    }


@st.cache_data(show_spinner=False)
def discover_file_sources(file_bytes: bytes, filename: str) -> List[Dict[str, object]]:
    """Return one source entry per CSV file or Excel sheet.

    The format is detected from the file contents, not the extension: portals often
    export HTML or CSV text with an .xls name, or real xlsx with an .xls name.
    """
    sources: List[Dict[str, object]] = []
    file_id = hashlib.md5(file_bytes).hexdigest()
    head = file_bytes[:2048].lstrip().lower()

    if not file_bytes.strip():
        raise ValueError("The file is empty.")

    try:
        if file_bytes.startswith(b"PK"):
            sheets = read_xlsx_all_sheets(file_bytes)
        elif file_bytes.startswith(b"\xd0\xcf\x11\xe0"):
            sheets = read_legacy_xls_sheets(file_bytes)
        elif b"<table" in file_bytes[:200000].lower() or head.startswith((b"<html", b"<!doctype")):
            sheets = {"": read_html_table(file_bytes)}
        else:
            sheets = {"": read_delimited_text(file_bytes)}
    except Exception as exc:
        raise ValueError(
            f"this file could not be opened. It may be corrupted or password-protected ({exc})."
        ) from exc

    for sheet_name, df in sheets.items():
        if df.empty and len(df.columns) == 0:
            continue
        keep = True if not sheet_name else not looks_like_summary_sheet(sheet_name, df)
        sources.append(make_source_entry(file_id, filename, sheet_name, df, keep))

    if not sources:
        raise ValueError("No data found in this file.")
    return sources


@st.cache_data(show_spinner=False)
def read_uploaded_file_cached(file_bytes: bytes, filename: str) -> pd.DataFrame:
    sources = discover_file_sources(file_bytes, filename)
    if not sources:
        return pd.DataFrame()
    return sources[0]["df"]  # type: ignore[return-value]


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    file_bytes = uploaded_file.getvalue()
    return read_uploaded_file_cached(file_bytes, uploaded_file.name)


CLEANED_COLUMN_ORDER = [
    "Full Name",
    "First Name",
    "Last Name",
    "Gender",
    "Phone",
    "Date of Birth",
    "Age",
    "Category",
    "T-Shirt Size",
    "Blood Group",
    "Address",
    "City",
    "State",
    "Country",
    "Emergency First Name",
    "Emergency Last Name",
    "Emergency Full Name",
    "Emergency Phone",
    "Emergency Relation",
]


def merge_cleaned_frames(
    parts: List[Tuple[pd.DataFrame, Dict[Tuple[int, str], str], str, str]]
) -> Tuple[pd.DataFrame, Dict[Tuple[int, str], str]]:
    """Union-merge cleaned frames; keep every column; remapped cell flags."""
    if not parts:
        return pd.DataFrame(), {}

    frames: List[pd.DataFrame] = []
    merged_flags: Dict[Tuple[int, str], str] = {}
    offset = 0
    for cleaned_df, cell_flags, file_name, sheet_name in parts:
        part = cleaned_df.copy()
        part = part.rename(
            columns={c: f"{c} (original)" for c in ("Source File", "Source Sheet") if c in part.columns}
        )
        part["Source File"] = file_name
        part["Source Sheet"] = sheet_name or ""
        frames.append(part)
        for (row_idx, col_name), flag in cell_flags.items():
            merged_flags[(row_idx + offset, col_name)] = flag
        offset += len(part)

    merged = pd.concat(frames, ignore_index=True, sort=False).fillna("")
    cleaned_order = [c for c in CLEANED_COLUMN_ORDER if c in merged.columns]
    extra_cols = [c for c in merged.columns if c not in cleaned_order]
    # Keep Source File / Source Sheet near the front of extras for clarity.
    source_cols = [c for c in ("Source File", "Source Sheet") if c in extra_cols]
    other_extras = [c for c in extra_cols if c not in source_cols]
    final = merged[cleaned_order + source_cols + other_extras].copy()
    return sort_cleaned_df(final, merged_flags)


def mark_name_issue(cell_flags: Dict[Tuple[int, str], str], row_idx: int, columns: List[str]):
    for column in columns:
        cell_flags[(row_idx, column)] = "name"


def render_error_legend(error_counts: Optional[Dict[str, int]] = None) -> None:
    """Colour legend; with counts it doubles as the results summary."""
    cards = []
    for error_type, info in ERROR_LEGEND.items():
        count_html = ""
        if error_counts is not None:
            count = int(error_counts.get(error_type, 0))
            if count == 0:
                continue
            count_html = f'<span class="err-count">{count}</span>'
        cards.append(
            f'<div class="err-card">'
            f'<span class="err-dot" style="background:#{info["hex"]};box-shadow:0 0 12px #{info["hex"]}88"></span>'
            f'<span class="err-name">{info["label"]}</span>{count_html}</div>'
        )
    st.markdown(f'<div class="err-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def count_errors_by_type(cell_flags: Dict[Tuple[int, str], str]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for error_type in cell_flags.values():
        counts[error_type] = counts.get(error_type, 0) + 1
    return counts


def style_preview_dataframe(df: pd.DataFrame, cell_flags: Dict[Tuple[int, str], str]):
    def row_style(row: pd.Series):
        styles = []
        for col_name in row.index:
            error_type = cell_flags.get((row.name, col_name))
            if error_type and error_type in ERROR_LEGEND:
                color = ERROR_LEGEND[error_type]["hex"]
                # Dark text keeps flagged cells readable on the dark grid theme.
                styles.append(f"background-color: #{color}; color: #14060a")
            else:
                styles.append("")
        return styles

    return df.style.apply(row_style, axis=1)


def apply_cleaning(
    df: pd.DataFrame,
    mapping: Dict[str, Optional[str]],
    category_map: Dict[str, str],
    tshirt_map: Dict[str, str],
    age_as_on: date,
) -> Tuple[pd.DataFrame, Dict[Tuple[int, str], str]]:
    out = df.copy()
    cell_flags: Dict[Tuple[int, str], str] = {}

    def mark_cell(row_idx: int, column: str, flag: Optional[str]):
        if not flag:
            return
        error_type = COLUMN_FLAG_TO_ERROR.get(column, {}).get(flag)
        if error_type:
            cell_flags[(row_idx, column)] = error_type

    def mark_empty(row_idx: int, column: str):
        error_type = EMPTY_FIELD_ERROR.get(column)
        if error_type:
            cell_flags[(row_idx, column)] = error_type

    # Name handling: if full name exists use it to derive first/last/full by your rules.
    full_name_col = mapping.get("full_name")
    first_name_col = mapping.get("first_name")
    last_name_col = mapping.get("last_name")

    if full_name_col or first_name_col or last_name_col:
        cleaned_first = []
        cleaned_last = []
        cleaned_full = []

        blanks = [""] * len(out)
        full_vals = out[full_name_col].tolist() if full_name_col else blanks
        first_vals = out[first_name_col].tolist() if first_name_col else blanks
        last_vals = out[last_name_col].tolist() if last_name_col else blanks
        for i, (full_raw, first_raw, last_raw) in enumerate(zip(full_vals, first_vals, last_vals)):
            first, last, full = merge_name_fields(full_raw, first_raw, last_raw)
            cleaned_first.append(first)
            cleaned_last.append(last)
            cleaned_full.append(full)
            if first == "." and last == "." and full == ".":
                mark_name_issue(cell_flags, i, NAME_OUTPUT_COLUMNS)

        out["First Name"] = cleaned_first
        out["Last Name"] = cleaned_last
        out["Full Name"] = cleaned_full

    for target, cleaner in [
        ("Gender", clean_gender),
        ("Phone", clean_phone),
        ("Blood Group", clean_blood_group),
        ("Address", clean_address_like),
        ("City", clean_address_like),
        ("State", clean_address_like),
        ("Emergency Phone", clean_phone),
        ("Emergency Relation", clean_emergency_relation),
    ]:
        field_key = {
            "Gender": "gender",
            "Phone": "phone",
            "Blood Group": "blood_group",
            "Address": "address",
            "City": "city",
            "State": "state",
            "Emergency Phone": "emergency_phone",
            "Emergency Relation": "emergency_relation",
        }[target]
        src_col = mapping.get(field_key)
        if src_col:
            cleaned_values = []
            for i, raw in enumerate(out[src_col].tolist()):
                cleaned, flag = cleaner(raw)
                cleaned_values.append(cleaned)
                if is_blank(raw):
                    mark_empty(i, target)
                mark_cell(i, target, flag)
            out[target] = cleaned_values

    # DOB: merge first non-empty across duplicate DOB columns.
    dob_primary = mapping.get("dob")
    dob_cols = find_dob_source_columns(out, dob_primary)
    if dob_cols:
        dob_columns = [out[c].tolist() for c in dob_cols]
        raw_dobs = [
            next((v for v in values if not is_blank(v)), "") for values in zip(*dob_columns)
        ]
        month_first = detect_month_first(raw_dobs)
        dob_values = []
        age_values = []
        for i, raw in enumerate(raw_dobs):
            cleaned, flag = clean_dob(raw, month_first)
            age = calculate_age_from_dob(cleaned, age_as_on)
            if not flag and age.isdigit() and int(age) < 3:
                flag = "red"
            dob_values.append(cleaned)
            age_values.append(age)
            if is_blank(raw):
                mark_empty(i, "Date of Birth")
            mark_cell(i, "Date of Birth", flag)
        out["Date of Birth"] = dob_values
        out["Age"] = age_values

    # Emergency name: same first/last/full rules as attendee name.
    emergency_name_col = mapping.get("emergency_name")
    if emergency_name_col:
        em_first, em_last, em_full = [], [], []
        for i, raw in enumerate(out[emergency_name_col].tolist()):
            first, last, full = split_name(raw)
            em_first.append(first)
            em_last.append(last)
            em_full.append(full)
            if first == "." and last == "." and full == ".":
                mark_name_issue(cell_flags, i, EMERGENCY_NAME_OUTPUT_COLUMNS)
        out["Emergency First Name"] = em_first
        out["Emergency Last Name"] = em_last
        out["Emergency Full Name"] = em_full

    # Category mapping
    cat_col = mapping.get("category")
    if cat_col:
        category_values = []
        for raw in out[cat_col].tolist():
            if is_blank(raw):
                category_values.append(".")
                continue
            raw_clean = clean_spaces(strip_formula(raw))
            mapped = category_map.get(raw_clean, raw_clean)
            category_values.append(mapped)
        out["Category"] = category_values

    # T-shirt mapping (manual mapping only, same behavior as category mapping)
    tshirt_col = mapping.get("tshirt_size")
    if tshirt_col:
        tshirt_values = []
        for raw in out[tshirt_col].tolist():
            if is_blank(raw):
                tshirt_values.append(".")
                continue
            raw_clean = clean_spaces(strip_formula(raw))
            mapped = tshirt_map.get(raw_clean, raw_clean)
            tshirt_values.append(mapped)
        out["T-Shirt Size"] = tshirt_values

    # Country logic
    country_col = mapping.get("country")
    if country_col:
        values = []
        for i, raw in enumerate(out[country_col].tolist()):
            cleaned, flag = clean_country(raw)
            values.append(cleaned)
            mark_cell(i, "Country", flag)
        out["Country"] = values
    else:
        out["Country"] = ["India"] * len(out)

    # Keep cleaned-first order.
    cleaned_order = [c for c in CLEANED_COLUMN_ORDER if c in out.columns]

    mapped_sources = {src for src in mapping.values() if src}
    if dob_primary:
        mapped_sources.update(find_dob_source_columns(out, dob_primary))

    extra_cols = [
        c for c in df.columns if c not in cleaned_order and c not in mapped_sources
    ]
    final_order = cleaned_order + extra_cols
    final = out[final_order].copy()
    for col in cleaned_order:
        final[col] = [clean_spaces(v) for v in final[col]]
    return sort_cleaned_df(final, cell_flags)


GENDER_SORT_ORDER = {"Male": 0, "Female": 1, "Others": 2}


def sort_cleaned_df(
    df: pd.DataFrame, cell_flags: Dict[Tuple[int, str], str]
) -> Tuple[pd.DataFrame, Dict[Tuple[int, str], str]]:
    """Sort by category, then gender (Male, Female, Others)."""
    if df.empty:
        return df, cell_flags

    sorted_df = df.copy()
    sorted_df["_orig_idx"] = range(len(sorted_df))

    sort_cols: List[str] = []
    ascending: List[bool] = []

    if "Category" in sorted_df.columns:
        sort_cols.append("Category")
        ascending.append(True)

    if "Gender" in sorted_df.columns:
        sorted_df["_gender_order"] = sorted_df["Gender"].map(lambda g: GENDER_SORT_ORDER.get(g, 3))
        sort_cols.append("_gender_order")
        ascending.append(True)

    if sort_cols:
        sorted_df = sorted_df.sort_values(by=sort_cols, ascending=ascending, kind="stable")

    orig_indices = sorted_df["_orig_idx"].astype(int).tolist()
    drop_cols = [c for c in ["_orig_idx", "_gender_order"] if c in sorted_df.columns]
    sorted_df = sorted_df.drop(columns=drop_cols).reset_index(drop=True)

    new_position = {orig_idx: new_idx for new_idx, orig_idx in enumerate(orig_indices)}
    new_flags: Dict[Tuple[int, str], str] = {
        (new_position[row_idx], col_name): flag
        for (row_idx, col_name), flag in cell_flags.items()
        if row_idx in new_position
    }
    return sorted_df, new_flags


def df_to_excel_with_highlight(df: pd.DataFrame, cell_flags: Dict[Tuple[int, str], str]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Cleaned")
        ws = writer.book["Cleaned"]

        col_name_to_idx = {name: idx + 1 for idx, name in enumerate(df.columns)}
        error_fills = {
            error_type: PatternFill(
                start_color=info["hex"], end_color=info["hex"], fill_type="solid"
            )
            for error_type, info in ERROR_LEGEND.items()
        }

        dob_col_idx = col_name_to_idx.get("Date of Birth")
        phone_col_idx = col_name_to_idx.get("Phone")
        emergency_phone_col_idx = col_name_to_idx.get("Emergency Phone")
        age_col_idx = col_name_to_idx.get("Age")

        for (row_idx, col_name), error_type in cell_flags.items():
            excel_row = row_idx + 2
            col_idx = col_name_to_idx.get(col_name)
            fill = error_fills.get(error_type)
            if col_idx and fill:
                ws.cell(row=excel_row, column=col_idx).fill = fill

        # Write valid DOB values as real Excel dates for proper sorting.
        if dob_col_idx:
            for row_idx, raw_val in enumerate(df["Date of Birth"].tolist()):
                excel_row = row_idx + 2
                parsed = parse_dob_for_excel(raw_val)
                if parsed:
                    cell = ws.cell(row=excel_row, column=dob_col_idx)
                    cell.value = parsed
                    cell.number_format = "DD-MM-YYYY"

        # Write valid phone values as numbers to avoid Excel text warnings.
        for col_idx, col_name in [
            (phone_col_idx, "Phone"),
            (emergency_phone_col_idx, "Emergency Phone"),
        ]:
            if not col_idx or col_name not in df.columns:
                continue
            for row_idx, value in enumerate(df[col_name].tolist()):
                excel_row = row_idx + 2
                raw_val = clean_spaces(value)
                if re.fullmatch(r"\d{10}", raw_val):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    cell.value = int(raw_val)
                    cell.number_format = "0"

        # Write age as numeric cells to avoid Excel text warnings.
        if age_col_idx and "Age" in df.columns:
            for row_idx, value in enumerate(df["Age"].tolist()):
                excel_row = row_idx + 2
                raw_val = clean_spaces(value)
                if re.fullmatch(r"\d+", raw_val):
                    cell = ws.cell(row=excel_row, column=age_col_idx)
                    cell.value = int(raw_val)
                    cell.number_format = "0"
    output.seek(0)
    return output.read()


@st.cache_resource
def get_persisted_state_store():
    return {}


def get_session_persist_id() -> str:
    sid = st.query_params.get("sid")
    if isinstance(sid, list):
        sid = sid[0] if sid else ""
    if not sid:
        sid = uuid.uuid4().hex
        st.query_params["sid"] = sid
    return sid


def restore_persisted_state():
    store = get_persisted_state_store()
    sid = get_session_persist_id()
    persisted = store.get(sid, {})
    for key, value in persisted.items():
        if key not in st.session_state:
            st.session_state[key] = value


def save_persisted_state():
    store = get_persisted_state_store()
    sid = get_session_persist_id()
    # Never persist file bytes — that OOMs Streamlit Cloud on large Excel files.
    keys_to_persist = [
        "loaded_batch_id",
        "source_keep",
        "category_value_map",
        "tshirt_value_map",
        "age_as_on_date",
        "upload_widget_nonce",
        "active_map_source_label",
    ]
    snapshot = {k: st.session_state[k] for k in keys_to_persist if k in st.session_state}
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith("map::"):
            snapshot[key] = st.session_state[key]
    store[sid] = snapshot


def clear_loaded_file_state():
    sid = get_session_persist_id()
    get_persisted_state_store().pop(sid, None)
    for key in [
        "loaded_files",
        "loaded_batch_id",
        "source_keep",
        "category_value_map",
        "tshirt_value_map",
        "category_edit_df",
        "category_edit_fp",
        "tshirt_edit_df",
        "tshirt_edit_fp",
        "category_map_df",
        "category_map_source",
        "category_groups",
        "tshirt_map_df",
        "tshirt_map_source",
        "tshirt_groups",
        "age_as_on_date",
        "active_map_source_label",
    ]:
        st.session_state.pop(key, None)
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and (
            key.startswith("map::")
            or key.startswith("map_")
            or key.startswith("keep_")
            or key.startswith("bulk_")
            or key.startswith("tshirt_bulk_")
            or key.startswith("category_value_editor")
            or key.startswith("tshirt_value_editor")
            or key in {"source_keep_editor"}
        ):
            st.session_state.pop(key, None)
    # Legacy single-file keys
    for key in ("active_file_bytes", "active_file_name", "active_file_id"):
        st.session_state.pop(key, None)
    st.session_state["upload_widget_nonce"] = st.session_state.get("upload_widget_nonce", 0) + 1


def source_label(source: Dict[str, object]) -> str:
    sheet = str(source.get("sheet_name") or "")
    file_name = str(source.get("file_name") or "")
    if sheet:
        return f"{file_name} / {sheet}"
    return file_name


def mapped_field_count(mapping: Dict[str, Optional[str]]) -> int:
    return sum(1 for value in mapping.values() if value)


def column_value_counts(df: pd.DataFrame, col_name: str) -> List[Tuple[str, int]]:
    if col_name not in df.columns:
        return []
    series = df[col_name].fillna("").astype(str).str.strip().replace("", ".")
    counts = series.value_counts()
    return [(str(k), int(v)) for k, v in counts.items()]


def build_value_breakdown(
    sources: List[Dict[str, object]],
    source_mappings: Dict[str, Dict[str, Optional[str]]],
    field_key: str,
) -> pd.DataFrame:
    """Unique values per source for category/tshirt from in-memory dataframes."""
    rows: List[Dict[str, object]] = []
    for source in sources:
        sid = str(source["source_id"])
        col_name = source_mappings.get(sid, {}).get(field_key)
        if not col_name:
            continue
        df = source["df"]  # type: ignore[assignment]
        label = source_label(source)
        for raw_value, count in column_value_counts(df, col_name):
            rows.append({"Source": label, "Raw Value": raw_value, "Rows": count})
    if not rows:
        return pd.DataFrame(columns=["Source", "Raw Value", "Rows"])
    return pd.DataFrame(rows).sort_values(["Source", "Raw Value"]).reset_index(drop=True)


def render_value_map_editor(
    breakdown: pd.DataFrame,
    state_map_key: str,
    editor_key: str,
) -> Dict[str, str]:
    """Simple text inputs — avoids Streamlit data_editor crashes on cloud."""
    if state_map_key not in st.session_state:
        st.session_state[state_map_key] = {}

    saved_map: Dict[str, str] = dict(st.session_state[state_map_key])
    if breakdown.empty:
        return saved_map

    new_map: Dict[str, str] = {}
    # Cap widgets for safety on weird files with huge unique sets.
    max_rows = 80
    view = breakdown.head(max_rows)
    for source, group in view.groupby("Source", sort=True):
        st.markdown(f'<span class="val-src">{source}</span>', unsafe_allow_html=True)
        grid = st.columns(6)
        for pos, (_, row) in enumerate(group.iterrows()):
            raw = str(row["Raw Value"])
            rows_n = int(row["Rows"])
            widget_key = f"{editor_key}::{source}::{raw}"
            mapped = grid[pos % 6].text_input(
                f"{raw} · {rows_n}",
                value=saved_map.get(raw, raw),
                key=widget_key,
            )
            mapped = clean_spaces(mapped) or "."
            # The same raw value can appear in several sheets; a rename in any of them wins.
            if mapped != raw or raw not in new_map:
                new_map[raw] = mapped
    if len(breakdown) > max_rows:
        st.warning(f"Showing first {max_rows} values only ({len(breakdown)} total).")
    st.session_state[state_map_key] = new_map
    return new_map


def render_value_section(
    breakdown: pd.DataFrame,
    state_map_key: str,
    editor_key: str,
    empty_message: str,
) -> Dict[str, str]:
    if breakdown.empty:
        st.info(empty_message)
        return {}

    head_left, head_right = st.columns([5, 1], vertical_alignment="center")
    summary_slot = head_left.empty()
    if head_right.button("Reset", key=f"reset_{editor_key}", use_container_width=True):
        st.session_state[state_map_key] = {}
        for key in list(st.session_state.keys()):
            if isinstance(key, str) and key.startswith(f"{editor_key}::"):
                st.session_state.pop(key, None)
        st.rerun()

    value_map = render_value_map_editor(breakdown, state_map_key, editor_key)
    renamed = sum(1 for raw, mapped in value_map.items() if mapped != raw)
    summary_slot.markdown(
        f'<div class="val-summary"><b>{breakdown["Raw Value"].nunique()}</b> values across '
        f'<b>{breakdown["Source"].nunique()}</b> sheets'
        f'<span class="val-renamed{" on" if renamed else ""}">{renamed} renamed</span></div>',
        unsafe_allow_html=True,
    )
    return value_map


@st.cache_resource
def load_hero_image_b64() -> str:
    path = Path(__file__).parent / "assets" / "goku.png"
    try:
        return base64.b64encode(path.read_bytes()).decode("ascii")
    except OSError:
        return ""


LANDING_CSS = """
<style>
.stApp {
    background:
        radial-gradient(60vw 60vh at 20% 80%, rgba(255, 110, 20, 0.14), transparent 70%),
        #000 !important;
}

:root {
    /* goku.png is 1532x1024; palms sit at the far right edge, half-way down. */
    --gw: min(56vw, calc(92vh * 1.496));
    --gh: calc(var(--gw) / 1.496);
    --palm-x: calc(var(--gw) * 0.975);
    --palm-y: calc(100vh - var(--gh) * 0.5);
    --content-x: calc(var(--gw) + 4.5rem);
}
.block-container {
    max-width: none !important;
    padding-left: var(--content-x) !important;
    padding-right: clamp(1.5rem, 5vw, 4rem) !important;
    padding-top: max(1.5rem, calc(var(--palm-y) - 264px)) !important;
}
.ember:nth-child(n+7) { display: none; }

.goku-hero {
    position: fixed;
    left: 0; bottom: 0;
    width: var(--gw);
    height: var(--gh);
    z-index: 1;
    pointer-events: none;
    -webkit-mask-image: linear-gradient(to bottom, transparent 0%, #000 9%, #000 76%, transparent 100%);
    mask-image: linear-gradient(to bottom, transparent 0%, #000 9%, #000 76%, transparent 100%);
    filter: drop-shadow(0 0 14px rgba(255, 190, 70, 0.35));
    animation: gokuIn 1.1s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
@keyframes gokuIn {
    from { opacity: 0; transform: translateX(-70px); filter: brightness(2.2) drop-shadow(0 0 40px rgba(255, 200, 80, 0.9)); }
    to   { opacity: 1; transform: none; filter: drop-shadow(0 0 14px rgba(255, 190, 70, 0.35)); }
}
.goku-aura {
    position: fixed;
    left: calc(var(--gw) * 0.05);
    bottom: 0;
    width: calc(var(--gw) * 0.8);
    height: calc(var(--gh) * 1.05);
    z-index: 0;
    pointer-events: none;
    background: radial-gradient(closest-side, rgba(255, 150, 40, 0.30), rgba(255, 90, 0, 0.10) 55%, transparent 75%);
    filter: blur(30px);
    animation: auraFlare 1.4s ease-out 0.2s both, auraBreath 4.5s ease-in-out 1.6s infinite;
}
@keyframes auraFlare {
    0%   { opacity: 0; transform: scale(0.6); }
    45%  { opacity: 1; transform: scale(1.15); }
    100% { opacity: 0.7; transform: scale(1); }
}
@keyframes auraBreath {
    0%, 100% { opacity: 0.55; transform: scale(1); }
    50%      { opacity: 0.9;  transform: scale(1.04); }
}

.ki-core {
    position: fixed;
    left: calc(var(--palm-x) - 26px);
    top: calc(var(--palm-y) - 26px);
    width: 52px; height: 52px;
    border-radius: 50%;
    z-index: 3;
    pointer-events: none;
    mix-blend-mode: screen;
    background: radial-gradient(circle, #fff 0%, #fff3d6 20%, rgba(255, 179, 71, 0.8) 42%, rgba(255, 122, 26, 0.25) 62%, transparent 72%);
    animation: coreCharge 1s ease-out 0.7s both, coreIdle 2.8s ease-in-out 1.7s infinite;
}
@keyframes coreCharge {
    0%   { opacity: 0; transform: scale(0.2); }
    60%  { opacity: 1; transform: scale(1.6); }
    100% { opacity: 0.85; transform: scale(1); }
}
@keyframes coreIdle {
    0%, 100% { opacity: 0.55; transform: scale(0.9); }
    50%      { opacity: 0.95; transform: scale(1.08); }
}
.ki-beam {
    position: fixed;
    left: var(--palm-x);
    top: calc(var(--palm-y) - 2px);
    width: calc(var(--content-x) - var(--palm-x));
    height: 4px;
    z-index: 3;
    pointer-events: none;
    border-radius: 4px;
    transform-origin: left center;
    background: linear-gradient(90deg, #fff 0%, #fff3d6 30%, var(--accent-2) 70%, var(--accent) 100%);
    box-shadow: 0 0 12px rgba(255, 179, 71, 0.9), 0 0 28px rgba(255, 122, 26, 0.6);
    animation: beamFire 0.45s cubic-bezier(0.3, 0.9, 0.3, 1) 1.2s both, beamIdle 2.8s ease-in-out 1.7s infinite;
}
@keyframes beamFire { from { transform: scaleX(0); } to { transform: scaleX(1); } }
@keyframes beamIdle {
    0%, 100% { opacity: 0.35; }
    50%      { opacity: 0.8; }
}
.stApp:has([data-testid="stFileUploader"]:hover) .ki-beam { opacity: 1 !important; height: 5px; }
.stApp:has([data-testid="stFileUploader"]:hover) .ki-core { opacity: 1 !important; transform: scale(1.3) !important; }

.landing-head, [data-testid="stFileUploader"] { max-width: 600px; }
.landing-kicker {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 0.72rem;
    letter-spacing: 0.32em;
    text-transform: uppercase;
    color: var(--muted);
    font-weight: 600;
    animation: fadeUp 0.6s ease 0.2s both;
}
.landing-kicker i { width: 7px; height: 7px; border-radius: 50%; background: var(--accent); box-shadow: 0 0 10px var(--accent); }
.landing-title {
    font-family: 'Bebas Neue', sans-serif;
    font-size: clamp(2.4rem, 4.2vw, 4.6rem);
    letter-spacing: 0.04em;
    line-height: 0.95;
    white-space: nowrap;
    color: var(--text);
    margin: 10px 0 8px;
    animation: fadeUp 0.6s ease 0.3s both;
}
.landing-sub {
    color: var(--muted);
    font-size: 1.02rem;
    margin: 0 0 22px;
    animation: fadeUp 0.6s ease 0.4s both;
}

@property --ring { syntax: '<angle>'; initial-value: 270deg; inherits: false; }
[data-testid="stFileUploader"] {
    position: relative;
    z-index: 4;
    background: linear-gradient(160deg, rgba(24, 24, 32, 0.92), rgba(10, 10, 14, 0.94)) !important;
    border: 1px solid var(--line-strong) !important;
    border-radius: 20px !important;
    padding: 24px 24px 20px !important;
    backdrop-filter: blur(10px);
    box-shadow: 0 20px 60px rgba(0, 0, 0, 0.6);
    animation: fadeUp 0.7s cubic-bezier(0.2, 0.8, 0.2, 1) 0.5s both;
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}
[data-testid="stFileUploader"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 24px 70px rgba(0, 0, 0, 0.6), 0 0 40px rgba(255, 122, 26, 0.16);
}
[data-testid="stFileUploader"]::before {
    content: "";
    position: absolute;
    inset: -1px;
    padding: 1.5px;
    border-radius: 21px;
    pointer-events: none;
    background: conic-gradient(from var(--ring), transparent 0deg 290deg, var(--accent) 330deg, #fff3d6 352deg, transparent 360deg);
    -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0);
    -webkit-mask-composite: xor;
    mask-composite: exclude;
    animation: ringRun 1.3s cubic-bezier(0.4, 0, 0.2, 1) 1.6s both, ringLoop 7s linear 2.9s infinite;
}
[data-testid="stFileUploader"]::after {
    content: "";
    position: absolute;
    inset: 0;
    border-radius: 20px;
    pointer-events: none;
    box-shadow: 0 0 0 1px rgba(255, 122, 26, 0.5);
    opacity: 0;
    transition: opacity 0.3s ease;
}
[data-testid="stFileUploader"]:hover::after { opacity: 1; }
@keyframes ringRun  { from { --ring: 270deg; } to { --ring: 630deg; } }
@keyframes ringLoop { from { --ring: 270deg; } to { --ring: 630deg; } }
[data-testid="stFileUploader"] [data-testid="stWidgetLabel"] p {
    font-size: 1.12rem !important;
    font-weight: 700 !important;
    color: var(--text) !important;
}


@media (max-width: 900px) {
    .goku-hero, .goku-aura, .ki-core, .ki-beam { display: none; }
    .block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        padding-top: 2rem !important;
    }
    .landing-head, [data-testid="stFileUploader"] { max-width: none; }
}
</style>
"""


def render_landing_scene() -> None:
    img_b64 = load_hero_image_b64()
    hero_img = (
        f'<img class="goku-hero" src="data:image/png;base64,{img_b64}" alt="">'
        if img_b64
        else ""
    )
    st.markdown(LANDING_CSS, unsafe_allow_html=True)
    st.markdown(
        f'<div class="goku-aura"></div>{hero_img}<div class="ki-beam"></div><div class="ki-core"></div>'
        '<div class="landing-head">'
        '<div class="landing-kicker"><i></i>Sports Timing Solutions &middot; FileSort Cleaner</div>'
        '<div class="landing-title">Upload your <span>file</span></div>'
        '<p class="landing-sub">Clean, merge and flag errors in seconds.</p>'
        "</div>",
        unsafe_allow_html=True,
    )


WORKSPACE_CSS = """
<style>
/* Loaded files are listed in the header; the uploader stays mounted but hidden. */
[data-testid="stElementContainer"]:has([data-testid="stFileUploader"]) { display: none; }
.block-container {
    max-width: none !important;
    padding-left: clamp(1rem, 3vw, 3rem) !important;
    padding-right: clamp(1rem, 3vw, 3rem) !important;
}
</style>
"""


def render_stats(items: List[Tuple[str, int]]) -> None:
    cards = "".join(
        f'<div class="stat"><div class="stat-label">{label}</div>'
        f'<div class="stat-num" style="--to:{int(value)}"></div></div>'
        for label, value in items
    )
    st.markdown(f'<div class="stats">{cards}</div>', unsafe_allow_html=True)


def is_source_kept(source: Dict[str, object]) -> bool:
    sid = str(source["source_id"])
    return bool(
        st.session_state.get(
            f"keep_{sid}",
            st.session_state["source_keep"].get(sid, source["default_keep"]),
        )
    )


def keep_all_sources(sources: List[Dict[str, object]]) -> None:
    for s in sources:
        sid = str(s["source_id"])
        st.session_state["source_keep"][sid] = True
        st.session_state[f"keep_{sid}"] = True


restore_persisted_state()
if "upload_widget_nonce" not in st.session_state:
    st.session_state["upload_widget_nonce"] = 0
if "source_keep" not in st.session_state:
    st.session_state["source_keep"] = {}
if "loaded_files" not in st.session_state:
    st.session_state["loaded_files"] = {}

upload_key = f"source_upload_{st.session_state['upload_widget_nonce']}"
show_landing = not st.session_state.get("loaded_files") and not st.session_state.get(upload_key)
if show_landing:
    render_landing_scene()
uploaded_files = st.file_uploader(
    "Drop your CSV / XLSX files here",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    key=upload_key,
)

if uploaded_files:
    batch_files = []
    for uploaded in uploaded_files:
        file_bytes = uploaded.getvalue()
        file_id = hashlib.md5(file_bytes).hexdigest()
        batch_files.append(
            {
                "file_id": file_id,
                "file_name": uploaded.name,
                "file_bytes": file_bytes,
            }
        )
    batch_id = hashlib.md5(
        "|".join(sorted(f["file_id"] for f in batch_files)).encode("utf-8")
    ).hexdigest()
    if st.session_state.get("loaded_batch_id") == batch_id and not st.session_state.get("loaded_files"):
        # Same files re-uploaded after a refresh: keep restored mappings, reload the data.
        st.session_state["loaded_files"] = {
            f["file_id"]: {"file_name": f["file_name"], "file_bytes": f["file_bytes"]}
            for f in batch_files
        }
    elif st.session_state.get("loaded_batch_id") != batch_id:
        # Reset maps/keep flags for the new batch, but do not remount the uploader
        # (that would wipe the just-selected files).
        for key in [
            "category_value_map",
            "tshirt_value_map",
            "category_edit_df",
            "category_edit_fp",
            "tshirt_edit_df",
            "tshirt_edit_fp",
            "category_map_df",
            "category_map_source",
            "category_groups",
            "tshirt_map_df",
            "tshirt_map_source",
            "tshirt_groups",
            "age_as_on_date",
            "active_map_source_label",
        ]:
            st.session_state.pop(key, None)
        for key in list(st.session_state.keys()):
            if isinstance(key, str) and (
                key.startswith("map::")
                or key.startswith("keep_")
                or key.startswith("bulk_")
                or key.startswith("tshirt_bulk_")
                or key.startswith("category_value_editor")
                or key.startswith("tshirt_value_editor")
                or key in {"source_keep_editor"}
            ):
                st.session_state.pop(key, None)
        st.session_state["loaded_files"] = {
            f["file_id"]: {"file_name": f["file_name"], "file_bytes": f["file_bytes"]}
            for f in batch_files
        }
        st.session_state["loaded_batch_id"] = batch_id
        st.session_state["source_keep"] = {}

all_sources: List[Dict[str, object]] = []
load_errors: List[str] = []
if st.session_state.get("loaded_files"):
    for file_id, meta in st.session_state["loaded_files"].items():
        try:
            discovered = discover_file_sources(meta["file_bytes"], meta["file_name"])
            all_sources.extend(discovered)
        except Exception as exc:
            load_errors.append(f"{meta['file_name']}: {exc}")

if load_errors:
    for msg in load_errors:
        st.error(f"Could not read file: {msg}")

if all_sources:
    for source in all_sources:
        sid = str(source["source_id"])
        if sid not in st.session_state["source_keep"]:
            st.session_state["source_keep"][sid] = bool(source["default_keep"])
        if f"keep_{sid}" not in st.session_state:
            st.session_state[f"keep_{sid}"] = st.session_state["source_keep"][sid]

    file_names = [str(meta["file_name"]) for meta in st.session_state["loaded_files"].values()]
    kept_sources = [s for s in all_sources if is_source_kept(s)]
    total_runners = sum(int(s.get("runner_count", s["rows"])) for s in kept_sources)

    st.markdown(WORKSPACE_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="ki-flash"></div>'
        '<div class="app-bar"><div>'
        '<div class="app-kicker"><i></i>Sports Timing Solutions</div>'
        '<div class="app-title">FileSort <span>Cleaner</span></div>'
        "</div>"
        f'<div class="app-files">{", ".join(file_names[:3])}{" …" if len(file_names) > 3 else ""}</div>'
        "</div>",
        unsafe_allow_html=True,
    )
    render_stats(
        [
            ("Files", len(file_names)),
            ("Sheets kept", len(kept_sources)),
            ("Runners", total_runners),
        ]
    )

    tab_keep, tab_map, tab_values, tab_clean = st.tabs(
        ["Keep sheets", "Map columns", "Map values", "Clean & download"],
        key="workspace_step",
    )

    with tab_keep:
        q1, q2, _ = st.columns([1, 1, 2])
        with q1:
            if st.button("Keep all", use_container_width=True):
                keep_all_sources(all_sources)
                st.rerun()
        with q2:
            if st.button("Start over", use_container_width=True):
                clear_loaded_file_state()
                st.rerun()

        ordered = sorted(
            all_sources,
            key=lambda s: (-int(s.get("runner_count", s["rows"])), -int(s["rows"])),
        )
        multi_file = len(file_names) > 1
        for source in ordered:
            sid = str(source["source_id"])
            runners = int(source.get("runner_count", source["rows"]))
            sheet = str(source.get("sheet_name") or source["file_name"])
            label = f"{sheet} · {source['file_name']}" if multi_file and source.get("sheet_name") else sheet
            runner_word = "runner" if runners == 1 else "runners"
            label += f"  —  {runners} {runner_word}" if source["default_keep"] else "  —  summary sheet"
            keep = st.toggle(label, key=f"keep_{sid}")
            st.session_state["source_keep"][sid] = keep

        kept_sources = [s for s in all_sources if st.session_state["source_keep"].get(str(s["source_id"]), False)]
        if not kept_sources:
            st.warning("Keep at least one sheet to continue.")
            save_persisted_state()
            st.stop()

    with tab_map:
        source_labels = {source_label(s): s for s in kept_sources}
        selected_label = st.selectbox("Working on", list(source_labels.keys()), key="active_map_source_label")
        active_source = source_labels[selected_label]
        active_sid = str(active_source["source_id"])
        active_df = active_source["df"]  # type: ignore[assignment]
        active_cols = list(active_source["columns"])  # type: ignore[arg-type]
        active_guess = guess_mapping(active_cols)
        active_options = ["<None>"] + active_cols

        # Resolve mappings for all kept sources (defaults + saved picks).
        source_mappings: Dict[str, Dict[str, Optional[str]]] = {}
        for source in kept_sources:
            sid = str(source["source_id"])
            cols = list(source["columns"])  # type: ignore[arg-type]
            guessed = guess_mapping(cols)
            options = set(cols) | {"<None>"}
            mapping: Dict[str, Optional[str]] = {}
            for field in CANONICAL_FIELDS:
                key = map_key(sid, field)
                picked = st.session_state.get(key)
                if picked not in options:
                    picked = guessed.get(field) if guessed.get(field) in options else "<None>"
                mapping[field] = None if picked == "<None>" else picked
            source_mappings[sid] = mapping

        health_slot = st.empty()

        essential_ui = st.columns(4)
        for idx, field in enumerate(ESSENTIAL_MAPPING_FIELDS):
            key = map_key(active_sid, field)
            current = source_mappings[active_sid].get(field) or active_guess.get(field) or "<None>"
            if current not in active_options:
                current = "<None>"
            picked = essential_ui[idx % 4].selectbox(
                FIELD_LABELS[field],
                options=active_options,
                index=active_options.index(current),
                key=key,
            )
            source_mappings[active_sid][field] = None if picked == "<None>" else picked

        with st.expander("More columns (optional)", expanded=False):
            advanced_fields = [f for f in CANONICAL_FIELDS if f not in ESSENTIAL_MAPPING_FIELDS]
            adv_ui = st.columns(4)
            for idx, field in enumerate(advanced_fields):
                key = map_key(active_sid, field)
                current = source_mappings[active_sid].get(field) or active_guess.get(field) or "<None>"
                if current not in active_options:
                    current = "<None>"
                picked = adv_ui[idx % 4].selectbox(
                    FIELD_LABELS[field],
                    options=active_options,
                    index=active_options.index(current),
                    key=key,
                )
                source_mappings[active_sid][field] = None if picked == "<None>" else picked

        for field in CANONICAL_FIELDS:
            picked = st.session_state.get(map_key(active_sid, field), "<None>")
            source_mappings[active_sid][field] = None if picked == "<None>" else picked

        has_name = any(
            source_mappings[active_sid].get(f) for f in ("full_name", "first_name", "last_name")
        )
        checks = [has_name] + [
            bool(source_mappings[active_sid].get(f))
            for f in ESSENTIAL_MAPPING_FIELDS
            if f not in ("full_name", "first_name", "last_name")
        ]
        matched, total = sum(checks), len(checks)
        health_slot.markdown(
            f'<div class="health"><span class="health-text">{matched}/{total} key fields matched</span>'
            f'<div class="health-bar"><i style="width:{int(100 * matched / max(total, 1))}%"></i></div></div>',
            unsafe_allow_html=True,
        )

        with st.expander("Preview sheet", expanded=False):
            st.dataframe(active_df.head(8), use_container_width=True)

        age_as_on = st.date_input(
            "Calculate age as on",
            value=date.today(),
            format="DD-MM-YYYY",
            key="age_as_on_date",
        )

    with tab_values:
        category_map: Dict[str, str] = {}
        tshirt_map: Dict[str, str] = {}
        sub_cat, sub_tshirt = st.tabs(["🏃  Categories", "👕  T-shirt sizes"], key="values_step")

        with sub_cat:
            try:
                category_map = render_value_section(
                    build_value_breakdown(kept_sources, source_mappings, "category"),
                    "category_value_map",
                    "category_value_editor",
                    "Map a Category column in step 2 to edit category values.",
                )
            except Exception as exc:
                st.error(f"Category mapping failed: {exc}")

        with sub_tshirt:
            try:
                tshirt_map = render_value_section(
                    build_value_breakdown(kept_sources, source_mappings, "tshirt_size"),
                    "tshirt_value_map",
                    "tshirt_value_editor",
                    "Map a T-Shirt column in step 2 to edit T-shirt values.",
                )
            except Exception as exc:
                st.error(f"T-Shirt mapping failed: {exc}")

    with tab_clean:
        if st.button("Run cleaning", type="primary", use_container_width=True):
            parts: List[Tuple[pd.DataFrame, Dict[Tuple[int, str], str], str, str]] = []
            progress = st.progress(0.0, text="Starting…")
            stage = "starting"
            try:
                for idx, source in enumerate(kept_sources):
                    sid = str(source["source_id"])
                    stage = f"cleaning {source_label(source)}"
                    progress.progress(idx / len(kept_sources), text=f"Cleaning {source_label(source)}…")
                    cleaned_df, cell_flags = apply_cleaning(
                        source["df"],  # type: ignore[arg-type]
                        source_mappings[sid],
                        category_map,
                        tshirt_map,
                        age_as_on,
                    )
                    parts.append(
                        (
                            cleaned_df,
                            cell_flags,
                            str(source["file_name"]),
                            str(source.get("sheet_name") or ""),
                        )
                    )
                stage = "merging sheets"
                progress.progress(0.95, text="Merging sheets…")
                cleaned_df, cell_flags = merge_cleaned_frames(parts)
                stage = "building the Excel file"
                excel_bytes = df_to_excel_with_highlight(cleaned_df, cell_flags)
            except Exception as exc:
                progress.empty()
                st.error(f"Cleaning stopped while {stage}: {exc}")
                save_persisted_state()
                st.stop()
            progress.empty()

            error_counts = count_errors_by_type(cell_flags)
            flagged = sum(error_counts.values())
            st.success(
                f"{len(cleaned_df)} rows cleaned · "
                + (f"{flagged} cells flagged" if flagged else "no issues found")
            )
            st.download_button(
                label="Download cleaned Excel",
                data=excel_bytes,
                file_name="cleaned_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                on_click="ignore",
            )
            if flagged:
                render_error_legend(error_counts)
            preview_df = cleaned_df.head(20)
            if cell_flags:
                st.dataframe(style_preview_dataframe(preview_df, cell_flags), use_container_width=True)
            else:
                st.dataframe(preview_df, use_container_width=True)
    save_persisted_state()
