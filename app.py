import hashlib
import io
import re
import uuid
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


st.set_page_config(page_title="FileSort Cleaner", layout="wide", page_icon="🔥")

THEME_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Orbitron:wght@600;900&family=Rajdhani:wght@500;600;700&display=swap');

:root {
    --fire-core: #fff3b0;
    --fire-hot: #ffd60a;
    --fire-mid: #ff8a00;
    --fire-deep: #ff3d00;
    --ember: #ff2e63;
    --neon: #00e5ff;
    --void: #05040c;
    --panel: rgba(18, 14, 32, 0.72);
    --edge: rgba(255, 138, 0, 0.35);
}

[data-testid="stToolbar"] {display: none !important;}
[data-testid="stDecoration"] {display: none !important;}
header {display: none !important;}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}

/* ---------- Base canvas ---------- */
.stApp {
    background:
        radial-gradient(1200px 600px at 12% -10%, rgba(255, 61, 0, 0.20), transparent 60%),
        radial-gradient(900px 500px at 88% 0%, rgba(0, 229, 255, 0.14), transparent 55%),
        radial-gradient(800px 700px at 50% 110%, rgba(255, 46, 99, 0.18), transparent 60%),
        linear-gradient(180deg, #05040c 0%, #0b0716 45%, #120a1e 100%);
    background-attachment: fixed;
    color: #f2ecff;
    font-family: 'Rajdhani', 'Segoe UI', sans-serif;
}

/* Animated aurora haze over the base */
.stApp::before {
    content: "";
    position: fixed;
    inset: -20%;
    pointer-events: none;
    z-index: 0;
    background:
        conic-gradient(from 0deg at 30% 40%, rgba(255,61,0,0.10), rgba(0,229,255,0.06), rgba(255,46,99,0.10), rgba(255,61,0,0.10));
    filter: blur(90px);
    animation: auroraSpin 24s linear infinite;
    opacity: 0.85;
}

/* Scanline / grid texture for the anime tech feel */
.stApp::after {
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    z-index: 0;
    background-image:
        linear-gradient(rgba(255, 255, 255, 0.030) 1px, transparent 1px),
        linear-gradient(90deg, rgba(255, 255, 255, 0.030) 1px, transparent 1px);
    background-size: 48px 48px;
    mask-image: radial-gradient(circle at 50% 30%, black 10%, transparent 78%);
    -webkit-mask-image: radial-gradient(circle at 50% 30%, black 10%, transparent 78%);
    opacity: 0.5;
}

[data-testid="stAppViewContainer"] > .main { position: relative; z-index: 1; }
.block-container { position: relative; z-index: 2; padding-top: 1.5rem !important; }

@keyframes auroraSpin {
    from { transform: rotate(0deg) scale(1.05); }
    to   { transform: rotate(360deg) scale(1.05); }
}

/* ---------- Ember particle field ---------- */
.ember-field {
    position: fixed;
    inset: 0;
    pointer-events: none;
    z-index: 1;
    overflow: hidden;
}
.ember {
    position: absolute;
    bottom: -40px;
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: radial-gradient(circle, var(--fire-core) 0%, var(--fire-mid) 45%, rgba(255, 61, 0, 0) 70%);
    box-shadow: 0 0 12px 3px rgba(255, 138, 0, 0.65);
    animation: emberRise linear infinite;
    opacity: 0;
}
@keyframes emberRise {
    0%   { transform: translate3d(0, 0, 0) scale(0.6); opacity: 0; }
    12%  { opacity: 1; }
    70%  { opacity: 0.85; }
    100% { transform: translate3d(var(--drift, 40px), -105vh, 0) scale(1.25); opacity: 0; }
}

/* Base of the screen burning */
.fire-floor {
    position: fixed;
    left: 0; right: 0; bottom: 0;
    height: 190px;
    pointer-events: none;
    z-index: 1;
    background: linear-gradient(0deg, rgba(255, 61, 0, 0.38) 0%, rgba(255, 138, 0, 0.16) 38%, transparent 100%);
    filter: blur(22px);
    animation: firePulse 3.4s ease-in-out infinite alternate;
}
@keyframes firePulse {
    from { opacity: 0.55; transform: scaleY(0.92); }
    to   { opacity: 1;    transform: scaleY(1.10); }
}

/* ---------- Hero ---------- */
.hero {
    position: relative;
    padding: 34px 38px 30px;
    margin-bottom: 26px;
    border-radius: 22px;
    background: linear-gradient(135deg, rgba(26, 16, 46, 0.86), rgba(10, 8, 22, 0.92));
    border: 1px solid var(--edge);
    box-shadow:
        0 0 0 1px rgba(255, 255, 255, 0.04) inset,
        0 24px 70px rgba(255, 61, 0, 0.18),
        0 0 90px rgba(0, 229, 255, 0.07);
    overflow: hidden;
}
.hero::before {
    content: "";
    position: absolute;
    top: -60%; left: -30%;
    width: 60%; height: 220%;
    background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.16), transparent);
    transform: rotate(18deg);
    animation: slashSweep 5.5s ease-in-out infinite;
}
@keyframes slashSweep {
    0%, 62%  { left: -40%; }
    100%     { left: 130%; }
}
.hero-kicker {
    font-family: 'Orbitron', sans-serif;
    font-size: 0.74rem;
    letter-spacing: 0.42em;
    text-transform: uppercase;
    color: var(--neon);
    text-shadow: 0 0 14px rgba(0, 229, 255, 0.75);
    margin-bottom: 6px;
}
.hero-title {
    font-family: 'Bebas Neue', 'Orbitron', sans-serif;
    font-size: clamp(2.9rem, 7.5vw, 5.4rem);
    line-height: 0.94;
    letter-spacing: 0.045em;
    margin: 0;
    background: linear-gradient(180deg, #fff8d6 0%, var(--fire-hot) 32%, var(--fire-mid) 58%, var(--fire-deep) 86%);
    -webkit-background-clip: text;
    background-clip: text;
    color: transparent;
    -webkit-text-fill-color: transparent;
    filter: drop-shadow(0 0 18px rgba(255, 106, 0, 0.55)) drop-shadow(0 0 44px rgba(255, 46, 99, 0.35));
    animation: heatFlicker 3.6s ease-in-out infinite;
}
@keyframes heatFlicker {
    0%, 100% { filter: drop-shadow(0 0 18px rgba(255, 106, 0, 0.55)) drop-shadow(0 0 44px rgba(255, 46, 99, 0.32)); }
    45%      { filter: drop-shadow(0 0 26px rgba(255, 160, 0, 0.85)) drop-shadow(0 0 66px rgba(255, 61, 0, 0.50)); }
    72%      { filter: drop-shadow(0 0 16px rgba(255, 106, 0, 0.50)) drop-shadow(0 0 38px rgba(255, 46, 99, 0.30)); }
}
.hero-sub {
    margin-top: 12px;
    font-size: 1.02rem;
    font-weight: 600;
    letter-spacing: 0.07em;
    color: #cdbfe8;
}
.hero-sub .step { color: var(--fire-hot); text-shadow: 0 0 10px rgba(255, 214, 10, 0.5); }
.hero-sub .arrow { color: var(--ember); margin: 0 6px; }
.hero-rule {
    margin-top: 18px;
    height: 3px;
    border-radius: 3px;
    background: linear-gradient(90deg, var(--fire-deep), var(--fire-hot), var(--neon), transparent);
    box-shadow: 0 0 18px rgba(255, 138, 0, 0.7);
    animation: rulePulse 2.8s ease-in-out infinite alternate;
}
@keyframes rulePulse {
    from { opacity: 0.65; }
    to   { opacity: 1; }
}

/* ---------- Typography ---------- */
h1, h2, h3, h4 {
    font-family: 'Orbitron', 'Bebas Neue', sans-serif !important;
    color: #fff1e0 !important;
    letter-spacing: 0.03em;
}
[data-testid="stMarkdownContainer"] h3,
[data-testid="stHeadingWithActionElements"] h3 {
    position: relative;
    padding-left: 16px;
    text-shadow: 0 0 22px rgba(255, 138, 0, 0.45);
}
[data-testid="stMarkdownContainer"] h3::before,
[data-testid="stHeadingWithActionElements"] h3::before {
    content: "";
    position: absolute;
    left: 0; top: 12%;
    height: 76%;
    width: 5px;
    border-radius: 4px;
    background: linear-gradient(180deg, var(--fire-hot), var(--fire-deep));
    box-shadow: 0 0 14px rgba(255, 138, 0, 0.9);
}
p, label, span, li { color: #ded4f2; }

/* ---------- Buttons ---------- */
.stButton > button, .stDownloadButton > button {
    font-family: 'Orbitron', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: 0.10em;
    text-transform: uppercase;
    font-size: 0.80rem !important;
    color: #fff4e2 !important;
    background: linear-gradient(135deg, rgba(60, 22, 60, 0.92), rgba(24, 14, 38, 0.92)) !important;
    border: 1px solid rgba(255, 138, 0, 0.55) !important;
    border-radius: 12px !important;
    padding: 0.60rem 1.1rem !important;
    position: relative;
    overflow: hidden;
    transition: transform 0.16s ease, box-shadow 0.22s ease, border-color 0.22s ease;
    box-shadow: 0 6px 22px rgba(0, 0, 0, 0.45), 0 0 0 1px rgba(255, 255, 255, 0.03) inset;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-2px) scale(1.015);
    border-color: var(--fire-hot) !important;
    color: #fffdf5 !important;
    box-shadow: 0 10px 30px rgba(255, 61, 0, 0.40), 0 0 26px rgba(255, 214, 10, 0.35);
}
.stButton > button::after, .stDownloadButton > button::after {
    content: "";
    position: absolute;
    top: 0; left: -120%;
    width: 60%; height: 100%;
    background: linear-gradient(90deg, transparent, rgba(255, 226, 150, 0.42), transparent);
    transform: skewX(-22deg);
    transition: left 0.55s ease;
}
.stButton > button:hover::after, .stDownloadButton > button:hover::after { left: 130%; }

/* Primary = full inferno */
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--fire-deep), var(--fire-mid) 55%, var(--fire-hot)) !important;
    border: 1px solid #ffe08a !important;
    color: #2a0d00 !important;
    text-shadow: 0 1px 0 rgba(255, 255, 255, 0.35);
    animation: infernoPulse 2.2s ease-in-out infinite;
}
@keyframes infernoPulse {
    0%, 100% { box-shadow: 0 0 22px rgba(255, 94, 0, 0.55), 0 8px 26px rgba(0,0,0,0.45); }
    50%      { box-shadow: 0 0 44px rgba(255, 160, 0, 0.90), 0 0 76px rgba(255, 61, 0, 0.45), 0 8px 26px rgba(0,0,0,0.45); }
}
.stButton > button[kind="primary"]:hover {
    color: #1c0800 !important;
    transform: translateY(-2px) scale(1.03);
}

/* ---------- Panels: uploader, alerts, expanders, tables ---------- */
[data-testid="stFileUploader"] {
    background: var(--panel);
    border: 1.5px dashed rgba(255, 138, 0, 0.55);
    border-radius: 18px;
    padding: 12px 16px;
    backdrop-filter: blur(9px);
    transition: border-color 0.25s ease, box-shadow 0.25s ease;
    box-shadow: 0 0 34px rgba(255, 61, 0, 0.10) inset;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--fire-hot);
    box-shadow: 0 0 42px rgba(255, 138, 0, 0.28), 0 0 34px rgba(255, 61, 0, 0.12) inset;
}
[data-testid="stFileUploaderDropzone"] { background: transparent !important; }
[data-testid="stFileUploaderDropzoneInstructions"],
[data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stFileUploaderDropzoneInstructions"] small,
[data-testid="stFileUploaderDropzoneInstructions"] div {
    color: #d9cbf5 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] svg { fill: var(--fire-mid) !important; }
[data-testid="stFileUploader"] button {
    background: linear-gradient(135deg, var(--fire-deep), var(--fire-mid) 60%, var(--fire-hot)) !important;
    color: #2a0d00 !important;
    border: 1px solid #ffe08a !important;
    font-family: 'Orbitron', sans-serif !important;
    font-weight: 800 !important;
    letter-spacing: 0.10em;
    text-transform: uppercase;
    border-radius: 11px !important;
    box-shadow: 0 0 22px rgba(255, 122, 0, 0.55);
}
[data-testid="stFileUploader"] button:hover {
    color: #1c0800 !important;
    box-shadow: 0 0 34px rgba(255, 170, 0, 0.85);
}
[data-testid="stFileUploader"] button p,
[data-testid="stFileUploader"] button span,
[data-testid="stFileUploader"] button div { color: #2a0d00 !important; }

[data-testid="stAlert"] {
    border-radius: 14px;
    border-left: 5px solid var(--fire-mid);
    background: linear-gradient(100deg, rgba(30, 18, 48, 0.90), rgba(14, 10, 26, 0.86)) !important;
    backdrop-filter: blur(8px);
    box-shadow: 0 10px 34px rgba(0, 0, 0, 0.42);
    color: #efe6ff !important;
}
[data-testid="stAlert"] p { color: #efe6ff !important; }

[data-testid="stExpander"] {
    border: 1px solid var(--edge) !important;
    border-radius: 14px !important;
    background: var(--panel) !important;
    backdrop-filter: blur(8px);
}

[data-testid="stDataFrame"], [data-testid="stTable"] {
    border: 1px solid var(--edge);
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 14px 44px rgba(0, 0, 0, 0.48), 0 0 26px rgba(255, 61, 0, 0.10);
}

/* Dark grid palette. Flagged cells set their own dark text in the styler. */
[data-testid="stDataFrame"] {
    --gdg-bg-cell: #100b20;
    --gdg-bg-cell-medium: #17102b;
    --gdg-bg-header: #1f1436;
    --gdg-bg-header-hovered: #2c1c4b;
    --gdg-bg-header-has-focus: #33205a;
    --gdg-text-dark: #f2ecff;
    --gdg-text-medium: #cbbde8;
    --gdg-text-light: #9a8cba;
    --gdg-text-header: #ffcf94;
    --gdg-text-header-selected: #fff6e6;
    --gdg-border-color: rgba(255, 138, 0, 0.20);
    --gdg-horizontal-border-color: rgba(255, 138, 0, 0.12);
    --gdg-accent-color: var(--fire-mid);
    --gdg-accent-light: rgba(255, 138, 0, 0.16);
    --gdg-accent-fg: #2a0d00;
    --gdg-bg-bubble: #1c1233;
    --gdg-bg-bubble-selected: #2c1c4b;
    --gdg-bg-search-result: rgba(255, 214, 10, 0.30);
}

/* Inputs */
.stSelectbox div[data-baseweb="select"] > div,
.stMultiSelect div[data-baseweb="select"] > div,
.stTextInput input, .stDateInput input {
    background: rgba(16, 11, 30, 0.88) !important;
    border: 1px solid rgba(255, 138, 0, 0.32) !important;
    border-radius: 11px !important;
    color: #f4ecff !important;
}
.stSelectbox div[data-baseweb="select"] > div:hover,
.stMultiSelect div[data-baseweb="select"] > div:hover {
    border-color: var(--fire-hot) !important;
    box-shadow: 0 0 18px rgba(255, 138, 0, 0.30);
}
.stMultiSelect span[data-baseweb="tag"] {
    background: linear-gradient(135deg, var(--fire-deep), var(--fire-mid)) !important;
    color: #2a0d00 !important;
    font-weight: 700;
    border-radius: 8px !important;
}

/* Legend chips */
.legend-chip {
    padding: 13px 8px;
    border-radius: 12px;
    text-align: center;
    color: #14060a;
    font-family: 'Orbitron', sans-serif;
    font-weight: 800;
    font-size: 0.72rem;
    letter-spacing: 0.06em;
    border: 1.5px solid rgba(255, 255, 255, 0.55);
    box-shadow: 0 8px 22px rgba(0, 0, 0, 0.45);
    transition: transform 0.18s ease, box-shadow 0.18s ease;
}
.legend-chip:hover {
    transform: translateY(-3px) scale(1.03);
    box-shadow: 0 12px 30px rgba(0, 0, 0, 0.55), 0 0 26px currentColor;
}

/* Scrollbar */
::-webkit-scrollbar { width: 11px; height: 11px; }
::-webkit-scrollbar-track { background: #0a0716; }
::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg, var(--fire-mid), var(--fire-deep));
    border-radius: 8px;
    border: 2px solid #0a0716;
}
::-webkit-scrollbar-thumb:hover { background: linear-gradient(180deg, var(--fire-hot), var(--fire-mid)); }

@media (prefers-reduced-motion: reduce) {
    .stApp::before, .ember, .fire-floor, .hero::before,
    .hero-title, .hero-rule, .stButton > button[kind="primary"] { animation: none !important; }
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


def resolve_numeric_date_parts(raw: str) -> Optional[Tuple[int, int, int]]:
    """
    Decide (day, month, year) from a numeric date string, without checking that
    the calendar date actually exists.
    - Default numeric day/month input to DD-MM-YYYY.
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
    elif p2 > 12:
        # Obvious MM-DD-YYYY input (e.g. 5/15/1991) -> convert to DD-MM-YYYY.
        month, day, year = p1, p2, p3
    else:
        day, month, year = p1, p2, p3

    if year < 100:
        year += 2000 if year < 30 else 1900

    return day, month, year


def parse_numeric_date(raw: str) -> Optional[pd.Timestamp]:
    """Parse numeric date parts into a timestamp (NaT if the date is impossible)."""
    parts = resolve_numeric_date_parts(raw)
    if parts is None:
        return None
    day, month, year = parts
    try:
        return pd.Timestamp(year=year, month=month, day=day)
    except ValueError:
        return pd.NaT


def clean_dob(raw: str) -> Tuple[str, Optional[str]]:
    if is_blank(raw):
        return "", None

    value = clean_spaces(strip_formula(raw))
    value = value.strip("'\"")

    ordinal_fixed = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", value, flags=re.IGNORECASE)

    # Handle pure numeric Excel serial date.
    if re.fullmatch(r"\d{5}(\.\d+)?", ordinal_fixed):
        serial = float(ordinal_fixed)
        dt = pd.to_datetime(serial, origin="1899-12-30", unit="D", errors="coerce")
    else:
        dt = parse_numeric_date(ordinal_fixed)
        if dt is None:
            dt = pd.to_datetime(ordinal_fixed, dayfirst=True, errors="coerce")

    if pd.isna(dt):
        # Impossible calendar date: still reorder to DD-MM-YYYY so an obvious
        # MM-DD-YYYY input (e.g. 2-30-2025) is shown as 30-02-2025, flagged red.
        parts = resolve_numeric_date_parts(ordinal_fixed)
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


def first_non_blank_value(row: pd.Series, columns: List[str]) -> str:
    for col in columns:
        value = row[col]
        if not is_blank(value):
            return value
    return ""


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


def worksheet_to_dataframe(worksheet) -> pd.DataFrame:
    headers: Optional[List[str]] = None
    data: List[List[str]] = []
    for row in worksheet.iter_rows():
        cells = [cell_to_text(cell) for cell in row]
        if headers is None:
            headers = cells
        else:
            data.append(cells)

    if not headers:
        return pd.DataFrame()

    columns = dedupe_headers(headers)
    normalized = [
        (values + [""] * (len(columns) - len(values)))[: len(columns)]
        for values in data
    ]
    return pd.DataFrame(normalized, columns=columns).fillna("")


def read_xlsx_all_sheets(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    sheets: Dict[str, pd.DataFrame] = {}
    for sheet_name in workbook.sheetnames:
        sheets[sheet_name] = worksheet_to_dataframe(workbook[sheet_name])
    workbook.close()
    return sheets


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
    """Return one source entry per CSV file or Excel sheet."""
    lower = filename.lower()
    sources: List[Dict[str, object]] = []
    file_id = hashlib.md5(file_bytes).hexdigest()

    if lower.endswith(".csv"):
        uploaded_file = io.BytesIO(file_bytes)
        df = pd.read_csv(uploaded_file, dtype=str, keep_default_na=False).fillna("")
        sources.append(make_source_entry(file_id, filename, "", df, True))
        return sources

    if lower.endswith(".xlsx"):
        sheets = read_xlsx_all_sheets(file_bytes)
        for sheet_name, df in sheets.items():
            keep = not looks_like_summary_sheet(sheet_name, df)
            sources.append(make_source_entry(file_id, filename, sheet_name, df, keep))
        return sources

    if lower.endswith(".xls"):
        uploaded_file = io.BytesIO(file_bytes)
        try:
            excel = pd.ExcelFile(uploaded_file)
            for sheet_name in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name=sheet_name, dtype=str).fillna("")
                keep = not looks_like_summary_sheet(sheet_name, df)
                sources.append(make_source_entry(file_id, filename, sheet_name, df, keep))
            return sources
        except Exception:
            uploaded_file.seek(0)
            html = uploaded_file.read().decode("utf-8", errors="replace")
            df = pd.read_html(io.StringIO(html), header=0)[0]
            if len(df) > 0 and str(df.iloc[0, 0]).strip().lower() in {"first name", "name"}:
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)
            df = df.astype(str).fillna("")
            sources.append(make_source_entry(file_id, filename, "", df, True))
            return sources

    raise ValueError("Unsupported file format")


@st.cache_data(show_spinner=False)
def cached_column_value_counts(
    file_id: str, sheet_name: str, col_name: str, file_bytes: bytes, filename: str
) -> Tuple[Tuple[str, int], ...]:
    """Cached unique value counts for one column (avoids re-scanning 20k+ rows)."""
    for source in discover_file_sources(file_bytes, filename):
        if str(source.get("sheet_name") or "") != (sheet_name or ""):
            continue
        df = source["df"]  # type: ignore[assignment]
        if col_name not in df.columns:
            return tuple()
        series = df[col_name].fillna("").astype(str).str.strip().replace("", ".")
        counts = series.value_counts()
        return tuple((str(k), int(v)) for k, v in counts.items())
    return tuple()


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


def render_error_legend():
    st.markdown("**Error Color Legend**")
    legend_cols = st.columns(len(ERROR_LEGEND))
    for col, (_, info) in zip(legend_cols, ERROR_LEGEND.items()):
        with col:
            st.markdown(
                f'<div class="legend-chip" style="background-color:#{info["hex"]}; '
                f'color:#{info["hex"]};">'
                f'<span style="color:#14060a;">{info["label"]}</span></div>',
                unsafe_allow_html=True,
            )


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

        for i in range(len(out)):
            full_raw = out.at[i, full_name_col] if full_name_col else ""
            first_raw = out.at[i, first_name_col] if first_name_col else ""
            last_raw = out.at[i, last_name_col] if last_name_col else ""
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
            for i in range(len(out)):
                raw = out.at[i, src_col]
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
        dob_values = []
        age_values = []
        for i in range(len(out)):
            raw = first_non_blank_value(out.iloc[i], dob_cols)
            cleaned, flag = clean_dob(raw)
            dob_values.append(cleaned)
            age_values.append(calculate_age_from_dob(cleaned, age_as_on))
            if is_blank(raw):
                mark_empty(i, "Date of Birth")
            mark_cell(i, "Date of Birth", flag)
        out["Date of Birth"] = dob_values
        out["Age"] = age_values

    # Emergency name: same first/last/full rules as attendee name.
    emergency_name_col = mapping.get("emergency_name")
    if emergency_name_col:
        em_first, em_last, em_full = [], [], []
        for i in range(len(out)):
            first, last, full = split_name(out.at[i, emergency_name_col])
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
        for i in range(len(out)):
            raw = out.at[i, cat_col]
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
        for i in range(len(out)):
            raw = out.at[i, tshirt_col]
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
        for i in range(len(out)):
            cleaned, flag = clean_country(out.at[i, country_col])
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

    new_flags: Dict[Tuple[int, str], str] = {}
    for new_idx, orig_idx in enumerate(orig_indices):
        for col_name in df.columns:
            key = (orig_idx, col_name)
            if key in cell_flags:
                new_flags[(new_idx, col_name)] = cell_flags[key]
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
            for row_idx in range(len(df)):
                excel_row = row_idx + 2
                raw_val = df.iloc[row_idx]["Date of Birth"]
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
            for row_idx in range(len(df)):
                excel_row = row_idx + 2
                raw_val = clean_spaces(df.iloc[row_idx][col_name])
                if re.fullmatch(r"\d{10}", raw_val):
                    cell = ws.cell(row=excel_row, column=col_idx)
                    cell.value = int(raw_val)
                    cell.number_format = "0"

        # Write age as numeric cells to avoid Excel text warnings.
        if age_col_idx and "Age" in df.columns:
            for row_idx in range(len(df)):
                excel_row = row_idx + 2
                raw_val = clean_spaces(df.iloc[row_idx]["Age"])
                if re.fullmatch(r"\d+", raw_val):
                    cell = ws.cell(row=excel_row, column=age_col_idx)
                    cell.value = int(raw_val)
                    cell.number_format = "0"
    output.seek(0)
    return output.read()


st.markdown(
    """
    <div class="hero">
        <div class="hero-kicker">Sports Timing Solutions</div>
        <h1 class="hero-title">FileSort Cleaner</h1>
        <div class="hero-sub">
            <span class="step">Upload</span><span class="arrow">&#10148;</span>
            <span class="step">Keep</span><span class="arrow">&#10148;</span>
            <span class="step">Map</span><span class="arrow">&#10148;</span>
            <span class="step">Clean</span><span class="arrow">&#10148;</span>
            <span class="step">Download</span>
        </div>
        <div class="hero-rule"></div>
    </div>
    """,
    unsafe_allow_html=True,
)

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
    keys_to_persist = [
        "loaded_files",
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


def build_value_breakdown(
    sources: List[Dict[str, object]],
    source_mappings: Dict[str, Dict[str, Optional[str]]],
    field_key: str,
    loaded_files: Dict[str, Dict[str, object]],
) -> pd.DataFrame:
    """Unique values per source for category/tshirt (cached column scans)."""
    rows: List[Dict[str, object]] = []
    for source in sources:
        sid = str(source["source_id"])
        col_name = source_mappings.get(sid, {}).get(field_key)
        if not col_name:
            continue
        file_id = str(source["file_id"])
        meta = loaded_files.get(file_id)
        if not meta:
            continue
        counts = cached_column_value_counts(
            file_id,
            str(source.get("sheet_name") or ""),
            col_name,
            meta["file_bytes"],  # type: ignore[arg-type]
            str(meta["file_name"]),
        )
        label = source_label(source)
        for raw_value, count in counts:
            rows.append({"Source": label, "Raw Value": raw_value, "Rows": count})
    if not rows:
        return pd.DataFrame(columns=["Source", "Raw Value", "Rows", "Mapped Value"])
    out = pd.DataFrame(rows).sort_values(["Source", "Raw Value"]).reset_index(drop=True)
    out["Mapped Value"] = out["Raw Value"]
    return out


def render_value_map_editor(
    title: str,
    breakdown: pd.DataFrame,
    state_map_key: str,
    editor_key: str,
) -> Dict[str, str]:
    """Safe lightweight mapper: store only raw->mapped dict (no df session writes)."""
    if state_map_key not in st.session_state:
        st.session_state[state_map_key] = {}

    saved_map: Dict[str, str] = dict(st.session_state[state_map_key])
    display = breakdown.copy()
    display["Mapped Value"] = [
        saved_map.get(str(raw), str(raw)) for raw in display["Raw Value"].tolist()
    ]
    fp = hashlib.md5(
        "|".join(f"{a}:{b}:{c}" for a, b, c in zip(
            display["Source"].astype(str),
            display["Raw Value"].astype(str),
            display["Rows"].astype(str),
        )).encode("utf-8")
    ).hexdigest()[:10]

    st.markdown(f"**{title}**")
    st.caption("Sources: " + " | ".join(sorted(display["Source"].unique().tolist())))
    edited = st.data_editor(
        display,
        hide_index=True,
        use_container_width=True,
        disabled=["Source", "Raw Value", "Rows"],
        column_config={
            "Mapped Value": st.column_config.TextColumn("Mapped Value"),
            "Rows": st.column_config.NumberColumn("Rows"),
        },
        key=f"{editor_key}_{fp}",
    )

    new_map: Dict[str, str] = {}
    for raw, mapped in zip(edited["Raw Value"].tolist(), edited["Mapped Value"].tolist()):
        raw_s = clean_spaces(str(raw))
        mapped_s = clean_spaces(str(mapped)) or "."
        if raw_s:
            new_map[raw_s] = mapped_s
    st.session_state[state_map_key] = new_map
    return new_map


restore_persisted_state()
if "upload_widget_nonce" not in st.session_state:
    st.session_state["upload_widget_nonce"] = 0
if "source_keep" not in st.session_state:
    st.session_state["source_keep"] = {}
if "loaded_files" not in st.session_state:
    st.session_state["loaded_files"] = {}

upload_key = f"source_upload_{st.session_state['upload_widget_nonce']}"
uploaded_files = st.file_uploader(
    "Upload CSV/XLSX/XLS (multiple files allowed)",
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
    if st.session_state.get("loaded_batch_id") != batch_id:
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
    # Initialize keep flags (default from heuristic; preserve existing choices).
    for source in all_sources:
        sid = str(source["source_id"])
        if sid not in st.session_state["source_keep"]:
            st.session_state["source_keep"][sid] = bool(source["default_keep"])

    top_left, top_right = st.columns([3, 1])
    with top_left:
        kept_sources_preview = [
            s for s in all_sources if st.session_state["source_keep"].get(str(s["source_id"]), False)
        ]
        kept_count = len(kept_sources_preview)
        total_runners = sum(int(s.get("runner_count", s["rows"])) for s in kept_sources_preview)
        st.success(
            f"Loaded {len(st.session_state['loaded_files'])} file(s) → "
            f"{len(all_sources)} sheet(s) | Keeping {kept_count} | "
            f"Selected runners ≈ {total_runners}"
        )
    with top_right:
        if st.button("Remove Loaded Files", type="secondary", use_container_width=True):
            clear_loaded_file_state()
            st.rerun()

    st.subheader("1) Analyse & Keep Files / Sheets")
    st.caption("Keep runner sheets. Ignore summary sheets like Count.")

    overview = pd.DataFrame(
        [
            {
                "File": str(s["file_name"]),
                "Sheet": str(s.get("sheet_name") or "(whole file)"),
                "Runners": int(s.get("runner_count", s["rows"])),
                "Rows": int(s["rows"]),
                "Cols": len(s["columns"]),  # type: ignore[arg-type]
                "Suggestion": "IGNORE" if not s["default_keep"] else "KEEP",
            }
            for s in all_sources
        ]
    ).sort_values(["Runners", "Rows"], ascending=[False, False])
    st.dataframe(overview, hide_index=True, use_container_width=True)

    q1, q2 = st.columns(2)
    with q1:
        if st.button("Use Suggestions", use_container_width=True):
            for s in all_sources:
                st.session_state["source_keep"][str(s["source_id"])] = bool(s["default_keep"])
            st.rerun()
    with q2:
        if st.button("Keep All", use_container_width=True):
            for s in all_sources:
                st.session_state["source_keep"][str(s["source_id"])] = True
            st.rerun()

    for source in all_sources:
        sid = str(source["source_id"])
        label = source_label(source)
        runners = int(source.get("runner_count", source["rows"]))
        tip = "summary" if not source["default_keep"] else "runners"
        keep = st.checkbox(
            f"{label} — {runners} runners ({tip})",
            value=bool(st.session_state["source_keep"].get(sid, source["default_keep"])),
            key=f"keep_{sid}",
        )
        st.session_state["source_keep"][sid] = keep

    kept_sources = [
        s for s in all_sources if st.session_state["source_keep"].get(str(s["source_id"]), False)
    ]
    if not kept_sources:
        st.warning("Keep at least one source before mapping.")
        save_persisted_state()
        st.stop()

    st.subheader("2) Map Columns (one source at a time)")
    source_labels = {source_label(s): s for s in kept_sources}
    selected_label = st.selectbox("Working on", list(source_labels.keys()), key="active_map_source_label")
    active_source = source_labels[selected_label]
    active_sid = str(active_source["source_id"])
    active_df = active_source["df"]  # type: ignore[assignment]
    active_cols = list(active_source["columns"])  # type: ignore[arg-type]
    active_guess = guess_mapping(active_cols)
    active_options = ["<None>"] + active_cols

    m1, m2, m3 = st.columns(3)
    m1.metric("Runners", int(active_source.get("runner_count", active_source["rows"])))
    m2.metric("Rows", int(active_source["rows"]))
    m3.metric("Columns", len(active_cols))

    with st.expander("Preview (first 8 rows)", expanded=False):
        st.dataframe(active_df.head(8), use_container_width=True)

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

    st.markdown("**Essential columns**")
    essential_ui = st.columns(2)
    for idx, field in enumerate(ESSENTIAL_MAPPING_FIELDS):
        key = map_key(active_sid, field)
        current = source_mappings[active_sid].get(field) or active_guess.get(field) or "<None>"
        if current not in active_options:
            current = "<None>"
        picked = essential_ui[idx % 2].selectbox(
            FIELD_LABELS[field],
            options=active_options,
            index=active_options.index(current),
            key=key,
        )
        source_mappings[active_sid][field] = None if picked == "<None>" else picked

    with st.expander("More columns (optional)", expanded=False):
        advanced_fields = [f for f in CANONICAL_FIELDS if f not in ESSENTIAL_MAPPING_FIELDS]
        adv_ui = st.columns(2)
        for idx, field in enumerate(advanced_fields):
            key = map_key(active_sid, field)
            current = source_mappings[active_sid].get(field) or active_guess.get(field) or "<None>"
            if current not in active_options:
                current = "<None>"
            picked = adv_ui[idx % 2].selectbox(
                FIELD_LABELS[field],
                options=active_options,
                index=active_options.index(current),
                key=key,
            )
            source_mappings[active_sid][field] = None if picked == "<None>" else picked

    for field in CANONICAL_FIELDS:
        picked = st.session_state.get(map_key(active_sid, field), "<None>")
        source_mappings[active_sid][field] = None if picked == "<None>" else picked

    status_rows = [
        {
            "Source": source_label(s),
            "Category": source_mappings[str(s["source_id"])].get("category") or "—",
            "T-Shirt": source_mappings[str(s["source_id"])].get("tshirt_size") or "—",
            "Phone": source_mappings[str(s["source_id"])].get("phone") or "—",
            "DOB": source_mappings[str(s["source_id"])].get("dob") or "—",
        }
        for s in kept_sources
    ]
    st.caption("Column map status")
    st.dataframe(pd.DataFrame(status_rows), hide_index=True, use_container_width=True)

    age_as_on = st.date_input(
        "Age calculation date (as on)",
        value=date.today(),
        format="DD-MM-YYYY",
        key="age_as_on_date",
    )

    st.subheader("3) Map Category & T-Shirt Values")
    st.caption("Values from every kept source. Edit Mapped Value only.")

    category_map: Dict[str, str] = {}
    tshirt_map: Dict[str, str] = {}
    loaded_files = st.session_state.get("loaded_files", {})

    try:
        cat_breakdown = build_value_breakdown(
            kept_sources, source_mappings, "category", loaded_files
        )
        if cat_breakdown.empty:
            st.info("Map a Category column in step 2 to edit category values.")
        else:
            category_map = render_value_map_editor(
                "Category values by source",
                cat_breakdown,
                "category_value_map",
                "category_value_editor",
            )
            if st.button("Reset Category Mappings", key="reset_category_maps"):
                st.session_state["category_value_map"] = {}
                st.rerun()
    except Exception as exc:
        st.error(f"Category mapping failed: {exc}")

    try:
        tshirt_breakdown = build_value_breakdown(
            kept_sources, source_mappings, "tshirt_size", loaded_files
        )
        if tshirt_breakdown.empty:
            st.info("Map a T-Shirt column in step 2 to edit T-shirt values.")
        else:
            tshirt_map = render_value_map_editor(
                "T-Shirt values by source",
                tshirt_breakdown,
                "tshirt_value_map",
                "tshirt_value_editor",
            )
            if st.button("Reset T-Shirt Mappings", key="reset_tshirt_maps"):
                st.session_state["tshirt_value_map"] = {}
                st.rerun()
    except Exception as exc:
        st.error(f"T-Shirt mapping failed: {exc}")

    st.subheader("4) Clean & Download")
    render_error_legend()

    if st.button("Run Cleaning", type="primary"):
        parts: List[Tuple[pd.DataFrame, Dict[Tuple[int, str], str], str, str]] = []
        with st.spinner(f"Cleaning {len(kept_sources)} source(s)..."):
            for source in kept_sources:
                sid = str(source["source_id"])
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
            cleaned_df, cell_flags = merge_cleaned_frames(parts)

        error_counts = count_errors_by_type(cell_flags)
        if error_counts:
            summary_parts = [
                f"{ERROR_LEGEND[error_type]['label']}: {count}"
                for error_type, count in sorted(error_counts.items())
                if error_type in ERROR_LEGEND
            ]
            st.success(
                f"Cleaning completed for {len(kept_sources)} source(s), "
                f"{len(cleaned_df)} total rows. Flagged cells -> {' | '.join(summary_parts)}"
            )
        else:
            st.success(
                f"Cleaning completed for {len(kept_sources)} source(s), "
                f"{len(cleaned_df)} total rows. No flagged errors found."
            )

        render_error_legend()
        preview_df = cleaned_df.head(20)
        if cell_flags:
            st.dataframe(style_preview_dataframe(preview_df, cell_flags), use_container_width=True)
        else:
            st.dataframe(preview_df, use_container_width=True)

        excel_bytes = df_to_excel_with_highlight(cleaned_df, cell_flags)
        st.download_button(
            label="Download Cleaned Excel",
            data=excel_bytes,
            file_name="cleaned_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    save_persisted_state()
else:
    st.info(
        "Upload one or more files to start. Multi-sheet Excel files show Keep/Ignore per sheet. "
        "Loaded progress stays until you click 'Remove Loaded Files'."
    )
