import io
import json
import os
import re
import tempfile
import zipfile
from hashlib import md5
from pathlib import Path
from urllib.request import urlopen

import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from streamlit_autorefresh import st_autorefresh
except Exception:
    st_autorefresh = None


st.set_page_config(
    page_title="NQAS Internal Assessment Dashboard",
    layout="wide",
    page_icon="bar_chart",
)


DEFAULT_LOCAL_DATA_ROOT = Path(
    r"C:\Users\hari\National Health Systems Resource Centre\Shraiya Srivastava - NQAS Internal Assessors' Data"
)
SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv"}

MAX_SCORE = 40.0
PASS_MARKS = 24.0

REQUIRED_OUTPUT_COLUMNS = [
    "Sno.",
    "State",
    "district",
    "Aspirational Block Name",
    "Name",
    "Designation",
    "Place of posting",
    "Mobile no.",
    "email",
    "score",
    "Percentage",
    "result",
    "Qualified or not qualified",
]

EXTRA_OUTPUT_COLUMNS = [
    "Certificate No.",
    "Assessment Year",
    "Certificate State Code",
    "Batch No",
    "Participant Serial",
    "Source File",
    "Source Sheet",
]

CERTIFICATE_REGEX = re.compile(
    r"IA/(?P<year>20\d{2})/(?P<state_code>[A-Z]{1,3})(?P<batch>\d{0,3})/(?P<serial>\d+)",
    re.IGNORECASE,
)

STATE_NAMES_36 = [
    "Andhra Pradesh",
    "Arunachal Pradesh",
    "Assam",
    "Bihar",
    "Chhattisgarh",
    "Goa",
    "Gujarat",
    "Haryana",
    "Himachal Pradesh",
    "Jharkhand",
    "Karnataka",
    "Kerala",
    "Madhya Pradesh",
    "Maharashtra",
    "Manipur",
    "Meghalaya",
    "Mizoram",
    "Nagaland",
    "Odisha",
    "Punjab",
    "Rajasthan",
    "Sikkim",
    "Tamil Nadu",
    "Telangana",
    "Tripura",
    "Uttar Pradesh",
    "Uttarakhand",
    "West Bengal",
    "Andaman and Nicobar Islands",
    "Chandigarh",
    "Dadra and Nagar Haveli and Daman and Diu",
    "Delhi",
    "Jammu and Kashmir",
    "Ladakh",
    "Lakshadweep",
    "Puducherry",
]

STATE_ALIASES = {
    "Andhra Pradesh": ["andhra pradesh", "andra pradesh", "andhra", "ap"],
    "Arunachal Pradesh": ["arunachal pradesh", "arunachal", "ar"],
    "Assam": ["assam", "as"],
    "Bihar": ["bihar", "br"],
    "Chhattisgarh": ["chhattisgarh", "chattisgarh", "cg"],
    "Goa": ["goa", "ga"],
    "Gujarat": ["gujarat", "gujurat", "gj"],
    "Haryana": ["haryana", "hr"],
    "Himachal Pradesh": ["himachal pradesh", "himachal", "hp"],
    "Jharkhand": ["jharkhand", "jh"],
    "Karnataka": ["karnataka", "ka"],
    "Kerala": ["kerala", "kl"],
    "Madhya Pradesh": ["madhya pradesh", "mp"],
    "Maharashtra": ["maharashtra", "mh"],
    "Manipur": ["manipur", "mn"],
    "Meghalaya": ["meghalaya", "ml"],
    "Mizoram": ["mizoram", "mz"],
    "Nagaland": ["nagaland", "nl"],
    "Odisha": ["odisha", "orissa", "od"],
    "Punjab": ["punjab", "pb"],
    "Rajasthan": ["rajasthan", "raj", "rj"],
    "Sikkim": ["sikkim", "sk"],
    "Tamil Nadu": ["tamil nadu", "tn"],
    "Telangana": ["telangana", "telengana", "tg", "ts"],
    "Tripura": ["tripura", "tr"],
    "Uttar Pradesh": ["uttar pradesh", "up"],
    "Uttarakhand": ["uttarakhand", "uttaranchal", "uk", "ua"],
    "West Bengal": ["west bengal", "wb", "bengal"],
    "Andaman and Nicobar Islands": [
        "andaman and nicobar islands",
        "andaman nicobar",
        "andaman",
        "an",
    ],
    "Chandigarh": ["chandigarh", "ch"],
    "Dadra and Nagar Haveli and Daman and Diu": [
        "dadra and nagar haveli and daman and diu",
        "dadra nagar haveli daman diu",
        "daman and diu",
        "daman diu",
        "dnh dd",
        "dnh and dd",
        "dn",
    ],
    "Delhi": ["delhi", "new delhi", "nct delhi", "dl"],
    "Jammu and Kashmir": [
        "jammu and kashmir",
        "jammu kashmir",
        "j and k",
        "j k",
        "jammu",
        "jk",
    ],
    "Ladakh": ["ladakh", "la"],
    "Lakshadweep": ["lakshadweep", "ld"],
    "Puducherry": ["puducherry", "pondicherry", "py"],
}

STATE_CODE_TO_STATE = {
    "AP": "Andhra Pradesh",
    "AR": "Arunachal Pradesh",
    "AS": "Assam",
    "BR": "Bihar",
    "CG": "Chhattisgarh",
    "GA": "Goa",
    "GJ": "Gujarat",
    "HR": "Haryana",
    "HP": "Himachal Pradesh",
    "JH": "Jharkhand",
    "KA": "Karnataka",
    "KL": "Kerala",
    "MP": "Madhya Pradesh",
    "MH": "Maharashtra",
    "MN": "Manipur",
    "ML": "Meghalaya",
    "MZ": "Mizoram",
    "NL": "Nagaland",
    "OD": "Odisha",
    "PB": "Punjab",
    "RJ": "Rajasthan",
    "SK": "Sikkim",
    "TN": "Tamil Nadu",
    "TG": "Telangana",
    "TS": "Telangana",
    "TR": "Tripura",
    "UP": "Uttar Pradesh",
    "UK": "Uttarakhand",
    "UA": "Uttarakhand",
    "WB": "West Bengal",
    "AN": "Andaman and Nicobar Islands",
    "CH": "Chandigarh",
    "DD": "Dadra and Nagar Haveli and Daman and Diu",
    "DN": "Dadra and Nagar Haveli and Daman and Diu",
    "DL": "Delhi",
    "JK": "Jammu and Kashmir",
    "LA": "Ladakh",
    "LD": "Lakshadweep",
    "PY": "Puducherry",
}

COLUMN_ALIASES = {
    "Sno.": [
        "sno",
        "s no",
        "s no.",
        "sr no",
        "sr. no",
        "serial no",
        "serial number",
        "sl no",
        "sl. no",
    ],
    "State": ["state", "name of state", "state name", "state ut", "state/ut"],
    "district": [
        "district",
        "district name",
        "name of district",
        "district facility",
        "district / facility",
    ],
    "Aspirational Block Name": [
        "aspirational block name",
        "aspirational block",
        "name of aspirational block",
        "block name",
    ],
    "Name": [
        "name",
        "participant name",
        "participants name",
        "name of the participant",
        "assessor name",
        "participant",
    ],
    "Designation": ["designation", "post", "cadre"],
    "Place of posting": [
        "place of posting",
        "name of facility",
        "facility name",
        "place posting",
        "posting place",
        "facility",
    ],
    "Mobile no.": [
        "mobile no",
        "mobile number",
        "mobile",
        "phone number",
        "phone no",
        "contact number",
        "contact no",
    ],
    "email": ["email", "email id", "e mail id", "e-mail", "mail id"],
    "score": [
        "score",
        "marks",
        "marks obtained",
        "obtained marks",
        "score out of 40",
    ],
    "Total Marks": ["total marks", "maximum marks", "max marks", "marks out of"],
    "Percentage": ["percentage", "percent", "score %", "% score", "percentage %"],
    "result": ["result", "status", "remark", "remarks"],
    "Certificate No.": [
        "certificate",
        "certificate no",
        "certificate number",
        "certificates",
        "cert no",
    ],
    "Assessment Year": ["year", "assessment year", "batch year"],
}

HEADER_SIGNALS = [
    "s no",
    "serial no",
    "name",
    "designation",
    "district",
    "mobile",
    "email",
    "score",
    "marks obtained",
    "percentage",
    "result",
    "certificate",
]

GEOJSON_URLS = [
    "https://gist.githubusercontent.com/jbrobst/56c13bbbf9d97d187fea01ca62ea5112/raw/e388c4cae20aa53cb5090210a42ebb9b765c0a36/india_states.geojson",
    "https://raw.githubusercontent.com/geohacker/india/master/state/india_state.geojson",
]


def normalize_text(value):
    if pd.isna(value):
        text = ""
    else:
        text = str(value).strip().lower()
    text = re.sub(r"[&/_\-]+", " ", text)
    text = re.sub(r"[^a-z0-9.% ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clean_string_series(series):
    return (
        series.astype("string")
        .fillna("")
        .str.replace("\u00a0", " ", regex=False)
        .str.strip()
    )


ALIAS_TO_STATE = {}
for state_name in STATE_NAMES_36:
    ALIAS_TO_STATE[normalize_text(state_name)] = state_name
for state_name, aliases in STATE_ALIASES.items():
    for alias in aliases:
        ALIAS_TO_STATE[normalize_text(alias)] = state_name


def infer_state_from_text(text_value):
    text = normalize_text(text_value)
    if not text:
        return "Unknown"
    best_match = None
    best_size = -1
    for alias, state_name in ALIAS_TO_STATE.items():
        pattern = rf"\b{re.escape(alias)}\b"
        if re.search(pattern, text) and len(alias) > best_size:
            best_match = state_name
            best_size = len(alias)
    return best_match if best_match else "Unknown"


def normalize_state_value(value):
    if pd.isna(value):
        value = ""
    else:
        value = str(value).strip()
    if not value:
        return "Unknown"
    inferred = infer_state_from_text(value)
    if inferred != "Unknown":
        return inferred
    return "Unknown"


def canonicalize_certificate_text(value):
    if pd.isna(value):
        text = ""
    else:
        text = str(value)
    text = text.upper().replace("\\", "/")
    text = re.sub(r"\s+", "", text)
    return text


def parse_numeric(series):
    cleaned = clean_string_series(series).str.replace(",", "", regex=False)
    cleaned = cleaned.str.replace(r"[^0-9.\-]", "", regex=True)
    cleaned = cleaned.replace("", pd.NA)
    return pd.to_numeric(cleaned, errors="coerce")


def parse_percentage(series):
    pct = clean_string_series(series).str.replace(",", "", regex=False)
    pct = pct.str.replace("%", "", regex=False)
    pct = pct.str.replace(r"[^0-9.\-]", "", regex=True).replace("", pd.NA)
    pct = pd.to_numeric(pct, errors="coerce")
    pct = pct.where(pct > 1, pct * 100)
    return pct


def parse_result_bool(series):
    normalized = clean_string_series(series).str.lower()
    pass_mask = normalized.str.contains(r"\b(?:pass|qualified|q)\b", regex=True)
    fail_mask = normalized.str.contains(r"\b(?:fail|not qualified|nq)\b", regex=True)
    output = pd.Series(pd.NA, index=series.index, dtype="boolean")
    output.loc[pass_mask] = True
    output.loc[fail_mask] = False
    return output


def canonical_certificate(series):
    cert = clean_string_series(series).map(canonicalize_certificate_text)
    cert = cert.replace("", pd.NA)
    parsed = cert.str.extract(CERTIFICATE_REGEX)
    return cert, parsed


def infer_year_from_path(file_path):
    stem_year = re.search(r"(20\d{2})", file_path.stem)
    if stem_year:
        return int(stem_year.group(1))
    for part in reversed(file_path.parts):
        match = re.search(r"(20\d{2})", part)
        if match:
            return int(match.group(1))
    return int(pd.Timestamp(file_path.stat().st_mtime, unit="s").year)


def is_supported_data_file(file_path):
    if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return False
    if file_path.name.startswith("~$"):
        return False
    return True


def list_input_files(root_path):
    return sorted(
        [
            file_path
            for file_path in root_path.rglob("*")
            if file_path.is_file() and is_supported_data_file(file_path)
        ]
    )


def build_data_signature(root_path):
    entries = []
    for file_path in list_input_files(root_path):
        stats = file_path.stat()
        entries.append(f"{file_path}|{stats.st_size}|{stats.st_mtime_ns}")
    return md5("\n".join(entries).encode("utf-8")).hexdigest()


def detect_header_row(preview_df):
    best_row = None
    best_score = 0
    max_rows = min(len(preview_df), 60)
    for idx in range(max_rows):
        row_values = [
            normalize_text(cell)
            for cell in preview_df.iloc[idx].tolist()
            if normalize_text(cell)
        ]
        if len(row_values) < 3:
            continue
        row_score = 0
        for signal in HEADER_SIGNALS:
            if any(signal in cell for cell in row_values):
                row_score += 1
        if row_score > best_score:
            best_score = row_score
            best_row = idx
    if best_row is None or best_score < 3:
        return None
    return best_row


def candidate_header_rows(preview_df):
    candidates = []
    primary = detect_header_row(preview_df)
    if primary is not None:
        candidates.append(primary)

    max_rows = min(len(preview_df), 40)
    for idx in range(max_rows):
        row_values = [
            normalize_text(cell)
            for cell in preview_df.iloc[idx].tolist()
            if normalize_text(cell)
        ]
        if len(row_values) < 2:
            continue
        row_score = 0
        for signal in HEADER_SIGNALS:
            if any(signal in cell for cell in row_values):
                row_score += 1
        if row_score >= 3:
            candidates.append(idx)

    candidates.append(0)
    deduped = []
    seen = set()
    for row_id in candidates:
        if row_id not in seen and 0 <= row_id < max(len(preview_df), 1):
            deduped.append(row_id)
            seen.add(row_id)
    return deduped[:12]


def find_matching_column(columns, aliases, avoid_percent=False):
    normalized_columns = {col: normalize_text(col) for col in columns}
    normalized_aliases = [normalize_text(alias) for alias in aliases]

    for alias in normalized_aliases:
        for column_name, norm_name in normalized_columns.items():
            if norm_name == alias:
                if avoid_percent and ("percent" in norm_name or "%" in str(column_name)):
                    continue
                return column_name

    for alias in normalized_aliases:
        for column_name, norm_name in normalized_columns.items():
            if alias and alias in norm_name:
                if avoid_percent and ("percent" in norm_name or "%" in str(column_name)):
                    continue
                return column_name
    return None


def _select_best_column(normalized_columns, used_columns, exact_aliases=None, contains_aliases=None, excludes=None):
    exact_aliases = [normalize_text(a) for a in (exact_aliases or [])]
    contains_aliases = [normalize_text(a) for a in (contains_aliases or [])]
    excludes = [normalize_text(a) for a in (excludes or [])]

    best_col = None
    best_score = 0
    for col_name, norm_name in normalized_columns.items():
        if col_name in used_columns:
            continue
        if not norm_name:
            continue
        if any(ex in norm_name for ex in excludes if ex):
            continue

        score = 0
        if norm_name in exact_aliases:
            score = max(score, 100)

        for alias in contains_aliases:
            if alias and alias in norm_name:
                score = max(score, 70 + min(len(alias), 20))

        if score > best_score:
            best_score = score
            best_col = col_name

    return best_col if best_score > 0 else None


def resolve_standard_columns(dataframe):
    normalized_columns = {col: normalize_text(col) for col in dataframe.columns}
    used = set()

    def pick(exact=None, contains=None, excludes=None):
        col = _select_best_column(
            normalized_columns=normalized_columns,
            used_columns=used,
            exact_aliases=exact or [],
            contains_aliases=contains or [],
            excludes=excludes or [],
        )
        if col is not None:
            used.add(col)
        return col

    sno_col = pick(
        exact=["s no", "s no.", "sno", "sr no", "sl no", "serial no", "serial number"],
        contains=["serial no", "sr no", "sl no", "s no", "sno"],
    )
    state_col = pick(
        exact=["state", "state ut", "state/ut", "name of state", "state name"],
        contains=["name of state", "state"],
        excludes=["district", "facility"],
    )
    district_col = pick(
        exact=["district", "district name", "name of district"],
        contains=["district"],
        excludes=["state"],
    )
    block_col = pick(
        exact=["aspirational block name", "aspirational block", "block name"],
        contains=["aspirational block", "block name"],
    )
    name_col = pick(
        exact=["name", "participant name", "participants name", "name of the participant", "assessor name"],
        contains=["participant name", "participants name", "name of the participant", "assessor name", "name"],
        excludes=[
            "district",
            "facility",
            "state",
            "block",
            "sheet",
            "list",
            "training",
            "certificate",
        ],
    )
    designation_col = pick(
        exact=["designation", "cadre", "post"],
        contains=["designation", "cadre", "post"],
        excludes=["place of posting"],
    )
    posting_col = pick(
        exact=["place of posting", "facility name", "name of facility"],
        contains=["place of posting", "facility name", "name of facility", "posting"],
        excludes=["district", "state"],
    )
    mobile_col = pick(
        exact=["mobile no", "mobile number", "mobile", "phone number", "phone no", "contact number", "contact no"],
        contains=["mobile", "phone", "contact"],
    )
    email_col = pick(
        exact=["email", "email id", "e mail id", "e-mail", "mail id"],
        contains=["email", "mail id"],
    )
    score_col = pick(
        exact=["marks obtained", "score", "marks", "obtained marks", "score out of 40"],
        contains=["marks obtained", "obtained marks", "score out of 40", "score", "marks"],
        excludes=["percentage", "%", "total", "maximum", "max", "result", "certificate"],
    )
    total_col = pick(
        exact=["total marks", "maximum marks", "max marks", "marks out of"],
        contains=["total marks", "maximum marks", "max marks", "marks out of"],
    )
    pct_col = pick(
        exact=["percentage", "percentage %", "score %", "% score", "percent"],
        contains=["percentage", "score %", "% score", "percent"],
    )
    result_col = pick(
        exact=["result", "status", "remark", "remarks"],
        contains=["result", "status", "remark"],
    )
    cert_col = pick(
        exact=["certificate", "certificate no", "certificate number", "certificates", "cert no"],
        contains=["certificate", "cert no"],
    )
    year_col = pick(
        exact=["year", "assessment year", "batch year"],
        contains=["assessment year", "batch year", "year"],
        excludes=["certificate year"],
    )

    return {
        "Sno.": sno_col,
        "State": state_col,
        "district": district_col,
        "Aspirational Block Name": block_col,
        "Name": name_col,
        "Designation": designation_col,
        "Place of posting": posting_col,
        "Mobile no.": mobile_col,
        "email": email_col,
        "score": score_col,
        "Total Marks": total_col,
        "Percentage": pct_col,
        "result": result_col,
        "Certificate No.": cert_col,
        "Assessment Year": year_col,
    }


def read_sheet_with_detected_header(file_path, sheet_name):
    if file_path.suffix.lower() == ".csv":
        preview = pd.read_csv(file_path, header=None, nrows=60, dtype="string", on_bad_lines="skip")
        header_row = detect_header_row(preview)
        if header_row is None:
            return pd.DataFrame()
        return pd.read_csv(
            file_path,
            header=header_row,
            dtype="string",
            on_bad_lines="skip",
        )

    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=60, dtype="string")
    header_row = detect_header_row(preview)
    if header_row is None:
        return pd.DataFrame()
    return pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype="string")


def read_sheet_candidates(file_path, sheet_name):
    candidates = []
    if file_path.suffix.lower() == ".csv":
        preview = pd.read_csv(file_path, header=None, nrows=60, dtype="string", on_bad_lines="skip")
        for header_row in candidate_header_rows(preview):
            try:
                df = pd.read_csv(
                    file_path,
                    header=header_row,
                    dtype="string",
                    on_bad_lines="skip",
                )
                if not df.empty:
                    candidates.append(df)
            except Exception:
                continue
        return candidates

    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=60, dtype="string")
    for header_row in candidate_header_rows(preview):
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype="string")
            if not df.empty:
                candidates.append(df)
        except Exception:
            continue
    return candidates


def to_series(dataframe, column_name):
    if column_name and column_name in dataframe.columns:
        return clean_string_series(dataframe[column_name])
    return pd.Series([""] * len(dataframe), index=dataframe.index, dtype="string")


def standardize_dataframe(dataframe, file_path, sheet_name, file_state, file_year):
    if dataframe.empty:
        return pd.DataFrame()
    dataframe = dataframe.copy()
    dataframe = dataframe.dropna(axis=0, how="all").dropna(axis=1, how="all")
    if dataframe.empty:
        return pd.DataFrame()

    column_map = resolve_standard_columns(dataframe)
    sno_col = column_map["Sno."]
    state_col = column_map["State"]
    district_col = column_map["district"]
    block_col = column_map["Aspirational Block Name"]
    name_col = column_map["Name"]
    designation_col = column_map["Designation"]
    posting_col = column_map["Place of posting"]
    mobile_col = column_map["Mobile no."]
    email_col = column_map["email"]
    score_col = column_map["score"]
    total_col = column_map["Total Marks"]
    pct_col = column_map["Percentage"]
    result_col = column_map["result"]
    cert_col = column_map["Certificate No."]
    year_col = column_map["Assessment Year"]

    has_identity_columns = name_col is not None and any([designation_col, mobile_col, email_col, district_col])
    has_assessment_columns = any([score_col, pct_col, result_col, cert_col])
    if not (has_identity_columns and has_assessment_columns):
        return pd.DataFrame()

    output = pd.DataFrame(index=dataframe.index)
    output["Sno."] = to_series(dataframe, sno_col)
    output["State"] = to_series(dataframe, state_col)
    output["district"] = to_series(dataframe, district_col)
    output["Aspirational Block Name"] = to_series(dataframe, block_col)
    output["Name"] = to_series(dataframe, name_col)
    output["Designation"] = to_series(dataframe, designation_col)
    output["Place of posting"] = to_series(dataframe, posting_col)
    output["Mobile no."] = to_series(dataframe, mobile_col)
    output["email"] = to_series(dataframe, email_col).str.lower()
    output["result"] = to_series(dataframe, result_col)
    output["Certificate No."] = to_series(dataframe, cert_col)
    output["Assessment Year"] = to_series(dataframe, year_col)

    output["State"] = output["State"].replace("", pd.NA).fillna(file_state)
    output["State"] = output["State"].apply(normalize_state_value)
    output["district"] = clean_string_series(output["district"]).replace("", "Unknown")
    output["Designation"] = clean_string_series(output["Designation"]).replace("", "Unknown")
    output["Name"] = clean_string_series(output["Name"]).replace("", "Unknown")
    output["Place of posting"] = clean_string_series(output["Place of posting"]).replace("", "Unknown")
    output["Aspirational Block Name"] = clean_string_series(output["Aspirational Block Name"]).replace("", "No")

    score_values = parse_numeric(to_series(dataframe, score_col))
    total_values = parse_numeric(to_series(dataframe, total_col))
    pct_values = parse_percentage(to_series(dataframe, pct_col))

    valid_total = score_values.notna() & total_values.notna() & (total_values > 0) & (total_values != MAX_SCORE)
    score_values.loc[valid_total] = (score_values.loc[valid_total] / total_values.loc[valid_total]) * MAX_SCORE

    score_missing = score_values.isna() & pct_values.notna()
    score_values.loc[score_missing] = (pct_values.loc[score_missing] / 100.0) * MAX_SCORE
    score_over_max_with_pct = score_values.notna() & (score_values > MAX_SCORE) & pct_values.notna()
    score_values.loc[score_over_max_with_pct] = (pct_values.loc[score_over_max_with_pct] / 100.0) * MAX_SCORE

    percentage_from_score = (score_values / MAX_SCORE) * 100.0
    pct_values = pct_values.where(pct_values.notna(), percentage_from_score)
    pct_values = pct_values.round(2)
    score_values = score_values.round(2)

    cert_values, cert_parts = canonical_certificate(output["Certificate No."])
    output["Certificate No."] = cert_values.fillna("")
    output["Certificate State Code"] = cert_parts["state_code"].fillna("").str.upper()
    output["Batch No"] = cert_parts["batch"].fillna("")
    output["Participant Serial"] = cert_parts["serial"].fillna("")
    valid_certificate_mask = cert_parts["year"].notna() & cert_parts["state_code"].notna() & cert_parts["serial"].notna()

    cert_year = pd.to_numeric(cert_parts["year"], errors="coerce")
    year_values = pd.to_numeric(output["Assessment Year"], errors="coerce")
    year_values = year_values.where(year_values.between(2000, 2100), pd.NA)
    year_values = year_values.where(year_values.notna(), cert_year)
    year_values = year_values.where(year_values.notna(), float(file_year))
    output["Assessment Year"] = year_values.round().fillna(float(file_year)).map(
        lambda x: str(int(x)) if pd.notna(x) else str(int(file_year))
    )

    inferred_state_from_cert = output["Certificate State Code"].map(STATE_CODE_TO_STATE)
    unknown_state_mask = output["State"].eq("Unknown") & inferred_state_from_cert.notna()
    output.loc[unknown_state_mask, "State"] = inferred_state_from_cert[unknown_state_mask]

    score_pass_mask = score_values.notna() & (score_values >= PASS_MARKS)
    score_fail_mask = score_values.notna() & (score_values < PASS_MARKS)

    cert_pass_mask = score_values.isna() & valid_certificate_mask
    result_bool = parse_result_bool(output["result"])

    qualified = pd.Series(False, index=output.index)
    qualified.loc[score_pass_mask] = True
    qualified.loc[score_fail_mask] = False
    qualified.loc[cert_pass_mask] = True

    unresolved = score_values.isna() & output["Certificate No."].eq("") & result_bool.notna()
    qualified.loc[unresolved] = result_bool.loc[unresolved].astype(bool)

    generated_result = qualified.map({True: "Pass", False: "Fail"})
    output["result"] = output["result"].replace("", pd.NA).fillna(generated_result)
    output["Qualified or not qualified"] = qualified.map({True: "Qualified", False: "Not Qualified"})
    output["score"] = score_values
    output["Percentage"] = pct_values

    output["Mobile no."] = (
        output["Mobile no."].str.replace(r"\.0$", "", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
    )
    output["Mobile no."] = output["Mobile no."].str.replace(r"[^0-9+]", "", regex=True)

    district_norm = output["district"].str.lower().str.strip()
    designation_norm = output["Designation"].str.lower().str.strip()
    designation_like_district = {
        "medical officer",
        "mo",
        "staff nurse",
        "nursing officer",
        "consultant",
        "professor",
        "associate professor",
        "assistant professor",
        "smo",
        "rmo",
        "cmo",
    }
    district_is_contact = district_norm.str.contains(r"^[0-9+]{8,}$", regex=True, na=False) | district_norm.str.contains(
        "@", na=False
    )
    district_same_as_designation = district_norm.eq(designation_norm) & district_norm.ne("unknown")
    district_is_designation_word = district_norm.isin(designation_like_district)
    output.loc[district_is_contact | district_same_as_designation | district_is_designation_word, "district"] = "Unknown"

    for required_col in REQUIRED_OUTPUT_COLUMNS + EXTRA_OUTPUT_COLUMNS:
        if required_col not in output.columns:
            output[required_col] = ""

    output["Source File"] = str(file_path)
    output["Source Sheet"] = str(sheet_name)

    name_is_headerish = output["Name"].str.lower().isin(
        {
            "name",
            "participant name",
            "participants name",
            "name of the participant",
            "assessor name",
        }
    )
    has_contact_or_loc = (
        output["district"].ne("Unknown")
        | output["Designation"].ne("Unknown")
        | output["Mobile no."].ne("")
        | output["email"].ne("")
    )
    has_assessment_signal = score_values.notna() | pct_values.notna() | valid_certificate_mask | result_bool.notna()
    keep_mask = (~name_is_headerish) & output["Name"].ne("Unknown") & has_contact_or_loc & has_assessment_signal
    output = output[keep_mask].copy()

    if output.empty:
        return pd.DataFrame()

    return output[REQUIRED_OUTPUT_COLUMNS + EXTRA_OUTPUT_COLUMNS]


def parse_file(file_path):
    file_state = infer_state_from_text(file_path.as_posix())
    file_year = infer_year_from_path(file_path)

    parsed_frames = []
    errors = []

    try:
        if file_path.suffix.lower() == ".csv":
            best_standardized = pd.DataFrame()
            for dataframe in read_sheet_candidates(file_path, sheet_name=None):
                standardized = standardize_dataframe(
                    dataframe=dataframe,
                    file_path=file_path,
                    sheet_name=file_path.stem,
                    file_state=file_state,
                    file_year=file_year,
                )
                if len(standardized) > len(best_standardized):
                    best_standardized = standardized
            if not best_standardized.empty:
                parsed_frames.append(best_standardized)
            return parsed_frames, errors

        workbook = pd.ExcelFile(file_path)
        for sheet_name in workbook.sheet_names:
            try:
                sheet_state = infer_state_from_text(str(sheet_name))
                effective_state = file_state if file_state != "Unknown" else sheet_state
                if effective_state == "Unknown":
                    effective_state = file_state
                best_standardized = pd.DataFrame()
                candidate_frames = read_sheet_candidates(file_path, sheet_name=sheet_name)
                for dataframe in candidate_frames:
                    standardized = standardize_dataframe(
                        dataframe=dataframe,
                        file_path=file_path,
                        sheet_name=sheet_name,
                        file_state=effective_state,
                        file_year=file_year,
                    )
                    if len(standardized) > len(best_standardized):
                        best_standardized = standardized
                if not best_standardized.empty:
                    parsed_frames.append(best_standardized)
            except Exception as sheet_error:
                errors.append(f"{file_path.name} | {sheet_name}: {sheet_error}")
    except Exception as file_error:
        errors.append(f"{file_path.name}: {file_error}")

    return parsed_frames, errors


@st.cache_data(ttl=180, show_spinner=False)
def compile_assessment_data(root_path_str, signature):
    del signature
    root_path = Path(root_path_str)
    files = list_input_files(root_path)

    all_frames = []
    all_errors = []
    parsed_file_count = 0

    for file_path in files:
        file_frames, file_errors = parse_file(file_path)
        all_errors.extend(file_errors)
        if file_frames:
            parsed_file_count += 1
            all_frames.extend(file_frames)

    if not all_frames:
        empty_df = pd.DataFrame(columns=REQUIRED_OUTPUT_COLUMNS + EXTRA_OUTPUT_COLUMNS)
        return empty_df, all_errors, len(files), len(files), parsed_file_count

    combined = pd.concat(all_frames, ignore_index=True)
    combined = combined.reset_index(drop=True)
    combined["Sno."] = range(1, len(combined) + 1)
    return combined, all_errors, len(files), len(files), parsed_file_count


def dataframe_to_excel_bytes(dataframe):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Results")
    buffer.seek(0)
    return buffer.getvalue()


def state_summary_table(dataframe):
    state_grouped = (
        dataframe.assign(_qualified=dataframe["Qualified or not qualified"].eq("Qualified"))
        .groupby("State", as_index=False)["_qualified"]
        .sum()
        .rename(columns={"_qualified": "Qualified Assessors"})
        .sort_values("Qualified Assessors", ascending=False)
    )
    return state_grouped


@st.cache_data(show_spinner=False)
def load_india_geojson(data_root_str):
    data_root = Path(data_root_str)
    local_candidates = [
        data_root / "india_goi_official_map.geojson",
        Path(__file__).with_name("india_goi_official_map.geojson"),
    ]
    for local_path in local_candidates:
        if local_path.exists():
            with open(local_path, "r", encoding="utf-8") as local_file:
                return json.load(local_file), str(local_path), True

    for source_url in GEOJSON_URLS:
        try:
            with urlopen(source_url, timeout=20) as response:
                return json.load(response), source_url, False
        except Exception:
            continue
    return None, "", False


def detect_geojson_state_property(geojson):
    if not geojson or "features" not in geojson or not geojson["features"]:
        return None
    sample_properties = geojson["features"][0].get("properties", {})
    candidates = ["ST_NM", "state_name", "STATE", "NAME_1", "State", "st_nm", "NAME", "name"]
    for key in candidates:
        if key in sample_properties:
            return key
    return next(iter(sample_properties.keys()), None)


def build_map_dataframe(dataframe, geojson, state_property):
    qualified_counts = (
        dataframe.assign(
            _qualified=dataframe["Qualified or not qualified"].eq("Qualified"),
            _state_norm=dataframe["State"].apply(normalize_state_value),
        )
        .groupby("_state_norm", as_index=False)["_qualified"]
        .sum()
        .rename(columns={"_state_norm": "State", "_qualified": "Qualified Assessors"})
    )
    state_to_count = dict(zip(qualified_counts["State"], qualified_counts["Qualified Assessors"]))

    rows = []
    for feature in geojson.get("features", []):
        raw_state = str(feature.get("properties", {}).get(state_property, "")).strip()
        canonical_state = normalize_state_value(raw_state)
        rows.append(
            {
                "Geo State": raw_state,
                "State": canonical_state,
                "Qualified Assessors": int(state_to_count.get(canonical_state, 0)),
            }
        )
    return pd.DataFrame(rows)


def rerun_app():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def resolve_data_root():
    candidates = []
    env_path = os.getenv("NQAS_DATA_ROOT", "").strip()
    if env_path:
        candidates.append(("Environment Variable", Path(env_path)))

    try:
        secret_path = str(st.secrets.get("NQAS_DATA_ROOT", "")).strip()
    except Exception:
        secret_path = ""
    if secret_path:
        candidates.append(("Streamlit Secret", Path(secret_path)))

    candidates.append(("Local Default", DEFAULT_LOCAL_DATA_ROOT))

    for label, path in candidates:
        if path.exists():
            return path, f"{label}: {path}", True

    if candidates:
        label, path = candidates[0]
        return path, f"{label}: {path}", False
    return DEFAULT_LOCAL_DATA_ROOT, f"Local Default: {DEFAULT_LOCAL_DATA_ROOT}", False


def resolve_remote_zip_url():
    env_zip_url = os.getenv("NQAS_DATA_ZIP_URL", "").strip()
    if env_zip_url:
        return env_zip_url, "Environment Variable"
    try:
        secret_zip_url = str(st.secrets.get("NQAS_DATA_ZIP_URL", "")).strip()
    except Exception:
        secret_zip_url = ""
    if secret_zip_url:
        return secret_zip_url, "Streamlit Secret"
    return "", ""


def write_uploaded_files_to_temp(uploaded_files, uploaded_zip):
    temp_dir = Path(tempfile.mkdtemp(prefix="nqas_streamlit_data_"))

    if uploaded_zip is not None:
        zip_target = temp_dir / uploaded_zip.name
        zip_target.write_bytes(uploaded_zip.getbuffer())
        with zipfile.ZipFile(zip_target, "r") as zip_ref:
            zip_ref.extractall(temp_dir)

    for uploaded in uploaded_files:
        target_path = temp_dir / uploaded.name
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(uploaded.getbuffer())

    return temp_dir


def write_remote_zip_to_temp(zip_url):
    temp_dir = Path(tempfile.mkdtemp(prefix="nqas_streamlit_data_"))
    zip_target = temp_dir / "remote_data.zip"
    with urlopen(zip_url, timeout=120) as response:
        zip_target.write_bytes(response.read())
    with zipfile.ZipFile(zip_target, "r") as zip_ref:
        zip_ref.extractall(temp_dir)
    return temp_dir


st.title("Internal Assessment Dashboard")

active_data_root, data_source_label, data_root_exists = resolve_data_root()
st.caption(f"Data source: `{data_source_label}`")

if st_autorefresh is not None:
    st_autorefresh(interval=120000, key="ia_data_refresh")
    st.sidebar.caption("Auto-refresh: every 2 minutes")
else:
    st.sidebar.caption("Install `streamlit-autorefresh` for timed auto-refresh.")

if st.sidebar.button("Refresh Data Now"):
    st.cache_data.clear()
    rerun_app()

if not data_root_exists:
    st.warning(
        "Configured folder is not available in this environment. "
        "On Streamlit Cloud, local Windows paths are not accessible."
    )
    remote_zip_url, remote_zip_source = resolve_remote_zip_url()
    if remote_zip_url:
        try:
            active_data_root = write_remote_zip_to_temp(remote_zip_url)
            st.caption(
                f"Using remote ZIP from {remote_zip_source}. "
                f"Temporary extraction path: `{active_data_root}`"
            )
        except Exception as remote_error:
            st.error(f"Failed to load remote ZIP: {remote_error}")
            st.stop()
    else:
        st.markdown(
            "Upload a `.zip` of your data folder, or upload multiple result files directly."
        )
        uploaded_zip = st.file_uploader("Upload Data ZIP", type=["zip"])
        uploaded_files = st.file_uploader(
            "Or upload result files",
            type=["xlsx", "xls", "xlsm", "csv"],
            accept_multiple_files=True,
        )
        if uploaded_zip is None and not uploaded_files:
            st.info("Waiting for files.")
            st.stop()
        active_data_root = write_uploaded_files_to_temp(uploaded_files or [], uploaded_zip)
        st.caption(f"Using uploaded files from temporary path: `{active_data_root}`")

signature = build_data_signature(active_data_root)
compiled_df, parse_errors, total_files, files_read, parsed_files = compile_assessment_data(
    str(active_data_root), signature
)

if compiled_df.empty:
    st.warning("No valid assessment records found yet. Add/update Excel files and refresh.")
    if parse_errors:
        with st.expander("Parser Notes"):
            st.write("\n".join(parse_errors[:100]))
    st.stop()

for col in REQUIRED_OUTPUT_COLUMNS + EXTRA_OUTPUT_COLUMNS:
    if col not in compiled_df.columns:
        compiled_df[col] = ""

years_available = sorted(
    [int(year) for year in pd.to_numeric(compiled_df["Assessment Year"], errors="coerce").dropna().unique()]
)
states_available = sorted(
    [
        state
        for state in compiled_df["State"].dropna().unique()
        if str(state).strip() and normalize_state_value(state) != "Unknown"
    ]
)

st.sidebar.header("Navigation")
selected_page = st.sidebar.selectbox("Page", ["India Overview"] + states_available)

working_df = compiled_df.copy()
if selected_page != "India Overview":
    working_df = working_df[working_df["State"] == selected_page].copy()
else:
    include_unknown = st.sidebar.checkbox("Include Unknown State records", value=False)
    if not include_unknown:
        working_df = working_df[working_df["State"].apply(normalize_state_value) != "Unknown"].copy()

st.sidebar.header("Filters")
selected_years = st.sidebar.multiselect(
    "Assessment Year",
    options=years_available,
    default=years_available,
)
if selected_years:
    working_df = working_df[
        pd.to_numeric(working_df["Assessment Year"], errors="coerce").isin(selected_years)
    ].copy()

district_options = sorted([d for d in working_df["district"].dropna().unique() if str(d).strip()])
district_options = [d for d in district_options if str(d).strip().lower() not in {"unknown", "no", "nan"}]
selected_districts = st.sidebar.multiselect("District", options=district_options)
if selected_districts:
    working_df = working_df[working_df["district"].isin(selected_districts)].copy()

designation_options = sorted(
    [d for d in working_df["Designation"].dropna().unique() if str(d).strip()]
)
designation_options = [d for d in designation_options if str(d).strip().lower() not in {"unknown", "nan"}]
selected_designations = st.sidebar.multiselect("Designation", options=designation_options)
if selected_designations:
    working_df = working_df[working_df["Designation"].isin(selected_designations)].copy()

qualification_filter = st.sidebar.selectbox(
    "Qualified Status",
    ["All", "Qualified", "Not Qualified"],
)
if qualification_filter != "All":
    working_df = working_df[
        working_df["Qualified or not qualified"] == qualification_filter
    ].copy()

st.sidebar.subheader("Certificate Filter")
cert_only = st.sidebar.checkbox("Only rows with certificate number")
if cert_only:
    working_df = working_df[working_df["Certificate No."].astype("string").str.strip().ne("")].copy()

cert_query = st.sidebar.text_input(
    "Certificate Number Search",
    placeholder="IA/2024/UA6/01 or partial text",
)
if cert_query.strip():
    cert_series = clean_string_series(working_df["Certificate No."])
    raw_query = cert_query.strip()
    canonical_query = canonicalize_certificate_text(raw_query)
    cert_mask_raw = cert_series.str.contains(
        raw_query,
        case=False,
        na=False,
    )
    cert_mask_canonical = cert_series.map(canonicalize_certificate_text).str.contains(
        canonical_query,
        case=False,
        na=False,
        regex=False,
    )
    working_df = working_df[cert_mask_raw | cert_mask_canonical].copy()

certificate_year_options = sorted(
    [
        int(year)
        for year in pd.to_numeric(
            working_df["Certificate No."].astype("string").str.extract(CERTIFICATE_REGEX)["year"],
            errors="coerce",
        )
        .dropna()
        .unique()
    ]
)
selected_cert_years = st.sidebar.multiselect("Certificate Year", options=certificate_year_options)
if selected_cert_years:
    parsed_year = pd.to_numeric(
        working_df["Certificate No."].astype("string").str.extract(CERTIFICATE_REGEX)["year"],
        errors="coerce",
    )
    working_df = working_df[parsed_year.isin(selected_cert_years)].copy()

batch_options = sorted(
    [
        batch
        for batch in clean_string_series(working_df["Batch No"])
        .str.replace(r"\.0$", "", regex=True)
        .dropna()
        .unique()
        if str(batch).strip()
    ]
)
selected_batches = st.sidebar.multiselect("Batch No", options=batch_options)
if selected_batches:
    batch_series = clean_string_series(working_df["Batch No"]).str.replace(r"\.0$", "", regex=True)
    working_df = working_df[batch_series.isin(selected_batches)].copy()

search_query = st.text_input(
    "Search records",
    placeholder=(
        "Search by Sno., State, district, Aspirational Block, Name, Designation, "
        "Place of posting, Mobile no., email, score, Percentage, result, qualification"
    ),
)
if search_query.strip():
    searchable_columns = REQUIRED_OUTPUT_COLUMNS + ["Certificate No.", "Assessment Year"]
    haystack = working_df[searchable_columns].astype("string").fillna("").agg(" | ".join, axis=1)
    search_mask = haystack.str.contains(search_query.strip(), case=False, na=False)
    working_df = working_df[search_mask].copy()

working_df = working_df.reset_index(drop=True)
working_df["Sno."] = range(1, len(working_df) + 1)

total_participants = len(working_df)
qualified_count = int(working_df["Qualified or not qualified"].eq("Qualified").sum())
qualification_rate = (qualified_count / total_participants * 100.0) if total_participants else 0.0

m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Total Participants", f"{total_participants}")
with m2:
    st.metric("Qualified", f"{qualified_count}")
with m3:
    st.metric("Qualification %", f"{qualification_rate:.2f}")
with m4:
    st.metric("Files Read", f"{files_read} / {total_files}")

st.caption(f"Files with extractable participant records: {parsed_files} / {total_files}")

download_compiled_df = working_df[REQUIRED_OUTPUT_COLUMNS].copy()
download_detailed_df = working_df[REQUIRED_OUTPUT_COLUMNS + ["Certificate No.", "Assessment Year"]].copy()

d1, d2, d3 = st.columns(3)
with d1:
    st.download_button(
        "Download CSV (Required Headers)",
        data=download_compiled_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="compiled_internal_assessment.csv",
        mime="text/csv",
    )
with d2:
    st.download_button(
        "Download Filtered CSV",
        data=download_detailed_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="filtered_internal_assessment.csv",
        mime="text/csv",
    )
with d3:
    st.download_button(
        "Download Filtered Excel",
        data=dataframe_to_excel_bytes(download_detailed_df),
        file_name="filtered_internal_assessment.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if selected_page == "India Overview":
    geojson, geojson_source, is_local_geojson = load_india_geojson(str(active_data_root))
    if geojson is None:
        st.warning("Could not load India state geojson for map rendering.")
    else:
        state_property = detect_geojson_state_property(geojson)
        if state_property is None:
            st.warning("Geojson loaded, but no state property found for mapping.")
        else:
            map_df = build_map_dataframe(working_df, geojson, state_property)
            map_fig = px.choropleth(
                map_df,
                geojson=geojson,
                locations="Geo State",
                featureidkey=f"properties.{state_property}",
                color="Qualified Assessors",
                hover_data={"State": True, "Qualified Assessors": True},
                color_continuous_scale="YlGn",
            )
            map_fig.update_geos(fitbounds="locations", visible=False)
            map_fig.update_layout(margin={"r": 0, "t": 0, "l": 0, "b": 0}, height=650)
            st.subheader("India Heat Map: Qualified Assessors State-wise")
            st.plotly_chart(map_fig, use_container_width=True)
            if is_local_geojson:
                st.caption(f"Map source: local GOI file `{geojson_source}`")
            else:
                st.caption(
                    "Map source: fallback public boundary file. "
                    "For strict GOI official map, place `india_goi_official_map.geojson` in the data folder."
                )

    st.subheader("State-wise Qualified Count")
    st.dataframe(state_summary_table(working_df), use_container_width=True)
else:
    st.subheader(f"State Page: {selected_page}")

if working_df.empty:
    st.info("No rows match current filters.")
    st.stop()

chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    designation_chart_df = (
        working_df[working_df["Designation"].astype("string").str.strip().str.lower().ne("unknown")]
        .groupby("Designation", as_index=False)
        .size()
        .rename(columns={"size": "Participants"})
        .sort_values("Participants", ascending=False)
    )
    if not designation_chart_df.empty:
        fig_bar = px.bar(
            designation_chart_df,
            x="Designation",
            y="Participants",
            title="Participants by Designation",
        )
        fig_bar.update_layout(xaxis_title="Designation", yaxis_title="Participants")
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("Designation data not available for bar chart.")

with chart_col2:
    pie_df = (
        working_df.groupby("Qualified or not qualified", as_index=False)
        .size()
        .rename(columns={"size": "Participants"})
    )
    if not pie_df.empty:
        fig_pie = px.pie(
            pie_df,
            names="Qualified or not qualified",
            values="Participants",
            title="Qualified vs Not Qualified",
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("Qualification data not available for pie chart.")

st.subheader("Year-wise Segregation")
year_values = (
    pd.to_numeric(working_df["Assessment Year"], errors="coerce")
    .dropna()
    .astype(int)
    .sort_values()
    .unique()
    .tolist()
)
if year_values:
    year_tabs = [str(year) for year in year_values]
    tabs = st.tabs(year_tabs)
    for tab, year in zip(tabs, year_values):
        with tab:
            year_slice = working_df[
                pd.to_numeric(working_df["Assessment Year"], errors="coerce") == int(year)
            ].copy()
            st.write(f"Rows: {len(year_slice)}")
            st.dataframe(
                year_slice[REQUIRED_OUTPUT_COLUMNS + ["Certificate No.", "Assessment Year"]],
                use_container_width=True,
            )
            st.download_button(
                f"Download {year} CSV",
                data=year_slice[REQUIRED_OUTPUT_COLUMNS + ["Certificate No.", "Assessment Year"]]
                .to_csv(index=False)
                .encode("utf-8-sig"),
                file_name=f"internal_assessment_{year}.csv",
                mime="text/csv",
                key=f"download_csv_{year}",
            )
else:
    st.info("No assessment year values available after filtering.")

st.subheader("Filtered Records")
st.dataframe(
    working_df[REQUIRED_OUTPUT_COLUMNS + ["Certificate No.", "Assessment Year", "Batch No", "Participant Serial"]],
    use_container_width=True,
)

if parse_errors:
    with st.expander("Parser Notes (non-blocking)"):
        st.write(
            "Some files/sheets were skipped because they did not look like participant result sheets "
            "or had unreadable structure."
        )
        st.write(f"Skipped sheet/file parse events: {len(parse_errors)}")
        st.write("\n".join(parse_errors[:40]))
