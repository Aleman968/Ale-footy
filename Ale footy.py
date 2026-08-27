
import streamlit as st
import pandas as pd
import datetime as dt
import requests
import math
from zoneinfo import ZoneInfo

st.set_page_config(page_title="MG Auto Dati (FootyStats)", layout="wide")

# =========================
# GOOGLE SHEETS (SAFE)
# =========================
import json as _json

def _gs_available():
    try:
        _ = st.secrets["gsheets"]
        return True
    except Exception:
        return False

def _gs_append_row(row_values):
    """Append a row to the configured Google Sheet."""
    import gspread
    from google.oauth2.service_account import Credentials

    cfg = st.secrets["gsheets"]
    creds_info = _json.loads(cfg["creds_json"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(cfg["spreadsheet_id"])
    ws = sh.worksheet(cfg["worksheet_name"])
    ws.append_row(row_values, value_input_option="USER_ENTERED")

st.markdown(
    """
    <style>
    :root{
        --bg:#ffffff;
        --panel:#ffffff;
        --panel2:#ffffff;
        --line:#dfe6ee;
        --text:#111827;
        --muted:#64748b;
        --green:#24d366;
        --orange:#f59e0b;
        --red:#ff5b57;
        --blue:#2f86ff;
    }
    .stApp{
        background:#ffffff;
        color:var(--text);
    }
    .block-container{padding-top:.7rem;padding-bottom:2rem;max-width:1480px;}
    h1,h2,h3,h4{color:#111827!important;}
    .render-topline{
        display:flex;justify-content:space-between;align-items:center;
        padding:4px 2px 12px;border-bottom:1px solid #e2e8f0;margin-bottom:12px;
        color:#111827;font-weight:800;
    }
    .render-topline .league{font-size:1.05rem;letter-spacing:.02em}
    .render-topline .day{color:#30df88;margin-left:15px}
    .render-topline .date{font-size:.92rem;color:#64748b;font-weight:600}
    .render-card{
        background:#ffffff;
        border:1px solid #dfe6ee;border-radius:14px;padding:14px 16px;
        box-shadow:0 4px 16px rgba(15,23,42,.06);height:100%;
    }
    .render-title{
        text-align:center;text-transform:uppercase;letter-spacing:.04em;
        font-size:.78rem;font-weight:900;color:#334155;margin-bottom:10px;
    }
    .match-center{
        text-align:center;font-size:1.5rem;font-weight:950;color:#111827;
        margin:1px 0 12px;
    }
    .match-vs{color:#9baabd;font-size:1rem;margin:0 18px;font-weight:800}
    .pick-box{
        border:1px solid #168a58;background:#f8fffb;
        border-radius:12px;padding:14px 16px;text-align:center;
    }
    .pick-kicker{font-size:.79rem;color:#3bea85;font-weight:950;text-transform:uppercase}
    .pick-name{font-size:1.8rem;font-weight:950;color:#111827;margin:5px 0 10px}
    .pick-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:0}
    .pick-stat{padding:4px 8px;border-right:1px solid #e2e8f0}
    .pick-stat:last-child{border-right:0}
    .pick-stat-label{font-size:.67rem;color:#9cadbd;text-transform:uppercase}
    .pick-stat-val{font-size:1.36rem;font-weight:950;margin-top:3px}
    .green{color:#39e27d}.orange{color:#ffad27}.white{color:#111827}
    .alt-label{text-align:center;color:#98a8bb;text-transform:uppercase;font-size:.67rem;font-weight:900;margin:11px 0 6px}
    .alt-wrap{display:flex;gap:8px;flex-wrap:wrap;justify-content:center}
    .alt-pill{border:1px solid #2b4668;background:#0b1b2d;border-radius:7px;padding:7px 12px;color:#e8f0fb;font-size:.78rem;font-weight:850}
    .standings-wrap{max-height:300px;overflow:auto;padding-right:2px}
    table.render-table{width:100%;border-collapse:collapse;font-size:.75rem}
    .render-table th{font-size:.62rem;text-transform:uppercase;color:#64748b;padding:6px 5px;border-bottom:1px solid #e2e8f0;text-align:right}
    .render-table th:nth-child(1),.render-table th:nth-child(2){text-align:left}
    .render-table td{padding:7px 5px;border-bottom:1px solid #eef2f7;text-align:right;color:#1e293b}
    .render-table td:nth-child(1),.render-table td:nth-child(2){text-align:left}
    .render-table tr.sel-home td{color:#63e68d;font-weight:900}
    .render-table tr.sel-away td{color:#ffb12f;font-weight:900}
    .last5-team{font-weight:900;color:#111827;font-size:.82rem;margin:7px 0 4px}
    .last5-row{display:grid;grid-template-columns:48px 42px 1fr 46px;gap:6px;padding:5px 0;border-bottom:1px solid #eef2f7;font-size:.72rem;align-items:center}
    .last5-row .muted{color:#64748b}.last5-row .score{text-align:right;font-weight:900;color:white}
    .team-panel{
        background:#ffffff;
        border:1px solid #28528a;border-radius:14px;overflow:hidden;height:100%;
    }
    .team-panel.away{border-color:#8c5a12}
    .team-head{padding:11px 15px;background:#f2f7ff;font-weight:950;color:white}
    .team-panel.away .team-head{background:#fff8ed;color:#ffb020}
    .metric-strip{display:grid;grid-template-columns:1.15fr .75fr .75fr;align-items:center;gap:8px;margin:9px 10px;padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:9px}
    .metric-strip .lbl{font-size:.74rem;font-weight:900;color:#334155}
    .metric-strip .big{text-align:center;font-size:1.05rem;font-weight:950;color:#111827}
    .metric-strip .sub{text-align:center;font-size:.62rem;color:#64748b;text-transform:uppercase}
    .trend-strip{margin:9px 10px;padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:9px;display:flex;justify-content:space-between;align-items:center}
    .trend-strip .lbl{font-size:.74rem;font-weight:900;color:#334155}
    .trend-strip .val{font-size:.82rem;font-weight:950}
    .dist-box{margin:9px 10px;padding:10px 12px;background:#ffffff;border:1px solid #e2e8f0;border-radius:9px}
    .dist-title{font-size:.67rem;color:#334155;font-weight:900;text-transform:uppercase;margin-bottom:7px}
    .dist-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:5px;align-items:end}
    .dist-col{text-align:center;font-size:.62rem;color:#64748b}
    .dist-bar-wrap{height:43px;display:flex;align-items:end;justify-content:center;margin:3px 0}
    .dist-bar{width:72%;min-height:3px;border-radius:3px 3px 1px 1px}
    .dist-bar.gf{background:#37cf67}.dist-bar.ga{background:#ff5c59}
    .dist-pct{font-size:.7rem;font-weight:950}
    .footer-note{font-size:.72rem;color:#64748b;text-align:center;margin:10px 0}
    div[data-testid="stExpander"]{border:1px solid #dfe6ee!important;border-radius:13px!important;background:#ffffff!important}
    div[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden}

    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextInput"] input{
        background:#ffffff!important;
        color:#111827!important;
    }
    div[data-baseweb="popover"],
    div[data-baseweb="menu"]{
        background:#ffffff!important;
        color:#111827!important;
    }
    .stMarkdown, .stCaption, label, p, span{
        color:inherit;
    }


    /* V12 readability fixes on white background */
    .team-panel,
    .team-panel .team-head,
    .team-panel .metric-strip,
    .team-panel .metric-strip .lbl,
    .team-panel .metric-strip .big,
    .team-panel .metric-strip .sub,
    .team-panel .trend-strip,
    .team-panel .trend-strip .lbl,
    .team-panel .dist-box,
    .team-panel .dist-title,
    .team-panel .dist-col,
    .team-panel .dist-pct {
        color:#111827 !important;
    }

    .team-panel .team-head {
        color:#0f172a !important;
    }

    .last5-team,
    .last5-row,
    .last5-row span,
    .last5-row .score {
        color:#111827 !important;
    }

    .last5-row .muted {
        color:#64748b !important;
    }


    /* =========================
       RIEPILOGO GIOCATE - V14
       ========================= */
    .summary-hero{
        background:linear-gradient(135deg,#ffffff 0%,#f5f9ff 100%);
        border:1px solid #dfe6ee;
        border-radius:16px;
        padding:18px 20px;
        margin:8px 0 14px;
        box-shadow:0 4px 16px rgba(15,23,42,.05);
    }
    .summary-eyebrow{
        color:#2563eb;
        font-size:.68rem;
        font-weight:900;
        letter-spacing:.10em;
        margin-bottom:3px;
    }
    .summary-title{
        color:#0f172a;
        font-size:1.65rem;
        line-height:1.1;
        font-weight:950;
    }
    .summary-subtitle{
        color:#64748b;
        font-size:.82rem;
        margin-top:5px;
    }
    .summary-filter-title{
        color:#475569;
        font-size:.70rem;
        font-weight:900;
        letter-spacing:.08em;
        margin-bottom:2px;
    }
    .summary-countbar{
        display:flex;
        justify-content:space-between;
        align-items:center;
        gap:12px;
        color:#475569;
        font-size:.82rem;
        padding:12px 3px 7px;
    }
    .summary-countbar b{color:#0f172a}
    .summary-day{
        color:#0f172a;
        font-size:1.00rem;
        font-weight:950;
        margin:16px 0 8px;
        padding-bottom:6px;
        border-bottom:2px solid #e8eef5;
    }
    .summary-match{padding:2px 0}
    .summary-time{
        display:inline-block;
        color:#2563eb;
        font-size:.78rem;
        font-weight:950;
        background:#eff6ff;
        border-radius:999px;
        padding:3px 8px;
        margin-bottom:6px;
    }
    .summary-teams{
        color:#0f172a;
        font-size:1.06rem;
        font-weight:950;
        line-height:1.2;
    }
    .summary-teams span{
        color:#94a3b8!important;
        font-size:.72rem;
        font-weight:800;
        margin:0 5px;
        text-transform:uppercase;
    }
    .summary-league{
        color:#64748b;
        font-size:.72rem;
        margin-top:3px;
    }
    .summary-market{
        display:inline-block;
        color:#0f5132;
        font-size:.79rem;
        font-weight:950;
        margin-top:7px;
        margin-right:7px;
    }
    .summary-badge{
        display:inline-block;
        border-radius:999px;
        padding:3px 8px;
        font-size:.61rem;
        font-weight:950;
        letter-spacing:.03em;
    }
    .summary-badge.robusto{background:#e8f8ee;color:#16854a}
    .summary-badge.neutro{background:#fff5d8;color:#946200}
    .summary-badge.instabile{background:#feecec;color:#ba2d2d}
    .summary-prob{
        text-align:center;
        border-left:1px solid #edf1f5;
        border-right:1px solid #edf1f5;
        padding:7px 8px;
    }
    .summary-prob-label{
        color:#94a3b8;
        font-size:.61rem;
        font-weight:900;
        letter-spacing:.07em;
    }
    .summary-prob-value{
        color:#18a957;
        font-size:1.65rem;
        line-height:1.05;
        font-weight:950;
        margin-top:3px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)

API_BASE = "https://api.football-data-api.com"
APP_TIMEZONE = "Europe/Rome"

def load_api_key() -> str:
    try:
        key = str(st.secrets["FOOTYSTATS_API_KEY"]).strip()
        if key:
            return key
    except Exception:
        pass
    try:
        key = str(st.secrets.get("FOOTY_API_KEY", "")).strip()
        if key:
            return key
    except Exception:
        pass
    st.error("Manca la chiave API di FootyStats. Inserisci FOOTYSTATS_API_KEY in .streamlit/secrets.toml")
    st.stop()

TOKEN = load_api_key()

def _secret_bool(name: str, default: bool = False) -> bool:
    try:
        raw = str(st.secrets.get(name, str(default))).strip().lower()
    except Exception:
        raw = str(default).lower()
    return raw in {"1", "true", "yes", "y", "on"}

CHOSEN_LEAGUES_ONLY = _secret_bool("FOOTY_CHOSEN_LEAGUES_ONLY", False)
AUTO_RETRY_ALL_LEAGUES = _secret_bool("FOOTY_AUTO_RETRY_ALL_LEAGUES", True)

def safe_float(value, default=0.0):
    try:
        if value in (None, "", "null"):
            return default
        return float(value)
    except Exception:
        return default

def safe_int(value, default=0):
    try:
        if value in (None, "", "null"):
            return default
        return int(float(value))
    except Exception:
        return default

def normalize_text(text):
    return " ".join(str(text or "").strip().lower().split())

def clean_league_name(name):
    name = str(name or "").strip()
    if not name:
        return "Campionato"
    import re
    return re.sub(r"\s*\((?:19|20)\d{2}(?:/?(?:19|20)\d{2})?\)\s*$", "", name).strip()

def build_pretty_league_name(country, league_name):
    country = str(country or "").strip()
    league_name = str(league_name or "").strip()
    if country and league_name:
        if country.lower() in league_name.lower():
            return clean_league_name(league_name)
        return clean_league_name(f"{country} - {league_name}")
    return clean_league_name(league_name or country or "Campionato")

def deep_find_first(obj, candidate_keys):
    if isinstance(obj, dict):
        for k in candidate_keys:
            if k in obj and obj[k] not in (None, "", [], {}):
                return obj[k]
        for v in obj.values():
            found = deep_find_first(v, candidate_keys)
            if found not in (None, "", [], {}):
                return found
    elif isinstance(obj, (list, tuple)):
        for item in obj:
            found = deep_find_first(item, candidate_keys)
            if found not in (None, "", [], {}):
                return found
    return None

def deep_collect_ids(obj):
    found = set()
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in {"id", "season_id", "competition_id", "league_id"}:
                val = safe_int(v, 0)
                if val:
                    found.add(val)
            found.update(deep_collect_ids(v))
    elif isinstance(obj, (list, tuple)):
        for item in obj:
            found.update(deep_collect_ids(item))
    return found

def coerce_league_item(raw_item):
    if isinstance(raw_item, dict):
        return raw_item
    if isinstance(raw_item, (list, tuple)):
        merged = {}
        for part in raw_item:
            if isinstance(part, dict):
                for k, v in part.items():
                    if k not in merged or merged.get(k) in (None, "", [], {}):
                        merged[k] = v
        return merged
    return {}

@st.cache_data(show_spinner=False, ttl=60*30)
def api_get(path: str, params: dict | None = None) -> dict | None:
    url = f"{API_BASE}{path}"
    merged = {"key": TOKEN}
    if params:
        merged.update(params)
    try:
        r = requests.get(url, params=merged, timeout=45)
    except Exception as e:
        st.error(f"Errore di rete durante la chiamata API: {e}")
        return None
    if not r.ok:
        preview = r.text[:400] if r.text else "Nessun dettaglio disponibile"
        st.error(f"Errore API {r.status_code} su {path}: {preview}")
        return None
    try:
        return r.json()
    except Exception as e:
        st.error(f"Risposta API non valida (JSON): {e}")
        return None

@st.cache_data(show_spinner=False, ttl=60*15)
def fetch_match_details(match_id: int) -> dict:
    """Dettagli singola partita FootyStats, incluse le quote disponibili."""
    if not match_id:
        return {}
    data = api_get("/match", {"match_id": int(match_id)})
    if not isinstance(data, dict):
        return {}
    raw = data.get("data", data)
    if isinstance(raw, list):
        return raw[0] if raw and isinstance(raw[0], dict) else {}
    return raw if isinstance(raw, dict) else {}


def _valid_odd(value):
    try:
        v = float(value)
        return v if v > 1.0 else None
    except Exception:
        return None


def odds_for_candidate(candidate: dict, match_details: dict):
    """Restituisce la quota bookmaker FootyStats se esiste per il mercato proposto."""
    if not candidate or not isinstance(match_details, dict):
        return None
    name = str(candidate.get("name", "")).upper().replace("–", "-")
    kind = str(candidate.get("kind", "")).upper()

    # Mercati match documentati da FootyStats.
    if kind == "UNDER":
        if "UNDER 2.5" in name:
            return _valid_odd(match_details.get("odds_ft_under25"))
        if "UNDER 3.5" in name:
            return _valid_odd(match_details.get("odds_ft_under35"))
    if kind == "OVER":
        if "OVER 2.5" in name:
            return _valid_odd(match_details.get("odds_ft_over25"))
        # "Over 1.5 squadra" e' team-specifico: FootyStats non espone
        # una quota equivalente nel set standard usato qui.
        if "OVER 1.5" in name and "SQUADRA" not in name:
            return _valid_odd(match_details.get("odds_ft_over15"))
    return None


@st.cache_data(show_spinner=False, ttl=60*60)
def fetch_league_lookup(chosen_only: bool = True):
    data = api_get("/league-list", {"chosen_leagues_only": "true"} if chosen_only else {})
    if data is None:
        return {"leagues": [], "id_to_key": {}, "id_to_name": {}, "name_map": {}}
    raw_data = data.get("data", [])
    if not isinstance(raw_data, list):
        raw_data = []

    leagues = []
    id_to_key = {}
    id_to_name = {}
    name_map = {}

    for idx, raw_item in enumerate(raw_data):
        item = coerce_league_item(raw_item)
        if not item:
            continue

        season_blob = deep_find_first(item, ["season"]) or {}
        if not isinstance(season_blob, dict):
            season_blob = {}

        season_id = safe_int(deep_find_first(item, ["season_id"]) or season_blob.get("id"), 0)
        item_id = safe_int(deep_find_first(item, ["id"]), 0)
        extra_ids = sorted(deep_collect_ids(raw_item) | deep_collect_ids(item))

        country = str(deep_find_first(item, ["country"]) or "").strip()
        league_name = str(
            deep_find_first(item, ["league_name"])
            or deep_find_first(item, ["english_name"])
            or deep_find_first(item, ["name_it"])
            or deep_find_first(item, ["name"])
            or ""
        ).strip()

        clean_name = build_pretty_league_name(country, league_name)
        canonical_key = f"league_{season_id or item_id or idx}"
        row = {
            "key": canonical_key,
            "season_id": season_id,
            "item_id": item_id,
            "extra_ids": extra_ids,
            "name": clean_name,
            "country": country,
            "league_name": league_name,
        }
        leagues.append(row)

        ids_for_row = set(extra_ids)
        if season_id:
            ids_for_row.add(season_id)
        if item_id:
            ids_for_row.add(item_id)

        for found_id in ids_for_row:
            id_to_key[found_id] = canonical_key
            id_to_name[found_id] = row["name"]

        for label in {row["name"], row["league_name"]}:
            label_norm = normalize_text(label)
            if label_norm:
                name_map.setdefault(label_norm, set()).add(canonical_key)

    leagues.sort(key=lambda x: x["name"])
    return {
        "leagues": leagues,
        "id_to_key": id_to_key,
        "id_to_name": id_to_name,
        "name_map": name_map,
    }

def infer_name_from_match(match):
    for k in ["competition_name", "league_name", "league", "competition"]:
        val = str(match.get(k, "")).strip()
        if val and val not in ("0", "-1"):
            return clean_league_name(val)
    return "Campionato"

def resolve_match_league(match, chosen_lookup):
    comp_id = safe_int(match.get("competition_id"), 0)

    if comp_id and comp_id in chosen_lookup["id_to_key"]:
        key = chosen_lookup["id_to_key"][comp_id]
        name = chosen_lookup["id_to_name"].get(comp_id) or infer_name_from_match(match)
        return key, clean_league_name(name), comp_id

    for k in ["competition_name", "league_name", "league", "competition"]:
        val = normalize_text(match.get(k, ""))
        if val and val in chosen_lookup["name_map"]:
            key = sorted(chosen_lookup["name_map"][val])[0]
            chosen_row = next((x for x in chosen_lookup["leagues"] if x["key"] == key), None)
            resolved_season_id = safe_int((chosen_row or {}).get("season_id"), 0) or comp_id
            return key, clean_league_name(chosen_row["name"] if chosen_row else infer_name_from_match(match)), resolved_season_id

    return f"unmatched_{comp_id or '0'}", clean_league_name(infer_name_from_match(match)), comp_id

def infer_matchday(match):
    candidate_keys = [
        "game_week", "match_round", "round", "round_name", "week", "gw",
        "roundID", "round_id", "gameweek", "stage", "stage_name"
    ]
    for key in candidate_keys:
        val = match.get(key)
        if val not in (None, "", "null"):
            return val
    nested = deep_find_first(match, candidate_keys)
    return nested if nested not in (None, "", "null") else None

def is_completed_match(match):
    status = str(match.get("status", "")).strip().lower()
    if status in {"incomplete", "scheduled", "postponed", "cancelled"}:
        return False
    if safe_int(match.get("date_unix"), 0) > int(dt.datetime.now(dt.timezone.utc).timestamp()):
        return False
    hg = match.get("homeGoalCount")
    ag = match.get("awayGoalCount")
    return hg not in (None, "", "null") and ag not in (None, "", "null")

def map_match_status(match):
    if is_completed_match(match):
        return "FINISHED"
    status = str(match.get("status", "")).strip().lower()
    if status in {"live", "inplay", "in_play", "playing"}:
        return "IN_PLAY"
    if status in {"paused", "halftime", "half_time"}:
        return "PAUSED"
    return "TIMED"

def unix_to_utc_datetime(unix_ts):
    try:
        return pd.to_datetime(int(unix_ts), unit="s", utc=True)
    except Exception:
        return pd.NaT

@st.cache_data(show_spinner=False, ttl=60*15)
def fetch_matches_range(days_ahead: int) -> pd.DataFrame:
    chosen_lookup = fetch_league_lookup(CHOSEN_LEAGUES_ONLY)
    if not chosen_lookup.get("leagues") and AUTO_RETRY_ALL_LEAGUES:
        chosen_lookup = fetch_league_lookup(False)

    collected = []
    base_day = dt.datetime.now()
    for i in range(int(days_ahead) + 1):
        date_str = (base_day + dt.timedelta(days=i)).strftime("%Y-%m-%d")
        data = api_get("/todays-matches", {"date": date_str, "timezone": APP_TIMEZONE})
        if data is None:
            continue
        matches = data.get("data", [])
        if not isinstance(matches, list):
            continue
        for m in matches:
            if not isinstance(m, dict):
                continue
            league_key, league_name, season_id = resolve_match_league(m, chosen_lookup)
            home_id = safe_int(m.get("homeID"), 0)
            away_id = safe_int(m.get("awayID"), 0)
            row = {
                "match_id": safe_int(m.get("id"), 0),
                "utcDate": unix_to_utc_datetime(m.get("date_unix")),
                "status": map_match_status(m),
                "home_id": home_id,
                "home_name": m.get("home_name") or f"Team {home_id}",
                "away_id": away_id,
                "away_name": m.get("away_name") or f"Team {away_id}",
                "home_ft": safe_int(m.get("homeGoalCount"), None),
                "away_ft": safe_int(m.get("awayGoalCount"), None),
                "matchday": infer_matchday(m),
                "season_id": safe_int(season_id, 0),
                "league_key": league_key,
                "league_name": league_name,
            }
            collected.append(row)
    df = pd.DataFrame(collected)
    if not df.empty and "utcDate" in df.columns:
        df["utcDate"] = pd.to_datetime(df["utcDate"], errors="coerce", utc=True)
        df = df.sort_values(["league_name", "utcDate", "match_id"]).drop_duplicates(subset=["match_id"], keep="first")
    return df

@st.cache_data(show_spinner=False, ttl=60*60)
def fetch_league_season_matches(season_id: int) -> pd.DataFrame:
    if not season_id:
        return pd.DataFrame()
    data = api_get("/league-matches", {"season_id": int(season_id)})
    if data is None:
        return pd.DataFrame()
    raw = data.get("data", [])
    if not isinstance(raw, list):
        raw = []
    rows = []
    for m in raw:
        if not isinstance(m, dict):
            continue
        home_id = safe_int(m.get("homeID"), 0)
        away_id = safe_int(m.get("awayID"), 0)
        rows.append({
            "match_id": safe_int(m.get("id"), 0),
            "utcDate": unix_to_utc_datetime(m.get("date_unix")),
            "status": map_match_status(m),
            "home_id": home_id,
            "away_id": away_id,
            "home_name": m.get("home_name") or f"Team {home_id}",
            "away_name": m.get("away_name") or f"Team {away_id}",
            "home_ft": safe_int(m.get("homeGoalCount"), None),
            "away_ft": safe_int(m.get("awayGoalCount"), None),
            "matchday": infer_matchday(m),
            "season_id": int(season_id),
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df["utcDate"] = pd.to_datetime(df["utcDate"], errors="coerce", utc=True)
    return df

@st.cache_data(show_spinner=False, ttl=60*15)
def get_competition_matches(comp_key: str, date_from: str, date_to: str) -> pd.DataFrame:
    try:
        d1 = dt.date.fromisoformat(date_from)
        d2 = dt.date.fromisoformat(date_to)
        days_ahead = max((d2 - d1).days, 1)
    except Exception:
        days_ahead = 7
    df = fetch_matches_range(days_ahead)
    if df.empty:
        return df
    if comp_key:
        df = df[df["league_key"] == comp_key].copy()
    return df

@st.cache_data(show_spinner=False, ttl=60*60)
def get_team_season_matches(team_id: int, season_id: int) -> pd.DataFrame:
    df = fetch_league_season_matches(int(season_id))
    if df.empty:
        return df
    return df[(df["home_id"] == int(team_id)) | (df["away_id"] == int(team_id))].copy()

def goals_for_in_match(row: pd.Series, team_id: int):
    if row.get("status") != "FINISHED":
        return None
    if row.get("home_id") == team_id:
        return row.get("home_ft")
    if row.get("away_id") == team_id:
        return row.get("away_ft")
    return None

def goals_conceded_in_match(row: pd.Series, team_id: int):
    if row.get("status") != "FINISHED":
        return None
    if row.get("home_id") == team_id:
        return row.get("away_ft")
    if row.get("away_id") == team_id:
        return row.get("home_ft")
    return None

def bucket_0_4p(x: int) -> str:
    return f"G{x}" if x <= 3 else "G4+"

def dist_table(counts: pd.Series, total: int) -> pd.DataFrame:
    order = ["G0","G1","G2","G3","G4+"]
    out = []
    for k in order:
        c = int(counts.get(k, 0))
        p = (c / total) if total > 0 else 0.0
        out.append({"Bucket": k, "Count": c, "Percent": p})
    return pd.DataFrame(out)

def dist_compare_context(total_df: pd.DataFrame, ctx_df: pd.DataFrame, ctx_label: str) -> pd.DataFrame:
    """Tabella comparativa: % gol fatti Totale stagione vs contesto (Casa/Trasferta)."""
    order = ["G0","G1","G2","G3","G4+"]
    tot_counts = total_df["bucket_gf"].value_counts() if total_df is not None and len(total_df) else pd.Series(dtype=int)
    tot_n = int(len(total_df)) if total_df is not None else 0

    ctx_counts = ctx_df["bucket_gf"].value_counts() if ctx_df is not None and len(ctx_df) else pd.Series(dtype=int)
    ctx_n = int(len(ctx_df)) if ctx_df is not None else 0

    rows = []
    for k in order:
        rows.append({
            "Gol": k,
            "Totale": (float(tot_counts.get(k, 0)) / tot_n) if tot_n else 0.0,
            ctx_label: (float(ctx_counts.get(k, 0)) / ctx_n) if ctx_n else None,
        })
    return pd.DataFrame(rows)


def _fmt_pct_safe(x, decimals=1):
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):.{decimals}%}"
    except (TypeError, ValueError):
        return "-"


def build_standings_from_matches(season_matches: pd.DataFrame) -> pd.DataFrame:
    cols = ["Pos", "Squadra", "PG", "V", "N", "P", "GF", "GS", "DR", "Pt", "team_id"]
    if season_matches is None or season_matches.empty:
        return pd.DataFrame(columns=cols)
    finished = season_matches[season_matches["status"] == "FINISHED"].copy()
    if finished.empty:
        return pd.DataFrame(columns=cols)
    teams = {}
    def ensure(tid, name):
        if tid not in teams:
            teams[tid] = {"Squadra": name, "PG": 0, "V": 0, "N": 0, "P": 0, "GF": 0, "GS": 0, "Pt": 0, "team_id": tid}
    for _, r in finished.iterrows():
        hid = safe_int(r.get("home_id"), 0)
        aid = safe_int(r.get("away_id"), 0)
        if not hid or not aid:
            continue
        hg = safe_int(r.get("home_ft"), None)
        ag = safe_int(r.get("away_ft"), None)
        if hg is None or ag is None:
            continue
        ensure(hid, str(r.get("home_name") or f"Team {hid}"))
        ensure(aid, str(r.get("away_name") or f"Team {aid}"))
        h, a = teams[hid], teams[aid]
        h["PG"] += 1; a["PG"] += 1
        h["GF"] += hg; h["GS"] += ag
        a["GF"] += ag; a["GS"] += hg
        if hg > ag:
            h["V"] += 1; a["P"] += 1; h["Pt"] += 3
        elif hg < ag:
            a["V"] += 1; h["P"] += 1; a["Pt"] += 3
        else:
            h["N"] += 1; a["N"] += 1; h["Pt"] += 1; a["Pt"] += 1
    rows = []
    for v in teams.values():
        v = dict(v)
        v["DR"] = v["GF"] - v["GS"]
        rows.append(v)
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=cols)
    df = df.sort_values(["Pt", "DR", "GF", "Squadra"], ascending=[False, False, False, True]).reset_index(drop=True)
    df.insert(0, "Pos", range(1, len(df) + 1))
    return df[cols]


def _mean_num(df, col):
    try:
        if df is None or df.empty or col not in df.columns:
            return 0.0
        return float(pd.to_numeric(df[col], errors="coerce").dropna().mean())
    except Exception:
        return 0.0


def _trend_html(status):
    s = str(status or "DATI INSUFFICIENTI")
    if s == "NORMAL":
        return '<span class="mg-trend-ok">NORMAL</span>'
    if s == "WARNING":
        return '<span class="mg-trend-warn">WARNING</span>'
    if s == "CAMBIO CONFERMATO":
        return '<span class="mg-trend-bad">CAMBIO CONFERMATO</span>'
    return '<span class="mg-mini">DATI INSUFFICIENTI</span>'


# =========================
# RIEPILOGO GIOCATE
# =========================
view_options = ["Riepilogo giocate", "Dettaglio partita"]
default_view = st.session_state.get("view_mode", "Riepilogo giocate")
if default_view not in view_options:
    default_view = "Riepilogo giocate"
view_mode = st.radio("Vista", view_options, index=view_options.index(default_view), horizontal=True)
st.session_state["view_mode"] = view_mode


def _summary_match_day_label(ts):
    try:
        local_dt = pd.to_datetime(ts, utc=True).tz_convert(APP_TIMEZONE)
        giorni_it = {
            0: "Lunedì",
            1: "Martedì",
            2: "Mercoledì",
            3: "Giovedì",
            4: "Venerdì",
            5: "Sabato",
            6: "Domenica",
        }
        giorno = giorni_it.get(local_dt.weekday(), "")
        return f"{giorno} {local_dt.strftime('%d/%m/%Y')}"
    except Exception:
        return "Data non disponibile"


def _summary_match_time_label(ts):
    try:
        return pd.to_datetime(ts, utc=True).tz_convert(APP_TIMEZONE).strftime("%H:%M")
    except Exception:
        return "--:--"


def analyze_match_for_summary(sel: dict, min_prob: float = 0.72, include_all_candidates: bool = False) -> dict | None:
    """Analisi sintetica per la pagina riepilogo usando la stessa logica del software."""
    try:
        home_id = int(sel.get("home_id") or 0)
        away_id = int(sel.get("away_id") or 0)
        season_id = int(sel.get("season_id") or 0)
        home_name = str(sel.get("home_name") or "Casa")
        away_name = str(sel.get("away_name") or "Trasferta")
        league_name = str(sel.get("league_name") or "Campionato")
        match_id = sel.get("match_id")
        utc_date = sel.get("utcDate")
        if not home_id or not away_id or not season_id:
            return None

        N_LAST = 10
        CTX_MIN_MATCHES = 6
        MIN_LAMBDA_MATCHES = 4
        SPLIT_MIN_MATCHES = 6

        def safe_mean_local(s: pd.Series) -> float:
            try:
                s = pd.to_numeric(s, errors="coerce").dropna()
                return float(s.mean()) if len(s) else 0.0
            except Exception:
                return 0.0

        def pct_dict_from_buckets(series_buckets: pd.Series) -> dict:
            order = ["G0", "G1", "G2", "G3", "G4+"]
            total = int(series_buckets.sum()) if series_buckets is not None else 0
            out = {k: 0.0 for k in order}
            if total <= 0:
                return out
            for k in order:
                out[k] = float(series_buckets.get(k, 0)) / total
            return out

        def is_flat(d: dict) -> bool:
            vals = [d.get("G0", 0), d.get("G1", 0), d.get("G2", 0), d.get("G3", 0), d.get("G4+", 0)]
            mx = max(vals) if vals else 0
            pairs = [d.get("G0", 0) + d.get("G1", 0), d.get("G1", 0) + d.get("G2", 0), d.get("G2", 0) + d.get("G3", 0)]
            return (mx < 0.35) and all(p < 0.55 for p in pairs)

        def pois_p0(lam: float) -> float:
            lam = max(float(lam), 0.0)
            return math.exp(-lam)

        def pois_pmf(lam: float, k: int) -> float:
            lam = max(float(lam), 0.0)
            if k < 0:
                return 0.0
            p0 = math.exp(-lam)
            if k == 0:
                return p0
            p = p0
            for i in range(1, k + 1):
                p *= lam / i
            return p

        def pois_cdf(lam: float, k: int) -> float:
            if k < 0:
                return 0.0
            s = 0.0
            for i in range(0, k + 1):
                s += pois_pmf(lam, i)
            return min(max(s, 0.0), 1.0)

        def pois_range_prob(lam: float, lo: int, hi: int, hi_is_4plus: bool = False) -> float:
            lam = max(float(lam), 0.0)
            lo = int(lo)
            hi = int(hi)
            if lo > hi:
                return 0.0
            if hi_is_4plus and hi == 4:
                p_le_3 = pois_cdf(lam, 3)
                p_lo_3 = pois_cdf(lam, 3) - pois_cdf(lam, lo - 1)
                return p_lo_3 + (1.0 - p_le_3)
            return pois_cdf(lam, hi) - pois_cdf(lam, lo - 1)

        def compute_bar_local(dist: dict):
            g01 = dist["G0"] + dist["G1"]
            g12 = dist["G1"] + dist["G2"]
            g23 = dist["G2"] + dist["G3"]
            if g01 >= 0.55:
                return "BASSO", g01, g12, g23
            if g12 >= 0.55:
                return "MEDIO", g01, g12, g23
            if g23 >= 0.55:
                return "ALTO", g01, g12, g23
            return None, g01, g12, g23

        def ranges_for_bar_local(bar: str):
            if bar == "BASSO":
                return ["0–1", "0–2"]
            if bar == "MEDIO":
                return ["1–2", "1–3"]
            if bar == "ALTO":
                return ["2–3", "2–4"]
            return []

        def range_includes(range_str: str, k: str) -> bool:
            lo, hi = range_str.split("–")
            lo_i = int(lo)
            hi_i = int(hi)
            v = 4 if k == "G4+" else int(k[1])
            return lo_i <= v <= hi_i

        def mg_cover(range_str: str, distd: dict) -> float:
            return sum(distd[k] for k in ["G0", "G1", "G2", "G3", "G4+"] if range_includes(range_str, k))

        def excluded_strong_events(range_str: str, distd: dict, thr: float = 0.30):
            out = []
            for k, p in distd.items():
                if (not range_includes(range_str, k)) and p >= thr:
                    out.append((k, p))
            return out

        def trend_metrics_local(team_df: pd.DataFrame, team_label: str) -> dict:
            out = {
                "Squadra": team_label,
                "Match stagione (FINISHED)": int(len(team_df)),
                "Media gol stagione": float(team_df["gf"].mean()) if len(team_df) else 0.0,
                "Match usati ultime 6": int(min(6, len(team_df))),
                "Media gol ultime 6": 0.0,
                "Delta (ult6 - stag)": 0.0,
                "Evento estremo (ult6)": "",
                "Estremi (ult6)": 0,
                "Stato": "DATI INSUFFICIENTI",
            }
            if len(team_df) < 3:
                return out
            recent = team_df.sort_values("utcDate", ascending=False).head(6).copy()
            m6 = float(recent["gf"].mean()) if len(recent) else 0.0
            delta = m6 - float(out["Media gol stagione"])
            out["Media gol ultime 6"] = m6
            out["Delta (ult6 - stag)"] = delta
            if delta >= 0:
                out["Evento estremo (ult6)"] = "3+"
                out["Estremi (ult6)"] = int((recent["gf"] >= 3).sum())
            else:
                out["Evento estremo (ult6)"] = "0"
                out["Estremi (ult6)"] = int((recent["gf"] == 0).sum())
            if len(recent) < 6:
                out["Stato"] = "DATI INSUFFICIENTI"
                return out
            abs_delta = abs(delta)
            extremes = out["Estremi (ult6)"]
            if abs_delta >= 0.7 and extremes >= 3:
                out["Stato"] = "CAMBIO CONFERMATO"
            elif abs_delta >= 0.4:
                out["Stato"] = "WARNING"
            else:
                out["Stato"] = "NORMAL"
            return out

        def ctx_or_total_local(split_df: pd.DataFrame, total_df: pd.DataFrame) -> pd.DataFrame:
            return split_df if split_df is not None and len(split_df) >= CTX_MIN_MATCHES else total_df

        def apply_global_coherence_local(results: list, *, bar_team_ctx: str, bar_opp_ctx: str, lambda_total: float,
                                         team_for_ctx: dict, opp_for_ctx: dict, team_conc_ctx: dict, opp_conc_ctx: dict,
                                         trend_delta: float):
            notes = []
            if not results:
                return results, notes
            p0_team = float(team_for_ctx.get("G0", 0.0))
            p0_opp = float(opp_for_ctx.get("G0", 0.0))
            p0_avg_for = (p0_team + p0_opp) / 2.0
            low_match = (bar_team_ctx == "BASSO") or (bar_opp_ctx == "BASSO") or (p0_avg_for >= 0.28) or (lambda_total <= 2.20)
            high_match = (lambda_total >= 2.90) and (p0_team <= 0.22) and (p0_opp <= 0.22)
            filtered = []
            for r in results:
                name = (r.get("name") or "").upper()
                kind = (r.get("kind") or "").upper()
                if ("BTTS" in kind) and ("SI" in name) and low_match:
                    continue
                if ("BTTS" in kind) and ("NO" in name) and high_match:
                    continue
                if kind == "UNDER" and ("2.5" in name) and (high_match or trend_delta >= 0.7):
                    continue
                if kind == "OVER" and (low_match or trend_delta <= -0.7):
                    continue
                filtered.append(r)
            has_under = [r for r in filtered if (r.get("kind", "").upper() == "UNDER")]
            has_over = [r for r in filtered if (r.get("kind", "").upper() == "OVER")]
            if has_under and has_over:
                best_under = max(has_under, key=lambda x: float(x.get("prob", 0.0)))
                best_over = max(has_over, key=lambda x: float(x.get("prob", 0.0)))
                if float(best_under.get("prob", 0.0)) >= float(best_over.get("prob", 0.0)):
                    filtered = [r for r in filtered if r is best_under or (r.get("kind", "").upper() != "OVER")]
                else:
                    filtered = [r for r in filtered if r is best_over or (r.get("kind", "").upper() != "UNDER")]
            btts = [r for r in filtered if (r.get("kind", "").upper() == "BTTS")]
            if len(btts) >= 2:
                best = max(btts, key=lambda x: float(x.get("prob", 0.0)))
                filtered = [r for r in filtered if (r.get("kind", "").upper() != "BTTS") or (r is best)]
            return filtered, notes

        def pick_with_priority_local(results, conflict=False):
            mg = [r for r in results if r.get("kind") == "MG"]
            under = [r for r in results if r.get("kind") == "UNDER"]
            if conflict:
                mg = []
            for r in mg:
                if r.get("label") == "ROBUSTO" and float(r.get("cover", 0)) >= 0.60:
                    return r
            for r in under:
                if r.get("label") == "ROBUSTO":
                    if not any(x.get("label") == "ROBUSTO" for x in mg):
                        return r
            return results[0] if results else None

        home_season = get_team_season_matches(home_id, season_id)
        away_season = get_team_season_matches(away_id, season_id)
        hs = home_season[home_season["status"] == "FINISHED"].copy()
        aw = away_season[away_season["status"] == "FINISHED"].copy()
        if hs.empty or aw.empty:
            return None

        hs["gf"] = hs.apply(lambda r: goals_for_in_match(r, home_id), axis=1)
        aw["gf"] = aw.apply(lambda r: goals_for_in_match(r, away_id), axis=1)
        hs["ga"] = hs.apply(lambda r: goals_conceded_in_match(r, home_id), axis=1)
        aw["ga"] = aw.apply(lambda r: goals_conceded_in_match(r, away_id), axis=1)
        hs = hs.dropna(subset=["gf"])
        aw = aw.dropna(subset=["gf"])
        if hs.empty or aw.empty:
            return None

        hs["bucket_gf"] = hs["gf"].astype(int).apply(bucket_0_4p)
        aw["bucket_gf"] = aw["gf"].astype(int).apply(bucket_0_4p)
        hs_home_gf = hs[hs["home_id"] == home_id].copy()
        hs_away_gf = hs[hs["away_id"] == home_id].copy()
        aw_home_gf = aw[aw["home_id"] == away_id].copy()
        aw_away_gf = aw[aw["away_id"] == away_id].copy()
        hs_home_ga = hs[hs["home_id"] == home_id].copy()
        hs_away_ga = hs[hs["away_id"] == home_id].copy()
        aw_home_ga = aw[aw["home_id"] == away_id].copy()
        aw_away_ga = aw[aw["away_id"] == away_id].copy()
        for _df in [hs_home_gf, hs_away_gf, aw_home_gf, aw_away_gf]:
            if not _df.empty:
                _df["bucket_gf"] = _df["gf"].astype(int).apply(bucket_0_4p)

        home_tr = trend_metrics_local(hs, home_name)
        away_tr = trend_metrics_local(aw, away_name)
        home_home = hs[hs["home_id"] == home_id].copy()
        home_home["ga"] = home_home.apply(lambda r: goals_conceded_in_match(r, home_id), axis=1)
        home_home = home_home.dropna(subset=["ga"]).sort_values("utcDate", ascending=False).head(N_LAST)
        home_home["bucket_ga"] = home_home["ga"].astype(int).apply(bucket_0_4p)
        away_away = aw[aw["away_id"] == away_id].copy()
        away_away["ga"] = away_away.apply(lambda r: goals_conceded_in_match(r, away_id), axis=1)
        away_away = away_away.dropna(subset=["ga"]).sort_values("utcDate", ascending=False).head(N_LAST)
        away_away["bucket_ga"] = away_away["ga"].astype(int).apply(bucket_0_4p)

        gf_home = safe_mean_local(hs_home_gf["gf"]) if len(hs_home_gf) >= MIN_LAMBDA_MATCHES else safe_mean_local(hs["gf"])
        ga_home = safe_mean_local(hs_home_ga["ga"]) if len(hs_home_ga) >= MIN_LAMBDA_MATCHES else safe_mean_local(hs["ga"])
        gf_away = safe_mean_local(aw_away_gf["gf"]) if len(aw_away_gf) >= MIN_LAMBDA_MATCHES else safe_mean_local(aw["gf"])
        ga_away = safe_mean_local(aw_away_ga["ga"]) if len(aw_away_ga) >= MIN_LAMBDA_MATCHES else safe_mean_local(aw["ga"])
        lambda_home = (gf_home + ga_away) / 2.0
        lambda_away = (gf_away + ga_home) / 2.0
        lambda_total = lambda_home + lambda_away

        results_all = []
        for team_choice in ["Casa", "Trasferta"]:
            if team_choice == "Casa":
                team_name = home_name
                opp_name = away_name
                team_dist = pct_dict_from_buckets(hs["bucket_gf"].value_counts())
                opp_conc = pct_dict_from_buckets(away_away["bucket_ga"].value_counts())
                trend_row = home_tr
                team_for_ctx_df = ctx_or_total_local(hs_home_gf, hs)
                opp_for_ctx_df = ctx_or_total_local(aw_away_gf, aw)
                team_conc_ctx = pct_dict_from_buckets(home_home["bucket_ga"].value_counts())
                opp_conc_ctx = pct_dict_from_buckets(away_away["bucket_ga"].value_counts())
                split_src = hs_home_gf
                lambda_team_ctx = lambda_home
            else:
                team_name = away_name
                opp_name = home_name
                team_dist = pct_dict_from_buckets(aw["bucket_gf"].value_counts())
                opp_conc = pct_dict_from_buckets(home_home["bucket_ga"].value_counts())
                trend_row = away_tr
                team_for_ctx_df = ctx_or_total_local(aw_away_gf, aw)
                opp_for_ctx_df = ctx_or_total_local(hs_home_gf, hs)
                team_conc_ctx = pct_dict_from_buckets(away_away["bucket_ga"].value_counts())
                opp_conc_ctx = pct_dict_from_buckets(home_home["bucket_ga"].value_counts())
                split_src = aw_away_gf
                lambda_team_ctx = lambda_away

            if is_flat(team_dist):
                continue
            bar, g01, g12, g23 = compute_bar_local(team_dist)
            if bar is None:
                continue
            ranges = ranges_for_bar_local(bar)
            split_dist = pct_dict_from_buckets(split_src["bucket_gf"].value_counts()) if len(split_src) else {k: 0.0 for k in ["G0", "G1", "G2", "G3", "G4+"]}
            bar_split, _, _, _ = compute_bar_local(split_dist)
            split_ok = True
            if len(split_src) >= SPLIT_MIN_MATCHES and bar_split is not None and bar_split != bar:
                split_ok = False

            g0c = opp_conc["G0"]
            g3pc = opp_conc["G3"] + opp_conc["G4+"]
            push_low = g0c >= 0.30
            push_high = g3pc >= 0.20
            conflict = push_low and push_high
            chosen_range = None
            if not conflict:
                if push_low:
                    chosen_range = ranges[0]
                elif push_high and len(ranges) > 1:
                    chosen_range = ranges[1]
                else:
                    chosen_range = ranges[0]

            mg_results = []
            if not conflict and chosen_range:
                exc = excluded_strong_events(chosen_range, team_dist, 0.30)
                cover = mg_cover(chosen_range, team_dist)
                if not exc:
                    lose_ev = [(k, p) for k, p in team_dist.items() if not range_includes(chosen_range, k)]
                    lose_sum = sum(p for _, p in lose_ev)
                    label = "ROBUSTO" if lose_sum < 0.30 else ("NEUTRO" if lose_sum < 0.40 else "INSTABILE")
                    if (not split_ok) and (len(split_src) >= SPLIT_MIN_MATCHES):
                        if label == "ROBUSTO":
                            label = "NEUTRO"
                        elif label == "NEUTRO":
                            label = "INSTABILE"
                    if label != "INSTABILE":
                        mg_results.append({
                            "name": f"MG {chosen_range} {team_name}",
                            "label": label,
                            "kind": "MG",
                            "cover": cover,
                            "prob": pois_range_prob(lambda_team_ctx, int(chosen_range.split("–")[0]), int(chosen_range.split("–")[1]), hi_is_4plus=True),
                        })

            team_for_ctx = pct_dict_from_buckets(team_for_ctx_df["bucket_gf"].value_counts())
            opp_for_ctx = pct_dict_from_buckets(opp_for_ctx_df["bucket_gf"].value_counts())
            alt_results = []
            team_for_g3p = team_for_ctx["G3"] + team_for_ctx["G4+"]
            opp_for_g3p = opp_for_ctx["G3"] + opp_for_ctx["G4+"]
            team_conc_g3p = team_conc_ctx["G3"] + team_conc_ctx["G4+"]
            opp_conc_g3p = opp_conc_ctx["G3"] + opp_conc_ctx["G4+"]
            under_ok = (team_for_g3p < 0.25) and (opp_for_g3p < 0.25) and (team_conc_g3p < 0.20) and (opp_conc_g3p < 0.20)
            if under_ok:
                low_mass = (team_for_ctx["G0"] + team_for_ctx["G1"] + opp_for_ctx["G0"] + opp_for_ctx["G1"] + team_conc_ctx["G0"] + team_conc_ctx["G1"] + opp_conc_ctx["G0"] + opp_conc_ctx["G1"]) / 4.0
                under_choice = "Under 2.5" if low_mass >= 1.05 else "Under 3.5"
                alt_results.append({"name": under_choice, "label": "ROBUSTO", "kind": "UNDER", "prob": (pois_cdf(lambda_total, 2) if under_choice == "Under 2.5" else pois_cdf(lambda_total, 3))})

            p_team_scores = 1.0 - ((team_for_ctx["G0"] + opp_conc_ctx["G0"]) / 2.0)
            p_opp_scores = 1.0 - ((opp_for_ctx["G0"] + team_conc_ctx["G0"]) / 2.0)
            p_btts_yes = p_team_scores * p_opp_scores
            p_home_scores_pois = 1.0 - pois_p0(lambda_home)
            p_away_scores_pois = 1.0 - pois_p0(lambda_away)
            p_btts_yes_pois = p_home_scores_pois * p_away_scores_pois
            p_btts_no_pois = 1.0 - p_btts_yes_pois
            bar_team_ctx, _, _, _ = compute_bar_local(team_for_ctx)
            bar_opp_ctx, _, _, _ = compute_bar_local(opp_for_ctx)
            btts_yes_allowed = ((bar_team_ctx != "BASSO") and (bar_opp_ctx != "BASSO") and (team_for_ctx["G0"] <= 0.30) and (opp_for_ctx["G0"] <= 0.30) and (opp_conc_ctx["G0"] <= 0.35) and (team_conc_ctx["G0"] <= 0.35))
            if btts_yes_allowed and (p_btts_yes >= 0.55) and (p_team_scores >= 0.65) and (p_opp_scores >= 0.65):
                alt_results.append({"name": "BTTS SI", "label": "NEUTRO", "kind": "BTTS", "prob": p_btts_yes_pois})
            else:
                if (p_btts_no_pois >= 0.55) and ((bar_team_ctx == "BASSO") or (bar_opp_ctx == "BASSO")):
                    alt_results.append({"name": "BTTS NO", "label": "NEUTRO", "kind": "BTTS", "prob": p_btts_no_pois})
                elif (p_btts_yes <= 0.45):
                    alt_results.append({"name": "BTTS NO", "label": "NEUTRO", "kind": "BTTS", "prob": p_btts_no_pois})
            team_g2p = team_dist["G2"] + team_dist["G3"] + team_dist["G4+"]
            team_g3p = team_dist["G3"] + team_dist["G4+"]
            opp_2p_conc = opp_conc["G2"] + opp_conc["G3"] + opp_conc["G4+"]
            over_ok = (team_g2p >= 0.55) and (opp_2p_conc >= 0.45) and (team_dist["G0"] < 0.25)
            if over_ok:
                over_choice = "Over 2.5" if team_g3p >= 0.25 else "Over 1.5 squadra"
                alt_results.append({"name": over_choice, "label": "NEUTRO", "kind": "OVER", "prob": (1.0 - pois_cdf(lambda_total, 2) if over_choice == "Over 2.5" else (1.0 - pois_p0(lambda_team_ctx)))})

            delta = float(trend_row.get("Delta (ult6 - stag)", 0.0))
            all_results = mg_results + alt_results
            all_results, _ = apply_global_coherence_local(
                all_results,
                bar_team_ctx=bar_team_ctx,
                bar_opp_ctx=bar_opp_ctx,
                lambda_total=lambda_total,
                team_for_ctx=team_for_ctx,
                opp_for_ctx=opp_for_ctx,
                team_conc_ctx=team_conc_ctx,
                opp_conc_ctx=opp_conc_ctx,
                trend_delta=delta,
            )
            order = {"ROBUSTO": 0, "NEUTRO": 1, "INSTABILE": 2}
            all_results = sorted(all_results, key=lambda x: (order.get(x.get("label", "NEUTRO"), 1), -float(x.get("prob", 0.0))))
            picked = pick_with_priority_local(all_results, conflict=conflict)
            if picked:
                picked = dict(picked)
                picked["team_choice"] = team_choice
                results_all.append(picked)
            results_all.extend([dict(r, team_choice=team_choice) for r in all_results])

        if not results_all:
            return None

        filtered_results = []
        seen = set()
        for r in results_all:
            name = str(r.get("name") or "")
            kind = str(r.get("kind") or "")
            prob = float(r.get("prob", 0.0) or 0.0)
            normalized = name.replace("-", "–")
            if (not include_all_candidates) and kind == "MG" and ("0–1" in normalized or "1–2" in normalized):
                continue
            if prob < float(min_prob):
                continue
            key = (name, kind, r.get("team_choice"))
            if key in seen:
                continue
            seen.add(key)
            filtered_results.append(r)

        if not filtered_results:
            return None

        label_rank = {"ROBUSTO": 0, "NEUTRO": 1, "INSTABILE": 2}
        # Nel riepilogo conta solo la probabilità più alta.
        # Robustezza e priorità del motore restano disponibili nel dettaglio partita.
        best = max(filtered_results, key=lambda r: float(r.get("prob", 0.0)))
        return {
            "match_id": match_id,
            "league_key": sel.get("league_key"),
            "league_name": league_name,
            "home_name": home_name,
            "away_name": away_name,
            "utcDate": utc_date,
            "best_name": best.get("name"),
            "best_kind": best.get("kind"),
            "best_prob": float(best.get("prob", 0.0)),
            "best_label": best.get("label", "NEUTRO"),
            "candidates": sorted(filtered_results, key=lambda r: (label_rank.get(r.get("label", "NEUTRO"), 1), -float(r.get("prob", 0.0)))) if include_all_candidates else sorted(filtered_results, key=lambda r: (label_rank.get(r.get("label", "NEUTRO"), 1), -float(r.get("prob", 0.0))))[:6],
        }
    except Exception:
        return None

if view_mode == "Riepilogo giocate":
    # Filtri persistenti: restano identici anche dopo apertura/chiusura dettaglio.
    if "_summary_days_saved" not in st.session_state:
        st.session_state["_summary_days_saved"] = 7
    if "_summary_prob_saved" not in st.session_state:
        st.session_state["_summary_prob_saved"] = 70

    if "summary_days" not in st.session_state:
        st.session_state["summary_days"] = int(st.session_state["_summary_days_saved"])
    if "summary_prob" not in st.session_state:
        st.session_state["summary_prob"] = int(st.session_state["_summary_prob_saved"])

    def _persist_summary_days():
        st.session_state["_summary_days_saved"] = int(st.session_state["summary_days"])

    def _persist_summary_prob():
        st.session_state["_summary_prob_saved"] = int(st.session_state["summary_prob"])

    st.markdown("""
        <div class="summary-hero">
            <div>
                <div class="summary-eyebrow">SELEZIONE AUTOMATICA</div>
                <div class="summary-title">Riepilogo giocate</div>
                <div class="summary-subtitle">Una sola proposta per partita: il mercato con la percentuale più alta.</div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown('<div class="summary-filter-title">FILTRI</div>', unsafe_allow_html=True)
        col_filters1, col_filters2 = st.columns([1, 1], gap="large")
        with col_filters1:
            days_ahead = st.slider(
                "Finestra partite (giorni avanti)",
                1, 14,
                key="summary_days",
                on_change=_persist_summary_days,
            )
        with col_filters2:
            min_summary_prob = st.slider(
                "Probabilità minima",
                68, 95,
                key="summary_prob",
                on_change=_persist_summary_prob,
            )

    with st.spinner("Carico partite da FootyStats..."):
        all_matches_df = fetch_matches_range(days_ahead)

    if all_matches_df.empty:
        st.warning("Nessuna partita trovata nella finestra scelta. Controlla la chiave API o aumenta i giorni.")
        st.stop()

    summary_base = all_matches_df[all_matches_df["status"].isin(["SCHEDULED", "TIMED", "IN_PLAY", "PAUSED"])].copy()
    if summary_base.empty:
        summary_base = all_matches_df.copy()
    summary_base = summary_base.sort_values(["utcDate", "league_name", "home_name", "away_name"]).reset_index(drop=True)

    with st.spinner("Analizzo le partite e filtro solo le giocate >= soglia..."):
        summary_rows = []
        progress = st.progress(0)
        total_matches = len(summary_base)
        for pos, (_, row) in enumerate(summary_base.iterrows(), start=1):
            res = analyze_match_for_summary(row.to_dict(), min_prob=float(min_summary_prob) / 100.0)
            if res is not None:
                summary_rows.append(res)
            progress.progress(pos / total_matches if total_matches else 1.0)
        progress.empty()

    st.markdown(
        f"""<div class="summary-countbar">
            <div><b>{len(summary_rows)}</b> partite selezionate</div>
            <div>Soglia <b>{min_summary_prob}%</b> · prossimi <b>{days_ahead}</b> giorni</div>
        </div>""",
        unsafe_allow_html=True,
    )

    if not summary_rows:
        st.warning("Nessuna partita supera la soglia richiesta con i filtri attuali.")
        st.stop()

    summary_df = pd.DataFrame(summary_rows)

    # Ordina usando la data locale italiana, cioè la stessa mostrata nel riepilogo.
    _utc_dates = pd.to_datetime(summary_df["utcDate"], utc=True, errors="coerce")
    _local_dates = _utc_dates.dt.tz_convert(APP_TIMEZONE)
    summary_df["_local_day_sort"] = _local_dates.dt.date
    summary_df["_local_time_sort"] = _local_dates

    # Prima il giorno; dentro lo stesso giorno probabilità decrescente.
    # L'orario locale viene usato solo come spareggio.
    summary_df = summary_df.sort_values(
        ["_local_day_sort", "best_prob", "_local_time_sort"],
        ascending=[True, False, True],
        na_position="last",
    ).drop(columns=["_local_day_sort", "_local_time_sort"])
    current_day = None
    for _, row in summary_df.iterrows():
        day_label = _summary_match_day_label(row.get("utcDate"))
        if day_label != current_day:
            current_day = day_label
            st.markdown(f'<div class="summary-day">📅 {day_label}</div>', unsafe_allow_html=True)

        prob_value = float(row.get("best_prob", 0.0))
        robust = str(row.get("best_label") or "NEUTRO")
        robust_class = "robusto" if robust == "ROBUSTO" else ("instabile" if robust == "INSTABILE" else "neutro")

        with st.container(border=True):
            c1, c2, c3 = st.columns([5.2, 2.0, 1.35], vertical_alignment="center")
            with c1:
                st.markdown(
                    f"""<div class="summary-match">
                        <div class="summary-time">{_summary_match_time_label(row.get('utcDate'))}</div>
                        <div class="summary-teams">{row.get('home_name')} <span>vs</span> {row.get('away_name')}</div>
                        <div class="summary-league">{row.get('league_name')}</div>
                        <div class="summary-market">{row.get('best_name')}</div>
                        <span class="summary-badge {robust_class}">{robust}</span>
                    </div>""",
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown(
                    f"""<div class="summary-prob">
                        <div class="summary-prob-label">PROBABILITÀ</div>
                        <div class="summary-prob-value">{prob_value:.0%}</div>
                    </div>""",
                    unsafe_allow_html=True,
                )
            with c3:
                btn_key = f"open_{row.get('match_id')}_{row.get('league_key')}"
                if st.button("Dettaglio →", key=btn_key, use_container_width=True):
                    st.session_state["_summary_days_saved"] = int(st.session_state.get("summary_days", st.session_state.get("_summary_days_saved", 7)))
                    st.session_state["_summary_prob_saved"] = int(st.session_state.get("summary_prob", st.session_state.get("_summary_prob_saved", 70)))
                    st.session_state["summary_target_match_id"] = row.get("match_id")
                    st.session_state["summary_target_league"] = row.get("league_key")
                    st.session_state["view_mode"] = "Dettaglio partita"
                    st.rerun()
    st.stop()
else:

    if st.button("← Torna al riepilogo giocate", key="back_to_summary"):
        st.session_state["view_mode"] = "Riepilogo giocate"
        st.rerun()

    col1, col2 = st.columns([1.2, 1])

    with col1:
        days_ahead = st.slider("Finestra partite (giorni avanti)", 1, 14, 7)
        today = dt.date.today()
        date_from = today.isoformat()
        date_to = (today + dt.timedelta(days=int(days_ahead))).isoformat()

    with st.spinner("Carico partite da FootyStats..."):
        all_matches_df = fetch_matches_range(days_ahead)

    if all_matches_df.empty:
        st.warning("Nessuna partita trovata nella finestra scelta. Controlla la chiave API o aumenta i giorni.")
        st.stop()

    available_leagues = (
        all_matches_df[["league_key", "league_name"]]
        .dropna()
        .drop_duplicates()
        .sort_values("league_name")
    )

    if available_leagues.empty:
        st.warning("FootyStats non ha restituito campionati utilizzabili per la finestra scelta.")
        st.stop()

    league_options = available_leagues["league_name"].tolist()
    league_key_by_name = dict(zip(available_leagues["league_name"], available_leagues["league_key"]))

    forced_league_key = st.session_state.get("summary_target_league")
    default_league_index = 0
    if forced_league_key in set(league_key_by_name.values()):
        for _i, _name in enumerate(league_options):
            if league_key_by_name.get(_name) == forced_league_key:
                default_league_index = _i
                break

    with col1:
        comp_label = st.selectbox("Campionato", league_options, index=default_league_index)
        comp = league_key_by_name[comp_label]

    matches_df = all_matches_df[all_matches_df["league_key"] == comp].copy()

    if matches_df.empty:
        st.warning("Nessuna partita trovata per il campionato scelto nella finestra indicata.")
        st.stop()

    upcoming = matches_df[matches_df["status"].isin(["SCHEDULED", "TIMED", "IN_PLAY", "PAUSED"])].copy()
    if upcoming.empty:
        upcoming = matches_df.copy()

    def match_label(r):
        dt_str = ""
        if pd.notna(r.get("utcDate")):
            try:
                dt_str = r["utcDate"].tz_convert(APP_TIMEZONE).strftime("%d/%m %H:%M")
            except Exception:
                dt_str = str(r.get("utcDate"))
        md = r.get("matchday")
        md_str = f"MD {md} - " if pd.notna(md) and md not in ("", None) else ""
        return f"{md_str}{dt_str} | {r['home_name']} vs {r['away_name']}"

    upcoming = upcoming.sort_values("utcDate") if "utcDate" in upcoming.columns else upcoming
    labels = [match_label(r) for _, r in upcoming.iterrows()]

    default_match_index = 0
    forced_match_id = st.session_state.get("summary_target_match_id")
    if forced_match_id is not None and not upcoming.empty and "match_id" in upcoming.columns:
        try:
            _ids = upcoming["match_id"].astype(str).tolist()
            if str(forced_match_id) in _ids:
                default_match_index = _ids.index(str(forced_match_id))
        except Exception:
            pass

    with col2:
        st.subheader("Seleziona partita")
        idx = st.selectbox("Partita", list(range(len(labels))), index=default_match_index, format_func=lambda i: labels[i])

    sel = upcoming.iloc[int(idx)]
    st.session_state.pop("summary_target_match_id", None)
    st.session_state.pop("summary_target_league", None)
    home_id = int(sel["home_id"])
    away_id = int(sel["away_id"])
    home_name = sel["home_name"]
    away_name = sel["away_name"]
    season_id = int(sel.get("season_id") or 0)
    comp = season_id

    # Toggle per mostrare dettagli dei calcoli Poisson (medie e λ)
    show_poisson_debug = st.checkbox("Mostra dettagli Poisson (medie & λ usati nel calcolo)", value=False)

    st.divider()
    st.subheader("Calcolo distribuzioni")

    N_LAST = 10

    with st.spinner("Scarico match stagione e calcolo..."):
        home_season = get_team_season_matches(home_id, comp)
        away_season = get_team_season_matches(away_id, comp)

        hs = home_season[home_season["status"] == "FINISHED"].copy()
        aw = away_season[away_season["status"] == "FINISHED"].copy()

        hs["gf"] = hs.apply(lambda r: goals_for_in_match(r, home_id), axis=1)
        aw["gf"] = aw.apply(lambda r: goals_for_in_match(r, away_id), axis=1)

        # Gol SUBITI (serve per Poisson / BTTS / Under match)
        hs["ga"] = hs.apply(lambda r: goals_conceded_in_match(r, home_id), axis=1)
        aw["ga"] = aw.apply(lambda r: goals_conceded_in_match(r, away_id), axis=1)

        hs = hs.dropna(subset=["gf"])
        aw = aw.dropna(subset=["gf"])

        hs["bucket_gf"] = hs["gf"].astype(int).apply(bucket_0_4p)
        aw["bucket_gf"] = aw["gf"].astype(int).apply(bucket_0_4p)

    
        # --- SPLIT CASA/TRASFERTA: gol FATTI (stagione) ---
        hs_home_gf = hs[hs["home_id"] == home_id].copy()
        hs_away_gf = hs[hs["away_id"] == home_id].copy()
        aw_home_gf = aw[aw["home_id"] == away_id].copy()
        aw_away_gf = aw[aw["away_id"] == away_id].copy()

        # --- SPLIT CASA/TRASFERTA: gol SUBITI (stagione) ---
        hs_home_ga = hs[hs["home_id"] == home_id].copy()
        hs_away_ga = hs[hs["away_id"] == home_id].copy()
        aw_home_ga = aw[aw["home_id"] == away_id].copy()
        aw_away_ga = aw[aw["away_id"] == away_id].copy()

        for _df in [hs_home_gf, hs_away_gf, aw_home_gf, aw_away_gf]:
            if not _df.empty:
                _df["bucket_gf"] = _df["gf"].astype(int).apply(bucket_0_4p)

    # --- Indicatori trend (ultime 6 vs stagione) sui gol FATTI ---
        def _trend_metrics(team_df: pd.DataFrame, team_label: str) -> dict:
            out = {
                "Squadra": team_label,
                "Match stagione (FINISHED)": int(len(team_df)),
                "Media gol stagione": float(team_df["gf"].mean()) if len(team_df) else 0.0,
                "Match usati ultime 6": int(min(6, len(team_df))),
                "Media gol ultime 6": 0.0,
                "Delta (ult6 - stag)": 0.0,
                "Evento estremo (ult6)": "",
                "Estremi (ult6)": 0,
                "Stato": "DATI INSUFFICIENTI",
            }
            if len(team_df) < 3:
                return out

            recent = team_df.sort_values("utcDate", ascending=False).head(6).copy()
            m6 = float(recent["gf"].mean()) if len(recent) else 0.0
            delta = m6 - float(out["Media gol stagione"])
            out["Media gol ultime 6"] = m6
            out["Delta (ult6 - stag)"] = delta

            if delta >= 0:
                out["Evento estremo (ult6)"] = "3+"
                out["Estremi (ult6)"] = int((recent["gf"] >= 3).sum())
            else:
                out["Evento estremo (ult6)"] = "0"
                out["Estremi (ult6)"] = int((recent["gf"] == 0).sum())

            if len(recent) < 6:
                out["Stato"] = "DATI INSUFFICIENTI"
                return out

            abs_delta = abs(delta)
            extremes = out["Estremi (ult6)"]
            if abs_delta >= 0.7 and extremes >= 3:
                out["Stato"] = "CAMBIO CONFERMATO"
            elif abs_delta >= 0.4:
                out["Stato"] = "WARNING"
            else:
                out["Stato"] = "NORMAL"
            return out

        home_tr = _trend_metrics(hs, home_name)
        away_tr = _trend_metrics(aw, away_name)

        # Gol subiti coerenti (ultime 10)
        home_home = hs[hs["home_id"] == home_id].copy()
        home_home["ga"] = home_home.apply(lambda r: goals_conceded_in_match(r, home_id), axis=1)
        home_home = home_home.dropna(subset=["ga"]).sort_values("utcDate", ascending=False).head(N_LAST)
        home_home["bucket_ga"] = home_home["ga"].astype(int).apply(bucket_0_4p)

        away_away = aw[aw["away_id"] == away_id].copy()
        away_away["ga"] = away_away.apply(lambda r: goals_conceded_in_match(r, away_id), axis=1)
        away_away = away_away.dropna(subset=["ga"]).sort_values("utcDate", ascending=False).head(N_LAST)
        away_away["bucket_ga"] = away_away["ga"].astype(int).apply(bucket_0_4p)


    dashboard_slot = st.container()

    with st.expander("Mostra analisi completa", expanded=False):
        st.subheader("Indicatori trend (automatici) – gol fatti: ultime 6 vs stagione")

        def _badge(s: str) -> str:
            if s == "CAMBIO CONFERMATO":
                return "🔴 CAMBIO CONFERMATO"
            if s == "WARNING":
                return "🟡 WARNING"
            if s == "NORMAL":
                return "🟢 NORMAL"
            return "⚪ DATI INSUFFICIENTI"

        trend_df = pd.DataFrame([home_tr, away_tr])
        trend_df["Stato"] = trend_df["Stato"].apply(_badge)

        st.dataframe(
            trend_df[[
                "Squadra",
                "Match stagione (FINISHED)",
                "Media gol stagione",
                "Match usati ultime 6",
                "Media gol ultime 6",
                "Delta (ult6 - stag)",
                "Evento estremo (ult6)",
                "Estremi (ult6)",
                "Stato",
            ]].style.format({
                "Media gol stagione": "{:.2f}",
                "Media gol ultime 6": "{:.2f}",
                "Delta (ult6 - stag)": "{:+.2f}",
            }),
            use_container_width=True,
            hide_index=True
        )
        st.caption("Regole: 🔴 CAMBIO CONFERMATO se |Δ| ≥ 0.7 e l’evento estremo (3+ se Δ≥0, altrimenti 0) esce ≥ 3 volte nelle ultime 6. 🟡 WARNING se |Δ| ≥ 0.4.")
        st.divider()

        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"### {home_name} – Gol fatti (Totale vs Casa)")
            df_h_cmp = dist_compare_context(hs, hs_home_gf, "Casa")
            st.dataframe(
                df_h_cmp.style.format({"Totale": _fmt_pct_safe, "Casa": _fmt_pct_safe}),
                use_container_width=True,
                hide_index=True
            )
            st.caption(f"Match usati: Totale={len(hs)} | Casa={len(hs_home_gf)}")

        with c2:
            st.markdown(f"### {away_name} – Gol subiti in trasferta (ultime {min(N_LAST, len(away_away))} partite)")
            st.dataframe(dist_table(away_away["bucket_ga"].value_counts(), len(away_away)).style.format({"Percent": _fmt_pct_safe}),
                         use_container_width=True, hide_index=True)

        st.divider()

        c3, c4 = st.columns(2)
        with c3:
            st.markdown(f"### {away_name} – Gol fatti (Totale vs Trasferta)")
            df_a_cmp = dist_compare_context(aw, aw_away_gf, "Trasferta")
            st.dataframe(
                df_a_cmp.style.format({"Totale": _fmt_pct_safe, "Trasferta": _fmt_pct_safe}),
                use_container_width=True,
                hide_index=True
            )
            st.caption(f"Match usati: Totale={len(aw)} | Trasferta={len(aw_away_gf)}")

        with c4:
            st.markdown(f"### {home_name} – Gol subiti in casa (ultime {min(N_LAST, len(home_home))} partite)")
            st.dataframe(dist_table(home_home["bucket_ga"].value_counts(), len(home_home)).style.format({"Percent": _fmt_pct_safe}),
                         use_container_width=True, hide_index=True)



        st.divider()

        # ===========================
        # CHECKLIST WIREFRAME (NO H2H)
        # ===========================

        st.subheader("Checklist guidata (wireframe)")

        # Scelta squadra da valutare (multigol squadra specifica)
        team_choice = st.radio("Squadra da valutare (mercato squadra specifica)", ["Casa", "Trasferta"], horizontal=True)

        # Costruisci distribuzioni % per gol fatti della squadra scelta
        def _pct_dict_from_buckets(series_buckets: pd.Series) -> dict:
            order = ["G0","G1","G2","G3","G4+"]
            total = int(series_buckets.sum()) if series_buckets is not None else 0
            out = {k: 0.0 for k in order}
            if total <= 0:
                return out
            for k in order:
                out[k] = float(series_buckets.get(k, 0)) / total
            return out

        if team_choice == "Casa":
            team_name = home_name
            opp_name = away_name
            team_dist = _pct_dict_from_buckets(hs["bucket_gf"].value_counts())
            opp_conc = _pct_dict_from_buckets(away_away["bucket_ga"].value_counts())
            trend_row = home_tr
        else:
            team_name = away_name
            opp_name = home_name
            team_dist = _pct_dict_from_buckets(aw["bucket_gf"].value_counts())
            opp_conc = _pct_dict_from_buckets(home_home["bucket_ga"].value_counts())
            trend_row = away_tr

        # Distribuzioni CONTEXT (gol fatti casa/trasferta coerenti col match) per mercati di PARTITA (Under/BTTS)
        # Fallback su stagione totale se split insufficiente.
        CTX_MIN_MATCHES = 6

        def _ctx_or_total(split_df: pd.DataFrame, total_df: pd.DataFrame) -> pd.DataFrame:
            return split_df if split_df is not None and len(split_df) >= CTX_MIN_MATCHES else total_df

        if team_choice == "Casa":
            team_for_ctx_df = _ctx_or_total(hs_home_gf, hs)          # gol fatti Inter in casa
            opp_for_ctx_df  = _ctx_or_total(aw_away_gf, aw)          # gol fatti avversaria in trasferta
            team_conc_ctx   = _pct_dict_from_buckets(home_home["bucket_ga"].value_counts())   # subiti casa (ult N)
            opp_conc_ctx    = _pct_dict_from_buckets(away_away["bucket_ga"].value_counts())   # subiti trasferta (ult N)
        else:
            team_for_ctx_df = _ctx_or_total(aw_away_gf, aw)          # gol fatti squadra in trasferta
            opp_for_ctx_df  = _ctx_or_total(hs_home_gf, hs)          # gol fatti avversaria in casa
            team_conc_ctx   = _pct_dict_from_buckets(away_away["bucket_ga"].value_counts())   # subiti trasferta (ult N)
            opp_conc_ctx    = _pct_dict_from_buckets(home_home["bucket_ga"].value_counts())   # subiti casa (ult N)

        team_for_ctx = _pct_dict_from_buckets(team_for_ctx_df["bucket_gf"].value_counts())
        opp_for_ctx  = _pct_dict_from_buckets(opp_for_ctx_df["bucket_gf"].value_counts())


        # --- STEP 0 inputs ---
        c0a, c0b = st.columns([2,1])
        with c0a:
            quota = st.number_input("Quota (facoltativa, per filtro Step 0)", min_value=1.01, max_value=100.0, value=1.62, step=0.01)
        with c0b:
            big_match = st.checkbox("Big match caotico", value=False)
        motivazioni = st.checkbox("Motivazioni anomale (derby/coppa/ultima giornata)", value=False)

        def _is_flat(d: dict) -> bool:
            # distribuzione "piatta" se nessuna coppia adiacente supera 55% e max < 35%
            vals = [d.get("G0",0),d.get("G1",0),d.get("G2",0),d.get("G3",0),d.get("G4+",0)]
            mx = max(vals) if vals else 0
            pairs = [d.get("G0",0)+d.get("G1",0), d.get("G1",0)+d.get("G2",0), d.get("G2",0)+d.get("G3",0)]
            return (mx < 0.35) and all(p < 0.55 for p in pairs)

        # Rendering stabile (NO f-strings multiline)
        def render_result(r: dict):
            text = (
                r.get("badge","")
                + " **"
                + r.get("name","")
                + " — "
                + r.get("label","")
                + "**\n"
                + (f"Prob (Poisson): **{r.get('prob', 0.0):.0%}**\n" if r.get("prob") is not None else "")
                + "Perde se: "
                + r.get("lose_if","")
            )
            st.markdown(text)


        def _pois_p0(lam: float) -> float:
            lam = max(float(lam), 0.0)
            return math.exp(-lam)

        def _pois_pmf(lam: float, k: int) -> float:
            lam = max(float(lam), 0.0)
            if k < 0:
                return 0.0
            # iterativa per stabilità
            p0 = math.exp(-lam)
            if k == 0:
                return p0
            p = p0
            for i in range(1, k+1):
                p *= lam / i
            return p

        def _pois_cdf(lam: float, k: int) -> float:
            # P(X <= k)
            if k < 0:
                return 0.0
            s = 0.0
            for i in range(0, k+1):
                s += _pois_pmf(lam, i)
            return min(max(s, 0.0), 1.0)

        def _pois_range_prob(lam: float, lo: int, hi: int, hi_is_4plus: bool=False) -> float:
            lam = max(float(lam), 0.0)
            lo = int(lo); hi = int(hi)
            if lo > hi:
                return 0.0
            if hi_is_4plus and hi == 4:
                # include 4+ tail
                p_le_3 = _pois_cdf(lam, 3)
                p_lo_3 = _pois_cdf(lam, 3) - _pois_cdf(lam, lo-1)
                # range lo..3 + tail (>=4)
                return (p_lo_3 + (1.0 - p_le_3))
            else:
                return _pois_cdf(lam, hi) - _pois_cdf(lam, lo-1)

        def _safe_mean(s: pd.Series) -> float:
            try:
                s = pd.to_numeric(s, errors="coerce").dropna()
                return float(s.mean()) if len(s) else 0.0
            except Exception:
                return 0.0


        def step_box(title: str, rows: list, status: str, kind: str="info"):
            # kind: info/success/warning/error
            body = "\n".join(rows)
            box = f"### {title}\n{body}\n\n**STATUS: {status}**"
            if kind == "success":
                st.success(box)
            elif kind == "warning":
                st.warning(box)
            elif kind == "error":
                st.error(box)
            else:
                st.info(box)

        # ---------- STEP 0 ----------
        pref_rows = []
        pref_rows.append(f"Quota ≥ 1.62: {'✔️' if quota >= 1.62 else '❌'} (quota={quota:.2f})")
        pref_rows.append(f"Big match caotico: {'❌' if big_match else '✔️ NO'}")
        pref_rows.append(f"Motivazioni anomale: {'❌' if motivazioni else '✔️ NO'}")
        pref_rows.append(f"Distribuzione non piatta: {'✔️' if not _is_flat(team_dist) else '❌'}")

        step0_ok = (quota >= 1.62) and (not big_match) and (not motivazioni) and (not _is_flat(team_dist))
        step_box("STEP 0 — PRE-FILTRO", pref_rows, "OK" if step0_ok else "NO BET", "success" if step0_ok else "error")

        if not step0_ok:
            st.stop()

        # ---------- STEP 1 ----------
        g01 = team_dist["G0"] + team_dist["G1"]
        g12 = team_dist["G1"] + team_dist["G2"]
        g23 = team_dist["G2"] + team_dist["G3"]

        bar = None
        if g01 >= 0.55:
            bar = "BASSO"
        elif g12 >= 0.55:
            bar = "MEDIO"
        elif g23 >= 0.55:
            bar = "ALTO"

        step1_rows = [
            f"G0+G1 = {g01:.0%}",
            f"G1+G2 = {g12:.0%}",
            f"G2+G3 = {g23:.0%}",
            f"➜ BARICENTRO: {bar if bar else 'NESSUNO'}",
        ]
        step1_ok = bar is not None
        step_box("STEP 1 — BARICENTRO (gol fatti)", step1_rows, "OK" if step1_ok else "NO BET", "success" if step1_ok else "error")
        if not step1_ok:
            st.stop()

        # ---------- STEP 2 ----------

        def compute_bar(dist: dict):
            g01 = dist["G0"] + dist["G1"]
            g12 = dist["G1"] + dist["G2"]
            g23 = dist["G2"] + dist["G3"]
            if g01 >= 0.55:
                return "BASSO", g01, g12, g23
            if g12 >= 0.55:
                return "MEDIO", g01, g12, g23
            if g23 >= 0.55:
                return "ALTO", g01, g12, g23
            return None, g01, g12, g23


        def apply_global_coherence(results: list, *, bar_team_ctx: str, bar_opp_ctx: str, lambda_total: float,
                                  team_for_ctx: dict, opp_for_ctx: dict, team_conc_ctx: dict, opp_conc_ctx: dict,
                                  trend_delta: float) -> tuple[list, list]:
            """
            Filtra esiti incoerenti tra loro (motore di coerenza globale).
            Restituisce (results_filtrati, note_filtri)
            """
            notes = []
            if not results:
                return results, notes

            # segnali match-level
            p0_team = float(team_for_ctx.get("G0", 0.0))
            p0_opp  = float(opp_for_ctx.get("G0", 0.0))
            p0_avg_for = (p0_team + p0_opp) / 2.0

            # match "chiuso" se lo 0 è strutturale o lambda bassa
            low_match = (bar_team_ctx == "BASSO") or (bar_opp_ctx == "BASSO") or (p0_avg_for >= 0.28) or (lambda_total <= 2.20)
            # match "aperto" se lambda alta e lo 0 basso (evita falsi positivi)
            high_match = (lambda_total >= 2.90) and (p0_team <= 0.22) and (p0_opp <= 0.22)

            # trend blocker: se trend offensivo forte evita UNDER, se trend in calo evita OVER
            if trend_delta >= 0.7:
                notes.append("Trend offensivo forte → blocco Under / mercati 'bassi'")
            if trend_delta <= -0.7:
                notes.append("Trend in calo → blocco Over / mercati 'alti'")

            filtered = []
            for r in results:
                name = (r.get("name") or "").upper()
                kind = (r.get("kind") or "").upper()

                # 1) BTTS SI incoerente se match è chiuso (lo 0 è strutturale)
                if ("BTTS" in kind) and ("SI" in name) and low_match:
                    notes.append("Rimosso BTTS SI: match con 0 strutturale (coerenza MG/BTTS)")
                    continue

                # 2) BTTS NO incoerente se match è molto aperto
                if ("BTTS" in kind) and ("NO" in name) and high_match:
                    notes.append("Rimosso BTTS NO: match molto aperto (coerenza BTTS/goal expectation)")
                    continue

                # 3) Under 2.5 incoerente se match è aperto o trend offensivo forte
                if kind == "UNDER" and ("2.5" in name) and (high_match or trend_delta >= 0.7):
                    notes.append("Rimosso Under 2.5: incoerente con match aperto/trend offensivo")
                    continue

                # 4) Over incoerente se match è chiuso o trend in calo
                if kind == "OVER" and (low_match or trend_delta <= -0.7):
                    notes.append("Rimosso Over: incoerente con match chiuso/trend in calo")
                    continue

                filtered.append(r)

            # 5) Evita di mostrare contemporaneamente Under e Over: tieni quello più robusto/probabile
            has_under = [r for r in filtered if (r.get("kind","").upper() == "UNDER")]
            has_over  = [r for r in filtered if (r.get("kind","").upper() == "OVER")]
            if has_under and has_over:
                # tieni il migliore per probabilità (e a parità, quello più prudente: Under)
                best_under = max(has_under, key=lambda x: float(x.get("prob", 0.0)))
                best_over  = max(has_over,  key=lambda x: float(x.get("prob", 0.0)))
                if float(best_under.get("prob",0.0)) >= float(best_over.get("prob",0.0)):
                    filtered = [r for r in filtered if r is best_under or (r.get("kind","").upper() != "OVER")]
                    notes.append("Rimosso Over: conflitto Under/Over, tenuto l'esito più coerente/probabile")
                else:
                    filtered = [r for r in filtered if r is best_over or (r.get("kind","").upper() != "UNDER")]
                    notes.append("Rimosso Under: conflitto Under/Over, tenuto l'esito più coerente/probabile")

            # 6) Non mostrare BTTS SI e BTTS NO insieme: tieni il più probabile
            btts = [r for r in filtered if (r.get("kind","").upper() == "BTTS")]
            if len(btts) >= 2:
                best = max(btts, key=lambda x: float(x.get("prob", 0.0)))
                filtered = [r for r in filtered if (r.get("kind","").upper() != "BTTS") or (r is best)]
                notes.append("Rimosso BTTS duplicato: tenuto solo l'esito BTTS più probabile/coerente")

            return filtered, notes

        def ranges_for_bar(bar: str):
            if bar == "BASSO":
                return ["0–1", "0–2"]
            if bar == "MEDIO":
                return ["1–2", "1–3"]
            if bar == "ALTO":
                return ["2–3", "2–4"]
            return []

        ranges = []
        if bar == "BASSO":
            ranges = ["0–1", "0–2"]
        elif bar == "MEDIO":
            ranges = ["1–2", "1–3"]
        else:
            ranges = ["2–3", "2–4"]

        step_box("STEP 2 — RANGE MULTIGOL CANDIDATI", [f"Range candidati: {', '.join(ranges)}"], "OK", "info")

        # ---------- STEP 2B (coerenza split casa/trasferta) ----------
        SPLIT_MIN_MATCHES = 6

        # Baseline (stagione totale) già calcolata: bar / ranges
        bar_base = bar
        ranges_base = ranges

        # Split coerente col match per la squadra valutata
        if team_choice == "Casa":
            split_src = hs_home_gf
        else:
            split_src = aw_away_gf

        split_dist = _pct_dict_from_buckets(split_src["bucket_gf"].value_counts()) if len(split_src) else {k: 0.0 for k in ["G0","G1","G2","G3","G4+"]}
        bar_split, s_g01, s_g12, s_g23 = compute_bar(split_dist)
        ranges_split = ranges_for_bar(bar_split) if bar_split else []

        st.subheader("STEP 2B — Verifica coerenza con split casa/trasferta (gol fatti)")
        msg1 = f"Baseline (stagione): baricentro = **{bar_base}** | range = **{', '.join(ranges_base)}**"
        msg2 = f"Split coerente (n={len(split_src)}): baricentro = **{bar_split if bar_split else 'NESSUNO'}** | range = **{', '.join(ranges_split) if ranges_split else '—'}**"
        st.markdown(msg1 + "\n\n" + msg2)

        split_ok = True
        if len(split_src) < SPLIT_MIN_MATCHES or bar_split is None:
            st.info(f"Split con pochi match (min {SPLIT_MIN_MATCHES}) o baricentro non determinabile: lo uso solo come indicatore (non blocca).")
        else:
            if bar_split == bar_base:
                st.success("✅ Coerente: lo split conferma il baricentro.")
            else:
                split_ok = False
                st.warning("⚠️ Non coerente: lo split sposta il baricentro → MG più rischioso. Consiglio: scegliere range più largo o ridurre stake.")


        # ---------- STEP 3 ----------
        g0c = opp_conc["G0"]
        g3pc = opp_conc["G3"] + opp_conc["G4+"]

        push_low = g0c >= 0.30
        push_high = g3pc >= 0.20

        step3_rows = [
            f"G0 subiti {opp_name} (ultime {min(N_LAST, len(away_away) if team_choice=='Casa' else len(home_home))}): {g0c:.0%} → {'↓ basso' if push_low else '—'}",
            f"G3+ subiti {opp_name}: {g3pc:.0%} → {'↑ alto' if push_high else '—'}",
        ]
        conflict = push_low and push_high
        if conflict:
            step3_rows.append("⚠️ Conflitto G0 alto + G3+ alto → MULTIGOL INSTABILE (passa a Under/BTTS)")
        step_box("STEP 3 — GOL SUBITI AVVERSARI", step3_rows, "CONFLITTO" if conflict else "OK", "warning" if conflict else "info")

        chosen_range = None
        if not conflict:
            # elimina un range
            if push_low:
                chosen_range = ranges[0]  # quello più basso
            elif push_high:
                chosen_range = ranges[1]  # quello più alto
            else:
                # nessuna spinta: scegli quello più centrale (di solito il primo)
                chosen_range = ranges[0]


        # ---------- POISSON: stima lambda (gol attesi) ----------
        # Ricava medie contestuali (casa/trasferta) da stagione corrente. Se campione troppo piccolo, fallback alla media stagione.
        MIN_LAMBDA_MATCHES = 4

        # Home team (casa)
        gf_home = _safe_mean(hs_home_gf["gf"]) if len(hs_home_gf) >= MIN_LAMBDA_MATCHES else _safe_mean(hs["gf"])
        ga_home = _safe_mean(hs_home_ga["ga"]) if "hs_home_ga" in locals() and len(hs_home_ga) >= MIN_LAMBDA_MATCHES else _safe_mean(hs["ga"])

        # Away team (trasferta)
        gf_away = _safe_mean(aw_away_gf["gf"]) if len(aw_away_gf) >= MIN_LAMBDA_MATCHES else _safe_mean(aw["gf"])
        ga_away = _safe_mean(aw_away_ga["ga"]) if "aw_away_ga" in locals() and len(aw_away_ga) >= MIN_LAMBDA_MATCHES else _safe_mean(aw["ga"])

        # Lambda match (media tra attacco e difesa avversaria)
        lambda_home = (gf_home + ga_away) / 2.0
        lambda_away = (gf_away + ga_home) / 2.0
        lambda_total = lambda_home + lambda_away

        # Lambda squadra nel contesto della valutazione (serve per prob MG squadra)
        lambda_team_ctx = lambda_home if team_choice == "Casa" else lambda_away

        # Mostra input Poisson (medie e lambda) se richiesto
        if show_poisson_debug:
            # flag: se sono stati usati i dati contestuali o fallback sul totale stagione
            used_ctx_h_gf = len(hs_home_gf) >= MIN_LAMBDA_MATCHES
            used_ctx_h_ga = ("hs_home_ga" in locals()) and (len(hs_home_ga) >= MIN_LAMBDA_MATCHES)
            used_ctx_a_gf = len(aw_away_gf) >= MIN_LAMBDA_MATCHES
            used_ctx_a_ga = ("aw_away_ga" in locals()) and (len(aw_away_ga) >= MIN_LAMBDA_MATCHES)

            with st.expander("Dettagli Poisson (medie e λ)", expanded=True):
                st.markdown("**Medie gol usate (contesto casa/trasferta):**")
                st.write({
                    "GF_home (squadra casa)": round(gf_home, 3),
                    "GA_home (squadra casa)": round(ga_home, 3),
                    "GF_away (squadra trasferta)": round(gf_away, 3),
                    "GA_away (squadra trasferta)": round(ga_away, 3),
                })
                st.markdown("**Campioni usati (n match):**")
                st.write({
                    "home_gf_n": int(len(hs_home_gf)),
                    "home_ga_n": int(len(hs_home_ga)) if "hs_home_ga" in locals() else 0,
                    "away_gf_n": int(len(aw_away_gf)),
                    "away_ga_n": int(len(aw_away_ga)) if "aw_away_ga" in locals() else 0,
                    "MIN_LAMBDA_MATCHES": int(MIN_LAMBDA_MATCHES),
                })
                st.markdown("**Fallback?** (se False = usato totale stagione)")
                st.write({
                    "GF_home_contesto": bool(used_ctx_h_gf),
                    "GA_home_contesto": bool(used_ctx_h_ga),
                    "GF_away_contesto": bool(used_ctx_a_gf),
                    "GA_away_contesto": bool(used_ctx_a_ga),
                })
                st.markdown("**λ calcolati:**")
                st.write({
                    "lambda_home": round(lambda_home, 3),
                    "lambda_away": round(lambda_away, 3),
                    "lambda_total": round(lambda_total, 3),
                    "lambda_team_ctx": round(lambda_team_ctx, 3),
                })


        # ---------- STEP 4 (MG) ----------
        def range_includes(range_str: str, k: str) -> bool:
            # k in ["G0","G1","G2","G3","G4+"]
            lo, hi = range_str.split("–")
            lo_i = int(lo)
            hi_i = int(hi)
            if k == "G4+":
                v = 4
            else:
                v = int(k[1])
            return lo_i <= v <= hi_i

        def mg_cover(range_str: str, distd: dict) -> float:
            return sum(distd[k] for k in ["G0","G1","G2","G3","G4+"] if range_includes(range_str, k))

        def excluded_strong_events(range_str: str, distd: dict, thr: float=0.30):
            out = []
            for k,p in distd.items():
                if (not range_includes(range_str, k)) and p >= thr:
                    out.append((k,p))
            return out

        mg_results = []
        if not conflict and chosen_range:
            exc = excluded_strong_events(chosen_range, team_dist, 0.30)
            cover = mg_cover(chosen_range, team_dist)
            step4_rows = [f"Multigol candidato: {chosen_range}", f"Copertura (solo gol fatti {team_name}): {cover:.0%}"]
            if exc:
                step4_rows.append("❌ Esclude evento ≥30%: " + ", ".join([f"{k}({p:.0%})" for k,p in exc]))
                mg_ok = False
            else:
                mg_ok = True
                step4_rows.append("✔️ Non esclude eventi ≥30%")
            step_box("STEP 4 — CONTROLLO ESTREMI (MULTIGOL)", step4_rows, "VALIDO" if mg_ok else "SCARTATO", "success" if mg_ok else "error")
            if mg_ok:
                # label robustezza: 1 scenario perdita principale (fuori range) = instabile se somma eventi fuori range >=35%
                lose_ev = [(k,p) for k,p in team_dist.items() if not range_includes(chosen_range,k)]
                lose_sum = sum(p for _,p in lose_ev)
                label = "ROBUSTO" if lose_sum < 0.30 else ("NEUTRO" if lose_sum < 0.40 else "INSTABILE")

                # Se lo split (casa/trasferta) è NON coerente e abbiamo abbastanza match, abbasso di 1 livello la robustezza
                if (not split_ok) and (len(split_src) >= SPLIT_MIN_MATCHES):
                    if label == "ROBUSTO":
                        label = "NEUTRO"
                    elif label == "NEUTRO":
                        label = "INSTABILE"

                badge = "🟢" if label=="ROBUSTO" else ("🟡" if label=="NEUTRO" else "🔴")

                # Se è INSTABILE non lo proponiamo tra le scelte finali
                if label != "INSTABILE":
                    mg_results.append({
                        "name": f"MG {chosen_range} {team_name}",
                        "label": label,
                        "badge": badge,
                        "lose_if": "Gol fuori range: " + ", ".join([k for k,p in lose_ev if p >= 0.10]) if lose_ev else "—",
                        "kind": "MG",
                        "cover": cover,
                        "prob": _pois_range_prob(lambda_team_ctx, int(chosen_range.split("–")[0]), int(chosen_range.split("–")[1]), hi_is_4plus=True),
                    })
        else:
            st.info("Multigol non valutato (conflitto Step 3) → passo ai mercati alternativi.")


        # ---------- STEP 4B (Under/Over/BTTS) ----------
        alt_results = []

        # Under decision (MATCH): proponilo solo se lo scenario è coerente per ENTRAMBE le squadre.
        # Proxy: code 3+ basse sia per gol FATTI che per gol SUBITI (in contesto coerente).
        team_for_g3p = team_for_ctx["G3"] + team_for_ctx["G4+"]
        opp_for_g3p  = opp_for_ctx["G3"] + opp_for_ctx["G4+"]
        team_conc_g3p = team_conc_ctx["G3"] + team_conc_ctx["G4+"]
        opp_conc_g3p  = opp_conc_ctx["G3"] + opp_conc_ctx["G4+"]

        under_ok = (team_for_g3p < 0.25) and (opp_for_g3p < 0.25) and (team_conc_g3p < 0.20) and (opp_conc_g3p < 0.20)

        if under_ok:
            # scelta 2.5 vs 3.5: più "2.5" se la massa è su 0-1 (fatti+subiti)
            low_mass = (team_for_ctx["G0"] + team_for_ctx["G1"] + opp_for_ctx["G0"] + opp_for_ctx["G1"] +
                        team_conc_ctx["G0"] + team_conc_ctx["G1"] + opp_conc_ctx["G0"] + opp_conc_ctx["G1"]) / 4.0
            under_choice = "Under 2.5" if low_mass >= 1.05 else "Under 3.5"  # 1.05 ~ media 0-1 >= 52.5%
            alt_results.append({
                "name": under_choice,
                "label": "ROBUSTO",
                "badge": "🟢",
                "lose_if": "3+ gol totali" if under_choice == "Under 2.5" else "4+ gol totali",
                "kind": "UNDER",
                "prob": (_pois_cdf(lambda_total, 2) if under_choice=="Under 2.5" else _pois_cdf(lambda_total, 3)),
            })
        # BTTS (MATCH) calibrato automaticamente.
        # Stima P(team segna) e P(opp segna) combinando: G0 fatti (contesto) + G0 subiti avversario (contesto).
        p_team_scores = 1.0 - ((team_for_ctx["G0"] + opp_conc_ctx["G0"]) / 2.0)
        p_opp_scores  = 1.0 - ((opp_for_ctx["G0"] + team_conc_ctx["G0"]) / 2.0)
        p_btts_yes = p_team_scores * p_opp_scores  # indipendenza (proxy)
        # Poisson: P(team>=1)=1-e^-lambda, P(opp>=1)=1-e^-lambda
        p_home_scores_pois = 1.0 - _pois_p0(lambda_home)
        p_away_scores_pois = 1.0 - _pois_p0(lambda_away)
        p_btts_yes_pois = p_home_scores_pois * p_away_scores_pois
        p_btts_no_pois = 1.0 - p_btts_yes_pois


        # ---- Coerenza BTTS (ragionamento d'insieme) ----
        # Se una (o entrambe) le squadre ha baricentro "BASSO" sui gol fatti (cioè range candidato 0–1/0–2),
        # allora lo 0 è strutturalmente presente → BTTS SI diventa incoerente come scelta primaria.
        bar_team_ctx, _, _, _ = compute_bar(team_for_ctx)
        bar_opp_ctx,  _, _, _ = compute_bar(opp_for_ctx)

        btts_yes_allowed = (
            (bar_team_ctx != "BASSO") and (bar_opp_ctx != "BASSO") and
            (team_for_ctx["G0"] <= 0.30) and (opp_for_ctx["G0"] <= 0.30) and
            (opp_conc_ctx["G0"] <= 0.35) and (team_conc_ctx["G0"] <= 0.35)
        )

        # Proposta BTTS YES/NO solo se segnale netto + coerenza
        if btts_yes_allowed and (p_btts_yes >= 0.55) and (p_team_scores >= 0.65) and (p_opp_scores >= 0.65):
            alt_results.append({
                "name": "BTTS SI",
                "label": "NEUTRO",
                "badge": "🟡",
                "lose_if": "Almeno una a 0 gol",
                "kind": "BTTS",
                "prob": p_btts_yes_pois,
            })
        else:
            # Se BTTS SI non è coerente (o segnale non netto), e lo 0 è rilevante, preferisci BTTS NO quando il Poisson lo supporta.
            if (p_btts_no_pois >= 0.55) and ((bar_team_ctx == "BASSO") or (bar_opp_ctx == "BASSO")):
                alt_results.append({
                    "name": "BTTS NO",
                    "label": "NEUTRO",
                    "badge": "🟡",
                    "lose_if": "Entrambe segnano",
                    "kind": "BTTS",
                    "prob": p_btts_no_pois,
                })
            elif (p_btts_yes <= 0.45):
                alt_results.append({
                    "name": "BTTS NO",
                    "label": "NEUTRO",
                    "badge": "🟡",
                    "lose_if": "Entrambe segnano",
                    "kind": "BTTS",
                    "prob": p_btts_no_pois,
                })
        # Over proxy: G2+ team e 2+ concessi opp
        team_g2p = team_dist["G2"] + team_dist["G3"] + team_dist["G4+"]
        team_g3p = team_dist["G3"] + team_dist["G4+"]
        opp_2p_conc = opp_conc["G2"] + opp_conc["G3"] + opp_conc["G4+"]
        over_ok = (team_g2p >= 0.55) and (opp_2p_conc >= 0.45) and (team_dist["G0"] < 0.25)
        if over_ok:
            over_choice = "Over 2.5" if team_g3p >= 0.25 else "Over 1.5 squadra"
            alt_results.append({"name":over_choice, "label":"NEUTRO", "badge":"🟡", "lose_if":"0-1 gol", "kind":"OVER", "prob": (1.0 - _pois_cdf(lambda_total, 2) if over_choice=="Over 2.5" else (1.0 - _pois_p0(lambda_team_ctx)))})

        # ---------- STEP 5 (Trend blocker) ----------
        delta = float(trend_row.get("Delta (ult6 - stag)", 0.0))
        trend_rows = [f"Delta (ult6 - stag) {team_name}: {delta:+.2f}"]
        trend_block = False
        if delta >= 0.7:
            trend_rows.append("⚠️ Trend offensivo forte: evita Under / MG bassi")
        elif delta <= -0.7:
            trend_rows.append("⚠️ Trend in calo: evita Over / MG alti")
        else:
            trend_rows.append("✔️ Trend neutro: nessun blocco")
        step_box("STEP 5 — TREND (ultime 6)", trend_rows, "BLOCCO" if abs(delta)>=0.7 else "OK", "warning" if abs(delta)>=0.7 else "success")

        # ---------- OUTPUT FINALE ----------
        all_results = mg_results + alt_results

        # --- COERENZA GLOBALE (match-level) ---
        all_results, coherence_notes = apply_global_coherence(
            all_results,
            bar_team_ctx=bar_team_ctx,
            bar_opp_ctx=bar_opp_ctx,
            lambda_total=lambda_total,
            team_for_ctx=team_for_ctx,
            opp_for_ctx=opp_for_ctx,
            team_conc_ctx=team_conc_ctx,
            opp_conc_ctx=opp_conc_ctx,
            trend_delta=delta,
        )
        if coherence_notes:
            step_box("COERENZA GLOBALE — filtri applicati", coherence_notes[:8], "OK", "info")

        # ordina per robustezza
        order = {"ROBUSTO": 0, "NEUTRO": 1, "INSTABILE": 2}
        all_results = sorted(all_results, key=lambda x: order.get(x.get("label","NEUTRO"), 1))
        all_results = all_results[:3]

        st.subheader("Esiti coerenti con i dati (ordinati per robustezza)")
        if not all_results:
            st.warning("Nessun esito supera i filtri → NO BET")
        else:
            for r in all_results:
                render_result(r)



        # Scelta sintetica senza filtro quote: migliore esito per robustezza e probabilita.
        # La dashboard finale ricalcola comunque entrambe le squadre insieme.
        picked = all_results[0] if all_results else None
        picked_q = None
        why = "Miglior esito coerente per robustezza e probabilita" if picked else "Nessun esito coerente"


        # =========================
        # LOG SU GOOGLE SHEET (AUTO)
        # =========================
        with st.expander("📌 Salva su Google Sheet (MG STORICO)", expanded=False):
            if not _gs_available():
                st.warning("Google Sheet non configurato. Crea `.streamlit/secrets.toml` con la sezione [gsheets].")
            else:
                quota_presa = st.number_input("Quota presa (reale)", min_value=1.01, max_value=50.0, value=float(picked_q), step=0.01)
                note_log = st.text_input("Note (opzionale)", value="")
                if st.button("✅ Salva questa giocata", type="primary"):
                    try:
                        data_str = dt.datetime.now().strftime("%Y-%m-%d")
                        campionato_str = str(comp_label) if 'comp_label' in globals() else ""
                        match_str = f"{home_name}-{away_name}" if 'home_name' in globals() and 'away_name' in globals() else ""
                        mercato_str = str(picked.get("kind", ""))  # MG / UNDER / BTTS / OVER
                        esito_str = str(picked.get("name", ""))
                        prob_str = float(picked.get("prob", 0.0)) if picked.get("prob") is not None else ""
                        # Formato colonne come logger manuale: data, campionato, match, mercato, esito, prob, quota, note, risultato
                        row = [data_str, campionato_str, match_str, mercato_str, esito_str, prob_str, float(quota_presa), note_log, ""]
                        _gs_append_row(row)
                        st.success("Salvato su MG STORICO ✅")
                    except Exception as e:
                        st.error(f"Errore salvataggio su Google Sheet: {e}")

    # =========================
    # DASHBOARD SINTETICA - STILE RENDER
    # =========================
    with dashboard_slot:
        matchday_val = sel.get("matchday")
        matchday_txt = f"Giornata {matchday_val}" if matchday_val not in (None, "", "null") and not pd.isna(matchday_val) else "Giornata n.d."
        try:
            local_dt = pd.to_datetime(sel.get("utcDate"), utc=True).tz_convert(APP_TIMEZONE)
            date_txt = local_dt.strftime("%a %d %b %Y %H:%M")
        except Exception:
            date_txt = "Data/ora non disponibile"

        # Valuta sempre ENTRAMBE le squadre, indipendentemente dalla radio della parte dettagliata.
        dashboard_eval = analyze_match_for_summary(sel, min_prob=0.0, include_all_candidates=True)
        dashboard_candidates_raw = (dashboard_eval or {}).get("candidates", [])
        dashboard_candidates = []
        _seen_dashboard = set()
        for _r in dashboard_candidates_raw:
            if str(_r.get("kind", "")).upper() == "BTTS":
                continue
            _key = (str(_r.get("name", "")), str(_r.get("kind", "")), str(_r.get("label", "")))
            if _key in _seen_dashboard:
                continue
            _seen_dashboard.add(_key)
            dashboard_candidates.append(_r)

        dashboard_picked = dashboard_candidates[0] if dashboard_candidates else None
        dashboard_alternatives = dashboard_candidates[1:] if len(dashboard_candidates) > 1 else []
        match_details_odds = fetch_match_details(safe_int(sel.get("match_id"), 0))

        season_matches_for_table = fetch_league_season_matches(season_id)
        standings_df = build_standings_from_matches(season_matches_for_table)

        st.markdown(
            f"""<div class="render-topline">
                <div><span class="league">{comp_label}</span><span class="day">{matchday_txt}</span></div>
                <div class="date">🗓 {date_txt}</div>
            </div>""",
            unsafe_allow_html=True,
        )

        def _last5_rows_html(team_df: pd.DataFrame, team_id: int):
            if team_df is None or team_df.empty:
                return '<div style="color:#64748b;font-size:.74rem">Nessun dato</div>'
            recent = team_df[team_df["status"] == "FINISHED"].sort_values("utcDate", ascending=False).head(5)
            out = []
            for _, m in recent.iterrows():
                is_home = safe_int(m.get("home_id"), 0) == int(team_id)
                gf = safe_int(m.get("home_ft" if is_home else "away_ft"), 0)
                ga = safe_int(m.get("away_ft" if is_home else "home_ft"), 0)
                opponent = str(m.get("away_name" if is_home else "home_name") or "-")
                try:
                    d = pd.to_datetime(m.get("utcDate"), utc=True).tz_convert(APP_TIMEZONE).strftime("%d/%m")
                except Exception:
                    d = "-"
                ct = "Casa" if is_home else "Trasf."
                out.append(f'<div class="last5-row"><span class="muted">{d}</span><span class="muted">{ct}</span><span>{opponent}</span><span class="score">{gf}-{ga}</span></div>')
            return "".join(out)

        # TOP: classifica / pick / ultimi 5
        top_l, top_c, top_r = st.columns([1.03, 1.62, 1.03], gap="medium")

        with top_l:
            if standings_df.empty:
                standings_html = '<div style="color:#91a1b6">Classifica non disponibile.</div>'
            else:
                rows_html = []
                for _, rr in standings_df.iterrows():
                    css = ""
                    if rr.get("Squadra") == home_name:
                        css = ' class="sel-home"'
                    elif rr.get("Squadra") == away_name:
                        css = ' class="sel-away"'
                    rows_html.append(
                        f"<tr{css}><td>{int(rr['Pos'])}</td><td>{rr['Squadra']}</td><td>{int(rr['PG'])}</td><td>{int(rr['GF'])}</td><td>{int(rr['GS'])}</td><td>{int(rr['DR']):+d}</td><td><b>{int(rr['Pt'])}</b></td></tr>"
                    )
                standings_html = f"""<div class="standings-wrap"><table class="render-table">
                <thead><tr><th>Pos</th><th>Squadra</th><th>PG</th><th>GF</th><th>GS</th><th>DR</th><th>Pt</th></tr></thead>
                <tbody>{''.join(rows_html)}</tbody></table></div>"""
            st.markdown(f'<div class="render-card"><div class="render-title">Classifica generale</div>{standings_html}</div>', unsafe_allow_html=True)

        with top_c:
            if dashboard_picked is None:
                pick_html = """<div class="pick-box">
                    <div class="pick-kicker">Giocata consigliata</div>
                    <div class="pick-name">NO BET</div>
                    <div style="color:#91a1b6">Nessun esito coerente disponibile</div>
                </div>"""
            else:
                prob_txt = f"{float(dashboard_picked.get('prob',0.0)):.0%}"
                rob_txt = str(dashboard_picked.get("label","-"))
                picked_odd = odds_for_candidate(dashboard_picked, match_details_odds)
                odd_txt = f"{picked_odd:.2f}" if picked_odd is not None else "-"
                alt_html = []
                for r in dashboard_alternatives:
                    _odd = odds_for_candidate(r, match_details_odds)
                    _odd_txt = f"{_odd:.2f}" if _odd is not None else "-"
                    _prob = f"{float(r.get('prob',0.0)):.0%}" if r.get("prob") is not None else "-"
                    alt_html.append(f'<span class="alt-pill">{r.get("name","-")} &nbsp; {_prob} &nbsp; q. {_odd_txt}</span>')
                alt_section = (
                    '<div class="alt-label">Alternative coerenti</div><div class="alt-wrap">' + "".join(alt_html) + "</div>"
                    if alt_html else '<div class="alt-label">Nessuna alternativa coerente</div>'
                )
                pick_html = f"""<div class="pick-box">
                    <div class="pick-kicker">Giocata consigliata</div>
                    <div class="pick-name">{dashboard_picked.get('name','-')}</div>
                    <div class="pick-stats">
                        <div class="pick-stat"><div class="pick-stat-label">Probabilità</div><div class="pick-stat-val green">{prob_txt}</div></div>
                        <div class="pick-stat"><div class="pick-stat-label">Quota</div><div class="pick-stat-val orange">{odd_txt}</div></div>
                        <div class="pick-stat"><div class="pick-stat-label">Robustezza</div><div class="pick-stat-val green">{rob_txt}</div></div>
                    </div>
                    {alt_section}
                </div>"""
            st.markdown(
                f'<div class="render-card"><div class="match-center">{home_name}<span class="match-vs">VS</span>{away_name}</div>{pick_html}</div>',
                unsafe_allow_html=True,
            )

        with top_r:
            last5_html = (
                f'<div class="last5-team">{home_name}</div>{_last5_rows_html(hs, home_id)}'
                f'<div class="last5-team" style="margin-top:12px">{away_name}</div>{_last5_rows_html(aw, away_id)}'
            )
            st.markdown(f'<div class="render-card"><div class="render-title">Ultimi 5 risultati</div>{last5_html}</div>', unsafe_allow_html=True)

        # DATI SQUADRE
        home_recent_ga = _mean_num(hs.sort_values("utcDate", ascending=False).head(6), "ga")
        away_recent_ga = _mean_num(aw.sort_values("utcDate", ascending=False).head(6), "ga")
        home_prob_score = 1.0 - math.exp(-max(lambda_home, 0.0))
        away_prob_score = 1.0 - math.exp(-max(lambda_away, 0.0))

        def _bucket_pct_from_df(df, col):
            if df is None or df.empty or col not in df.columns:
                return [0,0,0,0,0]
            vals = pd.to_numeric(df[col], errors="coerce").dropna().astype(int)
            n = len(vals)
            if not n:
                return [0,0,0,0,0]
            return [
                (vals == 0).mean(),
                (vals == 1).mean(),
                (vals == 2).mean(),
                (vals == 3).mean(),
                (vals >= 4).mean(),
            ]

        def _dist_html(vals, kind="gf"):
            labels = ["0","1","2","3","4+"]
            cols = []
            for lab, p in zip(labels, vals):
                h = max(4, min(42, round(float(p) * 90)))
                cols.append(
                    f'<div class="dist-col"><div>{lab}</div><div class="dist-bar-wrap"><div class="dist-bar {kind}" style="height:{h}px"></div></div><div class="dist-pct">{p:.0%}</div></div>'
                )
            return '<div class="dist-grid">' + "".join(cols) + "</div>"

        home_gf_dist = _bucket_pct_from_df(hs_home_gf if len(hs_home_gf) else hs, "gf")
        home_ga_dist = _bucket_pct_from_df(hs_home_ga if len(hs_home_ga) else hs, "ga")
        away_gf_dist = _bucket_pct_from_df(aw_away_gf if len(aw_away_gf) else aw, "gf")
        away_ga_dist = _bucket_pct_from_df(aw_away_ga if len(aw_away_ga) else aw, "ga")

        team_l, team_r = st.columns(2, gap="medium")
        with team_l:
            delta_h = safe_float(home_tr.get("Delta (ult6 - stag)"), 0.0)
            trend_h = _trend_html(home_tr.get("Stato"))
            st.markdown(
                f"""<div class="team-panel">
                    <div class="team-head">{home_name} &nbsp; (CASA)</div>
                    <div class="metric-strip"><div class="lbl">⚽ GOL FATTI</div><div><div class="big">{_mean_num(hs,'gf'):.2f}</div><div class="sub">media totale</div></div><div><div class="big">{_mean_num(hs_home_gf,'gf'):.2f}</div><div class="sub">media in casa</div></div></div>
                    <div class="metric-strip"><div class="lbl">🛡 GOL SUBITI</div><div><div class="big">{_mean_num(hs,'ga'):.2f}</div><div class="sub">media totale</div></div><div><div class="big">{_mean_num(hs_home_ga,'ga'):.2f}</div><div class="sub">media in casa</div></div></div>
                    <div class="metric-strip"><div class="lbl">🗓 ULTIME 6</div><div><div class="big">{float(home_tr.get('Media gol ultime 6',0.0)):.2f}</div><div class="sub">gol fatti</div></div><div><div class="big">{home_recent_ga:.2f}</div><div class="sub">gol subiti</div></div></div>
                    <div class="trend-strip"><div><div class="lbl">TREND</div><div style="font-size:.64rem;color:#64748b">{delta_h:+.2f} vs media stagione</div></div><div class="val">{trend_h}</div></div>
                    <div class="trend-strip"><div class="lbl">◎ PROBABILITÀ DI SEGNARE</div><div class="val green">{home_prob_score:.0%}</div></div>
                    <div class="dist-box"><div class="dist-title">Gol fatti - distribuzione (casa)</div>{_dist_html(home_gf_dist,'gf')}</div>
                    <div class="dist-box"><div class="dist-title">Gol subiti - distribuzione (casa)</div>{_dist_html(home_ga_dist,'ga')}</div>
                </div>""",
                unsafe_allow_html=True,
            )
        with team_r:
            delta_a = safe_float(away_tr.get("Delta (ult6 - stag)"), 0.0)
            trend_a = _trend_html(away_tr.get("Stato"))
            st.markdown(
                f"""<div class="team-panel away">
                    <div class="team-head">{away_name} &nbsp; (TRASFERTA)</div>
                    <div class="metric-strip"><div class="lbl">⚽ GOL FATTI</div><div><div class="big">{_mean_num(aw,'gf'):.2f}</div><div class="sub">media totale</div></div><div><div class="big">{_mean_num(aw_away_gf,'gf'):.2f}</div><div class="sub">media in trasferta</div></div></div>
                    <div class="metric-strip"><div class="lbl">🛡 GOL SUBITI</div><div><div class="big">{_mean_num(aw,'ga'):.2f}</div><div class="sub">media totale</div></div><div><div class="big">{_mean_num(aw_away_ga,'ga'):.2f}</div><div class="sub">media in trasferta</div></div></div>
                    <div class="metric-strip"><div class="lbl">🗓 ULTIME 6</div><div><div class="big">{float(away_tr.get('Media gol ultime 6',0.0)):.2f}</div><div class="sub">gol fatti</div></div><div><div class="big">{away_recent_ga:.2f}</div><div class="sub">gol subiti</div></div></div>
                    <div class="trend-strip"><div><div class="lbl">TREND</div><div style="font-size:.64rem;color:#64748b">{delta_a:+.2f} vs media stagione</div></div><div class="val">{trend_a}</div></div>
                    <div class="trend-strip"><div class="lbl">◎ PROBABILITÀ DI SEGNARE</div><div class="val orange">{away_prob_score:.0%}</div></div>
                    <div class="dist-box"><div class="dist-title">Gol fatti - distribuzione (trasferta)</div>{_dist_html(away_gf_dist,'gf')}</div>
                    <div class="dist-box"><div class="dist-title">Gol subiti - distribuzione (trasferta)</div>{_dist_html(away_ga_dist,'ga')}</div>
                </div>""",
                unsafe_allow_html=True,
            )

        st.markdown('<div class="footer-note">Tutta l’analisi tecnica resta nella tendina “Mostra analisi completa”.</div>', unsafe_allow_html=True)

