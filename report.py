#!/usr/bin/env python3
"""
report.py — Full multi-year AIC KPI report + sidebar with Dial Volume,
Connection Rate, Efficiency, and Call Timing Heatmap views.
Run: python3 report.py
"""
import csv, json, webbrowser, datetime, calendar, os, sys
from pathlib import Path
from collections import defaultdict

GAP_MINUTES = 20
DATE_FMTS = [
    "%m/%d/%Y, %I:%M %p",
    "%m/%d/%Y %I:%M %p",
    "%Y-%m-%dT%H:%M:%SZ",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
]

ALL_FILES = [
    # ── Edgar ──────────────────────────────────────────────────────────────
    "/Users/antonibustamante/Downloads/19383-call-history-4dbfe3c8-76e7-4dab-bae7-23145444459e.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-bb01b9e8-1b5c-46db-b612-1fd8543f354b.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-186a8ac2-7cbe-48f9-8ded-d0cd6e613f5b.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-f14cdb3c-c630-49ee-8734-4240c76351a7.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-518df07b-9303-4506-b53e-8468ab237805.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-24af617a-8334-4520-8f8f-274d2a221c23.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-4285e6e9-9b76-45bb-8701-d16a1f603b61.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-a990daf3-e8bc-4b8a-a31e-4614392e056c.csv",
    # ── Leandro ────────────────────────────────────────────────────────────
    "/Users/antonibustamante/Downloads/19383-call-history-c4c484b9-4b45-4331-85bf-af22e9cda02b.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-a792864a-89c6-48a7-b829-27abea9dcd77.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-ab929b50-a5ed-4135-b9f3-acd35cc6ced4.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-278d3735-2123-4592-bcf1-384c22a321db.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-98615102-1d96-444f-b073-c42de5bdd82c.csv",
    "/Users/antonibustamante/Downloads/19383-call-history-e7f3e7d2-3968-4ea3-a18e-7cefa99fec93.csv",
]

AGENT_ORDER   = ["Edgar Morales", "Leandro Greasebook"]
AGENT_DISPLAY = {"Edgar Morales": "Edgar", "Leandro Greasebook": "Leandro"}
AGENT_COLOR   = {"Edgar Morales": "#4F9CF9", "Leandro Greasebook": "#F97B4F"}

# Whitelist: only these dispositions count as a meaningful connection
# "Connection"/"Connect" = 2024/2025 Kixie labels for a live conversation
# "BOOKED"/"CALLBACK" = also meaningful outcomes that required live contact
CONN_DISPS = {"connection", "connect", "booked", "callback"}

# ── Static CSS (plain string — no f-string needed) ────────────────────────────
_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
     background:#0d1117;color:#c9d1d9;display:flex;min-height:100vh}
.sidebar{position:fixed;left:0;top:0;width:200px;height:100vh;
         background:#161b22;border-right:1px solid #30363d;
         padding:20px 10px;overflow-y:auto;z-index:100;
         display:flex;flex-direction:column}
.sb-brand{font-size:.68rem;font-weight:700;text-transform:uppercase;
          letter-spacing:.12em;color:#3fb950;padding:0 12px;margin-bottom:4px}
.sb-title{font-size:.98rem;font-weight:700;color:#fff;padding:0 12px;margin-bottom:22px}
.nav-sep{font-size:.62rem;text-transform:uppercase;letter-spacing:.08em;
         color:#484f58;padding:14px 12px 5px}
.nav-item{display:block;width:100%;padding:9px 12px;margin-bottom:2px;
          border-radius:7px;background:transparent;border:none;
          color:#8b949e;font-size:.82rem;cursor:pointer;text-align:left;
          transition:all .15s;font-family:inherit}
.nav-item:hover{background:#21262d;color:#c9d1d9}
.nav-item.active{background:#388bfd22;color:#58a6ff;font-weight:600}
.main-content{margin-left:210px;padding:32px 24px;width:100%;max-width:1260px}
h1{font-size:1.5rem;color:#fff;font-weight:700;margin-bottom:4px}
.meta{color:#8b949e;font-size:.82rem;margin-bottom:28px}
.cards{display:flex;gap:18px;flex-wrap:wrap;margin-bottom:26px}
.card{flex:1;min-width:240px;background:#161b22;border-radius:10px;
      padding:22px 24px;border:1px solid #30363d}
.card-name{font-size:.75rem;font-weight:700;text-transform:uppercase;
           letter-spacing:.08em;margin-bottom:14px}
.yr-row{display:flex;align-items:baseline;gap:10px;padding:8px 0;
        border-bottom:1px solid #21262d}
.yr-row:last-of-type{border-bottom:none}
.yr-label{font-size:.82rem;color:#8b949e;width:42px;flex-shrink:0}
.yr-stat{font-size:1.5rem;font-weight:700;color:#fff;flex-shrink:0}
.yr-sub{font-size:.75rem;color:#8b949e}
.card-calls{font-size:.75rem;color:#8b949e;margin-top:12px;padding-top:10px;
            border-top:1px solid #21262d}
.section{background:#161b22;border:1px solid #30363d;border-radius:10px;
         padding:24px;margin-bottom:20px}
.stitle{font-size:1rem;font-weight:600;margin-bottom:16px}
.chart-wrap{position:relative;height:260px}
.chart-wrap.tall{height:300px}
table{width:100%;border-collapse:collapse;font-size:.82rem}
th{text-align:left;padding:8px 10px;color:#8b949e;border-bottom:1px solid #21262d;
   font-size:.71rem;text-transform:uppercase;letter-spacing:.05em;font-weight:500}
td{padding:9px 10px;border-bottom:1px solid #21262d;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#1c2128}
.yr-sep td{background:#0d1117 !important;border-bottom:none}
.bar{height:7px;border-radius:4px;min-width:1px}
.db{padding:4px 8px;border-radius:4px;font-size:.82rem;white-space:nowrap;transition:width .3s}
small{color:#8b949e}
.em{color:#3d444d}
.dc{white-space:nowrap;color:#e6edf3;font-size:.82rem}
.wd{color:#8b949e;display:inline-block;width:28px}
.partial{background:#f0a50022;color:#f0a500;border:1px solid #f0a50044;
         border-radius:4px;font-size:.68rem;padding:1px 5px;margin-left:4px}
.filters{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:16px}
.fbt{background:#21262d;border:1px solid #30363d;color:#8b949e;border-radius:20px;
     padding:4px 12px;font-size:.75rem;cursor:pointer;transition:all .15s}
.fbt:hover{background:#2d333b;color:#c9d1d9}
.fbt.active{background:#388bfd22;color:#58a6ff;border-color:#388bfd}
.fyr{font-size:.72rem;color:#484f58;padding:0 4px;align-self:center}
.mbox{background:#0d1117;border:1px solid #30363d;border-radius:10px;
      padding:24px;margin-bottom:20px;line-height:1.75}
.mbox h2{color:#fff;font-size:1rem;margin-bottom:12px}
.mbox h3{color:#c9d1d9;font-size:.88rem;margin:16px 0 6px}
.mbox p,.mbox li{font-size:.83rem;color:#8b949e}
.mbox li{margin-left:20px;margin-bottom:4px}
.pill{display:inline-block;background:#21262d;border:1px solid #30363d;
      border-radius:5px;padding:1px 7px;font-size:.76rem;color:#58a6ff;font-family:monospace}
.ex-block{background:#161b22;border:1px solid #30363d;border-radius:8px;
          padding:14px 16px;margin-top:10px;font-size:.82rem;color:#c9d1d9;line-height:1.9}
.ex-block .arrow{color:#484f58;margin:0 4px}
.insight{margin-top:16px;padding:12px 16px;background:#0d1117;
         border-left:3px solid #388bfd;border-radius:0 6px 6px 0;
         font-size:.82rem;color:#8b949e;line-height:1.65}
.insight strong{color:#c9d1d9}
.insight .flag{color:#f85149}
.insight .good{color:#3fb950}
"""

# ── Static JS — chart logic (plain string, no f-prefix, real braces OK) ───────
_JS = """
Chart.defaults.color = "#8b949e";
Chart.defaults.borderColor = "#21262d";

var _ci = {};

function showPage(id) {
    document.querySelectorAll('.page').forEach(function(p){ p.style.display='none'; });
    document.getElementById('page-'+id).style.display='';
    document.querySelectorAll('.nav-item').forEach(function(n){ n.classList.remove('active'); });
    document.getElementById('nav-'+id).classList.add('active');
    if (!_ci[id]) {
        var fn = window['_init_'+id];
        if (fn) fn();
        _ci[id] = true;
    }
    location.hash = id;
}

function _base_opts(ytitle, ycb) {
    return {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#c9d1d9' } } },
        scales: {
            x: { ticks: { color: '#8b949e', maxRotation: 45 }, grid: { color: '#21262d' } },
            y: { ticks: { color: '#8b949e', callback: ycb || null }, grid: { color: '#21262d' },
                 title: { display: !!ytitle, text: ytitle || '', color: '#8b949e' } }
        }
    };
}

function _dual_opts(y0title, y1title) {
    return {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { labels: { color: '#c9d1d9' } } },
        scales: {
            x: { ticks: { color: '#8b949e', maxRotation: 45 }, grid: { color: '#21262d' } },
            y: { type: 'linear', position: 'left',
                 ticks: { color: '#8b949e' }, grid: { color: '#21262d' },
                 title: { display: true, text: y0title, color: '#8b949e' } },
            y1: { type: 'linear', position: 'right',
                  ticks: { color: '#8b949e', callback: function(v){ return v+'%'; } },
                  grid: { drawOnChartArea: false },
                  title: { display: true, text: y1title, color: '#8b949e' } }
        }
    };
}

function _init_aic() {
    new Chart(document.getElementById('monthlyChart'), {
        type: 'bar',
        data: { labels: MONTHLY_LABELS, datasets: AIC_MONTHLY_DS },
        options: _base_opts('AIC Hours', function(v){ return v+' hrs'; })
    });
    new Chart(document.getElementById('weeklyChart'), {
        type: 'line',
        data: { labels: WEEKLY_LABELS, datasets: AIC_WEEKLY_DS },
        options: _base_opts('AIC Hours', function(v){ return v+' hrs'; })
    });
}

function _init_dials() {
    new Chart(document.getElementById('dialMonthlyChart'), {
        type: 'bar',
        data: { labels: MONTHLY_LABELS, datasets: DIAL_MONTHLY_DS },
        options: _base_opts('Dials')
    });
    new Chart(document.getElementById('dialWeeklyChart'), {
        type: 'line',
        data: { labels: WEEKLY_LABELS, datasets: DIAL_WEEKLY_DS },
        options: _base_opts('Dials')
    });
}

function _init_connect() {
    new Chart(document.getElementById('connRateChart'), {
        type: 'line',
        data: { labels: MONTHLY_LABELS, datasets: CONN_RATE_DS },
        options: _base_opts('Connection %', function(v){ return v+'%'; })
    });
}

function _init_efficiency() {
    new Chart(document.getElementById('effChart'), {
        type: 'bar',
        data: { labels: MONTHLY_LABELS, datasets: EFF_DS },
        options: _base_opts('Dials / AIC Hour')
    });
}

function _init_heatmap() {
    new Chart(document.getElementById('hourChart'), {
        data: { labels: HOUR_LABELS, datasets: HOUR_DS },
        options: _dual_opts('Dials', 'Connection %')
    });
    new Chart(document.getElementById('dowChart'), {
        data: { labels: DOW_LABELS, datasets: DOW_DS },
        options: _dual_opts('Dials', 'Connection %')
    });
}

function filterDays(ym, btn) {
    document.querySelectorAll('.fbt').forEach(function(b){ b.classList.remove('active'); });
    btn.classList.add('active');
    document.querySelectorAll('#dailyBody tr').forEach(function(tr){
        tr.style.display = (ym === 'all' || tr.dataset.ym === ym) ? '' : 'none';
    });
}
"""

# ── Parse helpers ─────────────────────────────────────────────────────────────

def parse_dt(s):
    for fmt in DATE_FMTS:
        try: return datetime.datetime.strptime(s.strip(), fmt)
        except: pass
    return None

def parse_duration(s):
    try:
        parts = s.strip().split(":")
        if len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    except Exception:
        pass
    return 0

def is_connection(disp):
    return disp.strip().lower() in CONN_DISPS

# ── Load ──────────────────────────────────────────────────────────────────────

def load_all(files):
    records = []
    seen = set()
    for path in files:
        if not Path(path).exists():
            print(f"  WARNING: {path} not found, skipping")
            continue
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                if row.get("Type", "").strip().lower() not in ("outgoing", "outbound"):
                    continue
                dt = parse_dt(row.get("Date", ""))
                if not dt:
                    continue
                fn    = row.get("Agent First Name", "").strip()
                ln    = row.get("Agent Last Name",  "").strip()
                agent = f"{fn} {ln}".strip()
                to_num = row.get("To Number", "").strip()
                key = f"{agent}|{dt.isoformat()}|{to_num}"
                if key in seen:
                    continue
                seen.add(key)
                records.append({
                    "dt":           dt,
                    "agent":        agent,
                    "status":       row.get("Status",      "").strip(),
                    "disposition":  row.get("Disposition", "").strip(),
                    "duration_sec": parse_duration(row.get("Duration", "")),
                })
    return records

def load_aic_records(files):
    """Load all calls (inbound + outbound) for AIC computation.
    Each record has is_outbound=True/False so compute_aic() can apply
    the hybrid rule: sessions open only on outbound, inbounds extend them."""
    records = []
    seen = set()
    for path in files:
        if not Path(path).exists():
            continue
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.DictReader(fh):
                dt = parse_dt(row.get("Date", ""))
                if not dt:
                    continue
                fn    = row.get("Agent First Name", "").strip()
                ln    = row.get("Agent Last Name",  "").strip()
                agent = f"{fn} {ln}".strip()
                call_type = row.get("Type", "").strip().lower()
                is_outbound = call_type in ("outgoing", "outbound")
                to_num = row.get("To Number", "").strip()
                from_num = row.get("From Number", "").strip()
                number = to_num if is_outbound else from_num
                key = f"{agent}|{dt.isoformat()}|{number}|{call_type}"
                if key in seen:
                    continue
                seen.add(key)
                records.append({
                    "dt":          dt,
                    "agent":       agent,
                    "is_outbound": is_outbound,
                })
    return records

# ── Airtable data source ──────────────────────────────────────────────────────

AIRTABLE_TABLE = "Kixie Call Log"
AGENT_NAME_MAP = {
    "Edgar":   "Edgar Morales",
    "Leandro": "Leandro Greasebook",
}

def fetch_airtable(api_key, base_id):
    """Fetch all records from the Airtable Kixie Call Log table (paginated)."""
    try:
        import requests as _req
    except ImportError:
        print("ERROR: `requests` not installed. Run: pip install requests")
        sys.exit(1)

    url     = f"https://api.airtable.com/v0/{base_id}/{_req.utils.quote(AIRTABLE_TABLE)}"
    headers = {"Authorization": f"Bearer {api_key}"}
    records, params = [], {"pageSize": 100}
    while True:
        r = _req.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
        params["offset"] = offset
    return records


def transform_airtable_records(raw):
    """Convert Airtable records into the same dicts load_all() / load_aic_records() produce."""
    records, aic_records, seen = [], [], set()
    for rec in raw:
        f  = rec.get("fields", {})
        dt = parse_dt(f.get("call_datetime") or f.get("call_date") or "")
        if not dt:
            continue
        agent_short = f.get("agent", "")
        agent       = AGENT_NAME_MAP.get(agent_short, agent_short)
        direction   = (f.get("direction") or "outgoing").lower()
        is_outbound = direction in ("outgoing", "outbound")
        disposition = f.get("disposition", "")
        duration_s  = int(f.get("duration_sec") or 0)
        call_id     = f.get("call_id") or rec["id"]

        key = f"{agent}|{dt.isoformat()}|{call_id}"
        if key in seen:
            continue
        seen.add(key)

        aic_records.append({"dt": dt, "agent": agent, "is_outbound": is_outbound})
        if is_outbound:
            records.append({
                "dt": dt, "agent": agent,
                "status": "", "disposition": disposition, "duration_sec": duration_s,
            })
    return records, aic_records


def write_sheet_metrics(aic, counts, conn_daily):
    """Compute 24 BDR KPI summary values and write to Dashboard Summary Sheet Block E."""
    import json as _json

    gsheet_id  = os.environ.get("GSHEET_ID", "1RfGRwfRQlwPAenslF7X6yvBT1XLC-CHm85KfqNV9SK0")
    creds_raw  = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_raw:
        print("  GOOGLE_CREDENTIALS_JSON not set — skipping Sheet write")
        return

    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        print("  ERROR: gspread / google-auth not installed — skipping Sheet write")
        return

    creds = Credentials.from_service_account_info(
        _json.loads(creds_raw),
        scopes=["https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"],
    )
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(gsheet_id).worksheet("Dashboard Summary")

    today = datetime.date.today()

    # Last complete Mon-Fri week
    days_since_mon = today.weekday()           # 0 = Mon
    lw_start = today - datetime.timedelta(days=days_since_mon + 7)
    lw_end   = lw_start + datetime.timedelta(days=4)

    # 4wk window ending on lw_end
    w4_start = lw_end - datetime.timedelta(days=27)

    # 90d = last 3 complete calendar months
    cur_month_start = datetime.date(today.year, today.month, 1)
    months_90d = []
    for delta in (3, 2, 1):
        m = cur_month_start.replace(day=1)
        for _ in range(delta):
            m = (m - datetime.timedelta(days=1)).replace(day=1)
        months_90d.append((m.year, m.month))

    def _sum_period(daily, agent, start, end):
        return sum(v for d, v in daily.get(agent, {}).items() if start <= d <= end)

    def _sum_months(daily, agent, yms):
        return sum(monthly_sum(daily.get(agent, {}), y, m) for y, m in yms)

    REPS = [
        ("edgar",   "Edgar Morales"),
        ("leandro", "Leandro Greasebook"),
    ]

    rows = []
    for key, full_name in REPS:
        aic_lw      = _sum_period(aic, full_name, lw_start, lw_end) / 60
        aic_4wk_tot = _sum_period(aic, full_name, w4_start, lw_end) / 60
        aic_4wk     = aic_4wk_tot / 4
        aic_90d_tot = _sum_months(aic, full_name, months_90d) / 60
        aic_90d     = aic_90d_tot / 13

        calls_lw      = _sum_period(counts, full_name, lw_start, lw_end)
        calls_4wk_tot = _sum_period(counts, full_name, w4_start, lw_end)
        calls_4wk     = calls_4wk_tot / 4
        calls_90d_tot = _sum_months(counts, full_name, months_90d)
        calls_90d     = calls_90d_tot / 13

        conn_lw      = _sum_period(conn_daily, full_name, lw_start, lw_end)
        conn_4wk     = _sum_period(conn_daily, full_name, w4_start, lw_end)
        conn_90d_tot = _sum_months(conn_daily, full_name, months_90d)

        conn_rate_lw  = round(conn_lw  / calls_lw  * 100, 2) if calls_lw  else None
        conn_rate_4wk = round(conn_4wk / calls_4wk_tot * 100, 2) if calls_4wk_tot else None
        conn_rate_90d = round(conn_90d_tot / calls_90d_tot * 100, 2) if calls_90d_tot else None

        aic_min_lw  = _sum_period(aic, full_name, lw_start, lw_end)
        aic_min_4wk = _sum_period(aic, full_name, w4_start, lw_end)
        aic_min_90d = _sum_months(aic, full_name, months_90d)

        eff_lw  = round(calls_lw  / (aic_min_lw  / 60), 1) if aic_min_lw  > 0 else None
        eff_4wk = round(calls_4wk / (aic_min_4wk / 60 / 4), 1) if aic_min_4wk > 0 else None
        eff_90d = round(calls_90d_tot / (aic_min_90d / 60), 1) if aic_min_90d > 0 else None

        rows.append([f"{key}_aic_lw",   str(round(aic_lw,  2))])
        rows.append([f"{key}_aic_4wk",  str(round(aic_4wk, 2))])
        rows.append([f"{key}_aic_90d",  str(round(aic_90d, 2))])
        rows.append([f"{key}_calls_lw",  str(round(calls_lw,  0))])
        rows.append([f"{key}_calls_4wk", str(round(calls_4wk, 0))])
        rows.append([f"{key}_calls_90d", str(round(calls_90d, 0))])
        rows.append([f"{key}_conn_lw",  str(conn_rate_lw)  if conn_rate_lw  is not None else ""])
        rows.append([f"{key}_conn_4wk", str(conn_rate_4wk) if conn_rate_4wk is not None else ""])
        rows.append([f"{key}_conn_90d", str(conn_rate_90d) if conn_rate_90d is not None else ""])
        rows.append([f"{key}_eff_lw",   str(eff_lw)  if eff_lw  is not None else ""])
        rows.append([f"{key}_eff_4wk",  str(eff_4wk) if eff_4wk is not None else ""])
        rows.append([f"{key}_eff_90d",  str(eff_90d) if eff_90d is not None else ""])

    # Write all 24 rows to Block E starting at row 92
    ws.update("A92:B115", rows, value_input_option="RAW")
    print(f"  ✓ Sheet updated — {len(rows)} rows written to Block E (rows 92-115)")


# ── Compute ───────────────────────────────────────────────────────────────────

def compute_aic(aic_records):
    """Hybrid rule: a session opens only on an outbound call.
    Any call (inbound or outbound) within GAP_MINUTES of the last call
    in an open session keeps it alive. Inbounds outside an open session
    are ignored — they don't start new sessions."""
    by_agent_day = defaultdict(lambda: defaultdict(list))
    for r in aic_records:
        by_agent_day[r["agent"]][r["dt"].date()].append((r["dt"], r["is_outbound"]))
    result = {}
    gap = datetime.timedelta(minutes=GAP_MINUTES)
    for agent, days in by_agent_day.items():
        result[agent] = {}
        for day, calls in days.items():
            calls = sorted(calls, key=lambda x: x[0])
            # dedupe by timestamp
            seen_ts, unique = set(), []
            for dt, is_out in calls:
                if dt not in seen_ts:
                    seen_ts.add(dt)
                    unique.append((dt, is_out))
            calls = unique
            total = 0.0
            session_start = None
            session_last  = None
            for dt, is_outbound in calls:
                if session_start is None:
                    if is_outbound:
                        session_start = dt
                        session_last  = dt
                    # inbound with no open session — skip
                else:
                    if dt - session_last <= gap:
                        session_last = dt          # extend session
                    else:
                        total += (session_last - session_start).total_seconds() / 60
                        if is_outbound:
                            session_start = dt     # open new session
                            session_last  = dt
                        else:
                            session_start = None   # inbound after gap — don't open
                            session_last  = None
            if session_start is not None:
                total += (session_last - session_start).total_seconds() / 60
            result[agent][day] = round(total, 1)
    return result

def call_counts(records):
    out = defaultdict(lambda: defaultdict(int))
    for r in records:
        out[r["agent"]][r["dt"].date()] += 1
    return out

def compute_connections(records):
    out = defaultdict(lambda: defaultdict(int))
    for r in records:
        if is_connection(r.get("disposition", "")):
            out[r["agent"]][r["dt"].date()] += 1
    return out

def compute_heatmaps(records):
    hour_dials = defaultdict(lambda: defaultdict(int))
    hour_conns = defaultdict(lambda: defaultdict(int))
    dow_dials  = defaultdict(lambda: defaultdict(int))
    dow_conns  = defaultdict(lambda: defaultdict(int))
    for r in records:
        h = r["dt"].hour
        d = r["dt"].weekday()   # 0=Mon … 4=Fri (we ignore Sat/Sun)
        a = r["agent"]
        hour_dials[a][h] += 1
        dow_dials[a][d]  += 1
        if is_connection(r.get("disposition", "")):
            hour_conns[a][h] += 1
            dow_conns[a][d]  += 1
    return hour_dials, hour_conns, dow_dials, dow_conns

# ── Aggregation helpers ───────────────────────────────────────────────────────

def workdays_up_to(y, m, up_to):
    start = datetime.date(y, m, 1)
    end   = min(up_to, datetime.date(y, m, calendar.monthrange(y, m)[1]))
    return sum(1 for i in range((end - start).days + 1)
               if (start + datetime.timedelta(i)).weekday() < 5)

def monthly_sum(daily, y, m):
    return sum(v for d, v in daily.items() if d.year == y and d.month == m)

def month_stats(daily, y, m, up_to):
    entries  = {d: v for d, v in daily.items() if d.year == y and d.month == m}
    total    = sum(entries.values())
    days_act = len(entries)
    wdays    = workdays_up_to(y, m, up_to)
    return {
        "ym":          (y, m),
        "label":       datetime.date(y, m, 1).strftime("%b %Y"),
        "total_min":   round(total, 1),
        "total_hrs":   round(total / 60, 2),
        "days_active": days_act,
        "workdays":    wdays,
        "avg_active":  round(total / days_act, 1) if days_act else 0,
        "avg_workday": round(total / wdays,    1) if wdays    else 0,
        "partial":     (y == up_to.year and m == up_to.month),
    }

def year_stats(daily, y, up_to):
    entries  = {d: v for d, v in daily.items() if d.year == y}
    total    = sum(entries.values())
    days_act = len(entries)
    start = datetime.date(y, 1, 1)
    end   = min(up_to, datetime.date(y, 12, 31))
    wdays = sum(1 for i in range((end - start).days + 1)
                if (start + datetime.timedelta(i)).weekday() < 5)
    return {
        "year":        y,
        "total_min":   round(total, 1),
        "total_hrs":   round(total / 60, 2),
        "days_active": days_act,
        "workdays":    wdays,
        "avg_workday": round(total / wdays, 1) if wdays else 0,
    }

# ── HTML table/insight builders ───────────────────────────────────────────────

def _build_dial_tables(agents, counts, all_yms, today):
    parts = []
    for a in agents:
        c, name = AGENT_COLOR[a], AGENT_DISPLAY[a]
        tbody, prev_y, prev_d = "", None, None
        for y, m in all_yms:
            if y != prev_y:
                tbody += (f'<tr class="yr-sep"><td colspan="5" style="color:{c};'
                          f'padding:10px 10px 4px;font-weight:700;font-size:.82rem">{y}</td></tr>')
                prev_y = y
            d     = monthly_sum(counts.get(a, {}), y, m)
            wdays = workdays_up_to(y, m, today)
            avg   = round(d / wdays, 1) if wdays else 0
            lbl   = datetime.date(y, m, 1).strftime("%b %Y")
            ptag  = " <span class='partial'>partial</span>" if (y == today.year and m == today.month) else ""
            if prev_d and prev_d > 0:
                pct   = (d - prev_d) / prev_d * 100
                delta = f'<span class="{"good" if pct >= 0 else "flag"}">{pct:+.0f}%</span>'
            else:
                delta = "—"
            if d > 0:
                prev_d = d
            tbody += (f'<tr><td>{lbl}{ptag}</td><td><strong>{d:,}</strong></td>'
                      f'<td>{wdays}</td><td>{avg:.1f}</td><td>{delta}</td></tr>')
        parts.append(
            f'<div class="section">'
            f'<div class="stitle" style="color:{c}">{name} — Monthly Dial Volume</div>'
            f'<table><thead><tr><th>Month</th><th>Dials</th><th>Workdays</th>'
            f'<th>Avg / Workday</th><th>vs Prior Month</th></tr></thead>'
            f'<tbody>{tbody}</tbody></table></div>'
        )
    return "".join(parts)


def _build_conn_tables(agents, counts, conn_daily, all_yms, today):
    table_parts, insight_parts = [], []
    for a in agents:
        c, name = AGENT_COLOR[a], AGENT_DISPLAY[a]
        tbody, prev_y = "", None
        all_d, all_cn, best_rate, best_lbl = 0, 0, 0.0, ""
        for y, m in all_yms:
            if y != prev_y:
                tbody += (f'<tr class="yr-sep"><td colspan="5" style="color:{c};'
                          f'padding:10px 10px 4px;font-weight:700;font-size:.82rem">{y}</td></tr>')
                prev_y = y
            d   = monthly_sum(counts.get(a, {}), y, m)
            cn  = monthly_sum(conn_daily.get(a, {}), y, m)
            all_d  += d
            all_cn += cn
            rate = round(cn / d * 100, 2) if d > 0 else 0.0
            lbl  = datetime.date(y, m, 1).strftime("%b %Y")
            ptag = " <span class='partial'>partial</span>" if (y == today.year and m == today.month) else ""
            rcls = "good" if rate >= 2.0 else ("flag" if (rate < 1.0 and d > 50) else "")
            bw   = min(int(rate / 5 * 100), 100)
            tbody += (f'<tr><td>{lbl}{ptag}</td><td>{d:,}</td><td>{cn}</td>'
                      f'<td><span class="{rcls}">{rate:.2f}%</span></td>'
                      f'<td><div class="bar" style="width:{max(bw,1)}%;background:{c}"></div></td>'
                      f'</tr>')
            if rate > best_rate and d > 50:
                best_rate, best_lbl = rate, lbl
        all_time_rate = round(all_cn / all_d * 100, 2) if all_d else 0
        ins = (f'<strong>{name}</strong> all-time connection rate: '
               f'<strong>{all_time_rate:.2f}%</strong> '
               f'({all_cn:,} connections / {all_d:,} dials).')
        if best_lbl:
            ins += f' Peak month: <span class="good">{best_lbl} at {best_rate:.2f}%</span>.'
        insight_parts.append(f'<div class="insight">{ins}</div>')
        table_parts.append(
            f'<div class="section">'
            f'<div class="stitle" style="color:{c}">{name} — Monthly Connection Rate</div>'
            f'<table><thead><tr><th>Month</th><th>Dials</th><th>Connections</th>'
            f'<th>Conn %</th><th></th></tr></thead>'
            f'<tbody>{tbody}</tbody></table></div>'
        )
    return "".join(table_parts), "".join(insight_parts)


def _build_eff_insight(agents, counts, aic, all_yms):
    parts = []
    for a in agents:
        c, name = AGENT_COLOR[a], AGENT_DISPLAY[a]
        vals = []
        for y, m in all_yms:
            d       = monthly_sum(counts.get(a, {}), y, m)
            aic_min = monthly_sum(aic.get(a, {}), y, m)
            if aic_min > 60 and d > 0:
                vals.append((round(d / (aic_min / 60), 1), datetime.date(y, m, 1).strftime("%b %Y")))
        if not vals:
            continue
        avg_eff            = round(sum(e for e, _ in vals) / len(vals), 1)
        best_eff, best_lbl = max(vals, key=lambda x: x[0])
        low_eff,  low_lbl  = min(vals, key=lambda x: x[0])
        parts.append(
            f'<div class="insight">'
            f'<strong>{name}</strong> averages <strong>{avg_eff} dials / AIC hour</strong> '
            f'across {len(vals)} months. '
            f'<span class="good">Best: {best_lbl} ({best_eff} dials/hr)</span> · '
            f'<span class="flag">Lowest: {low_lbl} ({low_eff} dials/hr)</span>. '
            f'Low efficiency = chair time without dialing. High efficiency = fully locked in.'
            f'</div>'
        )
    return "".join(parts)


def _build_heatmap_insight(agents, hour_dials, hour_conns):
    parts = []
    for a in agents:
        name = AGENT_DISPLAY[a]
        h_d  = hour_dials.get(a, {})
        h_c  = hour_conns.get(a, {})
        if not h_d:
            continue
        peak_h = max(h_d, key=h_d.get)
        rates  = {h: h_c.get(h, 0) / h_d[h] * 100 for h in h_d if h_d[h] >= 50}
        txt = f'<strong>{name}</strong> dials most heavily at <strong>{peak_h:02d}:00</strong>.'
        if rates:
            best_rh  = max(rates, key=rates.get)
            best_rpct = round(rates[best_rh], 1)
            txt += (f' Best connect window: <span class="good">{best_rh:02d}:00 '
                    f'({best_rpct}% conn rate)</span>.')
        parts.append(f'<div class="insight">{txt}</div>')
    return "".join(parts)

# ── Main HTML builder ─────────────────────────────────────────────────────────

def build_html(aic, counts, conn_daily, hour_dials, hour_conns, dow_dials, dow_conns, records):
    today   = datetime.date.today()
    agents  = [a for a in AGENT_ORDER if a in aic]

    all_days  = sorted({r["dt"].date() for r in records})
    all_yms   = sorted({(d.year, d.month) for d in all_days})
    all_years = sorted({d.year for d in all_days})

    monthly_labels = [datetime.date(y, m, 1).strftime("%b '%y") for y, m in all_yms]

    # ── AIC monthly/yearly stats ──────────────────────────────────────────
    agent_monthly = {a: [month_stats(aic[a], y, m, today) for y, m in all_yms] for a in agents}
    agent_yearly  = {a: [year_stats(aic[a], y, today) for y in all_years]       for a in agents}

    # ── Weekly spine (shared by AIC and Dials weekly charts) ─────────────
    all_weeks_set, agent_week_aic, agent_week_dial = set(), {}, {}
    for a in agents:
        waic, wdial, wfirst = defaultdict(float), defaultdict(int), {}
        for d, v in aic[a].items():
            yw = (d.isocalendar()[0], d.isocalendar()[1])
            waic[yw] += v
            if yw not in wfirst or d < wfirst[yw]:
                wfirst[yw] = d
        for d, v in counts.get(a, {}).items():
            yw = (d.isocalendar()[0], d.isocalendar()[1])
            wdial[yw] += v
        agent_week_aic[a]  = (waic, wfirst)
        agent_week_dial[a] = wdial
        all_weeks_set.update(waic.keys())

    all_weeks   = sorted(all_weeks_set)
    week_labels = []
    for yw in all_weeks:
        d = None
        for a in agents:
            d = agent_week_aic[a][1].get(yw) or d
        week_labels.append(d.strftime("%-d %b '%y") if d else str(yw))

    # ── Chart datasets ────────────────────────────────────────────────────
    monthly_datasets = [
        {"label": AGENT_DISPLAY[a],
         "data": [agent_monthly[a][i]["total_hrs"] for i in range(len(all_yms))],
         "backgroundColor": AGENT_COLOR[a], "borderRadius": 4}
        for a in agents
    ]
    weekly_datasets = [
        {"label": AGENT_DISPLAY[a],
         "data": [round(agent_week_aic[a][0].get(yw, 0) / 60, 2) for yw in all_weeks],
         "borderColor": AGENT_COLOR[a], "backgroundColor": AGENT_COLOR[a] + "22",
         "fill": True, "tension": 0.3, "borderWidth": 2,
         "pointRadius": 2, "pointHoverRadius": 5, "spanGaps": True}
        for a in agents
    ]
    dial_monthly_ds = [
        {"label": AGENT_DISPLAY[a],
         "data": [monthly_sum(counts.get(a, {}), y, m) for y, m in all_yms],
         "backgroundColor": AGENT_COLOR[a], "borderRadius": 4}
        for a in agents
    ]
    dial_weekly_ds = [
        {"label": AGENT_DISPLAY[a],
         "data": [agent_week_dial[a].get(yw, 0) for yw in all_weeks],
         "borderColor": AGENT_COLOR[a], "backgroundColor": AGENT_COLOR[a] + "22",
         "fill": True, "tension": 0.3, "borderWidth": 2,
         "pointRadius": 2, "pointHoverRadius": 5, "spanGaps": True}
        for a in agents
    ]
    conn_rate_ds = []
    for a in agents:
        c = AGENT_COLOR[a]
        rates = [
            round(monthly_sum(conn_daily.get(a, {}), y, m) /
                  monthly_sum(counts.get(a, {}), y, m) * 100, 2)
            if monthly_sum(counts.get(a, {}), y, m) > 0 else None
            for y, m in all_yms
        ]
        conn_rate_ds.append({
            "label": AGENT_DISPLAY[a], "data": rates,
            "borderColor": c, "backgroundColor": c + "33",
            "fill": False, "tension": 0.3, "borderWidth": 2,
            "pointRadius": 3, "pointHoverRadius": 6, "spanGaps": True,
        })
    eff_ds = []
    for a in agents:
        effs = []
        for y, m in all_yms:
            d       = monthly_sum(counts.get(a, {}), y, m)
            aic_min = monthly_sum(aic.get(a, {}), y, m)
            effs.append(round(d / (aic_min / 60), 1) if aic_min > 60 else None)
        eff_ds.append({"label": AGENT_DISPLAY[a], "data": effs,
                       "backgroundColor": AGENT_COLOR[a], "borderRadius": 4})

    hour_labels = [f"{h:02d}:00" for h in range(24)]
    dow_labels  = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    hour_ds, dow_ds = [], []
    for a in agents:
        c = AGENT_COLOR[a]
        h_d = [hour_dials.get(a, {}).get(h, 0) for h in range(24)]
        h_c = [hour_conns.get(a, {}).get(h, 0) for h in range(24)]
        h_r = [round(h_c[h] / h_d[h] * 100, 1) if h_d[h] > 0 else None for h in range(24)]
        d_d = [dow_dials.get(a, {}).get(d, 0) for d in range(5)]
        d_c = [dow_conns.get(a, {}).get(d, 0) for d in range(5)]
        d_r = [round(d_c[d] / d_d[d] * 100, 1) if d_d[d] > 0 else None for d in range(5)]
        n = AGENT_DISPLAY[a]
        hour_ds += [
            {"type": "bar",  "label": f"{n} Dials", "data": h_d,
             "backgroundColor": c + "88", "borderRadius": 3, "yAxisID": "y"},
            {"type": "line", "label": f"{n} Conn%", "data": h_r,
             "borderColor": c, "backgroundColor": "transparent",
             "borderWidth": 2, "pointRadius": 3, "tension": 0.3,
             "yAxisID": "y1", "spanGaps": True},
        ]
        dow_ds += [
            {"type": "bar",  "label": f"{n} Dials", "data": d_d,
             "backgroundColor": c + "88", "borderRadius": 3, "yAxisID": "y"},
            {"type": "line", "label": f"{n} Conn%", "data": d_r,
             "borderColor": c, "backgroundColor": "transparent",
             "borderWidth": 2, "pointRadius": 4, "tension": 0.1,
             "yAxisID": "y1", "spanGaps": True},
        ]

    # ── HTML tables / insights for new pages ──────────────────────────────
    dial_tables_html               = _build_dial_tables(agents, counts, all_yms, today)
    conn_tables_html, conn_ins_html = _build_conn_tables(agents, counts, conn_daily, all_yms, today)
    eff_insight_html               = _build_eff_insight(agents, counts, aic, all_yms)
    heatmap_insight_html           = _build_heatmap_insight(agents, hour_dials, hour_conns)

    # ── AIC page: year cards ──────────────────────────────────────────────
    cards_html = ""
    for a in agents:
        c, name = AGENT_COLOR[a], AGENT_DISPLAY[a]
        yr_blocks = ""
        for ys in agent_yearly[a]:
            partial_note = (" <span style='color:#8b949e;font-size:.72rem'>(YTD)</span>"
                            if ys["year"] == today.year else "")
            yr_blocks += (
                f'<div class="yr-row">'
                f'<span class="yr-label">{ys["year"]}{partial_note}</span>'
                f'<span class="yr-stat">{ys["total_hrs"]:.0f} hrs</span>'
                f'<span class="yr-sub">{ys["avg_workday"]:.0f} min/workday · {ys["days_active"]} days</span>'
                f'</div>'
            )
        total_calls = sum(counts.get(a, {}).values())
        cards_html += (
            f'<div class="card" style="border-top:3px solid {c}">'
            f'<div class="card-name" style="color:{c}">{name}</div>'
            f'{yr_blocks}'
            f'<div class="card-calls">{total_calls:,} outbound calls total</div>'
            f'</div>'
        )

    # ── AIC page: monthly tables ──────────────────────────────────────────
    month_tables_html = ""
    for a in agents:
        c, name = AGENT_COLOR[a], AGENT_DISPLAY[a]
        tbody, prev_year = "", None
        for ms in agent_monthly[a]:
            y, m   = ms["ym"]
            yr_sep = ""
            if y != prev_year:
                yr_sep    = (f'<tr class="yr-sep"><td colspan="7" style="color:{c};'
                             f'padding:10px 10px 4px;font-weight:700;font-size:.82rem;'
                             f'letter-spacing:.04em">{y}</td></tr>')
                prev_year = y
            partial = " <span class='partial'>partial</span>" if ms["partial"] else ""
            bw      = min(int(ms["total_hrs"] / 80 * 100), 100)
            tbody  += yr_sep + (
                f'<tr>'
                f'<td>{ms["label"]}{partial}</td>'
                f'<td><strong>{ms["total_hrs"]:.1f}</strong> hrs</td>'
                f'<td>{ms["total_min"]:.0f} min</td>'
                f'<td>{ms["days_active"]} / {ms["workdays"]}</td>'
                f'<td>{ms["avg_active"]:.0f} min</td>'
                f'<td>{ms["avg_workday"]:.0f} min</td>'
                f'<td><div class="bar" style="width:{max(bw,1)}%;background:{c}"></div></td>'
                f'</tr>'
            )
        # Insight
        full_years    = [y for y in all_years if y < today.year]
        insight_parts = []
        for fy in full_years:
            yr_entries = {d: v for d, v in aic[a].items() if d.year == fy}
            if not yr_entries: continue
            total_fy  = sum(yr_entries.values())
            first_mo  = min(d.month for d in yr_entries)
            wd_fy     = sum(workdays_up_to(fy, mo, datetime.date(fy, 12, 31))
                            for mo in range(first_mo, 13))
            avg_wd_fy = round(total_fy / wd_fy, 1) if wd_fy else 0
            active_fy = len(yr_entries)
            ms_fy     = [ms for ms in agent_monthly[a] if ms["ym"][0] == fy]
            best_m    = max(ms_fy, key=lambda x: x["total_hrs"]) if ms_fy else None
            worst_m   = min([ms for ms in ms_fy if ms["total_hrs"] > 0],
                            key=lambda x: x["total_hrs"]) if ms_fy else None
            part = (f"<strong>{fy}:</strong> {total_fy/60:.0f} hrs total · "
                    f"<strong>{avg_wd_fy:.0f} min/workday</strong> · {active_fy} active days")
            if best_m and worst_m:
                part += (f" · <span class='good'>best: {best_m['label']} "
                         f"({best_m['total_hrs']:.1f} hrs)</span>"
                         f" · <span class='flag'>weakest: {worst_m['label']} "
                         f"({worst_m['total_hrs']:.1f} hrs)</span>")
            insight_parts.append(part)
        if len(full_years) >= 2:
            y1, y2 = full_years[-2], full_years[-1]
            t1 = sum(v for d, v in aic[a].items() if d.year == y1)
            t2 = sum(v for d, v in aic[a].items() if d.year == y2)
            if t1 > 0:
                pct = (t2 - t1) / t1 * 100
                cls = "good" if pct > 0 else "flag"
                insight_parts.append(
                    f"<span class='{cls}'>{y1}→{y2}: <strong>{pct:+.0f}%</strong> year-over-year</span>"
                )
        if full_years:
            last_fy       = full_years[-1]
            last_entries  = {d: v for d, v in aic[a].items() if d.year == last_fy}
            last_total    = sum(last_entries.values())
            last_wdays    = sum(workdays_up_to(last_fy, mo, datetime.date(last_fy, 12, 31))
                                for mo in range(1, 13))
            last_avg      = round(last_total / last_wdays, 1) if last_wdays else 0
            cur_months    = [ms for ms in agent_monthly[a] if ms["ym"][0] == today.year]
            cur_wd_vals   = [ms["avg_workday"] for ms in cur_months if ms["total_min"] > 0]
            if cur_wd_vals:
                cur_avg = round(sum(cur_wd_vals) / len(cur_wd_vals), 0)
                diff    = cur_avg - last_avg
                cls     = "flag" if diff < -10 else ("good" if diff > 5 else "")
                insight_parts.append(
                    f"<span class='{cls}'>{today.year} YTD averages <strong>{cur_avg:.0f} min/workday</strong>"
                    f" — {abs(diff):.0f} min/day {'below' if diff < 0 else 'above'} {last_fy} average.</span>"
                )
        agent_insight = " &nbsp;·&nbsp; ".join(insight_parts) if insight_parts else ""
        month_tables_html += (
            f'<div class="section">'
            f'<div class="stitle" style="color:{c}">{name} — Month by Month</div>'
            f'<table><thead><tr>'
            f'<th>Month</th><th>AIC Hours</th><th>AIC Min</th>'
            f'<th>Days Active / Workdays</th><th>Avg / Active Day</th>'
            f'<th>Avg / Workday</th><th></th>'
            f'</tr></thead><tbody>{tbody}</tbody></table>'
            f'<div class="insight">{agent_insight}</div>'
            f'</div>'
        )

    # ── AIC page: daily log ───────────────────────────────────────────────
    daily_rows_html = ""
    for d in reversed(all_days):
        ym_key = f"{d.year}-{d.month:02d}"
        cells  = (f'<td class="dc"><span class="wd">{d.strftime("%a")}</span>'
                  f'{d.strftime(" %b %-d, %Y")}</td>')
        for a in agents:
            mins = aic.get(a, {}).get(d)
            cnt  = counts.get(a, {}).get(d, 0)
            c    = AGENT_COLOR[a]
            if mins is not None:
                bw = min(int(mins / 300 * 100), 100)
                cells += (f'<td><div class="db" style="width:{max(bw,4)}%;'
                          f'background:{c}18;border-left:3px solid {c}">'
                          f'<span style="color:{c};font-weight:600">{mins/60:.1f}h</span>'
                          f' <small>({mins:.0f}m · {cnt} calls)</small></div></td>')
            else:
                cells += '<td class="em">—</td>'
        daily_rows_html += f'<tr data-ym="{ym_key}">{cells}</tr>'

    agent_ths = "".join(
        f'<th style="color:{AGENT_COLOR[a]}">{AGENT_DISPLAY[a]}</th>' for a in agents
    )

    # ── Month filter buttons ──────────────────────────────────────────────
    filter_btns = '<button class="fbt active" onclick="filterDays(\'all\',this)">All</button>'
    prev_fy = None
    for y, m in reversed(all_yms):
        if y != prev_fy:
            filter_btns += f'<span class="fyr">{y}</span>'
            prev_fy = y
        lbl = datetime.date(y, m, 1).strftime("%b")
        filter_btns += (f'<button class="fbt" onclick="filterDays(\'{y}-{m:02d}\',this)">'
                        f'{lbl}</button>')

    now_str    = datetime.datetime.now().strftime("%B %-d, %Y")
    date_range = f"{all_days[0].strftime('%B %Y')} – {all_days[-1].strftime('%B %-d, %Y')}"

    # ── Serialize all chart data to JSON ──────────────────────────────────
    data_js = (
        f"const MONTHLY_LABELS={json.dumps(monthly_labels)};\n"
        f"const WEEKLY_LABELS={json.dumps(week_labels)};\n"
        f"const AIC_MONTHLY_DS={json.dumps(monthly_datasets)};\n"
        f"const AIC_WEEKLY_DS={json.dumps(weekly_datasets)};\n"
        f"const DIAL_MONTHLY_DS={json.dumps(dial_monthly_ds)};\n"
        f"const DIAL_WEEKLY_DS={json.dumps(dial_weekly_ds)};\n"
        f"const CONN_RATE_DS={json.dumps(conn_rate_ds)};\n"
        f"const EFF_DS={json.dumps(eff_ds)};\n"
        f"const HOUR_LABELS={json.dumps(hour_labels)};\n"
        f"const HOUR_DS={json.dumps(hour_ds)};\n"
        f"const DOW_LABELS={json.dumps(dow_labels)};\n"
        f"const DOW_DS={json.dumps(dow_ds)};\n"
    )

    # ── Assemble full HTML ────────────────────────────────────────────────
    head = (
        '<!DOCTYPE html><html lang="en"><head>'
        '<meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<title>KPI Report — Edgar &amp; Leandro</title>'
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>'
        f'<style>{_CSS}</style>'
        '</head>'
    )

    sidebar = (
        '<nav class="sidebar">'
        '<div class="sb-brand">Greasebook</div>'
        '<div class="sb-title">KPI Dashboard</div>'
        '<div class="nav-sep">Views</div>'
        '<button class="nav-item active" id="nav-aic"        onclick="showPage(\'aic\')">&#9685; AIC Time</button>'
        '<button class="nav-item"        id="nav-dials"      onclick="showPage(\'dials\')">&#9685; Dial Volume</button>'
        '<button class="nav-item"        id="nav-connect"    onclick="showPage(\'connect\')">&#9685; Connection Rate</button>'
        '<button class="nav-item"        id="nav-efficiency" onclick="showPage(\'efficiency\')">&#9685; Efficiency</button>'
        '<button class="nav-item"        id="nav-heatmap"    onclick="showPage(\'heatmap\')">&#9685; Call Timing</button>'
        '</nav>'
    )

    page_aic = (
        f'<div id="page-aic" class="page">'
        f'<h1>Ass-in-Chair KPI Report</h1>'
        f'<div class="meta">Edgar Morales &amp; Leandro &nbsp;·&nbsp; {date_range}'
        f' &nbsp;·&nbsp; Generated {now_str} &nbsp;·&nbsp; Outbound-initiated sessions &nbsp;·&nbsp; 20-min gap rule</div>'
        f'<div class="cards">{cards_html}</div>'
        f'<div class="mbox">'
        f'<h2>What You\'re Looking At — How "Ass-in-Chair" Time Is Measured</h2>'
        f'<p>This report measures one thing: <strong style="color:#e6edf3">how many hours each BDR was actively sitting at their desk and dialing.</strong> It is not a call count. It is not talk time. It is <em>active dialing session time</em> — the window during which a rep is on the phones.</p>'
        f'<h3>The 20-Minute Gap Rule</h3>'
        f'<p>Every call has a timestamp. The calculation works like this:</p>'
        f'<ul>'
        f'<li>The first <strong style="color:#e6edf3">outbound call</strong> of the day opens a <strong style="color:#e6edf3">dialing block.</strong></li>'
        f'<li>Each next call — outbound <em>or inbound</em> — keeps the block open as long as it happens <strong style="color:#e6edf3">within 20 minutes</strong> of the previous one.</li>'
        f'<li>If the gap exceeds 20 minutes, the block closes. Only the next <strong style="color:#e6edf3">outbound</strong> call opens a new block — an inbound call after a gap does not.</li>'
        f'<li><strong style="color:#e6edf3">AIC time = total minutes across all blocks for the day.</strong></li>'
        f'</ul>'
        f'<h3>Worked Example</h3>'
        f'<div class="ex-block">'
        f'<span class="pill">9:00 AM</span> <span class="arrow">→</span>'
        f'<span class="pill">9:14 AM</span> <span class="arrow">→</span>'
        f'<span class="pill">9:28 AM</span> <span class="arrow">✗ 31-min gap ✗</span>'
        f'<span class="pill">9:59 AM</span> <span class="arrow">→</span>'
        f'<span class="pill">10:11 AM</span><br><br>'
        f'Block 1: 9:00 → 9:28 = <strong style="color:#4F9CF9">28 min</strong> &nbsp;|&nbsp;'
        f'Block 2: 9:59 → 10:11 = <strong style="color:#4F9CF9">12 min</strong> &nbsp;|&nbsp;'
        f'<strong style="color:#fff">Total AIC = 40 min</strong>'
        f'</div>'
        f'<h3>Avg / Workday vs Avg / Active Day</h3>'
        f'<p><strong style="color:#e6edf3">Avg / Active Day</strong> = average only on days they actually dialed. '
        f'<strong style="color:#e6edf3">Avg / Workday</strong> = average spread across every Monday–Friday in the period, '
        f'including days they didn\'t dial at all. The second number is the harder benchmark.</p>'
        f'</div>'
        f'<div class="section">'
        f'<div class="stitle">Monthly AIC Hours — Full History</div>'
        f'<div class="chart-wrap tall"><canvas id="monthlyChart"></canvas></div>'
        f'<div class="insight">Edgar has been dialing since <strong>January 2024</strong>; Leandro joined in <strong>April 2024</strong>. '
        f'Both reps held <strong>55–68 hrs/month</strong> across their strongest stretches. '
        f'<span class="flag">January and February 2026 are the softest back-to-back months in the entire dataset for both reps.</span></div>'
        f'</div>'
        f'<div class="section">'
        f'<div class="stitle">Weekly AIC Hours — Trend</div>'
        f'<div class="chart-wrap tall"><canvas id="weeklyChart"></canvas></div>'
        f'<div class="insight">Across the full history, <strong>Edgar is the more consistent dialer.</strong> '
        f'<strong>Leandro swings wider</strong> and has had near-zero weeks. '
        f'<span class="flag">The 2026 trend shows low weeks becoming more frequent for both reps.</span></div>'
        f'</div>'
        f'{month_tables_html}'
        f'<div class="section">'
        f'<div class="stitle">Daily Detail Log</div>'
        f'<div class="filters">{filter_btns}</div>'
        f'<table><thead><tr><th>Date</th>{agent_ths}</tr></thead>'
        f'<tbody id="dailyBody">{daily_rows_html}</tbody></table>'
        f'</div>'
        f'</div>'
    )

    page_dials = (
        f'<div id="page-dials" class="page" style="display:none">'
        f'<h1>Dial Volume</h1>'
        f'<div class="meta">Total outbound dials per period &nbsp;·&nbsp; {date_range}</div>'
        f'<div class="section">'
        f'<div class="stitle">Monthly Dials — Full History</div>'
        f'<div class="chart-wrap tall"><canvas id="dialMonthlyChart"></canvas></div>'
        f'</div>'
        f'<div class="section">'
        f'<div class="stitle">Weekly Dials — Trend</div>'
        f'<div class="chart-wrap tall"><canvas id="dialWeeklyChart"></canvas></div>'
        f'</div>'
        f'{dial_tables_html}'
        f'</div>'
    )

    page_connect = (
        f'<div id="page-connect" class="page" style="display:none">'
        f'<h1>Connection Rate</h1>'
        f'<div class="meta">Meaningful connections ÷ total dials &nbsp;·&nbsp; {date_range}</div>'
        f'<div class="section">'
        f'<div class="stitle">Monthly Connection Rate</div>'
        f'<div class="chart-wrap tall"><canvas id="connRateChart"></canvas></div>'
        f'<div class="insight">'
        f'A <strong>"connection"</strong> is any outbound call with a disposition of '
        f'<strong>Connection, Connect, Booked, or Callback</strong> — '
        f'dispositions that require a live conversation to receive.'
        f'</div>'
        f'</div>'
        f'{conn_tables_html}'
        f'{conn_ins_html}'
        f'</div>'
    )

    page_efficiency = (
        f'<div id="page-efficiency" class="page" style="display:none">'
        f'<h1>Efficiency — Dials per AIC Hour</h1>'
        f'<div class="meta">How intensely are they using their chair time? &nbsp;·&nbsp; {date_range}</div>'
        f'<div class="section">'
        f'<div class="stitle">Monthly Dials per AIC Hour</div>'
        f'<div class="chart-wrap tall"><canvas id="effChart"></canvas></div>'
        f'<div class="insight">'
        f'<strong>What this measures:</strong> total dials in a month divided by AIC hours in that month. '
        f'A rep can have high AIC hours but low efficiency — meaning they\'re sitting there but not dialing hard. '
        f'High efficiency = they\'re on the phone the moment the clock starts. '
        f'Months with fewer than 1 AIC hour are excluded.'
        f'</div>'
        f'</div>'
        f'{eff_insight_html}'
        f'</div>'
    )

    page_heatmap = (
        f'<div id="page-heatmap" class="page" style="display:none">'
        f'<h1>Call Timing Heatmap</h1>'
        f'<div class="meta">When are they dialing — and when does it work best? &nbsp;·&nbsp; {date_range}</div>'
        f'<div class="section">'
        f'<div class="stitle">Hour-of-Day: Dial Volume + Connection Rate</div>'
        f'<div class="chart-wrap tall"><canvas id="hourChart"></canvas></div>'
        f'<div class="insight">'
        f'Bars = total dials at that hour (all-time). '
        f'Line = connection rate at that hour. '
        f'Use this to identify the best windows and confirm reps are dialing in those windows.'
        f'</div>'
        f'</div>'
        f'<div class="section">'
        f'<div class="stitle">Day-of-Week: Dial Volume + Connection Rate</div>'
        f'<div class="chart-wrap"><canvas id="dowChart"></canvas></div>'
        f'</div>'
        f'{heatmap_insight_html}'
        f'</div>'
    )

    script = (
        f'<script>\n{data_js}\n{_JS}</script>'
    )

    return (
        head
        + '<body>'
        + sidebar
        + f'<main class="main-content">{page_aic}{page_dials}{page_connect}{page_efficiency}{page_heatmap}</main>'
        + script
        + '</body></html>'
    )


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    airtable_key = os.environ.get("AIRTABLE_API_KEY")
    base_id      = os.environ.get("AIRTABLE_BASE_ID")

    if airtable_key and base_id:
        print("Fetching from Airtable…")
        raw      = fetch_airtable(airtable_key, base_id)
        records, aic_records = transform_airtable_records(raw)
        print(f"  {len(raw)} Airtable records fetched")
    else:
        print("Loading CSVs… (set AIRTABLE_API_KEY + AIRTABLE_BASE_ID to use Airtable)")
        records     = load_all(ALL_FILES)
        aic_records = load_aic_records(ALL_FILES)

    print(f"Total unique outbound calls: {len(records):,}")
    print(f"Total calls for AIC (in+out): {len(aic_records):,}\n")

    aic        = compute_aic(aic_records)
    counts     = call_counts(records)
    conn_daily = compute_connections(records)
    hour_dials, hour_conns, dow_dials, dow_conns = compute_heatmaps(records)

    for a in AGENT_ORDER:
        if a not in aic:
            continue
        total_dials = sum(counts.get(a, {}).values())
        total_conns = sum(conn_daily.get(a, {}).values())
        total_aic   = sum(aic[a].values())
        rate = total_conns / total_dials * 100 if total_dials else 0
        print(f"  {a}: {total_aic/60:.1f} AIC hrs · {total_dials:,} dials · "
              f"{total_conns} connections ({rate:.2f}%)")

    html = build_html(aic, counts, conn_daily, hour_dials, hour_conns, dow_dials, dow_conns, records)
    out  = Path("greg_report.html")
    out.write_text(html, encoding="utf-8")
    print(f"\nReport → {out.resolve()}")

    # Write KPI summaries to Google Sheet (skipped if creds not set)
    if airtable_key:
        print("\nWriting summary metrics to Google Sheet…")
        write_sheet_metrics(aic, counts, conn_daily)

    # Open in browser only when running locally (not in CI)
    if not os.environ.get("CI"):
        webbrowser.open(f"file://{out.resolve()}")


if __name__ == "__main__":
    main()
