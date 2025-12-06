from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

import pandas as pd
import openpyxl  # needed by pandas for Excel writing
from copy import copy
import tempfile
import os
import requests
from datetime import datetime
import re
import base64
import traceback


# -------------------- FASTAPI APP & CORS --------------------

app = FastAPI()

# For now: allow everything (no credentials) so Blocks can talk to this.
# Once it's working, you can lock this down to your exact Blocks origin.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],        # ok because allow_credentials=False
    allow_credentials=False,    # IMPORTANT when using "*"
    allow_methods=["*"],
    allow_headers=["*"],
)


# -------------------- HELPERS --------------------

TIME_FORMATS = ["%I:%M %p", "%H:%M", "%I:%M%p", "%H:%M:%S"]
SNAP_WINDOWS = [(310, 360, 360), (790, 840, 840), (1270, 1320, 1320)]
MERGE_GAP_MIN = 60
MIN_BREAK_MIN = 45
MAX_SPAN_MIN = 14 * 60
AM_START, AM_END = 6 * 60, 14 * 60
PM_START, PM_END = 14 * 60, 22 * 60
BIRD_AM_START, BIRD_AM_END = 8 * 60, 14 * 60
BIRD_PM_START, BIRD_PM_END = 14 * 60, 20 * 60


def round_quarter(x: float) -> float:
    return round(float(x) * 4) / 4


def to_minutes(val):
    if val is None or str(val).strip() == "":
        return None
    s = str(val).strip()
    for fmt in TIME_FORMATS:
        try:
            dt = datetime.strptime(s, fmt)
            return dt.hour * 60 + dt.minute
        except ValueError:
            continue
    return None


def snap(minutes):
    if minutes is None:
        return None
    for low, high, target in SNAP_WINDOWS:
        if low <= minutes < high:
            return target
    return minutes


def as_float(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def norm_sched(s):
    return re.sub(r"[^0-9a-z:apm/\- ]", "", str(s).lower())


def is_bird_day(s):
    s = norm_sched(s)
    return (
        "eagles/hawks" in s
        or "eagleshawks" in s
        or "08:00 - 20:00" in s
        or "08:00-20:00" in s
        or "8am-8pm" in s
    )


def is_am_day(s):
    s = norm_sched(s)
    return (
        "06:00 - 14:00" in s
        or "06:00-14:00" in s
        or "6am-2pm" in s
        or "am 6-2" in s
    )


def is_pm_day(s):
    s = norm_sched(s)
    return (
        "14:00 - 22:00" in s
        or "14:00-22:00" in s
        or "2pm-10pm" in s
        or "pm 2-10" in s
    )


def is_bird_night(s):
    s = norm_sched(s)
    return (
        "20:00 - 08:00" in s
        or "20:00-08:00" in s
        or "8pm-8am" in s
        or "owls/falcons" in s
    )


def is_std_night(s):
    s = norm_sched(s)
    if is_bird_night(s):
        return False
    return (
        "20:00-06:00" in s
        or "22:00-06:00" in s
        or "8pm-6am" in s
        or "10pm-6am" in s
        or "nightshift22:00-06:00" in s
    )


def is_22_06_night(s):
    s = norm_sched(s)
    return (
        "22:00-06:00" in s
        or "22:00 - 06:00" in s
        or "nightshift22:00-06:00" in s
    )


def is_early_std_night(s):
    s = norm_sched(s)
    return "20:00-06:00" in s or "20:00 - 06:00" in s or "8pm-6am" in s


def minutes_to_timestr(m):
    m = int(round(m)) % (24 * 60)
    h = m // 60
    mi = m % 60
    return f"{h:02d}:{mi:02d}:00"


def get_work_envelope(row):
    TIME_COLS = [1, 2, 3, 4, 5, 6, 7, 8]
    times = []
    for c in TIME_COLS:
        if c < len(row):
            m = to_minutes(row[c])
            if m is not None:
                times.append(m)
    if not times:
        return None, None
    return min(times), max(times)


def parse_sched_window(schedule_str):
    s = str(schedule_str)
    m = re.search(r"(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})", s)
    if not m:
        return None, None
    start_min = to_minutes(m.group(1))
    end_min = to_minutes(m.group(2))
    if start_min is None or end_min is None:
        return None, None
    return start_min, end_min


# -------------------- CLEAN BIRD BREAKS --------------------

def clean_bird_breaks(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    TIME_COLS = [1, 2, 3, 4, 5, 6, 7, 8]

    for idx, row in df.iterrows():
        if len(row) <= 10:
            continue
        schedule = row[10]

        is_day = is_bird_day(schedule)
        is_night = is_bird_night(schedule)
        if not (is_day or is_night):
            continue

        sched_start, sched_end = parse_sched_window(schedule)
        if sched_start is None or sched_end is None:
            continue

        env_start, env_end = get_work_envelope(row)
        if env_start is None or env_end is None:
            continue

        if is_night and env_end <= env_start:
            env_end += 24 * 60

        envelope_min = env_end - env_start
        envelope_hours = envelope_min / 60.0
        total_hours = as_float(row[9]) if len(row) > 9 else 0.0

        if total_hours <= 0:
            continue

        break_hours = envelope_hours - total_hours

        if is_day:
            if not (9.0 <= total_hours <= 11.5):
                continue
        else:
            if not (10.0 <= total_hours <= 11.5):
                continue

        if break_hours <= 0.25 or break_hours > 2.0:
            continue

        break_min = break_hours * 60.0

        if is_day:
            pivot = 14 * 60
        else:
            shift_span_min = (sched_end - sched_start) % (24 * 60)
            pivot = sched_start + shift_span_min / 2.0
            if pivot < env_start:
                pivot += 24 * 60

        min_center = env_start + break_min / 2.0
        max_center = env_end - break_min / 2.0
        center = max(min(pivot, max_center), min_center)

        break_start = center - break_min / 2.0
        break_end = center + break_min / 2.0

        df.at[idx, 1] = minutes_to_timestr(env_start)
        df.at[idx, 2] = minutes_to_timestr(break_start)
        df.at[idx, 3] = minutes_to_timestr(break_end)
        df.at[idx, 4] = minutes_to_timestr(env_end)

        for c in TIME_COLS[4:]:
            if c < len(df.columns):
                df.at[idx, c] = ""

    return df


# -------------------- SHIFT RECONSTRUCTION --------------------

def repair_segments(day_row):
    spans = []
    offset = 0
    last_end = None

    for i in range(4):
        cin = to_minutes(day_row[1 + 2 * i])
        cout = to_minutes(day_row[2 + 2 * i])
        if cin is None or cout is None:
            continue

        cin = snap(cin)
        cout = snap(cout)
        start = cin + offset
        end = cout + offset

        if end <= start:
            end += 1440

        if last_end is not None and start <= last_end:
            shift = ((last_end - start) // 1440) + 1
            start += 1440 * shift
            end += 1440 * shift
            offset += 1440 * shift

        dur = end - start
        if dur > MAX_SPAN_MIN or dur <= MIN_BREAK_MIN:
            continue

        spans.append((start, end))
        last_end = end

    if not spans:
        return spans

    spans.sort()
    merged = [spans[0]]
    for s, e in spans[1:]:
        ps, pe = merged[-1]
        if s - pe <= MERGE_GAP_MIN:
            merged[-1] = (ps, max(pe, e))
        else:
            merged.append((s, e))
    return merged


def get_sched_start_minutes(schedule_str):
    s = norm_sched(schedule_str)
    if "06:00" in s or "6:00" in s:
        return 6 * 60
    if "08:00" in s or "8:00" in s:
        return 8 * 60
    if "14:00" in s or "2pm" in s:
        return 14 * 60
    if "20:00" in s or "8pm" in s:
        return 20 * 60
    if "22:00" in s or "10pm" in s:
        return 22 * 60
    return None


def apply_early_start_rule(segments, schedule_str):
    start_min = get_sched_start_minutes(schedule_str)
    if start_min is None or not segments:
        return segments
    threshold = start_min - 59
    adjusted = []
    for s, e in segments:
        local_start = s % 1440
        if threshold < local_start < start_min:
            shift = start_min - local_start
            s = s + shift
            if s >= e:
                continue
        adjusted.append((s, e))
    return adjusted


def overlap_hours(segments, start_min, end_min):
    tot = 0
    for s, e in segments:
        cur = s
        while cur < e:
            local = cur % 1440
            if local < start_min:
                nxt = min(e, cur - local + start_min)
            elif local >= end_min:
                nxt = min(e, cur - local + 1440 + start_min)
            else:
                nxt = min(e, cur - local + end_min)
                tot += nxt - cur
            cur = nxt
    return tot / 60.0


def window_accumulate(segments):
    mins = {"AM": 0, "PM": 0, "NIGHT": 0}
    for s, e in segments:
        cur = s
        while cur < e:
            local = cur % 1440
            if AM_START <= local < AM_END:
                boundary = cur - local + AM_END
                key = "AM"
            elif PM_START <= local < PM_END:
                boundary = cur - local + PM_END
                key = "PM"
            else:
                key = "NIGHT"
                boundary = cur - local + (
                    AM_START if local < AM_START else 1440 + AM_START
                )
            nxt = min(boundary, e)
            mins[key] += nxt - cur
            cur = nxt
    total = sum(mins.values())
    return (
        mins["AM"] / 60.0,
        mins["PM"] / 60.0,
        mins["NIGHT"] / 60.0,
        total / 60.0,
    )


def distribute_day(day_row, daily_total, schedule_str):
    s = schedule_str or ""
    tot = daily_total
    if tot <= 0:
        return 0.0, 0.0, 0.0, None

    if is_am_day(s) or is_pm_day(s):
        shift_type = "am_day" if is_am_day(s) else "pm_day"
        segs = repair_segments(day_row)
        segs = apply_early_start_rule(segs, s)
        if segs:
            am_w, pm_w, nt_w, used_h = window_accumulate(segs)
            if used_h > 0 and abs(used_h - tot) > 0.01:
                scale = tot / used_h
                am_w *= scale
                pm_w *= scale
                nt_w *= scale
            return am_w, pm_w, nt_w, shift_type
        return (
            (tot, 0.0, 0.0, shift_type)
            if shift_type == "am_day"
            else (0.0, tot, 0.0, shift_type)
        )

    if is_bird_day(s):
        segs = repair_segments(day_row)
        segs = apply_early_start_rule(segs, s)
        if segs:
            am_half = overlap_hours(segs, BIRD_AM_START, BIRD_AM_END)
            pm_half = overlap_hours(segs, BIRD_PM_START, BIRD_PM_END)
            used = am_half + pm_half
            if used > 0 and abs(used - tot) > 0.01:
                scale = tot / used
                am_half *= scale
                pm_half *= scale
            return am_half, pm_half, 0.0, "bird_day"
        return tot / 2.0, tot / 2.0, 0.0, "bird_day"

    if is_std_night(s):
        s_norm = norm_sched(s)
        if is_22_06_night(s_norm):
            segs = repair_segments(day_row)
            segs = apply_early_start_rule(segs, s)
            if segs:
                pm_overlap = overlap_hours(segs, 20 * 60, 22 * 60)
                pm = min(1.0, pm_overlap, tot)
                night = max(0.0, tot - pm)
                return 0.0, pm, night, "std_night_22_06"
            return 0.0, 0.0, tot, "std_night_22_06"

        segs = repair_segments(day_row)
        segs = apply_early_start_rule(segs, s)
        pm_cap = 2.0 if is_early_std_night(s) else 1.0
        if segs:
            pm_overlap = overlap_hours(segs, 20 * 60, 22 * 60)
            pm = min(pm_cap, pm_overlap, tot)
            night = max(0.0, tot - pm)
            return 0.0, pm, night, "std_night"
        pm = min(pm_cap, tot)
        return 0.0, pm, max(0.0, tot - pm), "std_night"

    if is_bird_night(s):
        segs = repair_segments(day_row)
        segs = apply_early_start_rule(segs, s)
        if segs:
            pm = overlap_hours(segs, 20 * 60, 22 * 60)
            am = overlap_hours(segs, 6 * 60, 8 * 60)
            pm = min(2.0, pm, tot)
            rem = max(0.0, tot - pm)
            am = min(2.0, rem)
            night = max(0.0, tot - pm - am)
            return am, pm, night, "bird_day"
        pm = min(2.0, tot)
        rem = max(0.0, tot - pm)
        am = min(2.0, rem)
        night = max(0.0, tot - pm - am)
        return am, pm, night, "bird_day"

    segs = repair_segments(day_row)
    segs = apply_early_start_rule(segs, s)
    if not segs:
        return 0.0, 0.0, 0.0, None
    am_w, pm_w, nt_w, used_h = window_accumulate(segs)
    if used_h > 0 and abs(used_h - tot) > 0.01:
        scale = tot / used_h
        am_w *= scale
        pm_w *= scale
        nt_w *= scale
    return am_w, pm_w, nt_w, "other"


def apply_weekly(dailies):
    if not dailies:
        return (0, 0, 0, 0, False, 0, 0, 0)

    pm_22 = sum(d["PM"] for d in dailies if d["type"] == "std_night_22_06")
    if pm_22 > 1.0:
        over = pm_22 - 1.0
        for d in dailies:
            if d["type"] == "std_night_22_06" and over > 1e-6:
                reduce = min(d["PM"], over)
                d["PM"] -= reduce
                d["NIGHT"] += reduce
                over -= reduce

    total_am = sum(d["AM"] for d in dailies)
    total_pm = sum(d["PM"] for d in dailies)
    total_nt = sum(d["NIGHT"] for d in dailies)
    total = total_am + total_pm + total_nt

    any_bird = any(d["type"] == "bird_day" for d in dailies)
    base_limit = 55.0 if any_bird else 37.5
    bird_flag = any_bird and (total > 22.0)

    remaining = base_limit
    base_week = {"AM": 0.0, "PM": 0.0, "NIGHT": 0.0}
    ot_week = {"AM": 0.0, "PM": 0.0, "NIGHT": 0.0}

    for d in sorted(dailies, key=lambda x: x["date"]):
        am = d["AM"]
        pm = d["PM"]
        nt = d["NIGHT"]
        day_total = am + pm + nt
        if day_total <= 0:
            continue

        base_for_day = min(day_total, remaining) if remaining > 0 else 0.0
        ot_for_day = max(0.0, day_total - base_for_day)

        if base_for_day > 0:
            factor = base_for_day / day_total
            base_week["AM"] += am * factor
            base_week["PM"] += pm * factor
            base_week["NIGHT"] += nt * factor
            remaining -= base_for_day

        if ot_for_day > 0:
            factor_ot = ot_for_day / day_total
            ot_week["AM"] += am * factor_ot
            ot_week["PM"] += pm * factor_ot
            ot_week["NIGHT"] += nt * factor_ot

    am_w = round_quarter(base_week["AM"])
    pm_w = round_quarter(base_week["PM"])
    nt_w = round_quarter(base_week["NIGHT"])
    am_ot = round_quarter(ot_week["AM"])
    pm_ot = round_quarter(ot_week["PM"])
    nt_ot = round_quarter(ot_week["NIGHT"])
    total_ot = round_quarter(am_ot + pm_ot + nt_ot)

    return am_w, pm_w, nt_w, total_ot, bird_flag, am_ot, pm_ot, nt_ot


def process_timesheet(df):
    out = []
    i = 0
    while i < len(df):
        row = df.iloc[i]
        if str(row[0]) != "Employee ID:":
            i += 1
            continue

        name = str(row[10]) if len(row) > 10 else ""
        i += 3

        daily_rows = []
        for _ in range(7):
            if i < len(df):
                daily_rows.append(df.iloc[i])
                i += 1

        total_hours = 0.0
        j = i
        while j < len(df):
            chk = df.iloc[j]
            if len(chk) > 9:
                v = as_float(chk[9])
                if v > 0:
                    total_hours = v
                    break
            if str(chk[0]) == "Employee ID:":
                break
            j += 1

        if total_hours == 0.0:
            continue

        dailies = []
        for dr in daily_rows:
            date_s = str(dr[0]).strip()
            if not date_s:
                continue
            date = datetime.strptime(date_s, "%a %d/%m/%Y")
            sched = str(dr[10]) if len(dr) > 10 else ""
            day_tot = as_float(dr[9])
            if day_tot == 0.0:
                continue

            am, pm, nt, typ = distribute_day(dr, day_tot, sched)
            dailies.append(
                {
                    "date": date,
                    "AM": am,
                    "PM": pm,
                    "NIGHT": nt,
                    "type": typ if typ else "other",
                }
            )

        am_w, pm_w, nt_w, ot_w, bird_flag, am_ot, pm_ot, nt_ot = apply_weekly(dailies)

        any_bird_week = any(d["type"] == "bird_day" for d in dailies)
        total_hours_q = round_quarter(total_hours)
        issue_flag = "Y" if (ot_w > 0 or bird_flag) else ""
        bird_shift_flag = "Y" if any_bird_week else ""

        out.append(
            {
                "CANDIDATE NAME": name,
                "TOTAL HOURS": total_hours_q,
                "AM HOURS": am_w,
                "PM HOURS": pm_w,
                "AM OVERTIME": am_ot,
                "PM OVERTIME": pm_ot,
                "NIGHTS HOURS": nt_w,
                "NIGHTS OVERTIME": nt_ot,
                "TOTAL OVERTIME": ot_w,
                "FLAG WITH POTENTIAL ISSUE": issue_flag,
                "BIRD SHIFT": bird_shift_flag,
                "HOLIDAYS": 0.0,
            }
        )

        i = max(i, j)

    cols = [
        "CANDIDATE NAME",
        "TOTAL HOURS",
        "AM HOURS",
        "PM HOURS",
        "AM OVERTIME",
        "PM OVERTIME",
        "NIGHTS HOURS",
        "NIGHTS OVERTIME",
        "TOTAL OVERTIME",
        "FLAG WITH POTENTIAL ISSUE",
        "BIRD SHIFT",
        "HOLIDAYS",
    ]
    return pd.DataFrame(out)[cols]


# -------------------- API MODELS & ENDPOINTS --------------------

class ProcessRequest(BaseModel):
    # Match what Blocks sends: { csvFileUrl, runName }
    csvFileUrl: str
    runName: str | None = None


@app.post("/process-payroll")
async def process_payroll(request: ProcessRequest):
    csv_url = request.csvFileUrl

    try:
        # Download CSV
        response = requests.get(csv_url, timeout=30)
        response.raise_for_status()

        # Save to temp file
        with tempfile.NamedTemporaryFile(
            mode="wb", delete=False, suffix=".csv"
        ) as tmp:
            tmp.write(response.content)
            csv_path = tmp.name

        # Step 1: Clean bird breaks
        raw_df = pd.read_csv(csv_path, header=None)
        cleaned_df = clean_bird_breaks(raw_df)

        # Step 2: Process timesheet
        result_df = process_timesheet(cleaned_df)

        # Step 3: Generate Excel output
        output_path = tempfile.mktemp(suffix=".xlsx")
        result_df.to_excel(output_path, index=False, sheet_name="Noatum Hours")

        # Read output file as base64
        with open(output_path, "rb") as f:
            file_content = base64.b64encode(f.read()).decode()

        # Cleanup temp files
        os.unlink(csv_path)
        os.unlink(output_path)

        # Calculate summary
        total_workers = len(result_df)
        total_hours = result_df["TOTAL HOURS"].sum()
        bird_workers = (result_df["BIRD SHIFT"] == "Y").sum()
        exceptions = (result_df["FLAG WITH POTENTIAL ISSUE"] == "Y").sum()

        return {
            "success": True,
            "summary": {
                "totalWorkers": int(total_workers),
                "totalHours": round(float(total_hours), 2),
                "birdShiftWorkers": int(bird_workers),
                "totalExceptions": int(exceptions),
                "runName": request.runName,
            },
            "exportFileBase64": file_content,
            "exportFileName": "Noatum_Timesheet_Output.xlsx",
        }

    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"{str(e)}\n{traceback.format_exc()}",
        )


@app.get("/health")
async def health():
    return {"status": "healthy"}


@app.get("/")
async def root():
    return {"message": "Noatum Payroll API", "status": "running"}
