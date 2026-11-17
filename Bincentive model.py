"""
Quality-Based Incentive Model — HBL Recovery
=============================================
Weighted average quality model.

Quality per case  = Cumulative Recovery / Original Outstanding
Officer quality   = Total Cumulative / Total Outstanding  (weighted by outstanding)
Adjusted Recovery = Monthly Recovery x Quality Multiplier
Incentive         = Proposed slab applied to Adjusted Recovery

Three models are compared side by side:
    Existing    — existing slab on raw recovery
    Proposal A  — QUALITY_SLABS + PROPOSED_INCENTIVE_SLABS
    Proposal B  — ALT_QUALITY_SLABS + ALT_INCENTIVE_SLABS (alternative figures)

Eligibility is always tested on RAW recovery (>= MIN_RECOVERY_FOR_INCENTIVE).
Adjusted recovery may fall below that figure once the multiplier is applied.

Outputs:
    incentive_dashboard.html            executive dashboard
    incentive_comparison_output.xlsx    detailed workbook

Usage:
    python incentive_model.py "recovery_data.xlsx"

Requires:
    pip install pandas openpyxl
"""

import pandas as pd
import numpy as np
import sys, os, json, re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ==============================================================
# CONFIGURATION
# ==============================================================

MONTH_ORDER = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
    'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5,
    'June': 6, 'July': 7, 'August': 8,
    'September': 9, 'October': 10, 'November': 11, 'December': 12,
}

# (min %, max %, multiplier, tier name)
QUALITY_SLABS = [
    (0,  10,  0.85,  "Needs Improvement"),
    (11, 20,  0.90,  "Low"),
    (21, 30,  0.95,  "Below Average"),
    (31, 40,  1.00,  "Base"),
    (41, 50,  1.05,  "Satisfactory"),
    (51, 60,  1.10, "Good"),
    (61, 70,  1.15,  "Strong"),
    (71, 80,  1.20,  "Very Strong"),
    (81, 100, 1.25,  "Outstanding"),
]

# (min recovery, max recovery, type, value, label)   type: 'flat' | 'pct'
EXISTING_INCENTIVE_SLABS = [
    (0,      74999,    'flat', 0,     '< 75K'),
    (75000,  124999,   'flat', 10000, '>= 75K - <125K'),
    (125000, 174999,   'pct',  0.10,  '>= 125K - <175K'),
    (175000, 249999,   'pct',  0.13,  '>= 175K - <250K'),
    (250000, 99999999, 'pct',  0.15,  '250K+'),
]
EXISTING_CAP = 200000

PROPOSED_INCENTIVE_SLABS = [
    (0,      74999,    'flat', 0,    '< 75K'),
    (75000,  124999,   'pct',  0.075, '>= 75K - <125K'),
    (125000, 174999,   'pct',  0.10, '>= 125K - <175K'),
    (175000, 249999,   'pct',  0.13, '>= 175K - <250K'),
    (250000, 99999999, 'pct',  0.15, '250K+'),
]
PROPOSED_CAP = 0

# --------------------------------------------------------------
# Comparison set-up: Existing vs Proposal A vs Proposal B
# --------------------------------------------------------------
# Proposal A = QUALITY_SLABS + PROPOSED_INCENTIVE_SLABS above.
# Proposal B = the alternative figures below, with their own multiplier AND
# incentive slab. Rename the labels to whatever you want shown on screen.
LABEL_A = "Proposal A"
LABEL_B = "Proposal B"

# PLACEHOLDER — Proposal B starts as a copy of Proposal A.
# Replace these with the alternative figures, or type them in the dashboard
# (gear icon > Proposal B) and the whole dashboard recalculates.
ALT_QUALITY_SLABS = [
    (0,  10,  0.85,  "Needs Improvement"),
    (11, 20,  0.90,  "Low"),
    (21, 30,  0.95,  "Below Average"),
    (31, 40,  1.00,  "Base"),
    (41, 50,  1.05,  "Satisfactory"),
    (51, 60,  1.10,  "Good"),
    (61, 70,  1.15,  "Strong"),
    (71, 80,  1.20,  "Very Strong"),
    (81, 100, 1.25,  "Outstanding"),
]

ALT_INCENTIVE_SLABS = [
    (0,      74999,    'flat', 0,     '< 75K'),
    (75000,  124999,   'pct',  0.075, '>= 75K - <125K'),
    (125000, 174999,   'pct',  0.10,  '>= 125K - <175K'),
    (175000, 249999,   'pct',  0.13,  '>= 175K - <250K'),
    (250000, 99999999, 'pct',  0.15,  '250K+'),
]
ALT_CAP = 0

# Two guardrails sit either side of the quality multiplier.
#
#   Eligibility floor — an officer whose RAW monthly recovery falls below this
#   earns nothing, so a weak month cannot be lifted over the line by a strong
#   quality score.
#
#   Adjusted-recovery floor — once an officer IS eligible, the multiplier may
#   reduce their adjusted recovery but never below this figure, so poor quality
#   cannot wipe out an incentive that volume has already earned.
#
# Both use the same number: the entry point of the incentive slab.
MIN_RECOVERY_FOR_INCENTIVE = 75000

# Set False to let a weak multiplier carry adjusted recovery below the threshold,
# so the incentive is read off whichever slab band the reduced figure lands in.
APPLY_ADJUSTED_FLOOR = True


def _apply_slab(recovery, slabs, cap):
    for low, high, typ, val, label in slabs:
        if low <= recovery <= high:
            inc = val if typ == 'flat' else round(recovery * val)
            if cap:
                inc = min(inc, cap)
            return inc, label
    _, _, typ, val, label = slabs[-1]
    inc = val if typ == 'flat' else round(recovery * val)
    if cap:
        inc = min(inc, cap)
    return inc, label


def calc_existing_incentive(recovery):
    return _apply_slab(recovery, EXISTING_INCENTIVE_SLABS, EXISTING_CAP)


def calc_proposed_incentive(adjusted_recovery):
    return _apply_slab(adjusted_recovery, PROPOSED_INCENTIVE_SLABS, PROPOSED_CAP)


def calc_alt_incentive(adjusted_recovery):
    return _apply_slab(adjusted_recovery, ALT_INCENTIVE_SLABS, ALT_CAP)


def get_quality_multiplier(quality_pct, slabs=None):
    slabs = slabs or QUALITY_SLABS
    for low, high, mult, tier in slabs:
        if low <= quality_pct <= high:
            return mult, tier
    return slabs[-1][2], slabs[-1][3]


def score_proposal(total_recovery, quality, q_slabs, slab_fn, zero_label):
    """Run one proposal (multiplier + slab) for one officer-month.

    Eligibility is tested on RAW recovery. Adjusted recovery may then come out
    below the threshold; APPLY_ADJUSTED_FLOOR decides whether it is held there.
    """
    multiplier, tier = get_quality_multiplier(round(quality), q_slabs)
    computed = round(total_recovery * multiplier)
    below = total_recovery < MIN_RECOVERY_FOR_INCENTIVE
    floored = (APPLY_ADJUSTED_FLOOR and not below
               and computed < MIN_RECOVERY_FOR_INCENTIVE)
    adjusted = MIN_RECOVERY_FOR_INCENTIVE if floored else computed
    if below:
        inc, slab = 0, zero_label
    else:
        inc, slab = slab_fn(adjusted)
    return {'tier': tier, 'mult': multiplier, 'computed': computed, 'below': below,
            'floored': floored, 'adjusted': adjusted, 'inc': inc, 'slab': slab}


def eligibility_note(q_slabs, slab_fn):
    """Plain-language note shown under each proposed slab."""
    t = MIN_RECOVERY_FOR_INCENTIVE
    lowest = min(s[2] for s in q_slabs)
    example = t + 1000
    computed = round(example * lowest)
    note = (f"Eligibility: RAW (actual) recovery must be at least {t:,}. "
            f"ADJUSTED recovery can be below {t:,}")
    if lowest < 1 and computed < t:
        note += (f" - e.g. {example:,} recovered at a {lowest}x multiplier gives "
                 f"{computed:,}. The officer is still eligible because the test is on raw recovery")
        if APPLY_ADJUSTED_FLOOR:
            note += f"; adjusted recovery is held at {t:,} for the slab ({slab_fn(t)[0]:,} incentive)."
        else:
            note += f"; the incentive is read from the band {computed:,} falls in ({slab_fn(computed)[0]:,})."
    else:
        note += " when the multiplier is under 1x; eligibility is still tested on raw recovery."
    return note


def _norm(val):
    """Normalise an identifier for matching."""
    if val is None or pd.isna(val):
        return None
    s = str(val).strip().upper()
    return None if s in ('', 'NAN', 'NONE', 'NULL', 'NA', '-', '.') else s


def _month_num(val):
    """Month number from 'Jan', 'January', 'Jan-26', '01-Jan-2026' or a real date cell.

    Excel often converts a typed 'Jan-26' into a date, so pandas may hand back a
    Timestamp rather than a string — that case is checked first.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.month
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'nat', 'none'):
        return None
    for token in re.split(r'[-/\s,]+', s):
        hit = MONTH_ORDER.get(token.strip().title())
        if hit:
            return hit
    return None


# ==============================================================
# DATA PROCESSING
# ==============================================================

def process_data(input_file):
    print(f"\n{'=' * 62}\n  QUALITY-BASED INCENTIVE MODEL\n{'=' * 62}")
    print(f"  Reading: {os.path.basename(input_file)}")

    df = pd.read_excel(input_file)
    df.columns = df.columns.str.strip()

    required = ['Recovery Month', 'EmpCode', 'Recovery Officer', 'Recovery',
                'Outstanding', 'AccountRaw', 'PRK', 'Region']
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"  ERROR — missing columns: {missing}")
        sys.exit(1)

    df['MonthNum'] = df['Recovery Month'].map(MONTH_ORDER)
    df = df.sort_values(['EmpCode', 'MonthNum'])

    for col in ['Recovery', 'Outstanding']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    if 'Total Waiver' in df.columns:
        df['Total Waiver'] = pd.to_numeric(df['Total Waiver'], errors='coerce').fillna(0)
    else:
        df['Total Waiver'] = 0

    # ---- Union-find: group rows that share an account or PRK into one case ----
    df['_acc'] = df['AccountRaw'].apply(_norm)
    df['_prk'] = df['PRK'].apply(_norm)

    parent = {}

    def find(x):
        while parent.setdefault(x, x) != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for _, row in df.iterrows():
        if row['_acc'] and row['_prk']:
            union('ACC:' + row['_acc'], 'PRK:' + row['_prk'])

    def root_of(r):
        if r['_acc']:
            return find('ACC:' + r['_acc'])
        if r['_prk']:
            return find('PRK:' + r['_prk'])
        return 'UNKNOWN'

    df['CaseKey'] = df.apply(root_of, axis=1)
    roots = {root: f"CASE-{i + 1:05d}" for i, root in enumerate(df['CaseKey'].unique())}
    df['CaseKey'] = df['CaseKey'].map(roots)

    print(f"  Officers: {df['EmpCode'].nunique():,}   Cases: {df['CaseKey'].nunique():,}   Rows: {len(df):,}")

    # ---- Settlement (closure payment) month per case ----
    closure_map = {}
    if 'Closure Payment Month' in df.columns:
        for _, row in df.iterrows():
            cpm = row.get('Closure Payment Month', '')
            if pd.notna(cpm) and str(cpm).strip():
                mn = _month_num(cpm)
                if mn:
                    closure_map[(row['EmpCode'], row['CaseKey'])] = mn

    # ---- Cumulative recovery and per-case quality, month by month ----
    months_sorted = sorted(df['MonthNum'].dropna().unique())
    records = []

    for month_num in months_sorted:
        upto = df[df['MonthNum'] <= month_num]
        current = df[df['MonthNum'] == month_num]

        for _, row in current.iterrows():
            case_key, emp = row['CaseKey'], row['EmpCode']
            cum = upto.loc[(upto['CaseKey'] == case_key) & (upto['EmpCode'] == emp), 'Recovery'].sum()
            os_val = row['Outstanding']
            quality = min(round((cum / os_val) * 100, 2), 100) if os_val > 0 else 0

            records.append({
                'Recovery Month': row['Recovery Month'], 'MonthNum': month_num,
                'EmpCode': emp, 'Recovery Officer': row['Recovery Officer'],
                'Region': row.get('Region', ''), 'Product': row.get('Product', ''),
                'Loan Category': row.get('Loan Category', ''),
                'CaseKey': case_key, 'AccountRaw': row['AccountRaw'], 'PRK': row['PRK'],
                'Outstanding': os_val,
                'Current Month Recovery': row['Recovery'],
                'Cumulative Recovery': cum,
                'Quality %': quality,
                'Total Waiver': row['Total Waiver'],
                'Settlement Status': row.get('Settlement Vlookup', ''),
            })

    df_cases = pd.DataFrame(records)

    # ---- Officer x month summary ----
    officer_rows = []
    for (emp, month_num), grp in df_cases.groupby(['EmpCode', 'MonthNum']):
        valid = grp[grp['Outstanding'] > 0]

        total_recovery = round(grp['Current Month Recovery'].sum())
        total_cum = valid['Cumulative Recovery'].sum()
        total_os = valid['Outstanding'].sum()

        quality = min(round((total_cum / total_os) * 100, 2), 100) if total_os > 0 else 0

        a = score_proposal(total_recovery, quality, QUALITY_SLABS,
                           calc_proposed_incentive, PROPOSED_INCENTIVE_SLABS[0][4])
        b = score_proposal(total_recovery, quality, ALT_QUALITY_SLABS,
                           calc_alt_incentive, ALT_INCENTIVE_SLABS[0][4])
        multiplier, tier, computed = a['mult'], a['tier'], a['computed']
        below_threshold, floored, adjusted = a['below'], a['floored'], a['adjusted']
        proposed, proposed_slab = a['inc'], a['slab']

        existing, existing_slab = calc_existing_incentive(total_recovery)

        officer_rows.append({
            'Recovery Month': grp['Recovery Month'].iloc[0], 'MonthNum': month_num,
            'EmpCode': emp, 'Recovery Officer': grp['Recovery Officer'].iloc[0],
            'Region': grp['Region'].iloc[0],
            'Total Cases': len(valid),
            'Total Outstanding': round(total_os),
            'Total Recovery': total_recovery,
            'Avg Quality %': quality,
            'Quality Tier': tier, 'Quality Multiplier': multiplier,
            'Computed Recovery': computed,
            'Adjusted Recovery': adjusted,
            'Below Threshold': below_threshold, 'Floored': floored,
            'Existing Slab': existing_slab, 'Existing Incentive': existing,
            'Proposed Slab': proposed_slab, 'Proposed Incentive': proposed,
            'Incentive Difference': proposed - existing,
            'B Tier': b['tier'], 'B Multiplier': b['mult'],
            'B Computed Recovery': b['computed'], 'B Adjusted Recovery': b['adjusted'],
            'B Floored': b['floored'], 'B Slab': b['slab'],
            'B Incentive': b['inc'], 'B Difference': b['inc'] - existing,
        })

    df_officer = pd.DataFrame(officer_rows).sort_values(['MonthNum', 'EmpCode'])

    # ---- Officer totals across all months ----
    overall_rows = []
    for emp, grp in df_officer.groupby('EmpCode'):
        existing = grp['Existing Incentive'].sum()
        proposed = grp['Proposed Incentive'].sum()
        diff = proposed - existing
        alt = grp['B Incentive'].sum()
        diff_b = alt - existing
        overall_rows.append({
            'EmpCode': emp, 'Recovery Officer': grp['Recovery Officer'].iloc[0],
            'Region': grp['Region'].iloc[0],
            'Months Active': grp['MonthNum'].nunique(),
            'Total Cases': grp['Total Cases'].sum(),
            'Total Recovery': grp['Total Recovery'].sum(),
            'Avg Quality %': round(grp['Avg Quality %'].mean(), 2),
            'Total Existing': existing, 'Total Proposed': proposed,
            'Total Difference': diff,
            'Impact': 'Benefited' if diff > 0 else ('Impacted' if diff < 0 else 'No Change'),
            'Total B': alt, 'B Total Difference': diff_b,
            'B Impact': 'Benefited' if diff_b > 0 else ('Impacted' if diff_b < 0 else 'No Change'),
        })

    df_overall = pd.DataFrame(overall_rows).sort_values('Total Difference')

    # ---- Case timelines for the deep dive ----
    case_timeline = []
    for (emp, case_key), grp in df_cases.groupby(['EmpCode', 'CaseKey']):
        grp = grp.sort_values('MonthNum')
        os_val = round(grp['Outstanding'].iloc[0])

        settlement_month = None
        mn = closure_map.get((emp, case_key))
        if mn:
            settlement_month = next((n for n, v in MONTH_ORDER.items() if v == mn and len(n) == 3), None)

        timeline, done = [], False
        for _, r in grp.iterrows():
            if done:
                break
            cum = round(r['Cumulative Recovery'])
            is_settlement = (settlement_month == r['Recovery Month'])
            is_closed = os_val > 0 and round((cum / os_val) * 100) >= 100
            timeline.append({
                'm': r['Recovery Month'],
                'rec': round(r['Current Month Recovery']),
                'cum': cum,
                'q': min(round(float(r['Quality %'])), 100),
                'settlement': is_settlement,
                'closed': is_closed,
            })
            if is_settlement or is_closed:
                done = True

        case_timeline.append({
            'emp': emp, 'case': case_key,
            'account': str(grp['AccountRaw'].iloc[0]),
            'os': os_val, 'timeline': timeline,
        })

    return df, df_cases, df_officer, df_overall, case_timeline


# ==============================================================
# EXCEL OUTPUT
# ==============================================================

def _label_columns(df):
    """Swap internal column names for the on-screen labels of A and B."""
    A, B = LABEL_A, LABEL_B
    fixed = {
        'Quality Tier': f'{A} Tier', 'Quality Multiplier': f'{A} Multiplier',
        'Computed Recovery': f'{A} Computed Recovery',
        'Adjusted Recovery': f'{A} Adjusted Recovery', 'Floored': f'{A} Floored',
        'Proposed Slab': f'{A} Slab', 'Proposed Incentive': f'{A} Incentive',
        'Incentive Difference': f'{A} vs Existing',
        'Total Proposed': f'Total {A}', 'Total Difference': f'{A} vs Existing',
        'Impact': f'{A} Impact',
        'Total B': f'Total {B}', 'B Total Difference': f'{B} vs Existing',
        'B Difference': f'{B} vs Existing',
    }

    def rename(c):
        if c in fixed:
            return fixed[c]
        if c.startswith('B '):
            return f'{B} {c[2:]}'
        return c
    return df.rename(columns=rename)


def monthly_summary(df_officer):
    """Month-wise recovery and incentive under all three models, plus a total row."""
    rows = []
    for _, g in df_officer.groupby('MonthNum'):
        ex, a, b = g['Existing Incentive'].sum(), g['Proposed Incentive'].sum(), g['B Incentive'].sum()
        rows.append({
            'Month': g['Recovery Month'].iloc[0],
            'Officers': g['EmpCode'].nunique(),
            f'Eligible (>= {MIN_RECOVERY_FOR_INCENTIVE:,})': int((~g['Below Threshold']).sum()),
            'Total Recovery': g['Total Recovery'].sum(),
            'Existing Incentive': ex,
            f'{LABEL_A} Incentive': a, f'{LABEL_A} vs Existing': a - ex,
            f'{LABEL_B} Incentive': b, f'{LABEL_B} vs Existing': b - ex,
        })
    df = pd.DataFrame(rows)
    total = {c: df[c].sum() for c in df.columns if c != 'Month'}
    total['Month'] = 'TOTAL'
    total['Officers'] = df_officer['EmpCode'].nunique()
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)[df.columns]


def _write_slab_sheet(wb, head_font, head_fill, head_align, thin):
    """Side-by-side slab comparison, with the eligibility note under each proposal."""
    ws = wb.create_sheet('Slab Comparison', 1)
    title_font = Font(name='Arial', bold=True, size=12, color='00857C')
    note_font = Font(name='Arial', italic=True, size=10, color='7A5C00')
    body_font = Font(name='Arial', size=10)

    def block(row, col, title, headers, data, note=None):
        ws.cell(row, col, title).font = title_font
        for j, h in enumerate(headers):
            c = ws.cell(row + 1, col + j, h)
            c.font, c.fill, c.alignment, c.border = head_font, head_fill, head_align, thin
        for i, rec in enumerate(data):
            for j, v in enumerate(rec):
                c = ws.cell(row + 2 + i, col + j, v)
                c.font, c.border = body_font, thin
        end = row + 2 + len(data)
        if note:
            ws.cell(end, col, note).font = note_font
            ws.cell(end, col).alignment = Alignment(wrap_text=True, vertical='top')
            ws.merge_cells(start_row=end, start_column=col, end_row=end, end_column=col + len(headers) - 1)
            ws.row_dimensions[end].height = 58
            end += 1
        return end + 1

    def inc_rows(slabs, cap):
        out = []
        for i, (lo, hi, typ, val, label) in enumerate(slabs):
            text = ('No Incentive' if val == 0 else f'Flat {val:,}') if typ == 'flat' else f'{val * 100:g}%'
            if cap and i == len(slabs) - 1:
                text += f' (cap {cap:,})'
            out.append((label, text))
        return out

    q_rows = lambda slabs: [(f'{lo}% - {hi}%', f'{m}x', t) for lo, hi, m, t in slabs]

    r = block(1, 1, 'Existing Incentive (on raw recovery)', ['Recovery band', 'Incentive'],
              inc_rows(EXISTING_INCENTIVE_SLABS, EXISTING_CAP))
    rA = block(1, 4, f'{LABEL_A} - Incentive (on adjusted recovery)', ['Adjusted band', 'Incentive'],
               inc_rows(PROPOSED_INCENTIVE_SLABS, PROPOSED_CAP),
               eligibility_note(QUALITY_SLABS, calc_proposed_incentive))
    rB = block(1, 7, f'{LABEL_B} - Incentive (on adjusted recovery)', ['Adjusted band', 'Incentive'],
               inc_rows(ALT_INCENTIVE_SLABS, ALT_CAP),
               eligibility_note(ALT_QUALITY_SLABS, calc_alt_incentive))
    start = max(r, rA, rB) + 1
    block(start, 4, f'{LABEL_A} - Quality Multiplier', ['Quality', 'Multiplier', 'Tier'], q_rows(QUALITY_SLABS))
    block(start, 8, f'{LABEL_B} - Quality Multiplier', ['Quality', 'Multiplier', 'Tier'], q_rows(ALT_QUALITY_SLABS))

    for col, w in zip('ABCDEFGHIJ', [22, 22, 3, 22, 22, 22, 22, 22, 14, 20]):
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False


def write_excel(input_file, df_cases, df_officer, df_overall):
    out_path = os.path.join(os.path.dirname(os.path.abspath(input_file)),
                            'incentive_comparison_output.xlsx')

    total_existing = df_officer['Existing Incentive'].sum()
    total_a = df_officer['Proposed Incentive'].sum()
    total_b = df_officer['B Incentive'].sum()

    df_summary = pd.DataFrame({
        'Metric': ['Total Officers', 'Total Recovery', 'Existing Incentive',
                   f'{LABEL_A} Incentive', f'{LABEL_A} vs Existing',
                   f'{LABEL_A} - Benefited', f'{LABEL_A} - Impacted',
                   f'{LABEL_B} Incentive', f'{LABEL_B} vs Existing',
                   f'{LABEL_B} - Benefited', f'{LABEL_B} - Impacted',
                   'Avg Quality %', 'Eligibility rule'],
        'Value': [len(df_overall), f"{df_officer['Total Recovery'].sum():,.0f}",
                  f"{total_existing:,.0f}",
                  f"{total_a:,.0f}", f"{total_a - total_existing:+,.0f}",
                  int((df_overall['Impact'] == 'Benefited').sum()),
                  int((df_overall['Impact'] == 'Impacted').sum()),
                  f"{total_b:,.0f}", f"{total_b - total_existing:+,.0f}",
                  int((df_overall['B Impact'] == 'Benefited').sum()),
                  int((df_overall['B Impact'] == 'Impacted').sum()),
                  f"{df_overall['Avg Quality %'].mean():.1f}%",
                  f"Raw recovery must be >= {MIN_RECOVERY_FOR_INCENTIVE:,}; "
                  f"adjusted recovery may fall below it (see Slab Comparison)"],
    })

    month_cols = ['EmpCode', 'Recovery Officer', 'Region', 'Total Cases',
                  'Total Outstanding', 'Total Recovery', 'Avg Quality %', 'Below Threshold',
                  'Existing Slab', 'Existing Incentive',
                  'Quality Tier', 'Quality Multiplier', 'Computed Recovery',
                  'Adjusted Recovery', 'Floored', 'Proposed Slab', 'Proposed Incentive',
                  'Incentive Difference',
                  'B Tier', 'B Multiplier', 'B Computed Recovery', 'B Adjusted Recovery',
                  'B Floored', 'B Slab', 'B Incentive', 'B Difference']

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Executive Summary', index=False)
        monthly_summary(df_officer).to_excel(writer, sheet_name='Monthly Summary', index=False)
        _label_columns(df_overall).to_excel(writer, sheet_name='Officer Overall', index=False)
        _label_columns(df_officer).to_excel(writer, sheet_name='Officer Monthly Detail', index=False)

        for month_num in sorted(df_officer['MonthNum'].unique()):
            dm = df_officer[df_officer['MonthNum'] == month_num]
            label = dm['Recovery Month'].iloc[0]
            cols = [c for c in month_cols if c in dm.columns]
            sheet = dm[cols].sort_values('Total Recovery', ascending=False).copy()
            sheet.insert(0, 'Rank', range(1, len(sheet) + 1))
            _label_columns(sheet).to_excel(writer, sheet_name=f'{label} Detail', index=False)

        df_cases.to_excel(writer, sheet_name='Case Level Detail', index=False)

    # ---- formatting ----
    wb = load_workbook(out_path)
    head_font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
    head_fill = PatternFill('solid', start_color='00857C', end_color='00857C')
    head_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin = Border(*[Side(style='thin', color='D9D9D9')] * 4)
    green = PatternFill('solid', start_color='E2EFDA', end_color='E2EFDA')
    red = PatternFill('solid', start_color='FCE4EC', end_color='FCE4EC')

    for name in wb.sheetnames:
        ws = wb[name]
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(1, col)
            cell.font, cell.fill, cell.alignment, cell.border = head_font, head_fill, head_align, thin
            width = max((len(str(ws.cell(r, col).value or ''))
                         for r in range(1, min(ws.max_row + 1, 80))), default=8)
            ws.column_dimensions[get_column_letter(col)].width = min(width + 4, 30)
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row, col).border = thin
                ws.cell(row, col).font = Font(name='Arial', size=10)
                v = ws.cell(row, col).value
                if (isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) >= 1000
                        and ws.cell(1, col).value not in ('EmpCode', 'Rank')):
                    ws.cell(row, col).number_format = '#,##0'
        ws.freeze_panes = 'A2'

    ws = wb['Executive Summary']
    ws.column_dimensions['B'].width = 70

    # Total row of the monthly summary in bold
    ws = wb['Monthly Summary']
    for col in range(1, ws.max_column + 1):
        c = ws.cell(ws.max_row, col)
        c.font = Font(name='Arial', size=10, bold=True)
        c.fill = PatternFill('solid', start_color='E0F5F3', end_color='E0F5F3')

    def shade(ws, header, positive=green, negative=red, match=None):
        target = next((c for c in range(1, ws.max_column + 1)
                       if ws.cell(1, c).value == header), None)
        if not target:
            return
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row, target).value
            fill = None
            if match:
                fill = positive if val == match[0] else (negative if val == match[1] else None)
            elif isinstance(val, (int, float)):
                fill = positive if val > 0 else (negative if val < 0 else None)
            if fill:
                ws.cell(row, target).fill = fill

    shade(wb['Officer Overall'], f'{LABEL_A} Impact', match=('Benefited', 'Impacted'))
    shade(wb['Officer Overall'], f'{LABEL_B} Impact', match=('Benefited', 'Impacted'))
    for name in wb.sheetnames:
        if name.endswith(' Detail') and name not in ('Officer Monthly Detail', 'Case Level Detail'):
            shade(wb[name], f'{LABEL_A} vs Existing')
            shade(wb[name], f'{LABEL_B} vs Existing')
    shade(wb['Monthly Summary'], f'{LABEL_A} vs Existing')
    shade(wb['Monthly Summary'], f'{LABEL_B} vs Existing')

    _write_slab_sheet(wb, head_font, head_fill, head_align, thin)

    wb.save(out_path)
    print(f"  Excel saved:     {os.path.basename(out_path)}")
    return out_path


# ==============================================================
# HTML DASHBOARD
# ==============================================================

DASHBOARD_CSS = """
:root{
  --teal:#00857C; --teal-dark:#006B63; --teal-light:#E0F5F3;
  --gold:#C5A236; --gold-light:#FBF5E4;
  --red:#D64545;  --red-light:#FDE8E8;
  --green:#2E7D32;--green-light:#E8F5E9;
  --bg:#F7FAFA;   --card:#FFFFFF; --text:#1A2B2A;
  --muted:#5F7A78;--border:#D4E4E2;
  --amber:#B7791F;--amber-light:#FDF3E1;
}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);line-height:1.5}

.header{background:var(--teal);padding:18px 32px;display:flex;align-items:center;gap:16px}
.header-logo{width:42px;height:42px;border-radius:8px;background:rgba(255,255,255,.15);
  display:flex;align-items:center;justify-content:center;font-weight:800;font-size:15px;color:#fff}
.header-title{font-size:18px;font-weight:600;color:#fff}
.header-right{margin-left:auto;font-size:12px;color:rgba(255,255,255,.65)}
.gear-btn{background:rgba(255,255,255,.15);border:none;color:#fff;width:36px;height:36px;
  border-radius:8px;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;margin-left:12px}
.gear-btn:hover{background:rgba(255,255,255,.3)}

.tab-bar{display:flex;padding:0 32px;background:var(--card);border-bottom:1px solid var(--border);overflow-x:auto}
.tab{padding:12px 24px;border:none;background:none;cursor:pointer;font-size:14px;color:var(--muted);
  border-bottom:3px solid transparent;white-space:nowrap;font-family:inherit}
.tab:hover{color:var(--teal)}
.tab.active{color:var(--teal);font-weight:600;border-bottom-color:var(--teal)}

.content{padding:24px 32px;max-width:1600px;margin:0 auto}
.tab-panel{display:none}
.tab-panel.active{display:block}
h2{font-size:20px;color:var(--teal);margin-bottom:20px;font-weight:600}
h3{font-size:15px;margin-bottom:12px;font-weight:600}

.stats{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:24px}
.stat{background:var(--card);border:1px solid var(--border);border-radius:12px;
  padding:16px 20px;flex:1 1 150px;min-width:130px}
.stat-label{font-size:12px;color:var(--muted);margin-bottom:4px}
.stat-value{font-size:20px;font-weight:700;white-space:nowrap}
.stat-sub{font-size:11px;color:var(--muted);margin-top:4px}

.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:24px;margin-bottom:24px}
.card-title{font-size:14px;color:var(--muted);margin-bottom:16px;font-weight:500}
.flex-row{display:flex;gap:24px;flex-wrap:wrap}
.flex-1{flex:1 1 320px}

table{width:100%;border-collapse:collapse;font-size:13px}
thead tr{background:var(--teal);color:#fff}
th{padding:10px 12px;font-weight:500;font-size:12px;white-space:nowrap;text-align:left}
td{padding:10px 12px;border-bottom:1px solid var(--border);white-space:nowrap}
.text-right{text-align:right}
.row-green{background:var(--green-light)}
.row-red{background:var(--red-light)}
.overflow-x{overflow-x:auto}

.badge{display:inline-block;padding:3px 10px;border-radius:6px;font-size:11px;font-weight:600}
.badge-green{color:var(--green);background:var(--green-light)}
.badge-red{color:var(--red);background:var(--red-light)}
.badge-gold{color:var(--gold);background:var(--gold-light)}
.badge-muted{color:var(--muted);background:var(--bg)}

.btn{padding:6px 16px;border-radius:6px;border:1px solid var(--border);background:var(--card);
  color:var(--text);cursor:pointer;font-size:13px;font-weight:500;font-family:inherit}
.btn.active{background:var(--teal);color:#fff;border-color:var(--teal)}
.btn:hover:not(.active){border-color:var(--teal);color:var(--teal)}

.filter-bar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:16px}
.filter-input{padding:8px 14px;border:1px solid var(--border);border-radius:8px;
  font-size:13px;min-width:240px;outline:none;font-family:inherit}
.filter-input:focus{border-color:var(--teal)}

.search-select{position:relative;display:inline-block;min-width:360px}
.search-select input{width:100%;padding:10px 14px;border:1px solid var(--border);
  border-radius:8px;font-size:14px;outline:none;font-family:inherit}
.search-select input:focus{border-color:var(--teal)}
.search-list{position:absolute;top:100%;left:0;right:0;background:var(--card);
  border:1px solid var(--border);border-top:none;border-radius:0 0 8px 8px;
  max-height:300px;overflow-y:auto;z-index:100;display:none;box-shadow:0 8px 24px rgba(0,0,0,.1)}
.search-list.open{display:block}
.search-item{padding:10px 14px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--border)}
.search-item:last-child{border-bottom:none}
.search-item:hover{background:var(--teal-light);color:var(--teal-dark)}

.officer-summary{display:flex;gap:28px;flex-wrap:wrap;background:var(--card);
  border:1px solid var(--border);border-radius:12px;padding:18px 24px;margin-bottom:20px}
.officer-summary-item{min-width:110px}
.officer-summary-label{font-size:11px;color:var(--muted);margin-bottom:3px}
.officer-summary-value{font-size:18px;font-weight:700}

.month-header{background:var(--teal);color:#fff;padding:10px 16px;border-radius:8px 8px 0 0;
  font-size:15px;font-weight:600;margin-top:20px}
.month-summary{background:var(--teal-light);padding:12px 16px;display:flex;gap:28px;flex-wrap:wrap;font-size:13px}
.month-summary-item{min-width:100px}
.month-summary-label{font-size:11px;color:var(--teal-dark);margin-bottom:2px}
.month-summary-value{font-size:15px;font-weight:700;color:var(--teal-dark)}
.case-table-wrap{background:var(--card);border:1px solid var(--border);border-top:none;
  border-radius:0 0 8px 8px;overflow:hidden;margin-bottom:16px}

.step-row{display:flex;gap:14px;padding:12px 0;border-bottom:1px solid var(--border);align-items:flex-start}
.step-row:last-child{border-bottom:none}
.step-num{width:28px;height:28px;border-radius:50%;background:var(--teal);color:#fff;
  display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;flex-shrink:0}
.step-title{font-size:14px;font-weight:600}
.step-desc{font-size:13px;color:var(--muted)}

.slab-table td{padding:8px 12px;border-bottom:1px solid var(--border);font-size:13px}
.slab-table tr:last-child td{border-bottom:none}

.slider-row{margin-bottom:16px}
.slider-row label{display:flex;justify-content:space-between;font-size:13px;color:var(--muted);margin-bottom:6px}
.slider-row label span:last-child{font-weight:600;color:var(--text)}
input[type=range]{width:100%;accent-color:var(--teal)}
.sim-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
.sim-box{padding:16px;border-radius:8px;text-align:center;background:var(--bg)}
.sim-box-red{background:var(--red-light)}
.sim-box-green{background:var(--green-light)}
.sim-label{font-size:12px;color:var(--muted)}
.sim-value{font-size:20px;font-weight:700;margin-top:2px}
.impact-box{padding:12px 16px;border-radius:8px;text-align:center;font-weight:600;font-size:16px;margin-top:16px}

.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:1000;
  justify-content:center;align-items:flex-start;padding:40px 20px;overflow-y:auto}
.modal-overlay.open{display:flex}
.modal{background:var(--card);border-radius:16px;width:100%;max-width:900px;
  max-height:90vh;overflow-y:auto;padding:32px;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal h2{margin-bottom:6px}
.modal-intro{color:var(--muted);font-size:13px;margin-bottom:20px}
.slab-section{background:var(--bg);border:1px solid var(--border);border-radius:10px;padding:16px;margin-bottom:20px}
.slab-section h3{font-size:14px;color:var(--teal-dark);margin-bottom:10px}
.slab-edit{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:12px}
.slab-edit th{background:var(--teal);color:#fff;padding:8px 10px;font-size:11px;font-weight:500;text-align:center}
.slab-edit td{padding:6px 8px;border-bottom:1px solid var(--border)}
.slab-edit input{width:100%;padding:6px 8px;border:1px solid var(--border);
  border-radius:4px;font-size:13px;text-align:center;font-family:inherit}
.slab-edit select{padding:5px 6px;border:1px solid var(--border);border-radius:4px;font-size:12px;font-family:inherit}
.cap-row{display:flex;gap:12px;align-items:center;font-size:13px}
.cap-row input{width:130px;padding:6px 8px;border:1px solid var(--border);border-radius:4px;
  font-size:13px;font-family:inherit}
.add-btn{padding:7px 16px;border:1px dashed var(--teal);border-radius:6px;background:none;
  color:var(--teal);cursor:pointer;font-size:13px;font-weight:500;font-family:inherit}
.remove-btn{padding:4px 9px;border:1px solid var(--border);border-radius:4px;
  background:var(--red-light);color:var(--red);cursor:pointer;font-size:12px;font-family:inherit}
.modal-actions{display:flex;gap:12px;justify-content:flex-end;margin-top:20px;
  padding-top:16px;border-top:1px solid var(--border)}
.btn-apply{padding:10px 28px;background:var(--teal);color:#fff;border:none;border-radius:8px;
  font-size:14px;font-weight:600;cursor:pointer;font-family:inherit}
.btn-apply:hover{background:var(--teal-dark)}
.btn-reset{padding:10px 28px;background:var(--red-light);color:var(--red);border:1px solid var(--red);
  border-radius:8px;font-size:14px;cursor:pointer;font-family:inherit}
.btn-close{padding:10px 28px;background:var(--bg);color:var(--muted);border:1px solid var(--border);
  border-radius:8px;font-size:14px;cursor:pointer;font-family:inherit}

.toast{position:fixed;top:20px;right:20px;background:var(--green);color:#fff;padding:12px 24px;
  border-radius:8px;font-size:14px;font-weight:500;z-index:2000;display:none;box-shadow:0 4px 20px rgba(0,0,0,.2)}

.chart-container{position:relative;height:240px}

/* Three-way comparison: Existing / Proposal A / Proposal B */
thead th.th-ex{background:#5F7A78}
thead th.th-a{background:var(--teal)}
thead th.th-b{background:var(--amber)}
thead tr.grp th{padding:6px 12px;font-size:11px;font-weight:700;letter-spacing:.04em;text-transform:uppercase}
td.bl{border-left:2px solid var(--border)}
th.bl{border-left:2px solid rgba(255,255,255,.45)}
tfoot td{border-top:2px solid var(--teal)}
.card-head{padding:14px 20px;border-bottom:1px solid var(--border);font-size:14px;color:var(--muted);font-weight:500}
.note-box{margin-top:12px;padding:10px 14px;border-left:3px solid var(--gold);background:var(--gold-light);
  border-radius:6px;font-size:12.5px;line-height:1.55;color:var(--text)}
.banner{display:none;margin-bottom:18px;padding:12px 16px;border-radius:8px;background:var(--amber-light);
  border:1px solid var(--amber);color:#7A5000;font-size:13px}
.modal-group{border:2px solid var(--border);border-radius:12px;padding:14px 14px 0;margin-bottom:20px}
.modal-group-title{font-size:15px;font-weight:700;margin-bottom:10px}

@media(max-width:768px){
  .content{padding:16px}
  .header{padding:14px 16px}
  .tab-bar{padding:0 16px}
  .tab{padding:10px 14px;font-size:13px}
  .sim-grid{grid-template-columns:1fr}
}
"""


DASHBOARD_JS = """
const fmt = n => Math.round(n ?? 0).toLocaleString('en-PK');
const sfmt = d => (d > 0 ? '+' : '') + fmt(d);
const dColor = d => d > 0 ? 'var(--green)' : d < 0 ? 'var(--red)' : 'var(--muted)';
const pctOf = (a, b) => b ? (a / b * 100).toFixed(1) + '%' : '&ndash;';
const qColor = q => q >= 41 ? 'var(--green)' : q >= 31 ? 'var(--gold)' : 'var(--red)';
const qBadge = q => q >= 41 ? 'badge-green' : q >= 31 ? 'badge-gold' : 'badge-red';
const impactBadge = s => s === 'Benefited' ? 'badge-green' : s === 'Impacted' ? 'badge-red' : 'badge-muted';
const COL = { ex:'#8FA9A6', a:'#00857C', b:'#B7791F', rec:'#1A2B2A' };

/* Banker's rounding — mirrors Python's round() so the dashboard and the
   Excel workbook never disagree by a rupee. */
function bround(x){
  const f = Math.floor(x), d = x - f;
  if (Math.abs(d - 0.5) > 1e-9) return Math.round(x);
  return f % 2 === 0 ? f : f + 1;
}

function getQMult(q, slabs){
  slabs = slabs || Q_SLABS;
  for (const s of slabs) if (q >= s.min && q <= s.max) return s;
  return slabs[slabs.length - 1];
}
function applySlab(amount, slabs, cap){
  for (const s of slabs){
    if (amount >= s.min && amount <= s.max){
      let inc = s.type === 'flat' ? s.val : bround(amount * s.val);
      if (cap) inc = Math.min(inc, cap);
      return { inc, slab: s.label };
    }
  }
  const s = slabs[slabs.length - 1];
  let inc = s.type === 'flat' ? s.val : bround(amount * s.val);
  if (cap) inc = Math.min(inc, cap);
  return { inc, slab: s.label };
}
const calcExisting = r => applySlab(r, EX_SLABS, EX_CAP);
const SCEN = {
  A: { label: LABEL_A, q: () => Q_SLABS,  slabs: () => PR_SLABS,  cap: () => PR_CAP },
  B: { label: LABEL_B, q: () => QB_SLABS, slabs: () => PRB_SLABS, cap: () => PRB_CAP },
};

/* One proposal (multiplier + slab) for one recovery / quality pair.
   Eligibility is tested on RAW recovery; adjusted may fall below the threshold. */
function score(recovery, quality, key){
  const sc = SCEN[key];
  const s = getQMult(Math.min(Math.round(quality), 100), sc.q());
  const computed = bround(recovery * s.mult);
  const below = recovery < MIN_RECOVERY;
  const floored = APPLY_FLOOR && !below && computed < MIN_RECOVERY;
  const adjusted = floored ? MIN_RECOVERY : computed;
  const inc = below ? 0 : applySlab(adjusted, sc.slabs(), sc.cap()).inc;
  return { tier: s.tier, mult: s.mult, computed, below, floored, adjusted, inc };
}

const slabText = s => s.type === 'flat'
  ? (s.val === 0 ? 'No Incentive' : 'Flat ' + fmt(s.val))
  : +(s.val * 100).toFixed(2) + '%';

function eligibilityNote(key){
  const sc = SCEN[key];
  const lowest = Math.min(...sc.q().map(s => s.mult));
  const example = MIN_RECOVERY + 1000;
  const computed = bround(example * lowest);
  let html = `<strong>Eligibility:</strong> <strong>raw</strong> (actual) recovery must be at least
    <strong>${fmt(MIN_RECOVERY)}</strong>. <strong>Adjusted</strong> recovery can be below ${fmt(MIN_RECOVERY)}`;
  if (lowest < 1 && computed < MIN_RECOVERY){
    html += ` &mdash; e.g. ${fmt(example)} recovered at a ${lowest}x multiplier gives ${fmt(computed)}.
      The officer is still eligible because the test is on raw recovery`;
    html += APPLY_FLOOR
      ? `; adjusted recovery is held at ${fmt(MIN_RECOVERY)} for the slab
         (${fmt(applySlab(MIN_RECOVERY, sc.slabs(), sc.cap()).inc)} incentive).`
      : `; the incentive is read from the band ${fmt(computed)} falls in
         (${fmt(applySlab(computed, sc.slabs(), sc.cap()).inc)}).`;
  } else {
    html += ` when the multiplier is under 1x; eligibility is still tested on raw recovery.`;
  }
  return html;
}

function bSameAsA(){
  return JSON.stringify(Q_SLABS) === JSON.stringify(QB_SLABS)
      && JSON.stringify(PR_SLABS) === JSON.stringify(PRB_SLABS) && PR_CAP === PRB_CAP;
}

function showTab(i){
  document.querySelectorAll('.tab').forEach((t, j) => t.classList.toggle('active', j === i));
  document.querySelectorAll('.tab-panel').forEach((p, j) => p.classList.toggle('active', j === i));
}

/* ---------- Tab 0 — Dashboard ---------- */

let dashChart;

function monthTotals(){
  return MONTHS.map(m => {
    const md = DATA.filter(r => r.month === m);
    const sum = k => md.reduce((s, r) => s + r[k], 0);
    return { m, officers: md.length, eligible: md.filter(r => !r.below).length,
             recovery: sum('recovery'), existing: sum('existing'), a: sum('proposed'), b: sum('proposedB') };
  });
}

function renderDashboard(){
  const rows = monthTotals();
  const T = rows.reduce((t, r) => ({ recovery: t.recovery + r.recovery, existing: t.existing + r.existing,
                                     a: t.a + r.a, b: t.b + r.b }),
                        { recovery: 0, existing: 0, a: 0, b: 0 });

  document.getElementById('dashBanner').style.display = bSameAsA() ? 'block' : 'none';

  document.getElementById('dashStats').innerHTML = `
    <div class="stat"><div class="stat-label">Total Recovery</div>
      <div class="stat-value" style="color:var(--text)">${fmt(T.recovery)}</div>
      <div class="stat-sub">${MONTHS.length} months &middot; ${OFFICERS.length} officers</div></div>
    <div class="stat" style="border-top:3px solid ${COL.ex}"><div class="stat-label">Existing Incentive</div>
      <div class="stat-value" style="color:var(--muted)">${fmt(T.existing)}</div>
      <div class="stat-sub">${pctOf(T.existing, T.recovery)} of recovery</div></div>
    <div class="stat" style="border-top:3px solid ${COL.a}"><div class="stat-label">${LABEL_A} Incentive</div>
      <div class="stat-value" style="color:var(--teal)">${fmt(T.a)}</div>
      <div class="stat-sub"><span style="color:${dColor(T.a - T.existing)};font-weight:600">${sfmt(T.a - T.existing)}</span>
        vs existing &middot; ${pctOf(T.a, T.recovery)} of recovery</div></div>
    <div class="stat" style="border-top:3px solid ${COL.b}"><div class="stat-label">${LABEL_B} Incentive</div>
      <div class="stat-value" style="color:var(--amber)">${fmt(T.b)}</div>
      <div class="stat-sub"><span style="color:${dColor(T.b - T.existing)};font-weight:600">${sfmt(T.b - T.existing)}</span>
        vs existing &middot; ${pctOf(T.b, T.recovery)} of recovery</div></div>`;

  document.getElementById('dashBody').innerHTML = rows.map(r => `
    <tr>
      <td style="font-weight:600">${r.m}</td>
      <td class="text-right">${r.officers}</td>
      <td class="text-right">${r.eligible}</td>
      <td class="text-right" style="font-weight:500">${fmt(r.recovery)}</td>
      <td class="text-right bl" style="color:var(--muted)">${fmt(r.existing)}</td>
      <td class="text-right bl" style="color:var(--teal);font-weight:600">${fmt(r.a)}</td>
      <td class="text-right" style="color:${dColor(r.a - r.existing)}">${sfmt(r.a - r.existing)}</td>
      <td class="text-right bl" style="color:var(--amber);font-weight:600">${fmt(r.b)}</td>
      <td class="text-right" style="color:${dColor(r.b - r.existing)}">${sfmt(r.b - r.existing)}</td>
    </tr>`).join('');

  document.getElementById('dashFoot').innerHTML = `
    <tr style="background:var(--teal-light);font-weight:700">
      <td>Total</td>
      <td class="text-right">${OFFICERS.length}</td>
      <td class="text-right" style="color:var(--muted)">&ndash;</td>
      <td class="text-right">${fmt(T.recovery)}</td>
      <td class="text-right bl" style="color:var(--muted)">${fmt(T.existing)}</td>
      <td class="text-right bl" style="color:var(--teal)">${fmt(T.a)}</td>
      <td class="text-right" style="color:${dColor(T.a - T.existing)}">${sfmt(T.a - T.existing)}</td>
      <td class="text-right bl" style="color:var(--amber)">${fmt(T.b)}</td>
      <td class="text-right" style="color:${dColor(T.b - T.existing)}">${sfmt(T.b - T.existing)}</td>
    </tr>`;

  // Scenario comparison
  const scen = [
    { name: 'Existing', total: T.existing, color: 'var(--muted)', ben: null, imp: null },
    { name: LABEL_A, total: T.a, color: 'var(--teal)', ben: SUMMARY.benefited, imp: SUMMARY.impacted },
    { name: LABEL_B, total: T.b, color: 'var(--amber)', ben: SUMMARY.benefitedB, imp: SUMMARY.impactedB },
  ];
  document.getElementById('dashCompare').innerHTML = scen.map(s => {
    const d = s.total - T.existing;
    const p = T.existing ? (d / T.existing * 100).toFixed(1) + '%' : '&ndash;';
    return `<tr>
      <td style="font-weight:600;color:${s.color}">${s.name}</td>
      <td class="text-right" style="font-weight:600">${fmt(s.total)}</td>
      <td class="text-right">${pctOf(s.total, T.recovery)}</td>
      <td class="text-right" style="color:${dColor(d)}">${s.ben === null ? '&ndash;' : sfmt(d)}</td>
      <td class="text-right" style="color:${dColor(d)}">${s.ben === null ? '&ndash;' : (d > 0 ? '+' : '') + p}</td>
      <td class="text-right" style="color:var(--green)">${s.ben ?? '&ndash;'}</td>
      <td class="text-right" style="color:var(--red)">${s.imp ?? '&ndash;'}</td>
    </tr>`;
  }).join('');

  if (dashChart) dashChart.destroy();
  if (typeof Chart === 'undefined') return;
  dashChart = new Chart(document.getElementById('dashChart'), {
    data: {
      labels: rows.map(r => r.m),
      datasets: [
        { type:'bar', label:'Existing', data: rows.map(r => r.existing), backgroundColor: COL.ex, borderRadius: 3, yAxisID:'y' },
        { type:'bar', label: LABEL_A, data: rows.map(r => r.a), backgroundColor: COL.a, borderRadius: 3, yAxisID:'y' },
        { type:'bar', label: LABEL_B, data: rows.map(r => r.b), backgroundColor: COL.b, borderRadius: 3, yAxisID:'y' },
        { type:'line', label:'Recovery', data: rows.map(r => r.recovery), borderColor: COL.rec, borderWidth: 2,
          borderDash:[5,4], pointRadius: 4, pointBackgroundColor: COL.rec, yAxisID:'y1' },
      ]
    },
    options: { responsive:true, maintainAspectRatio:false, interaction:{ mode:'index', intersect:false },
      scales:{
        y:{ title:{ display:true, text:'Incentive' }, ticks:{ callback: v => (v/1000).toFixed(0) + 'K' }, grid:{ color:'#EEF3F3' } },
        y1:{ position:'right', title:{ display:true, text:'Recovery' }, grid:{ display:false },
             ticks:{ callback: v => (v/1e6).toFixed(2) + 'M' } },
        x:{ grid:{ display:false } } },
      plugins:{ legend:{ position:'bottom' },
        tooltip:{ callbacks:{ label: c => c.dataset.label + ': ' + fmt(c.raw) } } } }
  });
}

/* ---------- Tab 1 — Model Overview ---------- */

function renderQualityTable(){
  const rowsFor = slabs => slabs.map((s, i) => {
    const color = s.mult < 1 ? 'var(--red)' : s.mult === 1 ? 'var(--gold)' : 'var(--green)';
    const bg = i % 2 === 0 ? 'var(--card)' : 'var(--bg)';
    return `<tr style="background:${bg}">
      <td>${s.min}% &ndash; ${s.max}%</td>
      <td style="text-align:center;color:${color};font-weight:600">${s.mult}x</td>
      <td style="color:${color};font-weight:500">${s.tier}</td>
    </tr>`;
  }).join('');
  document.getElementById('qualitySlabBody').innerHTML = rowsFor(Q_SLABS);
  document.getElementById('qualitySlabBodyB').innerHTML = rowsFor(QB_SLABS);
}

function renderOverviewSlabs(){
  const rows = (slabs, cap, color) => slabs.map((s, i) => {
    const c = (cap && i === slabs.length - 1) ? ` <span style="color:var(--muted);font-weight:400">(cap ${fmt(cap)})</span>` : '';
    return `<tr><td>${s.label}</td><td class="text-right" style="color:${color}"><strong>${slabText(s)}</strong>${c}</td></tr>`;
  }).join('');
  document.getElementById('existingSlabBody').innerHTML = rows(EX_SLABS, EX_CAP, 'var(--text)');
  document.getElementById('proposedSlabBody').innerHTML = rows(PR_SLABS, PR_CAP, 'var(--teal)');
  document.getElementById('proposedBSlabBody').innerHTML = rows(PRB_SLABS, PRB_CAP, 'var(--amber)');
  document.getElementById('eligNoteA').innerHTML = eligibilityNote('A');
  document.getElementById('eligNoteB').innerHTML = eligibilityNote('B');
}

function renderCaseExample(){
  const host = document.getElementById('caseExample');
  if (!host) return;

  const officers = [
    { name: 'Officer A', recovery: 150000, outstanding: 300000 },
    { name: 'Officer B', recovery: 150000, outstanding: 1000000 },
  ];

  const results = officers.map(o => {
    const quality = Math.min(Math.round((o.recovery / o.outstanding) * 100), 100);
    const a = score(o.recovery, quality, 'A');
    const b = score(o.recovery, quality, 'B');
    const existing = calcExisting(o.recovery).inc;
    return { ...o, quality, a, b, existing };
  });

  const adjCell = s => s.below ? '&mdash;'
    : s.floored ? `${fmt(s.adjusted)} <span class="badge badge-gold">Floor</span>` : fmt(s.adjusted);

  host.innerHTML = results.map(r => `<tr>
      <td style="font-weight:600">${r.name}</td>
      <td class="text-right">${fmt(r.recovery)}</td>
      <td class="text-right">${fmt(r.outstanding)}</td>
      <td class="text-right"><span class="badge ${qBadge(r.quality)}">${r.quality}%</span></td>
      <td class="text-right" style="color:var(--muted)">${fmt(r.existing)}</td>
      <td class="text-right bl">${r.a.mult}x</td>
      <td class="text-right">${adjCell(r.a)}</td>
      <td class="text-right" style="color:var(--teal);font-weight:600">${fmt(r.a.inc)}</td>
      <td class="text-right bl">${r.b.mult}x</td>
      <td class="text-right">${adjCell(r.b)}</td>
      <td class="text-right" style="color:var(--amber);font-weight:600">${fmt(r.b.inc)}</td>
    </tr>`).join('');

  const [x, y] = results;
  const gapA = x.a.inc - y.a.inc, gapB = x.b.inc - y.b.inc;
  const note = document.getElementById('caseExampleNote');
  if (note) {
    note.innerHTML = `Under the existing model both earn ${fmt(x.existing)} &mdash; no gap.
      Officer A earns <strong>${fmt(gapA)} more</strong> under ${LABEL_A} and
      <strong>${fmt(gapB)} more</strong> under ${LABEL_B}, for collecting the same amount against a smaller balance.`;
  }
}

function renderWorkedExamples(){
  const tA = document.getElementById('thresholdA');
  if (tA) tA.textContent = fmt(MIN_RECOVERY);
  const n2 = document.getElementById('note2Text');
  if (n2) n2.innerHTML = APPLY_FLOOR
    ? `Once an officer is eligible, the multiplier may reduce their adjusted recovery but never
       below <strong>${fmt(MIN_RECOVERY)}</strong>.`
    : `The multiplier is allowed to carry adjusted recovery below <strong>${fmt(MIN_RECOVERY)}</strong>,
       so the incentive is read off whichever slab band the reduced figure lands in.`;
}

function updateSimulator(){
  const q = +document.getElementById('simQuality').value;
  const r = +document.getElementById('simRecovery').value;
  document.getElementById('simQualityLabel').textContent = q + '%';
  document.getElementById('simRecoveryLabel').textContent = fmt(r);

  const ex = calcExisting(r).inc;
  const row = (name, color, s) => {
    const d = s.inc - ex;
    const adj = s.below ? '&mdash;'
      : fmt(s.adjusted) + (s.floored ? `<div style="font-size:11px;color:var(--gold)">raised from ${fmt(s.computed)}</div>` : '');
    return `<tr>
      <td style="font-weight:600;color:${color}">${name}</td>
      <td>${s.below ? '<span style="color:var(--muted)">below threshold</span>' : s.tier + ' (' + s.mult + 'x)'}</td>
      <td class="text-right">${adj}</td>
      <td class="text-right" style="font-weight:700;color:${color}">${fmt(s.inc)}</td>
      <td class="text-right" style="font-weight:600;color:${dColor(d)}">${sfmt(d)}</td>
    </tr>`;
  };
  document.getElementById('simResult').innerHTML = `
    <tr><td style="font-weight:600;color:var(--muted)">Existing</td><td style="color:var(--muted)">on raw recovery</td>
      <td class="text-right">${fmt(r)}</td><td class="text-right" style="font-weight:700">${fmt(ex)}</td><td></td></tr>
    ${row(LABEL_A, 'var(--teal)', score(r, q, 'A'))}
    ${row(LABEL_B, 'var(--amber)', score(r, q, 'B'))}`;
}

/* ---------- Tab 2 — Monthly Performance ---------- */

let currentMonth = MONTHS[0];
let currentSort = 'recovery';

function renderMonthly(){
  document.getElementById('monthButtons').innerHTML = MONTHS.map(m =>
    `<button class="btn ${m === currentMonth ? 'active' : ''}" onclick="currentMonth='${m}';renderMonthly()">${m}</button>`
  ).join('');

  const search = (document.getElementById('officerFilter')?.value || '').toLowerCase();
  let rows = DATA.filter(r => r.month === currentMonth);
  if (search) rows = rows.filter(r =>
    r.name.toLowerCase().includes(search) || String(r.emp).includes(search));

  if (currentSort === 'recovery')      rows.sort((a, b) => b.recovery - a.recovery);
  else if (currentSort === 'quality')  rows.sort((a, b) => b.quality - a.quality);
  else if (currentSort === 'diffB')    rows.sort((a, b) => b.diffB - a.diffB);
  else                                 rows.sort((a, b) => b.diff - a.diff);

  const sum = k => rows.reduce((s, r) => s + r[k], 0);
  const totalRec = sum('recovery'), totalEx = sum('existing');
  const totalA = sum('proposed'), totalB = sum('proposedB');
  const totalAdjA = rows.reduce((s, r) => s + (r.below ? 0 : r.adjusted), 0);
  const totalAdjB = rows.reduce((s, r) => s + (r.below ? 0 : r.adjustedB), 0);
  const avgQ = rows.length ? Math.round(rows.reduce((s, r) => s + r.quality, 0) / rows.length) : 0;

  document.getElementById('monthStats').innerHTML = `
    <div class="stat"><div class="stat-label">Total Recovery</div>
      <div class="stat-value">${fmt(totalRec)}</div>
      <div class="stat-sub">${rows.filter(r => !r.below).length} of ${rows.length} eligible</div></div>
    <div class="stat"><div class="stat-label">Avg Quality</div>
      <div class="stat-value" style="color:${qColor(avgQ)}">${avgQ}%</div></div>
    <div class="stat" style="border-top:3px solid ${COL.ex}"><div class="stat-label">Existing Incentive</div>
      <div class="stat-value" style="color:var(--muted)">${fmt(totalEx)}</div></div>
    <div class="stat" style="border-top:3px solid ${COL.a}"><div class="stat-label">${LABEL_A}</div>
      <div class="stat-value" style="color:var(--teal)">${fmt(totalA)}</div>
      <div class="stat-sub" style="color:${dColor(totalA - totalEx)}">${sfmt(totalA - totalEx)} vs existing</div></div>
    <div class="stat" style="border-top:3px solid ${COL.b}"><div class="stat-label">${LABEL_B}</div>
      <div class="stat-value" style="color:var(--amber)">${fmt(totalB)}</div>
      <div class="stat-sub" style="color:${dColor(totalB - totalEx)}">${sfmt(totalB - totalEx)} vs existing</div></div>`;

  const multCell = (below, mult, tier) => below
    ? `<td class="text-right bl" style="color:var(--muted)">&mdash;</td>`
    : `<td class="text-right bl" title="${tier}" style="font-weight:600;color:${mult < 1 ? 'var(--red)' : mult > 1 ? 'var(--green)' : 'var(--gold)'}">${mult}x</td>`;
  const adjCell = (below, floored, adjusted, computed) => below
    ? `<td class="text-right" style="color:var(--muted)">&mdash;</td>`
    : floored
      ? `<td class="text-right" title="Computed ${fmt(computed)}, raised to the ${fmt(MIN_RECOVERY)} floor">${fmt(adjusted)} <span class="badge badge-gold">Floor</span></td>`
      : `<td class="text-right">${fmt(adjusted)}</td>`;

  document.getElementById('monthlyBody').innerHTML = rows.map((r, i) => `
    <tr>
      <td class="text-right" style="color:var(--muted)">${i + 1}</td>
      <td style="font-weight:500">${r.name}</td>
      <td style="color:var(--muted)">${r.region}</td>
      <td class="text-right">${r.cases}</td>
      <td class="text-right" style="font-weight:500">${fmt(r.recovery)}${r.below ? ' <span class="badge badge-muted">&lt; ' + fmt(MIN_RECOVERY) + '</span>' : ''}</td>
      <td class="text-right"><span class="badge ${qBadge(r.quality)}">${r.quality}%</span></td>
      <td class="text-right bl" style="color:var(--muted)">${fmt(r.existing)}</td>
      ${multCell(r.below, r.mult, r.tier)}${adjCell(r.below, r.floored, r.adjusted, r.computed)}
      <td class="text-right" style="color:var(--teal);font-weight:600">${fmt(r.proposed)}</td>
      <td class="text-right" style="font-weight:600;color:${dColor(r.diff)}">${sfmt(r.diff)}</td>
      ${multCell(r.below, r.multB, r.tierB)}${adjCell(r.below, r.flooredB, r.adjustedB, r.computedB)}
      <td class="text-right" style="color:var(--amber);font-weight:600">${fmt(r.proposedB)}</td>
      <td class="text-right" style="font-weight:600;color:${dColor(r.diffB)}">${sfmt(r.diffB)}</td>
    </tr>`).join('');

  document.getElementById('monthlyFoot').innerHTML = `
    <tr style="background:var(--teal-light);font-weight:700">
      <td colspan="4">Total &mdash; ${rows.length} officers</td>
      <td class="text-right">${fmt(totalRec)}</td>
      <td class="text-right">${avgQ}%</td>
      <td class="text-right bl" style="color:var(--muted)">${fmt(totalEx)}</td>
      <td class="bl"></td><td class="text-right">${fmt(totalAdjA)}</td>
      <td class="text-right" style="color:var(--teal)">${fmt(totalA)}</td>
      <td class="text-right" style="color:${dColor(totalA - totalEx)}">${sfmt(totalA - totalEx)}</td>
      <td class="bl"></td><td class="text-right">${fmt(totalAdjB)}</td>
      <td class="text-right" style="color:var(--amber)">${fmt(totalB)}</td>
      <td class="text-right" style="color:${dColor(totalB - totalEx)}">${sfmt(totalB - totalEx)}</td>
    </tr>`;
}

function sortMonthly(key, btn){
  currentSort = key;
  btn.parentElement.querySelectorAll('.btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderMonthly();
}

/* ---------- Tab 3 — Impact Analysis ---------- */

let impactChart, trendChart, impactSort = 'diff';

function renderImpact(){
  const E = SUMMARY.totalExisting, A = SUMMARY.totalProposed, B = SUMMARY.totalProposedB;
  const pct = d => E > 0 ? ((d / E) * 100).toFixed(1) + '%' : '0%';

  document.getElementById('impactStats').innerHTML = `
    <div class="stat" style="border-top:3px solid ${COL.ex}"><div class="stat-label">Existing Incentive</div>
      <div class="stat-value" style="color:var(--muted)">${fmt(E)}</div></div>
    <div class="stat" style="border-top:3px solid ${COL.a}"><div class="stat-label">${LABEL_A}</div>
      <div class="stat-value" style="color:var(--teal)">${fmt(A)}</div>
      <div class="stat-sub"><span style="color:${dColor(A - E)};font-weight:600">${sfmt(A - E)}</span> (${pct(A - E)})
        &middot; ${SUMMARY.benefited} benefited &middot; ${SUMMARY.impacted} impacted</div></div>
    <div class="stat" style="border-top:3px solid ${COL.b}"><div class="stat-label">${LABEL_B}</div>
      <div class="stat-value" style="color:var(--amber)">${fmt(B)}</div>
      <div class="stat-sub"><span style="color:${dColor(B - E)};font-weight:600">${sfmt(B - E)}</span> (${pct(B - E)})
        &middot; ${SUMMARY.benefitedB} benefited &middot; ${SUMMARY.impactedB} impacted</div></div>
    <div class="stat"><div class="stat-label">Officers</div>
      <div class="stat-value">${SUMMARY.officers}</div></div>`;

  if (typeof Chart !== 'undefined'){
    if (impactChart) impactChart.destroy();
    const n = SUMMARY.officers;
    impactChart = new Chart(document.getElementById('impactChart'), {
      type: 'bar',
      data: {
        labels: ['Benefited', 'Impacted', 'No Change'],
        datasets: [
          { label: LABEL_A, data: [SUMMARY.benefited, SUMMARY.impacted, n - SUMMARY.benefited - SUMMARY.impacted],
            backgroundColor: COL.a, borderRadius: 3 },
          { label: LABEL_B, data: [SUMMARY.benefitedB, SUMMARY.impactedB, n - SUMMARY.benefitedB - SUMMARY.impactedB],
            backgroundColor: COL.b, borderRadius: 3 },
        ]
      },
      options: { responsive:true, maintainAspectRatio:false,
        scales:{ y:{ ticks:{ precision:0 }, title:{ display:true, text:'Officers' } }, x:{ grid:{ display:false } } },
        plugins:{ legend:{ position:'bottom' } } }
    });

    const trend = monthTotals();
    if (trendChart) trendChart.destroy();
    trendChart = new Chart(document.getElementById('trendChart'), {
      type: 'line',
      data: {
        labels: trend.map(d => d.m),
        datasets: [
          { label:'Existing', data:trend.map(d=>d.existing), borderColor:COL.ex, borderWidth:2, pointRadius:4 },
          { label:LABEL_A,    data:trend.map(d=>d.a),        borderColor:COL.a,  borderWidth:2, pointRadius:4 },
          { label:LABEL_B,    data:trend.map(d=>d.b),        borderColor:COL.b,  borderWidth:2, pointRadius:4, borderDash:[6,4] }
        ]
      },
      options: { responsive:true, maintainAspectRatio:false, interaction:{ mode:'index', intersect:false },
        scales:{ y:{ ticks:{ callback: v => (v/1000).toFixed(0) + 'K' } } },
        plugins:{ legend:{ position:'bottom' }, tooltip:{ callbacks:{ label: c => c.dataset.label + ': ' + fmt(c.raw) } } } }
    });
  }

  const sorted = [...OVERALL].sort((a, b) => b[impactSort] - a[impactSort]);
  document.getElementById('impactBody').innerHTML = sorted.map((o, i) => `
    <tr>
      <td class="text-right" style="color:var(--muted)">${i + 1}</td>
      <td style="font-weight:500">${o.name}</td>
      <td style="color:var(--muted)">${o.region}</td>
      <td class="text-right"><span class="badge ${qBadge(o.quality)}">${o.quality}%</span></td>
      <td class="text-right">${fmt(o.recovery)}</td>
      <td class="text-right bl" style="color:var(--muted)">${fmt(o.existing)}</td>
      <td class="text-right bl" style="color:var(--teal);font-weight:600">${fmt(o.proposed)}</td>
      <td class="text-right" style="font-weight:600;color:${dColor(o.diff)}">${sfmt(o.diff)}</td>
      <td><span class="badge ${impactBadge(o.impact)}">${o.impact}</span></td>
      <td class="text-right bl" style="color:var(--amber);font-weight:600">${fmt(o.proposedB)}</td>
      <td class="text-right" style="font-weight:600;color:${dColor(o.diffB)}">${sfmt(o.diffB)}</td>
      <td><span class="badge ${impactBadge(o.impactB)}">${o.impactB}</span></td>
    </tr>`).join('');
}

function sortImpact(key, btn){
  impactSort = key;
  btn.parentElement.querySelectorAll('.btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderImpact();
}

/* ---------- Tab 4 — Officer Deep Dive ---------- */

let selectedOfficer = null;

function initOfficerSearch(){
  document.getElementById('officerList').innerHTML = OFFICERS.map(o =>
    `<div class="search-item" onclick="selectOfficer('${o.emp}')">${o.name} <span style="color:var(--muted)">(${o.emp})</span></div>`
  ).join('');
}

function filterOfficers(){
  const q = document.getElementById('officerSearch').value.toLowerCase();
  const list = document.getElementById('officerList');
  list.classList.add('open');
  list.querySelectorAll('.search-item').forEach(el => {
    el.style.display = el.textContent.toLowerCase().includes(q) ? 'block' : 'none';
  });
}

function selectOfficer(emp){
  const officer = OFFICERS.find(o => o.emp == emp);
  selectedOfficer = emp;
  document.getElementById('officerSearch').value = `${officer.name} (${officer.emp})`;
  document.getElementById('officerList').classList.remove('open');
  renderDeepDive();
}

document.addEventListener('click', e => {
  if (!e.target.closest('.search-select'))
    document.getElementById('officerList')?.classList.remove('open');
});

function renderDeepDive(){
  const container = document.getElementById('deepDiveContent');
  if (!selectedOfficer){ container.innerHTML = ''; return; }

  const months = DATA.filter(r => r.emp == selectedOfficer);
  if (!months.length){ container.innerHTML = ''; return; }

  const cases = CASES.filter(c => c.emp == selectedOfficer);
  const sum = k => months.reduce((s, r) => s + r[k], 0);
  const totalRec = sum('recovery'), totalEx = sum('existing'), totalA = sum('proposed'), totalB = sum('proposedB');
  const avgQ = Math.round(sum('quality') / months.length);
  const item = (label, value, color) => `<div class="officer-summary-item"><div class="officer-summary-label">${label}</div>
        <div class="officer-summary-value" style="color:${color}">${value}</div></div>`;

  let html = `<h3 style="margin-bottom:14px">${months[0].name} <span style="color:var(--muted);font-weight:400">&middot; ${months[0].region}</span></h3>
    <div class="officer-summary">
      ${item('Total Recovery', fmt(totalRec), 'var(--text)')}
      ${item('Avg Quality', avgQ + '%', qColor(avgQ))}
      ${item('Existing', fmt(totalEx), 'var(--muted)')}
      ${item(LABEL_A, fmt(totalA) + ` <span style="font-size:12px;color:${dColor(totalA - totalEx)}">${sfmt(totalA - totalEx)}</span>`, 'var(--teal)')}
      ${item(LABEL_B, fmt(totalB) + ` <span style="font-size:12px;color:${dColor(totalB - totalEx)}">${sfmt(totalB - totalEx)}</span>`, 'var(--amber)')}
    </div>`;

  MONTHS.forEach(month => {
    const mo = months.find(r => r.month === month);
    if (!mo) return;
    const monthCases = cases
      .map(c => ({ ...c, entry: c.timeline.find(t => t.m === month) }))
      .filter(c => c.entry)
      .sort((a, b) => b.entry.rec - a.entry.rec || b.os - a.os);
    if (!monthCases.length) return;

    const ms = (label, value, color, sub) => `<div class="month-summary-item"><div class="month-summary-label">${label}</div>
          <div class="month-summary-value"${color ? ` style="color:${color}"` : ''}>${value}</div>${sub || ''}</div>`;
    const raised = (f, c) => f ? `<div style="font-size:10px;color:var(--gold)">raised from ${fmt(c)}</div>` : '';

    html += `<div class="month-header">${month}${mo.below ? ' <span style="font-size:12px;font-weight:400;opacity:.85">&middot; below ' + fmt(MIN_RECOVERY) + ' threshold, no incentive</span>' : ''}</div>
      <div class="month-summary">
        ${ms('Outstanding', fmt(mo.outstanding))}
        ${ms('Recovery', fmt(mo.recovery))}
        ${ms('Avg Quality', mo.quality + '%', qColor(mo.quality))}
        ${ms('Existing', fmt(mo.existing), 'var(--muted)')}
        ${ms(LABEL_A + ' adj.', mo.below ? '&mdash;' : fmt(mo.adjusted) + ` <span style="font-size:11px">(${mo.mult}x)</span>`, null, raised(mo.floored, mo.computed))}
        ${ms(LABEL_A, fmt(mo.proposed), 'var(--teal)')}
        ${ms(LABEL_B + ' adj.', mo.below ? '&mdash;' : fmt(mo.adjustedB) + ` <span style="font-size:11px">(${mo.multB}x)</span>`, null, raised(mo.flooredB, mo.computedB))}
        ${ms(LABEL_B, fmt(mo.proposedB), 'var(--amber)')}
      </div>
      <div class="case-table-wrap"><table>
        <thead><tr>
          <th>Case</th><th>Account</th>
          <th class="text-right">Outstanding</th>
          <th class="text-right">This Month</th>
          <th class="text-right">Cumulative</th>
          <th class="text-right">Quality %</th>
          <th>Status</th>
        </tr></thead><tbody>`;

    monthCases.forEach(c => {
      const t = c.entry;
      const status = t.settlement ? '<span class="badge badge-gold">Settlement</span>'
                   : t.closed     ? '<span class="badge badge-muted">Closed</span>' : '';
      html += `<tr>
        <td style="font-weight:500">${c.case}</td>
        <td style="color:var(--muted)">${c.account}</td>
        <td class="text-right">${fmt(c.os)}</td>
        <td class="text-right">${fmt(t.rec)}</td>
        <td class="text-right" style="font-weight:500">${fmt(t.cum)}</td>
        <td class="text-right"><span class="badge ${qBadge(t.q)}">${t.q}%</span></td>
        <td>${status}</td>
      </tr>`;
    });

    html += `</tbody></table></div>`;
  });

  container.innerHTML = html;
}

/* ---------- Settings modal ---------- */

const clone = x => JSON.parse(JSON.stringify(x));
const DEFAULTS = {
  q: clone(Q_SLABS), qb: clone(QB_SLABS),
  ex: clone(EX_SLABS), pr: clone(PR_SLABS), prb: clone(PRB_SLABS),
  exCap: EX_CAP, prCap: PR_CAP, prbCap: PRB_CAP, minRecovery: MIN_RECOVERY, applyFloor: APPLY_FLOOR
};

function openSettings(){
  document.getElementById('settingsModal').classList.add('open');
  renderEditors();
}
function closeSettings(){
  document.getElementById('settingsModal').classList.remove('open');
}

const qualityRows = slabs => slabs.map(s => `
  <tr>
    <td><input type="number" value="${s.min}"></td>
    <td><input type="number" value="${s.max}"></td>
    <td><input type="number" step="0.01" value="${s.mult}"></td>
    <td><input value="${s.tier}"></td>
    <td><button class="remove-btn" onclick="removeRow(this)">&times;</button></td>
  </tr>`).join('');

const incentiveRows = slabs => slabs.map(s => `
  <tr>
    <td><input type="number" value="${s.min}"></td>
    <td><input type="number" value="${s.max}"></td>
    <td><select>
      <option value="flat" ${s.type === 'flat' ? 'selected' : ''}>Flat</option>
      <option value="pct" ${s.type === 'pct' ? 'selected' : ''}>Percent</option>
    </select></td>
    <td><input type="number" step="0.001" value="${s.val}"></td>
    <td><input value="${s.label}"></td>
    <td><button class="remove-btn" onclick="removeRow(this)">&times;</button></td>
  </tr>`).join('');

function renderEditors(){
  document.getElementById('qualityEditor').innerHTML = qualityRows(Q_SLABS);
  document.getElementById('qualityBEditor').innerHTML = qualityRows(QB_SLABS);
  document.getElementById('existingEditor').innerHTML = incentiveRows(EX_SLABS);
  document.getElementById('proposedEditor').innerHTML = incentiveRows(PR_SLABS);
  document.getElementById('proposedBEditor').innerHTML = incentiveRows(PRB_SLABS);
  document.getElementById('existingCap').value = EX_CAP;
  document.getElementById('proposedCap').value = PR_CAP;
  document.getElementById('proposedBCap').value = PRB_CAP;
  document.getElementById('minRecovery').value = MIN_RECOVERY;
  document.getElementById('applyFloor').checked = APPLY_FLOOR;
}

/* Copy whatever is currently typed in Proposal A's editors into Proposal B's. */
function copyAtoB(){
  document.getElementById('qualityBEditor').innerHTML = qualityRows(readQuality('qualityEditor'));
  document.getElementById('proposedBEditor').innerHTML = incentiveRows(readIncentive('proposedEditor'));
  document.getElementById('proposedBCap').value = document.getElementById('proposedCap').value;
}

function removeRow(btn){
  const body = btn.closest('tbody');
  if (body.rows.length <= 1){ alert('At least one row is required.'); return; }
  btn.closest('tr').remove();
}

function addRow(id){
  const body = document.getElementById(id);
  if (id.startsWith('quality')){
    body.insertAdjacentHTML('beforeend', qualityRows([{ min:0, max:100, mult:1.0, tier:'New Tier' }]));
  } else {
    body.insertAdjacentHTML('beforeend', incentiveRows([{ min:0, max:99999999, type:'pct', val:0.10, label:'New Slab' }]));
  }
}

function readQuality(id){
  const out = [];
  document.querySelectorAll('#' + id + ' tr').forEach(tr => {
    const i = tr.querySelectorAll('input');
    if (i.length >= 4) out.push({ min:+i[0].value, max:+i[1].value, mult:+i[2].value, tier:i[3].value });
  });
  return out;
}
function readIncentive(id){
  const out = [];
  document.querySelectorAll('#' + id + ' tr').forEach(tr => {
    const i = tr.querySelectorAll('input');
    const sel = tr.querySelector('select');
    if (i.length >= 3 && sel)
      out.push({ min:+i[0].value, max:+i[1].value, type:sel.value, val:+i[2].value, label:i[3]?.value || '' });
  });
  return out;
}

function readEditors(){
  Q_SLABS.splice(0, Q_SLABS.length, ...readQuality('qualityEditor'));
  QB_SLABS.splice(0, QB_SLABS.length, ...readQuality('qualityBEditor'));
  EX_SLABS.splice(0, EX_SLABS.length, ...readIncentive('existingEditor'));
  PR_SLABS.splice(0, PR_SLABS.length, ...readIncentive('proposedEditor'));
  PRB_SLABS.splice(0, PRB_SLABS.length, ...readIncentive('proposedBEditor'));
  EX_CAP = +document.getElementById('existingCap').value;
  PR_CAP = +document.getElementById('proposedCap').value;
  PRB_CAP = +document.getElementById('proposedBCap').value;
  MIN_RECOVERY = +document.getElementById('minRecovery').value;
  APPLY_FLOOR = document.getElementById('applyFloor').checked;
}

function recalculate(){
  DATA.forEach(r => {
    const a = score(r.recovery, r.quality, 'A');
    const b = score(r.recovery, r.quality, 'B');
    r.below = a.below;
    r.existing = calcExisting(r.recovery).inc;
    r.tier = a.tier;  r.mult = a.mult;  r.computed = a.computed;  r.floored = a.floored;
    r.adjusted = a.adjusted;  r.proposed = a.inc;  r.diff = a.inc - r.existing;
    r.tierB = b.tier; r.multB = b.mult; r.computedB = b.computed; r.flooredB = b.floored;
    r.adjustedB = b.adjusted; r.proposedB = b.inc; r.diffB = b.inc - r.existing;
  });

  const byOfficer = {};
  DATA.forEach(r => {
    const t = byOfficer[r.emp] = byOfficer[r.emp] || { existing:0, a:0, b:0 };
    t.existing += r.existing; t.a += r.proposed; t.b += r.proposedB;
  });
  const status = d => d > 0 ? 'Benefited' : d < 0 ? 'Impacted' : 'No Change';
  OVERALL.forEach(o => {
    const t = byOfficer[o.emp] || { existing:0, a:0, b:0 };
    o.existing = t.existing;
    o.proposed = t.a;  o.diff = t.a - t.existing;  o.impact = status(o.diff);
    o.proposedB = t.b; o.diffB = t.b - t.existing; o.impactB = status(o.diffB);
  });

  SUMMARY.totalExisting  = DATA.reduce((s, r) => s + r.existing, 0);
  SUMMARY.totalProposed  = DATA.reduce((s, r) => s + r.proposed, 0);
  SUMMARY.totalProposedB = DATA.reduce((s, r) => s + r.proposedB, 0);
  SUMMARY.benefited  = OVERALL.filter(o => o.impact === 'Benefited').length;
  SUMMARY.impacted   = OVERALL.filter(o => o.impact === 'Impacted').length;
  SUMMARY.benefitedB = OVERALL.filter(o => o.impactB === 'Benefited').length;
  SUMMARY.impactedB  = OVERALL.filter(o => o.impactB === 'Impacted').length;
}

function refreshAll(){
  renderDashboard();
  renderQualityTable();
  renderOverviewSlabs();
  renderWorkedExamples();
  renderCaseExample();
  updateSimulator();
  renderMonthly();
  renderImpact();
  if (selectedOfficer) renderDeepDive();
}

function toast(message){
  const t = document.getElementById('toast');
  t.textContent = message;
  t.style.display = 'block';
  setTimeout(() => { t.style.display = 'none'; }, 2500);
}

function applySettings(){
  readEditors();
  recalculate();
  refreshAll();
  closeSettings();
  toast('Changes applied — all tabs recalculated');
}

function resetSettings(){
  Q_SLABS.splice(0, Q_SLABS.length, ...clone(DEFAULTS.q));
  QB_SLABS.splice(0, QB_SLABS.length, ...clone(DEFAULTS.qb));
  EX_SLABS.splice(0, EX_SLABS.length, ...clone(DEFAULTS.ex));
  PR_SLABS.splice(0, PR_SLABS.length, ...clone(DEFAULTS.pr));
  PRB_SLABS.splice(0, PRB_SLABS.length, ...clone(DEFAULTS.prb));
  EX_CAP = DEFAULTS.exCap;
  PR_CAP = DEFAULTS.prCap;
  PRB_CAP = DEFAULTS.prbCap;
  MIN_RECOVERY = DEFAULTS.minRecovery;
  APPLY_FLOOR = DEFAULTS.applyFloor;
  renderEditors();
  recalculate();
  refreshAll();
  closeSettings();
  toast('Reset to default configuration');
}

window.addEventListener('DOMContentLoaded', () => {
  refreshAll();
  initOfficerSearch();
});
"""


def generate_html(input_file, df_officer, df_overall, case_timeline):
    months = sorted(df_officer['Recovery Month'].unique(), key=lambda m: MONTH_ORDER.get(m, 99))

    data = [{
        'month': r['Recovery Month'], 'emp': r['EmpCode'], 'name': r['Recovery Officer'],
        'region': r['Region'], 'cases': int(r['Total Cases']),
        'outstanding': int(r['Total Outstanding']), 'recovery': int(r['Total Recovery']),
        'quality': round(float(r['Avg Quality %'])),
        'below': bool(r['Below Threshold']), 'existing': int(r['Existing Incentive']),
        'tier': r['Quality Tier'], 'mult': float(r['Quality Multiplier']),
        'computed': int(r['Computed Recovery']), 'adjusted': int(r['Adjusted Recovery']),
        'floored': bool(r['Floored']), 'proposed': int(r['Proposed Incentive']),
        'diff': int(r['Incentive Difference']),
        'tierB': r['B Tier'], 'multB': float(r['B Multiplier']),
        'computedB': int(r['B Computed Recovery']), 'adjustedB': int(r['B Adjusted Recovery']),
        'flooredB': bool(r['B Floored']), 'proposedB': int(r['B Incentive']),
        'diffB': int(r['B Difference']),
    } for _, r in df_officer.iterrows()]

    overall = [{
        'emp': r['EmpCode'], 'name': r['Recovery Officer'], 'region': r['Region'],
        'recovery': int(r['Total Recovery']), 'quality': round(float(r['Avg Quality %'])),
        'existing': int(r['Total Existing']),
        'proposed': int(r['Total Proposed']), 'diff': int(r['Total Difference']), 'impact': r['Impact'],
        'proposedB': int(r['Total B']), 'diffB': int(r['B Total Difference']), 'impactB': r['B Impact'],
    } for _, r in df_overall.iterrows()]

    summary = {
        'officers': len(df_overall),
        'benefited': int((df_overall['Impact'] == 'Benefited').sum()),
        'impacted': int((df_overall['Impact'] == 'Impacted').sum()),
        'benefitedB': int((df_overall['B Impact'] == 'Benefited').sum()),
        'impactedB': int((df_overall['B Impact'] == 'Impacted').sum()),
        'totalExisting': int(df_officer['Existing Incentive'].sum()),
        'totalProposed': int(df_officer['Proposed Incentive'].sum()),
        'totalProposedB': int(df_officer['B Incentive'].sum()),
    }

    officers = [{'emp': r['EmpCode'], 'name': r['Recovery Officer']}
                for _, r in df_overall[['EmpCode', 'Recovery Officer']].drop_duplicates().iterrows()]
    officers.sort(key=lambda o: o['name'])

    q_list = lambda slabs: [{'min': s[0], 'max': s[1], 'mult': s[2], 'tier': s[3]} for s in slabs]
    inc_list = lambda slabs: [{'min': s[0], 'max': s[1], 'type': s[2], 'val': s[3], 'label': s[4]}
                              for s in slabs]

    def js_default(o):
        if isinstance(o, np.integer):
            return int(o)
        if isinstance(o, np.floating):
            return float(o)
        return str(o)

    dump = lambda x: json.dumps(x, default=js_default)
    payload = (
        f"const DATA={dump(data)};\n"
        f"const OVERALL={dump(overall)};\n"
        f"const CASES={dump(case_timeline)};\n"
        f"const MONTHS={dump(months)};\n"
        f"const OFFICERS={dump(officers)};\n"
        f"const SUMMARY={dump(summary)};\n"
        f"const LABEL_A={dump(LABEL_A)};\n"
        f"const LABEL_B={dump(LABEL_B)};\n"
        f"const Q_SLABS={dump(q_list(QUALITY_SLABS))};\n"
        f"const QB_SLABS={dump(q_list(ALT_QUALITY_SLABS))};\n"
        f"const EX_SLABS={dump(inc_list(EXISTING_INCENTIVE_SLABS))};\n"
        f"const PR_SLABS={dump(inc_list(PROPOSED_INCENTIVE_SLABS))};\n"
        f"const PRB_SLABS={dump(inc_list(ALT_INCENTIVE_SLABS))};\n"
        f"let EX_CAP={EXISTING_CAP or 0};\n"
        f"let PR_CAP={PROPOSED_CAP or 0};\n"
        f"let PRB_CAP={ALT_CAP or 0};\n"
        f"let MIN_RECOVERY={MIN_RECOVERY_FOR_INCENTIVE};\n"
        f"let APPLY_FLOOR={'true' if APPLY_ADJUSTED_FLOOR else 'false'};\n"
    )

    period = f"{months[0]} – {months[-1]}" if len(months) > 1 else months[0]
    A, B = LABEL_A, LABEL_B

    def incentive_editor(title, body_id, cap_id, extra=''):
        return f"""
    <div class="slab-section">
      <h3>{title}{extra}</h3>
      <table class="slab-edit">
        <thead><tr><th>Min</th><th>Max</th><th>Type</th><th>Value</th><th>Label</th><th></th></tr></thead>
        <tbody id="{body_id}"></tbody>
      </table>
      <button class="add-btn" onclick="addRow('{body_id}')">+ Add row</button>
      <div class="cap-row" style="margin-top:10px">
        <label>Cap</label><input type="number" id="{cap_id}">
        <span style="color:var(--muted)">0 removes the cap</span>
      </div>
    </div>"""

    def quality_editor(title, body_id, extra=''):
        return f"""
    <div class="slab-section">
      <h3>{title}{extra}</h3>
      <table class="slab-edit">
        <thead><tr><th>Min %</th><th>Max %</th><th>Multiplier</th><th>Tier</th><th></th></tr></thead>
        <tbody id="{body_id}"></tbody>
      </table>
      <button class="add-btn" onclick="addRow('{body_id}')">+ Add row</button>
    </div>"""

    copy_btn = (' <button class="add-btn" style="float:right;padding:3px 10px;font-size:12px" '
                f'onclick="copyAtoB()">Copy from {A}</button>')

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Quality-Based Incentive Model</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>{DASHBOARD_CSS}</style>
</head>
<body>

<div class="header">
  <div class="header-logo">HBL</div>
  <div class="header-title">Quality-Based Incentive Model</div>
  <div class="header-right">{period}</div>
  <button class="gear-btn" onclick="openSettings()" title="Configure slabs">&#9881;</button>
</div>

<div class="toast" id="toast"></div>

<div class="modal-overlay" id="settingsModal">
  <div class="modal">
    <h2>Configure Slabs</h2>
    <p class="modal-intro">Edit the values below and apply. Every tab recalculates instantly.
      {A} and {B} each have their own multiplier and incentive slab; the existing model is the baseline.</p>

    {incentive_editor('Existing Incentive', 'existingEditor', 'existingCap')}

    <div class="modal-group" style="border-color:var(--teal)">
      <div class="modal-group-title" style="color:var(--teal)">{A}</div>
      {quality_editor('Quality Multiplier', 'qualityEditor')}
      {incentive_editor('Incentive Slab', 'proposedEditor', 'proposedCap')}
    </div>

    <div class="modal-group" style="border-color:var(--amber)">
      <div class="modal-group-title" style="color:var(--amber)">{B}{copy_btn}</div>
      {quality_editor('Quality Multiplier', 'qualityBEditor')}
      {incentive_editor('Incentive Slab', 'proposedBEditor', 'proposedBCap')}
    </div>

    <div class="slab-section">
      <h3>Thresholds (apply to both proposals)</h3>
      <div class="cap-row" style="margin-bottom:12px">
        <label>Minimum monthly recovery</label><input type="number" id="minRecovery">
        <span style="color:var(--muted)">Tested on raw recovery. Below this, no incentive is paid and the multiplier is not applied</span>
      </div>
      <div class="cap-row">
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer;width:auto">
          <input type="checkbox" id="applyFloor" style="width:auto;margin:0;cursor:pointer">
          <span>Hold adjusted recovery at the minimum</span>
        </label>
        <span style="color:var(--muted)">Untick to let a weak multiplier carry it below</span>
      </div>
    </div>

    <div class="modal-actions">
      <button class="btn-close" onclick="closeSettings()">Cancel</button>
      <button class="btn-reset" onclick="resetSettings()">Reset to default</button>
      <button class="btn-apply" onclick="applySettings()">Apply changes</button>
    </div>
  </div>
</div>

<div class="tab-bar">
  <button class="tab active" onclick="showTab(0)">Dashboard</button>
  <button class="tab" onclick="showTab(1)">Model Overview</button>
  <button class="tab" onclick="showTab(2)">Monthly Performance</button>
  <button class="tab" onclick="showTab(3)">Impact Analysis</button>
  <button class="tab" onclick="showTab(4)">Officer Deep Dive</button>
</div>

<div class="content">

  <!-- Tab 0: Dashboard -->
  <div class="tab-panel active" id="tab0">
    <h2>Recovery &amp; Incentive Dashboard &mdash; {period}</h2>
    <div class="banner" id="dashBanner">
      {B} currently uses the same figures as {A}. Open the &#9881; settings to enter {B}&rsquo;s
      multiplier and incentive slab, or set ALT_QUALITY_SLABS / ALT_INCENTIVE_SLABS in the script.
    </div>
    <div class="stats" id="dashStats"></div>

    <div class="card" style="padding:0">
      <div class="card-head">Month-wise recovery and incentive</div>
      <div class="overflow-x">
        <table>
          <thead>
            <tr class="grp">
              <th colspan="4"></th>
              <th class="th-ex bl">Existing</th>
              <th class="th-a bl" colspan="2" style="text-align:center">{A}</th>
              <th class="th-b bl" colspan="2" style="text-align:center">{B}</th>
            </tr>
            <tr>
              <th>Month</th>
              <th class="text-right">Officers</th>
              <th class="text-right">Eligible</th>
              <th class="text-right">Recovery</th>
              <th class="text-right th-ex bl">Incentive</th>
              <th class="text-right th-a bl">Incentive</th>
              <th class="text-right th-a">vs Existing</th>
              <th class="text-right th-b bl">Incentive</th>
              <th class="text-right th-b">vs Existing</th>
            </tr>
          </thead>
          <tbody id="dashBody"></tbody>
          <tfoot id="dashFoot"></tfoot>
        </table>
      </div>
    </div>

    <div class="flex-row">
      <div class="card flex-1" style="flex-grow:3">
        <div class="card-title">Incentive by month (bars) against recovery (dashed line)</div>
        <div class="chart-container" style="height:300px"><canvas id="dashChart"></canvas></div>
      </div>
      <div class="card flex-1" style="flex-grow:2;padding:0">
        <div class="card-head">Scenario comparison &mdash; full period</div>
        <div class="overflow-x">
          <table>
            <thead><tr>
              <th>Scenario</th>
              <th class="text-right">Incentive</th>
              <th class="text-right">% of Recovery</th>
              <th class="text-right">vs Existing</th>
              <th class="text-right">%</th>
              <th class="text-right">Benefited</th>
              <th class="text-right">Impacted</th>
            </tr></thead>
            <tbody id="dashCompare"></tbody>
          </table>
        </div>
        <div style="padding:12px 18px;font-size:12px;color:var(--muted)">
          Eligible = officers whose raw recovery reached the minimum that month.
          Benefited / Impacted counts officers over the full period.
        </div>
      </div>
    </div>
  </div>

  <!-- Tab 1: Model Overview -->
  <div class="tab-panel" id="tab1">
    <h2>Quality-Based Incentive Model</h2>
    <p style="color:var(--muted);max-width:840px;margin-bottom:10px">
      Two officers can recover the same amount and be doing very different work. One collects
      150,000 against a 300,000 balance &mdash; half the debt cleared. The other collects the same
      150,000 against a 1,000,000 balance, leaving 85% still outstanding. The existing model pays
      both identically.
    </p>
    <p style="color:var(--muted);max-width:840px;margin-bottom:22px">
      The proposed model keeps recovered value as the basis for payment and weights it by how much
      of the outstanding was actually cleared, so the officer closing more of their book earns more
      for the same collection. {A} and {B} are two sets of figures for the same method.
    </p>

    <div class="card" style="border-left:4px solid var(--teal);background:var(--teal-light);border-color:var(--teal)">
      <div style="font-size:13px;color:var(--teal-dark);font-weight:600;margin-bottom:6px">In one line</div>
      <div style="font-size:15px;color:var(--text)">
        Monthly Recovery &times; Quality Multiplier = Adjusted Recovery, and the incentive slab is
        read against the adjusted recovery instead of actual recovery.
      </div>
    </div>

    <div class="card">
      <h3>How It Works</h3>
      <div class="step-row"><div class="step-num">1</div><div>
        <div class="step-title">Measure quality on each case</div>
        <div class="step-desc">Cumulative Recovery &divide; Original Outstanding.</div></div></div>

      <div class="step-row"><div class="step-num">2</div><div>
        <div class="step-title">Weighted Average across the officer&rsquo;s cases</div>
        <div class="step-desc">Total Cumulative Recovery &divide; Total Outstanding for every case worked that month.</div></div></div>

      <div class="step-row"><div class="step-num">3</div><div>
        <div class="step-title">Apply the quality multiplier</div>
        <div class="step-desc">Monthly Recovery &times; Multiplier = Adjusted Recovery.</div></div></div>

      <div class="step-row"><div class="step-num">4</div><div>
        <div class="step-title">Read the incentive off the slab</div>
        <div class="step-desc">The proposed slab is applied to Adjusted Recovery rather than actual recovery.</div></div></div>

      <div class="step-row"><div class="step-num" style="background:var(--gold)">&#9888;</div><div>
        <div class="step-title">Note 1</div>
        <div class="step-desc">
          An officer whose <strong>actual</strong> recovery for the month is below
          <strong id="thresholdA">75,000</strong> earns nothing, exactly as today. The multiplier is
          not applied at all.
        </div></div></div>

      <div class="step-row"><div class="step-num" style="background:var(--gold)">&#9888;</div><div>
        <div class="step-title">Note 2</div>
        <div class="step-desc" id="note2Text"></div></div></div>
    </div>

    <div class="card">
      <h3>The Case In Numbers</h3>
      <p style="font-size:13px;color:var(--muted);margin:-4px 0 14px">
        Both officers recover 150,000 in the month. Only the size of the balance they recovered it from differs.
      </p>
      <div class="overflow-x">
        <table>
          <thead>
            <tr class="grp">
              <th colspan="5"></th>
              <th class="th-a bl" colspan="3" style="text-align:center">{A}</th>
              <th class="th-b bl" colspan="3" style="text-align:center">{B}</th>
            </tr>
            <tr>
              <th>Officer</th>
              <th class="text-right">Recovered</th>
              <th class="text-right">Outstanding</th>
              <th class="text-right">Quality</th>
              <th class="text-right">Existing</th>
              <th class="text-right th-a bl">Mult.</th>
              <th class="text-right th-a">Adjusted</th>
              <th class="text-right th-a">Incentive</th>
              <th class="text-right th-b bl">Mult.</th>
              <th class="text-right th-b">Adjusted</th>
              <th class="text-right th-b">Incentive</th>
            </tr>
          </thead>
          <tbody id="caseExample"></tbody>
        </table>
      </div>
      <div id="caseExampleNote" style="margin-top:14px;padding:12px 16px;background:var(--teal-light);
        border-radius:8px;font-size:13px;color:var(--teal-dark)"></div>
    </div>

    <div class="flex-row">
      <div class="flex-1">
        <div class="card" style="border-color:var(--teal)">
          <h3 style="color:var(--teal)">Quality Multiplier &mdash; {A}</h3>
          <table class="slab-table" style="margin-top:8px">
            <thead><tr><th>Quality Range</th><th style="text-align:center">Multiplier</th><th>Tier</th></tr></thead>
            <tbody id="qualitySlabBody"></tbody>
          </table>
        </div>
        <div class="card" style="border-color:var(--amber)">
          <h3 style="color:var(--amber)">Quality Multiplier &mdash; {B}</h3>
          <table class="slab-table" style="margin-top:8px">
            <thead><tr><th class="th-b">Quality Range</th><th class="th-b" style="text-align:center">Multiplier</th><th class="th-b">Tier</th></tr></thead>
            <tbody id="qualitySlabBodyB"></tbody>
          </table>
        </div>
      </div>
      <div class="flex-1">
        <div class="card">
          <h3>Existing Slab <span style="font-weight:400;color:var(--muted);font-size:12px">&middot; on raw recovery</span></h3>
          <table class="slab-table"><tbody id="existingSlabBody"></tbody></table>
        </div>
        <div class="card" style="border-color:var(--teal)">
          <h3 style="color:var(--teal)">{A} Slab <span style="font-weight:400;color:var(--muted);font-size:12px">&middot; on adjusted recovery</span></h3>
          <table class="slab-table"><tbody id="proposedSlabBody"></tbody></table>
          <div class="note-box" id="eligNoteA"></div>
        </div>
        <div class="card" style="border-color:var(--amber)">
          <h3 style="color:var(--amber)">{B} Slab <span style="font-weight:400;color:var(--muted);font-size:12px">&middot; on adjusted recovery</span></h3>
          <table class="slab-table"><tbody id="proposedBSlabBody"></tbody></table>
          <div class="note-box" id="eligNoteB"></div>
        </div>
      </div>
    </div>

    <div class="card">
      <h3>Interactive Simulator</h3>
      <p style="font-size:13px;color:var(--muted);margin-bottom:18px">
        Move the sliders to see how quality and recovery play out under each model.</p>
      <div class="flex-row">
        <div class="flex-1" style="max-width:420px">
          <div class="slider-row">
            <label><span>Weighted Quality</span><span id="simQualityLabel">47%</span></label>
            <input type="range" min="0" max="100" value="47" id="simQuality" oninput="updateSimulator()">
          </div>
          <div class="slider-row">
            <label><span>Monthly Recovery</span><span id="simRecoveryLabel">170,000</span></label>
            <input type="range" min="10000" max="800000" step="1000" value="170000" id="simRecovery" oninput="updateSimulator()">
          </div>
        </div>
        <div class="flex-1 overflow-x">
          <table>
            <thead><tr><th>Model</th><th>Tier</th><th class="text-right">Recovery used</th>
              <th class="text-right">Incentive</th><th class="text-right">vs Existing</th></tr></thead>
            <tbody id="simResult"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- Tab 2: Monthly Performance -->
  <div class="tab-panel" id="tab2">
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:20px">
      <h2 style="margin:0">Monthly Performance</h2>
      <div style="display:flex;gap:6px;margin-left:auto;flex-wrap:wrap" id="monthButtons"></div>
    </div>
    <div class="stats" id="monthStats"></div>
    <div class="filter-bar">
      <input class="filter-input" id="officerFilter" placeholder="Search officer name or code" oninput="renderMonthly()">
      <span style="font-size:13px;color:var(--muted);margin-left:auto">Sort</span>
      <span style="display:flex;gap:6px">
        <button class="btn active" onclick="sortMonthly('recovery',this)">Recovery</button>
        <button class="btn" onclick="sortMonthly('quality',this)">Quality</button>
        <button class="btn" onclick="sortMonthly('diff',this)">{A} impact</button>
        <button class="btn" onclick="sortMonthly('diffB',this)">{B} impact</button>
      </span>
    </div>
    <div class="card" style="padding:0">
      <div class="overflow-x">
        <table>
          <thead>
            <tr class="grp">
              <th colspan="6"></th>
              <th class="th-ex bl">Existing</th>
              <th class="th-a bl" colspan="4" style="text-align:center">{A}</th>
              <th class="th-b bl" colspan="4" style="text-align:center">{B}</th>
            </tr>
            <tr>
              <th class="text-right">#</th><th>Officer</th><th>Region</th>
              <th class="text-right">Cases</th>
              <th class="text-right">Recovery</th>
              <th class="text-right">Avg Quality</th>
              <th class="text-right th-ex bl">Incentive</th>
              <th class="text-right th-a bl">Mult.</th>
              <th class="text-right th-a">Adjusted</th>
              <th class="text-right th-a">Incentive</th>
              <th class="text-right th-a">vs Existing</th>
              <th class="text-right th-b bl">Mult.</th>
              <th class="text-right th-b">Adjusted</th>
              <th class="text-right th-b">Incentive</th>
              <th class="text-right th-b">vs Existing</th>
            </tr>
          </thead>
          <tbody id="monthlyBody"></tbody>
          <tfoot id="monthlyFoot"></tfoot>
        </table>
      </div>
    </div>
  </div>

  <!-- Tab 3: Impact Analysis -->
  <div class="tab-panel" id="tab3">
    <h2>Impact Analysis &mdash; {period}</h2>
    <div class="stats" id="impactStats"></div>
    <div class="flex-row">
      <div class="card flex-1">
        <div class="card-title">Officers benefited / impacted vs existing</div>
        <div class="chart-container"><canvas id="impactChart"></canvas></div>
      </div>
      <div class="card flex-1">
        <div class="card-title">Monthly incentive trend</div>
        <div class="chart-container"><canvas id="trendChart"></canvas></div>
      </div>
    </div>
    <div class="card" style="padding:0">
      <div class="card-head" style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
        Officer Summary &mdash; full period
        <span style="margin-left:auto;display:flex;gap:6px">
          <button class="btn active" onclick="sortImpact('diff',this)">Sort by {A}</button>
          <button class="btn" onclick="sortImpact('diffB',this)">Sort by {B}</button>
        </span>
      </div>
      <div class="overflow-x">
        <table>
          <thead>
            <tr class="grp">
              <th colspan="5"></th>
              <th class="th-ex bl">Existing</th>
              <th class="th-a bl" colspan="3" style="text-align:center">{A}</th>
              <th class="th-b bl" colspan="3" style="text-align:center">{B}</th>
            </tr>
            <tr>
              <th class="text-right">#</th><th>Officer</th><th>Region</th>
              <th class="text-right">Avg Quality</th>
              <th class="text-right">Total Recovery</th>
              <th class="text-right th-ex bl">Incentive</th>
              <th class="text-right th-a bl">Incentive</th>
              <th class="text-right th-a">vs Existing</th>
              <th class="th-a">Status</th>
              <th class="text-right th-b bl">Incentive</th>
              <th class="text-right th-b">vs Existing</th>
              <th class="th-b">Status</th>
            </tr>
          </thead>
          <tbody id="impactBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Tab 4: Officer Deep Dive -->
  <div class="tab-panel" id="tab4">
    <h2>Officer Deep Dive</h2>
    <p style="color:var(--muted);margin-bottom:16px">
      Search for an officer to see their month-by-month case detail.</p>
    <div class="search-select">
      <input type="text" id="officerSearch" placeholder="Search officer name or code"
             oninput="filterOfficers()" onfocus="filterOfficers()" autocomplete="off">
      <div class="search-list" id="officerList"></div>
    </div>
    <div id="deepDiveContent" style="margin-top:24px"></div>
  </div>

</div>

<script>
{payload}
{DASHBOARD_JS}
</script>
</body>
</html>"""

    out_path = os.path.join(os.path.dirname(os.path.abspath(input_file)), 'incentive_dashboard.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  Dashboard saved: {os.path.basename(out_path)}")
    return out_path


# ==============================================================
# MAIN
# ==============================================================

def main(input_file):
    df, df_cases, df_officer, df_overall, case_timeline = process_data(input_file)

    print()
    excel_path = write_excel(input_file, df_cases, df_officer, df_overall)
    html_path = generate_html(input_file, df_officer, df_overall, case_timeline)

    ms = monthly_summary(df_officer)
    a_col, b_col = f'{LABEL_A} Incentive', f'{LABEL_B} Incentive'

    print(f"\n{'=' * 74}")
    print(f"  {'Month':<8}{'Recovery':>16}{'Existing':>16}{LABEL_A:>16}{LABEL_B:>16}")
    print(f"{'-' * 74}")
    for _, r in ms.iterrows():
        if r['Month'] == 'TOTAL':
            print(f"{'-' * 74}")
        print(f"  {r['Month']:<8}{int(r['Total Recovery']):>16,}{int(r['Existing Incentive']):>16,}"
              f"{int(r[a_col]):>16,}{int(r[b_col]):>16,}")
    t = ms.iloc[-1]
    print(f"  {'vs Existing':<24}{'':>16}{int(t[a_col] - t['Existing Incentive']):>+16,}"
          f"{int(t[b_col] - t['Existing Incentive']):>+16,}")
    print(f"{'=' * 74}\n")
    print("  Open incentive_dashboard.html in a browser to review.\n")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('\nUsage: python incentive_model.py "recovery_data.xlsx"\n')
        sys.exit(1)

    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"\nERROR — file not found: {path}\n")
        sys.exit(1)

    main(path)
