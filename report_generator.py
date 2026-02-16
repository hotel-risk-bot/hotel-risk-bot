#!/usr/bin/env python3
"""
Enhanced Executive Claims Report Generator v2
Generates professional PDF reports with charts, loss ratios, and detailed analysis.
All charts split by Property / Liability. Improved colors, avg lines, and layout.
"""

import io
import os
import re
import tempfile
import unicodedata
from datetime import datetime, timedelta
from collections import defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from fpdf import FPDF
import requests as http_requests

# ── Configuration ──────────────────────────────────────────────────────────
AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")
AIRTABLE_API_URL = "https://api.airtable.com/v0"
CONSULTING_BASE_ID = "appOVp1eJUPbNgNXM"
POLICIES_TABLE_ID = "tblP5P6SZG1uIRaTt"

# Colors
HUB_DARK_BLUE = (0, 51, 102)
HUB_ELECTRIC_BLUE = (0, 102, 204)
HUB_LIGHT_BLUE = (200, 220, 240)
HUB_RED_LIGHT = (255, 220, 220)
HUB_WHITE = (255, 255, 255)
HUB_GRAY = (128, 128, 128)
HUB_DARK = (40, 40, 40)

# Distinct chart palette - lighter colors for better text readability
PROP_COLORS = ["#5B9BD5", "#70AD47", "#FFC000", "#ED7D31", "#A5A5A5",
               "#4472C4", "#9DC3E6", "#A9D18E", "#F4B183", "#BDD7EE"]
LIAB_COLORS = ["#E88E8E", "#F4B183", "#B4A7D6", "#87CEEB", "#70AD47",
               "#FFC000", "#D5A6BD", "#A9D18E", "#9DC3E6", "#C9DAF8"]
# Lighter colors for bars in location/trending charts
PROP_BAR_COLOR = "#5B9BD5"
LIAB_BAR_COLOR = "#E88E8E"
# Light highlight colors for development
DEV_INCREASE_COLOR = "#FFCCCC"  # light red
DEV_DECREASE_COLOR = "#CCFFCC"  # light green
DEV_INCREASE_CHART = "#E88E8E"  # light red for chart bars
DEV_DECREASE_CHART = "#8ED18E"  # light green for chart bars


def sanitize_for_pdf(text):
    if not text or not isinstance(text, str):
        return str(text) if text else ""
    replacements = {
        '\u2013': '-', '\u2014': '--', '\u2018': "'", '\u2019': "'",
        '\u201c': '"', '\u201d': '"', '\u2026': '...', '\u2022': '*',
        '\u00a0': ' ', '\u200b': '', '\u2010': '-', '\u2011': '-',
        '\u2012': '-', '\u00b7': '*', '\u2032': "'", '\u2033': '"',
        '\u00ae': '(R)', '\u2122': '(TM)', '\u00a9': '(C)',
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    result = []
    for ch in text:
        try:
            ch.encode('latin-1')
            result.append(ch)
        except UnicodeEncodeError:
            decomposed = unicodedata.normalize('NFKD', ch)
            ascii_chars = decomposed.encode('ascii', 'ignore').decode('ascii')
            result.append(ascii_chars if ascii_chars else '?')
    return ''.join(result)


def get_val(f, field_name, default="N/A"):
    val = f.get(field_name, default)
    if isinstance(val, list):
        return ", ".join(str(v) for v in val) if val else default
    return str(val) if val else default


# ── Airtable Policies Lookup ──────────────────────────────────────────────

def fetch_policies(client_name):
    safe_name = client_name.replace('"', '\\"')
    # Search across multiple fields to find policies for this client
    search_fields = [
        'FIND(LOWER("{name}"), LOWER(ARRAYJOIN({{Corporate Name}}, ",")))',
        'FIND(LOWER("{name}"), LOWER(ARRAYJOIN({{Policy Name}}, ",")))',
        'FIND(LOWER("{name}"), LOWER(ARRAYJOIN({{DBA (from Locations)}}, ",")))',
        'FIND(LOWER("{name}"), LOWER(ARRAYJOIN({{Client (from Locations)}}, ",")))',
        'FIND(LOWER("{name}"), LOWER(ARRAYJOIN({{Clients}}, ",")))',
    ]
    conditions = [sf.format(name=safe_name) for sf in search_fields]
    formula = f'OR({", ".join(conditions)})'
    headers = {"Authorization": f"Bearer {AIRTABLE_PAT}"}
    url = f"{AIRTABLE_API_URL}/{CONSULTING_BASE_ID}/{POLICIES_TABLE_ID}"
    all_records = []
    offset = None
    while True:
        params = {"filterByFormula": formula, "pageSize": 100}
        if offset:
            params["offset"] = offset
        try:
            resp = http_requests.get(url, headers=headers, params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            all_records.extend(data.get("records", []))
            offset = data.get("offset")
            if not offset:
                break
        except Exception as e:
            break
    return all_records


def build_loss_ratio_data(policies):
    """Build loss ratio data with individual policy rows grouped by year.
    Returns a list of policy dicts, each representing one policy record."""
    policy_rows = []
    for rec in policies:
        f = rec.get("fields", {})
        year = f.get("Policy Year", "N/A")
        raw_ptype = f.get("Policy Type", "N/A")
        if isinstance(raw_ptype, list):
            ptype = raw_ptype[0] if raw_ptype else "N/A"
        else:
            ptype = str(raw_ptype) if raw_ptype else "N/A"
        base_premium = f.get("Base Premium", 0) or 0
        incurred = f.get("Incurred", 0) or 0
        claim_count = f.get("Claim Count", 0) or 0
        carrier = get_val(f, "Carrier Name")
        policy_num = f.get("Policy #", "N/A")
        policy_name = get_val(f, "Policy Name", "")
        if ptype not in ("Property", "Liability"):
            continue
        loss_ratio = incurred / base_premium if base_premium > 0 else 0
        policy_rows.append({
            "year": str(year), "type": ptype,
            "base_premium": base_premium, "incurred": incurred,
            "claim_count": claim_count, "carrier": carrier,
            "policy_num": policy_num, "policy_name": policy_name,
            "loss_ratio": loss_ratio,
        })
    return policy_rows


# ── Parsers ───────────────────────────────────────────────────────────────

def parse_activity_comments(raw_data):
    if not raw_data:
        return []
    entries = raw_data.split("[[Break]]")
    comments = []
    for entry in entries:
        entry = entry.strip().strip(",").strip()
        if not entry or "Update/Comment:" not in entry:
            continue
        date_match = re.match(r'([\w]+\s+\d{1,2},\s+\d{4}\s+\d{1,2}:\d{2}\s*[AP]M)', entry)
        date_str = date_match.group(1) if date_match else "Unknown"
        comment_match = re.search(r'Update/Comment:\s*(.+?)(?:-{10,}|$)', entry, re.DOTALL)
        if comment_match:
            comment_text = comment_match.group(1).strip()
            if len(comment_text) > 10:
                comments.append({"date": date_str, "text": comment_text})
    return comments


def parse_claims_development(raw_data):
    if not raw_data:
        return []
    entries = raw_data.split("[[Break]]")
    valuations = []
    for entry in entries:
        entry = entry.strip().strip(",").strip()
        if not entry or ("Valuation" not in entry and "Total Incurred:" not in entry):
            continue
        date_match = re.match(r'([\w]+\s+\d{1,2},\s+\d{4})', entry)
        date_str = date_match.group(1) if date_match else "Unknown"
        paid_match = re.search(r'Paid:\s*\$?([\d,.]+)', entry)
        reserved_match = re.search(r'Reserved:\s*\$?([\d,.]+)', entry)
        expenses_match = re.search(r'Expenses:\s*\$?([\d,.]+)', entry)
        incurred_match = re.search(r'Total Incurred:\s*\$?([\d,.]+)', entry)

        def parse_amount(m):
            if m:
                try:
                    return float(m.group(1).replace(",", ""))
                except (ValueError, TypeError):
                    return 0.0
            return 0.0

        total_incurred = parse_amount(incurred_match)
        if total_incurred > 0 or paid_match or reserved_match:
            parsed_date = None
            for fmt in ("%B %d, %Y", "%b %d, %Y"):
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    break
                except (ValueError, TypeError):
                    pass
            valuations.append({
                "date": date_str, "parsed_date": parsed_date,
                "paid": parse_amount(paid_match), "reserved": parse_amount(reserved_match),
                "expenses": parse_amount(expenses_match), "total_incurred": total_incurred,
            })
    return valuations


def calculate_development_delta(valuations, months=15):
    if not valuations:
        return 0
    cutoff = datetime.now() - timedelta(days=months * 30)
    latest = valuations[-1]["total_incurred"] if valuations else 0
    baseline = 0
    for v in valuations:
        if v["parsed_date"] and v["parsed_date"] <= cutoff:
            baseline = v["total_incurred"]
    if baseline == 0 and valuations:
        first = valuations[0]
        if first["parsed_date"] and first["parsed_date"] > cutoff:
            baseline = 0
    return latest - baseline


# ── Chart Generators (all split by Property / Liability) ──────────────────

def _split_by_type(results):
    prop = [r for r in results if get_val(r["fields"], "Claim Type") == "Property"]
    liab = [r for r in results if get_val(r["fields"], "Claim Type") == "Liability"]
    return prop, liab


def create_incurred_by_type_year_chart(results, tmpdir, min_policy_year=None):
    """Split Property (top) and Liability (bottom) with avg, total incurred, and claim count.
    Always shows ALL years from min to max policy year, even if no claims in some years."""
    data = defaultdict(lambda: {"Property": {"incurred": 0, "count": 0}, "Liability": {"incurred": 0, "count": 0}})
    for r in results:
        f = r["fields"]
        year = get_val(f, "Policy Year", "Unknown")
        ctype = get_val(f, "Claim Type", "Unknown")
        if ctype in ("Property", "Liability"):
            data[year][ctype]["incurred"] += r["incurred"]
            data[year][ctype]["count"] += 1
    if not data:
        return None

    # Determine full year range - include all years even with 0 claims
    numeric_years = []
    for y in data.keys():
        try:
            numeric_years.append(int(y))
        except (ValueError, TypeError):
            pass
    if numeric_years:
        yr_min = min(numeric_years)
        yr_max = max(numeric_years)
        if min_policy_year and min_policy_year < yr_min:
            yr_min = min_policy_year
        current_year = datetime.now().year
        if yr_max < current_year:
            yr_max = current_year
        years = [str(y) for y in range(yr_min, yr_max + 1)]
    else:
        years = sorted(data.keys())

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 10))

    for ax, ptype, color in [
        (ax1, "Property", PROP_BAR_COLOR),
        (ax2, "Liability", LIAB_BAR_COLOR),
    ]:
        incurred = [data[y][ptype]["incurred"] for y in years]
        counts = [data[y][ptype]["count"] for y in years]
        x = np.arange(len(years))

        bars = ax.bar(x, incurred, color=color, alpha=0.8, label="Total Incurred")
        ax.set_xlabel("Policy Year")
        ax.set_ylabel("Total Incurred ($)")
        ax.set_title(f"{ptype} - Incurred Losses by Policy Year")
        ax.set_xticks(x)
        ax.set_xticklabels(years)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: f"${x:,.0f}"))

        # Avg incurred line
        avg_inc = np.mean(incurred) if incurred else 0
        ax.axhline(y=avg_inc, color="orange", linewidth=1.5, linestyle="--", label=f"Avg Incurred: ${avg_inc:,.0f}")

        # Total incurred annotation
        total_inc = sum(incurred)
        ax.text(0.98, 0.95, f"Total: ${total_inc:,.0f}", transform=ax.transAxes,
                ha="right", va="top", fontsize=9, fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", edgecolor="gray"))

        # Set y-axis minimum to 0
        ax.set_ylim(bottom=0)

        # Annotate bars with incurred amount
        for bar, inc_val in zip(bars, incurred):
            h = bar.get_height()
            if h > 0:
                ax.annotate(f"${h:,.0f}", xy=(bar.get_x() + bar.get_width()/2, h),
                           xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=7)

        # Claim count on secondary axis - DARK BLUE
        ax_twin = ax.twinx()
        ax_twin.plot(x, counts, color="#003366", marker="o", linewidth=2, label="Claim Count")
        ax_twin.set_ylabel("Claim Count", color="#003366")
        ax_twin.set_ylim(bottom=0)  # Ensure count axis starts at 0
        avg_count = np.mean(counts) if counts else 0
        ax_twin.axhline(y=avg_count, color="#003366", linewidth=1, linestyle=":", alpha=0.6, label=f"Avg Count: {avg_count:.1f}")

        lines1, labels1 = ax.get_legend_handles_labels()
        lines2, labels2 = ax_twin.get_legend_handles_labels()
        ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=7)

    plt.tight_layout()
    path = os.path.join(tmpdir, "incurred_by_type_year.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def create_location_impact_chart(results, tmpdir):
    """Split by Property (top) and Liability (bottom) with avg incurred line."""
    prop_claims, liab_claims = _split_by_type(results)

    def _build_loc_data(claims):
        loc = defaultdict(lambda: {"incurred": 0, "count": 0})
        for r in claims:
            f = r["fields"]
            p = get_val(f, "DBA (from Location)", "Unknown")
            loc[p]["incurred"] += r["incurred"]
            loc[p]["count"] += 1
        return loc

    prop_loc = _build_loc_data(prop_claims)
    liab_loc = _build_loc_data(liab_claims)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 10))

    for ax, loc_data, title, color in [
        (ax1, prop_loc, "Property - Location Impact", PROP_BAR_COLOR),
        (ax2, liab_loc, "Liability - Location Impact", LIAB_BAR_COLOR),
    ]:
        if not loc_data:
            ax.text(0.5, 0.5, "No claims in filtered results", ha="center", va="center",
                   transform=ax.transAxes, fontsize=10, color="gray", style="italic")
            ax.set_title(title)
            ax.set_xticks([])
            ax.set_yticks([])
            for spine in ax.spines.values():
                spine.set_visible(False)
            continue
        sorted_locs = sorted(loc_data.items(), key=lambda x: x[1]["incurred"], reverse=True)[:10]
        labels = [sanitize_for_pdf(l[0][:28]) for l in sorted_locs]
        incurred = [l[1]["incurred"] for l in sorted_locs]
        counts = [l[1]["count"] for l in sorted_locs]
        x = np.arange(len(labels))

        bars = ax.bar(x, incurred, color=color, alpha=0.8, label="Total Incurred")
        ax.set_ylabel("Total Incurred ($)")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: f"${x:,.0f}"))

        # Average incurred line
        avg_inc = np.mean(incurred) if incurred else 0
        ax.axhline(y=avg_inc, color="orange", linewidth=1.5, linestyle="--", label=f"Avg: ${avg_inc:,.0f}")

        ax2_twin = ax.twinx()
        ax2_twin.plot(x, counts, color="#003366", marker="o", linewidth=2, label="Claim Count")
        ax2_twin.set_ylabel("Claim Count", color="#003366")

        lines1, labels1 = ax.get_legend_handles_labels()
        lines2, labels2 = ax2_twin.get_legend_handles_labels()
        ax.legend(lines1 + lines2, labels1 + labels2, loc="upper right", fontsize=7)
        ax.set_title(title)

    plt.tight_layout()
    path = os.path.join(tmpdir, "location_impact.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def create_cause_of_loss_chart(results, tmpdir):
    """Split by Property (top) and Liability (bottom). Include hazard in labels. Better colors."""
    prop_claims, liab_claims = _split_by_type(results)

    def _build_cause_data(claims):
        cause = defaultdict(lambda: {"incurred": 0, "count": 0, "hazards": set()})
        for r in claims:
            f = r["fields"]
            c = get_val(f, "Cause of Loss Rollup Output", "Unknown")
            if c == "N/A":
                c = get_val(f, "Cause of Loss (from Cause of Loss)", "Unknown")
            h = get_val(f, "Risk/Hazard (From Risk/Hazard)", "")
            cause[c]["incurred"] += r["incurred"]
            cause[c]["count"] += 1
            if h and h != "N/A":
                cause[c]["hazards"].add(h)
        return cause

    prop_cause = _build_cause_data(prop_claims)
    liab_cause = _build_cause_data(liab_claims)

    fig, axes = plt.subplots(2, 2, figsize=(14, 10))

    for row, (cause_data, type_name, palette) in enumerate([
        (prop_cause, "Property", PROP_COLORS),
        (liab_cause, "Liability", LIAB_COLORS),
    ]):
        ax_pie = axes[row][0]
        ax_bar = axes[row][1]

        if not cause_data:
            ax_pie.text(0.5, 0.5, "No claims in filtered results", ha="center", va="center",
                       transform=ax_pie.transAxes, fontsize=10, color="gray", style="italic")
            ax_bar.text(0.5, 0.5, "No claims in filtered results", ha="center", va="center",
                       transform=ax_bar.transAxes, fontsize=10, color="gray", style="italic")
            ax_pie.set_title(f"{type_name} - Incurred by Cause of Loss")
            ax_bar.set_title(f"{type_name} - Claim Frequency by Cause")
            # Remove axes for cleaner look
            ax_pie.set_xticks([])
            ax_pie.set_yticks([])
            ax_bar.set_xticks([])
            ax_bar.set_yticks([])
            for spine in ax_pie.spines.values():
                spine.set_visible(False)
            for spine in ax_bar.spines.values():
                spine.set_visible(False)
            continue

        sorted_causes = sorted(cause_data.items(), key=lambda x: x[1]["incurred"], reverse=True)[:8]
        # Build labels with top hazard
        labels = []
        for c_name, c_data in sorted_causes:
            hazards = list(c_data["hazards"])[:2]
            lbl = sanitize_for_pdf(c_name[:22])
            if hazards:
                lbl += f" ({sanitize_for_pdf(', '.join(hazards)[:20])})"
            labels.append(lbl)

        incurred = [c[1]["incurred"] for c in sorted_causes]
        counts = [c[1]["count"] for c in sorted_causes]
        colors = palette[:len(labels)]

        wedges, texts, autotexts = ax_pie.pie(
            incurred, labels=None, autopct="%1.0f%%",
            colors=colors, startangle=90, pctdistance=0.8
        )
        for at in autotexts:
            at.set_fontsize(8)
        ax_pie.set_title(f"{type_name} - Incurred by Cause of Loss")
        ax_pie.legend(labels, loc="lower left", fontsize=6)

        x = np.arange(len(labels))
        ax_bar.barh(x, counts, color=colors)
        ax_bar.set_yticks(x)
        ax_bar.set_yticklabels(labels, fontsize=6)
        ax_bar.set_xlabel("Claim Count")
        ax_bar.set_title(f"{type_name} - Claim Frequency by Cause")
        ax_bar.invert_yaxis()

    plt.tight_layout()
    path = os.path.join(tmpdir, "cause_of_loss.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def create_claim_trending_chart(results, tmpdir, min_policy_year=None):
    """Split Property (top) and Liability (bottom). Shows incurred with avg line + claim count with avg line.
    Always shows ALL years from min to max policy year, even if no claims in some years."""
    data = defaultdict(lambda: {"Property": {"incurred": 0, "count": 0}, "Liability": {"incurred": 0, "count": 0}})
    for r in results:
        f = r["fields"]
        year = get_val(f, "Policy Year", "Unknown")
        ctype = get_val(f, "Claim Type", "Unknown")
        if ctype in ("Property", "Liability"):
            data[year][ctype]["incurred"] += r["incurred"]
            data[year][ctype]["count"] += 1
    if not data:
        return None

    # Determine full year range - include all years even with 0 claims
    numeric_years = []
    for y in data.keys():
        try:
            numeric_years.append(int(y))
        except (ValueError, TypeError):
            pass
    if numeric_years:
        yr_min = min(numeric_years)
        yr_max = max(numeric_years)
        # If min_policy_year specified, use it as floor
        if min_policy_year and min_policy_year < yr_min:
            yr_min = min_policy_year
        # Extend to current year if needed
        current_year = datetime.now().year
        if yr_max < current_year:
            yr_max = current_year
        years = [str(y) for y in range(yr_min, yr_max + 1)]
    else:
        years = sorted(data.keys())

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 10))

    for ax, ptype, color in [
        (ax1, "Property", PROP_BAR_COLOR),
        (ax2, "Liability", LIAB_BAR_COLOR),
    ]:
        incurred = [data[y][ptype]["incurred"] for y in years]
        counts = [data[y][ptype]["count"] for y in years]
        x = np.arange(len(years))

        # Incurred bars
        bars = ax.bar(x, incurred, color=color, alpha=0.7, label="Incurred")
        ax.set_xlabel("Policy Year")
        ax.set_ylabel("Total Incurred ($)")
        ax.set_title(f"{ptype} - Claim Trending by Policy Year")
        ax.set_xticks(x)
        ax.set_xticklabels(years)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: f"${x:,.0f}"))
        ax.grid(True, alpha=0.3, axis="y")

        # Avg incurred line
        avg_inc = np.mean(incurred) if incurred else 0
        ax.axhline(y=avg_inc, color="orange", linewidth=1.5, linestyle="--", label=f"Avg Incurred: ${avg_inc:,.0f}")

        # Set y-axis minimum to 0
        ax.set_ylim(bottom=0)

        # Annotate incurred on bars
        for bar, inc_val in zip(bars, incurred):
            h = bar.get_height()
            if h > 0:
                ax.annotate(f"${h:,.0f}", xy=(bar.get_x() + bar.get_width()/2, h),
                           xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=7)

        # Claim count on secondary axis - DARK BLUE
        ax_twin = ax.twinx()
        ax_twin.plot(x, counts, color="#003366", marker="o", linewidth=2, label="Claim Count")
        ax_twin.set_ylabel("Claim Count", color="#003366")
        ax_twin.set_ylim(bottom=0)  # Ensure count axis starts at 0

        # Avg claim count line
        avg_count = np.mean(counts) if counts else 0
        ax_twin.axhline(y=avg_count, color="#003366", linewidth=1, linestyle=":", alpha=0.6, label=f"Avg Count: {avg_count:.1f}")

        # Annotate counts
        for i, c in enumerate(counts):
            if c > 0:
                ax_twin.annotate(str(c), (x[i], c), textcoords="offset points", xytext=(0, 8), ha="center", fontsize=8, color="#003366")

        lines1, labels1 = ax.get_legend_handles_labels()
        lines2, labels2 = ax_twin.get_legend_handles_labels()
        ax.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=7)

    plt.tight_layout()
    path = os.path.join(tmpdir, "claim_trending.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def create_development_chart(results, tmpdir):
    """Claims with most development in last 15 months. Includes claimant, DOL, corp, DBA-city, cause, dev vs total."""
    dev_data = []
    for r in results:
        f = r["fields"]
        raw_activity = f.get("Activity Rollup Raw Data", "")
        if not raw_activity:
            continue
        valuations = parse_claims_development(raw_activity)
        delta = calculate_development_delta(valuations, months=15)
        if abs(delta) > 0:
            claim_num = get_val(f, "Claim #", "Unknown")
            claimant = get_val(f, "Involved Party (From Involved Party)", "")
            if claimant == "N/A":
                claimant = get_val(f, "Involved Party copy", "")
            dol = get_val(f, "Incident Date", "")
            if dol == "N/A" or not dol:
                dol = get_val(f, "DOL", "")
            corp = get_val(f, "Corporate Name", "")
            prop = get_val(f, "DBA (from Location)", "")
            cause = get_val(f, "Cause of Loss Rollup Output", "")
            if cause == "N/A":
                cause = get_val(f, "Cause of Loss (from Cause of Loss)", "")
            hazard = get_val(f, "Risk/Hazard (From Risk/Hazard)", "")

            # Build cause + hazard label
            cause_hazard = sanitize_for_pdf(cause[:20])
            if hazard and hazard != "N/A":
                cause_hazard = sanitize_for_pdf(f"{cause[:18]} / {hazard[:15]}")

            label = sanitize_for_pdf(claim_num[-10:]) if len(claim_num) > 10 else sanitize_for_pdf(claim_num)
            dev_data.append({
                "label": label, "delta": delta, "current": r["incurred"],
                "claimant": sanitize_for_pdf(claimant[:25]),
                "dol": sanitize_for_pdf(dol[:10]),
                "corp": sanitize_for_pdf(corp[:25]),
                "prop": sanitize_for_pdf(prop[:25]),
                "cause": cause_hazard,
            })
    if not dev_data:
        return None, []

    dev_data.sort(key=lambda x: abs(x["delta"]), reverse=True)
    dev_data = dev_data[:12]

    labels = [d["label"] for d in dev_data]
    deltas = [d["delta"] for d in dev_data]
    colors = [DEV_INCREASE_CHART if d > 0 else DEV_DECREASE_CHART for d in deltas]

    fig, ax = plt.subplots(figsize=(11, 5))
    x = np.arange(len(labels))
    bars = ax.bar(x, deltas, color=colors)
    ax.set_xlabel("Claim #")
    ax.set_ylabel("Incurred Change ($)")
    ax.set_title("Claims Development - Last 15 Months (Increase / Decrease)")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: f"${x:,.0f}"))
    ax.axhline(y=0, color="black", linewidth=0.5)

    for bar, delta in zip(bars, deltas):
        h = bar.get_height()
        va = "bottom" if h >= 0 else "top"
        off = 3 if h >= 0 else -3
        ax.annotate(f"${delta:+,.0f}", xy=(bar.get_x() + bar.get_width()/2, h),
                   xytext=(0, off), textcoords="offset points", ha="center", va=va, fontsize=7)

    plt.tight_layout()
    path = os.path.join(tmpdir, "development_chart.png")
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path, dev_data


# ── PDF Report Class ──────────────────────────────────────────────────────

class ExecutiveReportPDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(*HUB_DARK_BLUE)
        self.cell(0, 8, "Risk Analysis Report", ln=True, align="L")
        self.set_draw_color(*HUB_ELECTRIC_BLUE)
        self.set_line_width(0.5)
        # Extend line to full page width (handles both portrait and landscape)
        page_w = self.w - self.r_margin
        self.line(10, self.get_y(), page_w, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(*HUB_GRAY)
        self.cell(0, 10,
                  f"Confidential  |  Page {self.page_no()}/{{nb}}  |  Generated {datetime.now().strftime('%m/%d/%Y')}",
                  align="C")

    def _reset_margins(self):
        """Reset left margin to default for consistent content alignment."""
        self.set_left_margin(10)
        self.set_x(10)


# ── Main Report Generator ─────────────────────────────────────────────────

def generate_executive_pdf(client_name, results, query_params):
    tmpdir = tempfile.mkdtemp(prefix="report_")

    # Fetch policies for loss ratio
    policies = fetch_policies(client_name)
    lr_data = build_loss_ratio_data(policies)

    # Filter loss ratio data by policy year if specified
    min_py = query_params.get("min_policy_year")
    if min_py is not None:
        lr_data = [p for p in lr_data
                   if str(p["year"]).isdigit() and int(p["year"]) >= min_py]

    # Split results
    property_claims = sorted(
        [r for r in results if get_val(r["fields"], "Claim Type") == "Property"],
        key=lambda x: x["incurred"], reverse=True)
    liability_claims = sorted(
        [r for r in results if get_val(r["fields"], "Claim Type") == "Liability"],
        key=lambda x: x["incurred"], reverse=True)
    other_claims = [r for r in results if get_val(r["fields"], "Claim Type") not in ("Property", "Liability")]

    # Pre-compute 15-month development for each claim
    dev_lookup = {}
    for r in results:
        raw = r["fields"].get("Activity Rollup Raw Data", "")
        if raw:
            vals = parse_claims_development(raw)
            delta = calculate_development_delta(vals, months=15)
            dev_lookup[r.get("record_id", id(r))] = delta

    # Generate charts
    chart_incurred = create_incurred_by_type_year_chart(results, tmpdir, min_policy_year=min_py)
    chart_location = create_location_impact_chart(results, tmpdir)
    chart_cause = create_cause_of_loss_chart(results, tmpdir)
    chart_trending = create_claim_trending_chart(results, tmpdir, min_policy_year=min_py)
    dev_result = create_development_chart(results, tmpdir)
    if isinstance(dev_result, tuple):
        chart_development, dev_table_data = dev_result
    else:
        chart_development, dev_table_data = dev_result, []

    # Build PDF
    pdf = ExecutiveReportPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # ── Title ──
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(*HUB_DARK_BLUE)
    pdf.cell(0, 14, "Executive Claims Report", ln=True)
    pdf.set_font("Helvetica", "", 16)
    pdf.set_text_color(*HUB_DARK)
    pdf.cell(0, 10, sanitize_for_pdf(f"Client: {client_name}"), ln=True)
    pdf.ln(2)

    # Filters
    pdf.set_font("Helvetica", "I", 10)
    pdf.set_text_color(80, 80, 80)
    filter_parts = []
    if query_params.get("status"):
        filter_parts.append(f"Status: {query_params['status'].title()}")
    if query_params.get("claim_type"):
        filter_parts.append(f"Type: {query_params['claim_type'].title()}")
    if query_params.get("min_incurred") is not None:
        filter_parts.append(f"Min Incurred: ${query_params['min_incurred']:,.0f}")
    if query_params.get("max_incurred") is not None:
        filter_parts.append(f"Max Incurred: ${query_params['max_incurred']:,.0f}")
    if query_params.get("min_policy_year") is not None:
        filter_parts.append(f"Policy Year >= {query_params['min_policy_year']}")
    if filter_parts:
        pdf.cell(0, 6, "Filters: " + " | ".join(filter_parts), ln=True)
    pdf.cell(0, 6, f"Report Date: {datetime.now().strftime('%B %d, %Y')}", ln=True)
    pdf.ln(4)

    # ── Executive Summary Box ──
    total_incurred = sum(r["incurred"] for r in results)
    total_paid = total_reserved = open_count = closed_count = attorney_count = 0
    prop_open = prop_closed = liab_open = liab_closed = 0
    prop_inc = sum(r["incurred"] for r in property_claims)
    liab_inc = sum(r["incurred"] for r in liability_claims)
    prop_paid = prop_rsv = liab_paid = liab_rsv = 0

    for r in results:
        flds = r["fields"]
        ctype = get_val(flds, "Claim Type")
        is_open = flds.get("Status") == "Open"
        if is_open:
            open_count += 1
            if ctype == "Property": prop_open += 1
            elif ctype == "Liability": liab_open += 1
        else:
            closed_count += 1
            if ctype == "Property": prop_closed += 1
            elif ctype == "Liability": liab_closed += 1
        if flds.get("Attorney Representation"):
            attorney_count += 1
        p = flds.get("Paid - Rollup", 0)
        try:
            pv = float(p) if p else 0
        except (ValueError, TypeError):
            pv = 0
        total_paid += pv
        if ctype == "Property": prop_paid += pv
        elif ctype == "Liability": liab_paid += pv

        rv = flds.get("Reserved Helper", [])
        if isinstance(rv, list) and rv:
            try:
                rv_val = float(rv[-1])
            except (ValueError, TypeError):
                rv_val = 0
        else:
            try:
                rv_val = float(rv) if rv else 0
            except (ValueError, TypeError):
                rv_val = 0
        total_reserved += rv_val
        if ctype == "Property": prop_rsv += rv_val
        elif ctype == "Liability": liab_rsv += rv_val

    # Compute loss ratios for summary using POLICY TABLE totals (not just filtered claims)
    prop_lr_list = [p for p in lr_data if p["type"] == "Property"]
    liab_lr_list = [p for p in lr_data if p["type"] == "Liability"]
    prop_prem = sum(d["base_premium"] for d in prop_lr_list)
    liab_prem = sum(d["base_premium"] for d in liab_lr_list)
    # Use policy table incurred/claims totals for accurate summary (covers all claims, not just filtered)
    prop_inc_lr = sum(d["incurred"] for d in prop_lr_list)
    liab_inc_lr = sum(d["incurred"] for d in liab_lr_list)
    prop_claims_lr = sum(d["claim_count"] for d in prop_lr_list)
    liab_claims_lr = sum(d["claim_count"] for d in liab_lr_list)
    prop_lr = prop_inc_lr / prop_prem if prop_prem > 0 else 0
    liab_lr = liab_inc_lr / liab_prem if liab_prem > 0 else 0

    # Draw the box (compact to keep loss ratio in white space)
    box_h = 35
    pdf.set_fill_color(*HUB_LIGHT_BLUE)
    pdf.set_draw_color(*HUB_ELECTRIC_BLUE)
    pdf.rect(10, pdf.get_y(), 190, box_h, style="DF")

    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(*HUB_DARK_BLUE)
    pdf.cell(0, 8, "  Executive Summary", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*HUB_DARK)

    col_w = 63
    pdf.cell(col_w, 5, f"  Total Claims: {len(results)}", ln=False)
    pdf.cell(col_w, 5, f"Open: {open_count}", ln=False)
    pdf.cell(col_w, 5, f"Closed: {closed_count}", ln=True)

    pdf.cell(col_w, 5, f"  Total Incurred: ${total_incurred:,.0f}", ln=False)
    pdf.cell(col_w, 5, f"Total Paid: ${total_paid:,.0f}", ln=False)
    pdf.cell(col_w, 5, f"Total Reserved: ${total_reserved:,.0f}", ln=True)

    pdf.cell(col_w, 5, f"  Attorney Rep: {attorney_count} claim(s)", ln=True)

    # Property sub-section - use policy table totals for incurred/claims/loss ratio
    pdf.set_font("Helvetica", "B", 8)
    lr_str_p = f"{prop_lr:.0%}" if prop_prem > 0 else "N/A"
    prop_display_claims = prop_claims_lr if prop_claims_lr > 0 else len(property_claims)
    prop_display_inc = prop_inc_lr if prop_inc_lr > 0 else prop_inc
    pdf.cell(col_w, 4, f"  PROPERTY: {prop_display_claims} claims (O:{prop_open} C:{prop_closed})", ln=False)
    pdf.cell(col_w, 4, f"Incurred: ${prop_display_inc:,.0f}", ln=False)
    pdf.cell(col_w, 4, f"Loss Ratio: {lr_str_p}", ln=True)

    # Liability sub-section - use policy table totals for incurred/claims/loss ratio
    lr_str_l = f"{liab_lr:.0%}" if liab_prem > 0 else "N/A"
    liab_display_claims = liab_claims_lr if liab_claims_lr > 0 else len(liability_claims)
    liab_display_inc = liab_inc_lr if liab_inc_lr > 0 else liab_inc
    pdf.cell(col_w, 4, f"  LIABILITY: {liab_display_claims} claims (O:{liab_open} C:{liab_closed})", ln=False)
    pdf.cell(col_w, 4, f"Incurred: ${liab_display_inc:,.0f}", ln=False)
    pdf.cell(col_w, 4, f"Loss Ratio: {lr_str_l}", ln=True)

    pdf.ln(4)

    # ── Loss Ratio Tables (split) ──
    def render_lr_table(lr_list, title):
        if not lr_list:
            return
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(*HUB_DARK_BLUE)
        pdf.cell(0, 8, title, ln=True)
        pdf.ln(2)

        # Header
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(*HUB_DARK_BLUE)
        pdf.set_text_color(*HUB_WHITE)
        pdf.cell(16, 7, "Year", fill=True, border=1)
        pdf.cell(30, 7, "Policy #", fill=True, border=1)
        pdf.cell(28, 7, "Base Premium", fill=True, border=1, align="R")
        pdf.cell(28, 7, "Incurred", fill=True, border=1, align="R")
        pdf.cell(22, 7, "Loss Ratio", fill=True, border=1, align="R")
        pdf.cell(14, 7, "Claims", fill=True, border=1, align="C")
        pdf.cell(52, 7, "Carrier", fill=True, border=1)
        pdf.ln()

        # Group by year
        from collections import OrderedDict
        year_groups = OrderedDict()
        for p in sorted(lr_list, key=lambda x: (x["year"], x["carrier"])):
            yr = p["year"]
            if yr not in year_groups:
                year_groups[yr] = []
            year_groups[yr].append(p)

        pdf.set_font("Helvetica", "", 7)
        t_prem = t_inc = t_claims = 0
        row_idx = 0

        for yr, policies_in_year in year_groups.items():
            yr_prem = sum(p["base_premium"] for p in policies_in_year)
            yr_inc = sum(p["incurred"] for p in policies_in_year)
            yr_claims = sum(p["claim_count"] for p in policies_in_year)

            if len(policies_in_year) == 1:
                # Single policy for this year - show on one row
                p = policies_in_year[0]
                bg = HUB_LIGHT_BLUE if row_idx % 2 == 0 else HUB_WHITE
                if p["loss_ratio"] > 1.0:
                    bg = HUB_RED_LIGHT
                pdf.set_fill_color(*bg)
                pdf.set_text_color(*HUB_DARK)
                lr_s = f"{p['loss_ratio']:.0%}" if p["base_premium"] > 0 else "N/A"
                carrier = sanitize_for_pdf(p["carrier"][:30])
                pol_num = sanitize_for_pdf(p["policy_num"][:18]) if p["policy_num"] != "N/A" else ""
                pdf.cell(16, 5.5, yr, fill=True, border=1)
                pdf.cell(30, 5.5, pol_num, fill=True, border=1)
                pdf.cell(28, 5.5, f"${p['base_premium']:,.0f}", fill=True, border=1, align="R")
                pdf.cell(28, 5.5, f"${p['incurred']:,.0f}", fill=True, border=1, align="R")
                pdf.cell(22, 5.5, lr_s, fill=True, border=1, align="R")
                pdf.cell(14, 5.5, str(p["claim_count"]), fill=True, border=1, align="C")
                pdf.cell(52, 5.5, carrier, fill=True, border=1)
                pdf.ln()
                row_idx += 1
            else:
                # Multiple policies - show year header then each policy, then subtotal
                # Year header row
                pdf.set_font("Helvetica", "B", 7)
                pdf.set_fill_color(220, 230, 241)  # Light steel blue for year header
                pdf.set_text_color(*HUB_DARK_BLUE)
                pdf.cell(16, 5.5, yr, fill=True, border=1)
                pdf.cell(174, 5.5, "", fill=True, border=1)
                pdf.ln()
                pdf.set_font("Helvetica", "", 7)

                for p in policies_in_year:
                    bg = HUB_LIGHT_BLUE if row_idx % 2 == 0 else HUB_WHITE
                    if p["base_premium"] > 0 and p["loss_ratio"] > 1.0:
                        bg = HUB_RED_LIGHT
                    pdf.set_fill_color(*bg)
                    pdf.set_text_color(*HUB_DARK)
                    lr_s = f"{p['loss_ratio']:.0%}" if p["base_premium"] > 0 else "N/A"
                    carrier = sanitize_for_pdf(p["carrier"][:30])
                    pol_num = sanitize_for_pdf(p["policy_num"][:18]) if p["policy_num"] != "N/A" else ""
                    pdf.cell(16, 5.5, "", fill=True, border=1)  # year blank for sub-rows
                    pdf.cell(30, 5.5, pol_num, fill=True, border=1)
                    pdf.cell(28, 5.5, f"${p['base_premium']:,.0f}", fill=True, border=1, align="R")
                    pdf.cell(28, 5.5, f"${p['incurred']:,.0f}", fill=True, border=1, align="R")
                    pdf.cell(22, 5.5, lr_s, fill=True, border=1, align="R")
                    pdf.cell(14, 5.5, str(p["claim_count"]), fill=True, border=1, align="C")
                    pdf.cell(52, 5.5, carrier, fill=True, border=1)
                    pdf.ln()
                    row_idx += 1

                # Year subtotal row
                pdf.set_font("Helvetica", "B", 7)
                yr_lr = yr_inc / yr_prem if yr_prem > 0 else 0
                sub_bg = HUB_RED_LIGHT if yr_lr > 1.0 else (200, 215, 235)
                pdf.set_fill_color(*sub_bg)
                pdf.set_text_color(*HUB_DARK_BLUE)
                pdf.cell(16, 5.5, "", fill=True, border=1)
                pdf.cell(30, 5.5, f"{yr} Subtotal", fill=True, border=1)
                pdf.cell(28, 5.5, f"${yr_prem:,.0f}", fill=True, border=1, align="R")
                pdf.cell(28, 5.5, f"${yr_inc:,.0f}", fill=True, border=1, align="R")
                pdf.cell(22, 5.5, f"{yr_lr:.0%}", fill=True, border=1, align="R")
                pdf.cell(14, 5.5, str(yr_claims), fill=True, border=1, align="C")
                pdf.cell(52, 5.5, "", fill=True, border=1)
                pdf.ln()
                pdf.set_font("Helvetica", "", 7)
                row_idx += 1

            t_prem += yr_prem
            t_inc += yr_inc
            t_claims += yr_claims

            # Page break check
            if pdf.get_y() > 260:
                pdf.add_page()
                # Redraw header
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_fill_color(*HUB_DARK_BLUE)
                pdf.set_text_color(*HUB_WHITE)
                pdf.cell(16, 7, "Year", fill=True, border=1)
                pdf.cell(30, 7, "Policy #", fill=True, border=1)
                pdf.cell(28, 7, "Base Premium", fill=True, border=1, align="R")
                pdf.cell(28, 7, "Incurred", fill=True, border=1, align="R")
                pdf.cell(22, 7, "Loss Ratio", fill=True, border=1, align="R")
                pdf.cell(14, 7, "Claims", fill=True, border=1, align="C")
                pdf.cell(52, 7, "Carrier", fill=True, border=1)
                pdf.ln()
                pdf.set_font("Helvetica", "", 7)

        # Grand Total row
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(*HUB_DARK_BLUE)
        pdf.set_text_color(*HUB_WHITE)
        olr = t_inc / t_prem if t_prem > 0 else 0
        pdf.cell(16, 6, "TOTAL", fill=True, border=1)
        pdf.cell(30, 6, "", fill=True, border=1)
        pdf.cell(28, 6, f"${t_prem:,.0f}", fill=True, border=1, align="R")
        pdf.cell(28, 6, f"${t_inc:,.0f}", fill=True, border=1, align="R")
        pdf.cell(22, 6, f"{olr:.0%}", fill=True, border=1, align="R")
        pdf.cell(14, 6, str(t_claims), fill=True, border=1, align="C")
        pdf.cell(52, 6, "", fill=True, border=1)
        pdf.ln(6)

    pdf.ln(4)  # Extra spacing to separate from executive summary box
    render_lr_table(prop_lr_list, "Loss Ratio Analysis - Property")
    render_lr_table(liab_lr_list, "Loss Ratio Analysis - Liability")

    # ── Charts ──
    def add_chart(chart_path, title):
        if chart_path and os.path.exists(chart_path):
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 14)
            pdf.set_text_color(*HUB_DARK_BLUE)
            pdf.cell(0, 10, title, ln=True)
            pdf.ln(2)
            chart_w = pdf.w - pdf.l_margin - pdf.r_margin
            pdf.image(chart_path, x=pdf.l_margin, w=chart_w)

    add_chart(chart_incurred, "Claim Trending by Policy Year")
    add_chart(chart_location, "Location Impact Analysis")
    add_chart(chart_cause, "Cause of Loss Analysis")
    # Removed duplicate trending chart (page 5) since incurred chart now serves same purpose
    add_chart(chart_development, "Claims Development - Last 15 Months")

    # Development detail table after chart - centered on page
    if dev_table_data:
        pdf.ln(4)
        # Column widths centered on portrait page (210mm)
        dc = [20, 22, 26, 26, 34, 22, 22, 18]  # total = 190 - wider cause/hazard column
        table_w = sum(dc)
        margin_x = (pdf.w - table_w) / 2
        pdf.set_x(margin_x)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_fill_color(*HUB_DARK_BLUE)
        pdf.set_text_color(*HUB_WHITE)
        headers = ["DOL", "Claim #", "Claimant", "Property", "Cause / Hazard", "Total Inc.", "15mo Dev", "Dev %"]
        for i, h in enumerate(headers):
            align = "R" if i >= 5 else "L"
            pdf.cell(dc[i], 6, h, fill=True, border=1, align=align)
        pdf.ln()

        pdf.set_font("Helvetica", "", 7)
        for idx, d in enumerate(dev_table_data):
            bg = HUB_RED_LIGHT if d["delta"] > 0 else (220, 255, 220)
            pdf.set_fill_color(*bg)
            pdf.set_text_color(*HUB_DARK)
            dev_pct = f"{d['delta']/d['current']*100:.0f}%" if d["current"] > 0 else "N/A"
            pdf.set_x(margin_x)
            pdf.cell(dc[0], 5, d["dol"][:10], fill=True, border=1)
            pdf.cell(dc[1], 5, d["label"][:14], fill=True, border=1)
            pdf.cell(dc[2], 5, d["claimant"][:20], fill=True, border=1)
            pdf.cell(dc[3], 5, d["prop"][:20], fill=True, border=1)
            pdf.cell(dc[4], 5, d["cause"][:28], fill=True, border=1)
            pdf.cell(dc[5], 5, f"${d['current']:,.0f}", fill=True, border=1, align="R")
            pdf.cell(dc[6], 5, f"${d['delta']:+,.0f}", fill=True, border=1, align="R")
            pdf.cell(dc[7], 5, dev_pct, fill=True, border=1, align="R")
            pdf.ln()

    # ── Claims Summary Tables (split, with extra columns for liability) ──
    def render_claims_table(claims_list, section_title, is_liability=False):
        if not claims_list:
            return
        pdf.add_page("L")  # Always landscape for summary tables
        pdf.set_font("Helvetica", "B", 14)
        pdf.set_text_color(*HUB_DARK_BLUE)
        pdf.cell(0, 10, section_title, ln=True)
        pdf.ln(2)

        # Column widths - wider to fill landscape page (297mm - 20mm margins = 277mm)
        if is_liability:
            # DOL, Yr, Claim#, Status, Claimant, Corp-DBA, Cause, Incurred, Paid, Rsv, 15mo Dev
            cw = [20, 12, 26, 14, 32, 42, 34, 24, 24, 20, 29]  # total = 277
        else:
            # DOL, Yr, Claim#, Status, Corp-DBA, Cause, Incurred, Paid, Rsv
            cw = [24, 16, 34, 18, 55, 45, 30, 30, 25]  # total = 277

        def draw_header():
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_fill_color(*HUB_DARK_BLUE)
            pdf.set_text_color(*HUB_WHITE)
            pdf.cell(cw[0], 7, "DOL", fill=True, border=1)
            pdf.cell(cw[1], 7, "Yr", fill=True, border=1)
            pdf.cell(cw[2], 7, "Claim #", fill=True, border=1)
            pdf.cell(cw[3], 7, "Status", fill=True, border=1)
            if is_liability:
                pdf.cell(cw[4], 7, "Claimant", fill=True, border=1)
                pdf.cell(cw[5], 7, "Corporate - DBA", fill=True, border=1)
                pdf.cell(cw[6], 7, "Cause of Loss", fill=True, border=1)
                pdf.cell(cw[7], 7, "Incurred", fill=True, border=1, align="R")
                pdf.cell(cw[8], 7, "Paid", fill=True, border=1, align="R")
                pdf.cell(cw[9], 7, "Rsv", fill=True, border=1, align="R")
                pdf.cell(cw[10], 7, "15mo Dev", fill=True, border=1, align="R")
            else:
                pdf.cell(cw[4], 7, "Corporate - DBA", fill=True, border=1)
                pdf.cell(cw[5], 7, "Cause of Loss", fill=True, border=1)
                pdf.cell(cw[6], 7, "Incurred", fill=True, border=1, align="R")
                pdf.cell(cw[7], 7, "Paid", fill=True, border=1, align="R")
                pdf.cell(cw[8], 7, "Rsv", fill=True, border=1, align="R")
            pdf.ln()

        draw_header()
        pdf.set_font("Helvetica", "", 6)

        for idx, r in enumerate(claims_list):
            flds = r["fields"]
            is_open = flds.get("Status") == "Open"
            if is_open:
                bg = HUB_RED_LIGHT
            elif idx % 2 == 0:
                bg = (248, 248, 248)
            else:
                bg = HUB_WHITE
            pdf.set_fill_color(*bg)
            pdf.set_text_color(*HUB_DARK)

            dol = sanitize_for_pdf(get_val(flds, "Incident Date", ""))
            if dol == "N/A" or not dol:
                dol = sanitize_for_pdf(get_val(flds, "DOL", ""))
            year = sanitize_for_pdf(get_val(flds, "Policy Year", ""))
            cnum = sanitize_for_pdf(get_val(flds, "Claim #", ""))
            st = sanitize_for_pdf(get_val(flds, "Status", ""))
            prop = sanitize_for_pdf(get_val(flds, "DBA (from Location)", ""))
            inc = r["incurred"]
            p = get_val(flds, "Paid - Rollup", "0")
            try:
                p_val = float(p)
            except (ValueError, TypeError):
                p_val = 0
            rv = flds.get("Reserved Helper", [])
            if isinstance(rv, list) and rv:
                try:
                    rsv = float(rv[-1])
                except (ValueError, TypeError):
                    rsv = 0
            else:
                try:
                    rsv = float(rv) if rv else 0
                except (ValueError, TypeError):
                    rsv = 0

            # Build Corporate - DBA and Cause of Loss for all types
            corp = sanitize_for_pdf(get_val(flds, "Corporate Name", ""))
            dba = sanitize_for_pdf(get_val(flds, "DBA (from Location)", ""))
            corp_dba = f"{corp[:15]} - {dba[:15]}" if corp != "N/A" else dba[:30]
            cause = sanitize_for_pdf(get_val(flds, "Cause of Loss Rollup Output", ""))
            if cause == "N/A":
                cause = sanitize_for_pdf(get_val(flds, "Cause of Loss (from Cause of Loss)", ""))
            claimant = sanitize_for_pdf(get_val(flds, "Involved Party (From Involved Party)", ""))
            if claimant == "N/A" or not claimant:
                claimant = sanitize_for_pdf(get_val(flds, "Involved Party copy", ""))

            pdf.cell(cw[0], 5, dol[:10], fill=True, border=1)
            pdf.cell(cw[1], 5, year[:4], fill=True, border=1)
            pdf.cell(cw[2], 5, cnum[:16], fill=True, border=1)
            pdf.cell(cw[3], 5, st[:6], fill=True, border=1)
            if is_liability:
                pdf.cell(cw[4], 5, claimant[:20], fill=True, border=1)
                pdf.cell(cw[5], 5, corp_dba[:28], fill=True, border=1)
                pdf.cell(cw[6], 5, cause[:22], fill=True, border=1)
                pdf.cell(cw[7], 5, f"${inc:,.0f}", fill=True, border=1, align="R")
                pdf.cell(cw[8], 5, f"${p_val:,.0f}", fill=True, border=1, align="R")
                pdf.cell(cw[9], 5, f"${rsv:,.0f}", fill=True, border=1, align="R")
            else:
                pdf.cell(cw[4], 5, corp_dba[:36], fill=True, border=1)
                pdf.cell(cw[5], 5, cause[:30], fill=True, border=1)
                pdf.cell(cw[6], 5, f"${inc:,.0f}", fill=True, border=1, align="R")
                pdf.cell(cw[7], 5, f"${p_val:,.0f}", fill=True, border=1, align="R")
                pdf.cell(cw[8], 5, f"${rsv:,.0f}", fill=True, border=1, align="R")

            if is_liability:
                rid = r.get("record_id", id(r))
                delta = dev_lookup.get(rid, 0)
                # Color the 15mo dev cell: light red for increase, light green for decrease
                if delta > 0:
                    pdf.set_fill_color(255, 204, 204)  # light red
                elif delta < 0:
                    pdf.set_fill_color(204, 255, 204)  # light green
                # else keep current bg
                pdf.cell(cw[10], 5, f"${delta:+,.0f}" if delta != 0 else "-", fill=True, border=1, align="R")
                # Reset fill color for next row
                pdf.set_fill_color(*bg)

            pdf.ln()

            max_y = 180  # Landscape page height limit
            if pdf.get_y() > max_y:
                pdf.add_page("L")  # Always landscape for summary tables
                draw_header()
                pdf.set_font("Helvetica", "", 6)

    render_claims_table(liability_claims, "Liability Claims Summary", is_liability=True)
    render_claims_table(property_claims, "Property Claims Summary", is_liability=False)
    if other_claims:
        render_claims_table(other_claims, "Other Claims Summary", is_liability=False)

    # ── Detailed Claims Analysis (split) ──
    def render_detailed_claims(claims_list, section_title):
        if not claims_list:
            return
        pdf.add_page()  # Portrait page
        # Reset margins explicitly after landscape pages
        pdf.set_left_margin(10)
        pdf.set_right_margin(10)
        pdf.set_x(10)
        pdf.set_font("Helvetica", "B", 14)
        pdf.set_text_color(*HUB_DARK_BLUE)
        pdf.cell(0, 10, section_title, ln=True)
        pdf.ln(2)

        for idx, r in enumerate(claims_list):
            flds = r["fields"]
            if pdf.get_y() > 210:
                pdf.add_page()
                pdf.set_left_margin(10)
                pdf.set_right_margin(10)
                pdf.set_x(10)

            dol = sanitize_for_pdf(get_val(flds, "Incident Date", "N/A"))
            if dol == "N/A":
                dol = sanitize_for_pdf(get_val(flds, "DOL", "N/A"))
            cnum = sanitize_for_pdf(get_val(flds, "Claim #", "N/A"))
            st = sanitize_for_pdf(get_val(flds, "Status", "N/A"))
            ct = sanitize_for_pdf(get_val(flds, "Claim Type", "N/A"))
            year = sanitize_for_pdf(get_val(flds, "Policy Year", "N/A"))
            prop = sanitize_for_pdf(get_val(flds, "DBA (from Location)", "N/A"))
            corp = sanitize_for_pdf(get_val(flds, "Corporate Name", "N/A"))
            claimant = sanitize_for_pdf(get_val(flds, "Involved Party (From Involved Party)", "N/A"))
            if claimant == "N/A":
                claimant = sanitize_for_pdf(get_val(flds, "Involved Party copy", "N/A"))
            col_text = sanitize_for_pdf(get_val(flds, "Cause of Loss Rollup Output", "N/A"))
            if col_text == "N/A":
                col_text = sanitize_for_pdf(get_val(flds, "Cause of Loss (from Cause of Loss)", "N/A"))
            hazard = sanitize_for_pdf(get_val(flds, "Risk/Hazard (From Risk/Hazard)", "N/A"))
            loc_inc = sanitize_for_pdf(get_val(flds, "Location of Incident", "N/A"))
            brief = sanitize_for_pdf(get_val(flds, "Brief Description", "N/A"))
            atty = flds.get("Attorney Representation", False)
            carrier = sanitize_for_pdf(get_val(flds, "Carrier", "N/A"))
            if carrier == "N/A":
                carrier = sanitize_for_pdf(get_val(flds, "Carrier (from Policies)", "N/A"))

            # Header bar - RED for open, dark blue for closed
            is_open = st == "Open"
            if is_open:
                pdf.set_fill_color(200, 40, 40)
            else:
                pdf.set_fill_color(*HUB_DARK_BLUE)
            pdf.set_text_color(*HUB_WHITE)
            pdf.set_font("Helvetica", "B", 9)
            status_icon = "** OPEN **" if is_open else "CLOSED"
            pdf.cell(0, 7, f"  {dol}  |  {ct}  |  {cnum}  |  {status_icon}  |  ${r['incurred']:,.0f}  |  PY: {year}", fill=True, ln=True)

            # Open claim background highlight
            if is_open:
                pdf.set_fill_color(*HUB_RED_LIGHT)
            else:
                pdf.set_fill_color(*HUB_WHITE)

            pdf.set_text_color(*HUB_DARK)
            pdf.set_font("Helvetica", "", 8)

            pdf.cell(0, 5, f"Property: {prop}", fill=is_open, ln=True)
            pdf.cell(0, 5, f"Corporate: {corp}", fill=is_open, ln=True)
            pdf.cell(95, 5, f"Claimant: {claimant}", fill=is_open, ln=False)
            pdf.cell(0, 5, f"Cause: {col_text}", fill=is_open, ln=True)

            detail_parts = []
            if hazard != "N/A":
                detail_parts.append(f"Hazard: {hazard}")
            if loc_inc != "N/A":
                detail_parts.append(f"Location: {loc_inc}")
            if atty:
                detail_parts.append("Attorney: Yes")
            if carrier != "N/A":
                detail_parts.append(f"Carrier: {carrier}")
            if detail_parts:
                pdf.cell(0, 5, "  |  ".join(detail_parts), fill=is_open, ln=True)

            if brief != "N/A":
                pdf.set_font("Helvetica", "I", 7)
                pdf.multi_cell(0, 4, sanitize_for_pdf(f"Description: {brief[:250]}"))

            # Claims Development
            raw_activity = flds.get("Activity Rollup Raw Data", "")
            if raw_activity:
                valuations = parse_claims_development(raw_activity)
                if valuations:
                    pdf.set_x(pdf.l_margin)  # Ensure left-aligned
                    pdf.set_font("Helvetica", "B", 7)
                    pdf.cell(0, 5, "Claims Development:", ln=True)
                    pdf.set_font("Helvetica", "", 7)
                    for v in valuations:
                        parts = []
                        if v["paid"] > 0:
                            parts.append(f"Paid: ${v['paid']:,.0f}")
                        if v["reserved"] > 0:
                            parts.append(f"Rsv: ${v['reserved']:,.0f}")
                        if v["expenses"] > 0:
                            parts.append(f"Exp: ${v['expenses']:,.0f}")
                        detail = f" ({', '.join(parts)})" if parts else ""
                        pdf.cell(0, 4, f"    {v['date']}: ${v['total_incurred']:,.0f}{detail}", ln=True)

                comments = parse_activity_comments(raw_activity)
                if comments:
                    pdf.set_x(pdf.l_margin)  # Ensure left-aligned
                    pdf.set_font("Helvetica", "B", 7)
                    pdf.cell(0, 5, "Activity Comments:", ln=True)
                    pdf.set_font("Helvetica", "", 7)
                    for c in comments[:5]:
                        comment_text = sanitize_for_pdf(c["text"][:300])
                        if len(c["text"]) > 300:
                            comment_text += "..."
                        pdf.set_font("Helvetica", "I", 6)
                        pdf.cell(0, 4, f"  [{c['date']}]", ln=True)
                        pdf.set_font("Helvetica", "", 7)
                        pdf.multi_cell(0, 3.5, f"    {comment_text}")

            pdf.ln(4)

    render_detailed_claims(liability_claims, "Detailed Liability Claims Analysis")
    render_detailed_claims(property_claims, "Detailed Property Claims Analysis")
    if other_claims:
        render_detailed_claims(other_claims, "Detailed Other Claims Analysis")

    # Save
    filepath = os.path.join(tmpdir, f"Claims_Report_{client_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf")
    pdf.output(filepath)
    return filepath
