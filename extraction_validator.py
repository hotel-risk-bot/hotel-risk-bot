"""
Post-extraction validation framework.
Runs automated checks on extracted data before proposal generation.
Catches issues early and auto-corrects where possible.
"""

import logging
import re

logger = logging.getLogger(__name__)


def validate_extraction(data, sov_data=None):
    """
    Run all validation checks on extracted data.
    Returns a dict with:
      - warnings: list of warning messages (issues found but not auto-corrected)
      - corrections: list of corrections applied automatically
      - errors: list of critical errors that may need user attention
    """
    result = {
        "warnings": [],
        "corrections": [],
        "errors": [],
    }

    _validate_premiums(data, result)
    _validate_locations(data, sov_data, result)
    _validate_named_insureds(data, sov_data, result)
    _validate_tiv_data(data, sov_data, result)
    _validate_coverages(data, result)

    # Log summary
    if result["errors"]:
        logger.error(f"Validation ERRORS ({len(result['errors'])}): {result['errors']}")
    if result["warnings"]:
        logger.warning(f"Validation WARNINGS ({len(result['warnings'])}): {result['warnings']}")
    if result["corrections"]:
        logger.info(f"Validation CORRECTIONS ({len(result['corrections'])}): {result['corrections']}")

    return result


def _parse_currency(val):
    """Parse a currency string or number into a float."""
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        cleaned = re.sub(r'[^\d.]', '', val.replace(',', ''))
        if cleaned:
            try:
                return float(cleaned)
            except ValueError:
                pass
    return 0.0


def _validate_premiums(data, result):
    """Check premium consistency for each coverage."""
    coverages = data.get("coverages", {})

    for cov_key, cov in coverages.items():
        if not isinstance(cov, dict):
            continue

        carrier = cov.get("carrier", "")
        premium = _parse_currency(cov.get("premium", 0))
        total_premium = _parse_currency(cov.get("total_premium", 0))
        taxes_fees = _parse_currency(cov.get("taxes_fees", 0))

        if not carrier or carrier.upper() == "TBD":
            continue  # Skip placeholder coverages

        # Check 1: total_premium should be >= premium
        if total_premium > 0 and premium > 0 and total_premium < premium:
            result["warnings"].append(
                f"{cov_key}: total_premium (${total_premium:,.2f}) < premium (${premium:,.2f}). "
                f"Total should include taxes/fees."
            )
            # Auto-correct: swap if total < premium
            cov["total_premium"] = premium
            cov["premium"] = total_premium
            result["corrections"].append(
                f"{cov_key}: Swapped premium and total_premium (likely reversed)"
            )

        # Check 2: total_premium should approximately equal premium + taxes_fees
        if total_premium > 0 and premium > 0 and taxes_fees > 0:
            expected_total = premium + taxes_fees
            diff = abs(total_premium - expected_total)
            if diff > 100:  # Allow $100 tolerance for rounding
                result["warnings"].append(
                    f"{cov_key}: total_premium (${total_premium:,.2f}) != "
                    f"premium (${premium:,.2f}) + taxes_fees (${taxes_fees:,.2f}) = "
                    f"${expected_total:,.2f}. Difference: ${diff:,.2f}"
                )

        # Check 3: If total_premium is 0 but premium exists, use premium as total
        if premium > 0 and total_premium == 0:
            cov["total_premium"] = premium
            result["corrections"].append(
                f"{cov_key}: Set total_premium to premium (${premium:,.2f}) since total was 0"
            )

        # Check 4: Premium should not be unreasonably small or large
        if premium > 0 and premium < 100:
            result["warnings"].append(
                f"{cov_key}: Premium ${premium:,.2f} seems unusually low"
            )


def _normalize_addr_simple(addr):
    """Simple address normalization for comparison."""
    if not addr:
        return ""
    addr = addr.upper().strip()
    addr = re.sub(r'\s+', ' ', addr)
    addr = re.sub(r'\bSTREET\b', 'ST', addr)
    addr = re.sub(r'\bROAD\b', 'RD', addr)
    addr = re.sub(r'\bDRIVE\b', 'DR', addr)
    addr = re.sub(r'\bAVENUE\b', 'AVE', addr)
    addr = re.sub(r'\bBOULEVARD\b', 'BLVD', addr)
    addr = re.sub(r'\bHIGHWAY\b', 'HWY', addr)
    addr = re.sub(r'\bPIKE\b', 'PK', addr)
    addr = re.sub(r'\bPLACE\b', 'PL', addr)
    return addr


def _validate_locations(data, sov_data, result):
    """Check location count consistency between SOV, GL, and locations array."""
    locations = data.get("locations", [])
    sov_locs = sov_data.get("locations", []) if sov_data else []
    coverages = data.get("coverages", {})
    gl_cov = coverages.get("general_liability", {})
    gl_classes = gl_cov.get("schedule_of_classes", []) if isinstance(gl_cov, dict) else []

    # Count unique GL locations (skip non-physical entries)
    _skip = {"hired auto", "non-owned auto", "loss control", "package store",
             "category vi", "liquor", "sundry", "flat"}
    gl_addrs = set()
    for entry in gl_classes:
        if isinstance(entry, dict):
            classification = (entry.get("classification", "") or "").lower()
            if any(skip in classification for skip in _skip):
                continue
            addr = entry.get("address", "")
            if addr and addr.strip():
                gl_addrs.add(_normalize_addr_simple(addr.split(",")[0]))

    sov_count = len(sov_locs)
    gl_count = len(gl_addrs)
    loc_count = len(locations)

    # Check: SOV and GL should have similar location counts
    if sov_count > 0 and gl_count > 0:
        if abs(sov_count - gl_count) > 2:
            result["warnings"].append(
                f"Location count mismatch: SOV has {sov_count} locations, "
                f"GL schedule has {gl_count} unique locations"
            )

    # Check: locations array should match SOV count (when SOV is available)
    if sov_count > 0 and loc_count > 0:
        if loc_count > sov_count * 2:
            result["warnings"].append(
                f"Locations array ({loc_count}) has significantly more entries than SOV ({sov_count}). "
                f"May include non-physical addresses (mailing, carrier office)."
            )

    # Check: If SOV exists, ensure locations array uses SOV data
    if sov_count > 0 and loc_count > 0:
        locs_with_tiv = sum(1 for loc in locations if (loc.get("tiv", 0) or 0) > 0)
        if locs_with_tiv == 0:
            result["warnings"].append(
                f"No locations have TIV values despite SOV having {sov_count} locations with TIVs. "
                f"SOV data may not be flowing to locations array."
            )


def _validate_named_insureds(data, sov_data, result):
    """Check named insured completeness."""
    named_insureds = data.get("named_insureds", [])
    additional = data.get("additional_named_insureds", [])

    if not named_insureds and not additional:
        result["warnings"].append("No named insureds found in extracted data")
        return

    # Check: If GL has additional_named_insureds, they should be in named_insureds
    coverages = data.get("coverages", {})
    gl_cov = coverages.get("general_liability", {})
    if isinstance(gl_cov, dict):
        gl_additional = gl_cov.get("additional_named_insureds", [])
        if gl_additional and len(gl_additional) > len(named_insureds):
            result["warnings"].append(
                f"GL quote has {len(gl_additional)} additional named insureds but "
                f"only {len(named_insureds)} are in the named_insureds list"
            )


def _validate_tiv_data(data, sov_data, result):
    """Check TIV data presence and consistency."""
    if not sov_data or not sov_data.get("locations"):
        return

    sov_locs = sov_data["locations"]
    totals = sov_data.get("totals", {})

    # Check: SOV locations should have building values
    locs_with_bldg = sum(1 for loc in sov_locs if (loc.get("building_value", 0) or 0) > 0)
    if locs_with_bldg == 0:
        result["warnings"].append("SOV locations have no building values")

    # Check: SOV totals should be populated
    total_tiv = totals.get("tiv", 0) or 0
    if total_tiv == 0:
        # Try to calculate from individual locations
        calc_tiv = sum(loc.get("tiv", 0) or 0 for loc in sov_locs)
        if calc_tiv > 0:
            totals["tiv"] = calc_tiv
            result["corrections"].append(
                f"Calculated SOV total TIV from locations: ${calc_tiv:,.0f}"
            )
        else:
            result["warnings"].append("SOV has no TIV data (neither totals nor individual locations)")

    # Check: sov_data should be in the main data dict
    if not data.get("sov_data"):
        data["sov_data"] = sov_data
        result["corrections"].append("Injected sov_data into main data dict for generator access")


def _validate_coverages(data, result):
    """Check coverage data completeness."""
    coverages = data.get("coverages", {})

    for cov_key, cov in coverages.items():
        if not isinstance(cov, dict):
            continue

        carrier = (cov.get("carrier", "") or "").strip()
        if not carrier or carrier.upper() == "TBD":
            # Optional coverage placeholder — skip validation
            continue

        # Check: Coverage should have forms/endorsements
        forms = cov.get("forms_endorsements", [])
        if not forms:
            result["warnings"].append(
                f"{cov_key}: No forms/endorsements extracted for {carrier}"
            )

        # Check: Coverage should have limits
        limits = cov.get("coverage_limits", []) or cov.get("limits", [])
        if not limits:
            result["warnings"].append(
                f"{cov_key}: No coverage limits extracted for {carrier}"
            )

        # Check: GL should have schedule_of_classes
        if cov_key == "general_liability":
            classes = cov.get("schedule_of_classes", [])
            if not classes:
                result["warnings"].append(
                    "general_liability: No schedule_of_classes extracted. "
                    "Liability locations may not be identified."
                )
            # Check: GL should have designated_premises or schedule_of_locations
            premises = cov.get("designated_premises", [])
            schedule_locs = cov.get("schedule_of_locations", [])
            if not premises and not schedule_locs and not classes:
                result["warnings"].append(
                    "general_liability: No designated premises or schedule of locations found. "
                    "Liability checkmarks may be incomplete."
                )
