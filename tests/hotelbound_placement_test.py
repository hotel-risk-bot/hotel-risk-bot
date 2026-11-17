"""Regression checks for the HotelBound pass when a submission carries more than one property policy.

Run from the repo root:  python3 tests/hotelbound_placement_test.py
Self-contained: synthetic quote text in the layout pdftotext -layout produces, no OpenAI calls.

Why (Nov 16 2026, WS Management Inc): the HotelBound master quote and a separate property policy
for one hotel (BlueSky / Canopius) were uploaded together. The model read both correctly but put
the separate policy in "property" and HotelBound in "property_alt_1"; the HotelBound pass then wrote
the HotelBound figures over "property". HotelBound showed twice ($407,562.03 x 2), the separate
policy disappeared, and its carrier participation, UMR and quote number stayed inside the
HotelBound section. The same quote also wraps each full address inside the Cost Break Out Street
column, so hotels came out as "11855 Avenue of" and lost their Schedule of Locations checkmark.

  * hotelbound_quote.place_hotelbound_quote - the HotelBound quote's own entry becomes "property",
    any other property policy keeps a property_alt slot, a second reading of HotelBound is dropped
  * hotelbound_quote.apply_hotelbound_to_data - the separate policy's hotel and payment row survive
  * hotelbound_quote._parse_schedule - wrapped Street / City cells in the Cost Break Out
  * proposal_generator - "Property (Blue Harbor)" instead of "Property (Option 2)", the separate
    hotel counted and checked as property, the TIV covers both policies
  * client_email - the second property policy is named by its carrier
"""
import copy
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import hotelbound_quote as hq  # noqa: E402
import proposal_extractor as pe  # noqa: E402
from premium_reconcile import run_premium_pass  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


HB_FILE = "Harbor Light Hotels - HotelBound Quote.pdf"
OTHER_FILE = "1350 Harbor Light - Blue Harbor Primary Quote.pdf"

HB_TEXT = """
--- Page 1 ---
                                                                                                          RT Specialty
                                           Insurance Proposal
                 Applicant: Harbor Light Hotels LLC
              Policy Period: 11/15/2026 12:01 AM To 11/15/2027 12:01 AM
                  Coverage: Property
          Issuing Company: Certain Underwriters at Lloyd's, London

--- Page 3 ---
Cost Summary
     Property Premium                                                  $120,000.00
     CA Surplus Lines Tax                                                $3,600.00
     CA Stamp Fee                                                          $216.00
     Carrier Policy Fee                                                 $18,300.00
 Total Policy Cost                                                     $142,116.00

--- Page 5 ---
We are pleased to deliver this quotation confirmation for the insured captioned above. This quote is issued
as part of the HotelBound insurance program.

--- Page 7 ---
QUOTATION MEMORANDUM

                                                  RISK DETAILS

              TYPE: First-Party Property Insurance on an "All Risk" basis excluding Flood and Earth Movement.

           PERIOD: From November 15, 2026 to November 15, 2027 both days at 12:01 a.m. Local Time

    LEAD INSURER: Certain Underwriters of Lloyds of London

               TIV: $40,000,000

 LIMIT OF LIABILITY $40,000,000 Overall Program Limit of Liability any one Occurrence Except Perils of Flood
   AND SUBLIMITS: and Earth Movement as noted below.

        TERRORISM: Included - This premium is included in the Total Premium shown above

PREMIUM AND FEES:          Non-Admitted Premium:                                 $120,000
                           Program Brokerage Fee:                                 $18,000
                           Program Inspection Fee:                                     $0
                                 *Association Fee:                                  $300

--- Page 24 ---
                                                                                                     Cost Break Out By Location
ACCOUNT NAME: Harbor Light Hotels LLC                                                                                                         CONTROL #: 10001

      LID             Street                 City        State     Zip        County          Agg       Premium          Fee       Bldg. Value        BPP Value             BI Value                 TIV
                                                                                            Premiu
                                                                                               m
11111111101   500 Ocean View Road,      Del Mar         CA        92014   San Diego        $9,000      $54,000    $8,100       $14,000,000       $1,000,000       $3,000,000           $18,000,000
              Del Mar, California,
              92014, United States


11111111102   8800 Avenue of            Rancho          CA        91730   San Bernardino   $11,000     $66,000    $9,900       $17,000,000       $1,500,000       $3,500,000           $22,000,000
              Commerce, Rancho          Cucamonga
              Cucamonga, California,
              91730, United States

                        SHARED CAPACITY AND DEDICATED LIMITS DISCLOSURE
The Master insurance policies provide a program limit of insurance of $400,000,000 per Occurrence and
are further subject to certain program sublimits.
"""

OTHER_TEXT = """
--- Page 1 ---
                                           Insurance Proposal
                 Applicant: Harbor Light Hotels LLC
          Issuing Company: Blue Harbor Specialty Insurance Company

--- Page 3 ---
Cost Summary
     Property Premium                                                   $50,476.00
     CA Surplus Lines Tax                                                $1,533.03
     CA Stamp Fee                                                           $91.98
     Carrier Inspection Fee                                                $125.00
     Policy Fee                                                            $500.00
     Carrier Policy Fee                                                    $500.00
 Total Policy Cost                                                      $53,226.01

--- Page 6 ---
                                        Primary Coverage Summary - Blue Harbor Specialty Insurance Company
Total Insurable Value (TIV):                                                        $32,251,339.00
Loss Limit:                                                                         $25,000,000.00
                                                 Location Schedule
   1         1       1350 Bayshore Drive                Burlingame     CA      94010       Frame
"""

HB_GPT = {
    "carrier": "Certain Underwriters at Lloyd's, London", "carrier_admitted": False,
    "premium": 120000, "taxes_fees": 22116, "total_premium": 142116, "tiv": "$40,000,000",
    "limits": [{"description": "Overall Program Limit of Liability", "limit": "$40,000,000"}],
    "forms_endorsements": [{"form_number": "HB Evidence of Coverage", "description": "HotelBound Evidence of Coverage"}],
}
OTHER = {
    "carrier": "Blue Harbor Specialty Insurance Company", "carrier_admitted": False,
    "premium": 50476, "taxes_fees": 2750.01, "total_premium": 53226.01, "tiv": "$32,251,339",
    "limits": [{"description": "Loss Limit", "limit": "$25,000,000"},
               {"description": "Building", "limit": "$28,190,871"}],
    "deductibles": [{"description": "All Other Perils", "amount": "$25,000"}],
    "coverage_by_location": [{"premise": 1, "building": 1, "address": "1350 Bayshore Drive, Burlingame, CA 94010",
                              "building_value": "$28,190,871", "bpp_value": "$1,160,468",
                              "business_income": "$2,900,000"}],
    "carrier_participation": [{"carrier": "Blue Harbor Specialty Insurance Company", "percentage_share": "50.00%"},
                              {"carrier": "Certain Underwriters at Lloyd's Syndicate 9999", "percentage_share": "50.00%"}],
    "umr": "UMR-TEST-0001", "quote_number": "BH-100",
    "forms_endorsements": [{"form_number": "CP 00 10 06 07", "description": "Building and Personal Property Coverage Form"}],
}
STARR = {"carrier": "Starr Surplus Lines Insurance Company", "premium": 98000, "total_premium": 104000.55,
         "tiv": "$40,000,000"}
STARR_TEXT = "STARR SURPLUS LINES INSURANCE COMPANY\nTotal Premium $98,000.00\nTotal Cost $104,000.55\n"
GARBLED_HB = {"carrier": "RT Specialty", "premium": 138300, "total_premium": 138300, "tiv": "$40,000,000"}

LOCATIONS = [
    {"number": "1", "name": "Harbor Light Del Mar", "address": "500 Ocean View Road", "city": "Del Mar", "state": "CA",
     "zip": "92014", "tiv": 18000000},
    {"number": "2", "name": "Harbor Light Airport", "address": "1350 Bayshore Drive", "city": "Burlingame",
     "state": "CA", "zip": "94010", "tiv": 32251339},
]

HB_TOTAL, OTHER_TOTAL = 142116.0, 53226.01


def make(prop=None, alt1=None, alt2=None):
    covs = {}
    for k, v in (("property", prop), ("property_alt_1", alt1), ("property_alt_2", alt2)):
        if v is not None:
            covs[k] = copy.deepcopy(v)
    return {
        "client_name": "Harbor Light Hotels LLC",
        "client_info": {"named_insured": "Harbor Light Hotels LLC", "effective_date": "11/15/2026",
                        "expiration_date": "11/15/2027"},
        "named_insureds": [{"name": "Harbor Light Hotels LLC"}],
        "coverages": covs,
        "locations": copy.deepcopy(LOCATIONS),
        "payment_options": [
            {"carrier": "Blue Harbor Specialty Insurance Company", "coverage_type": "Property",
             "terms": "Due within 20 days of inception", "mep": "25%"},
            {"carrier": "Certain Underwriters at Lloyd's, London", "coverage_type": "Property",
             "terms": "Due within 15 days", "mep": "35%"},
        ],
        "warnings": [],
    }


ITEMS = [{"filename": HB_FILE, "text": HB_TEXT}, {"filename": OTHER_FILE, "text": OTHER_TEXT}]
HB_ONLY = [ITEMS[0]]


def run(data, items):
    check(pe._run_hotelbound_pass(data, items, ""), "HotelBound pass applied")
    run_premium_pass(data, items)
    return data["coverages"]


def is_hb(cov):
    return isinstance(cov, dict) and (cov.get("hotelbound") or {}).get("detected") and cov.get("total_premium") == HB_TOTAL


def is_other(cov):
    return (isinstance(cov, dict) and str(cov.get("carrier", "")).startswith("Blue Harbor")
            and cov.get("total_premium") == OTHER_TOTAL and not cov.get("hotelbound"))


# ---------------------------------------------------------------- parser: wrapped Cost Break Out cells
hb = hq.parse_hotelbound_quote(HB_TEXT)
check(hb["detected"] and hb["total_policy_cost"] == HB_TOTAL and hb["base_premium"] == 120000, hb)
streets = [(r["street"], r["city"]) for r in hb["schedule"]]
check(streets == [("500 Ocean View Road", "Del Mar"), ("8800 Avenue of Commerce", "Rancho Cucamonga")], streets)
check(sum(r["tiv"] for r in hb["schedule"]) == 40000000, "schedule TIV")
check(hq._street_from_cell("4055 Street Road,", "Bensalem", "PA") == "4055 Street Road", "plain street, trailing comma")
check(hq._street_from_cell("4055 Street Road", "Bensalem", "PA") == "4055 Street Road", "plain street unchanged")
check(hq._street_from_cell("1041 Harriman Place, 1041 E, San Bernardino, California, 92408, United States",
                           "San Bernadino", "CA") == "1041 Harriman Place 1041 E",
      "full address cut at the state when the City column is spelled differently")

# ---------------------------------------------------------------- the production case: other policy in "property"
d = run(make(OTHER, HB_GPT), ITEMS)
covs = d
check(is_hb(covs["property"]), covs["property"])
check(is_other(covs["property_alt_1"]), covs.get("property_alt_1"))
check("property_alt_2" not in covs, "no third slot")
for k in ("umr", "quote_number", "carrier_participation", "deductibles"):
    if k == "deductibles":
        check(all("All Other Perils" != x.get("description") for x in covs["property"].get(k) or []), "deductibles")
    else:
        check(k not in covs["property"], f"{k} of the other policy is not in the HotelBound section")
check(covs["property_alt_1"]["umr"] == "UMR-TEST-0001" and covs["property_alt_1"]["quote_number"] == "BH-100",
      "the other policy keeps its own details")
full = make(OTHER, HB_GPT)
pe._run_hotelbound_pass(full, ITEMS, "")
check([l["address"] for l in full["locations"]] == ["500 Ocean View Road", "8800 Avenue of Commerce", "1350 Bayshore Drive"],
      full["locations"])
check(full["locations"][-1]["name"] == "Harbor Light Airport" and full["locations"][-1]["tiv"] == 32251339,
      full["locations"][-1])
po = [(p["carrier"], p["coverage_type"]) for p in full["payment_options"]]
check(("Blue Harbor Specialty Insurance Company", "Property") in po and sum("(HotelBound)" in c for c, _ in po) == 1, po)
check(any(w.startswith("Property check") and OTHER_FILE in w for w in full["warnings"]), full["warnings"])

# ---------------------------------------------------------------- HotelBound already in "property"
covs = run(make(HB_GPT, OTHER), ITEMS)
check(is_hb(covs["property"]) and is_other(covs["property_alt_1"]), "in place, alt untouched")
covs = run(make(HB_GPT), HB_ONLY)
check(is_hb(covs["property"]) and list(covs) == ["property"], "single HotelBound quote")

# ---------------------------------------------------------------- the model skipped HotelBound
covs = run(make(OTHER), ITEMS)
check(is_hb(covs["property"]) and is_other(covs["property_alt_1"]), "other policy moved out, HotelBound built")

# ---------------------------------------------------------------- a garbled HotelBound reading stays put
covs = run(make(GARBLED_HB), HB_ONLY)
check(is_hb(covs["property"]) and list(covs) == ["property"], "garbled reading, single file")
covs = run(make(GARBLED_HB, OTHER), ITEMS)
check(is_hb(covs["property"]) and is_other(covs["property_alt_1"]) and "property_alt_2" not in covs, "garbled + other")

# ---------------------------------------------------------------- HotelBound read twice
covs = run(make(HB_GPT, copy.deepcopy(HB_GPT)), ITEMS)
check(list(covs) == ["property"] and is_hb(covs["property"]), "duplicate dropped")

# ---------------------------------------------------------------- three property quotes, HotelBound last
items3 = ITEMS + [{"filename": "Starr Property Quote.pdf", "text": STARR_TEXT * 3}]
covs = run(make(OTHER, STARR, HB_GPT), items3)
check(is_hb(covs["property"]), "HotelBound first")
check({covs["property_alt_1"]["carrier"].split()[0], covs["property_alt_2"]["carrier"].split()[0]} == {"Starr", "Blue"},
      "both other quotes kept")
covs = run(make(STARR), items3)
check(is_hb(covs["property"]) and covs["property_alt_1"]["carrier"].startswith("Starr"),
      "a competing quote with the same TIV is not taken for HotelBound")

# ---------------------------------------------------------------- one combined text (bot path)
covs = run(make(OTHER, HB_GPT), [{"filename": "combined", "text": HB_TEXT + OTHER_TEXT}])
check(is_hb(covs["property"]) and is_other(covs["property_alt_1"]), "combined text: figures decide")

# ---------------------------------------------------------------- idempotent
twice = make(OTHER, HB_GPT)
pe._run_hotelbound_pass(twice, ITEMS, "")
snap = copy.deepcopy(twice)
pe._run_hotelbound_pass(twice, ITEMS, "")
check(twice["coverages"] == snap["coverages"] and twice["warnings"] == snap["warnings"]
      and twice["locations"] == snap["locations"], "second pass changes nothing")

# ---------------------------------------------------------------- proposal and email
import docx  # noqa: E402
import proposal_generator as pg  # noqa: E402
from client_email import build_email_context  # noqa: E402

gen = make(OTHER, HB_GPT)
pe._run_hotelbound_pass(gen, ITEMS, "")
run_premium_pass(gen, ITEMS)
with tempfile.TemporaryDirectory() as td:
    out = os.path.join(td, "p.docx")
    pg.generate_proposal(copy.deepcopy(gen), out)
    d = docx.Document(out)
    paras = [p.text for p in d.paragraphs]
    tables = [[[c.text.strip() for c in r.cells] for r in t.rows] for t in d.tables]
prem = next(t for t in tables if t and t[0][:3] == ["Coverage", "Carrier", "Proposed Premium"])
check(["Property", "Certain Underwriters at Lloyd's, London", "$142,116.00"] in prem, prem)
check(any(r[0] == "Property (Blue Harbor Specialty)" and r[1].startswith("Blue Harbor") and r[2] == "$53,226.01"
          for r in prem), prem)
check(not any("Option 2" in r[0] for r in prem), "no 'Option 2' for a second property policy")
check(prem[-1][-1] == "$195,342.01", prem[-1])
check("Property Coverage — Blue Harbor Specialty" in paras and "Property Coverage — Option 2" not in paras, "section title")
rows = {r[0]: r[1] for t in tables for r in t if len(r) >= 2}
check(rows.get("Total Number of Locations") == "3" and rows.get("Property Locations") == "3", rows)
check(rows.get("Total Insured Value (TIV)") == "$72,251,339", rows.get("Total Insured Value (TIV)"))
sched = next(t for t in tables if t and t[0][:3] == ["#", "Property Name", "Address"])
by_addr = {r[2]: r for r in sched[1:] if r[2]}
check(set(by_addr) == {"500 Ocean View Road", "8800 Avenue of Commerce", "1350 Bayshore Drive"}, by_addr)
check(all(r[6] == "✓" for r in by_addr.values()), "every hotel is on a property policy")
hb_start = paras.index("Property Coverage")
hb_end = paras.index("Property Coverage — Blue Harbor Specialty")
check(not any("Blue Harbor" in p or "UMR-TEST" in p for p in paras[hb_start:hb_end]), "HotelBound section is HotelBound only")

ctx = build_email_context(copy.deepcopy(gen))
labels = [l["coverage"] for l in ctx.get("coverages", [])]
check("Property (Blue Harbor Specialty)" in labels and not any("Option" in l for l in labels), labels)

# GL alt labels are unchanged by the shared helper
check(pg._gl_alt_label({"general_liability_alt_1": {"carrier": "Burlington Insurance Company"}},
                       "general_liability_alt_1", 2) == "General Liability (Burlington)", "GL label")
check(pg._gl_alt_section_title({}, "general_liability_alt_1", 2) == "General Liability Coverage — Option 2", "GL fallback")

print(f"hotelbound_placement_test: {PASS} checks passed")
