"""Regression checks for a HotelBound account whose liability quote lists a different set of hotels.

Run from the repo root:  python3 tests/liability_only_locations_test.py
Self-contained: synthetic quote text in the layout pdftotext -layout produces, no OpenAI calls.

Why (Stefan, Nov 24 2026, WG Hotel LLC): "why didn't this read with 7 locations". The HotelBound
quote scheduled six addresses and the General Liability quote six, five of them the same. One
hotel is on the property schedule only (its liability is written elsewhere) and one is on the
liability quote only (its property is written elsewhere): seven hotels. The proposal showed six.

  * the HotelBound pass replaced the whole location list with its own schedule, so the
    liability-only hotel disappeared from the review screen
  * the property-only hotel got a Liability checkmark from a count rule ("the GL lists as many
    locations as the property schedule, so it covers them all")
  * the Information Summary counted 6 / 6 / 6
  * the Cost Break Out prints a long street and its city with one space between them
    ("815 North Pottstown Pike Exton"), so two hotels kept the wrapped Schedule of Values
    reading ("7630 Boulevard", "Philadelphi")

  * hotelbound_quote._split_street_city / _parse_schedule - joined street and city cells
  * hotelbound_quote.merge_liability_only_locations - the liability-only hotel stays, with no
    insured value, and a review note says what differs
  * proposal_generator - Schedule of Locations checkmarks and Information Summary counts
"""
import copy
import logging
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
os.environ.setdefault("OPENAI_API_KEY", "test")
logging.disable(logging.CRITICAL)

import docx  # noqa: E402
import hotelbound_quote as hq  # noqa: E402
import proposal_generator as pg  # noqa: E402

PASS = 0


def check(cond, msg):
    global PASS
    if not cond:
        raise AssertionError(msg)
    PASS += 1


# ------------------------------------------------------------------ 1. street / city in one cell
split = hq._split_street_city
check(split("615 North Lancaster Pike Exton", ["Exton"]) == ("615 North Lancaster Pike", "Exton"), "city hint")
# the Schedule of Values column cut the city short
check(split("5630 Germantown Boulevard Philadelphia", ["Philadelphi"]) == ("5630 Germantown Boulevard", "Philadelphia"),
      split("5630 Germantown Boulevard Philadelphia", ["Philadelphi"]))
check(split("5426 Lower York Road New Hope", ["New Hope"]) == ("5426 Lower York Road", "New Hope"), "two-word city")
# no hint: the street type (and a direction after it) ends the street
check(split("2100 Oxford Boulevard West Langhorne") == ("2100 Oxford Boulevard West", "Langhorne"),
      split("2100 Oxford Boulevard West Langhorne"))
check(split("615 North Lancaster Pike Exton") == ("615 North Lancaster Pike", "Exton"), "street type, no hint")
# nothing says where the city starts: the cell is left whole rather than guessed
check(split("77 Old Quarry Camden") == ("77 Old Quarry Camden", ""), split("77 Old Quarry Camden"))

# ------------------------------------------------------------------ 2. the quote
HB_TEXT = """
--- Page 1 ---
                                                                                                          RT Specialty
                                           Insurance Proposal
                 Applicant: Keystone Lodging Group LLC
              Policy Period: 12/01/2026 12:01 AM To 12/01/2027 12:01 AM
                  Coverage: Property
          Issuing Company: Certain Underwriters at Lloyd's, London

--- Page 3 ---
Cost Summary
     Property Premium                                                  $113,729.00
     PA Surplus Lines Tax                                                $3,411.87
     PA Stamping Fee                                                        $25.00
     Carrier Policy Fee                                                 $17,360.00
 Total Policy Cost                                                     $134,525.87

--- Page 5 ---
We are pleased to deliver this quotation confirmation for the insured captioned above. This quote is issued
as part of the HotelBound insurance program.

--- Page 7 ---
QUOTATION MEMORANDUM

                                                  RISK DETAILS

              TYPE: First-Party Property Insurance on an "All Risk" basis excluding Flood and Earth Movement.

           PERIOD: From December 1, 2026 to December 1, 2027 both days at 12:01 a.m. Local Time

    LEAD INSURER: Certain Underwriters of Lloyds of London

               TIV: $53,585,302

 LIMIT OF LIABILITY $53,585,302 Overall Program Limit of Liability any one Occurrence Except Perils of Flood
   AND SUBLIMITS: and Earth Movement as noted below.

        TERRORISM: Included - This premium is included in the Total Premium shown above

     SCHEDULE OF         A. $500,000 per Occurrence per Insured Member as respects Extra Expense;
       SUBLIMITS:
                         B. Lesser of $100,000 or 30 days per Occurrence per Insured Member as respects
                            Ingress / Egress; limit 5 mile radius;
              RT Specialty is a division of RSG Specialty, LLC, a Delaware limited liability company
                                  and a subsidiary of Ryan Specialty Group, LLC.
                   In California: RSG Specialty Insurance Services, LLC (License # 0G97516)
                              150 South US Highway 1, 3rd Floor, Jupiter, FL 33477
                                    Phone - (561) 615-5552 Fax - (561) 615-5592                     Page | 6
UQT 10 2022

--- Page 8 ---
QUOTATION MEMORANDUM

                         C. $100,000 per Occurrence per Insured Member as respects Consequential

              RT Specialty is a division of RSG Specialty, LLC, a Delaware limited liability company
                                  and a subsidiary of Ryan Specialty Group, LLC.
                   In California: RSG Specialty Insurance Services, LLC (License # 0G97516)
                              150 South US Highway 1, 3rd Floor, Jupiter, FL 33477
                                    Phone - (561) 615-5552 Fax - (561) 615-5592                     Page | 7
UQT 10 2022

--- Page 9 ---
QUOTATION MEMORANDUM

                              Reduction in Value
                          D. $100,000 per Occurrence per Insured Member as respects Expediting Expense

                      The above may not include all sublimits. Please review the full policy form and
                      endorsements.

              FORM: HotelBound Manuscript Wording

   ATTACHMENTS: HB Evidence of Coverage
                HB Shared Limits Disclosure
                HB Manuscript
              RT Specialty is a division of RSG Specialty, LLC, a Delaware limited liability company
                                  and a subsidiary of Ryan Specialty Group, LLC.
                   In California: RSG Specialty Insurance Services, LLC (License # 0G97516)
                              150 South US Highway 1, 3rd Floor, Jupiter, FL 33477
                                    Phone - (561) 615-5552 Fax - (561) 615-5592                     Page | 8
UQT 10 2022

--- Page 10 ---
QUOTATION MEMORANDUM

                     HB Schedule of Carriers with Terrorism

     WARRANTIES: Warrant loss runs on file cover all locations on the broker's statement of values.
                 Warrant no deterioration of loss history since the valuation date of loss runs on file.

              RT Specialty is a division of RSG Specialty, LLC, a Delaware limited liability company
                                  and a subsidiary of Ryan Specialty Group, LLC.
                   In California: RSG Specialty Insurance Services, LLC (License # 0G97516)
                              150 South US Highway 1, 3rd Floor, Jupiter, FL 33477
                                    Phone - (561) 615-5552 Fax - (561) 615-5592                     Page | 9
UQT 10 2022

--- Page 11 ---
QUOTATION MEMORANDUM

                     Warrant no aluminum wiring inside any building.
                     Warrant active automatic sprinkler system covering 100% of all buildings that are reported
                     sprinklered.

   SUBJECTIVITIES: Due at binding:
                      · Request to bind must be in writing
                      · Statement of Values signed by the insured prior to inception of coverage

PREMIUM AND FEES:          Non-Admitted Premium:                                 $113,729
                           Program Brokerage Fee:                                 $17,060
                           Program Inspection Fee:                                     $0
                                 *Association Fee:                                  $300

--- Page 22 ---
                                                                                                     Schedule of Values
    LID         Street        City     State   Zip   County   Flood   Year Built   ATC Occ.   Num     Num of    Num of   ISO Code   Sprinklered   Gross     Bldg. Value    BPP Value     BI Value          TIV
                                                              Zone                  Code       of     Stories    Units                 (Y/N)      Sq.Ft.
                                                                                              Bldg.



21111111101 2901 Old       Bensalem      PA 19020 Bucks         X     1960            3          1       1         0        1          N          2275     $318,500       $25,000      $10,000       $353,500
            Mill Road
21111111102 2901 Old       Bensalem      PA 19020 Bucks         X     1960            3          1       1         8        1          N          7360     $1,030,400     $100,000     $468,000      $1,598,400
            Mill Road
21111111103 6890 Lindbergh Philadelphi   PA 19153 Philade     X500    1990            3          1       2        18        1          Y          9528     $1,429,200     $355,000     $211,752      $1,995,952
            Avenue       a                        lphia
21111111104 615 North      Exton         PA 19341 Chester       X     1990            3          1       5       215        4          Y          114456 $20,000,000      $1,100,000   $1,400,000    $22,500,000
            Lancaster
            Pike
21111111105 2100 Oxford    Langhorne     PA 19047 Bucks         X     1987            3          1       3        91        1          N          28830    $4,324,500     $650,000     $700,000      $5,674,500
            Boulevard

            West
21111111106 5426 Lower   New Hope        PA 18938 Bucks          X     1983           3          1        4       135          4        N          87052    $11,500,000    $825,000     $752,250     $13,077,250
            York Road
21111111107 5630         Philadelphi     PA 19152 Philade        X     1967           3          1        2       105          2        N          45000    $7,010,700     $675,000     $700,000     $8,385,700
            Germantown   a                        lphia
            Boulevard

For Reference in the above table: ATC Occupancy Code 3 = Hotel / Motel

--- Page 27 ---
                                                                                            Cost Break Out By Location
ACCOUNT NAME: Keystone Lodging Group                                                                                                                                                             CONTROL #:
10002

      LID            Street                City      State     Zip         County      Agg      Premium          Fee       Bldg. Value       BPP Value             BI Value                TIV
                                                                                      Premiu
                                                                                        m
21111111101   2901 Old Mill Road     Bensalem       PA        19020   Bucks          $150      $859       $129         $318,500          $25,000         $10,000              $353,500




21111111102   2901 Old Mill Road     Bensalem       PA        19020   Bucks          $677      $3,813     $572         $1,030,400        $100,000        $468,000             $1,598,400




21111111103   6890 Lindbergh Avenue  Philadelphia   PA        19153   Philadelphia   $846      $4,683     $702         $1,429,200        $355,000        $211,752             $1,995,952




21111111104   615 North Lancaster Pike Exton          PA        19341   Chester        $9,536    $48,339    $7,251       $20,000,000       $1,100,000       $1,400,000         $22,500,000




21111111105   2100 Oxford Boulevard   Langhorne       PA        19047   Bucks          $2,405    $12,424    $1,864       $4,324,500        $650,000         $700,000           $5,674,500
              West




21111111106   5426 Lower York Road    New Hope        PA        18938   Bucks          $5,543    $26,207    $3,931       $11,500,000       $825,000         $752,250           $13,077,250




                                                           RT Specialty is a division of RSG Specialty, LLC, a Delaware limited liability company
                                                                                 Phone - (561) 615-5552 Fax - (561) 615-5592                     Page | 22
UQT 10 2022

--- Page 28 ---
QUOTATION MEMORANDUM

      LID             Street                 City      State     Zip         County     AggPre    Premium       Fee       Bldg. Value       BPP Value          BI Value                TIV
                                                                                         mium
21111111107   5630 Germantown Boulevard Philadelphia   PA        19152   Philadelphia   $3,554    $17,404    $2,611    $7,010,700        $675,000        $700,000          $8,385,700




                                                           RT Specialty is a division of RSG Specialty, LLC, a Delaware limited liability company

                        SHARED CAPACITY AND DEDICATED LIMITS DISCLOSURE
The Master insurance policies provide a program limit of insurance of $400,000,000 per Occurrence and
are further subject to certain program sublimits.
"""

hb = hq.parse_hotelbound_quote(HB_TEXT)
check(hb.get("detected"), "HotelBound quote detected")
sched = hb["schedule"]
check(len(sched) == 7, f"7 buildings, got {len(sched)}")
got = [(r["street"], r["city"], r["state"], r["zip"]) for r in sched]
want = [("2901 Old Mill Road", "Bensalem", "PA", "19020"), ("2901 Old Mill Road", "Bensalem", "PA", "19020"),
        ("6890 Lindbergh Avenue", "Philadelphia", "PA", "19153"), ("615 North Lancaster Pike", "Exton", "PA", "19341"),
        ("2100 Oxford Boulevard West", "Langhorne", "PA", "19047"), ("5426 Lower York Road", "New Hope", "PA", "18938"),
        ("5630 Germantown Boulevard", "Philadelphia", "PA", "19152")]
for g, w in zip(got, want):
    check(g == w, f"{g} != {w}")
check(sum(r["premium"] for r in sched) == 113729, sum(r["premium"] for r in sched))
check(sum(r["tiv"] for r in sched) == 53585302, sum(r["tiv"] for r in sched))
check(len({r["location_number"] for r in sched}) == 6, "six addresses")
# the Schedule of Values attributes still ride along with the corrected address
check(sched[6]["year_built"] == "1967" and sched[6]["num_units"] == 105, sched[6])

# lists that run over a page break carry neither the page marker nor the wholesaler's footer
check(hb["attachments"] == ["HB Evidence of Coverage", "HB Shared Limits Disclosure", "HB Manuscript"], hb["attachments"])
subl = {x["letter"]: (x["description"], x["limit"]) for x in hb["sublimits"]}
check(subl == {"A": ("Extra Expense", "$500,000 per Occurrence"),
               "B": ("Ingress / Egress — limit 5 mile radius", "Lesser of $100,000 or 30 days per Occurrence"),
               "C": ("Consequential Reduction in Value", "$100,000 per Occurrence"),
               "D": ("Expediting Expense", "$100,000 per Occurrence")}, subl)
check(hb["warranties"] == [
    "Warrant loss runs on file cover all locations on the broker's statement of values.",
    "Warrant no deterioration of loss history since the valuation date of loss runs on file.",
    "Warrant no aluminum wiring inside any building.",
    "Warrant active automatic sprinkler system covering 100% of all buildings that are reported sprinklered."],
    hb["warranties"])
check(hb["subjectivities"] == ["Request to bind must be in writing",
                               "Statement of Values signed by the insured prior to inception of coverage"],
      hb["subjectivities"])
import json as _json  # noqa: E402
_blob = _json.dumps({k: v for k, v in hb.items() if k != "program_brokerage_fee_text"}, default=str)
for junk in ("--- Page", "RSG Specialty", "Ryan Specialty", "(561)", "UQT 10 2022", "Jupiter, FL"):
    check(junk not in _blob, f"{junk!r} left in the parsed quote")

# ------------------------------------------------------------------ 3. the model's reading, then the pass
GL_PREMISES = ["6890 LINDBERGH AVENUE, Philadelphia, PA, 19153", "615 North Lancaster Pike, Exton, PA, 19341",
               "2100 Oxford Boulevard West, Langhorne, PA, 19047", "5426 Lower York Road, New Hope, PA, 18938",
               "5630 Germantown Boulevard, Philadelphia, PA, 19152", "2055 Street Rd, Bensalem, PA, 19020"]
# location labels exactly as two quotes print them: "location#4" and "Location #4" are one premises
CLASSES = [("Primary", 0, "45190"), ("location#2", 1, "45191"), ("location#2", 1, "16916"), ("location#3", 2, "45192"),
           ("location#4", 3, "45190"), ("location#4", 3, "16916"), ("location#5", 4, "45192"),
           ("location#6", 5, "45192"), ("Location #4", 3, "58175"), ("Location #2", 1, "58177")]


def model_data(premises=GL_PREMISES, classes=CLASSES):
    locs = []
    for i, a in enumerate(premises, 1):
        st, city, state, z = [p.strip() for p in a.split(",")]
        locs.append({"number": str(i), "name": "", "corporate_entity": "KEYSTONE LODGING GROUP LLC", "address": st,
                     "city": city, "state": state, "zip": z, "description": "Hotel/Motel", "tiv": 0})
    return {
        "client_info": {"named_insured": "KEYSTONE LODGING GROUP LLC", "address": "10 Main Street, Langhorne, PA 19047",
                        "effective_date": "12/01/2026", "expiration_date": "12/01/2027"},
        "locations": locs,
        "coverages": {
            "property": {"carrier": "Certain Underwriters at Lloyd's, London", "premium": 113729,
                         "total_premium": 134525.87, "tiv": "$53,585,302"},
            "general_liability": {
                "carrier": "Lakeshore Specialty Insurance Company", "premium": 200000, "total_premium": 209000,
                "designated_premises": list(premises),
                "schedule_of_classes": [{"location": l, "address": premises[i], "class_code": c,
                                         "classification": "Hotels and Motels", "exposure": "$1,000,000", "premium": "$10,000"}
                                        for l, i, c in classes if i < len(premises)],
                "limits": [{"description": "Each Occurrence", "limit": "$1,000,000"}]},
        },
    }


ITEMS = [{"filename": "Keystone Lodging Group - Hotelbound Quote.pdf", "text": HB_TEXT}]


DOC_TEXT = {}


def run(data, tag=""):
    hq.apply_hotelbound_to_data(data, hq.parse_hotelbound_quote(HB_TEXT), ITEMS)
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "p.docx")
        pg.generate_proposal(copy.deepcopy(data), out)
        d = docx.Document(out)
        tables = [[[c.text for c in r.cells] for r in t.rows] for t in d.tables]
        DOC_TEXT[tag] = "\n".join([p.text for p in d.paragraphs] + [c for t in tables for r in t for c in r])
    sched_tbl = next(t for t in tables if t[0][:3] == ["#", "Property Name", "Address"])
    info = {r[0]: r[1] for t in tables for r in t if len(r) == 2 and r[0].startswith(("Total Number", "Property Loc",
                                                                                      "Liability Loc", "Location Type"))}
    return sched_tbl, info


data = model_data()
sched_tbl, info = run(data)

# the review screen: six scheduled hotels, then the one the liability quote alone lists
locs = data["locations"]
check(len(locs) == 7, f"7 locations on the review screen, got {len(locs)}")
check([l["address"] for l in locs[:6]] == ["2901 Old Mill Road", "6890 Lindbergh Avenue", "615 North Lancaster Pike",
                                             "2100 Oxford Boulevard West", "5426 Lower York Road",
                                             "5630 Germantown Boulevard"], [l["address"] for l in locs])
last = locs[6]
check((last["address"], last["city"], last["state"], last["zip"]) == ("2055 Street Road", "Bensalem", "PA", "19020"), last)
check(last["tiv"] == 0 and last["number"] == "7", last)
check(sum(l["tiv"] for l in locs) == 53585302, "the liability-only hotel adds no insured value")
notes = [w for w in data.get("warnings", []) if str(w).startswith("Locations check")]
check(len(notes) == 1, notes)
check("On the liability quote only: 2055 Street Road, Bensalem." in notes[0], notes[0])
check("On the property schedule only: 2901 Old Mill Road, Bensalem." in notes[0], notes[0])
check("--" not in notes[0], "no double dashes")

# the proposal: Schedule of Locations
body = [r for r in sched_tbl[1:] if r[0].strip().isdigit()]
check(len(body) == 7, [r[2] for r in body])
marks = {r[2]: (r[-2], r[-1]) for r in body}
check(marks["2901 Old Mill Road"] == ("✓", "—"), f"property only: {marks['2901 Old Mill Road']}")
check(marks["2055 Street Road"] == ("—", "✓"), f"liability only: {marks['2055 Street Road']}")
for a in ("6890 Lindbergh Avenue", "615 North Lancaster Pike", "2100 Oxford Boulevard West", "5426 Lower York Road",
          "5630 Germantown Boulevard"):
    check(marks[a] == ("✓", "✓"), f"{a}: {marks[a]}")
check(next(r for r in body if r[2] == "2055 Street Road")[5].strip() in ("", "—"), "no insured value printed")
check(sched_tbl[-1][5] == "$53,585,302", sched_tbl[-1])

# the proposal: Information Summary
check(info.get("Total Number of Locations") == "7", info)
check(info.get("Property Locations") == "6", info)
check(info.get("Liability Locations") == "6", info)
check(info.get("Location Types") == "7 Hotel(s)", info)          # was "6 Hotel(s)" beside 7 locations
# nothing of the wholesaler's page footer or the page markers reaches the client
for junk in ("--- Page", "RSG Specialty", "Ryan Specialty", "(561) 615", "Jupiter, FL", "License # 0G97516"):
    check(junk not in DOC_TEXT[""], f"{junk!r} printed in the proposal")
check("Consequential Reduction in Value" in DOC_TEXT[""], "sublimit that wraps over a page break")

# running the pass again (Refresh, or a second Generate) adds nothing
before = copy.deepcopy(data["locations"])
hq.apply_hotelbound_to_data(data, hq.parse_hotelbound_quote(HB_TEXT), ITEMS)
check([l["address"] for l in data["locations"]] == [l["address"] for l in before], "idempotent locations")
check(len([w for w in data["warnings"] if str(w).startswith("Locations check")]) == 1, "one note, not two")

# ------------------------------------------------------------------ 4. the liability quote lists the same hotels
same = ["2901 Old Mill Road, Bensalem, PA, 19020"] + GL_PREMISES[:5]
data2 = model_data(premises=same, classes=[("Primary", 0, "45190")] + [(f"location#{i}", i - 1, "45192") for i in range(2, 7)])
sched2, info2 = run(data2)
check(len(data2["locations"]) == 6, "nothing added when both quotes list the same six")
check(not [w for w in data2.get("warnings", []) if str(w).startswith("Locations check")], data2.get("warnings"))
body2 = [r for r in sched2[1:] if r[0].strip().isdigit()]
check(len(body2) == 6 and all((r[-2], r[-1]) == ("✓", "✓") for r in body2), [(r[2], r[-2], r[-1]) for r in body2])
check((info2.get("Total Number of Locations"), info2.get("Property Locations"), info2.get("Liability Locations"))
      == ("6", "6", "6"), info2)
check(info2.get("Location Types") == "6 Hotel(s)", info2)

# ------------------------------------------------------------------ 5. one extra hotel on the liability quote
extra = same + ["2055 Street Rd, Bensalem, PA, 19020"]
data3 = model_data(premises=extra, classes=[("Primary", 0, "45190")] + [(f"location#{i}", i - 1, "45192") for i in range(2, 8)])
sched3, info3 = run(data3)
check(len(data3["locations"]) == 7, "liability-only hotel kept")
n3 = [w for w in data3.get("warnings", []) if str(w).startswith("Locations check")]
check(len(n3) == 1 and "on the liability quote but not on the property schedule: 2055 Street Road, Bensalem" in n3[0], n3)
body3 = {r[2]: (r[-2], r[-1]) for r in sched3[1:] if r[0].strip().isdigit()}
check(body3["2055 Street Road"] == ("—", "✓"), body3)
check(body3["2901 Old Mill Road"] == ("✓", "✓"), body3)
check((info3.get("Total Number of Locations"), info3.get("Property Locations"), info3.get("Liability Locations"))
      == ("7", "6", "7"), info3)
check(info3.get("Location Types") == "7 Hotel(s)", info3)

# ------------------------------------------------------------------ 6. the street type is spelled out, nothing else
sp = hq._spell_street_suffix
check(sp("2055 Street Rd") == "2055 Street Road", sp("2055 Street Rd"))
check(sp("100 Main St.") == "100 Main Street", sp("100 Main St."))
check(sp("3100 Cabot Blvd W") == "3100 Cabot Boulevard W", sp("3100 Cabot Blvd W"))
check(sp("12 St Charles Ave") == "12 St Charles Avenue", sp("12 St Charles Ave"))
check(sp("815 North Pottstown Pike") == "815 North Pottstown Pike", "already spelled out")

# ------------------------------------------------------------------ 7. premises labels on the liability quote
check(pg._gl_location_id("location#4") == pg._gl_location_id("Location #4") == "loc4", pg._gl_location_id("Location #4"))
check(pg._gl_location_id("Primary") == "loc1", pg._gl_location_id("Primary"))

print(f"liability_only_locations_test: {PASS} checks passed")
