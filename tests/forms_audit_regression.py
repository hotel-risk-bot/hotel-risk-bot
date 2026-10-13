"""Regression tests for forms_schedule_audit.py.

Run with plain Python (no dependencies):   python tests/forms_audit_regression.py
or with pytest if it is installed:          pytest tests/forms_audit_regression.py

Cases:
  * Berkley Aspire GL quote (Oct 13 2026, Triumph Investments Group): 50 forms with
    letter-suffixed numbers, a glued edition, titles wrapped around a lone number,
    page footers beside form numbers and a TRIA notice after the schedule.
  * Houston Specialty style excess schedule (title first, number last, upper-case
    wrapped titles), the layout the audit was built for in Sep 2026 (Neptune Lodging).
  * Futuristic style package schedule with a Line column (Oct 6 2026, Bloom Ventures).
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

from forms_schedule_audit import (  # noqa: E402
    parse_forms_schedule,
    reconcile_coverage_forms,
    run_forms_schedule_audit,
)

# The Forms & Endorsements schedule printed on the Berkley quote, in order.
BERKLEY_FORMS = [
    "EMN2002 1024", "EMJ0001 1023", "IL0017 1198", "IL0021 0908", "IL 02 77 0726",
    "EMN2000 1023", "EMN2003 1023", "EMN2006 1023", "PN0017 0719", "EMD1000 0725",
    "CG0001 0413", "CG2029 1219", "CG2107 0514", "CG2136 0305", "CG2144 0417",
    "CG2147 1207", "CG2173 0115", "CG2196 0305", "CG2245 0413", "EMO30011023",
    "CLOC 0502", "BR013 0313", "EGD1001 1024", "BR1400 0124", "BR1411 0124",
    "BR5269G 1024", "BR902I 0816", "BR5294G 0124", "BR5191G 0124", "BR5159G 0124",
    "BR5121G 0621", "BR5116G 0321", "BR5328G 0124", "BR5287G 0124", "BR5259G 0124",
    "BR5175G 0124", "BR5154G 0124", "BR5118G 0521", "BR5006G 1012", "BR5304G 0124",
    "BR5280G 0124", "BR5239G 0124", "BR5172G 0124", "BR5128G 0525", "BR5302G 0124",
    "BR5275G 0124", "BR5194G 0124", "BR5163G 0124", "EMO3000 1023", "BR5127G 0523",
]
BERKLEY_FILE = "Commercial+Quote-2642731876 (1).pdf"


def _berkley_text() -> str:
    return BERKLEY_QUOTE_TEXT.strip("\n")


def _berkley_extracted(skip=(), swap=None):
    """What the extractor hands the audit: every schedule row, minus `skip`, with any
    number replaced per `swap` (simulating GPT misreads)."""
    rows = parse_forms_schedule(_berkley_text())
    by_num = {r["form_number"]: r["description"] for r in rows}
    out = []
    for num in BERKLEY_FORMS:
        if num in skip:
            continue
        out.append({"form_number": (swap or {}).get(num, num), "description": by_num[num]})
    return out


# ── Berkley Aspire ──────────────────────────────────────────────────────────

def test_berkley_reads_all_50_rows_in_quote_order():
    rows = parse_forms_schedule(_berkley_text())
    assert [r["form_number"] for r in rows] == BERKLEY_FORMS, [r["form_number"] for r in rows]


def test_berkley_page_footers_are_not_rows():
    rows = parse_forms_schedule(_berkley_text())
    nums = {r["form_number"] for r in rows}
    assert "ENP9014 1024" not in nums and "ENP5101 1124" not in nums
    assert not any("Page " in r["description"] for r in rows)


def test_berkley_wrapped_titles_are_rebuilt():
    d = {r["form_number"]: r["description"] for r in parse_forms_schedule(_berkley_text())}
    assert d["CG2107 0514"] == ("Exclusion - Access Or Disclosure Of Confidential Or Personal Information "
                                "And Data-Related Liability - Limited Bodily Injury Exception Not Included")
    assert d["BR5287G 0124"] == ("Additional Insured - Primary and Noncontributory - Automatic Status When "
                                 "Required In a Written Contract Or Agreement (Including Completed Operations)")
    assert d["BR5280G 0124"] == ("Waiver of Transfer of Rights of Recovery Against Others to Us -Automatic "
                                 "Status When Required in Contract or Agreement (Blanket)")
    assert d["EMO30011023"] == "L500 - Liability for Guests' Property"
    assert d["BR902I 0816"] == "Changes - Civil Union or Domestic Partnership"


def test_berkley_complete_extraction_raises_no_warning():
    data = {"coverages": {"general_liability": {
        "carrier": "Berkley Specialty Insurance Company",
        "forms_endorsements": _berkley_extracted()}}}
    warnings = run_forms_schedule_audit(data, [{"filename": BERKLEY_FILE, "text": _berkley_text()}])
    assert warnings == [], warnings
    audit = data["coverages"]["general_liability"]["forms_audit"]
    assert audit["schedule_rows"] == 50 and audit["added"] == [] and audit["final_count"] == 50, audit


def test_berkley_missing_rows_are_added_in_quote_order():
    missing = ("CG2107 0514", "EMO30011023", "BR5118G 0521", "BR5287G 0124")
    cov = {"forms_endorsements": _berkley_extracted(skip=missing)}
    audit = reconcile_coverage_forms(cov, _berkley_text(), BERKLEY_FILE, "general_liability")
    assert sorted(a["form_number"] for a in audit["added"]) == sorted(missing), audit["added"]
    assert [f["form_number"] for f in cov["forms_endorsements"]] == BERKLEY_FORMS
    assert all(f.get("schedule_verified") for f in cov["forms_endorsements"] if f["form_number"] in missing)


def test_gpt_character_swaps_and_spacing_are_not_added_twice():
    cov = {"forms_endorsements": _berkley_extracted(swap={
        "BR902I 0816": "BR9021 0816",        # I read as 1
        "EMO3000 1023": "EM03000 1023",      # O read as 0
        "CG2107 0514": "CG 21 07 05 14",     # ISO spacing
        "EMO30011023": "EMO3001 1023",       # edition split back out
    })}
    audit = reconcile_coverage_forms(cov, _berkley_text(), BERKLEY_FILE, "general_liability")
    assert audit["added"] == [], audit["added"]
    assert len(cov["forms_endorsements"]) == 50


# ── Houston Specialty style excess schedule (title first, number last) ───────

EXCESS_TEXT = """--- Page 3 ---
                         SCHEDULE OF FORMS AND ENDORSEMENTS
Form Title                                                              Form Number
Commercial Excess Liability Policy Jacket                               HSIC JACKET 01 21
COMMERCIAL EXCESS LIABILITY COVERAGE FORM                               HSIC CX ES 00 01 04 21
EXCLUSION - ASSAULT OR BATTERY                                          HSIC CX ES 01 44 04 21
EXCLUSION - TOTAL FIREARMS                                              HSIC CX ES 01 52 04 21
EXCLUSION OF COVERAGE SUBJECT TO SUBLIMITS OR OTHER REDUCED LIMITS      HSIC CX ES 01 60 04 21
IN UNDERLYING INSURANCE
EXCLUSION - SEXUAL MISCONDUCT                                           HSIC CX ES 01 61 04 21
NUCLEAR ENERGY LIABILITY EXCLUSION                                      HSIC CX ES 01 02 04 21
PUNITIVE OR EXEMPLARY DAMAGES EXCLUSION                                 HSIC CX ES 01 09 04 21
DISCLOSURE PURSUANT TO TERRORISM RISK INSURANCE ACT                     IL 09 85 12 20
SERVICE OF SUIT                                                         HSIC IL 70 01 04 21

HSIC CX DEC 04 21                                                       Page 2 of 3
--- Page 4 ---
                    COMMERCIAL EXCESS LIABILITY DECLARATIONS
Named Insured:  Sample Lodging Partners LLC
"""


def test_excess_schedule_title_first_layout():
    rows = parse_forms_schedule(EXCESS_TEXT)
    nums = [r["form_number"] for r in rows]
    assert nums == ["HSIC JACKET 01 21", "HSIC CX ES 00 01 04 21", "HSIC CX ES 01 44 04 21",
                    "HSIC CX ES 01 52 04 21", "HSIC CX ES 01 60 04 21", "HSIC CX ES 01 61 04 21",
                    "HSIC CX ES 01 02 04 21", "HSIC CX ES 01 09 04 21", "IL 09 85 12 20",
                    "HSIC IL 70 01 04 21"], nums
    d = {r["form_number"]: r["description"] for r in rows}
    assert d["HSIC CX ES 01 60 04 21"] == ("EXCLUSION OF COVERAGE SUBJECT TO SUBLIMITS OR OTHER REDUCED "
                                           "LIMITS IN UNDERLYING INSURANCE")
    cov = {"forms_endorsements": [{"form_number": n, "description": d[n]} for n in nums[:6]]}
    audit = reconcile_coverage_forms(cov, EXCESS_TEXT, "xs.pdf", "umbrella")
    assert [a["form_number"] for a in audit["added"]] == nums[6:], audit["added"]


# ── Futuristic style package schedule with a Line column ─────────────────────

PACKAGE_TEXT = """--- Page 6 ---
FORMS AND ENDORSEMENTS
Form Number   Edition   Form Title                                         Line
IL 00 17      11/98     Common Policy Conditions                           Interline
IL 00 21      09/08     Nuclear Energy Liability Exclusion Endorsement     Interline
CG 00 01      04/13     Commercial General Liability Coverage Form         General Liability
CG 21 47      12/07     Employment-Related Practices Exclusion             General Liability
FUT-SS 01     01/22     Assault And Battery Sublimit                       General Liability
GLF 101       01/22     Human Trafficking Exclusion                        General Liability
CG 00 33      04/13     Liquor Liability Coverage Form                     Liquor Liability
CG 21 51      09/99     Amendment Of Liquor Liability Exclusion            Liquor Liability
CA 00 01      10/13     Business Auto Coverage Form                        Business Auto
FUT 1025      01/25     Vehicle Schedule                                   Business Auto
CA 23 84      10/13     Exclusion Of Terrorism                             Business Auto
"""


def test_package_schedule_line_column():
    data = {"coverages": {
        "general_liability": {"carrier": "Futuristic Underwriters", "forms_endorsements": []},
        "commercial_auto": {"carrier": "Futuristic Underwriters", "forms_endorsements": []},
    }}
    text = PACKAGE_TEXT + "\nGeneral Liability Business Auto Futuristic Underwriters general liability business auto\n"
    run_forms_schedule_audit(data, [{"filename": "package.pdf", "text": text}])
    gl = [f["form_number"] for f in data["coverages"]["general_liability"]["forms_endorsements"]]
    auto = [f["form_number"] for f in data["coverages"]["commercial_auto"]["forms_endorsements"]]
    assert gl == ["IL 00 17 11/98", "IL 00 21 09/08", "CG 00 01 04/13", "CG 21 47 12/07",
                  "FUT-SS 01 01/22", "GLF 101 01/22", "CG 00 33 04/13", "CG 21 51 09/99"], gl
    assert auto == ["IL 00 17 11/98", "IL 00 21 09/08", "CA 00 01 10/13", "FUT 1025 01/25",
                    "CA 23 84 10/13"], auto


# ── Page handling ─────────────────────────────────────────────────────────────

def test_schedule_runs_across_a_page_break_without_a_new_header():
    text = """--- Page 1 ---
Forms & Endorsements
EMN2002 1024    Surplus Lines Producer Information

BR5269G 1024    Privacy Breach Coverage

ENP9014 1024                                         Page 1 of 2
--- Page 2 ---
BR5118G 0521    Exclusion - Total Liquor Liability

ENP9014 1024                                         Page 2 of 2
"""
    nums = [r["form_number"] for r in parse_forms_schedule(text)]
    assert nums == ["EMN2002 1024", "BR5269G 1024", "BR5118G 0521"], nums


def test_schedule_ends_before_the_next_document():
    # A lone numbered line in the notice after the schedule must not join the schedule.
    text = """--- Page 1 ---
Forms & Endorsements
CG0001 0413     Commercial General Liability Coverage Form

CG2147 1207     Employment-Related Practices Exclusion
--- Page 2 ---
POLICYHOLDER NOTICE
You are hereby notified that under the Terrorism Risk Insurance Act you have a right to purchase coverage.
NOTICE ABC 1234    Acceptance Or Rejection Of Terrorism Insurance Coverage
"""
    nums = [r["form_number"] for r in parse_forms_schedule(text)]
    assert nums == ["CG0001 0413", "CG2147 1207"], nums


def test_copyright_footer_lines_are_not_rows():
    text = """SCHEDULE OF FORMS AND ENDORSEMENTS
Commercial General Liability Coverage Form                       CG 00 01 04 13
Employment-Related Practices Exclusion                           CG 21 47 12 07
CG 21 47 12 07          © Insurance Services Office, Inc., 2006          Page 1 of 1
Includes copyrighted material of ISO Properties, Inc.            CG 21 70 01 15
"""
    nums = [r["form_number"] for r in parse_forms_schedule(text)]
    assert nums == ["CG 00 01 04 13", "CG 21 47 12 07"], nums


# Berkley Aspire (Berkley Specialty Insurance Co) GL quote, pages 2-5, exactly as the proposal
# extractor reads it (pdftotext -layout per page, "--- Page N ---" markers). Insured name and
# premises address replaced; the forms schedule, page footers and TRIA notice are verbatim.
BERKLEY_QUOTE_TEXT = """
--- Page 2 ---
Commercial General Liability Quote
     Premises/Operations and Products/Completed Operations – Occurrence

Limits of Liability                                                                                                Deductible
Each Occurrence                                               $1,000,000                                                    Per Claim
General Aggregate                                             $2,000,000                                                       $2,500
Products/Completed Operations Aggregate                       INCLUDED
Personal/Advertising Injury                                   $1,000,000
Damage To Premises Rented To You                                $100,000
Medical Payments (Per Person)                                 EXCLUDED



 Classification Description               Code      Exposure        Basis(per)     Prem/Ops          Pr/Co        Advance Premium
                                                                                      Rate            Rate
Hotels and Motels - With Pools or         45191     $2,700,000       $1,000          $7.000                                 $18,900.00
Beaches - 4 to 8 Stories                                           Gross Sales
                                                                                                    Included                 Included

Additional Coverages and Other Premium Charges
Hired Auto and Non-Owned Auto Liability       $1,000,000    Hired Auto Limit of Insurance     Fully Earned                   $1,000.00
(Specified Limits of Insurance)                             (Any One Accident)
                                              $1,000,000    Non-Owned Auto Limit of
                                                            Insurance (Any One
                                                            Accident)
Privacy Breach Coverage                           $25,000   Endorsement Aggregate                                              $130.00
                                                            Coverage Limit
                                                  $25,000   First Party Coverage Limit
                                                  $25,000   Regulatory Proceeding
                                                            Coverage Limit
Deductible                                         $1,000
Terrorism Coverage                                                                                             See Terrorism Disclosure
                                                                                                                                Notice
                                                                                 Total Advance Premium                      $20,030.00
                                                                                 Minimum Term Premium                            $0.00


Schedule of Surcharges, Taxes and Fees                                                                                        Amount
Policy Fee                                                                                                                     $500.00
Surplus Lines Tax                                                                                                              $955.88
Surplus Lines Service Charge                                                                                                    $39.83
Stamping Fee                                                                                                                    $35.93
                                                                         Total Surcharges, Taxes and Fees                    $1,531.64
                                                    Total including Premium, Surcharges, Taxes and Fees                   $21,561.64




ENP9014 1024                                                                                                              Page 2 of 4

--- Page 3 ---
Locations

 Location #     GL Class Code                Address
 1                     45191                 100 Sample Road, Anytown, LA 70000


Forms & Endorsements
EMN2002 1024    Surplus Lines Producer Information

EMJ0001 1023    Commercial Lines Policy Jacket

IL0017 1198     Common Policy Conditions

IL0021 0908     Nuclear Energy Liability Exclusion Endorsement

IL 02 77 0726   Louisiana Changes - Cancellation And Nonrenewal

EMN2000 1023    U.S. Treasury Department's Office of Foreign Assets Control (OFAC) Advisory Notice To Policyholders

EMN2003 1023    W.R. Berkley Corporation Notice of Privacy Policies

EMN2006 1023    Claim Reporting Information

PN0017 0719     Important Notice - Privacy Breach Coverage Insurance

EMD1000 0725    Common Policy Declarations

CG0001 0413     Commercial General Liability Coverage Form

CG2029 1219     Additional Insured - Grantor Of Franchise

                Exclusion - Access Or Disclosure Of Confidential Or Personal Information And Data-Related Liability - Limited
CG2107 0514
                Bodily Injury Exception Not Included

CG2136 0305     Exclusion - New Entities

CG2144 0417     Limitation Of Coverage To Designated Premises, Project Or Operation

CG2147 1207     Employment-Related Practices Exclusion

CG2173 0115     Exclusion of Certified Acts of Terrorism

CG2196 0305     Silica or Silica-Related Dust Exclusion

CG2245 0413     Exclusion - Specified Therapeutic Or Cosmetic Services

EMO30011023     L500 - Liability for Guests' Property

CLOC 0502       Location Schedule

BR013 0313      Minimum Earned Premium Endorsement

EGD1001 1024    Commercial General Liability Coverage Part Declarations

BR1400 0124     Service Of Suit

BR1411 0124     Economic Or Trade Sanctions Endorsement

BR5269G 1024    Privacy Breach Coverage

BR902I 0816     Changes - Civil Union or Domestic Partnership

BR5294G 0124    Deductible Liability Insurance (Including Allocated Loss Adjustment Expense)

BR5191G 0124    Conditional Exclusion - Swimming Pool Barrier Requirement (Self-Closing and Positive Self-Latching Mechanism)

BR5159G 0124    Amendment of Definitions -Insured Contract (Limited Form)

BR5121G 0621    Exclusion - Punitive or Exemplary Damages

BR5116G 0321    Exclusion - Injury to Employees, Contractors, Volunteers and Other Workers

BR5328G 0124    Exclusion - Asbestos

                Additional Insured - Primary and Noncontributory - Automatic Status When Required In a Written Contract Or
BR5287G 0124
                Agreement (Including Completed Operations)

BR5259G 0124    Exclusion - Communicable or Infectious Disease

BR5175G 0124    Exclusion - Microorganisms, Biological Organisms, Bioaerosols or Organic Contaminants


ENP9014 1024                                                                                                       Page 3 of 4

--- Page 4 ---
Forms & Endorsements (Cont'd)
BR5154G 0124   Exclusion - Medical Payments

BR5118G 0521   Exclusion - Total Liquor Liability

BR5006G 1012   Exclusion - Cross Suits

BR5304G 0124   Conditional Exclusion -Water Related Hazard Sign Posting Requirement

               Waiver of Transfer of Rights of Recovery Against Others to Us -Automatic Status When Required in Contract or
BR5280G 0124
               Agreement (Blanket)

BR5239G 0124   Exclusion - Continuous or Progressive Damage

BR5172G 0124   Exclusion -Toxic Metals

BR5128G 0525   Total Perfluoroalkyl And Polyfluoroalkyl Substances (Pfas) Exclusion

BR5302G 0124   Conditional Exclusion -Lifesaving Equipment

BR5275G 0124   Amendment of Conditions - Premium Audit

BR5194G 0124   Hired Auto and Non-Owned Auto Liability (Specified Limits of Insurance)

BR5163G 0124   Exclusion -Total Pollution

EMO3000 1023   Schedule Of Forms And Endorsements

BR5127G 0523   Exclusion - Biometric Information




ENP9014 1024                                                                                                   Page 4 of 4

--- Page 5 ---
POLICYHOLDER NOTICE
                    ACCEPTANCE OR REJECTION OF TERRORISM INSURANCE COVERAGE
        POLICYHOLDER NOTICE ACCEPTANCE OR REJECTION OF TERRORISM INSURANCE COVERAGE_EN P5 001_10/23
You are hereby notified that under the Terrorism Risk Insurance Act, as amended, (the “Act”), you have a right to purchase
insurance coverage for losses resulting from acts of terrorism, as defined in Section 102(1) of the Act: The term “act of
terrorism” means any act that is certified by the Secretary of the Treasury—in consultation with the Secretary of Homeland
Security, and the Attorney General of the United States—to be an act of terrorism; to be a violent act or an act that is
dangerous to human life, property, or infrastructure; to have resulted in damage within the United States, or outside the United
States in the case of certain air carriers or vessels or the premises of a United States mission; and to have been committed by
an individual or individuals as part of an effort to coerce the civilian population of the United States or to influence the policy or
affect the conduct of the United States Government by coercion. Coverage under your policy may be affected as follows:
IF YOU ARE PURCHASING COMMERCIAL PROPERTY COVERAGE IN THE STATES OF CALIFORNIA, GEORGIA,
HAWAII, ILLINOIS, IOWA, MAINE, MISSOURI, NEW JERSEY, NEW YORK, NORTH CAROLINA, OREGON, RHODE
ISLAND, WASHINGTON, WISCONSIN OR WEST VIRGINIA; AND/OR PURCHASING COMMERCIAL INLAND MARINE
COVERAGE IN THE STATES OF CALIFORNIA, MAINE, MISSOURI, OREGON OR WISCONSIN THERE ARE STATE
STATUTORY EXCEPTIONS COVERING CERTAIN FIRE LOSSES IF YOU DECLINE COVERAGE FOR “ACTS OF
TERRORISM” DEFINED UNDER THE ACT. IF AN “ACT OF TERRORISM” CERTIFIED UNDER THE ACT RESULTS IN
FIRE, WE ARE REQUIRED TO PAY FOR THE LOSS OR DAMAGE CAUSED BY THAT FIRE. SUCH COVERAGE FOR FIRE
APPLIES ONLY TO DIRECT LOSS OR DAMAGE BY FIRE TO COVERED PROPERTY AND IS SUBJECT TO ANY
LIMITATIONS OF ANY TERRORISM EXCLUSION, OR INAPPLICABILITY OR OMISSION OF A TERRORISM EXCLUSION.
THIS NOTICE DOES NOT SERVE TO CREATE COVERAGE FOR ANY LOSS WHICH WOULD OTHERWISE BE
EXCLUDED UNDER YOUR POLICY.
THE PORTION OF YOUR PREMIUM THAT IS ATTRIBUTABLE TO COVERAGE FOR DIRECT LOSS OR DAMAGE THAT IS
CAUSED BY AN “ACT OF TERRORISM” CERTIFIED UNDER THE ACT AND WHERE FIRE ENSUES IS $25, AND DOES
NOT INCLUDE ANY CHARGES FOR THE PORTION OF LOSSES COVERED BY THE UNITED STATES GOVERNMENT
UNDER THE ACT. NOTE – THIS PREMIUM IS APPLIED TO YOUR POLICY REGARDLESS IF YOU ACCEPT OR DECLINE
COVERAGE FOR “ACTS OF TERRORISM” BELOW.
YOU SHOULD KNOW THAT WHERE COVERAGE IS PROVIDED BY THIS POLICY FOR LOSSES RESULTING FROM
CERTIFIED ACTS OF TERRORISM, SUCH LOSSES MAY BE PARTIALLY REIMBURSED BY THE UNITED STATES
GOVERNMENT UNDER A FORMULA ESTABLISHED BY FEDERAL LAW. HOWEVER, YOUR POLICY MAY CONTAIN
OTHER EXCLUSIONS WHICH MIGHT AFFECT YOUR COVERAGE, SUCH AS AN EXCLUSION FOR NUCLEAR EVENTS.
UNDER THE FORMULA, THE UNITED STATES GOVERNMENT GENERALLY REIMBURSES 80% BEGINNING ON
JANUARY 1, 2020 OF COVERED TERRORISM LOSSES EXCEEDING THE STATUTORILY ESTABLISHED DEDUCTIBLE
PAID BY THE INSURANCE COMPANY PROVIDING THE COVERAGE. THE PREMIUM CHARGED FOR THIS COVERAGE
IS PROVIDED BELOW AND DOES NOT INCLUDE ANY CHARGES FOR THE PORTION OF LOSS THAT MAY BE
COVERED BY THE FEDERAL GOVERNMENT UNDER THE ACT.
YOU SHOULD ALSO KNOW THAT THE ACT, CONTAINS A $100 BILLION CAP THAT LIMITS U.S. GOVERNMENT
REIMBURSEMENT AS WELL AS INSURERS’ LIABILITY FOR LOSSES RESULTING FROM CERTIFIED ACTS OF
TERRORISM WHEN THE AMOUNT OF SUCH LOSSES IN ANY ONE CALENDAR YEAR EXCEEDS $100 BILLION. IF THE
AGGREGATE INSURED LOSSES FOR ALL INSURERS EXCEED $100 BILLION, YOUR COVERAGE MAY BE REDUCED.
    Acceptance or Rejection of Terrorism Insurance Coverage
     I hereby elect to purchase terrorism coverage, subject to the limitations of the Act, for acts of terrorism as defined in
     the Act, for a prospective premium of $125 plus applicable lines taxes and fees.
     I hereby decline to purchase terrorism coverage for certified acts of terrorism. I understand that I will have no coverage
     for losses resulting from certified acts of terrorism.


                                                                                     Berkley Specialty Insurance Company
             Policyholder/Applicant’s Signature                                               Insurance Company


                         Print Name                                                               Policy Number

                                                                                        Sample Hotel Group LLC
                             Date                                                               Named Insured




ENP5101 1124                                                                                                                 Page 1 of 1
                    Includes copyrighted material of National Association of Insurance Commissioners, with its permission.
"""


if __name__ == "__main__":
    failed = 0
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    for name, fn in tests:
        try:
            fn()
            print(f"PASS  {name}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"FAIL  {name}: {exc!r}")
    print(f"\n{len(tests) - failed}/{len(tests)} passed")
    sys.exit(1 if failed else 0)
