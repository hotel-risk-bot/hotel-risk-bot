"""
Constants and lookup tables for the HUB proposal generator.
Extracted from proposal_generator.py for maintainability.
"""

import os
from docx.shared import RGBColor

# HUB Design System Colors
ELECTRIC_BLUE = RGBColor(0x16, 0x7B, 0xD4)  # #167BD4
CLASSIC_BLUE = RGBColor(0x26, 0x38, 0x45)   # #263845
ARCTIC_GRAY = RGBColor(0xB8, 0xC4, 0xCE)    # #B8C4CE
EGGSHELL = RGBColor(0xF3, 0xF5, 0xF1)       # #F3F5F1
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x4A, 0x4A, 0x4A)

ELECTRIC_BLUE_HEX = "167BD4"
CLASSIC_BLUE_HEX = "263845"
ARCTIC_GRAY_HEX = "B8C4CE"
EGGSHELL_HEX = "F3F5F1"

# Logo path
LOGO_PATH = os.path.join(os.path.dirname(__file__), "templates", "hub_logo_horizontal.png")

# AM Best Rating Lookup Table - common hospitality insurance carriers
# Updated periodically; used as fallback when quote doesn't include rating
AM_BEST_RATINGS = {
    # Property carriers
    "vantage risk": "A- (VII)",
    "vantage risk specialty insurance": "A- (VII)",
    "tower hill insurance": "A- (VII)",
    "tower hill prime insurance": "A- (VII)",
    "tower hill preferred insurance": "A- (VII)",
    "starr surplus lines": "A (XV)",
    "starr indemnity": "A (XV)",
    "lexington insurance": "A+ (XV)",
    "scottsdale insurance": "A+ (XV)",
    "zurich": "A+ (XV)",
    "zurich american insurance": "A+ (XV)",
    "great lakes insurance": "A+ (XV)",
    "lloyd's of london": "A (XV)",
    "lloyds": "A (XV)",
    "certain underwriters at lloyd's, london": "A (XV)",
    "certain underwriters at lloyd's": "A (XV)",
    "owners first association": "A (XV)",
    "nautilus insurance": "A+ (XV)",
    "empire indemnity": "A (VIII)",
    "colony insurance": "A (VIII)",
    "james river insurance": "A- (VIII)",
    "canopius": "A- (VII)",
    # GL carriers
    "amtrust": "A- (VIII)",
    "amtrust e&s": "A- (VIII)",
    "amtrust financial": "A- (VIII)",
    "associated industries": "A- (VIII)",
    "associated industries insurance": "A- (VIII)",
    "technology insurance company": "A- (VIII)",
    "security national insurance": "A- (VIII)",
    "wesco insurance": "A- (VIII)",
    "southlake specialty insurance": "A- (VIII)",
    "southlake specialty": "A- (VIII)",
    "futuristic underwriters": "A- (VIII)",
    "kinsale insurance": "A (VIII)",
    "kinsale insurance company": "A (VIII)",
    "gotham insurance": "A- (VIII)",
    "gotham insurance company": "A- (VIII)",
    "coaction": "A- (VIII)",
    "coaction specialty insurance": "A- (VIII)",
    "markel insurance": "A (XV)",
    "evanston insurance": "A+ (XV)",
    "general star indemnity": "A++ (XV)",
    "essentia insurance": "A- (VII)",
    "mount vernon fire insurance": "A++ (XV)",
    # Umbrella/Excess carriers
    "palms insurance": "A- (VII)",
    "palms insurance company": "A- (VII)",
    "starstone": "A- (VII)",
    "starstone national insurance": "A- (VII)",
    "ironshore specialty insurance": "A (XV)",
    "westchester surplus lines": "A+ (XV)",
    "great american insurance": "A+ (XV)",
    "argo group": "A- (VIII)",
    "hudson insurance": "A+ (XV)",
    # Crime carriers
    "federal insurance company": "A++ (XV)",
    "federal insurance": "A++ (XV)",
    "chubb": "A++ (XV)",
    "vigilant insurance": "A++ (XV)",
    # WC carriers
    "employers insurance": "A (VIII)",
    "employers compensation insurance": "A (VIII)",
    "zenith insurance": "A (VII)",
    "pinnacol assurance": "A (VII)",
    "texas mutual insurance": "A (VIII)",
    "state compensation insurance fund": "A (VIII)",
    # Auto carriers
    "national interstate insurance": "A (VIII)",
    # Flood carriers
    "selective insurance": "A+ (VIII)",
    "selective": "A+ (VIII)",
    "wright flood": "N/A (NFIP)",
    # Multi-line carriers
    "travelers": "A++ (XV)",
    "hartford": "A+ (XV)",
    "cna": "A (XV)",
    "liberty mutual": "A (XV)",
    "nationwide": "A+ (XV)",
    "berkshire hathaway": "A++ (XV)",
    "aig": "A (XV)",
}
