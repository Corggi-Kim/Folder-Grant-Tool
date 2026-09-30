from typing import Final

DOMAIN_EMAIL_SUFFIX: Final = "lskglobal.com"

LEVEL3_CHOICES: Final = {
    "ARS", "CO", "DM", "ER", "MW", "PM", "PV", "RA", "SSU", "STAT",
    "STAT_IDMC", "ETC",
}

ROLE_MAP: Final = {
    "Trial STAT/SP": ["3.Dataset", "4.Analysis", "5.SDTM", "6.Validation"],
    "Verification SP": ["3.Dataset", "6.Validation", "8.Verification"],
    "SDTM": ["3.Dataset", "5.SDTM", "6.Validation"],
    "Manager": ["3.Dataset", "4.Analysis", "5.SDTM", "6.Validation", "8.Verification"],
}
STUDY_ROLES: Final = set(ROLE_MAP)

ISOLATED_ROLE_MAP: Final = {
    "Randomization Statistician": ["Random"],
    "Blind Reviewer": ["Reviewer"],
    "Unblind Reviewer": ["Random", "Reviewer"],
}
ISOLATED_ROLES: Final = set(ISOLATED_ROLE_MAP)

ISOLATED_STAT_IDMC_BLOCKED_MAP: Final = {
    "Trial STAT/SP": ["8.Verification"],
    "Verification SP": ["4.Analysis", "5.SDTM"],
    "SDTM": ["4.Analysis", "8.Verification"],
    "Manager": [],
}

LEGACY_STUDY_MAP: Final = {
    "Trial STAT/SP": [3, 4, 5],
    "Verification SP": [3, 5, 8],
    "SDTM": [3, 5],
    "Manager": [3, 4, 5, 8],
}

NEW_THRESHOLD: Final = 25069
FORCE_NEW_CODES: Final = {
    "25-038", "25-032", "25-028", "25-006", "25-004",
    "24-084", "24-076", "24-072", "24-061", "24-037", "24-033",
    "24-026", "24-009", "21-040",
}
STAT_IDMC_NEW_THRESHOLD: Final = 26012
STAT_IDMC_FORCE_NEW_CODES: Final = {"25-074", "25-077"}
