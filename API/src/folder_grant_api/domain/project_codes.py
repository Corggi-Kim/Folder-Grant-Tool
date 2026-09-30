import re

from .constants import (
    FORCE_NEW_CODES,
    NEW_THRESHOLD,
    STAT_IDMC_FORCE_NEW_CODES,
    STAT_IDMC_NEW_THRESHOLD,
)


class InvalidProjectCode(ValueError):
    pass


def insert_zero_middle_4digit(code: str) -> str:
    return code[:2] + "0" + code[2:]


def split_project_and_suffix(raw: str) -> tuple[str, str]:
    value = (raw or "").strip()
    if "-" in value:
        base, suffix = value.split("-", 1)
        return base.strip(), suffix.strip()
    return value, ""


def _project_digits(raw: str) -> tuple[str, str]:
    base, suffix = split_project_and_suffix(raw)
    digits = re.sub(r"\D", "", base)
    if not digits:
        raise InvalidProjectCode("프로젝트 코드에 숫자가 없습니다.")
    if len(digits) == 4:
        digits = insert_zero_middle_4digit(digits)
    return digits, suffix


def project_segment_for_folder(raw: str) -> str:
    digits, suffix = _project_digits(raw)
    return digits + (f"A{suffix}" if suffix else "")


def group_digits_and_suffix(raw: str) -> tuple[str, str]:
    digits, suffix = _project_digits(raw)
    last_five = digits[-5:] if len(digits) >= 5 else digits.zfill(5)
    return last_five, suffix


def canonical_project_code(raw: str) -> str:
    last_five, _ = group_digits_and_suffix(raw)
    return f"{last_five[:2]}-{last_five[2:]}"


def is_new_template(raw: str) -> bool:
    last_five, _ = group_digits_and_suffix(raw)
    return canonical_project_code(raw) in FORCE_NEW_CODES or int(last_five) >= NEW_THRESHOLD


def is_stat_idmc_new_policy(raw: str) -> bool:
    last_five, _ = group_digits_and_suffix(raw)
    return (
        canonical_project_code(raw) in STAT_IDMC_FORCE_NEW_CODES
        or int(last_five) >= STAT_IDMC_NEW_THRESHOLD
    )
