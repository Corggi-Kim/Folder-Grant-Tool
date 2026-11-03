import ntpath
import re

from .project_codes import insert_zero_middle_4digit, project_segment_for_folder, split_project_and_suffix
from .constants import LEVEL3_CHOICES


def normalize_level2(value: str) -> str:
    return "Isolated" if "iso" in (value or "").strip().lower() else "Study"


def normalize_level3(value: str) -> str:
    value = (value or '').strip()
    if not value:
        raise ValueError('Level3 폴더명을 입력하세요.')
    if (len(value)>50 or value in {'.','..'} or value.endswith('.')
            or any(c in r'<>:"/\|?*[]' or ord(c)<32 or ord(c)==127 for c in value)):
        raise ValueError('Level3는 경로 구분자·와일드카드가 없는 한 개의 폴더명으로 입력하세요.')
    reserved = {'CON','PRN','AUX','NUL','CONIN$','CONOUT$','CLOCK$'} | {
        prefix+number for prefix in ('COM','LPT') for number in '123456789¹²³'}
    if value.split('.',1)[0].rstrip().upper() in reserved:
        raise ValueError('Level3에 사용할 수 없는 예약된 폴더명입니다.')
    # The fixed list supplies recommendations and canonical spellings, not policy.
    return value.upper() if value.upper() in LEVEL3_CHOICES else value


def build_active_path(project_code: str, level2: str, level3: str, share_root: str) -> str:
    return share_root.format(
        proj_seg=project_segment_for_folder(project_code),
        lv2=normalize_level2(level2),
        lv3=normalize_level3(level3),
    )


def closed_segment(project_code: str) -> str:
    value = (project_code or "").strip()
    if re.fullmatch(r"\d{5}(?:A\d+)?", value):
        return value

    base, suffix = split_project_and_suffix(value)
    digits = re.sub(r"\D", "", base)
    if len(digits) == 4:
        digits = insert_zero_middle_4digit(digits)
    segment = digits[-5:].zfill(5)
    suffix_digits = re.sub(r"\D", "", suffix)
    return segment + (f"A{suffix_digits}" if suffix_digits else "")


def build_closed_candidates(project_code: str, closed_root: str, archive_root: str) -> list[str]:
    segment = closed_segment(project_code)
    return [ntpath.join(closed_root, segment), ntpath.join(archive_root, segment)]


def format_group_name(project_code: str, level2: str) -> str:
    from .project_codes import group_digits_and_suffix

    last_five, suffix = group_digits_and_suffix(project_code)
    name = f"LSK {last_five[:2]}-{last_five[2:]}"
    if suffix:
        name += f"-{suffix}"
    if normalize_level2(level2) == "Isolated":
        name += " Isolated"
    return name
