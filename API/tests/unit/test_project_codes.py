import pytest

from folder_grant_api.domain.project_codes import (
    InvalidProjectCode,
    canonical_project_code,
    is_new_template,
    project_segment_for_folder,
)


def test_five_digit_project_code_is_normalized() -> None:
    assert canonical_project_code("26012") == "26-012"
    assert project_segment_for_folder("26012") == "26012"


def test_four_digit_legacy_code_inserts_middle_zero() -> None:
    assert project_segment_for_folder("2412") == "24012"


def test_project_suffix_matches_existing_behavior() -> None:
    assert project_segment_for_folder("26012-2") == "26012A2"


def test_new_template_threshold_and_forced_codes() -> None:
    assert is_new_template("25069") is True
    assert is_new_template("25068") is False
    assert is_new_template("25038") is True


def test_project_code_requires_digits() -> None:
    with pytest.raises(InvalidProjectCode):
        canonical_project_code("ABC")
