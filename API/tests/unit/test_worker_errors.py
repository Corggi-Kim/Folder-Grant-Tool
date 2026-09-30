from folder_grant_api.workers.powershell_worker import _specific_error_code


def test_specific_project_verification_error_is_preserved_from_clixml() -> None:
    message = '#< CLIXML><S S="Error">Exception: ROOT_ACL_NOT_FOUND</S>'
    assert _specific_error_code(message, "POWERSHELL_FAILED") == "ROOT_ACL_NOT_FOUND"


def test_unknown_powershell_error_uses_fallback() -> None:
    assert _specific_error_code("unexpected", "POWERSHELL_FAILED") == "POWERSHELL_FAILED"
