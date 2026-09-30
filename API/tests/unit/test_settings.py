import pytest

from folder_grant_api.settings import Settings


@pytest.mark.parametrize("port", [0, 65536])
def test_runtime_safety_rejects_invalid_api_port(port: int) -> None:
    settings = Settings(port=port)

    with pytest.raises(ValueError, match="API port"):
        settings.validate_runtime_safety()


def test_runtime_safety_accepts_valid_api_port() -> None:
    Settings(port=8000).validate_runtime_safety()
