from test_client.models import ManualAccessRequest, new_request_id


def test_request_id_is_unique_and_prefixed() -> None:
    first = new_request_id()
    second = new_request_id()
    assert first.startswith("CLIENT-")
    assert first != second


def test_progress_request_keeps_folder_fields() -> None:
    request = ManualAccessRequest(
        request_id="CLIENT-1",
        operation="grant",
        project_status="progress",
        employee_id="12345",
        project_code="98765",
        level2="Study",
        level3="STAT",
        role="Manager",
        requested_by="tester",
    )
    payload = request.to_payload()
    assert payload["source"] == "test-client"
    assert payload["level3"] == "STAT"


def test_closed_request_removes_folder_specific_fields() -> None:
    request = ManualAccessRequest(
        request_id="CLIENT-2",
        operation="revoke",
        project_status="closed",
        employee_id="12345",
        project_code="98765",
        level2="Study",
        level3="STAT",
        role="Manager",
        requested_by="tester",
    )
    payload = request.to_payload()
    assert payload["level2"] is None
    assert payload["level3"] is None
    assert payload["role"] is None
