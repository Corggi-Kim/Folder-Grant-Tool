from dataclasses import asdict, dataclass
from datetime import datetime
from uuid import uuid4


def new_request_id(prefix: str = "CLIENT") -> str:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"{prefix}-{stamp}-{uuid4().hex[:6].upper()}"


@dataclass(frozen=True)
class ManualAccessRequest:
    request_id: str
    operation: str
    project_status: str
    employee_id: str
    project_code: str
    level2: str | None
    level3: str | None
    role: str | None
    requested_by: str
    source: str = "test-client"

    def to_payload(self) -> dict:
        payload = asdict(self)
        if self.project_status == "closed":
            payload["level2"] = None
            payload["level3"] = None
            payload["role"] = None
        return payload
