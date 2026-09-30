import json

from PyQt5.QtCore import QSettings, QThread, QTimer, Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .api_client import ApiResponse, FolderGrantApiClient
from .models import ManualAccessRequest, new_request_id
from .worker import ApiTask


LEVEL3_VALUES = ["ARS", "CO", "DM", "ER", "MW", "PM", "PV", "RA", "SSU", "STAT", "STAT_IDMC", "ETC"]
ROLE_VALUES = [
    "",
    "Trial STAT/SP",
    "Verification SP",
    "SDTM",
    "Manager",
    "Randomization Statistician",
    "Blind Reviewer",
    "Unblind Reviewer",
]
FINAL_STATUSES = {"simulated", "succeeded", "partially_succeeded", "failed", "cancelled"}


class TestClientWindow(QMainWindow):
    COL_SELECT = 0
    COL_REQUEST = 1
    COL_OPERATION = 2
    COL_PROJECT_STATUS = 3
    COL_EMPLOYEE = 4
    COL_PROJECT = 5
    COL_LEVEL2 = 6
    COL_LEVEL3 = 7
    COL_ROLE = 8
    COL_STATUS = 9
    COL_JOB = 10

    def __init__(self):
        super().__init__()
        self.settings = QSettings("FolderGrantTool", "ApiTestClient")
        self.requests: list[ManualAccessRequest] = []
        self.active_jobs: dict[int, str] = {}
        self.poll_in_flight: set[int] = set()
        self.threads: set[QThread] = set()
        self.tasks: set[ApiTask] = set()
        self.setWindowTitle("Folder Grant API Test Client")
        self.resize(1480, 850)
        self._build_ui()
        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(1500)
        self.poll_timer.timeout.connect(self._poll_jobs)
        self.poll_timer.start()

    def _build_ui(self) -> None:
        root = QWidget(self)
        layout = QVBoxLayout(root)
        layout.addWidget(self._connection_group())
        layout.addWidget(self._request_group())

        buttons = QHBoxLayout()
        self.preview_button = QPushButton("선택 Preview")
        self.execute_button = QPushButton("선택 실행")
        self.cancel_button = QPushButton("선택 Job 취소")
        self.delete_button = QPushButton("선택 행 삭제")
        buttons.addWidget(self.preview_button)
        buttons.addWidget(self.execute_button)
        buttons.addWidget(self.cancel_button)
        buttons.addWidget(self.delete_button)
        buttons.addStretch(1)
        layout.addLayout(buttons)

        headers = ["선택", "Request ID", "작업", "과제", "사번", "프로젝트", "Level2", "Level3", "Role", "상태", "Job ID"]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(self.COL_ROLE, QHeaderView.Stretch)
        layout.addWidget(self.table, 3)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(2000)
        layout.addWidget(QLabel("결과 / 로그"))
        layout.addWidget(self.log, 2)
        self.setCentralWidget(root)

        self.preview_button.clicked.connect(self._preview_selected)
        self.execute_button.clicked.connect(self._execute_selected)
        self.cancel_button.clicked.connect(self._cancel_selected)
        self.delete_button.clicked.connect(self._delete_selected)

    def _connection_group(self) -> QGroupBox:
        group = QGroupBox("API 연결")
        layout = QHBoxLayout(group)
        self.url_edit = QLineEdit(self.settings.value("base_url", "http://192.168.1.10:8000"))
        self.key_edit = QLineEdit()
        self.key_edit.setEchoMode(QLineEdit.Password)
        self.key_edit.setPlaceholderText("X-API-Key (저장하지 않음)")
        self.show_key = QCheckBox("Key 보기")
        self.connection_button = QPushButton("연결 확인")
        layout.addWidget(QLabel("URL"))
        layout.addWidget(self.url_edit, 2)
        layout.addWidget(QLabel("API Key"))
        layout.addWidget(self.key_edit, 2)
        layout.addWidget(self.show_key)
        layout.addWidget(self.connection_button)
        self.show_key.toggled.connect(
            lambda checked: self.key_edit.setEchoMode(QLineEdit.Normal if checked else QLineEdit.Password)
        )
        self.connection_button.clicked.connect(self._check_connection)
        return group

    def _request_group(self) -> QGroupBox:
        group = QGroupBox("수동 요청 추가")
        layout = QGridLayout(group)
        self.employee_edit = QLineEdit()
        self.project_edit = QLineEdit()
        self.requester_edit = QLineEdit()
        self.operation_combo = QComboBox()
        self.operation_combo.addItem("권한부여", "grant")
        self.operation_combo.addItem("권한해제", "revoke")
        self.project_status_combo = QComboBox()
        self.project_status_combo.addItem("진행", "progress")
        self.project_status_combo.addItem("종료", "closed")
        self.level2_combo = QComboBox()
        self.level2_combo.addItems(["Study", "Isolated"])
        self.level3_combo = QComboBox()
        self.level3_combo.addItems(LEVEL3_VALUES)
        self.role_combo = QComboBox()
        self.role_combo.addItems(ROLE_VALUES)
        self.add_button = QPushButton("목록에 추가")

        fields = [
            ("대상자 사번", self.employee_edit),
            ("프로젝트 코드", self.project_edit),
            ("요청자", self.requester_edit),
            ("작업", self.operation_combo),
            ("과제 상태", self.project_status_combo),
            ("Level2", self.level2_combo),
            ("Level3", self.level3_combo),
            ("Role", self.role_combo),
        ]
        for index, (label, widget) in enumerate(fields):
            row, column = divmod(index, 4)
            layout.addWidget(QLabel(label), row * 2, column)
            layout.addWidget(widget, row * 2 + 1, column)
        layout.addWidget(self.add_button, 4, 3)
        self.add_button.clicked.connect(self._add_request)
        self.project_status_combo.currentIndexChanged.connect(self._update_folder_fields)
        return group

    def _client(self) -> FolderGrantApiClient:
        base_url = self.url_edit.text().strip()
        self.settings.setValue("base_url", base_url)
        return FolderGrantApiClient(base_url, self.key_edit.text())

    def _add_request(self) -> None:
        employee = self.employee_edit.text().strip()
        project = self.project_edit.text().strip()
        requester = self.requester_edit.text().strip()
        if not employee or not project or not requester:
            QMessageBox.warning(self, "입력 확인", "대상자 사번, 프로젝트 코드, 요청자를 모두 입력하세요.")
            return
        closed = self.project_status_combo.currentData() == "closed"
        request = ManualAccessRequest(
            request_id=new_request_id(),
            operation=self.operation_combo.currentData(),
            project_status=self.project_status_combo.currentData(),
            employee_id=employee,
            project_code=project,
            level2=None if closed else self.level2_combo.currentText(),
            level3=None if closed else self.level3_combo.currentText(),
            role=None if closed else (self.role_combo.currentText() or None),
            requested_by=requester,
        )
        self.requests.append(request)
        self._append_row(request)

    def _append_row(self, request: ManualAccessRequest) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        checkbox = QCheckBox()
        checkbox.setChecked(True)
        holder = QWidget()
        holder_layout = QHBoxLayout(holder)
        holder_layout.setContentsMargins(0, 0, 0, 0)
        holder_layout.setAlignment(Qt.AlignCenter)
        holder_layout.addWidget(checkbox)
        self.table.setCellWidget(row, self.COL_SELECT, holder)
        values = [
            request.request_id,
            request.operation,
            request.project_status,
            request.employee_id,
            request.project_code,
            request.level2 or "",
            request.level3 or "",
            request.role or "",
            "대기",
            "",
        ]
        for column, value in enumerate(values, start=1):
            self.table.setItem(row, column, QTableWidgetItem(str(value)))

    def _selected_rows(self) -> list[int]:
        rows = []
        for row in range(self.table.rowCount()):
            holder = self.table.cellWidget(row, self.COL_SELECT)
            checkbox = holder.findChild(QCheckBox) if holder else None
            if checkbox and checkbox.isChecked():
                rows.append(row)
        return rows

    def _check_connection(self) -> None:
        client = self._client()
        self._run_task(
            client.ready,
            lambda response: self._log_json("연결 성공", response.data),
        )

    def _preview_selected(self) -> None:
        rows = self._selected_rows()
        if len(rows) != 1:
            QMessageBox.information(self, "Preview", "Preview할 행 하나만 선택하세요.")
            return
        row = rows[0]
        client = self._client()
        payload = self.requests[row].to_payload()
        self._set_status(row, "Preview 중")
        self._run_task(
            lambda: client.preview(payload),
            lambda response, target_row=row: self._preview_finished(target_row, response),
            lambda message, target_row=row: self._request_failed(target_row, message),
        )

    def _execute_selected(self) -> None:
        rows = self._selected_rows()
        if not rows:
            QMessageBox.information(self, "실행", "실행할 행을 선택하세요.")
            return
        answer = QMessageBox.question(
            self,
            "API 작업 실행",
            f"선택한 {len(rows)}건을 API로 전송합니다.\n서버가 PowerShell 모드이면 실제 권한이 변경됩니다. 계속하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return
        client = self._client()
        for row in rows:
            if row in self.active_jobs:
                continue
            self._set_status(row, "접수 중")
            payload = self.requests[row].to_payload()
            self._run_task(
                lambda target_payload=payload: client.register(target_payload),
                lambda response, target_row=row: self._registered(target_row, response),
                lambda message, target_row=row: self._request_failed(target_row, message),
            )

    def _cancel_selected(self) -> None:
        client = self._client()
        for row in self._selected_rows():
            job_id = self.active_jobs.get(row) or self.table.item(row, self.COL_JOB).text().strip()
            if not job_id:
                continue
            requested_by = self.requests[row].requested_by
            self._run_task(
                lambda target_job=job_id, actor=requested_by: client.cancel(target_job, actor),
                lambda response, target_row=row: self._job_updated(target_row, response),
                lambda message, target_row=row: self._request_failed(target_row, message),
            )

    def _delete_selected(self) -> None:
        if self.threads:
            QMessageBox.warning(self, "삭제 불가", "API 요청 또는 상태 조회가 끝난 후 삭제하세요.")
            return
        for row in reversed(self._selected_rows()):
            if row in self.active_jobs:
                QMessageBox.warning(self, "삭제 불가", "처리 중인 행은 삭제할 수 없습니다.")
                continue
            self.table.removeRow(row)
            self.requests.pop(row)
            self._reindex_active_jobs_after_delete(row)

    def _preview_finished(self, row: int, response: ApiResponse) -> None:
        self._set_status(row, "Preview 완료")
        self._log_json(f"Preview {self.requests[row].request_id}", response.data)

    def _registered(self, row: int, response: ApiResponse) -> None:
        data = response.data
        job_id = data["job_id"]
        self.active_jobs[row] = job_id
        self.table.setItem(row, self.COL_JOB, QTableWidgetItem(job_id))
        self._set_status(row, data["status"])
        self._log_json(f"Job 접수 {self.requests[row].request_id}", data)

    def _poll_jobs(self) -> None:
        client = self._client()
        for row, job_id in list(self.active_jobs.items()):
            if row in self.poll_in_flight or row >= self.table.rowCount():
                continue
            self.poll_in_flight.add(row)
            self._run_task(
                lambda target_job=job_id: client.get_job(target_job),
                lambda response, target_row=row: self._job_updated(target_row, response),
                lambda message, target_row=row: self._poll_failed(target_row, message),
                lambda target_row=row: self.poll_in_flight.discard(target_row),
            )

    def _job_updated(self, row: int, response: ApiResponse) -> None:
        if row >= self.table.rowCount():
            return
        data = response.data
        status = data["status"]
        self._set_status(row, status)
        if status in FINAL_STATUSES:
            self.active_jobs.pop(row, None)
            self._log_json(f"Job 완료 {self.requests[row].request_id}", data)

    def _request_failed(self, row: int, message: str) -> None:
        self._set_status(row, "실패")
        self.active_jobs.pop(row, None)
        self.log.appendPlainText(f"[오류] {message}")

    def _poll_failed(self, row: int, message: str) -> None:
        self.log.appendPlainText(f"[상태 조회 오류] {message}")

    def _set_status(self, row: int, status: str) -> None:
        item = self.table.item(row, self.COL_STATUS) or QTableWidgetItem()
        item.setText(status)
        colors = {
            "succeeded": QColor("#c8e6c9"),
            "simulated": QColor("#bbdefb"),
            "failed": QColor("#ffcdd2"),
            "partially_succeeded": QColor("#ffe0b2"),
            "cancelled": QColor("#eeeeee"),
        }
        item.setBackground(colors.get(status, QColor("white")))
        self.table.setItem(row, self.COL_STATUS, item)

    def _log_json(self, title: str, data: dict) -> None:
        self.log.appendPlainText(f"\n[{title}]\n{json.dumps(data, ensure_ascii=False, indent=2)}")

    def _run_task(self, operation, success, failure=None, done=None) -> None:
        thread = QThread(self)
        task = ApiTask(operation)
        task.moveToThread(thread)
        thread.started.connect(task.run)
        task.succeeded.connect(success)
        task.failed.connect(failure or (lambda message: self.log.appendPlainText(f"[오류] {message}")))
        task.finished.connect(thread.quit)
        if done:
            task.finished.connect(done)
        task.finished.connect(task.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(lambda: self.threads.discard(thread))
        task.finished.connect(lambda: self.tasks.discard(task))
        self.threads.add(thread)
        self.tasks.add(task)
        thread.start()

    def _update_folder_fields(self) -> None:
        enabled = self.project_status_combo.currentData() == "progress"
        self.level2_combo.setEnabled(enabled)
        self.level3_combo.setEnabled(enabled)
        self.role_combo.setEnabled(enabled)

    def _reindex_active_jobs_after_delete(self, deleted_row: int) -> None:
        self.active_jobs = {
            (row - 1 if row > deleted_row else row): job_id
            for row, job_id in self.active_jobs.items()
            if row != deleted_row
        }
