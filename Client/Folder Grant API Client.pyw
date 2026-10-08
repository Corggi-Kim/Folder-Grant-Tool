# =====================================
# 1) 폴더 구조
#    project/
#      ├─ Folder Grant API Client.pyw
#      └─ assets/
#            ├─ logo.png
#            └─ fgt.ico
#
# 2) 실행 시 생성/사용 경로 (코드에서 자동 생성)
#    - C:\FGT\Log       : 실행 로그(access_YYYYMMDD.log)
#    - C:\FGT\conf      : 설정(login.json, theme 등)
#    - C:\FGT\ef         : 엑셀파일 다운로드
#    - C:\FGT\debug   : 디버그 파일
#
# 3) EXE 압축 시 모듈 임포트 필요 (--collect-all 옵션)
#    - selenium
# =====================================

import sys, os, glob, datetime, re, shutil, json, time
import html as htmllib
from typing import Dict, List
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, StaleElementReferenceException
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QFileDialog, QTableWidget,
    QTableWidgetItem, QAbstractItemView, QVBoxLayout, QWidget, QTextEdit, QHBoxLayout,
    QLabel, QHeaderView, QCheckBox, QMenu, QToolButton, QMessageBox, QStyleOptionButton,
    QStyle, QDialog, QLineEdit, QComboBox, QDialogButtonBox, QFormLayout, QStatusBar,
    QSpinBox, QProgressBar, QAbstractButton
)
from PyQt5.QtGui import QFont, QGuiApplication, QKeySequence, QPainter, QTextOption, QTextCursor, QPixmap, QIcon
from PyQt5.QtCore import Qt, QRect, pyqtSignal, QThread, QObject, pyqtSlot, QTimer, QPoint, QSize
from openpyxl import load_workbook

APP_NAME = "Folder Grant API Client"
APP_VERSION = "1.0.0"  #구조변경, 기능추가/수정, 오류/버그수정
APP_BUILD = "2026-10-08"
APP_VERSION_STR = f"v{APP_VERSION}"

THEMES = {
    "light": {
        "bg": "#f5f5f5",
        "btn": "#d6dde0",
        "hover": "#b0bec5",
        "press": "#78909c",
        "panel": "#eceff1",
        "panel_border": "#cfd8dc",
        "select": "#d6e1e7",
        "alt": "#f6f9fb",
    },
    "dark":  {
        "bg": "#2f2f2f",
        "btn": "#424242",
        "hover": "#616161",
        "press": "#757575",
        "panel": "#3a3a3a",
        "panel_border": "#505050",
        "select": "#4a5a66",
        "alt": "#343838",
    },
}

HELP_TEXT = r"""[Folder Grant API Client 사용 안내]

1) API 설정: 서버 주소와 X-API-Key를 입력하고 연결 확인 후 저장합니다.
   설정 파일은 C:\FGT\conf\api_client.json이며 API Key가 평문 저장됩니다.
2) 설정: 기존 BUS 로그인 계정을 입력합니다.
3) 요청 확인: BUS 진행·종료 권한 요청을 조회합니다.
4) 신규 확인: BUS 프로젝트 생성 요청을 조회합니다. 수동 추가도 가능합니다.
5) 파일 선택 또는 수동 입력: API 작업용 요청을 추가합니다. BUS 완료 대상은 아닙니다.
6) 실행: 서버 API에 권한 작업을 등록하고 최종 Job 결과를 표시합니다.
   클라이언트에서는 AD·폴더 권한 명령을 실행하지 않습니다.
7) Dry Run: 서버 Preview만 호출합니다. 실제 Job과 BUS 완료 처리를 실행하지 않습니다.
8) 완료 처리: BUS 요청이며 서버 PowerShell 작업이 실제 성공한 경우에만 허용합니다.
   Mock·Preview·부분 성공·실패·취소는 BUS 완료 대상에서 제외됩니다.
   서버 성공 후 BUS 완료가 실패했으면 완료 처리만 다시 실행하세요.
9) 중지: 새 작업 등록을 멈추고 실행 Job에 취소를 요청합니다.
   이미 실행 중인 서버 PowerShell 단계는 즉시 종료되지 않을 수 있습니다.
10) BUS 조회 행을 편집하면 수동 요청으로 전환합니다. 기존 성공 결과는 무효화됩니다.
11) 로그: C:\FGT\Log\access_YYYYMMDD.log, 다운로드: C:\FGT\ef
12) 테마는 상단 버튼으로 전환합니다. 상세 실행·검증 절차는 Client/README.md를 확인하세요.
"""


LOG_DIR = r"C:\FGT\Log"
CONF_DIR = r"C:\FGT\conf"
DL_DIR = r"C:\FGT\ef"
DEBUG_DIR = r"C:\FGT\debug"

GROUP_OU_PATH = r"OU=Group Project Folder,OU=0.Management Object Group,OU=lskglobal,DC=lskglobal,DC=com"
TEMPLATE_ROOT = r"\\LSK_S010\Study folder\_Template"

CONF_FILE = os.path.join(CONF_DIR, "login.json")
LOGO_FILE = "logo.png"
ICON_FILE = "fgt.ico"



DOMAIN_EMAIL_SUFFIX = "lskglobal.com"

BUS_LOGIN_URL = "https://bus.lskglobal.com/L4/Common/Login.aspx"
BUS_PROGRESS_LIST_URL = "https://bus.lskglobal.com/L4/Common/Default.aspx?7VHoVKC6bjriQDXQa/t/dQ=="
BUS_END_LIST_URL = "https://bus.lskglobal.com/L4/Common/Default.aspx?oRrCZc631pq4qaUZnht8Cg=="
BUS_NEW_URL = "https://bus.lskglobal.com/L4/Common/Default.aspx?2QiSDYhMfx5ql2mmNbos4A=="

REQ_GRANT = "권한부여"
REQ_RELEASE = "권한해제"
END_FLAG_GRANT = "END_GRANT"
END_FLAG_RELEASE = "END_RELEASE"

RELEASE_HINT_HEADERS = {
    "해제요청일", "해제요청자사번", "해제요청자성명",
    "해제요청자부서/팀", "해제요청자직책"
}

SHARE_ROOT = r"\\LSK_S010\Study Folder\{proj_seg}\{lv2}\{lv3}"
CLOSED_ROOT = r"\\192.168.1.95\Study_Closed"
CLOSED_ARCHIVE_ROOT = r"\\192.168.1.95\Study_Archive"

LEVEL3_CHOICES = ["ARS","CO","DM","ER","MW","PM","PV","RA","SSU","STAT","STAT_IDMC","ETC"]

ROLE_MAP: Dict[str, List[str]] = {
    "Trial STAT/SP": ["3.Dataset", "4.Analysis", "5.SDTM", "6.Validation"],
    "Verification SP": ["3.Dataset", "6.Validation", "8.Verification"],
    "SDTM": ["3.Dataset", "5.SDTM", "6.Validation"],
    "Manager": ["3.Dataset", "4.Analysis", "5.SDTM", "6.Validation", "8.Verification"],
    "Randomization Statistician": ["Random"],
    "Blind Reviewer": ["Reviewer"],
    "Unblind Reviewer": ["Random", "Reviewer"],
}

STUDY_ROLES = {"Trial STAT/SP", "Verification SP", "SDTM", "Manager"}
ISOLATED_ROLES = {"Randomization Statistician", "Blind Reviewer", "Unblind Reviewer"}

ISOLATED_STAT_IDMC_ROLE_MAP: Dict[str, List[str]] = {
    "Trial STAT/SP": ["8.Verification"],
    "Verification SP": ["4.Analysis", "5.SDTM"],
    "SDTM": ["4.Analysis", "8.Verification"],
    "Manager": [],
}

LEGACY_STUDY_MAP = {
    "Trial STAT/SP": [3, 4, 5],
    "Verification SP": [3, 5, 8],
    "SDTM": [3, 5],
    "Manager": [3, 4, 5, 8],
}

NEW_THRESHOLD = 25069
FORCE_NEW_CODES = {
    "25-038", "25-032", "25-028", "25-006", "25-004",
    "24-084", "24-076", "24-072", "24-061", "24-037", "24-033", "24-026", "24-009",
    "21-040",
}
STAT_IDMC_NEW_THRESHOLD = 26012
STAT_IDMC_FORCE_NEW_CODES = {"25-074", "25-077"}

HEADER_ALIASES = {
    "user":   {"대상자사번"},
    "name": {"대상자성명"},
    "proj":   {"프로젝트코드"},
    "level1": {"폴더level1","폴더Level1"},
    "level2": {"폴더level2","폴더Level2"},
    "level3": {"폴더level3","폴더Level3"},
    "role":   {"statrole","STATROLE"},
    "dept":   {"대상자부서/팀"},
}

def build_root_from_proj(proj_raw: str) -> str:
    seg = proj_segment_for_folder(proj_raw)
    return r"\\LSK_S010\Study Folder\{}".format(seg)

def resource_path(rel_path: str) -> str:
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel_path)



def make_hidden_chrome_options(download_dir: str | None = None):
    options = webdriver.ChromeOptions()
    if download_dir:
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)

    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1280,900")
    options.add_argument("--window-position=-32000,-32000")
    return options

def _norm_k(h: str) -> str:
    return (h or "").strip().lower().replace(" ", "").replace("/", "").replace("_","")
RELEASE_HINT_HEADERS_NORM = {_norm_k(h) for h in RELEASE_HINT_HEADERS}

def _find_cols(header_row, wanted_norm_set):
    idxs = {}
    norm = [_norm_k(h) for h in header_row]
    for i, h in enumerate(norm):
        if h in wanted_norm_set and h not in idxs:
            idxs[h] = i
    return idxs

def _is_release_row_by_values(row_values, header_row):
    try:
        idxs = _find_cols(header_row, RELEASE_HINT_HEADERS_NORM)
        for k, i in idxs.items():
            if i < len(row_values):
                v = row_values[i]
                if v is not None and str(v).strip() != "":
                    return True
    except Exception:
        pass
    return False

def _norm_header(h: str) -> str:
    return (h or "").strip().lower().replace(" ", "").replace("/", "").replace("_","").replace("-","")

def auto_map_columns(header_row):
    idx = {}
    norm_headers = [_norm_header(h if h is not None else "") for h in header_row]
    for key, aliases in HEADER_ALIASES.items():
        aliases_norm = {_norm_header(a) for a in aliases}
        for i, h in enumerate(norm_headers):
            if h in aliases_norm:
                idx[key] = i
                break
    return idx

def normalize_lv2(s: str) -> str:
    t = (s or "").strip().lower()
    return "Isolated" if "iso" in t else "Study"

def is_stat_idmc_lv3(lv3: str) -> bool:
    return (lv3 or "").strip().upper() == "STAT_IDMC"

def is_stat_lv3(lv3: str) -> bool:
    return (lv3 or "").strip().upper() in {"STAT", "STAT_IDMC"}

def is_lv3_etc(lv3: str) -> bool:
    return (lv3 or "").strip().lower() == "etc"

def insert_zero_middle_4digit(code4: str) -> str:
    return code4[:2] + "0" + code4[2:]

def split_proj_and_suffix(raw: str):
    s = (raw or "").strip()
    if "-" in s:
        base, suf = s.split("-", 1)
        return base.strip(), suf.strip()
    return s, ""

def proj_segment_for_folder(raw: str) -> str:
    base, suf = split_proj_and_suffix(raw)
    digits = re.sub(r"\D", "", base or "")
    if len(digits) == 4:
        digits = insert_zero_middle_4digit(digits)
    return digits + ("A" + suf if suf else "")

def group_digits_from_proj_for_groupname(raw: str) -> (str, str):
    base, suf = split_proj_and_suffix(raw)
    digits = re.sub(r"\D", "", base or "")
    if len(digits) == 4:
        digits = insert_zero_middle_4digit(digits)
    last5 = digits[-5:] if len(digits) >= 5 else digits.zfill(5)
    return last5, suf

def format_group_name(proj_raw: str, lv2: str) -> str:
    last5, suf = group_digits_from_proj_for_groupname(proj_raw)
    yy, xxx = last5[:2], last5[2:]
    name = f"LSK {yy}-{xxx}"
    if suf:
        name += f"-{suf}"
    if normalize_lv2(lv2) == "Isolated":
        name += " Isolated"
    return name

def _yyxxx_from_proj_for_groupname(proj_raw: str) -> str:
    last5, _ = group_digits_from_proj_for_groupname(proj_raw)
    return last5

def _lsk_code_for_compare(proj_raw: str) -> str:
    last5 = _yyxxx_from_proj_for_groupname(proj_raw)
    return f"{last5[:2]}-{last5[2:]}"

def is_new_template(proj_raw: str) -> bool:
    code_yyxxx = int(_yyxxx_from_proj_for_groupname(proj_raw))
    lsk_code = _lsk_code_for_compare(proj_raw)
    if lsk_code in FORCE_NEW_CODES:
        return True
    return code_yyxxx >= NEW_THRESHOLD

def is_stat_idmc_new_policy(proj_raw: str) -> bool:
    lsk_code = _lsk_code_for_compare(proj_raw)
    if lsk_code in STAT_IDMC_FORCE_NEW_CODES:
        return True
    try:
        code_yyxxx = int(_yyxxx_from_proj_for_groupname(proj_raw))
    except Exception:
        return False
    return code_yyxxx >= STAT_IDMC_NEW_THRESHOLD

def build_path_l3(proj_raw: str, lv2: str, lv3: str) -> str:
    seg = proj_segment_for_folder(proj_raw)
    return SHARE_ROOT.format(
        proj_seg=seg,
        lv2=normalize_lv2(lv2),
        lv3=(lv3 or "").strip().replace("\\","").replace("/","")
    )

def closed_segment_from_proj(proj_raw: str) -> str:
    s = (proj_raw or "").strip()
    m = re.fullmatch(r"(\d{5})(?:A(\d+))?$", s)
    if m:
        return s

    base, suf = split_proj_and_suffix(s)
    digits = re.sub(r"\D", "", base or "")
    if len(digits) == 4:
        digits = insert_zero_middle_4digit(digits)
    seg = digits[-5:].zfill(5)

    suf_digits = re.sub(r"\D", "", suf or "")
    return seg + (("A" + suf_digits) if suf_digits else "")

def build_closed_path_from_proj(proj_raw: str) -> str:
    return build_closed_candidate_paths_from_proj(proj_raw)[0]

def build_closed_candidate_paths_from_proj(proj_raw: str) -> List[str]:
    seg = closed_segment_from_proj(proj_raw)
    return [
        os.path.join(CLOSED_ROOT, seg),
        os.path.join(CLOSED_ARCHIVE_ROOT, seg),
    ]


def _extract_tables_from_html(s: str):
    tables = re.findall(r"<table[^>]*>(.*?)</table>", s, re.I | re.S)
    out = []
    for tbl in tables:
        rows = re.findall(r"<tr[^>]*>(.*?)</tr>", tbl, re.I | re.S)
        parsed = []
        for r in rows:
            cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", r, re.I | re.S)
            vals = []
            for c in cells:
                t = re.sub(r"<[^>]+>", "", c)
                t = htmllib.unescape(t).strip()
                vals.append(t)
            if vals:
                parsed.append(vals)
        if parsed:
            out.append(parsed)
    return out

def _parse_html_best_table(path: str):
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        html = f.read()
    candidates = _extract_tables_from_html(html)
    if not candidates:
        return [], []
    for tbl in candidates:
        header = tbl[0]
        data = tbl[1:]
        return header, data
    return [], []

# API contract and transport. No server package dependency.
import hashlib
import threading
from uuid import uuid4
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit
from urllib.request import Request, build_opener, HTTPRedirectHandler

DEFAULT_API_URL = 'http://192.168.1.10:8000'
API_CONFIG_FILE = os.path.join(CONF_DIR, 'api_client.json')
FINAL_JOB_STATUSES = {'simulated', 'succeeded', 'partially_succeeded', 'failed', 'cancelled'}


def canonical_project_code(raw: str) -> str:
    value = str(raw).strip().upper()
    # BUS displays YY-NNN; the API accepts YYNNN and optional A suffixes.
    value = re.sub(r'^(\d{2})-(\d{3})(?=A\d+$|$)', r'\1\2', value)
    if not re.fullmatch(r'(?:\d{4}|\d{5}(?:A\d+|-\d+)?)', value):
        raise ValueError('프로젝트 코드 형식을 확인하세요: ' + value)
    return value


def make_access_payload(values: dict, request_id: str, source: str, requested_by: str) -> dict:
    closed = values.get('kind') == '종료'
    user = str(values.get('user', '')).strip()
    if not re.fullmatch(r'[A-Za-z0-9._-]+', user):
        raise ValueError('대상자 사번을 확인하세요.')
    operation = values.get('req', REQ_GRANT)
    if operation not in {REQ_GRANT, REQ_RELEASE}:
        raise ValueError('권한부여 또는 권한해제 요청이어야 합니다.')
    lv2 = normalize_lv2(str(values.get('lv2', '')).strip())
    lv3 = str(values.get('lv3', '')).strip().upper()
    if not closed and (lv2 not in {'Study', 'Isolated'} or lv3 not in LEVEL3_CHOICES):
        raise ValueError('진행 과제의 Level2·Level3를 확인하세요.')
    return dict(request_id=request_id, source=source,
                operation='revoke' if operation == REQ_RELEASE else 'grant',
                project_status='closed' if closed else 'progress', employee_id=user,
                project_code=canonical_project_code(values.get('proj', '')),
                level2=None if closed else lv2, level3=None if closed else lv3,
                role=None if closed else (str(values.get('role', '')).strip() or None),
                requested_by=requested_by)


def make_project_payload(project_code: str, project_name: str, request_id: str,
                         source: str, requested_by: str) -> dict:
    return dict(request_id=request_id, source=source,
                project_code=canonical_project_code(project_code),
                project_name=project_name.strip(), requested_by=requested_by)


def can_complete_bus(result: dict, source: str, stopped: bool = False) -> bool:
    return bool(not stopped and source == 'bus' and result.get('status') == 'succeeded'
                and result.get('executor_mode') == 'powershell')


def make_bus_request_id(values: dict, bus_identity: str = '') -> str:
    stable = {k: str(v or '').strip() for k, v in values.items()
              if k in {'kind', 'req', 'user', 'proj', 'lv2', 'lv3', 'role', 'project_name'}}
    if 'proj' in stable:
        stable['proj'] = canonical_project_code(stable['proj'])
    stable['bus_identity'] = bus_identity.strip()
    digest = hashlib.sha256(json.dumps(stable, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
    return 'BUS-' + digest


def new_client_request_id() -> str:
    return 'CLIENT-' + uuid4().hex


def bus_row_identity(header, values) -> str:
    fields = {str(h).strip(): str(v or '').strip() for h,v in zip(header,values)
              if str(h).strip().lower() not in {'','순번','번호','no','no.','#','선택'}}
    for name,value in fields.items():
        if re.sub(r'[\s_-]+','',name).lower() in {'요청번호','신청번호','requestid','요청id','신청id'} and value:
            return name + ':' + value
    # Keep request dates and other original matching fields when no ID is exported.
    return json.dumps(fields,ensure_ascii=False,sort_keys=True)


class ApiClientError(RuntimeError):
    def __init__(self, message, status_code=None):
        super().__init__(message)
        self.status_code = status_code


class _NoApiRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class FolderGrantApiClient:
    def __init__(self, base_url: str, api_key: str, timeout_seconds: int = 30):
        self.base_url = base_url.strip().rstrip('/')
        parsed = urlsplit(self.base_url)
        if parsed.scheme not in {'http', 'https'} or not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment:
            raise ValueError('API 주소는 http://서버:포트 또는 https://서버 형식이어야 합니다.')
        self.api_key = api_key.strip()
        self.timeout_seconds = timeout_seconds
        self.opener = build_opener(_NoApiRedirects())

    def request(self, method: str, path: str, payload: dict | None = None) -> dict:
        authenticated = path != '/health'
        if authenticated and not self.api_key:
            raise ApiClientError('API Key를 입력하세요.')
        headers = {'Accept': 'application/json'}
        if authenticated:
            headers['X-API-Key'] = self.api_key
        body = None
        if payload is not None:
            body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
            headers['Content-Type'] = 'application/json; charset=utf-8'
        try:
            with self.opener.open(Request(self.base_url + path, data=body, headers=headers,
                                         method=method), timeout=self.timeout_seconds) as response:
                data = json.loads(response.read().decode('utf-8'))
                if not isinstance(data, dict):
                    raise ValueError('JSON object required')
                return data
        except HTTPError as exc:
            raw = exc.read().decode('utf-8', errors='replace')
            try:
                detail = json.loads(raw).get('detail', raw)
                if isinstance(detail, dict):
                    detail = detail.get('message') or detail.get('code') or detail
                elif isinstance(detail, list):
                    detail = '; '.join(str(x.get('msg', x)) if isinstance(x, dict) else str(x) for x in detail)
            except (ValueError, AttributeError):
                detail = raw or exc.reason
            message = f'HTTP {exc.code}: {detail}'
            if self.api_key:
                message = message.replace(self.api_key, '[API Key]')
            raise ApiClientError(message, exc.code) from exc
        except (URLError, TimeoutError, OSError) as exc:
            raise ApiClientError('API 연결 실패 또는 시간 초과. 서버 주소와 연결을 확인하세요.') from exc
        except (ValueError, UnicodeError) as exc:
            raise ApiClientError('API 응답이 올바른 JSON 객체가 아닙니다.') from exc

    def health(self):
        return self.request('GET', '/health')

    def ready(self):
        return self.request('GET', '/ready')


class ApiJobWorker(QObject):
    progress = pyqtSignal(dict)
    succeeded = pyqtSignal(dict)
    failed = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, client: FolderGrantApiClient, kind: str, payload: dict, dry_run: bool = False):
        super().__init__()
        if kind not in {'access', 'project'}:
            raise ValueError('알 수 없는 API 작업 종류입니다.')
        self.client, self.kind, self.payload, self.dry_run = client, kind, dict(payload), dry_run
        self.poll_interval = 1.5
        self.max_wait_seconds = 4 * 60 * 60
        self._stop_event = threading.Event()
        self.job_id = None

    def stop(self) -> None:
        self._stop_event.set()

    def execute(self) -> dict:
        """Also used by the project batch worker in its own Qt thread."""
        path = '/api/v1/' + self.kind + '-jobs'
        if self._stop_event.is_set():
            return {'status': 'cancelled', 'client_stopped': True}
        if self.dry_run:
            preview = self.client.request('POST', path + '/preview', self.payload)
            return {**preview, 'status': 'preview', 'client_stopped': self._stop_event.is_set()}
        try:
            job = self.client.request('POST', path, self.payload)
        except ApiClientError as exc:
            if exc.status_code is not None and exc.status_code < 500:
                raise
            # An accepted POST can lose its response. Never generate a new ID here.
            try:
                job = self.client.request('GET', path + '/by-request/' + quote(self.payload['request_id'], safe=''))
            except ApiClientError as recovery:
                raise ApiClientError(f'등록 결과를 확인하지 못했습니다. 같은 요청으로 재시도하세요. request_id={self.payload["request_id"]}: {recovery}') from exc
        self.job_id = job.get('job_id')
        if not self.job_id:
            raise ApiClientError('API 응답에 job_id가 없습니다.')
        job_path = path + '/' + quote(str(self.job_id), safe='')
        deadline = time.monotonic() + self.max_wait_seconds
        cancel_sent = False
        while True:
            self.progress.emit(dict(job))
            status = job.get('status')
            if status in FINAL_JOB_STATUSES:
                return {**job, 'client_stopped': self._stop_event.is_set()}
            if status not in {'queued', 'running'}:
                raise ApiClientError('알 수 없는 API Job 상태: ' + str(status))
            if time.monotonic() >= deadline:
                raise ApiClientError(f'작업 조회 대기 시간이 초과되었습니다. job_id={self.job_id}')
            if self._stop_event.is_set() and not cancel_sent:
                self.client.request('POST', job_path + '/cancel', {'requested_by': self.payload['requested_by']})
                cancel_sent = True
            # Once cancelled, wait normally to avoid busy looping on a set Event.
            if cancel_sent:
                threading.Event().wait(self.poll_interval)
            else:
                self._stop_event.wait(self.poll_interval)
            job = self.client.request('GET', job_path)

    @pyqtSlot()
    def run(self) -> None:
        try:
            self.succeeded.emit(self.execute())
        except Exception as exc:
            message = str(exc)
            if self.client.api_key:
                message = message.replace(self.client.api_key, '[API Key]')
            self.failed.emit(message)
        finally:
            self.finished.emit()


def load_api_config() -> dict:
    try:
        with open(API_CONFIG_FILE,encoding='utf-8') as handle:
            data = json.load(handle)
        if isinstance(data,dict):
            return dict(base_url=data.get('base_url',DEFAULT_API_URL),api_key=data.get('api_key',''))
    except (OSError,ValueError):
        pass
    return dict(base_url=DEFAULT_API_URL,api_key='')


class ApiConnectionTask(QObject):
    succeeded = pyqtSignal(dict)
    failed = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self,client):
        super().__init__()
        self.client = client

    @pyqtSlot()
    def run(self):
        try:
            self.succeeded.emit(dict(health=self.client.health(),ready=self.client.ready()))
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            self.finished.emit()


class ApiSettingsDialog(QDialog):
    def __init__(self,config,parent=None):
        super().__init__(parent)
        self.setWindowTitle('API 연결 설정')
        self.resize(500,220)
        self.thread = None
        self.task = None
        layout = QFormLayout(self)
        self.url = QLineEdit(config.get('base_url',DEFAULT_API_URL))
        self.key = QLineEdit(config.get('api_key',''))
        self.key.setEchoMode(QLineEdit.Password)
        self.message = QLabel('API Key는 C:\\FGT\\conf\\api_client.json에 저장됩니다.')
        self.message.setWordWrap(True)
        layout.addRow('API 주소',self.url)
        layout.addRow('X-API-Key',self.key)
        self.check = QPushButton('연결 확인 (health / ready)')
        self.check.clicked.connect(self._check_connection)
        layout.addRow(self.check)
        layout.addRow(self.message)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self._save)
        self.buttons.rejected.connect(self.reject)
        layout.addRow(self.buttons)

    def values(self):
        return dict(base_url=self.url.text().strip().rstrip('/'),api_key=self.key.text().strip())

    def _validated_client(self):
        values = self.values()
        if not values['api_key']:
            raise ValueError('API Key를 입력하세요.')
        return FolderGrantApiClient(values['base_url'],values['api_key'])

    def _check_connection(self):
        if self.thread is not None:
            return
        try:
            client = self._validated_client()
        except ValueError as exc:
            self.message.setText(str(exc)); return
        self.thread = QThread(self)
        self.task = ApiConnectionTask(client)
        self.task.moveToThread(self.thread)
        self.thread.started.connect(self.task.run)
        self.task.succeeded.connect(lambda result:self.message.setText('API 연결 및 인증 확인 성공'))
        self.task.failed.connect(self.message.setText)
        self.task.finished.connect(self.task.deleteLater)
        self.task.finished.connect(self.thread.quit,type=Qt.DirectConnection)
        self.thread.finished.connect(self._connection_finished)
        self.thread.finished.connect(self.thread.deleteLater)
        self.check.setEnabled(False)
        self.buttons.setEnabled(False)
        self.url.setEnabled(False); self.key.setEnabled(False)
        self.message.setText('연결 확인 중…')
        self.thread.start()

    def _connection_finished(self):
        self.thread = None; self.task = None
        self.check.setEnabled(True); self.buttons.setEnabled(True)
        self.url.setEnabled(True); self.key.setEnabled(True)

    def _save(self):
        if self.thread is not None:
            return
        try:
            self._validated_client()
            os.makedirs(os.path.dirname(API_CONFIG_FILE),exist_ok=True)
            temporary = API_CONFIG_FILE+'.tmp'
            with open(temporary,'w',encoding='utf-8') as handle:
                json.dump(self.values(),handle,ensure_ascii=False,indent=2)
            if os.name != 'nt':
                os.chmod(temporary,0o600)
            os.replace(temporary,API_CONFIG_FILE)
        except (OSError,ValueError) as exc:
            self.message.setText(str(exc)); return
        self.accept()

    def reject(self):
        if self.thread is None:
            super().reject()

    def closeEvent(self,event):
        if self.thread is not None:
            event.ignore()
        else:
            event.accept()


class SettingsDialog(QDialog):
    def __init__(self, parent=None, saved=None, remembered=False, debug_on=False, fail_tol_default=5):
        try:
            self._fail_tol_default = int(fail_tol_default)
        except Exception:
            self._fail_tol_default = 5

        super().__init__(parent)
        self.setWindowTitle("로그인 계정")
        self.le_id = QLineEdit()
        self.le_pw = QLineEdit()
        self.le_pw.setEchoMode(QLineEdit.Password)
        if saved:
            self.le_id.setText(saved.get("id",""))
            self.le_pw.setText(saved.get("pw",""))

        self.chk_save = QCheckBox("저장")
        self.chk_save.setChecked(bool(remembered))
        
        self.chk_debug = QCheckBox("디버그 모드")
        self.chk_debug.setChecked(bool(debug_on))

        self.sb_tol = QSpinBox()
        self.sb_tol.setRange(0, 999)
        try:
            tol_saved = int((saved or {}).get("remove_fail_tol", self._fail_tol_default))
            self.sb_tol.setValue(tol_saved)
        except Exception:
            self.sb_tol.setValue(self._fail_tol_default)

        form = QFormLayout()

        self.sb_notify = QSpinBox()
        self.sb_notify.setRange(5, 60)
        self.sb_notify.setValue(int((saved or {}).get("notify_refresh_min", 10)))
        form.addRow("알림 리프레시(분)", self.sb_notify)

        form.addRow("아이디", self.le_id)
        form.addRow("비밀번호", self.le_pw)
        form.addRow("", self.chk_save)
        form.addRow("", self.chk_debug)
        form.addRow("폴더해제 실패 허용 개수", self.sb_tol)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        lay = QVBoxLayout(self)
        lay.addLayout(form)

        bottom_bar = QHBoxLayout()
        bottom_bar.setContentsMargins(0, 6, 0, 0)

        self.version_lbl = QLabel(f"{APP_VERSION_STR}")
        self.version_lbl.setObjectName("versionLabel")
        self.version_lbl.setStyleSheet("color:#8a8a8a; font-size:11px; padding:2px 4px;")
        self.version_lbl.setToolTip(f"Build: {APP_BUILD}")

        bottom_bar.addWidget(self.version_lbl, 0, Qt.AlignVCenter)
        bottom_bar.addStretch()
        bottom_bar.addWidget(btns)

        lay.addLayout(bottom_bar)

    def result(self):
        return (
            self.le_id.text().strip(),
            self.le_pw.text().strip(),
            self.chk_save.isChecked(),
            self.chk_debug.isChecked(),
            int(self.sb_tol.value()),
            int(self.sb_notify.value()),
        )

    def get_fail_tol(self) -> int:
        try:
            return int(self.sb_tol.value())
        except Exception:
            return 5

class BusSessionManager(QObject):
    stopped = pyqtSignal()
    readyChanged = pyqtSignal(bool, str)
    downloaded = pyqtSignal(str, str)
    busyChanged = pyqtSignal(bool)
    processed = pyqtSignal(object)
    newDownloaded = pyqtSignal(str, str)
    countsReady = pyqtSignal(dict)

    MAX_INIT_RETRY = 3
    INIT_RETRY_BASE_DELAY = 2.0     

    def __init__(self, dl_dir: str):
        super().__init__()
        self.dl_dir = dl_dir
        self.driver = None
        self._ready = False
        self._busy = False
        self._cancel = False
        self.debug_enabled = False
        self.debug_dir = DEBUG_DIR
        self._user = ""
        self._pw = ""

    def set_debug(self, enabled: bool, debug_dir: str | None = None):
        self.debug_enabled = bool(enabled)
        if debug_dir:
            self.debug_dir = debug_dir
        try:
            if self.debug_enabled:
                os.makedirs(self.debug_dir, exist_ok=True)
        except Exception:
            pass

    def _stabilize_grid(self, tries=2, nap=0.8):
        self._wait_overlay_gone(10)
        for _ in range(tries):
            self._click_search()
            self._wait_overlay_gone(10)
            self._go_iframe()
            time.sleep(nap)

    def set_creds(self, user: str, pw: str):
        self._user, self._pw = (user or "").strip(), (pw or "").strip()

    def is_ready(self) -> bool:
        return self._ready and self.driver is not None

    def is_busy(self) -> bool:
        return self._busy

    def _reset_filters(self):
        d = self.driver
        try:
            self._go_iframe()
            d.execute_script("""
              for (const id of ['approverYn','processYn','releaseYn']) {
                const el = document.getElementById(id);
                if (el) { el.value=''; el.dispatchEvent(new Event('change',{bubbles:true})); }
              }
            """)
        except Exception:
            pass

    def _nav_open_and_iframe(self, url: str, iframe_css: str = "iframe"):
        d = self.driver
        d.switch_to.default_content()
        d.get(url)
        WebDriverWait(d, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, iframe_css)))
        frame = d.find_element(By.CSS_SELECTOR, iframe_css)
        d.switch_to.frame(frame)

    def _set_select_value_and_fire(self, select_id: str, value: str):
        d = self.driver
        d.execute_script("""
            var sid = arguments[0], val = arguments[1];
            var s = document.getElementById(sid);
            if (s) {
                s.value = val;
                try { s.dispatchEvent(new Event('change', {bubbles:true})); } catch(e) {}
                if (typeof AllBtnYn === 'function') { try { AllBtnYn(); } catch(e) {} }
            }
        """, select_id, value)

    def _click_search_manual(self):
        d = self.driver
        try:
            WebDriverWait(d, 2).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
            return
        except Exception:
            pass

        d.switch_to.default_content()
        WebDriverWait(d, 5).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()

        WebDriverWait(d, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "iframe")))
        d.switch_to.frame(d.find_element(By.CSS_SELECTOR, "iframe"))

    def _goto_new_site(self):
        try:
            d = self.driver
            d.switch_to.default_content()
            d.get(BUS_NEW_URL)
            WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
            self._go_iframe()
            self._wait_overlay_gone(10)
            time.sleep(0.8)
            return True
        except Exception:
            return False

    def _goto_progress_site(self):
        try:
            d = self.driver
            d.switch_to.default_content()
            d.get(BUS_PROGRESS_LIST_URL)
            WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
            self._go_iframe()
            self._wait_overlay_gone(10)
            time.sleep(1.0)
            return True
        except Exception:
            return False

    def _goto_end_site(self):
        try:
            d = self.driver
            d.switch_to.default_content()
            d.get(BUS_END_LIST_URL)
            WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
            self._go_iframe()
            self._wait_overlay_gone(10)
            time.sleep(1.0)
            return True
        except Exception:
            return False

    @pyqtSlot()
    def download_new_list(self):
        if not self.is_ready():
            self.newDownloaded.emit("", "세션 준비 안됨"); return
        if self._busy:
            self.newDownloaded.emit("", "다른 작업 실행중"); return

        self._busy = True
        self.busyChanged.emit(True)
        self._cancel = False
        d = self.driver

        try:
            self._goto_new_site()
            d = self.driver
            self._go_iframe()
            try:
                d.execute_script("""
                  const p=document.getElementById('processYn');
                  if(p){ p.value='N'; p.dispatchEvent(new Event('change',{bubbles:true})); }
                """)
                time.sleep(0.2)
                self._go_iframe()
                WebDriverWait(d, 5).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
            except Exception:
                pass

            before = set(glob.glob(os.path.join(self.dl_dir, "*.xls"))) | set(glob.glob(os.path.join(self.dl_dir, "*.xlsx")))

            WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.ID, "btnExcel"))).click()
            try:
                WebDriverWait(d, 8).until(EC.visibility_of_element_located((By.ID, "txtExcelDownReason")))
                try:
                    el = d.find_element(By.ID, "txtExcelDownReason")
                    el.clear(); el.send_keys("신규확인-미완료(N) 목록")
                except Exception:
                    pass
                ok_btn = WebDriverWait(d, 6).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
                try: ok_btn.click()
                except ElementClickInterceptedException:
                    d.execute_script("arguments[0].click();", ok_btn)
            except TimeoutException:
                pass

            end = time.time() + 120
            latest = ""
            while time.time() < end:
                self._check_cancel()
                if not list(glob.glob(os.path.join(self.dl_dir, "*.crdownload"))):
                    now = set(glob.glob(os.path.join(self.dl_dir, "*.xls"))) | set(glob.glob(os.path.join(self.dl_dir, "*.xlsx")))
                    new_files = list(now - before)
                    if new_files:
                        latest = max(new_files, key=os.path.getmtime)
                        break
                time.sleep(0.3)

            if not latest:
                self.newDownloaded.emit("", "다운로드 실패"); return

            out_path = os.path.join(self.dl_dir, "신규_미완료.xls")
            try:
                if os.path.exists(out_path): os.remove(out_path)
            except Exception:
                pass
            try:
                shutil.move(latest, out_path)
            except Exception:
                try:
                    shutil.copyfile(latest, out_path)
                    try: os.remove(latest)
                    except Exception: pass
                except Exception:
                    self.newDownloaded.emit("", "파일 저장 실패"); return

            self.newDownloaded.emit(out_path, "")
        except SystemExit:
            self.newDownloaded.emit("", "사용자 취소")
        except Exception as e:
            self.newDownloaded.emit("", f"오류: {e}")
        finally:
            self._busy = False
            self.busyChanged.emit(False)

    def _read_total_entries(self) -> int:
        d = self.driver
        try:
            info_el = WebDriverWait(d, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".dataTables_info"))
            )
            txt = (info_el.text or "").strip().lower()
            m = re.search(r"(\d+)", txt)
            if m:
                return int(m.group(1))
        except Exception:
            pass

        try:
            self._stabilize_grid()
            WebDriverWait(d, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr")))
            return len(d.find_elements(By.CSS_SELECTOR, "tbody tr"))
        except Exception:
            return 0

    def collect_request_counts(self):
        if not self.is_ready():
            try:
                self.countsReady.emit(getattr(self, "_last_counts", {}))
            except Exception:
                pass
            return
        if getattr(self, "_busy", False):

            try:
                self.countsReady.emit(getattr(self, "_last_counts", {}))
            except Exception:
                pass
            return

        d = self.driver
        counts = {
            "신규미완료": 0,
            "진행-부여": 0,
            "진행-제거": 0,
            "종료-부여": 0,
            "종료-제거": 0,
        }

        try:
            try:
                self._nav_open_and_iframe(BUS_NEW_URL)

                try:
                    pre = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                except Exception:
                    pre = ""

                self._set_select_value_and_fire("processYn", "N")
                self._click_search_manual()

                try:
                    WebDriverWait(d, 7).until(
                        lambda x: (x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                                   if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else "") != (pre or "")
                    )
                except Exception:
                    time.sleep(0.6)

                txt = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip().lower()
                m = re.search(r"(\d+)", txt)
                counts["신규미완료"] = int(m.group(1)) if m else 0
            except Exception:
                pass

            try:
                self._nav_open_and_iframe(BUS_PROGRESS_LIST_URL)

                try:
                    pre = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                except Exception:
                    pre = ""

                self._set_select_value_and_fire("processYn", "N")
                self._click_search_manual()

                try:
                    WebDriverWait(d, 7).until(
                        lambda x: (x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                                   if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else "") != (pre or "")
                    )
                except Exception:
                    time.sleep(0.6)

                txt = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip().lower()
                m = re.search(r"(\d+)", txt)
                counts["진행-부여"] = int(m.group(1)) if m else 0
            except Exception:
                pass

            try:
                self._nav_open_and_iframe(BUS_PROGRESS_LIST_URL)

                try:
                    pre = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                except Exception:
                    pre = ""

                self._set_select_value_and_fire("releaseYn", "N")
                self._click_search_manual()

                try:
                    WebDriverWait(d, 7).until(
                        lambda x: (x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                                   if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else "") != (pre or "")
                    )
                except Exception:
                    time.sleep(0.6)

                txt = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip().lower()
                m = re.search(r"(\d+)", txt)
                counts["진행-제거"] = int(m.group(1)) if m else 0
            except Exception:
                pass

            try:
                self._nav_open_and_iframe(BUS_END_LIST_URL)

                try:
                    pre = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                except Exception:
                    pre = ""

                self._set_select_value_and_fire("processYn", "N")
                self._click_search_manual()

                try:
                    WebDriverWait(d, 7).until(
                        lambda x: (x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                                   if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else "") != (pre or "")
                    )
                except Exception:
                    time.sleep(0.6)

                txt = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip().lower()
                m = re.search(r"(\d+)", txt)
                counts["종료-부여"] = int(m.group(1)) if m else 0
            except Exception:
                pass

            try:
                self._nav_open_and_iframe(BUS_END_LIST_URL)

                try:
                    pre = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                except Exception:
                    pre = ""

                self._set_select_value_and_fire("releaseYn", "N")
                self._click_search_manual()

                try:
                    WebDriverWait(d, 7).until(
                        lambda x: (x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                                   if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else "") != (pre or "")
                    )
                except Exception:
                    time.sleep(0.6)

                txt = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip().lower()
                m = re.search(r"(\d+)", txt)
                counts["종료-제거"] = int(m.group(1)) if m else 0
            except Exception:
                pass

        finally:
            self._last_counts = counts
            try:
                self.countsReady.emit(counts)
            except Exception:
                pass
            
    @pyqtSlot()
    def start(self):
        def _cleanup_driver(drv):
            if not drv:
                return
            try:
                drv.quit()
            except Exception:
                try:
                    if hasattr(drv, "service") and drv.service and drv.service.process:
                        drv.service.process.kill()
                except Exception:
                    pass

        if not (self._user and self._pw):
            self._ready = False
            self.readyChanged.emit(False, "자격증명 없음")
            return

        if self.driver:
            self._ready = True
            self.readyChanged.emit(True, "이미 준비됨")
            return

        attempt = 0
        last_err = None

        while attempt < self.MAX_INIT_RETRY and not self._cancel:
            attempt += 1
            try:
                options = make_hidden_chrome_options(self.dl_dir)

                self.driver = webdriver.Chrome(options=options)
                d = self.driver

                d.get(BUS_LOGIN_URL)
                WebDriverWait(d, 20).until(EC.presence_of_element_located((By.ID, "windowsaccount")))
                d.find_element(By.ID, "windowsaccount").send_keys(self._user)
                d.find_element(By.ID, "password").send_keys(self._pw)
                d.find_element(By.ID, "btnLogin").click()
                WebDriverWait(d, 20).until(EC.url_contains("Common"))

                try:
                    d.switch_to.window(d.window_handles[-1])
                except Exception:
                    pass

                d.get(BUS_PROGRESS_LIST_URL)
                WebDriverWait(d, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, "iframe")))
                self._go_iframe()

                WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.ID, "approverYn")))
                Select(d.find_element(By.ID, "approverYn")).select_by_value("1")
                
                self._set_request_filter(REQ_GRANT)

                self._ready = True
                self.readyChanged.emit(True, f"준비됨(시도 {attempt}/{self.MAX_INIT_RETRY})")
                return

            except Exception as e:
                last_err = e
                _cleanup_driver(self.driver)
                self.driver = None
                self._ready = False
                self.readyChanged.emit(False, f"세션 초기화 실패({attempt}/{self.MAX_INIT_RETRY}): {e}")

                if attempt >= self.MAX_INIT_RETRY or self._cancel:
                    break

                delay = self.INIT_RETRY_BASE_DELAY * (2 ** (attempt - 1))
                end = time.time() + delay
                while time.time() < end and not self._cancel:
                    QApplication.processEvents()
                    time.sleep(0.05)

        self._ready = False
        self.readyChanged.emit(False, f"세션 초기화 최종 실패: {last_err}")

    @pyqtSlot()
    def stop(self):
        if getattr(self, "_stopping", False):
            return
        self._stopping = True
        try:
            self._cancel = True
            self._busy = False
            self._ready = False
            self.busyChanged.emit(False)
            self.readyChanged.emit(False, "정지됨")

            d = self.driver
            self.driver = None
            if d:
                try:
                    d.quit()
                except Exception:
                    try:
                        if hasattr(d, "service") and d.service and d.service.process:
                            d.service.process.kill()
                    except Exception:
                        pass
        finally:
            self._stopping = False
            self.stopped.emit()

    @pyqtSlot()
    def cancel_current(self):
        self._cancel = True

    def _check_cancel(self, context: str = "download"):
        if self._cancel:
            self._busy = False
            self.busyChanged.emit(False)
            self._cancel = False
            if context == "process":
                self.processed.emit([])
            else:
                self.downloaded.emit("", "사용자 취소")
            raise SystemExit

    @pyqtSlot()
    def download_list(self):
        if not self.is_ready():
            self.downloaded.emit("", "세션 준비 안됨")
            return
        if self._busy:
            self.downloaded.emit("", "다른 작업 실행중")
            return

        self._busy = True
        self.busyChanged.emit(True)
        self._cancel = False

        d = self.driver

        def _search_and_download_end(req_flag: str) -> str:
            self._goto_end_site()
            self._set_end_filter(req_flag)

            d = self.driver
            try:
                self._go_iframe()
                try:
                    WebDriverWait(d, 5).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
                except Exception:
                    d.switch_to.default_content()
                    WebDriverWait(d, 5).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
                    self._go_iframe()
            except Exception:
                pass

            before = set(glob.glob(os.path.join(self.dl_dir, "*.xls"))) | set(glob.glob(os.path.join(self.dl_dir, "*.xlsx")))

            try:
                self._go_iframe()
                WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.ID, "btnExcel"))).click()
            except Exception:
                d.switch_to.default_content()
                WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.ID, "btnExcel"))).click()

            try:
                WebDriverWait(d, 10).until(EC.visibility_of_element_located((By.ID, "txtExcelDownReason")))
                try:
                    reason = "종료권한리스트 수집-" + ("해제" if req_flag == END_FLAG_RELEASE else "부여")
                    el = d.find_element(By.ID, "txtExcelDownReason")
                    el.clear(); el.send_keys(reason)
                except Exception:
                    pass
                ok_btn = WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm")))
                try: ok_btn.click()
                except ElementClickInterceptedException:
                    d.execute_script("arguments[0].click();", ok_btn)
            except TimeoutException:
                pass

            end = time.time() + 120
            latest = ""
            while time.time() < end:
                self._check_cancel()
                if not list(glob.glob(os.path.join(self.dl_dir, "*.crdownload"))):
                    now = set(glob.glob(os.path.join(self.dl_dir, "*.xls"))) | set(glob.glob(os.path.join(self.dl_dir, "*.xlsx")))
                    new_files = list(now - before)
                    if new_files:
                        latest = max(new_files, key=os.path.getmtime)
                        break
                time.sleep(0.35)

            if not latest:
                return ""

            suffix = "해제" if req_flag == END_FLAG_RELEASE else "부여"
            out_path = os.path.join(self.dl_dir, f"종료권한리스트_{suffix}.xls")
            try:
                if os.path.exists(out_path): os.remove(out_path)
            except Exception:
                pass

            try:
                shutil.move(latest, out_path)
            except Exception:
                try:
                    shutil.copyfile(latest, out_path)
                    try: os.remove(latest)
                    except Exception: pass
                except Exception:
                    return ""

            return out_path

        def _search_and_download(reqtype: str) -> str:
            self._goto_progress_site()
            self._set_request_filter(reqtype)
            try:
                self._go_iframe()
                try:
                    WebDriverWait(d, 5).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
                except Exception:
                    d.switch_to.default_content()
                    WebDriverWait(d, 5).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
                    self._go_iframe()
            except Exception:
                pass

            before = set(glob.glob(os.path.join(self.dl_dir, "*.xls"))) | set(
                glob.glob(os.path.join(self.dl_dir, "*.xlsx"))
            )

            WebDriverWait(d, 10).until(EC.element_to_be_clickable((By.ID, "btnExcel"))).click()
            WebDriverWait(d, 10).until(EC.visibility_of_element_located((By.ID, "txtExcelDownReason")))
            try:
                reason = f"권한리스트 수집-{'해제' if reqtype == REQ_RELEASE else '부여'}"
                el = d.find_element(By.ID, "txtExcelDownReason")
                el.clear()
                el.send_keys(reason)
            except Exception:
                pass

            ok_btn = WebDriverWait(d, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
            )
            try:
                ok_btn.click()
            except ElementClickInterceptedException:
                d.execute_script("arguments[0].click();", ok_btn)

            end = time.time() + 120
            latest = ""
            while time.time() < end:
                self._check_cancel()
                if not list(glob.glob(os.path.join(self.dl_dir, "*.crdownload"))):
                    now = set(glob.glob(os.path.join(self.dl_dir, "*.xls"))) | set(
                        glob.glob(os.path.join(self.dl_dir, "*.xlsx"))
                    )
                    new_files = list(now - before)
                    if new_files:
                        latest = max(new_files, key=os.path.getmtime)
                        break
                time.sleep(0.35)

            if not latest:
                return ""

            suffix = "해제" if reqtype == REQ_RELEASE else "부여"
            out_path = os.path.join(self.dl_dir, f"권한리스트_{suffix}.xls")
            try:
                if os.path.exists(out_path):
                    os.remove(out_path)
            except Exception:
                pass

            try:
                shutil.move(latest, out_path)
            except Exception:
                try:
                    shutil.copyfile(latest, out_path)
                    try:
                        os.remove(latest)
                    except Exception:
                        pass
                except Exception:
                    return ""

            return out_path

        try:
            for f in glob.glob(os.path.join(self.dl_dir, "*")):
                try:
                    os.remove(f)
                except:
                    pass

            grant_path = _search_and_download(REQ_GRANT)
            release_path = _search_and_download(REQ_RELEASE)

            end_grant_path   = _search_and_download_end(END_FLAG_GRANT)
            end_release_path = _search_and_download_end(END_FLAG_RELEASE)

            header_g, data_g = _parse_html_best_table(grant_path) if grant_path else ([], [])
            header_r, data_r = _parse_html_best_table(release_path) if release_path else ([], [])
            has_normal = bool(grant_path or release_path)
            combined_path = ""
            if has_normal:
                header = header_g if len(header_g) >= len(header_r) else header_r
                def _pad(row, n): rr = list(row);  rr += [""] * max(0, n-len(rr)); return rr[:n]
                rows = []
                for r in data_g: rows.append(_pad(r, len(header)))
                for r in data_r: rows.append(_pad(r, len(header)))
                parts = ["<html><head><meta charset='utf-8'></head><body>",
                         "<table border='1'>",
                         "<thead><tr>" + "".join(f"<th>{htmllib.escape(str(h or ''))}</th>" for h in header) + "</tr></thead>",
                         "<tbody>"]
                for r in rows:
                    parts.append("<tr>" + "".join(f"<td>{htmllib.escape(str(c or ''))}</td>" for c in r) + "</tr>")
                parts.append("</tbody></table></body></html>")
                combined_path = os.path.join(self.dl_dir, "권한리스트_합본.xls")
                with open(combined_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(parts))

            header_eg, data_eg = _parse_html_best_table(end_grant_path) if end_grant_path else ([], [])
            header_er, data_er = _parse_html_best_table(end_release_path) if end_release_path else ([], [])
            has_end = bool(end_grant_path or end_release_path)
            combined_end_path = ""
            if has_end:
                header_e = header_eg if len(header_eg) >= len(header_er) else header_er
                def _pad_e(row, n):
                    rr = list(row); rr += [""] * max(0, n-len(rr)); return rr[:n]
                rows_e = []
                for r in data_eg: rows_e.append(_pad_e(r, len(header_e)))
                for r in data_er: rows_e.append(_pad_e(r, len(header_e)))
                parts_e = ["<html><head><meta charset='utf-8'></head><body>",
                           "<table border='1'>",
                           "<thead><tr>" + "".join(f"<th>{htmllib.escape(str(h or ''))}</th>" for h in header_e) + "</tr></thead>",
                           "<tbody>"]
                for r in rows_e:
                    parts_e.append("<tr>" + "".join(f"<td>{htmllib.escape(str(c or ''))}</td>" for c in r) + "</tr>")
                parts_e.append("</tbody></table></body></html>")
                combined_end_path = os.path.join(self.dl_dir, "종료권한리스트_합본.xls")
                with open(combined_end_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(parts_e))

            if not has_normal and not has_end:
                self.downloaded.emit("", "부여/해제 다운로드 실패")
                return

            primary = combined_path if has_normal else combined_end_path
            self.downloaded.emit(primary, "")

        except SystemExit:
            self.downloaded.emit("", "사용자 취소")
        except Exception as e:
            self.downloaded.emit("", f"오류: {e}")
        finally:
            self._busy = False
            self.busyChanged.emit(False)

    def _ensure_iframe(self):
        try:
            d = self.driver
            d.switch_to.default_content()
            frames = d.find_elements(By.TAG_NAME, "iframe")
            if frames:
                d.switch_to.frame(frames[0])
            return True
        except Exception:
            return False

    def _go_iframe(self):
        try:
            d = self.driver
            d.switch_to.default_content()
            frames = d.find_elements(By.TAG_NAME, "iframe")
            if frames:
                d.switch_to.frame(frames[0])
            return True
        except Exception:
            return False

    def _wait_overlay_gone(self, timeout=10):
        d = self.driver
        try:
            WebDriverWait(d, timeout).until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, ".swal2-container, .loading, .blockUI, .blockOverlay")
                )
            )
        except Exception:
            pass
            
    def _debug_dump(self, tag: str):
        if not self.debug_enabled:
            return
        try:
            d = self.driver
            ts = int(time.time())
            base = f"debug_{tag}_{ts}"
            html_path = os.path.join(self.debug_dir, f"{base}.html")
            png_path  = os.path.join(self.debug_dir, f"{base}.png")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(d.page_source)
            try:
                d.save_screenshot(png_path)
            except Exception:
                pass
        except Exception:
            pass
        
    def _set_request_filter(self, reqtype: str):
        self._goto_progress_site()
        d = self.driver
        self._go_iframe()

        try:
            d.execute_script("""
              for (const id of ['approverYn','processYn','releaseYn']) {
                const el = document.getElementById(id);
                if (el) { el.value=''; el.dispatchEvent(new Event('change',{bubbles:true})); }
              }
            """)
        except Exception:
            pass

        is_release = (reqtype or "").strip() in (REQ_RELEASE, "권한해제", "해제", "release", "rel")

        try:
            if is_release:
                d.execute_script("""
                  const a=document.getElementById('approverYn');
                  const p=document.getElementById('processYn');
                  const r=document.getElementById('releaseYn');
                  if (a) { a.value='';  a.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (p) { p.value='';  p.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (r) { r.value='N'; r.dispatchEvent(new Event('change',{bubbles:true})); }
                """)
            else:
                d.execute_script("""
                  const a=document.getElementById('approverYn');
                  const p=document.getElementById('processYn');
                  const r=document.getElementById('releaseYn');
                  if (a) { a.value='1'; a.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (p) { p.value='N'; p.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (r) { r.value='';  r.dispatchEvent(new Event('change',{bubbles:true})); }
                """)
        except Exception:
            pass

        time.sleep(0.3)
        try:
            vals = d.execute_script("""
              const g = id => (document.getElementById(id)||{}).value || '';
              return {a:g('approverYn'), p:g('processYn'), r:g('releaseYn')};
            """)
            if is_release:
                ok = (vals.get('a','') == '' and vals.get('p','') == '' and vals.get('r','') == 'N')
            else:
                ok = (vals.get('a','') == '1' and vals.get('p','') == 'N' and vals.get('r','') == '')
            if not ok:
                return self._set_request_filter(reqtype)
        except Exception:
            pass

        self._debug_dump("progress_release_filter" if is_release else "progress_grant_filter")

    def _set_end_filter(self, reqtype_flag: str):
        self._goto_end_site()
        d = self.driver
        self._go_iframe()

        try:
            d.execute_script("""
              for (const id of ['approverYn','processYn','releaseYn']) {
                const el = document.getElementById(id);
                if (el) { el.value=''; el.dispatchEvent(new Event('change',{bubbles:true})); }
              }
            """)
        except Exception:
            pass

        try:
            if reqtype_flag == END_FLAG_GRANT:
                d.execute_script("""
                  const a=document.getElementById('approverYn');
                  const p=document.getElementById('processYn');
                  const r=document.getElementById('releaseYn');
                  if (a) { a.value='1'; a.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (p) { p.value='N'; p.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (r) { r.value='';  r.dispatchEvent(new Event('change',{bubbles:true})); }
                """)
            else:
                d.execute_script("""
                  const a=document.getElementById('approverYn');
                  const p=document.getElementById('processYn');
                  const r=document.getElementById('releaseYn');
                  if (a) { a.value='';  a.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (p) { p.value='';  p.dispatchEvent(new Event('change',{bubbles:true})); }
                  if (r) { r.value='N'; r.dispatchEvent(new Event('change',{bubbles:true})); }
                """)
        except Exception:
            pass

        time.sleep(0.3)
        try:
            vals = d.execute_script("""
              const g = id => (document.getElementById(id)||{}).value || '';
              return {a:g('approverYn'), p:g('processYn'), r:g('releaseYn')};
            """)
            if reqtype_flag == END_FLAG_GRANT:
                ok = (vals.get('a','') == '1' and vals.get('p','') == 'N' and vals.get('r','') == '')
            else:
                ok = (vals.get('a','') == '' and vals.get('p','') == '' and vals.get('r','') == 'N')
            if not ok:
                return self._set_end_filter(reqtype_flag)
        except Exception:
            pass

        self._debug_dump("end_grant_filter" if reqtype_flag == END_FLAG_GRANT else "end_release_filter")

    def _click_search(self):
        d = self.driver
        try:
            self._go_iframe()
            try:
                WebDriverWait(d, 4).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
            except Exception:
                d.switch_to.default_content()
                WebDriverWait(d, 4).until(EC.element_to_be_clickable((By.ID, "btnSearch"))).click()
            self._wait_overlay_gone(10)
            self._go_iframe()
        except Exception:
            pass

    def _for_each_page(self, work_on_page):
        d = self.driver
        tried = 0

        self._check_cancel("process")
        if work_on_page():
            return True

        while tried < 200:
            self._check_cancel("process")            
            self._go_iframe()
            next_btn = None
            try:
                cands = d.find_elements(By.CSS_SELECTOR, "a[aria-label='Next'], button[aria-label='Next']")
                if not cands:
                    cands = d.find_elements(By.XPATH, "//a[normalize-space()='다음' or normalize-space()='>'] | //button[normalize-space()='다음' or normalize-space()='>']")
                if cands:
                    next_btn = cands[-1]
                    cls = (next_btn.get_attribute("class") or "").lower()
                    aria = (next_btn.get_attribute("aria-disabled") or "").lower()
                    if "disabled" in cls or aria in ("true", "1"):
                        break
            except Exception:
                next_btn = None

            if not next_btn:
                break

            try:
                d.execute_script("arguments[0].scrollIntoView({block:'center'});", next_btn)
            except Exception:
                pass

            self._check_cancel("process")
            try:
                next_btn.click()
            except ElementClickInterceptedException:
                d.execute_script("arguments[0].click();", next_btn)

            self._wait_overlay_gone(10)
            self._go_iframe()
            time.sleep(0.4)

            self._check_cancel("process")
            if work_on_page():
                return True

            tried += 1

        return False

    def _norm(self, s):
        return (s or "").strip().replace("\u00a0", " ").lower()

    def _bus_code_variants(self, proj_raw: str):
        try:
            code5 = _yyxxx_from_proj_for_groupname(proj_raw)
            c5_int = int(code5)
            code4 = None
            if c5_int < 20000 and len(code5) == 5 and code5[2] == "0":
                code4 = code5[:2] + code5[3:]
            return code5, code4
        except Exception:
            base = re.sub(r"\D", "", str(proj_raw or ""))
            if len(base) >= 5:
                base = base[-5:]
            return base, None

    def _goto_first_page(self):
        d = self.driver
        try:
            self._go_iframe()
            btns = d.find_elements(By.XPATH, "//a[normalize-space()='처음'] | //button[normalize-space()='처음']")
            if btns:
                try:
                    d.execute_script("arguments[0].scrollIntoView({block:'center'});", btns[0])
                except Exception:
                    pass
                try:
                    btns[0].click()
                except ElementClickInterceptedException:
                    d.execute_script("arguments[0].click();", btns[0])
                self._wait_overlay_gone(10)
                self._go_iframe()
                return
            
            page1 = d.find_elements(By.XPATH, "//a[normalize-space()='1'] | //button[normalize-space()='1']")
            if page1:
                try:
                    d.execute_script("arguments[0].scrollIntoView({block:'center'});", page1[0])
                except Exception:
                    pass
                try:
                    page1[0].click()
                except ElementClickInterceptedException:
                    d.execute_script("arguments[0].click();", page1[0])
                self._wait_overlay_gone(10)
                self._go_iframe()
        except Exception:
            pass

    def _narrow_by_path_strict(self, rows, lv2, lv3, path_hint):
        if not rows:
            return rows

        def N(s): return (s or "").replace("\u00a0"," ").strip().lower()
        lv2n = N(lv2)
        lv3n = N(lv3)
        tail = (path_hint or "").split("\\")[-1]
        tailn = N(tail)

        must_tokens = [t for t in [lv2n, lv3n, tailn] if t]

        narrowed = []
        for r in rows:
            try:
                txt = N(r.text).replace("\n", " ")
                if all(tok in txt for tok in must_tokens):
                    narrowed.append(r)
            except Exception:
                pass

        return narrowed

    @pyqtSlot(object)
    def process(self, targets):
        if not self.is_ready():
            self.processed.emit([{'row': t.get('row'), 'ok': False, 'msg': '세션 준비 안됨'} for t in (targets or [])])
            return
        if self._busy:
            self.processed.emit([{'row': t.get('row'), 'ok': False, 'msg': '다른 작업 실행중'} for t in (targets or [])])
            return

        self._busy = True
        self.busyChanged.emit(True)
        self._cancel = False

        self._check_cancel("process")

        results = []
        try:
            d = self.driver

            def count_user_rows(user_text: str) -> int:
                self._go_iframe()
                user_eq = f".//td[normalize-space()='{user_text}']"
                code_pred = f".//td[contains(normalize-space(), '{code5}')]"
                if code4:
                    code_pred = f"({code_pred} or .//td[contains(normalize-space(), '{code4}')])"
                xpath = f"//tbody/tr[{user_eq} and {code_pred}]"
                rows = d.find_elements(By.XPATH, xpath)
                rows = self._narrow_by_path_strict(rows, lv2, lv3, path_hint)
                return len(rows)

            self._go_iframe()

            def _row_signature(tr):
                try:
                    tds = tr.find_elements(By.TAG_NAME, "td")
                    txts = []
                    for td in tds:
                        txt = (td.text or "").replace("\u00a0"," ").strip()
                        txts.append(re.sub(r"\s+", " ", txt))
                    return " | ".join(txts)
                except Exception:
                    return ""

            def _signature_exists(sig: str) -> bool:
                found = False
                def _find_on_page():
                    nonlocal found
                    self._go_iframe()
                    rows = d.find_elements(By.CSS_SELECTOR, "tbody tr")
                    for rnode in rows:
                        if _row_signature(rnode) == sig:
                            found = True
                            return True
                    return False
                self._for_each_page(_find_on_page)
                return found

            for t in (targets or []):
                self._check_cancel("process")
                row_idx = t.get('row')
                user = (t.get('user') or '').strip()
                lv2 = (t.get('lv2') or '').strip()
                lv3 = (t.get('lv3') or '').strip()
                path_hint = (t.get('path') or '').strip()
                kind = (t.get('kind') or '진행').strip()
                req = (t.get('req') or REQ_GRANT).strip()
                proj = (t.get('proj') or '').strip()
                code5, code4 = self._bus_code_variants(proj)

                ok = False
                msg = ""

                try:
                    if kind == "종료":
                        self._set_end_filter(END_FLAG_RELEASE if req == REQ_RELEASE else END_FLAG_GRANT)
                    else:
                        self._set_request_filter(req)

                    self._click_search()
                    time.sleep(0.6)
                    self._goto_first_page()

                    self._stabilize_grid()

                    before_total = 0
                    def sum_on_page():
                        nonlocal before_total
                        before_total += count_user_rows(user)
                        return False
                    self._for_each_page(sum_on_page)
                    target_sig_holder = {"sig": ""}

                    def work_on_page():
                        self._check_cancel("process")
                        self._go_iframe()

                        user_eq = f".//td[normalize-space()='{user}']"
                        code_pred = f".//td[contains(normalize-space(), '{code5}')]"
                        if code4:
                            code_pred = f"({code_pred} or .//td[contains(normalize-space(), '{code4}')])"

                        xpath = f"//tbody/tr[{user_eq} and {code_pred}]"
                        cand_rows = d.find_elements(By.XPATH, xpath)
                        
                        if (t.get('kind') or '').strip() != '종료':
                            cand_rows = self._narrow_by_path_strict(cand_rows, lv2, lv3, path_hint)
                       
                        if not cand_rows:
                            self._debug_dump("no_rows_after_filter")
                            return False

                        target_row = cand_rows[0]

                        req_is_release = (t.get('req') or REQ_GRANT) == REQ_RELEASE
                        if req_is_release:
                            btn_xpath = (
                                ".//button[normalize-space()='해제처리' "
                                "or contains(@class,'btn') and (contains(normalize-space(),'해제') or contains(@onclick,'Release'))]"
                            )
                        else:
                            btn_xpath = (
                                ".//button[normalize-space()='완료처리' "
                                "or normalize-space()='처리' "
                                "or contains(@class,'btn-outline-aurora')]"
                            )

                        target_sig_holder["sig"] = _row_signature(target_row)

                        self._check_cancel("process")

                        try:
                            d.execute_script("arguments[0].scrollIntoView({block:'center'});", target_row)
                        except Exception:
                            pass

                        def _btn_ready(_):
                            try:
                                el = target_row.find_element(By.XPATH, btn_xpath)
                                return el if el.is_displayed() and el.is_enabled() else False
                            except StaleElementReferenceException:
                                return False

                        try:
                            btn = WebDriverWait(target_row, 10).until(_btn_ready)
                        except TimeoutException:
                            self._debug_dump("no_action_button_wait_timeout")
                            return False

                        try:
                            d.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                        except Exception:
                            pass

                        try:
                            btn.click()
                        except ElementClickInterceptedException:
                            d.execute_script("arguments[0].click();", btn)
                        except Exception:
                            pass

                        if self._cancel:
                            try:
                                self._wait_overlay_gone(0.2)
                                try:
                                    cancel_btn = WebDriverWait(d, 1).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-cancel"))
                                    )
                                    try:
                                        cancel_btn.click()
                                    except ElementClickInterceptedException:
                                        d.execute_script("arguments[0].click();", cancel_btn)
                                except TimeoutException:
                                    d.execute_script("if (window.Swal) { try{Swal.close();}catch(e){} }")
                            except Exception:
                                pass
                            self._check_cancel("process")
                            return False

                        try:
                            WebDriverWait(d, 3).until(
                                EC.visibility_of_element_located((By.CSS_SELECTOR, ".swal2-container"))
                            )
                            self._check_cancel("process")
                            if self._cancel:
                                try:
                                    cancel_btn = WebDriverWait(d, 1).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-cancel"))
                                    )
                                    try:
                                        cancel_btn.click()
                                    except ElementClickInterceptedException:
                                        d.execute_script("arguments[0].click();", cancel_btn)
                                except TimeoutException:
                                    d.execute_script("if (window.Swal) { try{Swal.close();}catch(e){} }")
                                self._check_cancel("process")
                                return False

                            okbtn = WebDriverWait(d, 5).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                            )
                            try:
                                okbtn.click()
                            except ElementClickInterceptedException:
                                d.execute_script("arguments[0].click();", okbtn)
                        except TimeoutException:
                            pass
                        
                        self._wait_overlay_gone(10)
                        self._stabilize_grid()                        
                        return True

                    self._goto_first_page()
                    did_click = self._for_each_page(work_on_page)
                    if not did_click:
                        ok, msg = False, "사번/프로젝트코드 불일치"
                    else:
                        timeout = time.time() + 12
                        disappeared = False
                        while time.time() < timeout:
                            self._check_cancel("process")
                            self._click_search()
                            time.sleep(0.8)
                            if target_sig_holder["sig"] and not _signature_exists(target_sig_holder["sig"]):
                                disappeared = True
                                break

                        if disappeared:
                            ok, msg = True, "처리 완료"
                        else:
                            after_total = 0
                            def sum_after_page():
                                nonlocal after_total
                                after_total += count_user_rows(user)
                                return False
                            self._for_each_page(sum_after_page)

                            ok = after_total < before_total
                            msg = f"처리 {'완료' if ok else '미확인'}: 전={before_total}, 후={after_total}"

                except SystemExit:
                    return
                except Exception as e:
                    ok, msg = False, f"오류: {e}"

                if not ok and self.debug_enabled:
                    try:
                        ts = int(time.time())
                        tag_user = re.sub(r'[^0-9A-Za-z_-]+', '_', user)[:30]
                        tag_lv3  = re.sub(r'[^0-9A-Za-z_-]+', '_', lv3)[:20]
                        base = f"fail_{ts}_{tag_user}_{tag_lv3}"
                        with open(os.path.join(self.debug_dir, f"{base}.html"), "w", encoding="utf-8") as f:
                            f.write(d.page_source)
                        try:
                            d.save_screenshot(os.path.join(self.debug_dir, f"{base}.png"))
                        except Exception:
                            pass
                    except Exception:
                        pass

                results.append({'row': row_idx, 'ok': ok, 'msg': msg})

            self._busy = False
            self.busyChanged.emit(False)
            self.processed.emit(results)

        except SystemExit:
            self._busy = False
            self.busyChanged.emit(False)
            if results:
                self.processed.emit(results)
            return
        except Exception as e:
            self._busy = False
            self.busyChanged.emit(False)
            self.processed.emit([{'row': t.get('row'), 'ok': False, 'msg': f'오류: {e}'} for t in (targets or [])])

class BusWatcher(QObject):
    stopped = pyqtSignal()
    readyChanged = pyqtSignal(bool, str)
    countsReady = pyqtSignal(dict)

    def __init__(self, dl_dir):
        super().__init__()
        self._user = ""
        self._pw = ""
        self._last_counts = {"신규미완료":0, "진행-부여":0, "진행-제거":0, "종료-부여":0, "종료-제거":0}
        self._mgr = BusSessionManager(dl_dir)

    def set_creds(self, user, pw):
        self._user, self._pw = (user or "").strip(), (pw or "").strip()

    def set_debug(self, on, debug_dir):
        try:
            self._mgr.set_debug(bool(on), debug_dir)
        except Exception:
            pass

    def is_ready(self):
        try:
            return self._mgr.is_ready()
        except Exception:
            return False

    @pyqtSlot()
    def start(self):
        if not (self._user and self._pw):
            self.readyChanged.emit(False, "워처 자격증명 없음")
            return
        try:
            self._mgr.set_creds(self._user, self._pw)
            self._mgr.start()
            self.readyChanged.emit(True, "워처 준비됨")
        except Exception:
            self.readyChanged.emit(False, "워처 시작 실패")

    @pyqtSlot()
    def stop(self):
        try:
            self._mgr.stop()
        except Exception:
            pass
        self.readyChanged.emit(False, "워처 정지됨")
        self.stopped.emit()

    @pyqtSlot()
    def collect_counts(self):
        try:
            self._mgr.collect_request_counts()
            self._last_counts = getattr(self._mgr, "_last_counts", self._last_counts)
            self.countsReady.emit(self._last_counts)
        except Exception:
            self.countsReady.emit(self._last_counts)

class CheckBoxHeader(QHeaderView):
    stateChanged = pyqtSignal(bool)

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self.setSectionsClickable(True)
        self.setHighlightSections(False)

        self._cb = QCheckBox(self)
        self._cb.setTristate(False)
        self._cb.setChecked(True)
        self._cb.setStyleSheet("QCheckBox{background:transparent; padding:0; margin:0;}")

        self._cb.stateChanged.connect(lambda st: self.stateChanged.emit(st == Qt.Checked))

        self.sectionResized.connect(self._reposition)
        self.sectionMoved.connect(self._reposition)
        self.geometriesChanged.connect(self._reposition)
        self.sectionCountChanged.connect(lambda *_: self._reposition())

        self._reposition()

    def setChecked(self, on: bool):
        self._cb.blockSignals(True)
        self._cb.setChecked(bool(on))
        self._cb.blockSignals(False)
        self.updateSection(0)

    def _reposition(self, *args):
        if self.count() == 0:
            return

        x0 = self.sectionViewportPosition(0)
        w0 = self.sectionSize(0)
        h  = self.height()

        sz = self._cb.sizeHint()
        x  = x0 + max(0, (w0 - sz.width()) // 2)
        y  = max(0, (h  - sz.height()) // 2)

        if w0 < sz.width():
            x = x0 + 2

        self._cb.setGeometry(x, y, sz.width(), sz.height())
        self._cb.raise_()
        self._cb.show()

    def mousePressEvent(self, e):
        if self.logicalIndexAt(e.pos()) == 0:
            self._cb.toggle()
            e.accept()
            return
        super().mousePressEvent(e)


class CopyTable(QTableWidget):
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            ranges = self.selectedRanges()
            if not ranges:
                super().keyPressEvent(event); return
            blocks = []
            for rng in ranges:
                rows = []
                for r in range(rng.topRow(), rng.bottomRow()+1):
                    cols = []
                    for c in range(rng.leftColumn(), rng.rightColumn()+1):
                        it = self.item(r, c)
                        cols.append("" if it is None else it.text())
                    rows.append("\t".join(cols))
                blocks.append("\n".join(rows))
            QGuiApplication.clipboard().setText("\n\n".join(blocks))
        else:
            super().keyPressEvent(event)

class ManualEntryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("수동 입력")
        self.setModal(True)

        self.cb_kind = QComboBox()
        self.cb_kind.addItems(["진행", "종료"])
        self.cb_kind.setCurrentText("진행")
        self.cb_kind.currentTextChanged.connect(self._on_kind_changed)

        self.cb_reqtype = QComboBox()
        self.le_user = QLineEdit()
        self.le_proj = QLineEdit()
        self.cb_lv2  = QComboBox()
        self.cb_lv3  = QComboBox()
        self.cb_role = QComboBox()
        
        self.cb_reqtype.addItems([REQ_GRANT, REQ_RELEASE])
        self.cb_reqtype.setCurrentText(REQ_GRANT)

        self.cb_lv2.addItems(["Study", "Isolated"])
        self.cb_lv2.setCurrentIndex(-1)

        self.cb_lv3.addItems(LEVEL3_CHOICES)
        self.cb_lv3.setCurrentIndex(-1)

        self.cb_role.addItem("")
        self.cb_role.setEnabled(False)

        self.cb_lv2.currentTextChanged.connect(self._refresh_role_candidates)
        self.cb_lv3.currentTextChanged.connect(self._refresh_role_candidates)
        self.le_proj.textChanged.connect(self._refresh_role_candidates)

        form = QFormLayout()
        form.addRow("구분*", self.cb_kind) 
        form.addRow("요청사항*", self.cb_reqtype) 
        form.addRow("사번*", self.le_user)
        form.addRow("프로젝트코드*", self.le_proj)
        form.addRow("Level2*", self.cb_lv2)
        form.addRow("Level3*", self.cb_lv3)
        form.addRow("STATROLE", self.cb_role)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)

        lay = QVBoxLayout(self)
        lay.addLayout(form)
        lay.addWidget(btns)

        self.result_row = None

    def _is_new_by_proj_text(self) -> bool:
        txt = self.le_proj.text().strip()
        if not txt:
            return True
        try:
            return is_new_template(txt)
        except Exception:
            return True

    def _refresh_role_candidates(self):
        try:
            lv2 = self.cb_lv2.currentText().strip() if self.cb_lv2.currentIndex() != -1 else ""
            lv3 = self.cb_lv3.currentText().strip() if self.cb_lv3.currentIndex() != -1 else ""
            is_new = self._is_new_by_proj_text()

            self.cb_role.blockSignals(True)
            self.cb_role.clear()
            self.cb_role.addItem("")
            self.cb_role.blockSignals(False)
            self.cb_role.setCurrentIndex(0)
            self.cb_role.setEnabled(False)

            if not is_stat_lv3(lv3):
                return

            candidates = []
            if lv2 == "Study":
                candidates = sorted(STUDY_ROLES)
            elif lv2 == "Isolated":
                if is_stat_idmc_lv3(lv3):
                    if is_stat_idmc_new_policy(self.le_proj.text().strip()):
                        candidates = sorted(STUDY_ROLES)
                else:
                    candidates = sorted(ISOLATED_ROLES) if is_new else ["Randomization Statistician"]

            if candidates:
                self.cb_role.blockSignals(True)
                for r in candidates:
                    self.cb_role.addItem(r)
                self.cb_role.blockSignals(False)
                self.cb_role.setEnabled(True)
        except Exception as e:
            parent = self.parent()
            if parent and hasattr(parent, "_log"):
                parent._log(f"[역할후보 갱신 예외] {e}")

    def _on_kind_changed(self, text: str):
        """구분이 '종료'면 Level2/Level3 입력을 비활성화하고 값을 비움."""
        is_end = (text or "").strip() == "종료"

        for cb in (self.cb_lv2, self.cb_lv3):
            try:
                cb.blockSignals(True)
                if is_end:
                    cb.setCurrentIndex(-1)
                cb.setEnabled(not is_end)
            finally:
                cb.blockSignals(False)

        try:
            self._refresh_role_candidates()
        except Exception:
            pass

    def _on_ok(self):
        try:
            kind   = self.cb_kind.currentText().strip()
            reqtype = self.cb_reqtype.currentText().strip()
            user = self.le_user.text().strip()
            proj = self.le_proj.text().strip()
            lv2  = self.cb_lv2.currentText().strip() if self.cb_lv2.currentIndex() != -1 else ""
            lv3  = self.cb_lv3.currentText().strip() if self.cb_lv3.currentIndex() != -1 else ""
            role = self.cb_role.currentText().strip() if self.cb_role.isEnabled() else ""

            missing = []
            if not user: missing.append("사번")
            if not proj: missing.append("프로젝트코드")
            if kind != "종료":
                if not lv2:  missing.append("Level2")
                if not lv3:  missing.append("Level3")
            if missing:
                QMessageBox.warning(self, "입력 누락", f"다음 항목을 입력해 주세요: {', '.join(missing)}")
                return

            if not is_stat_lv3(lv3):
                role = ""

            self.result_row = (reqtype, user, proj, lv2, lv3, "", role, kind)
            self.accept()
        except Exception as e:
            try:
                parent = self.parent()
                if parent and hasattr(parent, "_log"):
                    parent._log(f"[수동입력 예외] {e}")
            except:
                pass
            QMessageBox.critical(self, "오류", f"수동 입력 처리 중 오류:\n{e}")

    def _set_roles_for_lv2(self, lv2_text):
        self.cb_role.clear()
        self.cb_role.addItem("")
        if lv2_text == "Study":
            for r in sorted(STUDY_ROLES):
                self.cb_role.addItem(r)
        elif lv2_text == "Isolated":
            for r in sorted(ISOLATED_ROLES):
                self.cb_role.addItem(r)
        self.cb_role.setCurrentIndex(0)

class CheckBoxHeaderAt(QHeaderView):
    stateChanged = pyqtSignal(bool)

    def __init__(self, orientation, parent=None, target_index: int = 0, initially_checked: bool = True):
        super().__init__(orientation, parent)
        self._target = int(target_index)

        self._cb = QCheckBox(self)
        self._cb.setTristate(False)
        self._cb.setChecked(bool(initially_checked))
        self._cb.setStyleSheet("QCheckBox{background:transparent; padding:0; margin:0;}")
        self._cb.stateChanged.connect(lambda st: self.stateChanged.emit(st == Qt.Checked))

        self.setSectionsClickable(True)
        self.setHighlightSections(False)

        self.sectionResized.connect(self._reposition)
        self.sectionMoved.connect(self._reposition)
        self.sectionCountChanged.connect(lambda *_: self._reposition())

    def showEvent(self, e):
        super().showEvent(e)
        self._reposition()

    def setTargetSection(self, idx: int):
        self._target = int(idx)
        self._reposition()

    def setChecked(self, on: bool):
        self._cb.blockSignals(True)
        self._cb.setChecked(bool(on))
        self._cb.blockSignals(False)
        self.updateSection(self._target)

    def _reposition(self, *args):
        if self.count() == 0 or self._target < 0 or self._target >= self.count():
            self._cb.hide(); return
        x0 = self.sectionViewportPosition(self._target)
        w0 = self.sectionSize(self._target)
        h  = self.height()
        sz = self._cb.sizeHint()
        x  = x0 + max(0, (w0 - sz.width()) // 2)
        y  = max(0, (h  - sz.height()) // 2)
        self._cb.setGeometry(x, y, sz.width(), sz.height())
        self._cb.raise_()
        self._cb.show()

    def mousePressEvent(self, e):
        if self.logicalIndexAt(e.pos()) == self._target:
            self._cb.toggle()
            e.accept()
            return
        super().mousePressEvent(e)

class ManualNewRequestDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("수동 추가")
        lay = QVBoxLayout(self)

        form = QFormLayout()
        self.edt_proj = QLineEdit(self)
        self.edt_proj.setPlaceholderText("예: 25001")
        self.edt_proj.setMaxLength(5)
        self.edt_name = QLineEdit(self)
        self.edt_name.setPlaceholderText("프로젝트명")
        form.addRow("프로젝트 코드", self.edt_proj)
        form.addRow("프로젝트명", self.edt_name)
        lay.addLayout(form)

        btn_box = QDialogButtonBox(Qt.Horizontal, self)
        self.btn_ok = btn_box.addButton("생성", QDialogButtonBox.AcceptRole)
        self.btn_cancel = btn_box.addButton("취소", QDialogButtonBox.RejectRole)
        self.btn_ok.clicked.connect(self.accept)
        self.btn_cancel.clicked.connect(self.reject)
        lay.addWidget(btn_box)

    def values(self) -> tuple[str, str]:
        return self.edt_proj.text().strip(), self.edt_name.text().strip()


class NewItemsViewer(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("신규 폴더 생성요청 목록")
        self.resize(800, 520)

        lay = QVBoxLayout(self)
        self.tbl = QTableWidget(0, 0, self)

        hdr = CheckBoxHeader(Qt.Horizontal, self.tbl)
        hdr.stateChanged.connect(self._toggle_all_rows)
        self.tbl.setHorizontalHeader(hdr)

        self.tbl.setAlternatingRowColors(True)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.horizontalHeader().setStretchLastSection(True)

        btns = QHBoxLayout()
        self.btn_manual = QPushButton("수동 추가")
        self.btn_create = QPushButton("자동 생성")
        self.btn_refresh = QPushButton("↻ 새로고침")
        self.btn_close = QPushButton("닫기")
        self.btn_close.clicked.connect(self.accept)
        self.btn_create.clicked.connect(self._on_create_clicked)
        self.btn_manual.clicked.connect(self._on_manual_clicked)
        btns.addStretch()
        btns.addWidget(self.btn_manual)
        btns.addWidget(self.btn_create)
        btns.addWidget(self.btn_refresh)
        btns.addWidget(self.btn_close)
        lay.addWidget(self.tbl); lay.addLayout(btns)

        self.status_lbl = QLabel("")
        self.status_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.status_lbl.setMinimumWidth(420)
        self.status_lbl.setStyleSheet("padding-right: 12px; padding-bottom: 2px;")
        self.prg = QProgressBar(self)
        self.prg.setRange(0, 0)
        self.prg.setFixedSize(0, 0)
        self.prg.hide()

        bar = QHBoxLayout()
        bar.setContentsMargins(0, 0, 12, 6)
        bar.addStretch()
        bar.addWidget(self.status_lbl, 0, Qt.AlignRight | Qt.AlignVCenter)
        lay.addLayout(bar)

        self._last_worker_msg = ""
        self.worker_thread = None
        self.worker = None
        self.chk_dry = QCheckBox('Dry Run')
        self.chk_auto_complete = QCheckBox('실제 성공 후 BUS 완료')
        self.chk_auto_complete.setChecked(True)
        self.btn_stop = QPushButton('중지')
        self.btn_stop.clicked.connect(lambda: self.worker.stop() if self.worker else None)
        btns.insertWidget(0,self.chk_dry)
        btns.insertWidget(1,self.chk_auto_complete)
        btns.insertWidget(2,self.btn_stop)

    def _set_busy(self, on: bool, msg: str = ""):
        self.status_lbl.setText(msg or "")
        self.prg.hide()
        self.btn_manual.setEnabled(not on)
        self.btn_create.setEnabled(not on)
        self.btn_refresh.setEnabled(not on)
        self.btn_close.setEnabled(not on)
        self.chk_dry.setEnabled(not on)
        self.chk_auto_complete.setEnabled(not on)
        self.btn_stop.setEnabled(on)
        QApplication.processEvents()

    def _on_worker_progress(self, cur: int, total: int, message: str):
        self.status_lbl.setText(message)
        self._last_worker_msg = message
        QApplication.processEvents()

    def _on_worker_finished(self,ok_cnt,fail_cnt):
        self.status_lbl.setText(f'실제 성공 {ok_cnt} / 실패 {fail_cnt} — '+self._last_worker_msg)


    def _manual_add(self,proj,name):
        if not self.tbl.columnCount():
            self.set_data(['프로젝트코드','프로젝트명'],[])
        row = self.tbl.rowCount()
        self.tbl.insertRow(row)
        checkbox = QCheckBox(); checkbox.setChecked(True)
        holder = QWidget(); layout = QHBoxLayout(holder)
        layout.addWidget(checkbox); layout.setContentsMargins(0,0,0,0)
        self.tbl.setCellWidget(row,0,holder)
        self.tbl.setItem(row,self._hidx['프로젝트코드'],QTableWidgetItem(proj))
        self.tbl.setItem(row,self._hidx['프로젝트명'],QTableWidgetItem(name))
        self.tbl.item(row,self._hidx['프로젝트코드']).setData(Qt.UserRole,
            dict(source='test-client',request_id=new_client_request_id(),bus_done=False))
        self.status_lbl.setText('수동 요청을 추가했습니다. Dry Run 또는 API 생성을 실행하세요.')


    def _on_manual_clicked(self):
        dlg = ManualNewRequestDialog(self)
        if dlg.exec_() != QDialog.Accepted:
            return

        proj, name = dlg.values()
        if not proj:
            QMessageBox.warning(self, "알림", "프로젝트 코드를 입력하세요.")
            return

        self._manual_add(proj, name)

    def _hide_empty_columns(self, always_show: set[str] | None = None, skip_col_idx: int | None = None):
        t = self.tbl
        if t is None:
            return
        rows, cols = t.rowCount(), t.columnCount()
        always_show = always_show or set()

        for c in range(cols):
            t.setColumnHidden(c, False)

        for c in range(cols):
            if skip_col_idx is not None and c == skip_col_idx:
                continue
            header = t.horizontalHeaderItem(c).text() if t.horizontalHeaderItem(c) else ""
            if header in always_show:
                continue
            has_data = False
            for r in range(rows):
                it = t.item(r, c)
                if it and it.text().strip():
                    has_data = True
                    break
            t.setColumnHidden(c, not has_data)

    def set_data(self, header, rows):
        try:
            self.tbl.blockSignals(True)

            col_select = 0
            col_count = 1 + len(header)

            self.tbl.clear()
            self.tbl.setColumnCount(col_count)
            self.tbl.setRowCount(len(rows))

            it_sel = QTableWidgetItem("")
            it_sel.setToolTip("선택")
            self.tbl.setHorizontalHeaderItem(col_select, it_sel)
            self.tbl.horizontalHeader().setSectionResizeMode(col_select, QHeaderView.Fixed)
            self.tbl.setColumnWidth(col_select, 28)

            for i, h in enumerate(header, start=1):
                self.tbl.setHorizontalHeaderItem(i, QTableWidgetItem(h or ""))

            for r, row in enumerate(rows):
                cb = QCheckBox(self.tbl)
                cb.setChecked(True)
                wrap = QWidget(self.tbl)
                box = QHBoxLayout(wrap); box.setContentsMargins(0,0,0,0)
                box.addWidget(cb, 0, Qt.AlignCenter)
                self.tbl.setCellWidget(r, col_select, wrap)

                for c, v in enumerate(row, start=1):
                    it = QTableWidgetItem(str(v) if v is not None else "")
                    it.setTextAlignment(Qt.AlignCenter)
                    self.tbl.setItem(r, c, it)

            for c in range(1, col_count):
                self.tbl.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)
            self.tbl.horizontalHeader().setStretchLastSection(True)

            self._hide_empty_columns(always_show={"프로젝트코드","프로젝트명"}, skip_col_idx=0)
        finally:
            self.tbl.blockSignals(False)

        self._hidx = { (self.tbl.horizontalHeaderItem(i).text() or ""): i
                       for i in range(1, self.tbl.columnCount()) }
        col = self._hidx.get('프로젝트코드',-1)
        name_col = self._hidx.get('프로젝트명',self._hidx.get('과제명',self._hidx.get('제목',-1)))
        for row in range(self.tbl.rowCount()):
            item = self.tbl.item(row,col) if col >= 0 else None
            name_item = self.tbl.item(row,name_col) if name_col >= 0 else None
            if item:
                values = dict(proj=item.text(),project_name=name_item.text() if name_item else '')
                identity = bus_row_identity(header,rows[row])
                try:
                    request_id = make_bus_request_id(values,identity)
                except ValueError:
                    request_id = new_client_request_id()
                item.setData(Qt.UserRole,dict(source='bus',request_id=request_id,bus_done=False))

    def _toggle_all_rows(self, checked: bool):
        t = self.tbl
        for r in range(t.rowCount()):
            w = t.cellWidget(r, 0)
            if not w:
                continue
            cb = w.findChild(QCheckBox)
            if cb:
                cb.setChecked(checked)

    def _selected_targets(self):
        t = self.tbl
        items = []
        code_col = self._hidx.get("프로젝트코드", -1)
        name_col = self._hidx.get("프로젝트명", -1)

        if code_col < 0:
            QMessageBox.warning(self, "알림", "헤더 '프로젝트코드'를 찾을 수 없습니다.")
            return items

        for r in range(t.rowCount()):
            w = t.cellWidget(r, 0)
            cb = w.findChild(QCheckBox) if w else None
            if not (cb and cb.isChecked()):
                continue
            proj = (t.item(r, code_col).text().strip() if t.item(r, code_col) else "")
            pname = (t.item(r, name_col).text().strip() if (name_col >= 0 and t.item(r, name_col)) else "")
            if proj:
                items.append((proj, pname))
        return items


    def _on_create_clicked(self):
        parent = self.parent()
        if self.worker_thread is not None or parent.api_thread is not None or parent.session.is_busy():
            return
        if not parent.api_config.get('api_key'):
            parent._open_api_settings()
            if not parent.api_config.get('api_key'):
                return
        items = []
        code_col = self._hidx.get('프로젝트코드',-1)
        name_col = self._hidx.get('프로젝트명',self._hidx.get('과제명',self._hidx.get('제목',-1)))
        if code_col < 0:
            return
        for row in range(self.tbl.rowCount()):
            holder = self.tbl.cellWidget(row,0)
            checkbox = holder.findChild(QCheckBox) if holder else None
            code_item = self.tbl.item(row,code_col)
            if not checkbox or not checkbox.isChecked() or not code_item:
                continue
            meta = dict(code_item.data(Qt.UserRole) or {})
            if meta.get('bus_done'):
                continue
            project = code_item.text().strip()
            name_item = self.tbl.item(row,name_col) if name_col >= 0 else None
            name = name_item.text().strip() if name_item else ''
            try:
                canonical_project_code(project)
            except ValueError as exc:
                QMessageBox.warning(self,'입력 확인',str(exc)); return
            items.append(dict(proj=project,name=name,row=row,**meta))
        if not items:
            return
        parent.stop_requested = False
        parent._set_running_ui(True)
        self._set_busy(True,'API 프로젝트 작업 시작')
        self.worker_thread = QThread(self)
        self.worker = CreateWorker(items,dict(parent.creds),parent._api_client(),
                                   dry_run=self.chk_dry.isChecked(),
                                   auto_complete=self.chk_auto_complete.isChecked(),
                                   requested_by=parent._requester())
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.itemResult.connect(self._on_project_result)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.finished.connect(self.worker_thread.quit,type=Qt.DirectConnection)
        self.worker_thread.finished.connect(self._project_thread_finished)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()


    def _on_project_result(self,row,result):
        col = self._hidx.get('프로젝트코드',-1)
        item = self.tbl.item(row,col) if col >= 0 else None
        if item:
            meta = dict(item.data(Qt.UserRole) or {})
            meta.update(result=result,bus_done=result.get('bus_done',False))
            item.setData(Qt.UserRole,meta)
            item.setToolTip(json.dumps(result,ensure_ascii=False,indent=2))
        self.parent()._log('프로젝트 결과: '+json.dumps(result,ensure_ascii=False,indent=2))


    def _project_thread_finished(self):
        self.worker = None
        self.worker_thread = None
        self._set_busy(False,self.status_lbl.text())
        self.parent()._set_running_ui(False)


    def closeEvent(self,event):
        if self.worker_thread is not None and self.worker_thread.isRunning():
            self.worker.stop()
            event.ignore()
        else:
            event.accept()


    def reject(self):
        if self.worker_thread is not None:
            self.worker.stop()
            return
        super().reject()


class CreateWorker(QObject):
    itemResult = pyqtSignal(int,dict)
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(int, int)
    error = pyqtSignal(str)

    def __init__(self,items,creds,client,dry_run=False,auto_complete=True,requested_by='client-operator',parent=None):
        super().__init__(parent)
        self.items,self.creds,self.client = items,creds,client
        self.dry_run,self.auto_complete,self.requested_by = dry_run,auto_complete,requested_by
        self._stop = False
        self._active_job = None


    def stop(self):
        self._stop = True
        if self._active_job is not None:
            self._active_job.stop()


    def _bus_click_process(self, proj_code: str) -> (bool, str):
        if self._stop:
            return False, '사용자 중지'
        d = None
        stage = "INIT"
        bus_msg_title = ""
        bus_msg_body = ""
        try:
            options = make_hidden_chrome_options()
            d = webdriver.Chrome(options=options)

            def normalize_text(s: str) -> str:
                s = (s or "").replace("\u00a0", " ").strip()
                return re.sub(r"\s+", " ", s)

            def set_select_value_and_fire(driver, select_id, value):
                driver.execute_script("""
                    var s = document.getElementById(arguments[0]);
                    if (s) {
                        s.value = arguments[1];
                        try { s.dispatchEvent(new Event('change', {bubbles:true})); } catch(e) {}
                        if (typeof AllBtnYn === 'function') { try { AllBtnYn(); } catch(e) {} }
                    }
                """, select_id, value)

            def click_search_manual(driver):
                try:
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "btnSearch"))
                    ).click()
                except Exception:
                    driver.switch_to.default_content()
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.ID, "btnSearch"))
                    ).click()
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_all_elements_located((By.TAG_NAME, "iframe"))
                    )
                    driver.switch_to.frame(driver.find_elements(By.TAG_NAME, "iframe")[0])

            def find_swal_popup(driver):
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

                popups = driver.find_elements(By.CSS_SELECTOR, "div.swal2-container.swal2-shown")
                if popups:
                    return popups[0]

                try:
                    frames = driver.find_elements(By.TAG_NAME, "iframe")
                except Exception:
                    frames = []

                for fr in frames:
                    try:
                        driver.switch_to.frame(fr)
                        popups = driver.find_elements(By.CSS_SELECTOR, "div.swal2-container.swal2-shown")
                        if popups:
                            return popups[0]
                    except Exception:
                        pass
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass

                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                return None

            def enter_first_iframe(driver, timeout=15, reload_url=None):
                for attempt in range(2):
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass

                    try:
                        WebDriverWait(driver, timeout).until(
                            lambda drv: len(drv.find_elements(By.TAG_NAME, "iframe")) > 0
                        )
                        frames = driver.find_elements(By.TAG_NAME, "iframe")
                        if frames:
                            driver.switch_to.frame(frames[0])
                            return True
                    except Exception:
                        pass

                    if reload_url and attempt == 0:
                        try:
                            driver.get(reload_url)
                        except Exception:
                            pass
                        time.sleep(0.5)

                return False

            def refocus_first_iframe(driver, reload_url=None):
                return enter_first_iframe(driver, timeout=5, reload_url=reload_url)

            stage = "OPEN_LOGIN_PAGE"
            d.get(BUS_LOGIN_URL)
            WebDriverWait(d, 20).until(EC.presence_of_element_located((By.ID, "windowsaccount")))

            stage = "INPUT_IDPW"
            d.find_element(By.ID, "windowsaccount").clear()
            d.find_element(By.ID, "windowsaccount").send_keys(self.creds["id"])
            d.find_element(By.ID, "password").clear()
            d.find_element(By.ID, "password").send_keys(self.creds["pw"])

            stage = "CLICK_LOGIN"
            d.find_element(By.ID, "btnLogin").click()
            WebDriverWait(d, 20).until(EC.url_contains("Common"))

            stage = "OPEN_NEW_PAGE"
            d.get(BUS_NEW_URL)

            stage = "ENTER_IFRAME"
            if not enter_first_iframe(d, reload_url=BUS_NEW_URL):
                return False, f"IFRAME_ENTER_FAIL @ {stage}"

            try:
                pre_info = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
            except Exception:
                pre_info = ""

            stage = "SET_FILTER"
            try:
                set_select_value_and_fire(d, "processYn", "N")
            except Exception:
                pass

            stage = "CLICK_SEARCH"
            if not refocus_first_iframe(d, reload_url=BUS_NEW_URL):
                return False, f"IFRAME_LOST_BEFORE_SEARCH @ {stage}"
            click_search_manual(d)

            stage = "WAIT_FILTER_APPLY"
            try:
                WebDriverWait(d, 7).until(
                    lambda x: (
                        x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                        if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else ""
                    ) != (pre_info or "")
                )
            except Exception:
                time.sleep(0.8)

            stage = "WAIT_ROWS_BEFORE"

            def wait_rows_and_sample(driver):
                WebDriverWait(driver, 20).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                )
                rows_inner = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                return rows_inner

            def find_target_row(driver, proj_code_val):
                proj_digits_val = re.sub(r"\D", "", proj_code_val or "")
                samples_inner = []
                target_inner = None
                for attempt_inner in range(3):
                    try:
                        rows_inner = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                        samples_inner = []
                        for idx, tr in enumerate(rows_inner):
                            tds = tr.find_elements(By.CSS_SELECTOR, "td")
                            full_txt = " ".join([normalize_text(td.text) for td in tds])
                            full_digits = re.sub(r"\D", "", full_txt)
                            if idx < 3:
                                samples_inner.append(full_txt or "(empty)")
                            if (proj_code_val and proj_code_val in full_txt) or (proj_digits_val and proj_digits_val in full_digits):
                                target_inner = tr
                                break
                        if target_inner is not None:
                            break
                    except StaleElementReferenceException:
                        target_inner = None
                        samples_inner = []
                        refocus_first_iframe(driver)
                        time.sleep(0.3)
                return target_inner, samples_inner, proj_digits_val

            rows = wait_rows_and_sample(d)
            if not rows:
                return False, f"ROWS_EMPTY_BEFORE @ {stage}"

            stage = "FIND_TARGET_ROW_BEFORE"
            target, samples, proj_digits = find_target_row(d, proj_code)

            if not target:
                stage = "RETRY_SEARCH_BEFORE"
                try:
                    pre_info_retry = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                except Exception:
                    pre_info_retry = ""
                if refocus_first_iframe(d, reload_url=BUS_NEW_URL):
                    click_search_manual(d)
                try:
                    WebDriverWait(d, 7).until(
                        lambda x: (
                            x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                            if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else ""
                        ) != (pre_info_retry or "")
                    )
                except Exception:
                    time.sleep(0.8)
                if not enter_first_iframe(d, reload_url=BUS_NEW_URL):
                    return False, f"IFRAME_REENTER_FAIL @ {stage}"
                rows = wait_rows_and_sample(d)
                if not rows:
                    return False, f"ROWS_EMPTY_BEFORE @ {stage}"
                stage = "FIND_TARGET_ROW_BEFORE"
                target, samples, proj_digits = find_target_row(d, proj_code)

            if not target:
                return False, f"ROW_NOT_FOUND @ {stage}: proj={proj_code}, digits={proj_digits}, sample={samples}"

            stage = "FIND_PROCESS_BUTTON"
            try:
                btn = target.find_element(By.XPATH, ".//button[contains(.,'처리')]")
            except Exception:
                return False, f"BTN_NOT_FOUND @ {stage}"

            stage = "CLICK_PROCESS_BUTTON"
            if self._stop:
                return False, '사용자 중지 (BUS 완료 미실행)'
            try:
                d.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                try:
                    btn.click()
                except ElementClickInterceptedException:
                    d.execute_script("arguments[0].click();", btn)
            except Exception as e:
                return False, f"BTN_CLICK_FAIL @ {stage}: {e}"

            stage = "WAIT_CONFIRM_POPUP"

            popup = None
            for _ in range(40):  # 최대 8초 (0.2 * 40)
                try:
                    popup = find_swal_popup(d)
                    if popup:
                        break
                except Exception:
                    popup = None
                time.sleep(0.2)

            if popup is None:
                return False, f"CONFIRM_POPUP_NOT_FOUND @ {stage}"

            stage = "READ_CONFIRM_MESSAGE"
            try:
                bus_msg_title = (popup.find_element(By.CSS_SELECTOR, ".swal2-title").text or "").strip()
            except Exception:
                bus_msg_title = ""
            try:
                bus_msg_body = (popup.find_element(By.CSS_SELECTOR, ".swal2-html-container").text or "").strip()
            except Exception:
                bus_msg_body = ""

            stage = "CLICK_CONFIRM_POPUP"
            if self._stop:
                return False, '사용자 중지 (BUS 확인 미실행)'
            try:
                ok_btn = WebDriverWait(popup, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.swal2-confirm"))
                )
                try:
                    ok_btn.click()
                except ElementClickInterceptedException:
                    d.execute_script("arguments[0].click();", ok_btn)
            except Exception as e:
                return False, f"CONFIRM_POPUP_CLICK_FAIL @ {stage}: {e}"

            stage = "WAIT_CONFIRM_CLOSE"
            try:
                WebDriverWait(d, 10).until(lambda _:
                    find_swal_popup(d) is None
                )
            except Exception:
                pass

            stage = "RELOAD_IFRAME_AFTER_PROCESS"
            if not enter_first_iframe(d, reload_url=BUS_NEW_URL):
                return False, f"IFRAME_REENTER_FAIL @ {stage}"

            try:
                pre_info2 = d.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
            except Exception:
                pre_info2 = ""

            stage = "SET_FILTER_AFTER"
            try:
                set_select_value_and_fire(d, "processYn", "N")
            except Exception:
                pass

            stage = "CLICK_SEARCH_AFTER"
            if not refocus_first_iframe(d, reload_url=BUS_NEW_URL):
                return False, f"IFRAME_LOST_AFTER_SEARCH @ {stage}"
            click_search_manual(d)

            stage = "WAIT_FILTER_APPLY_AFTER"
            try:
                WebDriverWait(d, 7).until(
                    lambda x: (
                        x.find_element(By.CSS_SELECTOR, ".dataTables_info").text.strip()
                        if x.find_elements(By.CSS_SELECTOR, ".dataTables_info") else ""
                    ) != (pre_info2 or "")
                )
            except Exception:
                time.sleep(0.8)

            stage = "WAIT_ROWS_AFTER"
            try:
                WebDriverWait(d, 20).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                )
                rows2 = d.find_elements(By.CSS_SELECTOR, "tbody tr")
            except Exception:
                rows2 = []

            stage = "VERIFY_ROW_REMOVED"
            proj_digits = re.sub(r"\D", "", proj_code or "")
            still_exists = False
            for tr in rows2:
                tds = tr.find_elements(By.CSS_SELECTOR, "td")
                full_txt = " ".join([normalize_text(td.text) for td in tds])
                full_digits = re.sub(r"\D", "", full_txt)
                if (proj_code and proj_code in full_txt) or (proj_digits and proj_digits in full_digits):
                    still_exists = True
                    break

            if still_exists:
                if bus_msg_title or bus_msg_body:
                    return False, f"ROW_STILL_EXISTS_AFTER_PROCESS @ {stage}: title={bus_msg_title}, body={bus_msg_body}"
                return False, f"ROW_STILL_EXISTS_AFTER_PROCESS @ {stage}"

            if bus_msg_title or bus_msg_body:
                return True, f"OK @ {stage}: title={bus_msg_title}, body={bus_msg_body}"
            return True, f"OK @ {stage}"

        except Exception as e:
            return False, f"EXCEPTION @ {stage}: {e}"

        finally:
            if d is not None:
                try:
                    d.quit()
                except Exception:
                    pass


    def run(self):
        ok_count = fail_count = 0
        try:
            for index,item in enumerate(self.items):
                if self._stop:
                    break
                self.progress.emit(index,len(self.items),f"{item['proj']} API 작업 중")
                source = item.get('source','test-client')
                request_id = item.get('request_id') or new_client_request_id()
                payload = make_project_payload(item['proj'],item.get('name',''),request_id,source,self.requested_by)
                cached = item.get('result',{})
                if not self.dry_run and can_complete_bus(cached,source) and not item.get('bus_done'):
                    result = dict(cached)
                else:
                    self._active_job = ApiJobWorker(self.client,'project',payload,self.dry_run)
                    if self._stop:
                        self._active_job.stop()
                    try:
                        result = self._active_job.execute()
                    except Exception as exc:
                        result = dict(status='failed',error_message=str(exc))
                    self._active_job = None
                result = dict(result)
                real_success = result.get('status')=='succeeded' and result.get('executor_mode')=='powershell'
                if real_success:
                    ok_count += 1
                elif result.get('status') not in {'preview','simulated'}:
                    fail_count += 1
                message = f"{item['proj']} API {result.get('status')}: {result.get('error_message') or ''}"
                if self.auto_complete and can_complete_bus(result,source,self._stop or result.get('client_stopped',False)):
                    if not self.creds.get('id') or not self.creds.get('pw'):
                        result['bus_done'] = False
                        message += ' / BUS 계정 누락: 완료만 재시도 가능'
                    else:
                        bus_ok,bus_message = self._bus_click_process(item['proj'])
                        result['bus_done'] = bus_ok
                        message += ' / BUS '+bus_message
                self.itemResult.emit(item.get('row',index),result)
                self.progress.emit(index+1,len(self.items),message)
        except Exception as exc:
            message = str(exc)
            if self.client.api_key:
                message = message.replace(self.client.api_key,'[API Key]')
            self.error.emit(message)
            fail_count += 1
        finally:
            self._active_job = None
            self.finished.emit(ok_count,fail_count)


class BadgeToolButton(QWidget):
    def __init__(self, *args, **kwargs):
        target = None
        parent = None
        if len(args) == 1:
            parent = args[0]
        elif len(args) >= 2:
            if isinstance(args[0], QAbstractButton):
                target = args[0]
                parent = args[1]
            else:
                parent = args[1]
        else:
            parent = kwargs.get("parent", None)

        super().__init__(parent)
        self._target = None
        self._offset = QPoint(0, 1)
        self._value = 0
        self._lbl = QLabel(self)
        self._lbl.setObjectName("notifBadge")
        self._lbl.setAlignment(Qt.AlignCenter)
        f = self._lbl.font()
        f.setPointSize(7)
        f.setBold(True)
        self._lbl.setFont(f)
        self._base_px = 7
        self._apply_style(self._base_px)
        self._lbl.setVisible(False)

        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        self._lbl.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        if target is not None:
            self.setTarget(target)

    def setTarget(self, btn: QAbstractButton):
        if self._target is not None:
            try:
                self._target.removeEventFilter(self)
            except Exception:
                pass
        self._target = btn
        self._lbl.setParent(self._target)
        self._lbl.raise_()
        if self._target is not None:
            self._target.installEventFilter(self)
        self._reposition()

    def setBadge(self, n: int):
        self._value = int(n or 0)
        if self._value <= 0:
            self._lbl.hide()
            return
        self._lbl.setText("99+" if self._value >= 100 else str(self._value))
        self._lbl.adjustSize()
        self._lbl.show()
        self._lbl.raise_()
        self._reposition()

    def setOffset(self, dx: int, dy: int):
        self._offset = QPoint(dx, dy)
        self._reposition()

    def setBadgeSize(self, px: int = 1):
        self._base_px = px
        self._apply_style(px)
        self._lbl.adjustSize()
        self._reposition()

    def eventFilter(self, obj, ev):
        if obj is self._target and ev.type() in (ev.Resize, ev.Move, ev.Show):
            self._reposition()
        return super().eventFilter(obj, ev)

    def _reposition(self):
        if not (self._target and self._lbl.isVisible()):
            return
        bw = max(self._base_px, self._lbl.width())
        x = self._target.width() - bw + self._offset.x()
        y = self._offset.y()
        self._lbl.move(x, y)

    def _apply_style(self, px: int):
        r = px // 2
        pad = max(1, px // 4)
        maxw = px * 3

        self._lbl.setStyleSheet(f"""
            QLabel#notifBadge {{
                color: white;
                background: #E23;
                border-radius: {r}px;
                min-width: {px}px;
                min-height: {px}px;
                max-width: {maxw}px;
                padding-left: {pad}px;
                padding-right: {pad}px;
            }}
        """)

class AccessManager(QMainWindow):
    trigger_new_download = pyqtSignal()
    trigger_session_start = pyqtSignal()
    trigger_session_download = pyqtSignal()
    trigger_session_cancel = pyqtSignal()
    trigger_session_stop = pyqtSignal()
    trigger_session_process = pyqtSignal(object)
    trigger_watcher_start = pyqtSignal()
    trigger_watcher_stop  = pyqtSignal()
    trigger_watcher_collect = pyqtSignal()
    COL_SELECT  = 0
    COL_KIND    = 1
    COL_REQTYPE = 2
    COL_USER    = 3
    COL_NAME    = 4
    COL_PROJ    = 5
    COL_LV2     = 6
    COL_LV3     = 7
    COL_DEPT    = 8
    COL_ROLE    = 9
    COL_STATUS  = 10

    EDITABLE_COLS = {COL_KIND, COL_REQTYPE, COL_USER, COL_NAME, COL_PROJ, COL_LV2, COL_LV3, COL_DEPT, COL_ROLE}

    def _show_help(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("기능 설명")
        edit = QTextEdit(dlg)
        edit.setReadOnly(True)
        edit.setPlainText(HELP_TEXT)
        btns = QDialogButtonBox(QDialogButtonBox.Close, parent=dlg)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        lay = QVBoxLayout(dlg)
        lay.addWidget(edit)
        lay.addWidget(btns)
        dlg.resize(720, 560)
        dlg.exec_()

    def apply_theme(self, mode: str):
        theme = THEMES.get(mode, THEMES["light"])
        text_color = "#111111" if mode == "light" else "#E0E0E0"
        disabled_bg = "rgba(176,190,197,0.45)" if mode == "light" else "rgba(66,66,66,0.50)"
        disabled_tx = "#9aa3a8"
        disabled_border = "#c7c7c7" if mode == "light" else "#5a5a5a"

        qss = f"""
        QWidget {{
            background-color: {theme['bg']};
            color: {text_color};
        }}

        QPushButton {{
            background-color: {theme['btn']};
            border: 1px solid #9e9e9e;
            border-radius: 6px;
            padding: 6px 12px;
            color: {text_color};
        }}
        QPushButton:hover {{
            background-color: {theme['hover']};
        }}
        QPushButton:pressed {{
            background-color: {theme['press']};
        }}

        QPushButton:disabled,
        QPushButton:disabled:hover,
        QPushButton:disabled:pressed {{
            background-color: rgba(176,190,197,0.45);
            color: #9aa3a8;
            border: 1px solid {theme['panel_border']};
        }}

        QToolButton:disabled {{
            color: #9aa3a8;
        }}

        QLineEdit:disabled,
        QComboBox:disabled,
        QCheckBox:disabled,
        QTextEdit:disabled {{
            color: #9aa3a8;
            background-color: transparent;
        }}

        QHeaderView::section {{
            background-color: {theme['btn']};
            color: {text_color};
            border: 0px;
            padding: 4px 6px;
        }}

        QHeaderView::section:vertical {{
            background-color: {theme['btn']};
            color: {text_color};
            border: 0px;
            padding: 0px 6px;
            border-radius: 0px;
        }}

        QTableCornerButton::section {{
            background-color: {theme['btn']};
            border: 0px;
            padding: 0px;
            border-top-left-radius: 8px;
        }}

        QTableView, QTableWidget, QTextEdit {{
            background: {theme['panel']};
            border: 1px solid {theme['panel_border']};
            border-radius: 8px;
            selection-background-color: {theme['select']};
            selection-color: {text_color};
        }}

        QTableView, QTableWidget {{
            gridline-color: {theme['panel_border']};
            alternate-background-color: {theme['alt']};
        }}

        QTableWidget::item, QTableView::item {{
            padding: 4px 6px;
        }}

        QAbstractScrollArea {{
            background: transparent;
        }}

        QMenu {{
            background-color: {theme['bg']};
            color: {text_color};
            border: 1px solid #9e9e9e;
            border-radius: 6px;
            padding: 4px;
        }}
        QMenu::separator {{
            height: 1px;
            background: rgba(0,0,0,0.15);
            margin: 4px 8px;
        }}
        QMenu::item {{
            padding: 6px 12px;
            border-radius: 4px;
            background: transparent;
        }}
        QMenu::item:selected {{
            background: {theme['hover']};
            color: {text_color};
        }}

        QComboBox QAbstractItemView {{
            background-color: {theme['bg']};
            color: {text_color};
            selection-background-color: {theme['hover']};
            selection-color: {text_color};
            outline: 0;
        }}

        QHeaderView::section:horizontal:last {{
            border-top-right-radius: 8px;
        }}

        QHeaderView::section:horizontal:first {{
            border-top-left-radius: 0px;
        }}
        """
        self.setStyleSheet(qss)
        self.current_theme = mode
        self._refresh_theme_button_emoji()

        if hasattr(self, "btn_theme") and self.btn_theme:
            self.btn_theme.setText("🌙" if mode == "light" else "🌞")
            self.btn_theme.setToolTip("다크 모드로 전환" if mode == "light" else "라이트 모드로 전환")

        if hasattr(self, "table") and self.table and hasattr(self.table.horizontalHeader(), "_reposition"):
            self.table.horizontalHeader()._reposition()
        
    def _toggle_theme(self):
        next_mode = "dark" if getattr(self, "current_theme", "light") == "light" else "light"
        self.apply_theme(next_mode)
        try:
            data = {}
            if os.path.exists(CONF_FILE):
                with open(CONF_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            data["theme"] = next_mode
            with open(CONF_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def __init__(self):
        super().__init__()
        for folder in (DL_DIR, LOG_DIR, CONF_DIR, DEBUG_DIR):
            os.makedirs(folder, exist_ok=True)
        self._notif_counts = {"신규미완료": 0, "진행-부여": 0, "진행-제거": 0, "종료-부여": 0, "종료-제거": 0}
        self.setWindowTitle(f"{APP_NAME}")
        self.resize(900, 600)
        self._init_ui()
        self._notify_refresh_mins = 10
        try:
            if os.path.exists(CONF_FILE):
                with open(CONF_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self._notify_refresh_mins = int(data.get("notify_refresh_min", self._notify_refresh_mins))
        except Exception:
            pass
        self.watch_thread = QThread(self)
        self.watch_session = BusWatcher(DL_DIR)
        self.watch_session.moveToThread(self.watch_thread)
        self.watch_thread.start()
        self.trigger_watcher_start.connect(self.watch_session.start, type=Qt.QueuedConnection)
        self.trigger_watcher_stop.connect(self.watch_session.stop, type=Qt.QueuedConnection)
        self.trigger_watcher_collect.connect(self.watch_session.collect_counts, type=Qt.QueuedConnection)
        self.watch_session.countsReady.connect(self._on_counts_ready)
        self.notify_timer = QTimer(self)
        self.notify_timer.setInterval(max(5, int(self._notify_refresh_mins)) * 60 * 1000)
        self.notify_timer.timeout.connect(self.refresh_notifications)
        self.notify_timer.start()
        self._ensure_logo_and_set_title()
        self.session_thread = QThread(self)
        self.session = BusSessionManager(DL_DIR)
        self.session.moveToThread(self.session_thread)
        self.session.readyChanged.connect(self._on_session_ready)
        self.session.downloaded.connect(self._on_session_downloaded)
        self.session.processed.connect(self._on_session_processed)
        self.session.busyChanged.connect(self._on_session_busy)
        self.session_thread.start()
        self.watch_session.readyChanged.connect(self._on_watcher_ready)
        self.trigger_session_start.connect(self.session.start, type=Qt.QueuedConnection)
        self.trigger_session_download.connect(self.session.download_list, type=Qt.QueuedConnection)
        self.trigger_new_download.connect(self.session.download_new_list,type=Qt.QueuedConnection)
        self.trigger_session_process.connect(self.session.process, type=Qt.QueuedConnection)
        self.trigger_session_cancel.connect(self.session.cancel_current, type=Qt.QueuedConnection)
        self.trigger_session_stop.connect(self.session.stop, type=Qt.QueuedConnection)
        QTimer.singleShot(1000, self.refresh_notifications)
        self.api_config = load_api_config()
        self.api_thread = None
        self.api_worker = None
        self._closing = False
        self._shutdown_started = False
        QTimer.singleShot(0,self._show_initial_api_settings)
        self.run_queue = []
        self.total_jobs = 0
        self.done_jobs = 0
        self.current_row = -1
        self.current_seq = 0
        self.current_mode = ""
        self.setAcceptDrops(True)
        self._load_creds()
        if self.creds.get("id") and self.creds.get("pw"):
            self.watch_session.set_creds(self.creds["id"], self.creds["pw"])
            self.trigger_watcher_start.emit()
        self._log(f"{APP_NAME} 실행")
        self._buf_out = {}
        self.remove_fail_tolerance = int(self.creds.get("remove_fail_tol", 5))
        self.auto_complete_after_add = True
        self._waiting_for_bus = False
        self._pending_after_add_row = None
        self.bus_seq = 0
        self._ignore_bus_results = False
        self.stop_requested = False
        self._auto_create_group_decided = None
        self._retrying_after_group_create = False
        self._bus_mode = False
        self._bus_queue = []
        self._bus_done = 0
        self._bus_total = 0

        saved_theme = self._read_saved_theme()
        self.current_theme = saved_theme or "light"
        self.apply_theme(self.current_theme)

        self.debug_enabled = bool(self.creds.get("debug", False))
        self.session.set_debug(self.debug_enabled, DEBUG_DIR)
        
    def _init_ui(self):
        main_layout = QVBoxLayout()
        header_bar = QHBoxLayout()
        self.header_bar = header_bar
        header_bar.setContentsMargins(0, 0, 0, 0)
        header_bar.setSpacing(6)

        header_bar.addStretch()
        self._title_widget = QLabel("")
        header_bar.addWidget(self._title_widget, 0, Qt.AlignVCenter)
        header_bar.addStretch()

        self.btn_notif = QToolButton(self)
        self.btn_notif.setText("💡")
        self.btn_notif.setToolTip("요청건 알림")
        self.btn_notif.setAutoRaise(True)
        self.btn_notif.setStyleSheet("""
            QToolButton {
                border: none;
                background: transparent;
                font-size: 15px;
                padding: 0px;
            }
            QToolButton:hover {
                background: transparent;    /* 마우스 올렸을 때 배경 투명 */
            }
            QToolButton:pressed {
                background: transparent;    /* 클릭 중에도 투명 */
            }
            QToolButton:checked {
                background: transparent;    /* 체크 상태도 투명 */
            }
            QToolButton:focus {
                outline: none;              /* 포커스 테두리 제거 */
            }
        """)

        self.btn_notif.clicked.connect(self._open_notif_popup)
        self.badge_notif = BadgeToolButton(self.btn_notif, self)
        self.badge_notif.setBadgeSize(9)
        self.badge_notif.setOffset(0,0)

        self.btn_theme = QToolButton(self)
        self.btn_theme.setAutoRaise(True)
        self.btn_theme.setCursor(Qt.PointingHandCursor)
        self.btn_theme.setToolTip("테마 전환")
        self.btn_theme.setStyleSheet("""
            QToolButton {
                border: none;
                background: transparent;
                font-size: 15px;
                padding: 0 6px;
            }
            QToolButton:hover {
                background: transparent;
            }
        """)
        header_bar.addWidget(self.btn_theme, 0, Qt.AlignRight | Qt.AlignVCenter)
        header_bar.addWidget(self.btn_notif, 0, Qt.AlignRight | Qt.AlignVCenter)
        header_w = QWidget()
        header_w.setLayout(header_bar)
        main_layout.addWidget(header_w)

        file_bar = QHBoxLayout()
        #self.file_label = QLabel("요청 확인 버튼 클릭 또는 엑셀 파일 불러오기, 수동으로 리스트 입력")
        #self.file_label.setFont(QFont("Segoe UI", 9))

        self.btn_newcheck = QPushButton("𝙉 신규 확인")
        self.btn_newcheck.setFont(QFont("Segoe UI", 9))
        self.btn_newcheck.clicked.connect(self.open_new_viewer) 

        self.btn_request = QPushButton("🔍 요청 확인")
        self.btn_request.setFont(QFont("Segoe UI", 9))
        self.btn_request.clicked.connect(self._request_and_import)

        self.btn_manual = QPushButton("📝 수동 입력")
        self.btn_manual.setFont(QFont("Segoe UI", 9))
        self.btn_manual.clicked.connect(self.open_manual_dialog)
        
        self.btn_file = QPushButton("📂 파일 선택")
        self.btn_file.setFont(QFont("Segoe UI", 9))
        self.btn_file.clicked.connect(self.choose_file)

        self.btn_settings = QPushButton("⚙ 설정")
        self.btn_settings.setFont(QFont("Segoe UI", 9))
        self.btn_settings.clicked.connect(self._open_settings)
        self.btn_api_settings = QPushButton('API 설정')
        self.btn_api_settings.clicked.connect(self._open_api_settings)
        
        #file_bar.addWidget(self.file_label)
        file_bar.addStretch()
        file_bar.addWidget(self.btn_newcheck)
        file_bar.addWidget(self.btn_request)
        file_bar.addWidget(self.btn_manual)
        file_bar.addWidget(self.btn_file)
        file_bar.addWidget(self.btn_settings)
        file_bar.addWidget(self.btn_api_settings)
        main_layout.addLayout(file_bar)

        self.table = CopyTable()
        self.table.setProperty("hasRows", False)
        header = CheckBoxHeader(Qt.Horizontal, self.table)
        self.table.setHorizontalHeader(header)
        header.stateChanged.connect(self.toggle_all_rows)

        self.table.setFont(QFont("Consolas", 9))
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "", "구분", "요청사항", "사번", "이름", "프로젝트코드","Level2","Level3","부서","Role","상태"
        ])
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed)
        self.table.cellChanged.connect(self._on_cell_changed)

        hv = self.table.horizontalHeader()
        hv.setSectionResizeMode(QHeaderView.ResizeToContents)
        hv.setStretchLastSection(True)

        hv.setSectionResizeMode(self.COL_SELECT, QHeaderView.Fixed)
        self.table.setColumnWidth(self.COL_SELECT, 28)
        hv.setMinimumSectionSize(20)

        vh = self.table.verticalHeader()
        vh.setSectionResizeMode(QHeaderView.Fixed)
        vh.setDefaultSectionSize(28)
        vh.setMinimumSectionSize(22)
        vh.setFixedWidth(28)
        self.table.setCornerButtonEnabled(True)

        header.sectionResized.connect(lambda *_: header.updateSection(self.COL_SELECT))
        try:
            header.geometriesChanged.connect(lambda: header.updateSection(self.COL_SELECT))
        except Exception:
            pass
        
        main_layout.addWidget(self.table)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._open_table_menu)
        
        btn_bar = QHBoxLayout()

        self.btn_clear_log = QPushButton("🗑 로그 지우기")
        self.btn_clear_log.setFont(QFont("Segoe UI", 9))
        self.btn_clear_log.setToolTip("아래 로그 창 내용을 비웁니다")
        self.btn_clear_log.clicked.connect(self._clear_log)

        self.chk_dry = QCheckBox("Dry Run")
        self.chk_dry.setFont(QFont("Segoe UI", 9))

        self.chk_auto_complete = QCheckBox("실행 후 자동 완료처리")
        self.chk_auto_complete.setFont(QFont("Segoe UI", 9))
        self.chk_auto_complete.setChecked(True)
        self.chk_auto_complete.stateChanged.connect(
            lambda _: setattr(self, "auto_complete_after_add", self.chk_auto_complete.isChecked())
        )

        self.btn_run_execute = QPushButton("> 실행")
        self.btn_run_execute.setFont(QFont("Segoe UI", 9))
        self.btn_run_execute.clicked.connect(self.run_execute)

        self.btn_run_complete = QPushButton("v 완료 처리")
        self.btn_run_complete.setFont(QFont("Segoe UI", 9))
        self.btn_run_complete.clicked.connect(self.run_complete)

        self.btn_stop = QPushButton("x 중지")
        self.btn_stop.setFont(QFont("Segoe UI", 9))
        self.btn_stop.setToolTip("현재 실행 중인 작업과 대기 중인 모든 작업을 중지합니다")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self._stop_all)

        self.btn_theme.clicked.connect(self._toggle_theme)

        btn_bar.addWidget(self.btn_clear_log)
        btn_bar.addStretch()
        btn_bar.addWidget(self.chk_dry)
        btn_bar.addWidget(self.chk_auto_complete)
        btn_bar.addWidget(self.btn_run_execute)
        btn_bar.addWidget(self.btn_run_complete)
        btn_bar.addWidget(self.btn_stop)
        main_layout.addLayout(btn_bar)
        
        self.log = QTextEdit()
        self.log.setAcceptRichText(False)
        self.log.setLineWrapMode(QTextEdit.WidgetWidth)
        self.log.setWordWrapMode(QTextOption.WrapAnywhere)
        self.log.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.log.setFont(QFont("Consolas", 9))
        self.log.setReadOnly(True)
        main_layout.addWidget(self.log)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        sb = super().statusBar()
        if sb is None:
            sb = QStatusBar(self)
            self.setStatusBar(sb)
        self._statusbar = sb
        self._statusbar.setSizeGripEnabled(False)

        self.prg = QProgressBar()
        self.prg.setRange(0,100)
        self.prg.setValue(0)
        self.prg.setTextVisible(False)
        self.prg.setFixedSize(0, 0)
        self.prg.hide()
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.status_label.setMinimumWidth(420)
        self.status_label.setStyleSheet("padding-right: 12px; padding-bottom: 2px;")
        self.progress_panel = QWidget()
        self.progress_panel.setVisible(False)
        progress_layout = QVBoxLayout(self.progress_panel)
        progress_layout.setContentsMargins(0, 0, 12, 2)
        progress_layout.setSpacing(0)
        progress_layout.addWidget(self.status_label, 0, Qt.AlignRight | Qt.AlignVCenter)
        self._statusbar.addPermanentWidget(self.progress_panel, 1)
        self._progress_started_at = None
        self._progress_detail = ""
        self._progress_project = ""
        self._progress_done = 0
        self._progress_total = 0
        self._progress_running = False
        self._progress_timer = QTimer(self)
        self._progress_timer.setInterval(1000)
        self._progress_timer.timeout.connect(self._refresh_progress_status)

    def _format_elapsed(self) -> str:
        if not getattr(self, "_progress_started_at", None):
            return "00:00"
        elapsed = max(0, int(time.time() - self._progress_started_at))
        minutes, seconds = divmod(elapsed, 60)
        hours, minutes = divmod(minutes, 60)
        if hours:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def _set_progress_status(self, detail: str, project: str = "", done: int | None = None,
                             total: int | None = None, running: bool = True):
        if done is not None:
            self._progress_done = max(0, int(done))
        if total is not None:
            self._progress_total = max(0, int(total))
        self._progress_detail = detail or ""
        self._progress_project = project or ""
        self._progress_running = bool(running)

        if running and not getattr(self, "_progress_started_at", None):
            self._progress_started_at = time.time()
        if running and not self._progress_timer.isActive():
            self._progress_timer.start()
        self._refresh_progress_status()

    def _refresh_progress_status(self):
        total = max(0, int(getattr(self, "_progress_total", 0)))
        done = max(0, int(getattr(self, "_progress_done", 0)))
        running = bool(getattr(self, "_progress_running", False))
        detail = getattr(self, "_progress_detail", "") or "처리 중"
        project = getattr(self, "_progress_project", "") or ""

        if total > 0:
            if running:
                current = min(total, done + 1)
                prefix = f"처리 중 {current}/{total}"
                pct = int((current / total) * 100)
            else:
                current = min(total, done)
                prefix = f"완료 {current}/{total}" if done >= total else f"중지됨 {current}/{total}"
                pct = int((current / total) * 100)
        else:
            prefix = detail
            pct = 0 if running else 100

        parts = [prefix]
        if project:
            parts.append(f"프로젝트 {project}")
        if detail and detail != prefix:
            parts.append(detail)
        if running or getattr(self, "_progress_started_at", None):
            parts.append(self._format_elapsed())

        self.status_label.setText(" · ".join(parts))
        self.prg.setValue(max(0, min(100, pct)))
        self.progress_panel.setVisible(running or bool(getattr(self, "_progress_started_at", None)))
        self.prg.hide()

    def _finish_progress_status(self, detail: str = "완료", stopped: bool = False):
        self._progress_running = False
        if self._progress_timer.isActive():
            self._progress_timer.stop()
        if not stopped and self._progress_total:
            self._progress_done = self._progress_total
        self._progress_detail = detail or ("중지됨" if stopped else "완료")
        self._refresh_progress_status()

    def _hide_progress_status(self):
        self._progress_running = False
        self._progress_started_at = None
        if self._progress_timer.isActive():
            self._progress_timer.stop()
        self.status_label.setText("")
        self.prg.setValue(0)
        self.progress_panel.setVisible(False)
        self.prg.hide()

    def _set_plain_status(self, message: str):
        self._progress_running = False
        self._progress_started_at = None
        if self._progress_timer.isActive():
            self._progress_timer.stop()
        self.status_label.setText(message or "")
        self.progress_panel.setVisible(bool(message))
        self.prg.hide()

    def refresh_notifications(self):
        if hasattr(self, "watch_session") and self.watch_session and self.watch_session.is_ready():
            self.trigger_watcher_collect.emit()
            return
        self._on_counts_ready({"신규미완료":0,"진행-부여":0,"진행-제거":0,"종료-부여":0,"종료-제거":0})

    @pyqtSlot(dict)
    def _on_counts_ready(self, counts: dict):
        counts = counts or {}
        def _toi(x):
            try:
                return int(str(x).strip())
            except Exception:
                return 0

        self._notif_counts = {
            "신규미완료": _toi(counts.get("신규미완료", 0)),
            "진행-부여": _toi(counts.get("진행-부여", 0)),
            "진행-제거": _toi(counts.get("진행-제거", 0)),
            "종료-부여": _toi(counts.get("종료-부여", 0)),
            "종료-제거": _toi(counts.get("종료-제거", 0)),
        }

        total = sum(self._notif_counts.values())

        QTimer.singleShot(0, lambda t=total: self.badge_notif.setBadge(t))

    def _on_watcher_ready(self, ok: bool, msg: str):
        if hasattr(self, "btn_notif"):
            self.btn_notif.setEnabled(bool(ok))
            self.btn_notif.setToolTip("요청건 알림" if ok else "세션 준비 안됨")

        if ok:
            QTimer.singleShot(0, lambda: self.trigger_watcher_collect.emit())
        else:
            self._on_counts_ready({"신규미완료":0,"진행-부여":0,"진행-제거":0,"종료-부여":0,"종료-제거":0})

    def _open_notif_popup(self):
        if not hasattr(self, "_notif_counts") or not isinstance(self._notif_counts, dict):
            self._notif_counts = {"신규미완료":0,"진행-부여":0,"진행-제거":0,"종료-부여":0,"종료-제거":0}
        c = self._notif_counts
        menu = QMenu(self)
        n_new  = int(self._notif_counts.get("신규미완료", 0))
        n_pg_g = int(self._notif_counts.get("진행-부여", 0))
        n_pg_r = int(self._notif_counts.get("진행-제거", 0))
        n_ed_g = int(self._notif_counts.get("종료-부여", 0))
        n_ed_r = int(self._notif_counts.get("종료-제거", 0))
        menu.addAction(f"폴더 생성 요청: {n_new}")
        menu.addAction(f"진행-부여: {n_pg_g}")
        menu.addAction(f"진행-제거: {n_pg_r}")
        menu.addAction(f"종료-부여: {n_ed_g}")
        menu.addAction(f"종료-제거: {n_ed_r}")
        pos = self.btn_notif.mapToGlobal(self.btn_notif.rect().bottomRight())
        menu.popup(pos)

    def open_new_viewer(self):
        def _after_ready(ok: bool):
            if not ok:
                return
            self._newdlg = NewItemsViewer(self)
            self._newdlg.btn_refresh.clicked.connect(self._load_new_items)
            self._newdlg.show()
            self._load_new_items()

        self.ensure_bus_session_async("폴더 생성 요청 조회", _after_ready)

    def _load_new_items(self):
        if self.api_thread is not None or self.session.is_busy():
            return
        try:
            self.session.newDownloaded.disconnect(self._on_new_downloaded)
        except Exception:
            pass
        self.session.newDownloaded.connect(self._on_new_downloaded)
        self._set_running_ui(True)
        self.trigger_new_download.emit()


    def _on_new_downloaded(self, path: str, err: str):
        self._set_running_ui(False)
        if err:
            self._set_plain_status("폴더 생성 요청 로드 실패")
            QMessageBox.critical(self, "오류", f"폴더 생성 요청 로드 실패: {err}")
            return
        header, data = _parse_html_best_table(path)
        if not header:
            self._set_plain_status("표 데이터 없음")
            QMessageBox.warning(self, "안내", "표 데이터를 찾지 못했습니다.")
            return
        if getattr(self, "_newdlg", None):
            self._newdlg.set_data(header, data)
        self._set_plain_status(f"폴더 생성 요청 {len(data)}건")


    def _refresh_has_rows(self):
        has_rows = self.table.rowCount() > 0

        self.table.setProperty("hasRows", has_rows)
        hh = self.table.horizontalHeader()
        vh = self.table.verticalHeader()
        for w in (self.table, hh, vh):
            try:
                w.setProperty("hasRows", has_rows)
                st = w.style()
                st.unpolish(w); st.polish(w)
                w.update()
            except Exception:
                pass

    def _refresh_theme_button_emoji(self):
        if hasattr(self, "btn_theme") and self.btn_theme:
            self.btn_theme.setText("🌙" if self.current_theme == "light" else "🌞")
            self.btn_theme.setToolTip("다크 모드로 전환" if self.current_theme == "light" else "라이트 모드로 전환")

    def _read_saved_theme(self) -> str:
        try:
            if os.path.exists(CONF_FILE):
                with open(CONF_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    t = (data.get("theme") or "").strip().lower()
                    if t in ("light", "dark"):
                        return t
        except Exception:
            pass
        return "light"

    def _ensure_logo_and_set_title(self):
        try:
            os.makedirs(CONF_DIR, exist_ok=True)

            target_logo = os.path.join(CONF_DIR, LOGO_FILE)
            if not os.path.exists(target_logo):
                bundled_logo = resource_path(os.path.join("assets", LOGO_FILE))
                if os.path.exists(bundled_logo):
                    shutil.copyfile(bundled_logo, target_logo)

            target_icon = os.path.join(CONF_DIR, ICON_FILE)
            if not os.path.exists(target_icon):
                bundled_icon = resource_path(os.path.join("assets", ICON_FILE))
                if os.path.exists(bundled_icon):
                    shutil.copyfile(bundled_icon, target_icon)

            self._apply_title_logo(target_logo if os.path.exists(target_logo) else None)

            if os.path.exists(target_icon):
                app = QApplication.instance()
                icon = QIcon(target_icon)
                self.setWindowIcon(icon)
                if app:
                    app.setWindowIcon(icon)

        except Exception as e:
            self._apply_title_logo(None)
            self._log(f"[로고/아이콘 준비 오류] {e}")

    def _apply_title_logo(self, logo_path: str | None):
        try:
            if hasattr(self, "_title_widget") and self._title_widget is not None:
                self._title_widget.setParent(None)
        except Exception:
            pass

        if logo_path and os.path.exists(logo_path):
            lbl = QLabel()
            pm = QPixmap(logo_path)
            if not pm.isNull():
                pm = pm.scaledToHeight(48, Qt.SmoothTransformation)
                lbl.setPixmap(pm)
            else:
                lbl.setText(APP_NAME)
                lbl.setFont(QFont("Segoe UI", 18, QFont.Bold))
            lbl.setCursor(Qt.PointingHandCursor)
            lbl.mousePressEvent = lambda e: self._show_help()
            widget = lbl
        else:
            txt = QLabel(APP_NAME)
            txt.setFont(QFont("Segoe UI", 18, QFont.Bold))
            txt.setAlignment(Qt.AlignCenter)
            txt.setCursor(Qt.PointingHandCursor)
            txt.mousePressEvent = lambda e: self._show_help()
            widget = txt

        self.header_bar.insertWidget(1, widget, 0, Qt.AlignVCenter)
        self._title_widget = widget

    def _ensure_bus_session(self, purpose: str = "완료 처리") -> bool:
        if not self.creds.get("id") or not self.creds.get("pw"):
            self._open_settings()
            if not self.creds.get("id") or not self.creds.get("pw"):
                QMessageBox.warning(self, "알림", "아이디/비밀번호가 필요합니다.")
                return False

        self.session.set_creds(self.creds["id"], self.creds["pw"])

        if self.session.is_ready():
            return True

        self._set_running_ui(True)
        self.status_label.setText(f"BUS 세션 준비 중 ({purpose})")
        self.trigger_session_start.emit()

        end = time.time() + 60
        while time.time() < end:
            QApplication.processEvents()
            if self.session.is_ready():
                self.status_label.setText("세션 준비됨")
                return True
            time.sleep(0.05)

        self._set_running_ui(False)
        QMessageBox.critical(self, "오류", "세션 초기화 실패")
        return False

    def ensure_bus_session_async(self, purpose: str, on_ready):
        if not self.creds.get("id") or not self.creds.get("pw"):
            self._open_settings()
            if not self.creds.get("id") or not self.creds.get("pw"):
                QMessageBox.warning(self, "알림", "아이디/비밀번호가 필요합니다.")
                if on_ready:
                    try:
                        on_ready(False)
                    except Exception:
                        pass
                return

        self.session.set_creds(self.creds["id"], self.creds["pw"])

        if self.session.is_ready():
            if on_ready:
                try:
                    on_ready(True)
                except Exception:
                    pass
            return

        if not hasattr(self, "_pending_bus_callbacks"):
            self._pending_bus_callbacks = []
        if on_ready:
            self._pending_bus_callbacks.append(on_ready)

        self._set_running_ui(True)
        self.status_label.setText(f"BUS 세션 준비 중 ({purpose})")
        self.trigger_session_start.emit()

    def _start_next_bus_item(self):
        if self.stop_requested or not self._bus_queue:
            self._bus_mode = False
            self._finish_progress_status("완료" if not self.stop_requested else "중지됨", stopped=self.stop_requested)
            self._set_running_ui(False)
            self.stop_requested = False
            return

        t = self._bus_queue.pop(0)
        row = t.get('row', -1)
        self._set_progress_status("BUS 완료 처리 중", t.get('proj', ''), self._bus_done, self._bus_total, running=True)
        if 0 <= row < self.table.rowCount():
            it = self.table.item(row, self.COL_STATUS)
            if it:
                it.setText("완료 처리중")

        self.trigger_session_process.emit([t])

    def run_complete(self):
        if self.api_thread is not None or self.session.is_busy():
            return
        self.stop_requested = False
        self._ignore_bus_results = False
        targets = []
        for row in range(self.table.rowCount()):
            if not self._is_row_checked(row):
                continue
            meta = self._row_metadata(row)
            result = meta.get('result',{})
            if (not meta.get('bus_done') and can_complete_bus(result,meta.get('source'),result.get('client_stopped',False))):
                targets.append(self._bus_target(row))
            else:
                self._log(f'{row+1}행: BUS 요청의 실제 API 성공 결과가 없어 완료 처리에서 제외했습니다.')
        if not targets or not self._ensure_bus_session('완료 처리'):
            return
        self._waiting_for_bus = False
        self._bus_mode = True
        self._bus_queue = targets
        self._bus_total = len(targets)
        self._bus_done = 0
        self._progress_started_at = time.time()
        self._set_running_ui(True)
        self._start_next_bus_item()


    def _on_session_processed(self, results):
        if self._ignore_bus_results:
            return
        for result in results or []:
            row = result.get('row',-1)
            if 0 <= row < self.table.rowCount():
                meta = self._row_metadata(row)
                meta['bus_done'] = bool(result.get('ok'))
                self._set_row_metadata(row,meta)
                self.table.item(row,self.COL_STATUS).setText('BUS 완료' if result.get('ok') else 'BUS 실패 (완료만 재시도)')
                self._log(f"{row+1}행 BUS 결과: {result.get('msg','')}")
        if self._bus_mode:
            self._bus_done += len(results or [])
            self._start_next_bus_item()
        elif self._waiting_for_bus:
            if any(r.get('row')==self._pending_after_add_row for r in results or []):
                self._waiting_for_bus = False
                self._pending_after_add_row = None
                self._start_next_job()


    def _on_session_ready(self, ok: bool, msg: str):
        if ok:
            self._set_plain_status("세션 준비됨")
        else:
            self._set_plain_status(f"세션 미준비: {msg}")

        self._set_running_ui(False)
        cbs = getattr(self, "_pending_bus_callbacks", [])
        self._pending_bus_callbacks = []

        for cb in cbs:
            try:
                cb(bool(ok))
            except Exception as e:
                self._log(f"[세션 준비 콜백 오류] {e}")

    def _on_session_busy(self, b: bool):
        self._set_running_ui(b or self.api_thread is not None or self._waiting_for_bus)

    def _on_session_downloaded(self, path: str, err: str):
        self._set_running_ui(False)
        self.btn_request.setEnabled(True)
        self.btn_settings.setEnabled(True)
        
        if err:
            if err.strip() == "사용자 취소":
                self._set_plain_status("요청 확인 취소됨")
                return
            self._set_plain_status("로드 실패")
            QMessageBox.critical(self, "오류", f"로드 실패: {err}")
            return
        
        self._set_plain_status("로드 완료")
        self.load_excel(path, append=False, silent=False, source="bus")
        
        try:
            end_combined = os.path.join(DL_DIR, "종료권한리스트_합본.xls")
            normal_combined = os.path.join(DL_DIR, "권한리스트_합본.xls")
            if os.path.basename(path) == "권한리스트_합본.xls" and os.path.exists(end_combined):
                self.load_excel(end_combined, append=True, silent=True, source="bus")
            elif os.path.basename(path) == "종료권한리스트_합본.xls" and os.path.exists(normal_combined):
                self.load_excel(normal_combined, append=True, silent=True, source="bus")
        except Exception:
            pass

        prog_grant = prog_rel = end_grant = end_rel = 0
        for r in range(self.table.rowCount()):
            kind = (self._get(r, self.COL_KIND) or "").strip()
            req  = (self._get(r, self.COL_REQTYPE) or REQ_GRANT).strip()
            if   kind == "진행" and req == REQ_GRANT:   prog_grant += 1
            elif kind == "진행" and req == REQ_RELEASE: prog_rel   += 1
            elif kind == "종료" and req == REQ_GRANT:   end_grant  += 1
            elif kind == "종료" and req == REQ_RELEASE: end_rel    += 1

        parts = []
        if prog_grant: parts.append(f"진행-부여 {prog_grant}건")
        if prog_rel:   parts.append(f"진행-해제 {prog_rel}건")
        if end_grant:  parts.append(f"종료-부여 {end_grant}건")
        if end_rel:    parts.append(f"종료-해제 {end_rel}건")
        self._log("로드 완료: " + (" , ".join(parts) if parts else "항목 없음"))

    def _load_creds(self):
        self.creds = {"id":"", "pw":""}
        try:
            if os.path.exists(CONF_FILE):
                with open(CONF_FILE, "r", encoding="utf-8") as f:
                    self.creds = json.load(f)
                    self.creds.setdefault("debug", False)
        except:
            pass
        self.remove_fail_tolerance = int(self.creds.get("remove_fail_tol", 5))

    def _save_creds(self, id_, pw_, debug_):
        with open(CONF_FILE, "w", encoding="utf-8") as f:
            json.dump({"id": id_, "pw": pw_, "debug": bool(debug_), "theme": getattr(self, "current_theme", "light")},
                      f, ensure_ascii=False, indent=2)

    def closeEvent(self,event):
        self._closing = True
        if self.api_thread is not None and self.api_thread.isRunning():
            self.api_worker.stop()
            event.ignore()
            return
        newdlg = getattr(self,'_newdlg',None)
        if newdlg and getattr(newdlg,'worker_thread',None) and newdlg.worker_thread.isRunning():
            newdlg.worker.stop()
            event.ignore()
            QTimer.singleShot(200,self.close)
            return
        self.notify_timer.stop()
        self._progress_timer.stop()
        self.session.cancel_current()
        if self.session.is_busy() or self.watch_session._mgr.is_busy():
            event.ignore()
            QTimer.singleShot(200,self.close)
            return
        if not self._shutdown_started:
            self._shutdown_started = True
            self.session.stopped.connect(self.session_thread.quit, type=Qt.DirectConnection)
            self.watch_session.stopped.connect(self.watch_thread.quit, type=Qt.DirectConnection)
            self.trigger_session_stop.emit()
            self.trigger_watcher_stop.emit()
        if self.session_thread.isRunning() or self.watch_thread.isRunning():
            event.ignore()
            QTimer.singleShot(100,self.close)
            return
        event.accept()


    def _open_settings(self):
        current_debug = bool(self.creds.get("debug", False))
        remembered = bool(self.creds.get("id"))
        current_tol = int(self.creds.get("remove_fail_tol", 5))

        dlg = SettingsDialog(self, self.creds, remembered=remembered, debug_on=current_debug, fail_tol_default=current_tol)

        if dlg.exec_() == QDialog.Accepted:
            uid, pw, do_save, debug_on, fail_tol, notify_mins = dlg.result()
            tol = dlg.get_fail_tol()

            self.creds = {
                "id": uid if do_save else "",
                "pw": pw if do_save else "",
                "debug": bool(debug_on),
                "remove_fail_tol": int(tol),
                "notify_refresh_min": int(notify_mins),
            }

            if uid and pw:
                self.watch_session.set_creds(uid, pw)
                self.trigger_watcher_start.emit()
            else:
                self.trigger_watcher_stop.emit()

            data = {
                "theme": getattr(self, "current_theme", "light"),
                "debug": bool(debug_on),
                "remove_fail_tol": int(tol),
                "notify_refresh_min": int(notify_mins),
            }
            if do_save:
                data.update({"id": uid, "pw": pw})

            try:
                with open(CONF_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

            self._notify_refresh_mins = int(notify_mins)
            self.notify_timer.setInterval(max(5, self._notify_refresh_mins) * 60 * 1000)
            self.debug_enabled = bool(debug_on)
            self.session.set_debug(self.debug_enabled, DEBUG_DIR)
            self.remove_fail_tolerance = int(tol)

    def _request_and_import(self):
        if not self.creds.get("id") or not self.creds.get("pw"):
            self._open_settings()
            if not self.creds.get("id") or not self.creds.get("pw"):
                QMessageBox.warning(self, "알림", "아이디/비밀번호가 필요합니다.")
                return

        self.session.set_creds(self.creds["id"], self.creds["pw"])

        if not self.session.is_ready():
            self._set_running_ui(True)
            self._set_progress_status("세션 준비 중", "", 0, 0, running=True)
            self.trigger_session_start.emit()

            end = time.time() + 60
            while time.time() < end and self.isVisible():
                QApplication.processEvents()
                if self.session.is_ready():
                    self._set_plain_status("세션 준비됨")
                    break
                time.sleep(0.05)

            if not self.session.is_ready():
                self._set_running_ui(False)
                QMessageBox.critical(self, "오류", "세션 초기화 실패")
                return

        self.btn_request.setEnabled(False)
        self.btn_settings.setEnabled(False)
        self._set_running_ui(True)
        self._set_progress_status("요청 확인 중", "", 0, 0, running=True)
        self.trigger_session_download.emit()

    def _clear_log(self):
        self.log.clear()

    def _stop_all(self):
        self.run_queue = []
        self.stop_requested = True
        self._ignore_bus_results = True
        self._waiting_for_bus = False
        self._bus_queue = []
        self._bus_mode = False
        self.session.cancel_current()
        if self.api_worker is not None:
            self.api_worker.stop()
            self._log('API 취소 요청 후 최종 상태를 기다립니다.')
        newdlg = getattr(self,'_newdlg',None)
        if newdlg and getattr(newdlg,'worker',None):
            newdlg.worker.stop()
        if self.api_thread is None and not self.session.is_busy():
            self._set_running_ui(False)
        self._finish_progress_status('중지 요청됨',stopped=True)


    def _set_running_ui(self, running: bool):
        running = running or getattr(self,'api_thread',None) is not None or getattr(self,'_waiting_for_bus',False)
        newdlg = getattr(self,'_newdlg',None)
        running = running or bool(newdlg and getattr(newdlg,'worker_thread',None))
        if running:
            self._set_progress_status(
                getattr(self, "_progress_detail", "") or "처리 중",
                getattr(self, "_progress_project", ""),
                getattr(self, "done_jobs", 0),
                getattr(self, "total_jobs", 0),
                running=True,
            )

        to_disable = [
            getattr(self, "btn_api_settings", None),
            getattr(self, "btn_manual", None),
            getattr(self, "btn_file", None),
            getattr(self, "btn_run_execute", None),
            getattr(self, "btn_run_complete", None),
            getattr(self, "btn_clear_log", None),
            getattr(self, "btn_request", None),
            getattr(self, "btn_settings", None),
            getattr(self, "chk_dry", None),
            getattr(self, "chk_auto_complete", None),
            getattr(self, "btn_newcheck", None),
        ]
        for w in to_disable:
            if w:
                w.setEnabled(not running)

        if hasattr(self, "btn_stop") and self.btn_stop:
            self.btn_stop.setEnabled(running)

        self.setAcceptDrops(not running)
        if hasattr(self,'table'):
            self.table.setEditTriggers(QAbstractItemView.NoEditTriggers if running else (QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked | QAbstractItemView.EditKeyPressed))
        if hasattr(self, "table") and self.table:
            self.table.setContextMenuPolicy(Qt.PreventContextMenu if running else Qt.CustomContextMenu)



    def _start_next_job(self):
        if self.api_thread is not None or self._waiting_for_bus:
            return
        if self._closing or self.stop_requested or not self.run_queue:
            self._set_running_ui(False)
            self._finish_progress_status('중지됨' if self.stop_requested else '완료', stopped=self.stop_requested)
            return
        row,payload,dry = self.run_queue.pop(0)
        self.current_row = row
        self.current_seq = self.done_jobs+1
        self.current_mode = 'add' if payload['operation']=='grant' else 'remove'
        self.table.item(row,self.COL_STATUS).setText('Preview 조회 중' if dry else 'API 실행중')
        self._log(f"API {'Preview' if dry else 'Job'} 요청: {payload['request_id']}")
        thread = QThread(self)
        worker = ApiJobWorker(self._api_client(),'access',payload,dry)
        self.api_thread,self.api_worker = thread,worker
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self._on_api_progress)
        worker.succeeded.connect(self._on_api_result)
        worker.failed.connect(self._on_api_error)
        worker.finished.connect(worker.deleteLater)
        worker.finished.connect(thread.quit, type=Qt.DirectConnection)
        thread.finished.connect(self._api_thread_finished)
        thread.finished.connect(thread.deleteLater)
        thread.start()


    def open_manual_dialog(self):
        try:
            dlg = ManualEntryDialog(self)
            if dlg.exec_() == QDialog.Accepted and dlg.result_row:
                reqtype, user, proj, lv2, lv3, dept, role, kind = dlg.result_row
                self.add_table_row(reqtype, user, proj, lv2, lv3, dept, role, kind)   
        except Exception as e:
            self._log(f"[수동입력 오픈 예외] {e}")
            QMessageBox.critical(self, "오류", f"수동 입력 다이얼로그 실행 오류:\n{e}")

    def add_table_row(self, reqtype: str, user: str, proj: str, lv2: str, lv3: str, dept: str, role: str, kind: str = "진행"):
        self.table.blockSignals(True)
        try:
            r = self.table.rowCount()
            self.table.insertRow(r)

            chk = QCheckBox()
            chk.setChecked(True)
            
            wrapper = QWidget()
            layout = QHBoxLayout(wrapper)
            layout.addWidget(chk)
            layout.setAlignment(Qt.AlignCenter)
            layout.setContentsMargins(0,0,0,0)
            self.table.setCellWidget(r, self.COL_SELECT, wrapper)

            vals = [kind, reqtype or REQ_GRANT, user, "", proj, lv2, lv3, dept or "", role or "", "대기"]
            
            for j, val in enumerate(vals, start=1):
                col = j
                it = QTableWidgetItem(str(val))
                it.setTextAlignment(Qt.AlignCenter)

                if col in self.EDITABLE_COLS:
                    it.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                else:
                    it.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                self.table.setItem(r, col, it)
        finally:
            self.table.blockSignals(False)

        self._set_row_metadata(r,dict(source='test-client',request_id=new_client_request_id(),bus_done=False))
        if (kind or "").strip() == "종료":
            self._log(f"수동 입력 추가: 구분={kind}, 요청={reqtype}, 사번={user}, 프로젝트={proj}")
        else:
            self._log(f"수동 입력 추가: 구분={kind}, 요청={reqtype}, 사번={user}, 프로젝트={proj}, L2={lv2}, L3={lv3}, ROLE={'없음' if not role else role}")

    def _log(self, message: str, seq: int = None, dry: bool = False):
        api_key = getattr(self,'api_config',{}).get('api_key','')
        if api_key:
            message = message.replace(api_key,'[API Key]')
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        dry_str = " Dry" if dry else ""
        tag = f"{dry_str} #{seq}" if seq is not None else ""

        norm = (message or "").replace("\r", "\n").replace("\t", "    ").rstrip()
        block = f"[{ts}{tag}]\n{norm}\n"

        cursor = self.log.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log.setTextCursor(cursor)

        if self.log.toPlainText():
            self.log.insertPlainText("\n")

        self.log.insertPlainText(block)
        self.log.ensureCursorVisible()

        fn = os.path.join(LOG_DIR, f"access_{datetime.date.today().strftime('%Y%m%d')}.log")
        with open(fn, "a", encoding="utf-8") as f:
            f.write(block)

    def choose_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "승인 엑셀 선택", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        self.load_excel(path)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                lf = url.toLocalFile().lower()
                if lf.endswith(".xlsx") or lf.endswith(".xls"):
                    event.acceptProposedAction(); return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls: return
        path = urls[0].toLocalFile()
        if path.lower().endswith((".xlsx",".xls")):
            self.load_excel(path)

    def load_excel(self, file_path: str, append: bool = False, silent: bool = False, source: str = 'test-client'):
        try:
            self.table.blockSignals(True)
            #if not (append and silent):
                #self.file_label.setText(file_path)
            ext = os.path.splitext(file_path)[1].lower()

            start_offset = self.table.rowCount() if append else 0
            
            rows = []
            identities = []

            if ext == ".xlsx":
                wb = load_workbook(file_path, data_only=True)
                ws = wb.active
                header_raw = [str(c.value) if c.value is not None else "" for c in ws[1]]
                colmap = auto_map_columns(header_raw)
                
                is_end = any(("열람" in (str(h) or "")) for h in header_raw)
                kind = "종료" if is_end else "진행"
                
                required = {"user", "proj"} if is_end else {"user", "proj", "level2", "level3"}
                missing = [k for k in required if k not in colmap]
                if missing:
                    self._log(f"엑셀 로드 실패: 필요한 컬럼 없음 -> {missing} / 헤더: {header_raw}")
                    return

                def gv(row, key, default=""):
                    i = colmap.get(key)
                    if i is None: 
                        return default
                    v = row[i] if i < len(row) else None
                    return (str(v).strip() if v is not None else default)
                
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if not any(r):
                        continue
                    user_id = gv(r, "user")
                    name    = gv(r, "name")
                    proj    = gv(r, "proj")
                    lv2     = "" if is_end else gv(r, "level2")
                    lv3     = "" if is_end else gv(r, "level3")
                    role    = gv(r, "role") if not is_end else ""
                    dept    = gv(r, "dept")
                    
                    if not user_id or not proj:
                        continue
                    if not is_end and (not lv2 or not lv3):
                        continue
                    
                    reqtype = REQ_RELEASE if _is_release_row_by_values(list(r), header_raw) else REQ_GRANT
                    path = build_path_l3(proj, lv2, lv3)
                    rows.append((kind, reqtype, user_id, name, proj, lv2, lv3, dept, role, "대기"))
                    identities.append(bus_row_identity(header_raw,r))

                wb.close()

            elif ext == ".xls":
                header_raw, data_rows = _parse_html_best_table(file_path)
                if not header_raw:
                    self._log("엑셀 로드 실패: .xls(HTML) 테이블을 찾지 못했습니다.")
                    return
                
                colmap = auto_map_columns(header_raw)

                is_end = any(("열람" in (str(h) or "")) for h in header_raw)
                kind = "종료" if is_end else "진행"
                
                required = {"user", "proj"} if is_end else {"user", "proj", "level2", "level3"}
                missing = [k for k in required if k not in colmap]
                if missing:
                    self._log(f"엑셀 로드 실패: 필요한 컬럼 없음 -> {missing} / 헤더: {header_raw}")
                    return

                def gv(row, key, default=""):
                    i = colmap.get(key)
                    if i is None: 
                        return default
                    v = row[i] if i < len(row) else None
                    return (str(v).strip() if v is not None else default)

                for r in data_rows:
                    user_id = gv(r, "user")
                    name    = gv(r, "name")
                    proj    = gv(r, "proj")
                    lv2     = "" if is_end else gv(r, "level2")
                    lv3     = "" if is_end else gv(r, "level3")
                    role    = gv(r, "role") if not is_end else ""
                    dept    = gv(r, "dept")
                    
                    if not user_id or not proj:
                        continue
                    if not is_end and (not lv2 or not lv3):
                        continue
                    
                    reqtype = REQ_RELEASE if _is_release_row_by_values(r, header_raw) else REQ_GRANT
                    path = build_path_l3(proj, lv2, lv3)
                    rows.append((kind, reqtype, user_id, name, proj, lv2, lv3, dept, role, "대기"))
                    identities.append(bus_row_identity(header_raw,r))

            else:
                self._log(f"엑셀 로드 실패: 지원하지 않는 확장자 ({ext})")
                return

            if append:
                self.table.setRowCount(start_offset + len(rows))
                base = start_offset
            else:
                self.table.setRowCount(len(rows))
                base = 0

            for i, row in enumerate(rows):
                r = base + i
                
                chk = QCheckBox()
                chk.setChecked(True)
                
                wrapper = QWidget()
                layout = QHBoxLayout(wrapper)
                layout.addWidget(chk)
                layout.setAlignment(Qt.AlignCenter)
                layout.setContentsMargins(0, 0, 0, 0)
                self.table.setCellWidget(r, self.COL_SELECT, wrapper)

                for j, val in enumerate(rows[i]):
                    col = j + 1
                    it = QTableWidgetItem(str(val))
                    it.setTextAlignment(Qt.AlignCenter)
                    if col in self.EDITABLE_COLS:
                        it.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable)
                    else:
                        it.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    self.table.setItem(r, col, it)

            for row in range(base,base+len(rows)):
                values = self._row_values(row)
                try:
                    request_id = make_bus_request_id(values,identities[row-base]) if source=='bus' else new_client_request_id()
                except ValueError:
                    request_id = new_client_request_id()
                self._set_row_metadata(row,dict(source=source,request_id=request_id,bus_done=False))
            hv = self.table.horizontalHeader()
            hv.setStretchLastSection(True)

        except Exception as e:
            self._log(f"엑셀 로드 실패: {e}")

        finally:
            self.table.blockSignals(False)

    def _on_cell_changed(self,row,col):
        if col == self.COL_STATUS:
            return
        meta = self._row_metadata(row)
        if meta:
            self._set_row_metadata(row,dict(source='test-client',request_id=new_client_request_id()))
            self.table.item(row,self.COL_STATUS).setText('수정됨 (수동 요청)')

        
    def _open_table_menu(self, pos):
        row = self.table.indexAt(pos).row()
        m = QMenu(self)
        act_del_row = m.addAction("행 삭제")
        act_del_sel = m.addAction("선택 행 삭제")
        act_del_all = m.addAction("전체 삭제")
        
        gpos = self.table.viewport().mapToGlobal(pos)
        act = m.exec_(gpos)
        
        if act == act_del_row:
            self._delete_row(row)
        elif act == act_del_sel:
            self._delete_checked_rows()
        elif act == act_del_all:
            self._delete_all_rows()     

    def _delete_all_rows(self):
        self.table.blockSignals(True)
        try:
            self.table.setRowCount(0)
        finally:
            self.table.blockSignals(False)

    def _delete_row(self, row: int):
        if row < 0 or row >= self.table.rowCount():
            return
        self.table.removeRow(row)

    def _get_checkbox(self, row: int) -> QCheckBox:
        w = self.table.cellWidget(row, self.COL_SELECT)
        if not w:
            return None
        return w.findChild(QCheckBox)

    def _is_row_checked(self, row: int) -> bool:
        cb = self._get_checkbox(row)
        return bool(cb and cb.isChecked())

    def _delete_checked_rows(self):
        for r in range(self.table.rowCount()-1, -1, -1):
            if self._is_row_checked(r):
                self.table.removeRow(r)

    def validate_row(self, row, mode='add'):
        try:
            meta = self._row_metadata(row)
            make_access_payload(self._row_values(row), meta.get('request_id') or new_client_request_id(),
                                meta.get('source','test-client'), self._requester())
            return True, ''
        except ValueError as exc:
            return False, f'{row+1}행: {exc}'


    def toggle_all_rows(self, checked: bool):
        for r in range(self.table.rowCount()):
            cb = self._get_checkbox(r)
            if cb:
                cb.setChecked(checked)

    def run_execute(self):
        if self.api_thread is not None or self.session.is_busy() or self._waiting_for_bus:
            return
        if not self.api_config.get('api_key'):
            self._open_api_settings()
            if not self.api_config.get('api_key'):
                return
        self.stop_requested = False
        self._ignore_bus_results = False
        self._waiting_for_bus = False
        self.auto_complete_after_add = self.chk_auto_complete.isChecked()
        queue = []
        for row in range(self.table.rowCount()):
            if not self._is_row_checked(row):
                continue
            ok, reason = self.validate_row(row)
            if not ok:
                self.table.item(row,self.COL_STATUS).setText('검증실패')
                self._log(reason)
                continue
            meta = self._row_metadata(row)
            if not meta:
                meta = dict(source='test-client',request_id=new_client_request_id())
            values = self._row_values(row)
            # Keep the original operator when retrying the same idempotent payload.
            payload = meta.get('payload') or make_access_payload(values,meta['request_id'],meta['source'],self._requester())
            meta['payload'] = payload
            if not self.chk_dry.isChecked():
                meta['result'] = {}
            self._set_row_metadata(row,meta)
            queue.append((row,payload,self.chk_dry.isChecked()))
        if not queue:
            return
        self.run_queue = queue
        self.total_jobs = len(queue)
        self.done_jobs = 0
        self._progress_started_at = time.time()
        self._set_running_ui(True)
        self._start_next_job()


    def _get(self, row: int, col: int) -> str:
        it = self.table.item(row, col)
        return it.text().strip() if it else ""

    def _show_initial_api_settings(self):
        if not self.api_config.get('api_key'):
            self._open_api_settings()


    def _open_api_settings(self):
        if self.api_thread is not None or self._waiting_for_bus:
            return
        dialog = ApiSettingsDialog(self.api_config, self)
        if dialog.exec_() == QDialog.Accepted:
            self.api_config = dialog.values()
            self._log('API 연결 설정을 저장했습니다. API Key는 로그에 기록하지 않습니다.')


    def _api_client(self):
        return FolderGrantApiClient(self.api_config.get('base_url', DEFAULT_API_URL),
                                   self.api_config.get('api_key', ''))


    def _requester(self):
        import getpass
        return self.creds.get('id') or getpass.getuser() or 'client-operator'


    def _row_values(self, row):
        return dict(kind=self._get(row, self.COL_KIND), req=self._get(row, self.COL_REQTYPE),
                    user=self._get(row, self.COL_USER), proj=self._get(row, self.COL_PROJ),
                    lv2=self._get(row, self.COL_LV2), lv3=self._get(row, self.COL_LV3),
                    role=self._get(row, self.COL_ROLE))


    def _row_metadata(self, row):
        item = self.table.item(row, self.COL_STATUS)
        return dict(item.data(Qt.UserRole) or {}) if item else {}


    def _set_row_metadata(self, row, meta):
        item = self.table.item(row, self.COL_STATUS)
        if item:
            old = self.table.blockSignals(True)
            item.setData(Qt.UserRole, dict(meta))
            self.table.blockSignals(old)


    def _bus_target(self, row):
        values = self._row_values(row)
        result = self._row_metadata(row).get('result', {})
        # Prefer the server's executed path; never check UNC paths on the client.
        paths = [step.get('target', '') for step in result.get('steps', [])
                 if str(step.get('target', '')).startswith('\\\\')]
        path = paths[0] if paths else ('' if values['kind']=='종료' else build_path_l3(values['proj'],values['lv2'],values['lv3']))
        return {**values, 'row': row, 'path': path}


    def _on_api_progress(self, job):
        self._set_progress_status(f"API {job.get('status','')} / {job.get('job_id','')}",
                                 self._get(self.current_row,self.COL_PROJ),self.done_jobs,self.total_jobs,running=True)


    def _on_api_result(self, result):
        row = self.current_row
        meta = self._row_metadata(row)
        result = dict(result)
        meta['result'] = result
        self._set_row_metadata(row,meta)
        status = result.get('status')
        label = {'preview':'DryRun', 'simulated':'모의완료 (BUS 미처리)',
                 'succeeded':'API 성공', 'partially_succeeded':'부분 성공 (BUS 미처리)',
                 'failed':'API 실패', 'cancelled':'취소됨'}.get(status,'알 수 없는 결과')
        self.table.item(row,self.COL_STATUS).setText(label)
        self._log(json.dumps(result,ensure_ascii=False,indent=2))
        self.done_jobs += 1
        stopped = self.stop_requested or result.get('client_stopped',False)
        if (self.auto_complete_after_add and not meta.get('bus_done')
                and can_complete_bus(result,meta.get('source'),stopped)):
            # Session preparation pumps Qt events. Hold the queue before entering it.
            self._waiting_for_bus = True
            self._pending_after_add_row = row
            if self._ensure_bus_session('API 성공 후 완료 처리') and not self.stop_requested and not self._closing:
                self.table.item(row,self.COL_STATUS).setText('완료 처리중')
                self.trigger_session_process.emit([self._bus_target(row)])
            else:
                self._waiting_for_bus = False
                self._pending_after_add_row = None
                self._log('API는 성공했습니다. BUS 세션 준비 후 완료 처리만 다시 실행하세요.')
                QTimer.singleShot(0,self._start_next_job)


    def _on_api_error(self, message):
        self.table.item(self.current_row,self.COL_STATUS).setText('API 오류 (등록 결과 확인 필요)')
        self._log(message)
        self.done_jobs += 1


    def _api_thread_finished(self):
        self.api_worker = None
        self.api_thread = None
        if self._closing:
            QTimer.singleShot(0,self.close)
        elif not self._waiting_for_bus:
            self._start_next_job()


def _excepthook(etype, value, tb):
    try:
        import traceback
        msg = "".join(traceback.format_exception(etype, value, tb))
        print(msg)
    except:
        pass

    QMessageBox.critical(None, "치명적 오류", str(value))

import sys as _sys
if __name__ == "__main__":
    _sys.excepthook = _excepthook

if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 9))
    window = AccessManager()
    def _on_quit():
        try:
            window.close()
        except Exception:
            pass
    window.show()
    sys.exit(app.exec_())
