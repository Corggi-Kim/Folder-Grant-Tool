"""Run real PowerShell through a Windows-size-limited launch boundary.

These tests perform only stdout, local temporary-file and sleep operations.
No AD, SMB or permission commands are executed.
"""
import base64
import os
import shutil
import subprocess
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from folder_grant_api.infrastructure.powershell import PowerShellExecutionError, PowerShellRunner
from folder_grant_api.infrastructure.powershell_commands import quote


REAL_POWERSHELL = os.environ.get('FGT_TEST_POWERSHELL') or shutil.which('pwsh') or shutil.which('powershell.exe')
needs_powershell = pytest.mark.skipif(not REAL_POWERSHELL, reason='Set FGT_TEST_POWERSHELL for actual interpreter tests')


@pytest.fixture(autouse=True)
def writable_powershell_directories(tmp_path, monkeypatch):
    for key in ('XDG_CACHE_HOME','XDG_CONFIG_HOME','XDG_DATA_HOME'):
        path=tmp_path/key.lower();path.mkdir()
        monkeypatch.setenv(key,str(path))


def length_error():
    error = FileNotFoundError('The filename or extension is too long')
    error.winerror = 206
    return error


@pytest.fixture
def windows_size_limit(monkeypatch):
    native_run, native_popen = subprocess.run, subprocess.Popen
    def guard(args):
        # CreateProcessW counts UTF-16 code units, including its terminating NUL.
        if len(subprocess.list2cmdline(args).encode('utf-16-le')) // 2 + 1 > 32767:
            raise length_error()
    def run(args, *a, **kw):
        guard(args)
        return native_run(args, *a, **kw)
    def popen(args, *a, **kw):
        guard(args)
        return native_popen(args, *a, **kw)
    monkeypatch.setattr(subprocess, 'run', run)
    monkeypatch.setattr(subprocess, 'Popen', popen)


def runner(cancellable, seconds=10):
    return PowerShellRunner(REAL_POWERSHELL, seconds, cancel_check=(lambda: False) if cancellable else None)


@needs_powershell
@pytest.mark.parametrize('cancellable', [False, True])
def test_large_unicode_source_survives_windows_launch_limit(cancellable, windows_size_limit):
    value = "한글😀 & $variable `literal '인용' \\자료/" * 1300
    result = runner(cancellable).run('$text=' + quote(value) + '; Write-Output $text')
    assert result.return_code == 0
    assert result.stdout == value


@needs_powershell
@pytest.mark.parametrize('cancellable', [False, True])
def test_script_scope_functions_and_return_keep_their_results(cancellable):
    result = runner(cancellable).run("$script:checks=New-Object 'System.Collections.Generic.List[object]';"
        "function Record { [void]$script:checks.Add('검증😀') };Record;"
        "return ($script:checks.ToArray() | ConvertTo-Json -Compress)")
    assert result.stdout == '"검증😀"'


@needs_powershell
@pytest.mark.parametrize('cancellable', [False, True])
@pytest.mark.parametrize('command,message', [("throw '검증 실패'", '검증 실패'), ('exit 17', '17')])
def test_source_failure_never_becomes_success(cancellable, command, message):
    with pytest.raises(PowerShellExecutionError) as caught:
        runner(cancellable).run(command)
    assert caught.value.code == 'POWERSHELL_FAILED'
    assert message in str(caught.value)


@needs_powershell
@pytest.mark.parametrize('cancellable', [False, True])
@pytest.mark.parametrize('wire_kind', ['no_gate', 'gate_only', 'wrong_length', 'extra_source', 'bad_base64'])
def test_missing_or_incomplete_stdin_never_executes_source(tmp_path, monkeypatch, cancellable, wire_kind):
    marker = tmp_path / 'must-not-be-created'
    source = 'Set-Content -LiteralPath ' + quote(str(marker)) + " -Value 'changed'"
    encoded = base64.b64encode(source.encode('utf-8'))
    header = b'FGT_START\n' + str(len(encoded)).encode('ascii') + b'\n'
    wires = dict(no_gate=b'', gate_only=b'FGT_START\n',
                 wrong_length=b'FGT_START\n'+str(len(encoded)+4).encode('ascii')+b'\n'+encoded,
                 extra_source=header+encoded+b'AAAA', bad_base64=b'FGT_START\n4\n!!!!')
    native_popen = subprocess.Popen
    def launch(*args, **kwargs):
        process = native_popen(*args, **kwargs)
        communicate = process.communicate
        def send_incomplete(input=None, timeout=None):
            return communicate(input=wires[wire_kind], timeout=timeout)
        process.communicate = send_incomplete
        return process
    monkeypatch.setattr(subprocess, 'Popen', launch)
    with pytest.raises(PowerShellExecutionError) as caught:
        runner(cancellable).run(source)
    assert caught.value.code == 'POWERSHELL_FAILED'
    assert not marker.exists()


@needs_powershell
@pytest.mark.parametrize('cancellable', [False, True])
def test_long_source_still_obeys_timeout(cancellable, windows_size_limit):
    source = '#'+('padding' * 5500)+'\nStart-Sleep -Seconds 30'
    start = time.monotonic()
    with pytest.raises(PowerShellExecutionError) as caught:
        runner(cancellable, seconds=1).run(source)
    assert caught.value.code == 'POWERSHELL_TIMEOUT'
    assert time.monotonic()-start < 8


@needs_powershell
def test_long_source_can_be_cancelled_after_it_starts(tmp_path, windows_size_limit):
    marker=tmp_path/'started'
    stop=threading.Event()
    source='#'+('padding' * 5500)+'\nSet-Content -LiteralPath '+quote(str(marker))+" -Value 'started';Start-Sleep -Seconds 30"
    job=PowerShellRunner(REAL_POWERSHELL,15,cancel_check=stop.is_set)
    with ThreadPoolExecutor() as pool:
        future=pool.submit(job.run,source)
        try:
            deadline=time.monotonic()+5
            while not marker.exists() and not future.done() and time.monotonic()<deadline:
                threading.Event().wait(.02)
            assert marker.exists(), 'long script failed to start'
            stop.set()
            with pytest.raises(PowerShellExecutionError) as caught:future.result(timeout=5)
            assert caught.value.code=='POWERSHELL_CANCELLED'
        finally:
            stop.set()


@pytest.mark.parametrize('cancellable', [False, True])
@pytest.mark.parametrize('winerror,expected', [(206,'POWERSHELL_COMMAND_TOO_LONG'),(2,'POWERSHELL_NOT_FOUND'),(3,'POWERSHELL_NOT_FOUND'),(123,'POWERSHELL_START_FAILED')])
def test_windows_launch_error_keeps_its_actual_cause(monkeypatch, cancellable, winerror, expected):
    error=FileNotFoundError('Windows launch error');error.winerror=winerror
    def fail(*a,**kw):raise error
    monkeypatch.setattr(subprocess,'run',fail);monkeypatch.setattr(subprocess,'Popen',fail)
    with pytest.raises(PowerShellExecutionError) as caught:
        PowerShellRunner('configured-powershell.exe',10,cancel_check=(lambda:False) if cancellable else None).run('Write-Output 1')
    assert caught.value.code==expected
    assert str(winerror) in str(caught.value)


@pytest.mark.skipif(os.name=='nt',reason='POSIX executable fixture')
def test_large_stdin_cancellation_does_not_wait_for_an_unread_pipe(fake_powershell):
    executable,marker=fake_powershell
    stop=threading.Event()
    job=PowerShellRunner(str(executable),10,cancel_check=stop.is_set)
    with ThreadPoolExecutor() as pool:
        future=pool.submit(job.run,'#'+('x'*600000))
        try:
            deadline=time.monotonic()+5
            while not marker.exists() and not future.done() and time.monotonic()<deadline:
                threading.Event().wait(.02)
            assert marker.exists(), 'process failed to start before pipe cancellation'
            stop.set()
            with pytest.raises(PowerShellExecutionError) as caught:future.result(timeout=5)
            assert caught.value.code=='POWERSHELL_CANCELLED'
            previous=marker.read_text();threading.Event().wait(.15)
            assert marker.read_text()==previous
        finally:
            stop.set()
