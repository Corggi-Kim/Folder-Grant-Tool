import os
import subprocess
import threading
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

import pytest
from folder_grant_api.infrastructure.powershell import PowerShellRunner, PowerShellExecutionError



def wait_for_file(path):
    deadline=time.monotonic()+5
    while not path.exists() and time.monotonic()<deadline:
        threading.Event().wait(.02)
    assert path.exists(), 'test process did not start'


@pytest.mark.skipif(os.name=='nt',reason='POSIX executable fixture; Windows containment tested separately')
@pytest.mark.parametrize('parent_exits',[False,True])
def test_active_cancel_terminates_parent_and_child(fake_powershell,parent_exits):
    executable,marker=fake_powershell
    if parent_exits:
        executable.write_text(executable.read_text().replace('time.sleep(60)','sys.exit(0)'))
    stop=threading.Event()
    runner=PowerShellRunner(str(executable),10,cancel_check=stop.is_set)
    with ThreadPoolExecutor() as pool:
        future=pool.submit(runner.run,'ignored')
        try:
            wait_for_file(marker)
            stop.set()
            with pytest.raises(PowerShellExecutionError) as error: future.result(timeout=5)
            assert error.value.code=='POWERSHELL_CANCELLED'
            last=marker.read_text();threading.Event().wait(.15)
            assert marker.read_text()==last
        finally:
            stop.set()


@pytest.mark.skipif(os.name=='nt',reason='POSIX executable fixture')
def test_timeout_terminates_descendants(fake_powershell):
    executable,marker=fake_powershell
    with pytest.raises(PowerShellExecutionError) as error:
        PowerShellRunner(str(executable),1,cancel_check=lambda:False).run('ignored')
    assert error.value.code=='POWERSHELL_TIMEOUT'
    last=marker.read_text();threading.Event().wait(.15)
    assert marker.read_text()==last


def test_prestart_cancel_does_not_launch_process():
    from unittest.mock import patch
    with patch('subprocess.Popen') as popen:
        with pytest.raises(PowerShellExecutionError) as error:
            PowerShellRunner('unused',30,cancel_check=lambda:True).run('must not run')
        assert error.value.code=='POWERSHELL_CANCELLED'
        popen.assert_not_called()


@pytest.mark.skipif(os.name=='nt',reason='POSIX executable fixture')
def test_cancel_does_not_interrupt_another_job(fake_powershell):
    executable,marker=fake_powershell
    second=executable.with_name('second-powershell');second.write_text(executable.read_text());second.chmod(0o755)
    second_marker=second.with_suffix('.running')
    stops=[threading.Event(),threading.Event()]
    with ThreadPoolExecutor() as pool:
        futures=[pool.submit(PowerShellRunner(str(path),10,cancel_check=stop.is_set).run,'ignored')
                 for path,stop in zip((executable,second),stops)]
        try:
            wait_for_file(marker);wait_for_file(second_marker)
            stops[0].set()
            with pytest.raises(PowerShellExecutionError): futures[0].result(timeout=5)
            assert not futures[1].done()
            before=second_marker.read_text();threading.Event().wait(.15)
            assert second_marker.read_text()!=before
        finally:
            for stop in stops:stop.set()
            for future in futures:
                try:future.result(timeout=5)
                except PowerShellExecutionError:pass


@pytest.mark.skipif(os.name=='nt',reason='POSIX executable fixture')
def test_windows_assignment_failure_never_opens_startup_gate(fake_powershell):
    from unittest.mock import patch, Mock
    executable,marker=fake_powershell
    job=Mock();job.attach.side_effect=OSError('assignment denied')
    with patch('folder_grant_api.infrastructure.powershell.WINDOWS',True),patch('folder_grant_api.infrastructure.powershell.WindowsJob',return_value=job):
        with pytest.raises(PowerShellExecutionError) as error:
            PowerShellRunner(str(executable),10,cancel_check=lambda:False).run('ignored')
    assert error.value.code=='POWERSHELL_PROCESS_CONTROL_FAILED'
    assert not marker.exists()
    job.close.assert_called_once()


def test_windows_job_configures_only_its_own_handle_for_tree_termination():
    import ctypes
    from unittest.mock import Mock,patch
    from folder_grant_api.infrastructure.process_tree import WindowsJob
    kernel=Mock();kernel.CreateJobObjectW.return_value=1234
    kernel.SetInformationJobObject.return_value=True
    kernel.AssignProcessToJobObject.return_value=True
    kernel.TerminateJobObject.return_value=True
    observations=[]
    def query(handle,info_class,buffer,size,length):
        buffer._obj.ActiveProcesses=1 if not observations else 0
        observations.append(buffer._obj.ActiveProcesses)
        return True
    kernel.QueryInformationJobObject.side_effect=query
    process=Mock();process._handle=5678
    with patch.object(ctypes,'WinDLL',return_value=kernel,create=True):
        job=WindowsJob();job.attach(process);job.terminate();job.close()
    limits=kernel.SetInformationJobObject.call_args.args[2]._obj
    assert limits.BasicLimitInformation.LimitFlags==0x2000
    kernel.AssignProcessToJobObject.assert_called_once_with(1234,5678)
    kernel.TerminateJobObject.assert_called_once_with(1234,1)
    kernel.CloseHandle.assert_called_once_with(1234)
    assert observations==[1,0]


def real_powershell():
    import shutil
    return os.environ.get('FGT_TEST_POWERSHELL') or (shutil.which('powershell.exe') if os.name=='nt' else None)


@pytest.mark.skipif(not real_powershell(),reason='Set FGT_TEST_POWERSHELL for actual interpreter tests')
def test_cancellable_runner_executes_real_powershell_and_decodes_utf8(tmp_path,monkeypatch):
    for key in ('XDG_CACHE_HOME','XDG_CONFIG_HOME','XDG_DATA_HOME'): monkeypatch.setenv(key,str(tmp_path))
    result=PowerShellRunner(real_powershell(),10,cancel_check=lambda:False).run("Write-Output '정상'")
    assert result.stdout=='정상'


@pytest.mark.skipif(not real_powershell(),reason='Set FGT_TEST_POWERSHELL for actual interpreter tests')
def test_startup_gate_stays_closed_on_stdin_eof(tmp_path,monkeypatch):
    from unittest.mock import Mock,patch
    from folder_grant_api.infrastructure.powershell_commands import quote
    for key in ('XDG_CACHE_HOME','XDG_CONFIG_HOME','XDG_DATA_HOME'): monkeypatch.setenv(key,str(tmp_path))
    marker=tmp_path/'must-not-run'
    def attach(process):
        process.stdin.close();process.stdin=None
        process.wait(timeout=5)
        raise OSError('worker lost before assignment')
    job=Mock();job.attach.side_effect=attach
    with patch('folder_grant_api.infrastructure.powershell.WINDOWS',True),patch('folder_grant_api.infrastructure.powershell.WindowsJob',return_value=job):
        with pytest.raises(PowerShellExecutionError) as error:
            PowerShellRunner(real_powershell(),10,cancel_check=lambda:False).run('Set-Content -LiteralPath '+quote(str(marker))+" -Value 'changed'")
    assert error.value.code=='POWERSHELL_PROCESS_CONTROL_FAILED'
    assert not marker.exists()


@pytest.mark.skipif(os.name!='nt' or not real_powershell(),reason='Requires native Windows PowerShell and Job Objects')
def test_native_windows_cancel_terminates_powershell_and_child(tmp_path):
    import sys
    from folder_grant_api.infrastructure.powershell_commands import quote
    marker=tmp_path/'heartbeat'
    script=tmp_path/'child.py'
    script.write_text('import time\nfrom pathlib import Path\np=Path('+repr(str(marker))+')\ni=0\nwhile True:\n p.write_text(str(i));i+=1;time.sleep(.03)\n')
    stop=threading.Event()
    runner=PowerShellRunner(real_powershell(),10,cancel_check=stop.is_set)
    with ThreadPoolExecutor() as pool:
        future=pool.submit(runner.run,'& '+quote(sys.executable)+' '+quote(str(script)))
        try:
            wait_for_file(marker);stop.set()
            with pytest.raises(PowerShellExecutionError) as error: future.result(timeout=5)
            assert error.value.code=='POWERSHELL_CANCELLED'
            last=marker.read_text();threading.Event().wait(.15)
            assert marker.read_text()==last
        finally:stop.set()


def test_termination_failure_is_not_reported_as_confirmed_cancellation():
    from unittest.mock import Mock,patch
    from folder_grant_api.infrastructure.powershell import PowerShellCancelled
    job=Mock();job.terminate.side_effect=OSError('termination denied')
    process=Mock();process.poll.return_value=None
    checks=iter([False,True])
    with patch('folder_grant_api.infrastructure.powershell.WINDOWS',True),patch('folder_grant_api.infrastructure.powershell.WindowsJob',return_value=job),patch('subprocess.Popen',return_value=process):
        with pytest.raises(PowerShellExecutionError) as error:
            PowerShellRunner('unused',10,cancel_check=lambda:next(checks)).run('ignored')
    assert error.value.code=='POWERSHELL_STOP_FAILED'
    assert not isinstance(error.value,PowerShellCancelled)
    job.close.assert_called_once()


def test_windows_job_requires_all_descendants_to_exit_before_confirmation():
    import ctypes
    from unittest.mock import Mock,patch
    from folder_grant_api.infrastructure.process_tree import WindowsJob
    kernel=Mock();kernel.CreateJobObjectW.return_value=1234
    def query(handle,info_class,buffer,size,length):
        buffer._obj.ActiveProcesses=1
        return True
    kernel.QueryInformationJobObject.side_effect=query
    with patch.object(ctypes,'WinDLL',return_value=kernel,create=True):
        job=WindowsJob()
        try:
            with patch('folder_grant_api.infrastructure.process_tree.time.monotonic',side_effect=[0,6]):
                with pytest.raises(TimeoutError):job.terminate()
        finally:job.close()
