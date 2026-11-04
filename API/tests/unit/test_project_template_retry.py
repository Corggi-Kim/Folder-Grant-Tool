import os
import shutil

import pytest

from folder_grant_api.infrastructure.models import JobStepRecord
from folder_grant_api.infrastructure.powershell import PowerShellExecutionError, PowerShellRunner
from folder_grant_api.infrastructure.powershell_commands import build_step_command, quote


@pytest.fixture
def powershell(tmp_path, monkeypatch):
    executable = os.environ.get('FGT_TEST_POWERSHELL') or shutil.which('pwsh')
    if not executable:
        pytest.skip('Set FGT_TEST_POWERSHELL for actual interpreter tests')
    for key in ('XDG_CACHE_HOME', 'XDG_CONFIG_HOME', 'XDG_DATA_HOME'):
        monkeypatch.setenv(key, str(tmp_path))
    return PowerShellRunner(executable, 10)


def folder_step(root, template, *, incomplete):
    return JobStepRecord(
        step_order=2, step_type='ensure_project_folder', status='queued', target=str(root),
        description='template', details={'template_path':str(template), '_template_copy_incomplete':incomplete},
    )


def copy_boundary(root, template, invoked):
    return (
        'function robocopy {'
        f'Set-Content -LiteralPath {quote(str(invoked))} -Value invoked;'
        f'New-Item -ItemType Directory -Force -Path {quote(str(root))} | Out-Null;'
        f'Copy-Item -LiteralPath {quote(str(template / "template.txt"))} -Destination {quote(str(root / "template.txt"))} -Force;'
        '$global:LASTEXITCODE=1; Write-Output copied; };'
    )


def test_incomplete_template_copy_stops_before_existing_content_can_be_overwritten(tmp_path,powershell):
    root = tmp_path/'project'
    template = tmp_path/'template'
    root.mkdir(); template.mkdir()
    (root/'template.txt').write_text('operator content')
    (template/'template.txt').write_text('template content')
    invoked = tmp_path/'copy-invoked'
    command = copy_boundary(root,template,invoked) + build_step_command(folder_step(root,template,incomplete=True))
    with pytest.raises(PowerShellExecutionError) as error:
        powershell.run(command)
    assert 'PROJECT_TEMPLATE_REPAIR_REQUIRED' in str(error.value)
    assert '이전 템플릿 복사가 완료되지 않았습니다.' in str(error.value)
    assert (root/'template.txt').read_text() == 'operator content'
    assert not invoked.exists()


def test_incomplete_template_copy_can_create_missing_root(tmp_path,powershell):
    root = tmp_path/'project'
    template = tmp_path/'template'
    template.mkdir()
    (template/'template.txt').write_text('template content')
    invoked = tmp_path/'copy-invoked'
    result = powershell.run(copy_boundary(root,template,invoked) + build_step_command(folder_step(root,template,incomplete=True)))
    assert result.return_code == 0
    assert invoked.exists()
    assert (root/'template.txt').read_text() == 'template content'


def test_completed_folder_restart_preserves_existing_content(tmp_path,powershell):
    root = tmp_path/'project'
    template = tmp_path/'template'
    root.mkdir(); template.mkdir()
    (root/'template.txt').write_text('operator content')
    (template/'template.txt').write_text('template content')
    invoked = tmp_path/'copy-invoked'
    result = powershell.run(copy_boundary(root,template,invoked) + build_step_command(folder_step(root,template,incomplete=False)))
    assert result.return_code == 0
    assert result.stdout == 'FOLDER_EXISTS'
    assert (root/'template.txt').read_text() == 'operator content'
    assert not invoked.exists()
