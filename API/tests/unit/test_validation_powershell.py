"""Optional real PowerShell interpreter tests with stub AD/SMB cmdlets only.
Set FGT_TEST_POWERSHELL to a pwsh/powershell executable to enable.
"""
import base64
import json
import os
import subprocess
import pytest
from folder_grant_api.settings import Settings
from folder_grant_api.schemas.access_job import AccessJobRequest
from folder_grant_api.schemas.project_job import ProjectJobRequest
from folder_grant_api.api.routes.preview import create_plan
from folder_grant_api.api.routes.project_jobs import create_project_plan
from folder_grant_api.services.validation_service import access_probes, project_probes, READ_ONLY_SCRIPT
PWSH=os.environ.get('FGT_TEST_POWERSHELL')
pytestmark=pytest.mark.skipif(not PWSH,reason='Set FGT_TEST_POWERSHELL for interpreter tests with stub cmdlets')
STUBS=r'''
$ErrorActionPreference='Stop'
$fixture=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__FIXTURE__')) | ConvertFrom-Json
function Get-Item {
 [CmdletBinding()] param($LiteralPath)
 if (@($fixture.denied) -contains $LiteralPath) { throw [UnauthorizedAccessException]::new('access denied') }
 if (@($fixture.paths) -notcontains $LiteralPath) { throw [System.Management.Automation.ItemNotFoundException]::new('not found') }
 return [pscustomobject]@{FullName=$LiteralPath;PSIsContainer=$true}
}
function Get-ADUser {
 [CmdletBinding()] param($Identity,$Properties)
 return [pscustomobject]@{SID='user-sid';Enabled=$true;DistinguishedName='user-dn'}
}
function Get-ADGroup {
 [CmdletBinding()] param($Identity,$Properties)
 if ($fixture.group_error) {
   $id = switch ($fixture.group_error) {
     'remote_missing' { 'Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException,Microsoft.ActiveDirectory.Management.Commands.GetADGroup' }
     'denied' { 'Microsoft.ActiveDirectory.Management.ADUnauthorizedAccessException,Microsoft.ActiveDirectory.Management.Commands.GetADGroup' }
     'connection' { 'Microsoft.ActiveDirectory.Management.ADServerDownException,Microsoft.ActiveDirectory.Management.Commands.GetADGroup' }
     'other_missing' { 'UnrelatedObjectNotFound,Microsoft.ActiveDirectory.Management.Commands.GetADGroup' }
   }
   # Remoting may preserve the error ID while exposing a generic exception.
   throw [System.Management.Automation.ErrorRecord]::new([Exception]::new('AD lookup failed'),$id,[System.Management.Automation.ErrorCategory]::ObjectNotFound,$Identity)
 }
 if ($fixture.group_missing) { throw [System.Management.Automation.ItemNotFoundException]::new('group missing') }
 return [pscustomobject]@{SID='group-sid';GroupCategory='Security';member=@('user-dn')}
}
function Get-ADOrganizationalUnit {
 [CmdletBinding()] param($Identity)
 if ($fixture.deny_ou) { throw [UnauthorizedAccessException]::new('OU must not be queried') }
 return [pscustomobject]@{DistinguishedName=$Identity}
}
function Get-Acl {
 [CmdletBinding()] param($LiteralPath)
 if (@($fixture.denied_acl) -contains $LiteralPath) { throw [UnauthorizedAccessException]::new('ACL denied') }
 return [pscustomobject]@{Access=@()}
}
function Get-ChildItem {
 [CmdletBinding()] param($LiteralPath,[switch]$Directory)
 return @($fixture.children | ForEach-Object { [pscustomobject]@{Name=$_;FullName=($LiteralPath+'\'+$_)} })
}
'''


def run_script(probes,paths,**options):
    fixture=dict(paths=paths,denied=[],denied_acl=[],group_missing=False,children=[]);fixture.update(options)
    encoded=base64.b64encode(json.dumps(dict(probes=probes)).encode()).decode()
    fake=base64.b64encode(json.dumps(fixture).encode()).decode()
    script=STUBS.replace('__FIXTURE__',fake)+'\n'+READ_ONLY_SCRIPT.replace('__SPEC__',encoded)
    script=script.replace('[Security.Principal.WindowsIdentity]::GetCurrent().Name',"'TEST\\stub'")
    env=dict(os.environ,XDG_CACHE_HOME='/tmp/fgt-ps-cache',XDG_CONFIG_HOME='/tmp/fgt-ps-config',XDG_DATA_HOME='/tmp/fgt-ps-data')
    result=subprocess.run([PWSH,'-NoProfile','-NonInteractive','-EncodedCommand',base64.b64encode(script.encode('utf-16-le')).decode()],capture_output=True,timeout=15,env=env)
    assert result.returncode==0,result.stderr.decode(errors='replace')
    return {c['id']:c for c in json.loads(result.stdout.decode('utf-8-sig'))['checks']}


def access():
    s=Settings();p=AccessJobRequest(request_id='X',source='swagger',operation='grant',project_status='progress',employee_id='E1',project_code='26012',level2='Study',level3='DM',requested_by='op')
    probes=access_probes(p,create_plan(p,s),s)
    return probes,[p['target'] for p in probes if p['type']=='path']


def project():
    s=Settings();p=ProjectJobRequest(request_id='P',source='swagger',project_code='26012',requested_by='op')
    probes=project_probes(p,create_project_plan(p,s),s)
    return probes,{p['id']:p['target'] for p in probes}


def test_access_real_script_with_resources_present():
    probes,paths=access();checks=run_script(probes,paths)
    assert all(c['status']=='passed' for c in checks.values())
    assert checks['membership']['message'] == '그룹에 등록되어 있습니다.'


def test_missing_level3_vs_unreadable_share():
    probes,paths=access();checks=run_script(probes,paths[:-1])
    assert checks['level3']['status']=='failed' and checks['acl']['status']=='unknown'
    checks=run_script(probes,paths,denied=[paths[0]])
    assert checks['share']['status']=='unknown' and checks['level1']['status']=='unknown'


def test_acl_denied_is_unknown():
    probes,paths=access();checks=run_script(probes,paths,denied_acl=[paths[-1]])
    assert checks['level3']['status']=='passed' and checks['acl']['status']=='unknown'


def test_new_project_requires_template_structure_and_allows_absent_targets():
    probes,paths=project();present=[paths[k] for k in ('share','template','template_study')]
    checks=run_script(probes,present,group_missing=True)
    assert checks['group']['status']=='passed' and checks['study']['status']=='passed'
    assert checks['root_acl']['status']=='not_applicable'
    assert all(c['status'] in ('passed','not_applicable') for c in checks.values())
    checks=run_script(probes,present[:-1],group_missing=True)
    assert checks['template_study']['status']=='failed'


def test_existing_project_does_not_require_deleted_template():
    probes,paths=project();checks=run_script(probes,[paths[k] for k in ('share','project_root','study')])
    assert checks['template']['status']=='not_applicable'
    assert checks['root_acl']['status']=='passed'



def test_existing_isolated_etc_group_never_queries_preparation_ou_or_paths():
    settings=Settings()
    payload=AccessJobRequest(request_id='I',source='swagger',operation='grant',project_status='progress',employee_id='E1',project_code='26012',level2='Isolated',level3='ETC',requested_by='op')
    probes=access_probes(payload,create_plan(payload,settings),settings)
    # Every folder is denied and the OU stub throws if invoked. ETC still passes.
    paths=[p['target'] for p in probes if p['type']=='path']
    checks=run_script(probes,[],denied=paths,deny_ou=True)
    for id in ('ou','preparation_share','preparation_root','preparation_isolated'):
        assert checks[id]['status']=='not_applicable'
    assert all(c['status'] in ('passed','not_applicable') for c in checks.values())


def isolated_resources(group_error):
    from folder_grant_api.services.validation_service import with_group_preparation
    from folder_grant_api.schemas.validation import ValidationResponse, ValidationCheck
    from datetime import datetime, timezone
    settings = Settings(executor_mode='powershell')
    payload = AccessJobRequest(request_id='ISOLATED', source='swagger', operation='grant',
        project_status='progress', employee_id='E1', project_code='26086', level2='Isolated',
        level3='PV', requested_by='op')
    plan = create_plan(payload, settings)
    probes = access_probes(payload, plan, settings)
    checks = run_script(probes, [p['target'] for p in probes if p['type']=='path'], group_error=group_error)
    statuses = {c['status'] for c in checks.values()}
    result = ValidationResponse(request_id=payload.request_id, executor_mode='powershell',
        validation_status='failed' if 'failed' in statuses else 'unknown',
        checked_at=datetime.now(timezone.utc), checks=[ValidationCheck(**c) for c in checks.values()], warnings=[])
    return checks, with_group_preparation(payload, plan, settings, result)


def test_remote_ad_missing_runs_preparation_checks_and_offers_group_creation():
    checks, result = isolated_resources('remote_missing')
    assert checks['group']['status'] == 'failed'
    assert checks['group']['message'] == '대상 없음'
    assert all(checks[id]['status']=='passed' for id in ('ou','preparation_share','preparation_root','preparation_isolated'))
    assert result.group_preparation.group_name == 'LSK 26-086 Isolated'


@pytest.mark.parametrize('group_error', ['denied','connection','other_missing'])
def test_ad_lookup_errors_never_offer_group_creation(group_error):
    checks, result = isolated_resources(group_error)
    assert checks['group']['status'] == 'unknown'
    assert result.group_preparation is None


@pytest.mark.parametrize('group_error,created', [('remote_missing',True),('denied',False),('connection',False),('other_missing',False)])
def test_group_creation_handles_remote_missing_but_never_other_errors(group_error, created):
    from types import SimpleNamespace
    from folder_grant_api.infrastructure.powershell_commands import build_step_command
    command = build_step_command(SimpleNamespace(step_type='prepare_isolated_group',
        target='LSK 26-086 Isolated',details=dict(group_ou_path='OU=Groups,DC=example,DC=org',
        description_source_group='LSK 26-086',required_paths=['ROOT','ISOLATED'])))
    fixture = base64.b64encode(json.dumps(dict(paths=['ROOT','ISOLATED'],group_error=group_error)).encode()).decode()
    script = STUBS.replace('__FIXTURE__', fixture) + """
function New-ADGroup { [CmdletBinding()] param($Name,$SamAccountName,$GroupCategory,$GroupScope,$Path,$Description) Write-Output 'CREATED_MARKER' }
function Start-Sleep { param($Seconds) }
""" + command
    env=dict(os.environ,XDG_CACHE_HOME='/tmp/fgt-ps-cache',XDG_CONFIG_HOME='/tmp/fgt-ps-config',XDG_DATA_HOME='/tmp/fgt-ps-data')
    result = subprocess.run([PWSH,'-NoProfile','-NonInteractive','-EncodedCommand',base64.b64encode(script.encode('utf-16-le')).decode()],capture_output=True,timeout=15,env=env)
    assert ('CREATED_MARKER' in result.stdout.decode('utf-8-sig')) == created, result.stderr.decode(errors='replace')
    assert (result.returncode == 0) == created
    if not created:
        assert 'AD lookup failed' in result.stderr.decode(errors='replace')


def test_ad_missing_classifier_supports_native_wrapped_and_serialized_errors():
    from folder_grant_api.infrastructure.powershell_errors import AD_IDENTITY_MISSING_FUNCTION
    script = r'''
Add-Type 'namespace Microsoft.ActiveDirectory.Management { public class ADIdentityNotFoundException : System.Exception { public ADIdentityNotFoundException(string message) : base(message) {} } }'
$native = [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]::new('missing')
$wrapped = [Exception]::new('wrapper', $native)
$record = [System.Management.Automation.ErrorRecord]::new($native,'GenericError',[System.Management.Automation.ErrorCategory]::ObjectNotFound,$null)
$serialized = [System.Management.Automation.PSSerializer]::Deserialize([System.Management.Automation.PSSerializer]::Serialize($record))
''' + AD_IDENTITY_MISSING_FUNCTION + r'''
@(
 (IsAdIdentityMissing ([pscustomobject]@{Exception=$native})),
 (IsAdIdentityMissing ([pscustomobject]@{Exception=$wrapped})),
 (IsAdIdentityMissing $serialized),
 (IsAdIdentityMissing ([pscustomobject]@{Exception=[UnauthorizedAccessException]::new('ADIdentityNotFoundException')})),
 (IsAdIdentityMissing ([pscustomobject]@{FullyQualifiedErrorId='NotADIdentityNotFoundException,GetADGroup';Exception=[Exception]::new('missing')}))
) | ConvertTo-Json -Compress
'''
    env=dict(os.environ,XDG_CACHE_HOME='/tmp/fgt-ps-cache',XDG_CONFIG_HOME='/tmp/fgt-ps-config',XDG_DATA_HOME='/tmp/fgt-ps-data')
    result=subprocess.run([PWSH,'-NoProfile','-NonInteractive','-EncodedCommand',base64.b64encode(script.encode('utf-16-le')).decode()],capture_output=True,timeout=15,env=env)
    assert result.returncode==0, result.stderr.decode(errors='replace')
    assert json.loads(result.stdout.decode('utf-8-sig')) == [True,True,True,False,False]
