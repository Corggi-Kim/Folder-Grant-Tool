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
    assert '직접 가입' in checks['membership']['message']


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
