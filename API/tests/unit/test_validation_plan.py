from folder_grant_api.settings import Settings
from folder_grant_api.schemas.access_job import AccessJobRequest
from folder_grant_api.schemas.project_job import ProjectJobRequest
from folder_grant_api.api.routes.preview import create_plan
from folder_grant_api.api.routes.project_jobs import create_project_plan
from folder_grant_api.services.validation_service import access_probes, project_probes, READ_ONLY_SCRIPT


def request(**overrides):
    data=dict(request_id='X',source='swagger',operation='grant',project_status='progress',employee_id='E1',project_code='1987',level2='Study',level3='STAT',role='Manager',requested_by='op')
    data.update(overrides)
    return AccessJobRequest(**data)


def test_access_checks_all_levels_and_legacy_role_numbers():
    s=Settings();p=request();checks=access_probes(p,create_plan(p,s),s);by_id={c['id']:c for c in checks}
    assert by_id['level1']['target'].endswith('19087')
    assert by_id['level2']['target'].endswith(r'19087\Study')
    assert by_id['level3']['target'].endswith(r'19087\Study\STAT')
    assert by_id['group']['target']=='LSK 19-087'
    assert next(c for c in checks if c['type']=='children')['numbers']==[3,4,5,8]


def test_new_role_folders_and_isolated_exclusions_follow_execution_plan():
    s=Settings();p=request(project_code='26012');checks=access_probes(p,create_plan(p,s),s)
    assert any(c['target'].endswith(r'STAT\8.Verification') for c in checks)
    p=request(project_code='26012',level2='Isolated',level3='STAT_IDMC',role='Trial STAT/SP')
    checks=access_probes(p,create_plan(p,s),s)
    assert next(c for c in checks if c['type']=='children')['excluded']==['8.Verification']


def test_closed_candidates_and_etc_do_not_require_irrelevant_groups_or_folders():
    s=Settings();p=request(project_status='closed');checks=access_probes(p,create_plan(p,s),s)
    assert 'group' not in {c['id'] for c in checks}
    assert len(next(c for c in checks if c['type']=='closed')['candidates'])==2
    p=request(level3='ETC',role=None);checks=access_probes(p,create_plan(p,s),s)
    assert not any(c['type']=='path' for c in checks)


def test_new_project_checks_creation_prerequisites_and_allows_absent_targets():
    s=Settings();p=ProjectJobRequest(request_id='P',source='swagger',project_code='26012',requested_by='op')
    checks=project_probes(p,create_project_plan(p,s),s);by_id={c['id']:c for c in checks}
    assert by_id['group']['missing_ok'] and by_id['project_root']['missing_ok']
    assert by_id['template']['when_missing']=='project_root'
    assert by_id['ou']['when_missing']=='group'
    assert by_id['study']['missing_ok_if']=='project_root'
    assert not any('All' in c['target'] for c in checks)


def test_script_distinguishes_unknown_and_missing_and_does_not_recurse():
    assert 'ADIdentityNotFoundException|ItemNotFoundException' in READ_ONLY_SCRIPT
    assert '선행 대상이 없거나 확인되지 않아' in READ_ONLY_SCRIPT
    assert '-Recurse' not in READ_ONLY_SCRIPT



def test_isolated_preparation_probes_only_run_after_missing_group():
    settings=Settings();payload=request(level2='Isolated',level3='ETC',role=None)
    checks=access_probes(payload,create_plan(payload,settings),settings)
    preparation_ids={'ou','preparation_share','preparation_root','preparation_isolated'}
    assert {p['id'] for p in checks if p.get('when_group_missing')}==preparation_ids
    assert all(p.get('when_group_missing') for p in checks if p['type'] in {'path','ou'})
    guard="if ($p.when_group_missing -and $states['group'] -ne 'failed')"
    assert guard in READ_ONLY_SCRIPT
    assert READ_ONLY_SCRIPT.index(guard)<READ_ONLY_SCRIPT.index('    if ($p.when_missing)')
    guarded=READ_ONLY_SCRIPT.split(guard,1)[1].split('    if ($p.when_missing)',1)[0]
    assert "'not_applicable'" in guarded and 'continue' in guarded
