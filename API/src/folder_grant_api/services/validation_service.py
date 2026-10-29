"""Live, read-only preflight. Never dispatches a Job or a mutation command."""
import base64
import json
import ntpath
from datetime import datetime, timezone
from folder_grant_api.domain.project_codes import project_segment_for_folder
from folder_grant_api.infrastructure.powershell import PowerShellRunner
from folder_grant_api.schemas.validation import ValidationCheck, ValidationResponse

NOTICE = ['검증 시점의 조회 결과입니다. 실제 권한 변경 성공을 보장하지 않습니다.',
          'ACL은 사용자·프로젝트 그룹의 직접 설정 참고 정보이며 유효 권한 전체 계산은 아닙니다.',
          '요청에 필요한 경로만 확인하며 전체 파일 트리를 재귀 검사하지 않습니다.']


def probe(id, label, target, type='path', **options):
    return dict(id=id, label=label, target=target, type=type, **options)


def share_probe(path, id='share'):
    drive = ntpath.splitdrive(path)[0]
    root = drive if drive.startswith('\\\\') else drive + '\\'
    return probe(id, '공유 경로 접근', root or ntpath.dirname(path), root=True)


def access_probes(payload, plan, settings):
    checks = [probe('user', '대상 AD 사용자', payload.employee_id, 'user')]
    if payload.project_status == 'closed':
        # Candidate roots can independently fail; the selected target is the required check.
        checks.append(probe('closed', '종료 프로젝트 폴더', plan.target_paths[0], 'closed',
                            candidates=plan.target_paths))
        checks.append(probe('closed_acl', '종료 폴더 현재 ACL', plan.target_paths[0], 'acl', parent='closed'))
        return checks
    checks.extend([probe('group', '프로젝트 보안 그룹', plan.target_group, 'group'),
                   probe('membership', '그룹 등록 여부', plan.target_group, 'membership', parents=['user', 'group'])])
    if preparation_scope(payload):
        root, isolated = preparation_paths(payload, settings)
        checks.extend([{**p, 'when_group_missing': True} for p in [
            probe('ou', '그룹 생성 OU', settings.group_ou_path, 'ou'),
            share_probe(root, 'preparation_share'),
            probe('preparation_root', '프로젝트 루트', root, parent='preparation_share'),
            probe('preparation_isolated', 'Isolated 루트', isolated, parent='preparation_root')]])
    if (payload.level3 or '').upper() == 'ETC':
        checks.append(probe('etc', '폴더 검사', '', 'not_applicable', message='ETC는 그룹 작업만 수행합니다.'))
        return checks
    root = settings.share_root.format(proj_seg=project_segment_for_folder(payload.project_code), lv2='', lv3='').rstrip('\\')
    level2 = ntpath.dirname(plan.target_paths[0])
    checks.extend([share_probe(root), probe('level1', 'Level1 프로젝트 폴더', root, parent='share'),
                   probe('level2', 'Level2 폴더', level2, parent='level1'),
                   probe('level3', 'Level3 폴더', plan.target_paths[0], parent='level2'),
                   probe('acl', 'Level3 현재 ACL', plan.target_paths[0], 'acl', parent='level3')])
    for i, step in enumerate(plan.steps):
        if step.type != 'grant_folder_acl':
            continue
        if step.target != plan.target_paths[0]:
            id = f'role_{i}'
            checks.extend([probe(id, 'Role 하위 폴더', step.target, parent='level3'),
                           probe(id+'_acl', 'Role 폴더 현재 ACL', step.target, 'acl', parent=id)])
        numbers = step.metadata.get('legacy_subfolder_numbers')
        excluded = step.metadata.get('excluded_subfolders')
        if numbers or excluded is not None:
            checks.append(probe(f'children_{i}', 'Role 하위 폴더', step.target, 'children',
                                parent='level3', numbers=numbers or [], excluded=excluded))
    return checks


def project_probes(payload, plan, settings):
    return [share_probe(plan.root_path),
            probe('group', '프로젝트 보안 그룹', plan.group_name, 'group', missing_ok=True),
            probe('ou', '그룹 생성 OU', settings.group_ou_path, 'ou', when_missing='group'),
            probe('project_root', '프로젝트 생성 경로', plan.root_path, parent='share', missing_ok=True),
            {**share_probe(plan.template_path,'template_share'), 'when_missing':'project_root'},
            probe('template', '복사할 템플릿', plan.template_path, when_missing='project_root', parent='template_share'),
            probe('template_study', '템플릿 Study 폴더', ntpath.join(plan.template_path,'Study'), parent='template', when_missing='project_root'),
            probe('study', 'Level2 Study', ntpath.dirname(plan.study_all_path), parent='project_root', missing_ok_if='project_root'),
            probe('root_acl', '프로젝트 루트 현재 ACL', plan.root_path, 'acl', parent='project_root')]


# All targets are transported as data, not interpolated PowerShell source.
READ_ONLY_SCRIPT = r'''
$spec = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__SPEC__')) | ConvertFrom-Json
$script:checks = New-Object 'System.Collections.Generic.List[object]'
$states = @{}; $objects = @{}
function Record($p, $status, $message, $target) {
    if ($null -eq $target) { $target = $p.target }
    [void]$script:checks.Add([pscustomobject]@{id=[string]$p.id;label=[string]$p.label;target=[string]$target;status=$status;message=[string]$message})
    $states[$p.id] = $status
}
function IsMissing($err) {
    return ($err.Exception.GetType().FullName -match 'ADIdentityNotFoundException|ItemNotFoundException')
}
function AclInfo($path) {
    $acl = Get-Acl -LiteralPath $path -ErrorAction Stop
    $matches = @()
    foreach ($ace in $acl.Access) {
        try { $sid = $ace.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } catch { continue }
        $who = $null
        if ($objects['user'] -and $sid -eq [string]$objects['user'].SID) { $who = '사용자' }
        if ($objects['group'] -and $sid -eq [string]$objects['group'].SID) { $who = '프로젝트 그룹' }
        if ($who) { $matches += ($who + ': ' + [string]$ace.AccessControlType + ' ' + [string]$ace.FileSystemRights + ' (상속=' + [string]$ace.IsInherited + ')') }
    }
    if ($matches.Count -eq 0) { return 'ACL 읽기 성공 / 해석 가능한 ACE 중 대상 사용자·프로젝트 그룹과 일치하는 항목 없음 (유효 권한 판정 아님)' }
    return ($matches -join '; ')
}
foreach ($p in $spec.probes) {
    if ($p.when_group_missing -and $states['group'] -ne 'failed') {
        Record $p 'not_applicable' '그룹 누락 시에만 생성 준비를 검사합니다.' $null; continue
    }
    if ($p.when_missing) {
        if ($states[$p.when_missing] -notin @('passed','warning')) { Record $p 'unknown' '선행 대상 확인 불가' $null; continue }
        if ($objects[$p.when_missing]) { Record $p 'not_applicable' '기존 자원 재사용: 생성 준비 검사 불필요' $null; continue }
    }
    $parents = @(); if ($p.parent) { $parents += $p.parent }; if ($p.parents) { $parents += $p.parents }
    $blocked = @($parents | Where-Object { $states[$_] -notin @('passed','warning') })
    if ($blocked.Count -gt 0) { Record $p 'unknown' '선행 대상이 없거나 확인되지 않아 검사하지 못했습니다.' $null; continue }
    try {
        switch ($p.type) {
            'user' {
                $user = Get-ADUser -Identity $p.target -Properties Enabled -ErrorAction Stop
                $objects[$p.id] = $user
                if ($user.Enabled) { Record $p 'passed' '계정 확인 완료' $null }
                else { Record $p 'warning' '비활성화된 계정' $null }
            }
            'group' {
                $group = Get-ADGroup -Identity $p.target -Properties member -ErrorAction Stop
                if ([string]$group.GroupCategory -ne 'Security') { Record $p 'failed' '보안 그룹이 아닙니다.' $null }
                else { $objects[$p.id] = $group; Record $p 'passed' '존재 / 기존 그룹 사용' $null }
            }
            'membership' {
                if (@($objects['group'].member) -contains $objects['user'].DistinguishedName) { Record $p 'passed' '그룹에 등록되어 있습니다.' $null }
                else { Record $p 'passed' '그룹에 등록되어 있지 않습니다.' $null }
            }
            'ou' {
                $objects[$p.id] = Get-ADOrganizationalUnit -Identity $p.target -ErrorAction Stop
                Record $p 'passed' '생성 위치 존재 / 조회 가능 (생성 권한 보장 아님)' $null
            }
            'path' {
                $item = Get-Item -LiteralPath $p.target -ErrorAction Stop
                if (-not $item.PSIsContainer) { Record $p 'failed' '폴더가 아닌 파일입니다.' $null }
                else { $objects[$p.id] = $item; Record $p 'passed' '폴더 존재 / 조회 가능' $null }
            }
            'acl' {
                if (-not $objects[$p.parent]) { Record $p 'not_applicable' '생성 예정 폴더: 현재 ACL 없음' $null; continue }
                $path = $objects[$p.parent].FullName
                Record $p 'passed' (AclInfo $path) $path
            }
            'children' {
                $children = @(Get-ChildItem -LiteralPath $p.target -Directory -ErrorAction Stop | Where-Object { $_.Name -ine 'All' })
                if (@($p.numbers).Count -gt 0) {
                    foreach ($number in $p.numbers) {
                        $child = $children | Where-Object { $_.Name -like ([string]$number + '.*') } | Select-Object -First 1
                        $q = [pscustomobject]@{id=($p.id + '_' + $number);label='Role 필수 하위 폴더';target=($p.target + '\' + $number + '.*')}
                        if (-not $child) { Record $q 'failed' '필수 번호에 해당하는 폴더 없음' $null }
                        else { try { Record $q 'passed' (AclInfo $child.FullName) $child.FullName } catch { Record $q 'unknown' $_.Exception.Message $child.FullName } }
                    }
                } else {
                    $included = @($children | Where-Object { @($p.excluded) -notcontains $_.Name })
                    foreach ($child in $included) {
                        $q = [pscustomobject]@{id=($p.id + '_' + $child.Name);label='Role 적용 하위 폴더';target=$child.FullName}
                        try { Record $q 'passed' (AclInfo $child.FullName) $null } catch { Record $q 'unknown' $_.Exception.Message $null }
                    }
                }
                Record $p 'passed' '하위 폴더 조회 완료 / 필요한 항목의 검사 결과 참조' $null
            }
            'closed' {
                $found = $null; $unknown = $false
                foreach ($candidate in $p.candidates) {
                    $q = [pscustomobject]@{id=($p.id + '_' + [array]::IndexOf([array]$p.candidates,$candidate));label='종료 경로 후보';target=$candidate}
                    try {
                        # A failed share lookup must not be interpreted as an absent project.
                        $share = ([IO.Path]::GetPathRoot($candidate)).TrimEnd('\')
                        $null = Get-Item -LiteralPath $share -ErrorAction Stop
                    } catch { Record $q 'unknown' $_.Exception.Message $null; $unknown=$true; continue }
                    try {
                        $item = Get-Item -LiteralPath $candidate -ErrorAction Stop
                        if (-not $item.PSIsContainer) { Record $q 'warning' '폴더가 아닙니다.' $null; continue }
                        Record $q 'passed' '폴더 존재 / 실행 시 이 경로 선택' $null
                        $found=$item; break
                    } catch {
                        if (IsMissing $_) { Record $q 'warning' '이 후보에는 폴더 없음' $null }
                        else { Record $q 'unknown' $_.Exception.Message $null; $unknown=$true }
                    }
                }
                if ($found -and -not $unknown) { $objects[$p.id]=$found; Record $p 'passed' '실제 대상 경로 확인' $found.FullName }
                elseif ($found) { Record $p 'unknown' '다른 후보 조회 실패: 실제 실행의 경로 선택을 확정할 수 없습니다.' $found.FullName }
                elseif ($unknown) { Record $p 'unknown' '후보 경로 조회 실패로 대상 경로 확인 불가' $null }
                else { Record $p 'failed' '종료·Archive 후보 모두 대상 폴더 없음' $null }
            }
            'not_applicable' { Record $p 'not_applicable' $p.message $null }
            default { Record $p 'unknown' '지원하지 않는 검사 종류' $null }
        }
    } catch {
        $missing = IsMissing $_
        $allowed = $p.missing_ok -or ($p.missing_ok_if -and $states[$p.missing_ok_if] -eq 'passed' -and -not $objects[$p.missing_ok_if])
        if ($missing -and $allowed) { Record $p 'passed' '현재 없음 / 신규 생성 예정 (생성 권한 보장 아님)' $null }
        elseif ($missing -and -not $p.root) { Record $p 'failed' '대상 없음' $null }
        else { Record $p 'unknown' $_.Exception.Message $null }
    }
}
[pscustomobject]@{checks=@($script:checks.ToArray());checked_by=[Security.Principal.WindowsIdentity]::GetCurrent().Name} | ConvertTo-Json -Depth 8 -Compress
'''


def validate_resources(request_id, probes, settings):
    def response(checks, checked_by=None):
        statuses = {check.status for check in checks}
        state = ('failed' if 'failed' in statuses else 'unknown' if 'unknown' in statuses
                 else 'warning' if 'warning' in statuses else 'passed')
        return ValidationResponse(request_id=request_id, executor_mode=settings.executor_mode,
                                  validation_status=state, checked_at=datetime.now(timezone.utc),
                                  checked_by=checked_by, checks=checks, warnings=NOTICE)
    def unavailable(message):
        return response([ValidationCheck(id='live_check',label='실제 자원 검증',target='',status='unknown',message=message)])
    if settings.executor_mode != 'powershell':
        return unavailable('Mock 모드에서는 실제 AD·폴더를 조회하지 않습니다. 실제 서버의 PowerShell 모드에서 검증하세요.')
    data = base64.b64encode(json.dumps(dict(probes=probes),ensure_ascii=False).encode('utf-8')).decode('ascii')
    script = READ_ONLY_SCRIPT.replace('__SPEC__',data)
    try:
        output = PowerShellRunner(settings.powershell_path, settings.validation_timeout_seconds).run(script)
        raw = json.loads(output.stdout.lstrip('\ufeff'))
        checks = [ValidationCheck.model_validate(item) for item in raw['checks']]
        ids = [check.id for check in checks]
        if not checks or len(ids) != len(set(ids)) or not {p['id'] for p in probes}.issubset(ids):
            raise ValueError('검사 결과 누락 또는 중복')
        return response(checks, raw.get('checked_by'))
    except Exception as exc:
        return unavailable(f'실제 조회를 완료하지 못했습니다: {exc}')


def preparation_scope(payload):
    return (payload.operation == 'grant' and payload.project_status == 'progress'
            and (payload.level2 or '').lower() == 'isolated')


def preparation_paths(payload, settings):
    root = settings.share_root.format(proj_seg=project_segment_for_folder(payload.project_code), lv2='', lv3='').rstrip('\\')
    isolated = settings.share_root.format(proj_seg=project_segment_for_folder(payload.project_code), lv2='Isolated', lv3='').rstrip('\\')
    return root, isolated


def with_group_preparation(payload, plan, settings, result):
    """Offer only on an unambiguous missing group with every prerequisite readable."""
    from folder_grant_api.schemas.validation import GroupPreparation
    if settings.executor_mode != 'powershell' or not preparation_scope(payload):
        return result
    checks = {c.id: c for c in result.checks}
    required = {'user', 'group', 'membership', 'ou', 'preparation_share', 'preparation_root', 'preparation_isolated'}
    if not required.issubset(checks):
        return result
    if any(checks[id].status != 'passed' for id in required - {'group', 'membership'}):
        return result
    group = checks['group']
    if group.status != 'failed' or group.message != '대상 없음':
        return result
    for c in result.checks:
        if c.id == 'group':
            continue
        if c.id == 'membership' and c.status == 'unknown' and c.message == '선행 대상이 없거나 확인되지 않아 검사하지 못했습니다.':
            continue
        if c.status not in {'passed', 'not_applicable'}:
            return result
    root, isolated = preparation_paths(payload, settings)
    result.group_preparation = GroupPreparation(group_name=plan.target_group, root_path=root, isolated_path=isolated)
    return result


def validate_access_resources(payload, plan, settings):
    result = validate_resources(payload.request_id, access_probes(payload, plan, settings), settings)
    return with_group_preparation(payload, plan, settings, result)
