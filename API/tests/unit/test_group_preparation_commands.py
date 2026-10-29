from types import SimpleNamespace
from folder_grant_api.infrastructure.powershell_commands import build_step_command


def test_group_preparation_only_creates_after_specific_missing_error_and_checks_paths():
    command=build_step_command(SimpleNamespace(step_type='prepare_isolated_group',target="LSK 26-012 Isolated",details={
        'group_ou_path':"OU=Groups,DC=example,DC=org", 'description_source_group':'LSK 26-012',
        'required_paths':[r'\\server\share\26012',r'\\server\share\26012\Isolated']}))
    assert command.count('Get-Item -LiteralPath')==2
    assert command.index('Get-Item') < command.index('New-ADGroup')
    assert 'if (-not (IsAdIdentityMissing $_)) { throw }' in command
    assert '-ErrorAction SilentlyContinue' not in command
    assert "if ([string]$group.GroupCategory -ne 'Security')" in command
    assert "-GroupCategory Security -GroupScope DomainLocal" in command
    assert "-Identity 'LSK 26-012' -Properties Description -ErrorAction Stop" in command
    assert "catch { $description='' }" in command
    assert 'Start-Sleep -Seconds 15' in command
    assert all(x not in command for x in ['New-Item','robocopy','Add-ADGroupMember'])


def test_base_acl_uses_security_group_sid_and_plain_rx():
    command=build_step_command(SimpleNamespace(step_type='grant_isolated_base_acl',target="\\\\server\\share\\O'Brien",details={'principal':'LSK 26-012 Isolated'}))
    assert "O''Brien" in command
    assert "GroupCategory -ne 'Security'" in command
    assert "$sidPrincipal='*' + $groupSid" in command
    assert " /grant ($sidPrincipal + ':RX')" in command
    assert all(x not in command for x in ['/T','/C','(CI)','(OI)','New-Item'])
