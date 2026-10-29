"""Error identity checks shared by read-only validation and group preparation."""

AD_IDENTITY_MISSING_FUNCTION = r'''
function IsAdIdentityMissing($record) {
    # Windows PowerShell AD remoting can replace the native exception type.
    # Match the AD error identity, never localized text or ObjectNotFound alone.
    if ([string]$record.FullyQualifiedErrorId -match '(^|[.:,])ADIdentityNotFoundException(,|$)') { return $true }
    $exception = $record.Exception
    for ($depth = 0; $null -ne $exception -and $depth -lt 8; $depth++) {
        $names = @($exception.GetType().FullName) + @($exception.PSObject.TypeNames)
        if ($exception.SerializedRemoteException) {
            $names += @($exception.SerializedRemoteException.PSObject.TypeNames)
        }
        if (@($names | Where-Object { $_ -match '(^|\.)ADIdentityNotFoundException$' }).Count -gt 0) { return $true }
        $exception = $exception.InnerException
    }
    return $false
}
'''
