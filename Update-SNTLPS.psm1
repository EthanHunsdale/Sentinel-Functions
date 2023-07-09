

Function Enable-SSL {
    Try {
        # Set TLS 1.2 (3072), then TLS 1.1 (768), then TLS 1.0 (192)
        # Use integers because the enumeration values for TLS 1.2 and TLS 1.1 won't
        # exist in .NET 4.0, even though they are addressable if .NET 4.5+ is
        # installed (.NET 4.5 is an in-place upgrade).
        [System.Net.ServicePointManager]::SecurityProtocol = 3072 -bor 768 -bor 192
    }
    Catch {
        Write-Output 'Unable to set PowerShell to use TLS 1.2 and TLS 1.1 due to old .NET Framework installed. If you see underlying connection closed or trust errors, you may need to upgrade to .NET Framework 4.5+ and PowerShell v3+.'
    }
}

$FunctionsFolder = $ENV:SystemDrive + '\Sentinel\Functions'

New-Item -ItemType Directory -Force -Path $FunctionsFolder

