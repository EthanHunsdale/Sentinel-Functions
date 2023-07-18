Param (
    [Parameter(Mandatory = $True)]
    [String] $CSVPath,

    [Parameter(Mandatory = $False)]
    [String] $OutputPath,

    [Parameter(Mandatory = $False)]
    [String] $OutputName = "SNTL-PS-Functions"
)

Begin {
    Write-Host "Beginning File Merge"
    
    Write-Host "Creating Output File"
    
    If (-not $OutputPath) {
        $OutputPath = Get-Location
    }

    $Destination = "{0}\{1}.txt" -F $OutputPath, $OutputName
    New-Item -Path $Destination -ItemType File -Force

    Write-Host "Importing CSV Content"
    Try {
        $CSVContent = (Invoke-WebRequest $CSVPath -UseBasicParsing).Content | ConvertFrom-Csv -Delimiter ","
    }
    Catch {
        Write-Host "Error importing CSV Content"
    }
}

Process {
    $CSVContent | ForEach-Object {
        Write-Host "Importing $_.Verb"
        $TxtContent = (Invoke-WebRequest $_.URL -UseBasicParsing).Content

        Write-Host "Adding $_.Verb content to output file`n"
        Add-Content -Path $Destination -Value $TxtContent
    }
}