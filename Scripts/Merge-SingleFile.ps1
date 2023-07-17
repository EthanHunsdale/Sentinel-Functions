Param (
    [Parameter(Mandatory = $True)]
    [String] $CSVPath,

    [Parameter(Mandatory = $True)]
    [String] $OutputPath,

    [Parameter(Mandatory = $True)]
    [String] $OutputName
)

Begin {
    Write-Host "Beginning File Merge"
    
    Write-Host "Creating Output File"
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