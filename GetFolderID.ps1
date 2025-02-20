if ($psISE -and $psISE.CurrentPowerShellTab -and $psISE.CurrentPowerShellTab.Output) {
    $psISE.CurrentPowerShellTab.Output.Clear()
} else {
    Clear-Host
}


# Prompt user for input
Write-Host "`nEnter mailbox identity (e.g., DISCOON1)" -ForegroundColor Yellow
$mailboxIdentity = Read-Host 

Write-Host "Enter folder name to search (e.g., '99 Sensitivity Test')" -ForegroundColor Cyan

$folderName = Read-Host

# Clear screen after input
Clear-Host

# Show processing message
Write-Host "`nProcessing folder information..." -ForegroundColor Yellow
Write-Host "Please wait while we convert the FolderID...`n" -ForegroundColor Cyan

# Get folder statistics and filter by name
$folders = Get-MailboxFolderStatistics -Identity $mailboxIdentity | 
           Where-Object { $_.Name -eq $folderName }

# Initialize variables
$folderQueries = @()
$encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
$nibbler = $encoding.GetBytes("0123456789ABCDEF")

foreach ($folderStatistic in $folders) {
    # Show progress bar
    Write-Progress -Activity "Converting Folder ID" -Status "Processing $($folderStatistic.FolderPath)" -PercentComplete 50

    $folderId = $folderStatistic.FolderId
    $folderPath = $folderStatistic.FolderPath
    
    # Convert and process the folder ID
    $folderIdBytes = [Convert]::FromBase64String($folderId)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0
    
    $folderIdBytes | Select-Object -Skip 23 -First 24 | ForEach-Object {
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]
        $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0x0F]
    }
    
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
    
    # Create output object
    $folderStat = [PSCustomObject]@{
        FolderPath  = $folderPath
        FolderQuery = $folderQuery
    }
    
    $folderQueries += $folderStat
}

# Clear progress bar and previous messages
Clear-Host

# Display results
if ($folderQueries.Count -gt 0) {
    Write-Host "Conversion Results:`n" -ForegroundColor Green
    
    $folderQueries | ForEach-Object {
        Write-Host ("[Path]  ") -NoNewline -ForegroundColor Cyan
        Write-Host $_.FolderPath -ForegroundColor White
        
        Write-Host ("[Query] ") -NoNewline -ForegroundColor Green
        Write-Host $_.FolderQuery -ForegroundColor White
        Write-Host ("-" * ($Host.UI.RawUI.BufferSize.Width - 1)) -ForegroundColor DarkGray
    }
}
else {
    Write-Host "`nNo matching folders found!" -ForegroundColor Red
}
