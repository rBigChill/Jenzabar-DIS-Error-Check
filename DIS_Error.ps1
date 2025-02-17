if (Get-Module -ListAvailable SqlServer) {
    Import-Module -Name SqlServer
} else {
    Install-Module -Name SqlServer -Scope CurrentUser -Force
    Import-Module -Name SqlServer
}

$Computer = ''
$ErrorFileLocation = ''

$SqlServer = ""
$SqlDBName = ""
$SchemaName = ""
$TableName = ""

$smtpServer = ""
$fromEmail = ""
$toEmail = ""
$subject = ""

function Send-Email {
    param(
        [string]$body
    )
    Write-Host 'Sending email...'
    $smtpParams = @{
        From       = $fromEmail
        To         = $toEmail
        Subject    = $subject
        Body       = $body
        SmtpServer = $smtpServer
    }
    Send-MailMessage @smtpParams
    Write-Host 'Email sent...'
}

function Read-DIS_ImportLog {
    Write-Host 'Reading DIS Import Log from database'
    Invoke-Sqlcmd -ServerInstance $SqlServer -Encrypt Optional -Query "
        SELECT TOP 1 ImportLogID, StartTime, EndTime, Status, SequenceNumber, CurrentRecord, TotalRecords, ErrorCount
        FROM [ICS_NET].[dbo].[DIS_ImportLog]
        WHERE [SequenceNumber] = (SELECT MAX([SequenceNumber]) FROM [ICS_NET].[dbo].[DIS_ImportLog])
		ORDER BY StartTime DESC"
        # Status: 2 = success, 3 = failed, 1 = in-progress ???
}

function Delete-User {
    param(
        [string]$user
    )
    Write-Host "Deleting $user from databases..."
    Invoke-Sqlcmd -ServerInstance $SqlServer -Encrypt Optional -Query "
		DELETE FROM [TmsEPrd].[dbo].[TW_API_CRP] WHERE ID_Num IN ($user)
		DELETE FROM [TmsEPrd].[dbo].[TW_API_CST] WHERE ID_Num IN ($user)
        DELETE FROM [TmsEPrd].[dbo].[TW_API_PRS] WHERE ID_Num IN ($user)"
}

function Delete-Section {
    param(
        [string]$year,
        [string]$term,
        [string]$section
    )
    Write-Host "Deleting $year $term - $section from databases..."
    Invoke-Sqlcmd -ServerInstance $SqlServer -Encrypt Optional -Query "
		DELETE FROM [TmsEPrd].[dbo].[TW_API_SCH] WHERE (YR_CDE = $year AND TRM_CDE = $term) AND CRS_CDE IN ($section)
		DELETE FROM [TmsEPrd].[dbo].[TW_API_CRP] WHERE (YR_CDE = $year AND TRM_CDE = $term) AND CRS_CDE IN ($section)
        DELETE FROM [TmsEPrd].[dbo].[TW_API_CRS] WHERE (YR_CDE = $year AND TRM_CDE = $term) AND CRS_CDE IN ($section)"
}

function Repost-J1Batch {
    Write-Host 'Reposting J1 Batch...'
    Invoke-Sqlcmd -ServerInstance $SqlServer -Encrypt Optional -Query "
        USE TmsEPrd;

        GO
 
        DECLARE @JICSLastBatch INT, @J1LastBatch INT
	 
        SET @JICSLastBatch = (SELECT MAX([SequenceNumber]) FROM [ICS_NET].[dbo].[DIS_ImportLog])
        SET @J1LastBatch = (SELECT MAX(BATCH_NUMBER) FROM [TMSEPRD].[dbo].[TW_API_BATCH_STS])
 
        UPDATE tw_api_trans SET batch_number = 0 WHERE batch_number = @J1LastBatch
 
        DELETE FROM tw_api_batch_sts WHERE batch_number = @J1LastBatch
	 
        UPDATE [dbo].[TW_UI_CONFIG]
        SET [VALUE] = (@J1LastBatch - 1), [USER_NAME] = 'DIS Error Script', [JOB_NAME] = 'Error Script', JOB_TIME = getdate()
        WHERE DISPLAY_NAME = 'Last Successful Batch Number'"
}

function Stop-Services {
    Write-Host 'Stopping Services...'

	$WAS = get-service -computername $Computer -name JenzabarWebApplicationServices
    Write-Host "Stopping" $WAS.Name
	if ($WAS.Status -eq ‘Running’) {$WAS.stop()}

	$DIS = get-service -computername $Computer -name DIS
    Write-Host "Stopping" $DIS.Name
    if ($DIS.Status -eq ‘Running’) {$DIS.stop()}
}

function Start-Services {
    Write-Host 'Starting Services...'

    $count = 0

    $tableData = Read-DIS_ImportLog
    
    while ($tableData.Status -eq 3 -and $count -lt 5) {
        $WAS = get-service -computername $Computer -name JenzabarWebApplicationServices
        Write-Host "Starting" $WAS.Name
        if ($WAS.Status -ne ‘Running’) {$WAS.start()}

        Write-Host "Pausing 10 seconds while" $WAS.Name "boots..."
        Start-Sleep -Seconds 10
 
        $DIS = get-service -computername $Computer -name DIS
        Write-Host "Starting" $DIS.Name
        if ($DIS.Status -ne ‘Running’) {$DIS.start()}

        Write-Host "Pausing 60 seconds while" $DIS.Name "boots..."
        Start-Sleep -Seconds 60

        $count++
           
        $tableData = Read-DIS_ImportLog
    }
}

function Parse-File {
    param(
        [string]$fileName
    )
    Write-Host "Obtaining $fileName"
    $fileContent = Get-Content $fileName

    Write-Host 'Parsing file contents...'
    foreach ($line in $fileContent) {
        if ($line -like '*User not found*') {
            Write-Host 'User not found error initiated...'
            $userPattern = '(?<=,\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},)(\d+)(?=,)'

            $matches = $fileContent | Select-String -Pattern $userPattern -AllMatches | ForEach-Object { $_.Matches }

            foreach ($match in $matches) {
                $userID = $match.Groups[1].Value
                Write-Host "Found $userID..."
                Delete-User -user $userID
            }
        }
        if ($line -like '*Section not found*') {
            Write-Host 'Section not found error initiated...'
            $sectionPattern = '(\d{4} [A-Z]+),([A-Z]+\s\d+),([A-Z]+\d+\s[A-Z])'

            $matches = $fileContent | Select-String -Pattern $sectionPattern -AllMatches | ForEach-Object { $_.Matches }

            foreach ($match in $matches) {
                $year = $match.Groups[1].Value.Substring(0,4)
                $term = $match.Groups[1].Value.Substring(5,2) # Term
                $section = $match.Groups[2].Value + ' ' + $match.Groups[3].Value
                Write-Host "Found $year $term - $section..."
                Delete-Section -year $year -term $term -section $section
            }
        }
    }
}

function Main-Script {
    Write-Host 'Initiating script...'
    $tableData = Read-DIS_ImportLog

    Write-Host "Testing connection to $Computer..."
    if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
        if ($tableData.Status -eq 3) {
            Write-Host "Status is currently" $tableData.Status
            Stop-Services
            $file = $ErrorFileLocation + $tableData.ImportLogID + '.log'
            Parse-File -fileName $file
            Repost-J1Batch
            Start-Services
        }
    } else {
        $message = "Unable to reach $Computer"
        Write-Host $message
        Send-Email -body $message
    }
}

try {
    $count = 0

    $tableData = Read-DIS_ImportLog
    
    while ($tableData.Status -eq 3) {
        Main-Script
        $tableData = Read-DIS_ImportLog
        if ($tableData.Status -ne 3) {
            Write-Host 'Error has been corrected!'
            break
        } elseif ($tableData.Status -eq 3 -and $count -eq 2) {
            Write-Host 'Failed to resolve issue...'
            $fileLocation = $ErrorFileLocation + $tableData.ImportLogID + '.log'
            $message = "DIS quick fix did not resolve the issue.`n`nFailed file location: '$fileLocation'"
            Send-Email -body $message
            break
        } else {
            $count++
            Write-Host "Error still exist. Try # $count initiating..."
        }
    }
    Write-Host "No Error found..."
}
catch {
    $errorMessage = "Error occurred: $_"
    Write-Host $errorMessage
    Send-Email -body $errorMessage
}
