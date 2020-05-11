<#
V1.0

This script will:
1. Create a TempFile within the user's OneDrive folder (if found)
2. Will wait for $global:TestCycleSleepSeconds
3. Will check if the file is in sync by finding the TempFileName hash in
   the metadata and ensure that the file's attribute ReparsePoint was successfully set
4. Will send the results by mail (including log as an attachment)
#>

# Get current directory and set import file in variable 
$path = Split-Path -parent $MyInvocation.MyCommand.Definition 
$date = $(Get-Date -format "yyyyMMdd_HHmmss")
$log =  "$path\$($MyInvocation.MyCommand.Name)_$date.log"

# Enable old log cleanup?
$logsCleanupEnabled = $false
$logsCleanupOlderThan = 14

if ($logsCleanupEnabled) {
    $limit = (Get-Date).AddDays(-$logsCleanupOlderThan)
    Write-Debug "Removing logs older than $limit"
    # Delete files older than the $limit.
    try {
        Get-ChildItem -Path $path -Recurse -Force -File -Filter "$($MyInvocation.MyCommand.Name)_*.log" `
            | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } `
            | Remove-Item -Force -Confirm:$false -Verbose
    } catch {
        Write-Warning $_.Exception.Message
    }
}

try {
    Start-Transcript -Path $log -Force | Out-Null
} catch {
    Write-Warning $_.Exception.Message
}

<# 
Random file name and size
    Internally, GetRandomFileName uses RNGCryptoServiceProvider to generate 11-character (name:8+ext:3) string.
    The string represents a base-32 encoded number, so the total number of possible strings is 3211 or 255.
    Assuming uniform distribution, the chances of making a duplicate are about 2-55, or 1 in 36 quadrillion.
    That's pretty low: for comparison, your chances of winning NY lotto are roughly one million times higher.
#>
$global:testFileName = [System.IO.Path]::GetRandomFileName()
$global:testFileNameSize = 42KB

# OD Paths
$global:oneDrivePath = "$env:USERPROFILE\OneDrive"
$global:oneDriveAppDataPath = "$env:USERPROFILE\AppData\Local\Microsoft\OneDrive"
$global:DebugPreference = "Continue" #SilentlyContinue

# Test cycles parameters
$global:TestCycleSleepSeconds = 10
$global:TestCycles = 6

# MAIL & MESSAGING
# Global switch enable/disable SMTP send
$global:SMTPSendEmail = $true
# Send email on Error Only ?
$global:SMTPSendEmailOnErrorOnly = $true
$global:SMTPServer = "smtp.example.com"
$global:SMTPPort = 25
$global:SMTPFrom = "onedrive@example.com"
$global:SMTPTo = "admin@example.com"
$global:SMTPReplyTo = "helpdesk@example.com"
$global:SMTPCc = @("example@example.com", "example@example.com")
$global:SMTPSubject = "OneDriveTest@$env:COMPUTERNAME ({0})"
$global:SMTPBody = ""
# It's not secure to store user/password here
#$global:SMTPAuthUsername = ""
#$global:SMTPAuthPassword = ""

$returnCodes = @{
    InSync                       = @{ Id = 0; Desc = "Is in sync" }
    NotInSync                    = @{ Id = 1; Desc = "Sync failed" }
    OneDriveFolderIsMissing      = @{ Id = 2; Desc = "OneDrive folder is missing" }
    OneDriveAppDataPathIsMissing = @{ Id = 3; Desc = "OneDrive AppData is missing" }
    DatFilesWereFound            = @{ Id = 4; Desc = "Dat file(s) were found" }
    DatFilesWereNotFound         = @{ Id = 5; Desc = "Dat file(s) were not found" }
}

function SendEmail {
    [cmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
    param (
        [Parameter(Mandatory=$True)]
        [string]$From,

        [Parameter(Mandatory=$False)]
        [string]$ReplyTo,

        [Parameter(Mandatory=$True)]
        [string[]]$To,

        [Parameter(Mandatory=$False)]
        [string[]]$CC,

        [Parameter(Mandatory=$True)]
        [string]$Subject,

        [Parameter(Mandatory=$True)]
        [string]$Body,

        [Parameter(Mandatory=$False)]
        [string[]]$Files,

        [Parameter(Mandatory=$False)]
        [string]$smtpServer = $global:SMTPServer,

        [Parameter(Mandatory=$False)]
        [int]$smtpPort = $global:SMTPPort,

        [Parameter(Mandatory=$False)]
        [System.Net.Mail.MailPriority]$Priority = [System.Net.Mail.MailPriority]::Normal
    )

    #Creating a Mail object
    $msg = New-Object System.Net.Mail.MailMessage

    #Creating SMTP server object
    $SMTPClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)

    # Files to attachments
    if ($Files.Count -gt 0 ) {
        $Files | %{
            $att = New-Object Net.Mail.Attachment($_)
            $msg.Attachments.Add($att)
        }
    }

    # To
    if ($To.Count -gt 0 ) {
        $To | %{
            $msg.To.Add($_)
        }
    }

    #CC
    if ($CC.Count -gt 0 ) {
        $CC | %{
            $msg.CC.Add($_)
        }
    }

    #Email structure
    $encoding = [System.Text.Encoding]::UTF8
    $msg.From = $From
    $msg.ReplyTo = $ReplyTo
    $msg.Subject = $Subject
    $msg.Body = $Body
    $msg.IsBodyHtml = $true
    $msg.Priority = $Priority
    $msg.BodyEncoding = $encoding
    $msg.SubjectEncoding = $encoding

    Write-Debug @"
SMTP:
    From: $($msg.From.ToString())
    To: $($msg.To.ToString())
    CC: $($msg.CC.ToString())
    ReplyTo: $($msg.ReplyTo.ToString())
    Subject: $($msg.Subject.ToString())
    Attach: $((($msg.Attachments | %{$_.Name}) -join ", ").ToString())
"@

    Write-Debug "Trying to send mail _via_ $smtpServer`:$smtpPort "
    #Sending email
    try {

        # SSL if needed
        #$SMTPClient.EnableSsl = $True
        # SMTP AUTH
        #$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($global:SMTPAuthUsername, $global:SMTPAuthPassword)

        $SMTPClient.Send($msg)
    } catch [Exception] {
        throw $_.Exception.Message
    } finally {
        #Dispose
        $msg.Attachments | %{
            try { $_.Dispose() } catch {}
        }
    }
}

Function PreFlightCheck {
    if (-not (Test-Path $global:oneDrivePath)) {
        throw $returnCodes.OneDriveFolderIsMissing.Desc
    }
    if (-not (Test-Path $global:oneDriveAppDataPath)) {
        throw $returnCodes.OneDriveAppDataPathIsMissing.Desc
    }
    if (-not (Test-Path $global:oneDriveAppDataPath)) {
        throw $returnCodes.OneDriveAppDataPathIsMissing.Desc
    }
    
    $datFiles = $false
    $datFiles = gci -Path "$global:oneDriveAppDataPath\settings\Personal" -File -Force -Filter *.dat -ErrorAction SilentlyContinue

    if ($datFiles -and $datFiles.Count -gt 0) {
        Write-Debug "$($returnCodes.DatFilesWereFound.Desc)`n`t-- $($datFiles -join "`n`t-- ")"
        #Write-Debug 
        $global:DatFiles = $datFiles.FullName
    } else {
        throw $returnCodes.DatFilesWereNotFound.Desc
    }
}

Function ConvertToHexSearch([string]$string) {
    try {
        ([char[]]$string -join "`0").ToString()
    } catch {
        throw $_.Exception.Message
    }
}

Function BinaryFindInFile {
    [cmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
    param (
        [Parameter(Mandatory=$True)]
        [string[]]$path,

        [Parameter(Mandatory=$True)]
        [string]$stringToSearch,

        [Parameter(Mandatory=$False)]
        [System.Text.Encoding]$enc = [System.Text.Encoding]::ASCII
    )

    $numberOfBytesToRead = 1000000
    $stringToSearch = (ConvertToHexSearch $stringToSearch)
    
    foreach ($file in $path) {
        try {
            Write-Debug "Searching for '$stringToSearch' in '$file'`n$($stringToSearch | Format-Hex)"
            $fileStream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)

            # binary reader to search for the string 
            $binaryReader = New-Object System.IO.BinaryReader($fileStream)

            # get the contents of the beginning of the file
            [Byte[]] $byteArray = $binaryReader.ReadBytes($numberOfBytesToRead)

            # look for string
            $m = [Regex]::Match([Text.Encoding]::ASCII.GetString($byteArray), $stringToSearch)
            if ($m.Success) {    
                Write-Debug "Found '$stringToSearch' at position $($m.Index)"
                return [bigint]$m.Index
            } else {
                Write-Debug "'$stringToSearch' was not found in $file"
            }
        } catch {
            throw $_.Exception.Message
        } finally {
            try { $fileStream.Close() } catch {}
        }
    }
}

function TestSync ($path) {
    $result = $false
    try {
        Write-Debug "Creating test file '$path'"
        try {
            # First - we'll try to create a hidden file within OD folder
            $out = New-Object byte[] $global:testFileNameSize
            (New-Object Random).NextBytes($out)
            [IO.File]::WriteAllBytes($path, $out)
            [System.IO.File]::SetAttributes($path, [System.IO.FileAttributes]::Hidden)
        } catch {
            throw $_.Exception.Message
        }

        # Now we will look if sync is working by searching it inside the metadata storage and also checking the attribute
        $timeout = $metaFound = $attribFound = $false
        $i = 0
        while (-not $timeout -and (-not $metaFound -and -not $attribfound)) {
            # Sleep for a bit
            $i++
            Write-Debug "Sleeping for $global:TestCycleSleepSeconds`s (cyclye $i of $global:TestCycles)"
            Start-Sleep -Seconds $global:TestCycleSleepSeconds

            $metaFound = BinaryFindInFile -path $global:DatFiles -stringToSearch $global:testFileName

            # And just to make sure we'll check the attributes
            if ([bool]$metaFound) {
                try {
                    $attributes = [System.IO.File]::GetAttributes($path).ToString()
                    if ($attributes -match "ReparsePoint") {
                        Write-Debug "Attrib 'ReparsePoint' found at '$path'"
                        $attribFound = $true
                    }
                } catch {
                    Write-Error $_.Exception.Message
                }
            }
            
            if ($i -ge $global:TestCycles) {
                Write-Debug "Cycle timeout reached"
                $timeout = $true
            }
        }

        # switch return code
        switch ($true) {
            $timeout     {$result = $returnCodes.NotInSync}
            $metaFound   {$result = $returnCodes.InSync}
            $attribFound {$result = $returnCodes.InSync}
            Default {
                $result = $returnCodes.NotInSync
            }
        }
        # return result
        return $result

    } catch {
        Write-Error $_.Exception.Message
    } finally {
        try {
            Start-Sleep -Seconds 1
            Write-Debug "Deleting test file '$path'"
            [System.IO.File]::Delete($path)
        } catch {}
    }
}

Function Main {
    PreFlightCheck

    $result = TestSync -path "$global:oneDrivePath\$global:testFileName"

    if ($result.Id -eq 0) {
        Write-Host -ForegroundColor Green "OneDrive sync is OK"
    } else {
        Write-Host -ForegroundColor Red "Something went wrong... see the log '$log'"
    }

    try {
        Stop-Transcript | Out-Null
    } catch {}
    # Send email
    if ($global:SMTPSendEmail -and ($global:SMTPSendEmailOnErrorOnly -and $result.Id -ne 0 -or -not $global:SMTPSendEmailOnErrorOnly)) {
        Write-Debug "Sending mail..."

        $SMTPSubject = $global:SMTPSubject -f $result
        $SMTPBody = Get-Content $log | ConvertTo-HTML -Property @{Label='Text';Expression={$_}} | Out-String

        SendEmail -From $global:SMTPFrom `
                  -ReplyTo $global:SMTPReplyTo `
                  -To $global:SMTPTo `
                  -Cc $global:SMTPCc `
                  -Subject $SMTPSubject `
                  -Body $SMTPBody `
                  -Files @($log)
    } else {
        Write-Debug "SMTPSendEmail is not enabled or is set to SMTPSendEmailOnErrorOnly while there's no Errors"
    }
    Write-Host -ForegroundColor Cyan "Done!"
} Main
