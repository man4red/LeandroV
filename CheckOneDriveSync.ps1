<#
V1.2

This script will:
1. Create a TempFile within the user's OneDrive folder (if found)
2. Will wait for $global:TestCycleSleepSeconds
3. Will check if the file is in sync by finding the TempFileName hash in
   the metadata and ensure that the file's attribute ReparsePoint was successfully set
4. Will send the results by mail (including log as an attachment)
#>
$global:DebugPreference = "Continue" #SilentlyContinue

# Get current directory and set import file in variable
$path = Split-Path -parent $MyInvocation.MyCommand.Definition
$date = $(Get-Date -format "yyyyMMdd_HHmmss")
$log =  "$path\$($MyInvocation.MyCommand.Name)_$date.log"

try {
    Stop-Transcript | Out-Null
} catch {
    Write-Debug $_.Exception.Message
}

try {
    Start-Transcript -Path $log -Force | Out-Null
} catch {
    Write-Debug $_.Exception.Message
}

# Enable old log cleanup?
$global:logsCleanupEnabled = $true
$global:logsCleanupOlderThan = 14

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
# Look for business acount too (Default: personal only)
$global:oneDrivePersonalAndBusiness = $false

# Test cycles parameters
$global:TestCycleSleepSeconds = 10
$global:TestCycles = 6

# MAIL & MESSAGING
# Global switch enable/disable SMTP send
$global:SMTPSendEmail = $false
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
#$global:SMTPAuthUsername = ''
#$global:SMTPAuthPassword = ''

$returnCodes = @{
    InSync                       = @{ Id = 0; Desc = "Is in sync" }
    NotInSync                    = @{ Id = 1; Desc = "Sync failed" }
    OneDriveFolderIsMissing      = @{ Id = 2; Desc = "OneDrive folder is missing" }
    OneDriveAppDataPathIsMissing = @{ Id = 3; Desc = "OneDrive AppData is missing" }
    DatFilesWereFound            = @{ Id = 4; Desc = "Dat file(s) were found" }
    DatFilesWereNotFound         = @{ Id = 5; Desc = "Dat file(s) were not found" }
    SMTPIsNotEnabled             = @{ Id = 6; Desc = "SMTPSendEmail is not enabled or is set to SMTPSendEmailOnErrorOnly while there's no Errors" }
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
        [string]$SMTPServer = $global:SMTPServer,

        [Parameter(Mandatory=$False)]
        [int]$SMTPPort = $global:SMTPPort,

        [Parameter(Mandatory=$False)]
        [System.Net.Mail.MailPriority]$Priority = [System.Net.Mail.MailPriority]::Normal
    )

    #Creating a Mail object
    $msg = New-Object System.Net.Mail.MailMessage

    #Creating SMTP server object
    $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)

    # Files to attachments
    if ($Files.Count -gt 0 ) {
        $Files | %{
            if (Test-Path $_) {
                $att = New-Object Net.Mail.Attachment($_)
                $msg.Attachments.Add($att)
            }
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

    Write-Debug "Trying to send mail _via_ $SMTPServer`:$SMTPPort"
    #Sending email
    try {

        # Uncomment SSL and auth if needed
        #$SMTPClient.EnableSsl = $True
        # SMTP AUTH
        #$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($global:SMTPAuthUsername, $global:SMTPAuthPassword)

        $SMTPClient.Send($msg)
    } catch [Exception] {
        throw $_.Exception.Message
    } finally {
        #Dispose
        $msg.Attachments | %{
            try { $_.Dispose() } catch { Write-Warning $_.Exception.Message }
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

    $datFiles = $false
    if ($global:oneDrivePersonalAndBusiness) {
        $datFiles = Get-ChildItem -Path "$global:oneDriveAppDataPath\settings" -Recurse -File -Force -Filter *.dat -EA SilentlyContinue
    } else {
        $datFiles = Get-ChildItem -Path "$global:oneDriveAppDataPath\settings\Personal" -File -Force -Filter *.dat -EA SilentlyContinue
    }

    if ($datFiles -and $datFiles.Count -gt 0) {
        Write-Debug "$($returnCodes.DatFilesWereFound.Desc)`n`t-- $($datFiles -join "`n`t-- ")"
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
    [OutputType("System.Int32")]
    param (
        [Parameter(Mandatory=$True)]
        [string[]]$path,

        [Parameter(Mandatory=$True)]
        [string]$stringToSearch,

        [Parameter(Mandatory=$False)]
        [bool]$quiet = $false,

        [Parameter(Mandatory=$False)]
        [System.Text.Encoding]$encoding = [System.Text.Encoding]::ASCII
    )

    $numberOfBytesToRead = 10000
    $stringToSearch = (ConvertToHexSearch $stringToSearch)

    foreach ($file in $path) {
        try {
            if (-not $quiet) {
                Write-Debug "Searching for '$stringToSearch' in '$file'`n$($stringToSearch | Format-Hex)"
            }
            $fileStream = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)

            # Set the position so we can read bytes from the end
            if ($fileStream.Length -gt $numberOfBytesToRead) {
                $fileStream.Position = $fileStream.Length - $numberOfBytesToRead
            }

            # binary reader to search for the string
            $binaryReader = New-Object System.IO.BinaryReader($fileStream)
            # get the contents of the beginning of the file
            [Byte[]] $byteArray = $binaryReader.ReadBytes($numberOfBytesToRead)

            # look for string
            $m = [Regex]::Match($encoding.GetString($byteArray), $stringToSearch)
            if ($m.Success) {
                if (-not $quiet) {
                    Write-Debug "Found '$stringToSearch' at position $($m.Index)"
                }
                return [int32]$m.Index
            } else {
                if (-not $quiet) {
                    Write-Debug "'$stringToSearch' was not found in $file"
                }
                return [int32]0
            }
        } catch {
            throw $_.Exception.Message
        } finally {
            try { $fileStream.Close() } catch { Write-Warning $_.Exception.Message }
        }
    }
}

Function CleanUpLogs {
    [cmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Low')]
    param (
        [Parameter(Mandatory=$True)]
        [string[]]$path,

        [Parameter(Mandatory=$True)]
        [int]$days
    )

    $limit = (Get-Date).AddDays(-$days)
    Write-Debug "Removing logs older than $limit at '$path'"
    # Delete files older than the $limit.
    foreach ($folder in $path) {
        try {
            $logs = Get-ChildItem -Path $folder -Force -File -Filter "$($MyInvocation.MyCommand.Name)_*.log" `
                | Where-Object { -not $_.PSIsContainer -and $_.CreationTime -lt $limit }
            if ($logs -and $logs.Count -gt 0) {
                Write-Debug "CleanUpLogs: $($logs.Count) were found"
                $logs | Remove-Item -Force -Verbose -WhatIf
            } else {
                Write-Debug "ClenUpLogs: no old logs were found"
            }
        } catch {
            Write-Warning $_.Exception.Message
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
            <#
                I've no idea why but sometimes it returns "Could not find file" Exception
                which doesn't make sense
                So, trying to use Set-Content instead
            #>
            #[System.IO.File]::WriteAllBytes($path, $out)
            Set-Content $path -Value (([char[]]$out) -join "") -Force
            [System.IO.File]::SetAttributes($path, [System.IO.FileAttributes]::Hidden)
        } catch {
            throw $_.Exception.Message
        }

        # Now we will look if sync is working by searching it inside the metadata storage
        $timeout = $metaFound = $attribFound = $false
        $i = 0
        $quiet = $false
        while (-not $timeout -and (-not $metaFound -and -not $attribFound)) {
            # Sleep for a bit
            $i++
            Write-Debug "Sleeping for $global:TestCycleSleepSeconds`s (cyclye $i of $global:TestCycles)"
            Start-Sleep -Seconds $global:TestCycleSleepSeconds

            if ($i -gt 1) {$quiet = $true} 
            $metaFound = BinaryFindInFile -path $global:DatFiles -stringToSearch $global:testFileName -quiet $quiet

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
        } catch { Write-Warning $_.Exception.Message }
    }
}

Function Main {
    # PreFlights
    PreFlightCheck

    # CleanupLogs
    if ($global:logsCleanupEnabled) {
        CleanUpLogs -path $path -days $global:logsCleanupOlderThan
    }

    # Test Sync and get the result
    $result = TestSync -path "$global:oneDrivePath\$global:testFileName"

    if ($result -and $result.Id -eq 0) {
        Write-Host -ForegroundColor Green "OneDrive sync is OK"
    } else {
        Write-Host -ForegroundColor Red "Something went wrong... see the log '$log'"
    }

    try {
        Stop-Transcript | Out-Null
    } catch { Write-Debug $_.Exception.Message }
    # Send email
    if ($global:SMTPSendEmail -and ($global:SMTPSendEmailOnErrorOnly -and $result.Id -ne 0 -or -not $global:SMTPSendEmailOnErrorOnly)) {

        Write-Debug "Sending mail..."
        $SMTPSubject = $global:SMTPSubject -f $result.Desc
        $SMTPBody = $global:SMTPBody
        if (Test-Path $log) {
            $SMTPBody = Get-Content $log `
                | ConvertTo-HTML -Property @{Label="$SMTPSubject";Expression={$_}} -Title $SMTPSubject | Out-String
        }
        SendEmail -From $global:SMTPFrom `
                  -ReplyTo $global:SMTPReplyTo `
                  -To $global:SMTPTo `
                  -Cc $global:SMTPCc `
                  -Subject $SMTPSubject `
                  -Body $SMTPBody `
                  -Files @($log)
    } else {
        Write-Debug $returnCodes.SMTPIsNotEnabled.Desc
    }
    Write-Host -ForegroundColor Cyan "Done!"
} Main
