[CmdletBinding()]
param(
    [string]$JobsPath,

    [string]$PresetAttachmentFolder,
    [string]$CommonAttachmentFolder,
    [string]$SourceWorkbookFolder,

    [switch]$DryRun,
    [switch]$OpenDraftsFolder,
    [switch]$Interactive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:FormsReady = $false
$script:ExcelNormalizeApp = $null
$script:ExcelNormalizeDisabled = $false
$script:NormalizedAttachmentCache = @{}
$script:NormalizedTempFiles = New-Object System.Collections.Generic.List[string]
$script:NormalizedTempRoot = $null

function Resolve-RequiredPath {
    param(
        [Parameter(Mandatory = $true)][string]$PathValue,
        [Parameter(Mandatory = $true)][string]$Label
    )
    $resolved = Resolve-Path -LiteralPath $PathValue -ErrorAction Stop
    if (-not $resolved) {
        throw "$Label path could not be resolved: $PathValue"
    }
    return $resolved.Path
}

function Test-Blank {
    param([object]$Value)
    return [string]::IsNullOrWhiteSpace([string]$Value)
}

function Join-ExistingPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )
    $normalized = $RelativePath -replace '/', '\'
    $candidate = Join-Path -Path $Root -ChildPath $normalized
    if (-not (Test-Path -LiteralPath $candidate)) {
        throw "File not found: $candidate"
    }
    return (Resolve-Path -LiteralPath $candidate).Path
}

function Get-JobArray {
    param([Parameter(Mandatory = $true)]$Payload)
    if ($null -eq $Payload.jobs) {
        throw "Input JSON does not contain a 'jobs' array."
    }
    return @($Payload.jobs)
}

function Get-StringArray {
    param([object]$Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Array]) {
        return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return @() }
    return @($s)
}

function Resolve-AttachmentPath {
    param(
        [Parameter(Mandatory = $true)]$Attachment,
        [Parameter(Mandatory = $true)]$Job
    )

    $kind = [string]$Attachment.kind
    $rel = [string]$Attachment.relativePath
    $fileName = [string]$Attachment.fileName

    switch ($kind) {
        'matchedPreset' {
            if (Test-Blank $PresetAttachmentFolder) {
                throw "Job '$($Job.id)': matchedPreset attachment requires -PresetAttachmentFolder."
            }
            if (Test-Blank $rel) { $rel = $fileName }
            return Join-ExistingPath -Root $PresetAttachmentFolder -RelativePath $rel
        }
        'common' {
            if (Test-Blank $CommonAttachmentFolder) {
                throw "Job '$($Job.id)': common attachment requires -CommonAttachmentFolder."
            }
            if (Test-Blank $rel) { $rel = $fileName }
            return Join-ExistingPath -Root $CommonAttachmentFolder -RelativePath $rel
        }
        'sourceWorkbook' {
            if (Test-Blank $SourceWorkbookFolder) {
                throw "Job '$($Job.id)': sourceWorkbook attachment requires -SourceWorkbookFolder."
            }
            if (Test-Blank $rel) { $rel = $Job.sourceExcelRelativePath }
            return Join-ExistingPath -Root $SourceWorkbookFolder -RelativePath $rel
        }
        default {
            throw "Job '$($Job.id)': unknown attachment kind '$kind'."
        }
    }
}

function Add-ReplyRecipients {
    param(
        [Parameter(Mandatory = $true)]$MailItem,
        [Parameter(Mandatory = $true)][string[]]$ReplyTo
    )
    foreach ($addr in $ReplyTo) {
        if ([string]::IsNullOrWhiteSpace($addr)) { continue }
        [void]$MailItem.ReplyRecipients.Add($addr)
    }
}

function New-OutlookApplication {
    try {
        return New-Object -ComObject Outlook.Application
    } catch {
        throw "Could not start Outlook COM automation. Ensure Outlook Classic is installed on this PC. $($_.Exception.Message)"
    }
}

function Test-XlsxPath {
    param([string]$PathValue)
    if (Test-Blank $PathValue) { return $false }
    $ext = [System.IO.Path]::GetExtension([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($ext)) { return $false }
    return $ext.Equals('.xlsx', [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-ExcelNormalizationTempRoot {
    if (-not (Test-Blank $script:NormalizedTempRoot)) {
        return $script:NormalizedTempRoot
    }
    $base = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'OutlookDraftCreator-NormalizedXlsx'
    if (-not (Test-Path -LiteralPath $base -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $base -Force)
    }
    $runDir = Join-Path -Path $base -ChildPath ([Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $runDir -Force)
    $script:NormalizedTempRoot = $runDir
    return $script:NormalizedTempRoot
}

function Get-ExcelNormalizationApplication {
    if ($script:ExcelNormalizeDisabled) { return $null }
    if ($null -ne $script:ExcelNormalizeApp) { return $script:ExcelNormalizeApp }

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        try { $excel.ScreenUpdating = $false } catch { }
        try { $excel.EnableEvents = $false } catch { }
        $script:ExcelNormalizeApp = $excel
        return $script:ExcelNormalizeApp
    } catch {
        Write-Warning "Excel normalization is unavailable ($($_.Exception.Message)). Attaching original .xlsx files."
        $script:ExcelNormalizeDisabled = $true
        return $null
    }
}

function Get-NormalizedAttachmentPathForSend {
    param(
        [Parameter(Mandatory = $true)][string]$PathValue
    )

    if (-not (Test-XlsxPath -PathValue $PathValue)) {
        return $PathValue
    }

    if ($script:NormalizedAttachmentCache.ContainsKey($PathValue)) {
        return [string]$script:NormalizedAttachmentCache[$PathValue]
    }

    $excel = Get-ExcelNormalizationApplication
    if ($null -eq $excel) {
        $script:NormalizedAttachmentCache[$PathValue] = $PathValue
        return $PathValue
    }

    $tempRoot = Get-ExcelNormalizationTempRoot
    $originalFileName = [System.IO.Path]::GetFileName($PathValue)
    if ([string]::IsNullOrWhiteSpace($originalFileName)) {
        $originalFileName = 'attachment.xlsx'
    }
    $uniqueDir = Join-Path -Path $tempRoot -ChildPath ([Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $uniqueDir -Force)
    $tempPath = Join-Path -Path $uniqueDir -ChildPath $originalFileName

    $workbook = $null
    try {
        # Open in Excel and save a fresh copy. This produces a more canonical .xlsx that survives strict mail scanners.
        $workbook = $excel.Workbooks.Open($PathValue, 0, $true)
        $workbook.SaveCopyAs($tempPath)

        if (-not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
            throw "Excel did not create the normalized copy."
        }

        $script:NormalizedAttachmentCache[$PathValue] = $tempPath
        $script:NormalizedTempFiles.Add($tempPath) | Out-Null
        Write-Host "Normalized .xlsx for attachment: $([System.IO.Path]::GetFileName($PathValue))" -ForegroundColor DarkCyan
        return $tempPath
    } catch {
        Write-Warning "Could not normalize .xlsx attachment '$PathValue'. Attaching original file. $($_.Exception.Message)"
        $script:NormalizedAttachmentCache[$PathValue] = $PathValue
        return $PathValue
    } finally {
        if ($null -ne $workbook) {
            try { [void]$workbook.Close($false) } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } catch { }
            $workbook = $null
        }
    }
}

function Cleanup-ExcelNormalizationResources {
    foreach ($tempPath in @($script:NormalizedTempFiles.ToArray())) {
        try {
            if (-not (Test-Blank $tempPath) -and (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction Stop
            }
        } catch {
            Write-Warning "Could not delete temp normalized file '$tempPath'. $($_.Exception.Message)"
        }
    }

    if ($null -ne $script:ExcelNormalizeApp) {
        try { [void]$script:ExcelNormalizeApp.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelNormalizeApp) } catch { }
        $script:ExcelNormalizeApp = $null
    }

    if (-not (Test-Blank $script:NormalizedTempRoot) -and (Test-Path -LiteralPath $script:NormalizedTempRoot -PathType Container)) {
        try { Remove-Item -LiteralPath $script:NormalizedTempRoot -Force -Recurse -ErrorAction Stop } catch { }
    }
}

function Open-DraftsFolderIfRequested {
    param(
        [Parameter(Mandatory = $true)]$OutlookApp,
        [switch]$OpenDraftsFolder
    )
    if (-not $OpenDraftsFolder) { return }

    try {
        $namespace = $OutlookApp.GetNamespace("MAPI")
        $draftsFolder = $namespace.GetDefaultFolder(16) # olFolderDrafts
        if ($null -ne $draftsFolder) {
            [void]$draftsFolder.Display()
        }
    } catch {
        Write-Warning "Could not open Drafts folder automatically: $($_.Exception.Message)"
    }
}

function Initialize-WindowsForms {
    if ($script:FormsReady) { return }
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing | Out-Null
        $script:FormsReady = $true
    } catch {
        throw "Windows Forms dialogs are not available on this machine/session. $($_.Exception.Message)"
    }
}

function Get-ScriptDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    return (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

function Select-FilePathDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$InitialDirectory,
        [string]$Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    )
    Initialize-WindowsForms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = $Filter
    $dialog.Multiselect = $false
    if (-not (Test-Blank $InitialDirectory) -and (Test-Path -LiteralPath $InitialDirectory)) {
        $dialog.InitialDirectory = $InitialDirectory
    }
    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Selection cancelled."
    }
    return $dialog.FileName
}

function Select-FolderPathDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Description,
        [string]$InitialDirectory
    )
    Initialize-WindowsForms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $false
    if (-not (Test-Blank $InitialDirectory) -and (Test-Path -LiteralPath $InitialDirectory)) {
        $dialog.SelectedPath = $InitialDirectory
    }
    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        throw "Selection cancelled."
    }
    return $dialog.SelectedPath
}

function Show-ActionChoiceDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Caption
    )
    Initialize-WindowsForms

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Caption
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowInTaskbar = $false
    $form.TopMost = $true
    $form.ClientSize = New-Object System.Drawing.Size(480, 190)

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $false
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point(16, 16)
    $label.Size = New-Object System.Drawing.Size(448, 95)
    $label.TextAlign = 'TopLeft'

    $result = 'Cancel'
    $buttonWidth = 110
    $buttonY = 125
    $gap = 10
    $totalWidth = ($buttonWidth * 3) + ($gap * 2)
    $startX = [int](($form.ClientSize.Width - $totalWidth) / 2)

    $btnTest = New-Object System.Windows.Forms.Button
    $btnTest.Text = 'Test'
    $btnTest.Size = New-Object System.Drawing.Size($buttonWidth, 32)
    $btnTest.Location = New-Object System.Drawing.Point($startX, $buttonY)
    $btnTest.Add_Click({
            $script:ActionChoiceResult = 'Test'
            $form.Close()
        })

    $btnGenerate = New-Object System.Windows.Forms.Button
    $btnGenerate.Text = 'Generate'
    $btnGenerate.Size = New-Object System.Drawing.Size($buttonWidth, 32)
    $btnGenerate.Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $gap), $buttonY)
    $btnGenerate.Add_Click({
            $script:ActionChoiceResult = 'Generate'
            $form.Close()
        })

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size($buttonWidth, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(($startX + (($buttonWidth + $gap) * 2)), $buttonY)
    $btnCancel.Add_Click({
            $script:ActionChoiceResult = 'Cancel'
            $form.Close()
        })

    $form.Controls.Add($label)
    $form.Controls.Add($btnTest)
    $form.Controls.Add($btnGenerate)
    $form.Controls.Add($btnCancel)

    $script:ActionChoiceResult = 'Cancel'
    $form.AcceptButton = $btnGenerate
    $form.CancelButton = $btnCancel

    [void]$form.ShowDialog()
    $result = [string]$script:ActionChoiceResult
    Remove-Variable -Name ActionChoiceResult -Scope Script -ErrorAction SilentlyContinue
    return $result
}

function Show-OpenDraftsPromptDialog {
    param(
        [Parameter(Mandatory = $true)][int]$CreatedCount
    )
    Initialize-WindowsForms
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Draft creation is complete.`n`nCreated drafts: $CreatedCount`n`nOpen Outlook Drafts now?",
        "Outlook Draft Creator",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Find-AutoDetectedJobsFile {
    $scriptDir = Get-ScriptDir
    $preferred = Join-Path -Path $scriptDir -ChildPath 'outlook-draft-jobs.json'
    if (Test-Path -LiteralPath $preferred -PathType Leaf) {
        return (Resolve-Path -LiteralPath $preferred).Path
    }

    $jsonCandidates = @(Get-ChildItem -LiteralPath $scriptDir -Filter *.json -File -ErrorAction SilentlyContinue)
    $jobCandidates = @($jsonCandidates | Where-Object { $_.Name -match 'outlook.*draft.*jobs' })
    if ($jobCandidates.Count -eq 1) {
        return $jobCandidates[0].FullName
    }
    return $null
}

function Get-AttachmentRootRequirements {
    param(
        [Parameter(Mandatory = $true)]$Payload,
        [Parameter(Mandatory = $true)]$ReadyJobs
    )

    $req = [ordered]@{
        presetAttachmentFolder = $false
        commonAttachmentFolder = $false
        sourceWorkbookFolder   = $false
    }

    if ($null -ne $Payload.instructions -and $null -ne $Payload.instructions.folderRoots) {
        $roots = $Payload.instructions.folderRoots
        if ($null -ne $roots.presetAttachmentFolderRequired) { $req.presetAttachmentFolder = [bool]$roots.presetAttachmentFolderRequired }
        if ($null -ne $roots.commonAttachmentFolderRequired) { $req.commonAttachmentFolder = [bool]$roots.commonAttachmentFolderRequired }
        if ($null -ne $roots.sourceWorkbookFolderRequired) { $req.sourceWorkbookFolder = [bool]$roots.sourceWorkbookFolderRequired }
        return [pscustomobject]$req
    }

    foreach ($job in $ReadyJobs) {
        foreach ($a in @($job.attachments)) {
            switch ([string]$a.kind) {
                'matchedPreset' { $req.presetAttachmentFolder = $true }
                'common' { $req.commonAttachmentFolder = $true }
                'sourceWorkbook' { $req.sourceWorkbookFolder = $true }
            }
        }
    }

    return [pscustomobject]$req
}

function Find-AutoDetectedSourceWorkbookFolder {
    param(
        [Parameter(Mandatory = $true)][string]$JobsFilePath,
        [Parameter(Mandatory = $true)]$ReadyJobs
    )

    $jobsDir = Split-Path -Parent $JobsFilePath
    if (Test-Blank $jobsDir) { return $null }

    $candidate = Split-Path -Parent $jobsDir
    if (Test-Blank $candidate) { return $null }
    if (-not (Test-Path -LiteralPath $candidate -PathType Container)) { return $null }

    $sourceJobs = @($ReadyJobs | Where-Object {
            @($_.attachments) | Where-Object { [string]$_.kind -eq 'sourceWorkbook' }
        })
    if ($sourceJobs.Count -eq 0) {
        return $null
    }

    foreach ($job in $sourceJobs) {
        $rel = [string]$job.sourceExcelRelativePath
        if (Test-Blank $rel) {
            continue
        }
        try {
            [void](Join-ExistingPath -Root $candidate -RelativePath $rel)
            return (Resolve-Path -LiteralPath $candidate).Path
        } catch {
            # Keep checking other jobs in case the first one is unusual.
        }
    }

    return $null
}

Write-Host "Outlook Draft Creator (Classic Outlook COM)" -ForegroundColor Cyan
Write-Host "Date: $(Get-Date -Format s)"

$scriptDir = Get-ScriptDir
if (Test-Blank $JobsPath) {
    $JobsPath = Find-AutoDetectedJobsFile
}

if ((Test-Blank $JobsPath) -and $Interactive) {
    Write-Host "Select the exported outlook-draft-jobs.json file..." -ForegroundColor Yellow
    $JobsPath = Select-FilePathDialog -Title "Select outlook-draft-jobs.json" -InitialDirectory $scriptDir
}

if (Test-Blank $JobsPath) {
    throw "No JobsPath provided. Pass -JobsPath or use -Interactive."
}

$jobsFile = Resolve-RequiredPath -PathValue $JobsPath -Label "JobsPath"
if (-not (Test-Path -LiteralPath $jobsFile -PathType Leaf)) {
    throw "Jobs file not found: $jobsFile"
}

$rawJson = Get-Content -LiteralPath $jobsFile -Raw -Encoding UTF8
try {
    $payload = $rawJson | ConvertFrom-Json
} catch {
    throw "Jobs file is not valid JSON: $($_.Exception.Message)"
}

$jobs = Get-JobArray -Payload $payload
if ($jobs.Count -eq 0) {
    throw "Jobs file contains no jobs."
}

$readyJobs = @($jobs | Where-Object { [string]$_.status -eq 'ready' })
if ($readyJobs.Count -eq 0) {
    throw "Jobs file has no jobs with status 'ready'."
}

$promptOpenDraftsAfterCreate = $false

$rootReq = Get-AttachmentRootRequirements -Payload $payload -ReadyJobs $readyJobs

if ($Interactive) {
    if ($rootReq.presetAttachmentFolder -and (Test-Blank $PresetAttachmentFolder)) {
        Write-Host "Select the folder containing matched preset attachments..." -ForegroundColor Yellow
        $PresetAttachmentFolder = Select-FolderPathDialog -Description "Select folder containing matched preset attachments" -InitialDirectory $scriptDir
    }
    if ($rootReq.commonAttachmentFolder -and (Test-Blank $CommonAttachmentFolder)) {
        Write-Host "Select the folder containing common attachments..." -ForegroundColor Yellow
        $CommonAttachmentFolder = Select-FolderPathDialog -Description "Select folder containing common attachments" -InitialDirectory $scriptDir
    }
    if ($rootReq.sourceWorkbookFolder -and (Test-Blank $SourceWorkbookFolder)) {
        $autoDetectedSourceFolder = Find-AutoDetectedSourceWorkbookFolder -JobsFilePath $jobsFile -ReadyJobs $readyJobs
        if (-not (Test-Blank $autoDetectedSourceFolder)) {
            $SourceWorkbookFolder = $autoDetectedSourceFolder
            Write-Host "Auto-detected source Excel folder: $SourceWorkbookFolder" -ForegroundColor Green
        } else {
            Write-Host "Select the folder containing the Excel source files (Part 1 files)..." -ForegroundColor Yellow
            $SourceWorkbookFolder = Select-FolderPathDialog -Description "Select folder containing source Excel files" -InitialDirectory $scriptDir
        }
    }

    if (-not $PSBoundParameters.ContainsKey('DryRun')) {
        $choice = Show-ActionChoiceDialog -Caption "Outlook Draft Creator" -Text "Choose what to do next:`n`nTest = Dry Run (validate only, no drafts created)`nGenerate = Create Outlook drafts now`nCancel = Exit"
        if ($choice -eq 'Cancel') {
            throw "Cancelled by user."
        }
        if ($choice -eq 'Test') {
            $DryRun = $true
        } else {
            $DryRun = $false
            if (-not $PSBoundParameters.ContainsKey('OpenDraftsFolder')) {
                $promptOpenDraftsAfterCreate = $true
            }
        }
    }
}

$resolvedPresetRoot = if (-not (Test-Blank $PresetAttachmentFolder)) { Resolve-RequiredPath -PathValue $PresetAttachmentFolder -Label "PresetAttachmentFolder" } else { $null }
$resolvedCommonRoot = if (-not (Test-Blank $CommonAttachmentFolder)) { Resolve-RequiredPath -PathValue $CommonAttachmentFolder -Label "CommonAttachmentFolder" } else { $null }
$resolvedSourceRoot = if (-not (Test-Blank $SourceWorkbookFolder)) { Resolve-RequiredPath -PathValue $SourceWorkbookFolder -Label "SourceWorkbookFolder" } else { $null }

# Rebind to resolved absolute paths for helper functions
if ($resolvedPresetRoot) { $PresetAttachmentFolder = $resolvedPresetRoot }
if ($resolvedCommonRoot) { $CommonAttachmentFolder = $resolvedCommonRoot }
if ($resolvedSourceRoot) { $SourceWorkbookFolder = $resolvedSourceRoot }

Write-Host "Loaded jobs: $($jobs.Count) total / $($readyJobs.Count) ready" -ForegroundColor Green

$outlook = $null
if (-not $DryRun) {
    $outlook = New-OutlookApplication
}

$results = New-Object System.Collections.Generic.List[object]
$created = 0
$failed = 0
$index = 0

foreach ($job in $readyJobs) {
    $index++
    $jobId = [string]$job.id
    $sourceName = [string]$job.sourceExcelFileName
    $toList = @(Get-StringArray -Value $job.to)
    $ccList = @(Get-StringArray -Value $job.cc)
    $bccList = @(Get-StringArray -Value $job.bcc)
    $replyToList = @(Get-StringArray -Value $job.replyTo)
    $subject = [string]$job.subject
    $bodyText = [string]$job.bodyText
    $attachments = @($job.attachments)

    Write-Host ("[{0}/{1}] {2}" -f $index, $readyJobs.Count, $sourceName)

    try {
        if ($toList.Count -eq 0) {
            throw "No recipients in ready job."
        }

        $attachmentPaths = New-Object System.Collections.Generic.List[string]
        $attachmentItemsForSend = New-Object System.Collections.Generic.List[object]
        foreach ($attachment in $attachments) {
            $resolvedPath = Resolve-AttachmentPath -Attachment $attachment -Job $job
            $attachmentPaths.Add($resolvedPath) | Out-Null
            $displayName = [string]$attachment.fileName
            if (Test-Blank $displayName) {
                $displayName = [System.IO.Path]::GetFileName($resolvedPath)
            }

            $pathForSend = $resolvedPath
            if ($DryRun) {
                $pathForSend = $resolvedPath
            } else {
                $pathForSend = Get-NormalizedAttachmentPathForSend -PathValue $resolvedPath
            }

            $attachmentItemsForSend.Add([pscustomobject]@{
                    OriginalPath = $resolvedPath
                    PathForSend  = $pathForSend
                    DisplayName  = $displayName
                }) | Out-Null
        }

        if ($DryRun) {
            $results.Add([pscustomobject]@{
                    id                 = $jobId
                    sourceExcelFileName = $sourceName
                    status             = 'dry_run_ok'
                    to                 = ($toList -join '; ')
                    subject            = $subject
                    attachmentCount    = $attachmentPaths.Count
                    attachmentPaths    = @($attachmentPaths.ToArray())
                    error              = $null
                }) | Out-Null
            $created++
            continue
        }

        $mail = $outlook.CreateItem(0) # olMailItem
        $mail.Subject = $subject
        $mail.Body = $bodyText
        $mail.To = ($toList -join '; ')
        if ($ccList.Count -gt 0) { $mail.CC = ($ccList -join '; ') }
        if ($bccList.Count -gt 0) { $mail.BCC = ($bccList -join '; ') }
        if ($replyToList.Count -gt 0) { Add-ReplyRecipients -MailItem $mail -ReplyTo $replyToList }

        foreach ($attachmentItem in @($attachmentItemsForSend.ToArray())) {
            $sendPath = [string]$attachmentItem.PathForSend
            $displayName = [string]$attachmentItem.DisplayName
            if (Test-Blank $displayName) {
                [void]$mail.Attachments.Add($sendPath)
            } else {
                [void]$mail.Attachments.Add($sendPath, [Type]::Missing, [Type]::Missing, $displayName)
            }
        }

        # Save as draft in the user's default Drafts folder for visual review/send.
        [void]$mail.Save()

        $results.Add([pscustomobject]@{
                id                 = $jobId
                sourceExcelFileName = $sourceName
                status             = 'created'
                to                 = ($toList -join '; ')
                subject            = $subject
                attachmentCount    = $attachmentPaths.Count
                attachmentPaths    = @($attachmentPaths.ToArray())
                error              = $null
            }) | Out-Null
        $created++
    } catch {
        $failed++
        $msg = $_.Exception.Message
        Write-Warning "Failed for $sourceName : $msg"
        $results.Add([pscustomobject]@{
                id                 = $jobId
                sourceExcelFileName = $sourceName
                status             = 'error'
                to                 = ($toList -join '; ')
                subject            = $subject
                attachmentCount    = 0
                attachmentPaths    = @()
                error              = $msg
            }) | Out-Null
    }
}

$resultPayload = [pscustomobject]@{
    schemaVersion = '1.0-outlook-draft-results'
    generatedAt   = (Get-Date).ToString('s')
    dryRun        = [bool]$DryRun
    jobsPath      = $jobsFile
    summary       = [pscustomobject]@{
        totalReadyJobs = $readyJobs.Count
        ok            = $created
        failed        = $failed
    }
    results       = @($results.ToArray())
}

$resultPath = Join-Path -Path (Split-Path -Parent $jobsFile) -ChildPath 'outlook-draft-results.json'
$jsonContent = $resultPayload | ConvertTo-Json -Depth 100
$maxRetries = 5
$retryDelay = 2
for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
    try {
        [System.IO.File]::WriteAllText($resultPath, $jsonContent, [System.Text.Encoding]::UTF8)
        break
    } catch [System.IO.IOException] {
        if ($attempt -eq $maxRetries) { throw }
        Write-Warning "Results file is locked (attempt $attempt/$maxRetries), retrying in ${retryDelay}s..."
        Start-Sleep -Seconds $retryDelay
    }
}

Write-Host ""
Write-Host "Done. Created/validated: $created | Failed: $failed" -ForegroundColor Cyan
Write-Host "Results written to: $resultPath"

Cleanup-ExcelNormalizationResources

if (-not $DryRun -and $promptOpenDraftsAfterCreate -and $created -gt 0) {
    $OpenDraftsFolder = Show-OpenDraftsPromptDialog -CreatedCount $created
}

if (-not $DryRun -and $OpenDraftsFolder) {
    Open-DraftsFolderIfRequested -OutlookApp $outlook -OpenDraftsFolder
}
