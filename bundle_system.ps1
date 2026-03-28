[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Bundle', 'Restore', 'RestoreStructure')]
    [string]$Mode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Script:ExitCodes = @{
    Success             = 0
    NoInputFiles        = 10
    InvalidExtension    = 11
    ReadFailure         = 12
    WriteFailure        = 13
    BundleNotFound      = 20
    BundleMultiple      = 21
    InvalidFormat       = 22
    InvalidPath         = 23
    RestoreConflict     = 24
    PermissionDenied    = 30
    Unexpected          = 99
}

$Script:FormatName = 'BatchPsBundle'
$Script:FormatVersion = '1.1'
$Script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Script:ReservedNames = @(
    'CON', 'PRN', 'AUX', 'NUL',
    'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
    'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
)
$Script:AllowedExtensions = @('.bat', '.ps1', '.md', '.json', '.jsonl', '.bas')
$Script:AllowedNewlineStyles = @('None', 'CRLF', 'LF', 'CR', 'Mixed')
$Script:AllowedBomTypes = @('None', 'UTF32-LE', 'UTF32-BE', 'UTF8-BOM', 'UTF16-LE', 'UTF16-BE')

function Throw-HandledError {
    param(
        [int]$Code,
        [string]$Message
    )

    $exception = New-Object System.Exception($Message)
    $exception.Data['ExitCode'] = $Code
    throw $exception
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [System.IO.Directory]::CreateDirectory($Path) | Out-Null
    }
}

function Initialize-WorkingFolders {
    param([string]$RootPath)

    $paths = [ordered]@{
        Root         = $RootPath
        InputFiles   = Join-Path -Path $RootPath -ChildPath 'input_files'
        OutputBundle = Join-Path -Path $RootPath -ChildPath 'output_bundle'
        RestoreInput = Join-Path -Path $RootPath -ChildPath 'restore_input'
        RestoreOutput = Join-Path -Path $RootPath -ChildPath 'restore_output'
    }

    foreach ($path in $paths.Values) {
        Ensure-Directory -Path $path
    }

    return $paths
}

function Get-RelativePath {
    param(
        [string]$BasePath,
        [string]$TargetPath
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    $directorySeparator = [string][System.IO.Path]::DirectorySeparatorChar
    if (-not $baseFullPath.EndsWith($directorySeparator)) {
        $baseFullPath += $directorySeparator
    }

    $baseUri = New-Object System.Uri($baseFullPath)
    $targetUri = New-Object System.Uri([System.IO.Path]::GetFullPath($TargetPath))
    $relativeUri = $baseUri.MakeRelativeUri($targetUri)
    return [System.Uri]::UnescapeDataString($relativeUri.ToString()).Replace('/', '\')
}

function Get-Sha256Hex {
    param([byte[]]$Bytes)

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha256.ComputeHash($Bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-BomType {
    param([byte[]]$Bytes)

    if ($Bytes.Length -ge 4) {
        if ($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE -and $Bytes[2] -eq 0x00 -and $Bytes[3] -eq 0x00) { return 'UTF32-LE' }
        if ($Bytes[0] -eq 0x00 -and $Bytes[1] -eq 0x00 -and $Bytes[2] -eq 0xFE -and $Bytes[3] -eq 0xFF) { return 'UTF32-BE' }
    }
    if ($Bytes.Length -ge 3) {
        if ($Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { return 'UTF8-BOM' }
    }
    if ($Bytes.Length -ge 2) {
        if ($Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) { return 'UTF16-LE' }
        if ($Bytes[0] -eq 0xFE -and $Bytes[1] -eq 0xFF) { return 'UTF16-BE' }
    }
    return 'None'
}

function Get-NewlineStyle {
    param([byte[]]$Bytes)

    $hasCrLf = $false
    $hasLf = $false
    $hasCr = $false
    $index = 0

    while ($index -lt $Bytes.Length) {
        if ($Bytes[$index] -eq 13) {
            if (($index + 1) -lt $Bytes.Length -and $Bytes[$index + 1] -eq 10) {
                $hasCrLf = $true
                $index += 2
                continue
            }

            $hasCr = $true
        }
        elseif ($Bytes[$index] -eq 10) {
            $hasLf = $true
        }

        $index += 1
    }

    $styles = @()
    if ($hasCrLf) { $styles += 'CRLF' }
    if ($hasLf) { $styles += 'LF' }
    if ($hasCr) { $styles += 'CR' }

    if ($styles.Count -eq 0) { return 'None' }
    if ($styles.Count -eq 1) { return $styles[0] }
    return 'Mixed'
}

function Get-UniqueBundlePath {
    param([string]$OutputDirectory)

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseName = "bundle_$timestamp"
    $candidate = Join-Path -Path $OutputDirectory -ChildPath "$baseName.txt"
    $counter = 1

    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path -Path $OutputDirectory -ChildPath ('{0}_{1:00}.txt' -f $baseName, $counter)
        $counter += 1
    }

    return $candidate
}

function Get-FileNameExtension {
    param([string]$RelativePath)

    return [System.IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
}

function Test-ReservedName {
    param([string]$Name)

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Name).ToUpperInvariant()
    return $Script:ReservedNames -contains $stem
}

function Resolve-SafeRestorePath {
    param(
        [string]$RestoreRoot,
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message '復元対象の相対パスが空です。'
    }

    if ([System.IO.Path]::IsPathRooted($RelativePath)) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "絶対パスは復元できません: $RelativePath"
    }

    $normalizedRelativePath = $RelativePath.Replace('/', '\')
    $segments = $normalizedRelativePath.Split('\')

    if ($segments.Count -eq 0) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "相対パスを解釈できません: $RelativePath"
    }

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($segment in $segments) {
        if ([string]::IsNullOrWhiteSpace($segment)) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "空のパス要素は許可されません: $RelativePath"
        }
        if ($segment -eq '.' -or $segment -eq '..') {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "危険な相対パスを検出しました: $RelativePath"
        }
        if ($segment.EndsWith('.') -or $segment.EndsWith(' ')) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "末尾のドットまたは空白を含む名前は復元できません: $RelativePath"
        }
        if ($segment.IndexOfAny($invalidChars) -ge 0) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "禁止文字を含む相対パスです: $RelativePath"
        }
        if (Test-ReservedName -Name $segment) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "予約名を含む相対パスです: $RelativePath"
        }
    }

    $restoreRootFullPath = [System.IO.Path]::GetFullPath($RestoreRoot)
    $fullPath = [System.IO.Path]::GetFullPath((Join-Path -Path $restoreRootFullPath -ChildPath $normalizedRelativePath))
    $prefix = $restoreRootFullPath.TrimEnd('\') + '\'

    if (-not $fullPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "復元先の外側へ出る相対パスを検出しました: $RelativePath"
    }

    return $fullPath
}

function Get-BundleCandidates {
    param([string]$RestoreInputPath)

    return @(Get-ChildItem -LiteralPath $RestoreInputPath -File -Filter 'bundle*.txt' | Sort-Object Name)
}

function Show-PathList {
    param(
        [string]$Label,
        [string[]]$Paths
    )

    if ($null -eq $Paths -or $Paths.Count -eq 0) {
        return
    }

    Write-Host $Label
    foreach ($path in $Paths) {
        Write-Host " - $path"
    }
}

function Get-DirectoryRelativePaths {
    param(
        [string]$RelativePath,
        [bool]$IncludeLeaf
    )

    $normalizedRelativePath = $RelativePath.Replace('/', '\').Trim('\')
    if ([string]::IsNullOrWhiteSpace($normalizedRelativePath)) {
        return @()
    }

    $segments = $normalizedRelativePath.Split('\')
    $lastIndex = if ($IncludeLeaf) { $segments.Count - 1 } else { $segments.Count - 2 }
    if ($lastIndex -lt 0) {
        return @()
    }

    $directoryRelativePaths = New-Object System.Collections.ArrayList
    for ($index = 0; $index -le $lastIndex; $index++) {
        [void]$directoryRelativePaths.Add(($segments[0..$index] -join '\'))
    }

    return @($directoryRelativePaths)
}

function Test-IgnoredInputFile {
    param([System.IO.FileInfo]$File)

    $name = $File.Name.ToLowerInvariant()
    if ($name -in @('thumbs.db', 'desktop.ini')) {
        return $true
    }
    if ($name -like 'bundle*.txt') {
        return $true
    }
    if ($name -like '~$*') {
        return $true
    }
    if ($name -like '*.tmp' -or $name -like '*.temp') {
        return $true
    }

    return $false
}

function Assert-RestoreParentPathSafe {
    param(
        [string]$RestoreRoot,
        [string]$TargetPath,
        [string]$RelativePath
    )

    $restoreRootFullPath = [System.IO.Path]::GetFullPath($RestoreRoot)
    $prefix = $restoreRootFullPath.TrimEnd('\') + '\'
    $currentPath = [System.IO.Path]::GetDirectoryName($TargetPath)

    while (-not [string]::IsNullOrWhiteSpace($currentPath) -and
           $currentPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        if ($currentPath -eq $restoreRootFullPath) {
            break
        }

        if (Test-Path -LiteralPath $currentPath) {
            $item = Get-Item -LiteralPath $currentPath -Force
            if (-not $item.PSIsContainer) {
                Throw-HandledError -Code $Script:ExitCodes.RestoreConflict -Message "復元先の親パスがファイルと衝突しています: $RelativePath -> $currentPath"
            }
        }

        $currentPath = [System.IO.Path]::GetDirectoryName($currentPath)
    }
}

function Read-Bytes {
    param([string]$Path)

    try {
        return [System.IO.File]::ReadAllBytes($Path)
    }
    catch [System.UnauthorizedAccessException] {
        Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "読み取り権限がありません: $Path"
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.ReadFailure -Message "ファイルを読み取れませんでした: $Path"
    }
}

function Write-TextFile {
    param(
        [string]$Path,
        [string]$Content
    )

    try {
        [System.IO.File]::WriteAllText($Path, $Content, $Script:Utf8NoBom)
    }
    catch [System.UnauthorizedAccessException] {
        Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "書き込み権限がありません: $Path"
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.WriteFailure -Message "テキストファイルを書き込めませんでした: $Path"
    }
}

function Write-BytesFile {
    param(
        [string]$Path,
        [byte[]]$Bytes
    )

    try {
        [System.IO.File]::WriteAllBytes($Path, $Bytes)
    }
    catch [System.UnauthorizedAccessException] {
        Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "書き込み権限がありません: $Path"
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.WriteFailure -Message "ファイルを書き込めませんでした: $Path"
    }
}

function Read-JsonFile {
    param([string]$Path)

    try {
        return [System.IO.File]::ReadAllText($Path, $Script:Utf8NoBom)
    }
    catch [System.UnauthorizedAccessException] {
        Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "読み取り権限がありません: $Path"
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.ReadFailure -Message "集約ファイルを読み取れませんでした: $Path"
    }
}

function Show-Start {
    param(
        [string]$Operation,
        [System.Collections.IDictionary]$Paths
    )

    Write-Host ('=' * 60)
    Write-Host "処理種別 : $Operation"
    if ($Operation -eq '集約') {
        Write-Host "入力元     : $($Paths.InputFiles)"
        Write-Host "出力先     : $($Paths.OutputBundle)"
    }
    else {
        Write-Host "入力元     : $($Paths.RestoreInput)"
        Write-Host "復元先     : $($Paths.RestoreOutput)"
    }
    Write-Host ('=' * 60)
}

function Show-Summary {
    param(
        [string]$Operation,
        [int]$TargetCount,
        [int]$SuccessCount,
        [int]$FailureCount,
        [string]$OutputPath
    )

    Write-Host ''
    Write-Host "処理結果   : 正常終了"
    Write-Host "処理種別   : $Operation"
    Write-Host "対象件数   : $TargetCount"
    Write-Host "成功件数   : $SuccessCount"
    Write-Host "失敗件数   : $FailureCount"
    Write-Host "出力先     : $OutputPath"
    Write-Host '完了メッセージ: 処理が完了しました。'
}

function Invoke-Bundle {
    param([System.Collections.IDictionary]$Paths)

    Show-Start -Operation '集約' -Paths $Paths

    $allFiles = @(Get-ChildItem -LiteralPath $Paths.InputFiles -File -Recurse | Sort-Object FullName)
    $allDirectories = @(Get-ChildItem -LiteralPath $Paths.InputFiles -Directory -Recurse | Sort-Object FullName)
    $ignoredFiles = @($allFiles | Where-Object { Test-IgnoredInputFile -File $_ })
    $candidateFiles = @($allFiles | Where-Object { -not (Test-IgnoredInputFile -File $_) })
    $directoryEntries = New-Object System.Collections.ArrayList

    $sortedDirectories = @($allDirectories | Sort-Object { Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $_.FullName })
    $directoryId = 1
    foreach ($directory in $sortedDirectories) {
        [void]$directoryEntries.Add([ordered]@{
            id           = $directoryId
            relativePath = (Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $directory.FullName)
        })
        $directoryId += 1
    }

    if ($candidateFiles.Count -eq 0 -and $directoryEntries.Count -eq 0) {
        Throw-HandledError -Code $Script:ExitCodes.NoInputFiles -Message "対象ファイルなし: $($Paths.InputFiles) に対象ファイルまたはフォルダを配置してください。"
    }

    $invalidFiles = @($candidateFiles | Where-Object { $Script:AllowedExtensions -notcontains $_.Extension.ToLowerInvariant() })
    $supportedFiles = @($candidateFiles | Where-Object { $Script:AllowedExtensions -contains $_.Extension.ToLowerInvariant() })
    if ($invalidFiles.Count -gt 0) {
        Write-Host "補足       : 変換対象外の $($invalidFiles.Count) 件をスキップします。"
        Show-PathList -Label '変換対象外ファイル一覧:' -Paths @($invalidFiles | ForEach-Object { Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $_.FullName })
    }
    if ($supportedFiles.Count -eq 0 -and $invalidFiles.Count -gt 0) {
        $invalidList = ($invalidFiles | Select-Object -ExpandProperty FullName) -join ', '
        Throw-HandledError -Code $Script:ExitCodes.InvalidExtension -Message "変換可能な対象ファイルがありません。対象外ファイルのみです: $invalidList"
    }

    $bundleEntries = New-Object System.Collections.ArrayList
    $sortedFiles = @($supportedFiles | Sort-Object { Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $_.FullName })
    $id = 1

    foreach ($file in $sortedFiles) {
        $relativePath = Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $file.FullName
        $bytes = Read-Bytes -Path $file.FullName
        [void]$bundleEntries.Add([ordered]@{
            id           = $id
            relativePath = $relativePath
            fileName     = $file.Name
            extension    = $file.Extension.ToLowerInvariant()
            byteLength   = $bytes.Length
            sha256       = Get-Sha256Hex -Bytes $bytes
            newlineStyle = Get-NewlineStyle -Bytes $bytes
            bomType      = Get-BomType -Bytes $bytes
            contentBase64 = [System.Convert]::ToBase64String($bytes)
        })
        $id += 1
    }

    $bundleObject = [ordered]@{
        format    = $Script:FormatName
        version   = $Script:FormatVersion
        createdAt = (Get-Date).ToString('o')
        dirCount  = $directoryEntries.Count
        directories = @($directoryEntries)
        fileCount = $bundleEntries.Count
        files     = @($bundleEntries)
    }

    $outputPath = Get-UniqueBundlePath -OutputDirectory $Paths.OutputBundle
    $json = $bundleObject | ConvertTo-Json -Depth 5
    Write-TextFile -Path $outputPath -Content $json

    Show-Summary -Operation '集約' -TargetCount $bundleEntries.Count -SuccessCount $bundleEntries.Count -FailureCount 0 -OutputPath $outputPath
    Write-Host "対象フォルダ件数 : $($directoryEntries.Count)"
    if ($ignoredFiles.Count -gt 0) {
        Write-Host "補足       : 除外ルールに一致した $($ignoredFiles.Count) 件のファイルを無視しました。"
        Show-PathList -Label '除外ファイル一覧:' -Paths @($ignoredFiles | ForEach-Object { Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $_.FullName })
    }
}

function Get-RequiredString {
    param(
        [psobject]$Object,
        [string]$PropertyName,
        [string]$Context
    )

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "$Context に必須項目 $PropertyName がありません。"
    }

    return [string]$property.Value
}

function Get-RequiredInteger {
    param(
        [psobject]$Object,
        [string]$PropertyName,
        [string]$Context
    )

    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "$Context に必須項目 $PropertyName がありません。"
    }

    try {
        return [int]$property.Value
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "$Context の $PropertyName は整数である必要があります。"
    }
}

function Get-RequiredEnumString {
    param(
        [psobject]$Object,
        [string]$PropertyName,
        [string]$Context,
        [string[]]$AllowedValues
    )

    $value = Get-RequiredString -Object $Object -PropertyName $PropertyName -Context $Context
    if ($AllowedValues -notcontains $value) {
        $allowed = $AllowedValues -join ', '
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "$Context の $PropertyName が不正です。許可値: $allowed"
    }

    return $value
}

function Add-DirectoryRestoreTarget {
    param(
        [string]$RestoreRoot,
        [string]$RelativeDirectoryPath,
        [System.Collections.Generic.HashSet[string]]$KnownTargets,
        [System.Collections.ArrayList]$DirectoryPlan
    )

    if ([string]::IsNullOrWhiteSpace($RelativeDirectoryPath)) {
        return
    }

    $targetDirectoryPath = Resolve-SafeRestorePath -RestoreRoot $RestoreRoot -RelativePath $RelativeDirectoryPath
    if (-not $KnownTargets.Add($targetDirectoryPath)) {
        return
    }

    Assert-RestoreParentPathSafe -RestoreRoot $RestoreRoot -TargetPath $targetDirectoryPath -RelativePath $RelativeDirectoryPath
    if (Test-Path -LiteralPath $targetDirectoryPath) {
        $item = Get-Item -LiteralPath $targetDirectoryPath -Force
        if (-not $item.PSIsContainer) {
            Throw-HandledError -Code $Script:ExitCodes.RestoreConflict -Message "復元先のフォルダが既存ファイルと衝突しています: $RelativeDirectoryPath -> $targetDirectoryPath"
        }
    }

    [void]$DirectoryPlan.Add([pscustomobject]@{
        RelativePath = $RelativeDirectoryPath
        TargetPath   = $targetDirectoryPath
    })
}

function Invoke-Restore {
    param([System.Collections.IDictionary]$Paths)

    $operationLabel = if ($Mode -eq 'RestoreStructure') { 'フォルダ構成復元' } else { '復元' }
    Show-Start -Operation $operationLabel -Paths $Paths

    $bundleCandidates = @(Get-BundleCandidates -RestoreInputPath $Paths.RestoreInput)
    if ($bundleCandidates.Count -eq 0) {
        Throw-HandledError -Code $Script:ExitCodes.BundleNotFound -Message "復元用の集約ファイルが見つかりません: $($Paths.RestoreInput)"
    }
    if ($bundleCandidates.Count -gt 1) {
        $bundleList = ($bundleCandidates | Select-Object -ExpandProperty Name) -join ', '
        Throw-HandledError -Code $Script:ExitCodes.BundleMultiple -Message "復元用の集約ファイルは1件だけ配置してください: $bundleList"
    }

    $bundlePath = $bundleCandidates[0].FullName
    $jsonText = Read-JsonFile -Path $bundlePath

    try {
        $bundleObject = $jsonText | ConvertFrom-Json
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "集約ファイルのJSON形式が不正です: $bundlePath"
    }

    $format = Get-RequiredString -Object $bundleObject -PropertyName 'format' -Context '集約ファイル'
    if ($format -ne $Script:FormatName) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "想定外のフォーマット識別子です: $format"
    }

    [void](Get-RequiredString -Object $bundleObject -PropertyName 'version' -Context '集約ファイル')
    [void](Get-RequiredString -Object $bundleObject -PropertyName 'createdAt' -Context '集約ファイル')
    $dirCount = 0
    $dirRecords = @()
    if ($bundleObject.PSObject.Properties['directories']) {
        $dirRecords = @($bundleObject.directories)
    }
    if ($bundleObject.PSObject.Properties['dirCount']) {
        $dirCount = Get-RequiredInteger -Object $bundleObject -PropertyName 'dirCount' -Context '集約ファイル'
        if ($dirRecords.Count -ne $dirCount) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "dirCount とディレクトリ件数が一致しません。宣言件数: $dirCount / 実件数: $($dirRecords.Count)"
        }
    }
    else {
        $dirCount = $dirRecords.Count
    }
    $fileCount = Get-RequiredInteger -Object $bundleObject -PropertyName 'fileCount' -Context '集約ファイル'

    $fileRecords = @($bundleObject.files)
    if ($fileRecords.Count -ne $fileCount) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "fileCount と実データ件数が一致しません。宣言件数: $fileCount / 実件数: $($fileRecords.Count)"
    }

    $directoryPlan = New-Object System.Collections.ArrayList
    $plannedDirectoryTargets = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $directoryRecordIds = New-Object System.Collections.Generic.HashSet[int]
    $restorePlan = New-Object System.Collections.Generic.List[object]
    $plannedTargets = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    $recordIds = New-Object System.Collections.Generic.HashSet[int]

    foreach ($directoryRecord in $dirRecords) {
        $directoryRecordId = Get-RequiredInteger -Object $directoryRecord -PropertyName 'id' -Context 'ディレクトリレコード'
        $directoryRelativePath = Get-RequiredString -Object $directoryRecord -PropertyName 'relativePath' -Context "ディレクトリレコード $directoryRecordId"

        if ($directoryRecordId -le 0) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "ディレクトリ id は 1 以上である必要があります: $directoryRelativePath"
        }
        if (-not $directoryRecordIds.Add($directoryRecordId)) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "ディレクトリ id が重複しています: $directoryRecordId"
        }

        foreach ($directoryChainPath in (Get-DirectoryRelativePaths -RelativePath $directoryRelativePath -IncludeLeaf $true)) {
            Add-DirectoryRestoreTarget -RestoreRoot $Paths.RestoreOutput -RelativeDirectoryPath $directoryChainPath -KnownTargets $plannedDirectoryTargets -DirectoryPlan $directoryPlan
        }
    }

    foreach ($record in $fileRecords) {
        $recordId = Get-RequiredInteger -Object $record -PropertyName 'id' -Context 'ファイルレコード'
        $relativePath = Get-RequiredString -Object $record -PropertyName 'relativePath' -Context 'ファイルレコード'
        $fileName = Get-RequiredString -Object $record -PropertyName 'fileName' -Context "ファイルレコード $relativePath"
        $extension = Get-RequiredString -Object $record -PropertyName 'extension' -Context "ファイルレコード $relativePath"
        $contentBase64 = Get-RequiredString -Object $record -PropertyName 'contentBase64' -Context "ファイルレコード $relativePath"
        $byteLength = Get-RequiredInteger -Object $record -PropertyName 'byteLength' -Context "ファイルレコード $relativePath"
        $sha256 = Get-RequiredString -Object $record -PropertyName 'sha256' -Context "ファイルレコード $relativePath"
        $newlineStyle = Get-RequiredEnumString -Object $record -PropertyName 'newlineStyle' -Context "ファイルレコード $relativePath" -AllowedValues $Script:AllowedNewlineStyles
        $bomType = Get-RequiredEnumString -Object $record -PropertyName 'bomType' -Context "ファイルレコード $relativePath" -AllowedValues $Script:AllowedBomTypes

        if ($recordId -le 0) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "id は 1 以上である必要があります: $relativePath"
        }
        if (-not $recordIds.Add($recordId)) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "id が重複しています: $recordId"
        }

        if ($fileName -ne [System.IO.Path]::GetFileName($relativePath)) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "fileName と relativePath が一致しません: $relativePath"
        }
        if ($extension -ne (Get-FileNameExtension -RelativePath $relativePath)) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "extension と relativePath が一致しません: $relativePath"
        }
        if ($Script:AllowedExtensions -notcontains $extension.ToLowerInvariant()) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "許可されていない拡張子を含むレコードです: $relativePath"
        }

        try {
            $bytes = [System.Convert]::FromBase64String($contentBase64)
        }
        catch {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "Base64 データが不正です: $relativePath"
        }

        if ($bytes.Length -ne $byteLength) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "byteLength が一致しません: $relativePath"
        }

        if ((Get-Sha256Hex -Bytes $bytes) -ne $sha256.ToLowerInvariant()) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "SHA-256 が一致しません: $relativePath"
        }
        if ((Get-NewlineStyle -Bytes $bytes) -ne $newlineStyle) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "newlineStyle が実データと一致しません: $relativePath"
        }
        if ((Get-BomType -Bytes $bytes) -ne $bomType) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "bomType が実データと一致しません: $relativePath"
        }

        foreach ($directoryChainPath in (Get-DirectoryRelativePaths -RelativePath $relativePath -IncludeLeaf $false)) {
            Add-DirectoryRestoreTarget -RestoreRoot $Paths.RestoreOutput -RelativeDirectoryPath $directoryChainPath -KnownTargets $plannedDirectoryTargets -DirectoryPlan $directoryPlan
        }

        $targetPath = Resolve-SafeRestorePath -RestoreRoot $Paths.RestoreOutput -RelativePath $relativePath
        if (-not $plannedTargets.Add($targetPath)) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidFormat -Message "同一の復元先パスが重複しています: $relativePath"
        }

        if (Test-Path -LiteralPath $targetPath) {
            $existingTarget = Get-Item -LiteralPath $targetPath -Force
            if ($existingTarget.PSIsContainer) {
                Throw-HandledError -Code $Script:ExitCodes.RestoreConflict -Message "復元先のファイルパスが既存フォルダと衝突しています: $targetPath"
            }
            Throw-HandledError -Code $Script:ExitCodes.RestoreConflict -Message "復元先に同名ファイルが存在します: $targetPath"
        }
        Assert-RestoreParentPathSafe -RestoreRoot $Paths.RestoreOutput -TargetPath $targetPath -RelativePath $relativePath

        [void]$restorePlan.Add([pscustomobject]@{
            RelativePath = $relativePath
            TargetPath   = $targetPath
            Bytes        = $bytes
        })
    }

    foreach ($directoryItem in @($directoryPlan | Sort-Object { $_.TargetPath.Length }, TargetPath)) {
        Ensure-Directory -Path $directoryItem.TargetPath
    }

    if ($Mode -eq 'RestoreStructure') {
        Show-Summary -Operation $operationLabel -TargetCount $directoryPlan.Count -SuccessCount $directoryPlan.Count -FailureCount 0 -OutputPath $Paths.RestoreOutput
        Write-Host "対象フォルダ件数 : $($directoryPlan.Count)"
        return
    }

    foreach ($item in $restorePlan) {
        Write-BytesFile -Path $item.TargetPath -Bytes $item.Bytes
    }

    Show-Summary -Operation $operationLabel -TargetCount $restorePlan.Count -SuccessCount $restorePlan.Count -FailureCount 0 -OutputPath $Paths.RestoreOutput
    Write-Host "対象フォルダ件数 : $($directoryPlan.Count)"
}

try {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    $paths = Initialize-WorkingFolders -RootPath $scriptRoot

    switch ($Mode) {
        'Bundle' { Invoke-Bundle -Paths $paths }
        'Restore' { Invoke-Restore -Paths $paths }
        'RestoreStructure' { Invoke-Restore -Paths $paths }
        default { Throw-HandledError -Code $Script:ExitCodes.Unexpected -Message "未対応のモードです: $Mode" }
    }

    exit $Script:ExitCodes.Success
}
catch {
    $exitCode = $Script:ExitCodes.Unexpected
    if ($_.Exception -and $_.Exception.Data.Contains('ExitCode')) {
        $exitCode = [int]$_.Exception.Data['ExitCode']
    }
    elseif ($_.Exception -is [System.UnauthorizedAccessException]) {
        $exitCode = $Script:ExitCodes.PermissionDenied
    }

    Write-Host ''
    Write-Host "処理結果   : 異常終了"
    Write-Host "終了コード : $exitCode"
    Write-Host "エラー内容 : $($_.Exception.Message)"
    exit $exitCode
}
