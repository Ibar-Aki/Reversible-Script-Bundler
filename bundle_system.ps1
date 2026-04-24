[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Bundle', 'Restore', 'RestoreStructure', 'Verify')]
    [string]$Mode,
    [string]$InputPath,
    [string]$OutputPath,
    [string]$RestoreInputPath,
    [string]$RestoreOutputPath,
    [string]$IgnoreFilePath,
    [string]$BundleRootName,
    [string]$ResultJsonPath
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
    VerifyMismatch      = 25
    PermissionDenied    = 30
    Unexpected          = 99
}

$Script:FormatName = 'BatchPsBundle'
$Script:FormatVersion = '1.2'
$Script:ToolVersion = '1.2.0'
$Script:Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$Script:ReservedNames = @(
    'CON', 'PRN', 'AUX', 'NUL',
    'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
    'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
)
$Script:AllowedExtensions = @('.bat', '.ps1', '.psm1', '.md', '.json', '.jsonl', '.bas')
$Script:AllowedNewlineStyles = @('None', 'CRLF', 'LF', 'CR', 'Mixed')
$Script:AllowedBomTypes = @('None', 'UTF32-LE', 'UTF32-BE', 'UTF8-BOM', 'UTF16-LE', 'UTF16-BE')
$Script:Hostname = [System.Environment]::MachineName
$Script:ResultPath = $null
$Script:ScriptRoot = Split-Path -Parent $PSCommandPath

function Throw-HandledError {
    param(
        [int]$Code,
        [string]$Message
    )

    $exception = New-Object System.Exception($Message)
    $exception.Data['ExitCode'] = $Code
    throw $exception
}

function Get-CurrentUserName {
    try {
        return [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        if ([string]::IsNullOrWhiteSpace($env:USERNAME)) {
            return 'unknown-user'
        }

        return $env:USERNAME
    }
}

$Script:CreatedBy = Get-CurrentUserName

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        [System.IO.Directory]::CreateDirectory($Path) | Out-Null
    }
}

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path)
    }
    catch [System.UnauthorizedAccessException] {
        Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "権限不足でパスへアクセスできません: $Path"
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "パスが存在しません: $Path"
    }
}

function Get-DefaultRootPath {
    return [System.IO.Path]::GetFullPath($Script:ScriptRoot)
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

function Join-BundleRelativePath {
    param(
        [string]$BundleRootName,
        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($BundleRootName)) {
        return $RelativePath
    }

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        return $BundleRootName
    }

    return (Join-Path -Path $BundleRootName -ChildPath $RelativePath)
}

function Get-BundleRelativePath {
    param(
        [string]$InputRoot,
        [string]$TargetPath,
        [string]$BundleRootName
    )

    $relativePath = Get-RelativePath -BasePath $InputRoot -TargetPath $TargetPath
    return Join-BundleRelativePath -BundleRootName $BundleRootName -RelativePath $relativePath
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

function Convert-ToSafeFileNameSegment {
    param([string]$Value)

    $safeValue = $Value
    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safeValue = $safeValue.Replace([string]$invalidChar, '_')
    }

    $safeValue = $safeValue.Trim()
    if ([string]::IsNullOrWhiteSpace($safeValue)) {
        return 'input_files'
    }

    return $safeValue
}

function Get-BundleBaseName {
    param(
        [string]$InputDirectory,
        [bool]$UseDirectoryLeafName
    )

    $datePart = Get-Date -Format 'yyMMdd'
    if ($UseDirectoryLeafName) {
        $namePart = Convert-ToSafeFileNameSegment -Value (Split-Path -Leaf $InputDirectory)
        return "bundle_${datePart}_$namePart"
    }

    $topLevelDirectories = @(
        Get-ChildItem -LiteralPath $InputDirectory -Directory |
            Sort-Object Name
    )
    $topLevelFiles = @(
        Get-ChildItem -LiteralPath $InputDirectory -File |
            Where-Object { -not (Test-IgnoredInputFile -File $_) } |
            Sort-Object Name
    )

    if ($topLevelDirectories.Count -eq 1 -and $topLevelFiles.Count -eq 0) {
        $namePart = Convert-ToSafeFileNameSegment -Value $topLevelDirectories[0].Name
    }
    else {
        $namePart = 'input_files'
    }

    return "bundle_${datePart}_$namePart"
}

function Get-UniqueBundlePath {
    param(
        [string]$OutputDirectory,
        [string]$InputDirectory,
        [bool]$UseDirectoryLeafName
    )

    $baseName = Get-BundleBaseName -InputDirectory $InputDirectory -UseDirectoryLeafName $UseDirectoryLeafName
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
    if ($name -in @('thumbs.db', 'desktop.ini', '.gitkeep', '.bundleignore')) {
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

function Get-CreatedAt {
    return (Get-Date).ToString('o')
}

function Write-ResultObject {
    param([hashtable]$ResultObject)

    if ([string]::IsNullOrWhiteSpace($Script:ResultPath)) {
        return
    }

    $directoryPath = Split-Path -Parent $Script:ResultPath
    if (-not [string]::IsNullOrWhiteSpace($directoryPath)) {
        Ensure-Directory -Path $directoryPath
    }

    $json = $ResultObject | ConvertTo-Json -Depth 10
    Write-TextFile -Path $Script:ResultPath -Content $json
}

function Show-Start {
    param(
        [string]$Operation,
        [System.Collections.IDictionary]$Paths
    )

    Write-Host ('=' * 60)
    Write-Host "処理種別 : $Operation"
    Write-Host "入力元     : $($Paths.InputFiles)"
    if ($Operation -like '集約*' -or $Operation -eq '検証') {
        Write-Host "出力先     : $($Paths.OutputBundle)"
        if (-not [string]::IsNullOrWhiteSpace($Paths.IgnoreRulePath)) {
            Write-Host "ignore     : $($Paths.IgnoreRulePath)"
        }
    }
    else {
        Write-Host "集約入力元 : $($Paths.RestoreInput)"
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

function Resolve-BundlePaths {
    $scriptRoot = Get-DefaultRootPath
    $isDefaultInputPath = [string]::IsNullOrWhiteSpace($InputPath)
    $resolvedInputPath = if ($isDefaultInputPath) { Join-Path -Path $scriptRoot -ChildPath 'input_files' } else { Resolve-ExistingPath -Path $InputPath }
    $resolvedOutputPath = if ([string]::IsNullOrWhiteSpace($OutputPath)) { Join-Path -Path $scriptRoot -ChildPath 'output_bundle' } else { [System.IO.Path]::GetFullPath($OutputPath) }
    $effectiveIgnoreFilePath = $null
    $effectiveBundleRootName = $null

    if (-not (Test-Path -LiteralPath $resolvedInputPath)) {
        if ($isDefaultInputPath) {
            Ensure-Directory -Path $resolvedInputPath
        }
        else {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "InputPath が存在しません: $resolvedInputPath"
        }
    }

    $inputItem = Get-Item -LiteralPath $resolvedInputPath -Force
    if (-not $inputItem.PSIsContainer) {
        Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "InputPath はフォルダである必要があります: $resolvedInputPath"
    }

    Ensure-Directory -Path $resolvedOutputPath

    if (-not [string]::IsNullOrWhiteSpace($IgnoreFilePath)) {
        $effectiveIgnoreFilePath = Resolve-ExistingPath -Path $IgnoreFilePath
        if (-not (Get-Item -LiteralPath $effectiveIgnoreFilePath -Force).PSIsContainer) {
            # no-op
        }
        else {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "IgnoreFilePath はファイルである必要があります: $effectiveIgnoreFilePath"
        }
    }
    else {
        $defaultIgnore = Join-Path -Path $resolvedInputPath -ChildPath '.bundleignore'
        if (Test-Path -LiteralPath $defaultIgnore -PathType Leaf) {
            $effectiveIgnoreFilePath = [System.IO.Path]::GetFullPath($defaultIgnore)
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($BundleRootName)) {
        $effectiveBundleRootName = Convert-ToSafeFileNameSegment -Value $BundleRootName
        if ($effectiveBundleRootName -ne $BundleRootName.Trim()) {
            Throw-HandledError -Code $Script:ExitCodes.InvalidPath -Message "BundleRootName にファイル名として使えない文字が含まれています: $BundleRootName"
        }
    }

    return [ordered]@{
        Root                 = $scriptRoot
        InputFiles           = [System.IO.Path]::GetFullPath($resolvedInputPath)
        OutputBundle         = [System.IO.Path]::GetFullPath($resolvedOutputPath)
        IgnoreRulePath       = $effectiveIgnoreFilePath
        BundleRootName       = $effectiveBundleRootName
        UseDirectoryLeafName = ($isDefaultInputPath -eq $false) -or ((Split-Path -Leaf $resolvedInputPath) -ne 'input_files')
    }
}

function Resolve-RestorePaths {
    $scriptRoot = Get-DefaultRootPath
    $resolvedRestoreInputPath = $null
    $explicitBundleFile = $null

    if ([string]::IsNullOrWhiteSpace($RestoreInputPath)) {
        $resolvedRestoreInputPath = Join-Path -Path $scriptRoot -ChildPath 'restore_input'
        Ensure-Directory -Path $resolvedRestoreInputPath
    }
    else {
        $restoreInputFullPath = Resolve-ExistingPath -Path $RestoreInputPath
        $restoreInputItem = Get-Item -LiteralPath $restoreInputFullPath -Force
        if ($restoreInputItem.PSIsContainer) {
            $resolvedRestoreInputPath = $restoreInputFullPath
        }
        else {
            $explicitBundleFile = $restoreInputFullPath
            $resolvedRestoreInputPath = Split-Path -Parent $restoreInputFullPath
        }
    }

    $resolvedRestoreOutputPath = if ([string]::IsNullOrWhiteSpace($RestoreOutputPath)) { Join-Path -Path $scriptRoot -ChildPath 'restore_output' } else { [System.IO.Path]::GetFullPath($RestoreOutputPath) }
    Ensure-Directory -Path $resolvedRestoreOutputPath

    return [ordered]@{
        Root             = $scriptRoot
        InputFiles       = $null
        OutputBundle     = $null
        RestoreInput     = [System.IO.Path]::GetFullPath($resolvedRestoreInputPath)
        RestoreOutput    = [System.IO.Path]::GetFullPath($resolvedRestoreOutputPath)
        ExplicitBundle   = $explicitBundleFile
        IgnoreRulePath   = $null
        UseDirectoryLeafName = $false
    }
}

function Resolve-VerifyPaths {
    $bundlePaths = Resolve-BundlePaths
    $verifyRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("bundle_verify_{0}" -f ([Guid]::NewGuid().ToString('N')))
    Ensure-Directory -Path $verifyRoot

    $restoreInputPath = if ([string]::IsNullOrWhiteSpace($RestoreInputPath)) { Join-Path -Path $verifyRoot -ChildPath 'restore_input' } else { [System.IO.Path]::GetFullPath($RestoreInputPath) }
    $restoreOutputPath = if ([string]::IsNullOrWhiteSpace($RestoreOutputPath)) { Join-Path -Path $verifyRoot -ChildPath 'restore_output' } else { [System.IO.Path]::GetFullPath($RestoreOutputPath) }
    Ensure-Directory -Path $restoreInputPath
    Ensure-Directory -Path $restoreOutputPath

    $bundlePaths['RestoreInput'] = [System.IO.Path]::GetFullPath($restoreInputPath)
    $bundlePaths['RestoreOutput'] = [System.IO.Path]::GetFullPath($restoreOutputPath)
    $bundlePaths['VerifyWorkspace'] = [System.IO.Path]::GetFullPath($verifyRoot)
    $bundlePaths['CleanupVerifyWorkspace'] = $true
    return $bundlePaths
}

function Read-IgnoreRules {
    param([string]$IgnoreRulePath)

    if ([string]::IsNullOrWhiteSpace($IgnoreRulePath)) {
        return @()
    }

    try {
        $lines = Get-Content -LiteralPath $IgnoreRulePath -ErrorAction Stop
    }
    catch [System.UnauthorizedAccessException] {
        Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "ignore ファイルの読み取り権限がありません: $IgnoreRulePath"
    }
    catch {
        Throw-HandledError -Code $Script:ExitCodes.ReadFailure -Message "ignore ファイルを読み取れませんでした: $IgnoreRulePath"
    }

    $rules = New-Object System.Collections.ArrayList
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.StartsWith('#')) { continue }

        $isDirectoryRule = $trimmed.EndsWith('/') -or $trimmed.EndsWith('\')
        $normalized = $trimmed.TrimStart('.', '\', '/').Replace('\', '/').TrimEnd('/')
        if ([string]::IsNullOrWhiteSpace($normalized)) { continue }

        $hasWildcard = $normalized.IndexOfAny(@([char]'*', [char]'?', [char]'[')) -ge 0
        $hasPathSeparator = $normalized.Contains('/')

        $ruleType = if (-not $hasWildcard -and -not $hasPathSeparator) {
            'Segment'
        }
        elseif (-not $hasWildcard) {
            'Prefix'
        }
        else {
            'Wildcard'
        }

        $patternText = if ($ruleType -eq 'Wildcard' -and $isDirectoryRule) { "$normalized/*" } else { $normalized }

        [void]$rules.Add([pscustomobject]@{
            Original        = $trimmed
            Normalized      = $normalized.ToLowerInvariant()
            RuleType        = $ruleType
            IsDirectoryRule = $isDirectoryRule
            Pattern         = if ($ruleType -eq 'Wildcard') {
                [System.Management.Automation.WildcardPattern]::new($patternText.ToLowerInvariant(), [System.Management.Automation.WildcardOptions]::IgnoreCase)
            }
            else {
                $null
            }
        })
    }

    return @($rules)
}

function Test-IgnoreRuleMatch {
    param(
        [string]$RelativePath,
        [bool]$IsDirectory,
        [object[]]$Rules
    )

    if ($null -eq $Rules -or @($Rules).Count -eq 0) {
        return $false
    }

    $normalizedRelativePath = $RelativePath.Replace('\', '/').Trim('/').ToLowerInvariant()
    $segments = if ([string]::IsNullOrWhiteSpace($normalizedRelativePath)) { @() } else { $normalizedRelativePath.Split('/') }

    foreach ($rule in $Rules) {
        switch ($rule.RuleType) {
            'Segment' {
                if ($segments -contains $rule.Normalized) {
                    return $true
                }
            }
            'Prefix' {
                if ($normalizedRelativePath -eq $rule.Normalized -or $normalizedRelativePath.StartsWith($rule.Normalized + '/')) {
                    return $true
                }
            }
            'Wildcard' {
                if ($rule.Pattern.IsMatch($normalizedRelativePath)) {
                    return $true
                }
                if ($rule.IsDirectoryRule -and ($normalizedRelativePath -eq $rule.Normalized -or $normalizedRelativePath.StartsWith($rule.Normalized + '/'))) {
                    return $true
                }
            }
        }
    }

    return $false
}

function Get-InputSnapshot {
    param(
        [string]$InputRoot,
        [object[]]$IgnoreRules,
        [string]$BundleRootName
    )

    $ignoredBySystem = New-Object System.Collections.ArrayList
    $ignoredByRulesFiles = New-Object System.Collections.ArrayList
    $ignoredByRulesDirectories = New-Object System.Collections.ArrayList
    $candidateFiles = New-Object System.Collections.ArrayList
    $directoryEntries = New-Object System.Collections.ArrayList

    if (-not [string]::IsNullOrWhiteSpace($BundleRootName)) {
        [void]$directoryEntries.Add($BundleRootName)
    }

    $pendingDirectories = New-Object System.Collections.Queue
    $pendingDirectories.Enqueue((Get-Item -LiteralPath $InputRoot))

    while ($pendingDirectories.Count -gt 0) {
        $currentDirectory = $pendingDirectories.Dequeue()

        try {
            $childDirectories = @(Get-ChildItem -LiteralPath $currentDirectory.FullName -Directory -ErrorAction Stop | Sort-Object FullName)
            $childFiles = @(Get-ChildItem -LiteralPath $currentDirectory.FullName -File -ErrorAction Stop | Sort-Object FullName)
        }
        catch [System.UnauthorizedAccessException] {
            Throw-HandledError -Code $Script:ExitCodes.PermissionDenied -Message "読み取り権限がありません: $($currentDirectory.FullName)"
        }
        catch {
            Throw-HandledError -Code $Script:ExitCodes.ReadFailure -Message "フォルダを読み取れませんでした: $($currentDirectory.FullName)"
        }

        foreach ($directory in $childDirectories) {
            $relativePath = Get-RelativePath -BasePath $InputRoot -TargetPath $directory.FullName
            if (Test-IgnoreRuleMatch -RelativePath $relativePath -IsDirectory $true -Rules $IgnoreRules) {
                [void]$ignoredByRulesDirectories.Add($relativePath)
                continue
            }

            [void]$directoryEntries.Add((Join-BundleRelativePath -BundleRootName $BundleRootName -RelativePath $relativePath))
            $pendingDirectories.Enqueue($directory)
        }

        foreach ($file in $childFiles) {
            $relativePath = Get-RelativePath -BasePath $InputRoot -TargetPath $file.FullName
            if (Test-IgnoredInputFile -File $file) {
                [void]$ignoredBySystem.Add($relativePath)
                continue
            }
            if (Test-IgnoreRuleMatch -RelativePath $relativePath -IsDirectory $false -Rules $IgnoreRules) {
                [void]$ignoredByRulesFiles.Add($relativePath)
                continue
            }
            [void]$candidateFiles.Add($file)
        }
    }

    $supportedFiles = @($candidateFiles | Where-Object { $Script:AllowedExtensions -contains $_.Extension.ToLowerInvariant() } | Sort-Object FullName)
    $invalidFiles = @($candidateFiles | Where-Object { $Script:AllowedExtensions -notcontains $_.Extension.ToLowerInvariant() } | Sort-Object FullName)

    return [pscustomobject]@{
        DirectoryRelativePaths = @($directoryEntries | Sort-Object)
        IgnoredSystemPaths     = @($ignoredBySystem | Sort-Object)
        IgnoredByRulesFiles    = @($ignoredByRulesFiles | Sort-Object)
        IgnoredByRulesDirs     = @($ignoredByRulesDirectories | Sort-Object)
        SupportedFiles         = $supportedFiles
        InvalidFiles           = $invalidFiles
    }
}

function Invoke-Bundle {
    param([System.Collections.IDictionary]$Paths)

    Show-Start -Operation '集約' -Paths $Paths

    $ignoreRules = Read-IgnoreRules -IgnoreRulePath $Paths.IgnoreRulePath
    $snapshot = Get-InputSnapshot -InputRoot $Paths.InputFiles -IgnoreRules $ignoreRules -BundleRootName $Paths.BundleRootName

    if ($snapshot.SupportedFiles.Count -eq 0 -and $snapshot.InvalidFiles.Count -eq 0 -and $snapshot.DirectoryRelativePaths.Count -eq 0) {
        Throw-HandledError -Code $Script:ExitCodes.NoInputFiles -Message "対象ファイルなし: $($Paths.InputFiles) に対象ファイルまたはフォルダを配置してください。"
    }

    if ($snapshot.InvalidFiles.Count -gt 0) {
        Write-Host "補足       : 変換対象外の $($snapshot.InvalidFiles.Count) 件をスキップします。"
        Show-PathList -Label '変換対象外ファイル一覧:' -Paths @($snapshot.InvalidFiles | ForEach-Object { Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $_.FullName })
    }

    if ($snapshot.SupportedFiles.Count -eq 0 -and $snapshot.InvalidFiles.Count -gt 0) {
        $invalidList = ($snapshot.InvalidFiles | Select-Object -ExpandProperty FullName) -join ', '
        Throw-HandledError -Code $Script:ExitCodes.InvalidExtension -Message "変換可能な対象ファイルがありません。対象外ファイルのみです: $invalidList"
    }

    $directoryEntries = New-Object System.Collections.ArrayList
    $directoryId = 1
    foreach ($relativeDirectoryPath in $snapshot.DirectoryRelativePaths) {
        [void]$directoryEntries.Add([ordered]@{
            id           = $directoryId
            relativePath = $relativeDirectoryPath
        })
        $directoryId += 1
    }

    $bundleEntries = New-Object System.Collections.ArrayList
    $sortedFiles = @($snapshot.SupportedFiles | Sort-Object { Get-RelativePath -BasePath $Paths.InputFiles -TargetPath $_.FullName })
    $id = 1

    foreach ($file in $sortedFiles) {
        $relativePath = Get-BundleRelativePath -InputRoot $Paths.InputFiles -TargetPath $file.FullName -BundleRootName $Paths.BundleRootName
        $bytes = Read-Bytes -Path $file.FullName
        [void]$bundleEntries.Add([ordered]@{
            id            = $id
            relativePath  = $relativePath
            fileName      = $file.Name
            extension     = $file.Extension.ToLowerInvariant()
            byteLength    = $bytes.Length
            sha256        = Get-Sha256Hex -Bytes $bytes
            newlineStyle  = Get-NewlineStyle -Bytes $bytes
            bomType       = Get-BomType -Bytes $bytes
            contentBase64 = [System.Convert]::ToBase64String($bytes)
        })
        $id += 1
    }

    $bundleObject = [ordered]@{
        format              = $Script:FormatName
        version             = $Script:FormatVersion
        toolVersion         = $Script:ToolVersion
        createdAt           = Get-CreatedAt
        createdBy           = $Script:CreatedBy
        hostname            = $Script:Hostname
        sourceRoot          = [System.IO.Path]::GetFullPath($Paths.InputFiles)
        excludedDirectories = @($snapshot.IgnoredByRulesDirs)
        dirCount            = $directoryEntries.Count
        directories         = @($directoryEntries)
        fileCount           = $bundleEntries.Count
        files               = @($bundleEntries)
    }

    $outputPath = Get-UniqueBundlePath -OutputDirectory $Paths.OutputBundle -InputDirectory $Paths.InputFiles -UseDirectoryLeafName $Paths.UseDirectoryLeafName
    $json = $bundleObject | ConvertTo-Json -Depth 6
    Write-TextFile -Path $outputPath -Content $json

    Show-Summary -Operation '集約' -TargetCount $bundleEntries.Count -SuccessCount $bundleEntries.Count -FailureCount 0 -OutputPath $outputPath
    Write-Host "対象フォルダ件数 : $($directoryEntries.Count)"
    if ($snapshot.IgnoredSystemPaths.Count -gt 0) {
        Write-Host "補足       : 除外ルールに一致した $($snapshot.IgnoredSystemPaths.Count) 件のファイルを無視しました。"
        Show-PathList -Label '除外ファイル一覧:' -Paths $snapshot.IgnoredSystemPaths
    }
    if ($snapshot.IgnoredByRulesDirs.Count -gt 0 -or $snapshot.IgnoredByRulesFiles.Count -gt 0) {
        Write-Host "補足       : .bundleignore に一致したフォルダ $($snapshot.IgnoredByRulesDirs.Count) 件、ファイル $($snapshot.IgnoredByRulesFiles.Count) 件を無視しました。"
        Show-PathList -Label 'ignore 対象フォルダ一覧:' -Paths $snapshot.IgnoredByRulesDirs
        Show-PathList -Label 'ignore 対象ファイル一覧:' -Paths $snapshot.IgnoredByRulesFiles
    }

    return [ordered]@{
        Status               = 'Success'
        Operation            = 'Bundle'
        ExitCode             = $Script:ExitCodes.Success
        ToolVersion          = $Script:ToolVersion
        FormatVersion        = $Script:FormatVersion
        CreatedAt            = Get-CreatedAt
        CreatedBy            = $Script:CreatedBy
        Hostname             = $Script:Hostname
        SourceRoot           = [System.IO.Path]::GetFullPath($Paths.InputFiles)
        OutputDirectory      = [System.IO.Path]::GetFullPath($Paths.OutputBundle)
        BundlePath           = $outputPath
        BundledFileCount     = $bundleEntries.Count
        BundledDirectoryCount = $directoryEntries.Count
        SupportedExtensions  = $Script:AllowedExtensions
        IgnoreFilePath       = $Paths.IgnoreRulePath
        ExcludedDirectories  = @($snapshot.IgnoredByRulesDirs)
        IgnoredFilePaths     = @($snapshot.IgnoredSystemPaths + $snapshot.IgnoredByRulesFiles)
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

function Get-BundleCandidates {
    param(
        [string]$RestoreInputPath,
        [string]$ExplicitBundlePath
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitBundlePath)) {
        return ,(Get-Item -LiteralPath $ExplicitBundlePath -Force)
    }

    return @(Get-ChildItem -LiteralPath $RestoreInputPath -File -Filter 'bundle*.txt' | Sort-Object Name)
}

function Invoke-Restore {
    param(
        [System.Collections.IDictionary]$Paths,
        [string]$OperationLabel
    )

    Show-Start -Operation $OperationLabel -Paths $Paths

    $bundleCandidates = @(Get-BundleCandidates -RestoreInputPath $Paths.RestoreInput -ExplicitBundlePath $Paths.ExplicitBundle)
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
    $sourceRoot = if ($bundleObject.PSObject.Properties['sourceRoot']) { [string]$bundleObject.sourceRoot } else { $null }
    $excludedDirectories = if ($bundleObject.PSObject.Properties['excludedDirectories']) { @($bundleObject.excludedDirectories) } else { @() }
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
        Show-Summary -Operation $OperationLabel -TargetCount $directoryPlan.Count -SuccessCount $directoryPlan.Count -FailureCount 0 -OutputPath $Paths.RestoreOutput
        Write-Host "対象フォルダ件数 : $($directoryPlan.Count)"

        return [ordered]@{
            Status                = 'Success'
            Operation             = 'RestoreStructure'
            ExitCode              = $Script:ExitCodes.Success
            ToolVersion           = $Script:ToolVersion
            CreatedAt             = Get-CreatedAt
            CreatedBy             = $Script:CreatedBy
            Hostname              = $Script:Hostname
            BundlePath            = $bundlePath
            RestoreOutputPath     = [System.IO.Path]::GetFullPath($Paths.RestoreOutput)
            RestoredFileCount     = 0
            RestoredDirectoryCount = $directoryPlan.Count
            SourceRoot            = $sourceRoot
            ExcludedDirectories   = $excludedDirectories
        }
    }

    foreach ($item in $restorePlan) {
        Write-BytesFile -Path $item.TargetPath -Bytes $item.Bytes
    }

    Show-Summary -Operation $OperationLabel -TargetCount $restorePlan.Count -SuccessCount $restorePlan.Count -FailureCount 0 -OutputPath $Paths.RestoreOutput
    Write-Host "対象フォルダ件数 : $($directoryPlan.Count)"

    return [ordered]@{
        Status                = 'Success'
        Operation             = 'Restore'
        ExitCode              = $Script:ExitCodes.Success
        ToolVersion           = $Script:ToolVersion
        CreatedAt             = Get-CreatedAt
        CreatedBy             = $Script:CreatedBy
        Hostname              = $Script:Hostname
        BundlePath            = $bundlePath
        RestoreOutputPath     = [System.IO.Path]::GetFullPath($Paths.RestoreOutput)
        RestoredFileCount     = $restorePlan.Count
        RestoredDirectoryCount = $directoryPlan.Count
        SourceRoot            = $sourceRoot
        ExcludedDirectories   = $excludedDirectories
    }
}

function Get-RestoreSnapshot {
    param([string]$RestoreRoot)

    $fileMap = @{}
    $files = @(Get-ChildItem -LiteralPath $RestoreRoot -File -Recurse | Sort-Object FullName)
    foreach ($file in $files) {
        $relativePath = Get-RelativePath -BasePath $RestoreRoot -TargetPath $file.FullName
        $fileMap[$relativePath] = Get-Sha256Hex -Bytes (Read-Bytes -Path $file.FullName)
    }

    return [pscustomobject]@{
        FileMap = $fileMap
        DirectoryRelativePaths = @(
            Get-ChildItem -LiteralPath $RestoreRoot -Directory -Recurse |
                ForEach-Object { Get-RelativePath -BasePath $RestoreRoot -TargetPath $_.FullName } |
                Sort-Object
        )
    }
}

function Compare-StringArrays {
    param(
        [string[]]$Expected,
        [string[]]$Actual
    )

    if ($Expected.Count -ne $Actual.Count) {
        return $false
    }

    for ($index = 0; $index -lt $Expected.Count; $index += 1) {
        if ($Expected[$index] -ne $Actual[$index]) {
            return $false
        }
    }

    return $true
}

function Invoke-Verify {
    param([System.Collections.IDictionary]$Paths)

    Show-Start -Operation '検証' -Paths $Paths

    $ignoreRules = Read-IgnoreRules -IgnoreRulePath $Paths.IgnoreRulePath
    $sourceSnapshot = Get-InputSnapshot -InputRoot $Paths.InputFiles -IgnoreRules $ignoreRules -BundleRootName $Paths.BundleRootName

    $bundleResult = Invoke-Bundle -Paths $Paths

    $bundleTarget = Join-Path -Path $Paths.RestoreInput -ChildPath ([System.IO.Path]::GetFileName($bundleResult.BundlePath))
    Copy-Item -LiteralPath $bundleResult.BundlePath -Destination $bundleTarget -Force

    $restorePaths = [ordered]@{
        RestoreInput        = $Paths.RestoreInput
        RestoreOutput       = $Paths.RestoreOutput
        ExplicitBundle      = $bundleTarget
        InputFiles          = $Paths.InputFiles
        OutputBundle        = $Paths.OutputBundle
        IgnoreRulePath      = $Paths.IgnoreRulePath
        UseDirectoryLeafName = $Paths.UseDirectoryLeafName
    }
    $restoreResult = Invoke-Restore -Paths $restorePaths -OperationLabel '検証復元'

    $restoredSnapshot = Get-RestoreSnapshot -RestoreRoot $Paths.RestoreOutput
    $expectedFileMap = @{}
    foreach ($file in $sourceSnapshot.SupportedFiles) {
        $relativePath = Get-BundleRelativePath -InputRoot $Paths.InputFiles -TargetPath $file.FullName -BundleRootName $Paths.BundleRootName
        $expectedFileMap[$relativePath] = Get-Sha256Hex -Bytes (Read-Bytes -Path $file.FullName)
    }

    if ($expectedFileMap.Count -ne $restoredSnapshot.FileMap.Count) {
        Throw-HandledError -Code $Script:ExitCodes.VerifyMismatch -Message "検証失敗: 復元ファイル件数が一致しません。期待 $($expectedFileMap.Count) 件 / 実際 $($restoredSnapshot.FileMap.Count) 件"
    }

    foreach ($relativePath in $expectedFileMap.Keys) {
        if (-not $restoredSnapshot.FileMap.ContainsKey($relativePath)) {
            Throw-HandledError -Code $Script:ExitCodes.VerifyMismatch -Message "検証失敗: 復元結果にファイルがありません: $relativePath"
        }
        if ($restoredSnapshot.FileMap[$relativePath] -ne $expectedFileMap[$relativePath]) {
            Throw-HandledError -Code $Script:ExitCodes.VerifyMismatch -Message "検証失敗: SHA-256 が一致しません: $relativePath"
        }
    }

    $expectedDirectories = @($sourceSnapshot.DirectoryRelativePaths | Sort-Object)
    $actualDirectories = @($restoredSnapshot.DirectoryRelativePaths | Sort-Object)
    if (-not (Compare-StringArrays -Expected $expectedDirectories -Actual $actualDirectories)) {
        Throw-HandledError -Code $Script:ExitCodes.VerifyMismatch -Message '検証失敗: 復元フォルダ構成が一致しません。'
    }

    Write-Host ''
    Write-Host '検証結果   : 正常終了'
    Write-Host "検証対象   : $($expectedFileMap.Count) ファイル / $($expectedDirectories.Count) フォルダ"
    Write-Host "bundle     : $($bundleResult.BundlePath)"
    Write-Host "restore    : $($restoreResult.RestoreOutputPath)"

    return [ordered]@{
        Status                = 'Success'
        Operation             = 'Verify'
        ExitCode              = $Script:ExitCodes.Success
        ToolVersion           = $Script:ToolVersion
        CreatedAt             = Get-CreatedAt
        CreatedBy             = $Script:CreatedBy
        Hostname              = $Script:Hostname
        SourceRoot            = [System.IO.Path]::GetFullPath($Paths.InputFiles)
        BundlePath            = $bundleResult.BundlePath
        RestoreOutputPath     = $restoreResult.RestoreOutputPath
        VerifiedFileCount     = $expectedFileMap.Count
        VerifiedDirectoryCount = $expectedDirectories.Count
        IgnoreFilePath        = $Paths.IgnoreRulePath
        ExcludedDirectories   = @($sourceSnapshot.IgnoredByRulesDirs)
        VerifyWorkspace       = $Paths.VerifyWorkspace
    }
}

function Remove-VerifyWorkspaceIfInternal {
    param([System.Collections.IDictionary]$Paths)

    if ($null -eq $Paths) {
        return
    }

    if (-not $Paths.Contains('CleanupVerifyWorkspace') -or -not $Paths.CleanupVerifyWorkspace) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($Paths.VerifyWorkspace)) {
        return
    }

    if (Test-Path -LiteralPath $Paths.VerifyWorkspace) {
        try {
            Remove-Item -LiteralPath $Paths.VerifyWorkspace -Recurse -Force
        }
        catch {
            Write-Host "補足       : 検証用一時フォルダを削除できませんでした: $($Paths.VerifyWorkspace)"
        }
    }
}

try {
    $paths = $null
    $Script:ResultPath = if ([string]::IsNullOrWhiteSpace($ResultJsonPath)) { $null } else { [System.IO.Path]::GetFullPath($ResultJsonPath) }

    switch ($Mode) {
        'Bundle' {
            $paths = Resolve-BundlePaths
            $result = Invoke-Bundle -Paths $paths
        }
        'Restore' {
            $paths = Resolve-RestorePaths
            $result = Invoke-Restore -Paths $paths -OperationLabel '復元'
        }
        'RestoreStructure' {
            $paths = Resolve-RestorePaths
            $result = Invoke-Restore -Paths $paths -OperationLabel 'フォルダ構成復元'
        }
        'Verify' {
            $paths = Resolve-VerifyPaths
            $result = Invoke-Verify -Paths $paths
        }
        default {
            Throw-HandledError -Code $Script:ExitCodes.Unexpected -Message "未対応のモードです: $Mode"
        }
    }

    Remove-VerifyWorkspaceIfInternal -Paths $paths
    Write-ResultObject -ResultObject $result
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

    $errorResult = [ordered]@{
        Status      = 'Error'
        Operation   = $Mode
        ExitCode    = $exitCode
        ToolVersion = $Script:ToolVersion
        CreatedAt   = Get-CreatedAt
        CreatedBy   = $Script:CreatedBy
        Hostname    = $Script:Hostname
        Message     = $_.Exception.Message
    }
    Remove-VerifyWorkspaceIfInternal -Paths $paths
    Write-ResultObject -ResultObject $errorResult
    exit $exitCode
}
