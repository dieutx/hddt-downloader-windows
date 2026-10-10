Set-StrictMode -Version 2.0

# Chi nhung khoa nay duoc luu; khong luu tai khoan, token, mat khau hay proxy.
function Get-HddtDesktopSettingFields {
    return @(
        @{ Key = 'XML_CONCURRENCY'; Default = '4'; Kind = 'Integer' }
        @{ Key = 'XML_MAX_CONCURRENCY'; Default = '4'; Kind = 'Integer' }
        @{ Key = 'XML_REQUEST_INTERVAL_MS'; Default = '800'; Kind = 'Integer' }
        @{ Key = 'XML_RECOVERY_STEP_SECONDS'; Default = '10'; Kind = 'Integer' }
        @{ Key = 'REQUEST_DELAY_MS'; Default = '600'; Kind = 'Integer' }
        @{ Key = 'MAX_RETRIES'; Default = '4'; Kind = 'Integer' }
        @{ Key = 'HTTP_TIMEOUT_SECONDS'; Default = '90'; Kind = 'Integer' }
        @{ Key = 'PAGE_SIZE'; Default = '50'; Kind = 'Integer' }
        @{ Key = 'ADAPTIVE_THROTTLE'; Default = 'true'; Kind = 'Boolean' }
        @{ Key = 'LOG_TO_FILE'; Default = 'true'; Kind = 'Boolean' }
    )
}

function ConvertTo-HddtDesktopSettings {
    param([hashtable]$Values = @{})
    $settings = @{}
    foreach ($field in @(Get-HddtDesktopSettingFields)) {
        $text = (Get-EnvValue $Values $field.Key $field.Default).Trim()
        if ($field.Kind -eq 'Integer') {
            $number = 0
            if ($text -notmatch '^\d+$' -or -not [int]::TryParse($text, [ref]$number)) {
                throw ('{0} can la so nguyen khong am.' -f $field.Key)
            }
            $text = $number.ToString([Globalization.CultureInfo]::InvariantCulture)
        }
        else { $text = (ConvertTo-EnvBoolean $field.Key $text).ToString().ToLowerInvariant() }
        $settings[$field.Key] = $text
    }
    # Dung validator cua CLI de kiem tra khoang va quan he giua cac tham so.
    # Cac truong bat buoc chi la context cuc bo; khong dang nhap/goi mang.
    $validation = Copy-HddtConfigValues $settings
    $validation['GDT_USERNAME'] = 'settings-validation'
    $validation['GDT_PASSWORD'] = 'settings-validation'
    $validation['FROM_DATE'] = '01/01/2000'
    $validation['TO_DATE'] = '01/01/2000'
    $null = Get-HddtConfig -EnvFile (Join-Path $PSScriptRoot '.env') -RepositoryRoot $PSScriptRoot -Values $validation
    return $settings
}

function Read-HddtDesktopSettings {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{} }
    $document = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ($null -eq $document -or $document -is [array] -or $document -isnot [pscustomobject]) { throw 'File cai dat khong hop le.' }
    $values = @{}
    foreach ($field in @(Get-HddtDesktopSettingFields)) {
        $property = $document.PSObject.Properties[$field.Key]
        if ($null -ne $property) { $values[$field.Key] = [string]$property.Value }
    }
    return (ConvertTo-HddtDesktopSettings $values)
}

function Save-HddtDesktopSettings {
    param([Parameter(Mandatory = $true)][string]$Path, [hashtable]$Values = @{})
    $settings = ConvertTo-HddtDesktopSettings $Values
    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllText($temporary, (($settings | ConvertTo-Json) -replace "`r`n", "`n"), (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($temporary, $Path, [Management.Automation.Language.NullString]::Value) }
        else { [IO.File]::Move($temporary, $Path) }
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}
