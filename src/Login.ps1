Set-StrictMode -Version 2.0

# Bảng keyword path của font CAPTCHA SVG — sao y từ ListPathAllKeywords trong
# modDetectCaptcha.bas. Index 0-25 tương ứng A-Z, index 26-35 tương ứng 0-9.
# Các ký tự dễ nhầm (I, L, O, U, 0, 1) để rỗng nên bị bỏ qua khi tra bảng.
$script:CaptchaKeywords = @(
    'MQQQQQZMQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZMQQZ'
    'MQQQQQQQQQZMQQQQQQZMQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQZMQQQQQQQQZMQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQZMQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQZ'
    ''
    'MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZ'
    ''
    'MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZ'
    ''
    'MQQQQQQZMQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQZ'
    'MQQQQQQZMQQQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQZ'
    ''
    'MQQQQQQQQQQZMQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQZMQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZ'
    ''
    ''
    'MQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQZMQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZMQQQQQZ'
    'MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQZ'
    'MQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQZ'
    'MQQQQQQQQZMQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQZMQQQQQQQZ'
    'MQQQQQQQQZMQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQZ'
)

# Số đầu tiên trong chuỗi path, dùng để sắp xếp thứ tự ký tự theo tọa độ
# (tương ứng Val(ms(0).SubMatches(1)) trong VBA).
function ConvertTo-CaptchaPosition {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return 0.0 }
    $match = [regex]::Match($Text, '^\s*(-?(\d+(\.\d+)?|\.\d+))')
    if (-not $match.Success) { return 0.0 }
    $position = 0.0
    if ([double]::TryParse($match.Groups[1].Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$position)) {
        return $position
    }
    return 0.0
}

# Nhận diện mã CAPTCHA từ SVG (detectSVGCaptcha trong VBA): tách từng path,
# rút gọn chuỗi lệnh về keyword M/Q/Z, tra bảng và sắp xếp theo tọa độ.
function ConvertFrom-SvgCaptcha {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    # Phản hồi JSON escape dấu nháy kép thành \" — chuẩn hóa về lại ".
    $normalized = $Text.Replace('\"', '"')

    $dictionary = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([StringComparer]::Ordinal)
    for ($index = 0; $index -lt $script:CaptchaKeywords.Count; $index++) {
        $keyword = $script:CaptchaKeywords[$index]
        if ([string]::IsNullOrEmpty($keyword)) { continue }
        $dictionary[$keyword] = $index
    }

    $commandPattern = New-Object System.Text.RegularExpressions.Regex '(?i)([MQZ])([^MQZ]*)'
    $segments = $normalized -split ' d="'
    $found = New-Object System.Collections.Generic.List[object]
    for ($segmentIndex = 1; $segmentIndex -lt $segments.Count; $segmentIndex++) {
        $segment = $segments[$segmentIndex]
        $quoteIndex = $segment.IndexOf('"')
        if ($quoteIndex -ge 0) { $segment = $segment.Substring(0, $quoteIndex) }

        $matches = $commandPattern.Matches($segment)
        if ($matches.Count -eq 0) { continue }

        $keyword = $commandPattern.Replace($segment, '$1')
        if (-not $dictionary.ContainsKey($keyword)) { continue }

        $keywordIndex = $dictionary[$keyword]
        $character = if ($keywordIndex -le 26) { [string][char]($keywordIndex + 65) } else { [string]($keywordIndex - 26) }
        $found.Add([pscustomobject]@{
            Position = ConvertTo-CaptchaPosition ([string]$matches[0].Groups[2].Value)
            Character = $character
        })
    }

    if ($found.Count -eq 0) { return '' }
    return ((@($found | Sort-Object Position | ForEach-Object { $_.Character })) -join '')
}

# Lấy CAPTCHA mới từ GDT: phản hồi JSON {"key": "...", "content": "<svg .../>"}.
function Get-GdtCaptcha {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Config)

    $uri = '{0}/captcha' -f $Config.BaseUrl
    $responseText = Invoke-GdtRequest -Config $Config -Uri $uri -SkipAuthorization
    $key = ''
    $content = ''
    try {
        $payload = $responseText | ConvertFrom-Json
        $key = Get-JsonTextValue $payload 'key'
        $content = Get-JsonTextValue $payload 'content'
    }
    catch {
        throw 'Phản hồi CAPTCHA từ GDT không hợp lệ.'
    }

    $source = if ([string]::IsNullOrWhiteSpace($content)) { $responseText } else { $content }
    $code = ConvertFrom-SvgCaptcha -Text $source
    return [pscustomobject]@{ Key = $key; Code = $code }
}

# Đăng nhập bằng tài khoản/mật khẩu: lấy CAPTCHA, tự nhận diện rồi gửi kèm để
# lấy token. Tương ứng cmdDangNhap_Click trong frmDangNhap.frm.
function Invoke-GdtLogin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [int]$MaxAttempts = 3
    )

    if ([string]::IsNullOrWhiteSpace([string]$Config.Username) -or [string]::IsNullOrWhiteSpace([string]$Config.Password)) {
        throw 'Chưa cấu hình GDT_USERNAME và GDT_PASSWORD trong .env.'
    }

    $uri = '{0}/security-taxpayer/authenticate' -f $Config.BaseUrl
    $lastError = ''

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; không đăng nhập tiếp.' }

        try {
            $captcha = Get-GdtCaptcha -Config $Config
        }
        catch {
            if (Test-HddtStopRequested) { throw }
            $lastError = $_.Exception.Message
            Write-HddtLog WARN ('[ĐĂNG NHẬP] Lỗi lấy CAPTCHA ({0}/{1}): {2}' -f $attempt, $MaxAttempts, $lastError)
            continue
        }

        if ([string]::IsNullOrWhiteSpace($captcha.Key)) {
            $lastError = 'phản hồi CAPTCHA không có key.'
            Write-HddtLog WARN ('[ĐĂNG NHẬP] CAPTCHA thiếu key ({0}/{1}).' -f $attempt, $MaxAttempts)
            continue
        }
        if ([string]::IsNullOrWhiteSpace($captcha.Code)) {
            $lastError = 'không nhận diện được CAPTCHA.'
            Write-HddtLog WARN ('[ĐĂNG NHẬP] Không nhận diện được CAPTCHA ({0}/{1}); lấy CAPTCHA mới.' -f $attempt, $MaxAttempts)
            continue
        }

        Write-HddtLog INFO ('[ĐĂNG NHẬP] Thử {0}/{1} | tài khoản {2} | CAPTCHA tự nhận: {3}' -f $attempt, $MaxAttempts, $Config.Username, $captcha.Code)
        $body = @{
            username = [string]$Config.Username
            password = [string]$Config.Password
            cvalue = [string]$captcha.Code
            ckey = [string]$captcha.Key
        } | ConvertTo-Json

        try {
            $responseText = Invoke-GdtRequest -Config $Config -Uri $uri -Method Post -Body $body `
                -ContentType 'application/json' -SkipAuthorization
        }
        catch {
            if (Test-HddtStopRequested) { throw }
            $lastError = $_.Exception.Message
            Write-HddtLog WARN ('[ĐĂNG NHẬP] Chưa thành công ({0}/{1}): {2}' -f $attempt, $MaxAttempts, $lastError)
            continue
        }

        $payload = $null
        try { $payload = $responseText | ConvertFrom-Json }
        catch { throw 'Phản hồi đăng nhập từ GDT không hợp lệ.' }

        $messageProperty = $payload.PSObject.Properties['message']
        if ($null -ne $messageProperty -and $null -ne $messageProperty.Value -and -not [string]::IsNullOrWhiteSpace([string]$messageProperty.Value)) {
            $message = [string]$messageProperty.Value
            if ($message -match '(?i)captcha') {
                $lastError = $message
                Write-HddtLog WARN ('[ĐĂNG NHẬP] CAPTCHA không đúng ({0}/{1}): {2}' -f $attempt, $MaxAttempts, $message)
                continue
            }
            throw ('Đăng nhập thất bại: ' + $message)
        }

        $tokenProperty = $payload.PSObject.Properties['token']
        if ($null -eq $tokenProperty -or [string]::IsNullOrWhiteSpace([string]$tokenProperty.Value)) {
            throw 'Phản hồi đăng nhập không chứa token.'
        }
        return [string]$tokenProperty.Value
    }

    throw ('Đăng nhập thất bại sau {0} lần thử: {1}' -f $MaxAttempts, $lastError)
}
