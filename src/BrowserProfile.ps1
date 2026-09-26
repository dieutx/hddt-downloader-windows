Set-StrictMode -Version 2.0

# Browser profile dùng chung cho toàn phiên chạy: một User-Agent, một cặp
# client hint (sec-ch-ua) và một locale. Không random, không đổi theo request,
# không đổi sau HTTP 429 — một lần chạy giống một phiên Edge/Chromium trên
# Windows. BROWSER_USER_AGENT trong .env override được default.
$script:GdtDefaultBrowserUserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0'
$script:GdtBrowserProfile = $null

# Nguồn duy nhất của User-Agent: BROWSER_USER_AGENT nếu có, ngược lại default.
function Get-GdtBrowserUserAgent {
    param($Config)
    $value = [string](Get-HddtConfigValue $Config 'BrowserUserAgent' '')
    if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
    return $script:GdtDefaultBrowserUserAgent
}

# sec-ch-ua phải khớp với User-Agent đang dùng (cùng major version, cùng brand)
# để không tạo mismatch giữa UA và client hint.
function New-GdtBrowserProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$UserAgent)

    $majorVersion = ''
    if ($UserAgent -match 'Edg[A-Za-z]*\/(\d+)') { $majorVersion = $Matches[1] }
    elseif ($UserAgent -match 'Chrome\/(\d+)') { $majorVersion = $Matches[1] }
    if ([string]::IsNullOrEmpty($majorVersion)) { $majorVersion = '126' }

    if ($UserAgent -match 'Edg[A-Za-z]*\/\d+') {
        $secChUa = '"Chromium";v="{0}", "Microsoft Edge";v="{0}", "Not.A/Brand";v="24"' -f $majorVersion
    }
    elseif ($UserAgent -match 'Chrome\/\d+') {
        $secChUa = '"Chromium";v="{0}", "Google Chrome";v="{0}", "Not.A/Brand";v="24"' -f $majorVersion
    }
    else {
        $secChUa = '"Chromium";v="{0}", "Not-A.Brand";v="99"' -f $majorVersion
    }

    return [pscustomobject]@{
        UserAgent = $UserAgent
        SecChUa = $secChUa
        SecChUaMobile = '?0'
        SecChUaPlatform = '"Windows"'
        AcceptLanguage = 'vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7'
        Origin = 'https://hoadondientu.gdt.gov.vn'
        RootReferer = 'https://hoadondientu.gdt.gov.vn/'
        LookupReferer = 'https://hoadondientu.gdt.gov.vn/tra-cuu/tra-cuu-hoa-don'
    }
}

# Profile được tạo tối đa một lần cho mỗi User-Agent của logical run; các request
# sau đó lấy cùng một object (không tạo lại theo từng request).
function Get-GdtBrowserProfile {
    param($Config)
    $userAgent = Get-GdtBrowserUserAgent -Config $Config
    if ($null -ne $script:GdtBrowserProfile -and [string]$script:GdtBrowserProfile.UserAgent -eq $userAgent) {
        return $script:GdtBrowserProfile
    }
    $script:GdtBrowserProfile = New-GdtBrowserProfile -UserAgent $userAgent
    return $script:GdtBrowserProfile
}

function Reset-GdtBrowserProfile {
    $script:GdtBrowserProfile = $null
}

# Suy ra request profile từ URI để mọi call site đều qua header builder,
# kể cả endpoint chưa được gắn profile tường minh.
function Resolve-GdtRequestProfile {
    param([string]$Uri, [string]$RequestProfile = '')
    if (-not [string]::IsNullOrWhiteSpace($RequestProfile)) { return $RequestProfile }
    if ($Uri -match '/security-taxpayer/authenticate') { return 'Login' }
    if ($Uri -match '/captcha') { return 'Captcha' }
    if ($Uri -match '/invoices/export-xml') { return 'ExportXml' }
    if ($Uri -match '/invoices/(relative|related)') { return 'InvoiceRelation' }
    return 'InvoiceQuery'
}

# Chỉ ghi các header an toàn; Authorization, Cookie, proxy credential... không
# bao giờ được đưa vào log (denylist ngầm qua whitelist bên dưới).
function Write-GdtHttpProfileLog {
    param([string]$RequestProfile, $Headers, $Config)
    $enabled = [bool](Get-HddtConfigValue $Config 'LogHttpProfile' $false)
    if (-not $enabled) { return }
    Write-HddtLog INFO ('[HTTP] profile={0}' -f $RequestProfile)
    foreach ($name in @('User-Agent', 'Referer', 'Origin', 'Accept', 'Sec-Fetch-Site', 'Sec-Fetch-Mode', 'Sec-Fetch-Dest', 'sec-ch-ua', 'sec-ch-ua-platform')) {
        $property = $Headers[$name]
        if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property)) {
            Write-HddtLog INFO ('[HTTP] {0}={1}' -f $name, [string]$property)
        }
    }
}

# Log profile browser một lần khi khởi động (chỉ khi LOG_HTTP_PROFILE=true).
function Write-GdtBrowserProfileLog {
    param($Config)
    $enabled = [bool](Get-HddtConfigValue $Config 'LogHttpProfile' $false)
    if (-not $enabled) { return }
    $profile = Get-GdtBrowserProfile -Config $Config
    Write-HddtLog INFO ('[HTTP] Browser profile | UA={0}' -f $profile.UserAgent)
    Write-HddtLog INFO ('[HTTP] Browser profile | sec-ch-ua={0} | platform={1} | mobile={2} | locale={3}' -f $profile.SecChUa, $profile.SecChUaPlatform, $profile.SecChUaMobile, $profile.AcceptLanguage)
}

# Lắp ráp header theo thứ tự: BrowserProfile -> RequestProfile ->
# header riêng của endpoint (AdditionalHeaders) -> Authorization.
# AdditionalHeaders (Action, End-Point, Referer riêng...) luôn thắng profile
# nên header chức năng của API không bị mất.
function Get-GdtRequestHeaders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Captcha', 'Login', 'InvoiceQuery', 'InvoiceRelation', 'ExportXml')]
        [string]$RequestProfile,
        $Config,
        $BrowserProfile,
        [string]$AuthorizationToken = '',
        [hashtable]$AdditionalHeaders
    )

    if ($null -eq $BrowserProfile) { $BrowserProfile = Get-GdtBrowserProfile -Config $Config }

    $headers = @{
        'User-Agent' = [string]$BrowserProfile.UserAgent
        'Accept-Language' = [string]$BrowserProfile.AcceptLanguage
        'sec-ch-ua' = [string]$BrowserProfile.SecChUa
        'sec-ch-ua-mobile' = [string]$BrowserProfile.SecChUaMobile
        'sec-ch-ua-platform' = [string]$BrowserProfile.SecChUaPlatform
        'Sec-Fetch-Site' = 'same-origin'
        'Sec-Fetch-Mode' = 'cors'
        'Sec-Fetch-Dest' = 'empty'
    }

    switch ($RequestProfile) {
        'Captcha' {
            # GET CAPTCHA lấy ngay tại trang đăng nhập.
            $headers['Accept'] = 'application/json, text/plain, */*'
            $headers['Referer'] = [string]$BrowserProfile.RootReferer
        }
        'Login' {
            # POST JSON từ trang chủ: browser gửi Origin cho request CORS/POST.
            $headers['Accept'] = 'application/json, text/plain, */*'
            $headers['Referer'] = [string]$BrowserProfile.RootReferer
            $headers['Origin'] = [string]$BrowserProfile.Origin
        }
        'InvoiceQuery' {
            $headers['Accept'] = 'application/json, text/plain, */*'
            $headers['Referer'] = [string]$BrowserProfile.LookupReferer
        }
        'InvoiceRelation' {
            $headers['Accept'] = 'application/json, text/plain, */*'
            $headers['Referer'] = [string]$BrowserProfile.LookupReferer
        }
        'ExportXml' {
            # Tải ZIP/XML; Accept vẫn chứa */* nên không ép content negotiation.
            $headers['Accept'] = 'application/zip, application/xml, application/octet-stream, */*'
            $headers['Referer'] = [string]$BrowserProfile.LookupReferer
        }
    }

    if ($null -ne $AdditionalHeaders) {
        foreach ($headerName in $AdditionalHeaders.Keys) {
            $headers[[string]$headerName] = $AdditionalHeaders[$headerName]
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($AuthorizationToken)) {
        $headers['Authorization'] = 'Bearer ' + $AuthorizationToken
    }

    Write-GdtHttpProfileLog -RequestProfile $RequestProfile -Headers $headers -Config $Config
    return $headers
}
