param(
    [string]$HostAddress = "0.0.0.0",
    [int]$Port = 8501,
    [string]$BrowserAddress = ""
)

function Get-LocalIPv4 {
    $addresses = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object {
            $_.IPAddress -notlike "169.254.*" -and
            $_.IPAddress -ne "127.0.0.1" -and
            $_.PrefixOrigin -ne "WellKnown"
        } |
        Sort-Object InterfaceMetric

    if ($addresses) {
        return $addresses[0].IPAddress
    }

    return "127.0.0.1"
}

if (-not $BrowserAddress) {
    if ($HostAddress -eq "127.0.0.1" -or $HostAddress -eq "localhost") {
        $BrowserAddress = "127.0.0.1"
    }
    elseif ($HostAddress -eq "0.0.0.0") {
        $BrowserAddress = Get-LocalIPv4
    }
    else {
        $BrowserAddress = $HostAddress
    }
}

$env:STREAMLIT_SERVER_ADDRESS = $HostAddress
$env:STREAMLIT_SERVER_PORT = "$Port"
$env:STREAMLIT_BROWSER_SERVER_ADDRESS = $BrowserAddress
$env:STREAMLIT_BROWSER_SERVER_PORT = "$Port"

Write-Host "Local URL: http://127.0.0.1:$Port" -ForegroundColor Cyan

if ($HostAddress -eq "0.0.0.0" -or $HostAddress -eq $BrowserAddress) {
    Write-Host "LAN URL:   http://$BrowserAddress:$Port" -ForegroundColor Cyan
}

Write-Host "Binding Streamlit to $HostAddress:$Port" -ForegroundColor Cyan

streamlit run app.py --server.address $HostAddress --server.port $Port
