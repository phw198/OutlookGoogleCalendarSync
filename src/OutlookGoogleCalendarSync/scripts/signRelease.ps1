[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [String]$ReleaseDirectory,
    [String]$SignToolPath = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\signtool.exe"
)

if (!(Test-Path -LiteralPath $SignToolPath)) {
    throw "SignTool was not found at '$SignToolPath'."
}

$dllFiles = Get-ChildItem -LiteralPath $ReleaseDirectory -Recurse -File -Filter "*.dll"
$unsignedFiles = $dllFiles |
    Where-Object { (Get-AuthenticodeSignature -LiteralPath $_.FullName).Status -eq "NotSigned" }

$signingArguments = @(
    "sign",
    "/as",
    "/n", "Paul Woolcock",
    "/tr", "http://time.certum.pl/",
    "/td", "sha256",
    "/fd", "sha256",
    "/v"
)
$signingArguments += @($unsignedFiles | ForEach-Object { $_.FullName })

if ($unsignedFiles.Count -gt 0) {
    Write-Host "Signing $($unsignedFiles.Count) DLL(s) in one SignTool invocation"
    & $SignToolPath @signingArguments
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to sign the release DLLs."
    }
}

$invalidFiles = $dllFiles |
    Where-Object { (Get-AuthenticodeSignature -LiteralPath $_.FullName).Status -ne "Valid" }
if ($invalidFiles.Count -gt 0) {
    throw "The following DLL(s) do not have a valid Authenticode signature: $($invalidFiles.FullName -join ", ")"
}

Write-Host "Verified $($dllFiles.Count) release DLL(s); signed $($unsignedFiles.Count) previously unsigned DLL(s)."