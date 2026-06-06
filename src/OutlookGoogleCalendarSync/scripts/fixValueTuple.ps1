param (
    [string]$configPath
)

if (-not (Test-Path $configPath)) {
    [Console]::WriteLine("=== [ValueTuple Fix] Warning: Config file not found at $configPath ===")
    exit 0
}

try {
    [xml]$xml = Get-Content $configPath
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('asm', 'urn:schemas-microsoft-com:asm.v1')
    
    $nodes = $xml.SelectNodes('//asm:dependentAssembly[asm:assemblyIdentity[@name="System.ValueTuple"]]', $ns)
    
    if ($nodes.Count -eq 0) {
        [Console]::WriteLine("=== [ValueTuple Fix] Warning: No System.ValueTuple nodes found to patch. ===")
    }

    foreach ($node in $nodes) {
        $redirect = $node.SelectSingleNode('asm:bindingRedirect', $ns)
        if ($redirect) {
            $oldVer = $redirect.GetAttribute('newVersion')
            $redirect.SetAttribute('newVersion', '4.0.5.0')
            [Console]::WriteLine("=== [ValueTuple Fix] Success: Changed newVersion from $oldVer to 4.0.5.0. ===")
        }
    }
    
    $xml.Save($configPath)

} catch {
    [Console]::WriteLine("=== [ValueTuple Fix] ERROR: Failed to patch config file. Details: $_ ===")
}