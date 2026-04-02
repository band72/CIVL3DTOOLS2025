$dir = "c:\Users\Daryl Banks\source\repos\C3DKP26_FIXED\C3dProjects25"
$files = Get-ChildItem -Path $dir -Filter "*.cs" -Recurse

$classesWithCommands = @{}

foreach ($f in $files) {
    if ($f.Name -match "obj|bin") { continue }
    $content = Get-Content $f.FullName -Raw
    
    # We only care if it has CommandMethod
    if ($content -match "\[CommandMethod") {
        # parse namespace
        $ns = ""
        if ($content -match "(?m)^namespace\s+([\w\.]+)") {
             $ns = $matches[1]
        }
        
        $lines = $content -split '\r?\n'
        for ($i = 0; $i -lt $lines.Length; $i++) {
            if ($lines[$i] -match "class\s+(\w+)") {
                $cls = $matches[1]
                # look ahead for a CommandMethod in this class scope
                $hasCmd = $false
                for ($j = $i + 1; $j -lt $lines.Length; $j++) {
                    if ($lines[$j] -match "class\s+\w+") { break }
                    if ($lines[$j] -match "\[CommandMethod") { $hasCmd = $true; break }
                }
                
                if ($hasCmd) {
                    if ($ns) { 
                        $classesWithCommands["${ns}.${cls}"] = 1
                    } else { 
                        $classesWithCommands[$cls] = 1 
                    }
                }
            }
        }
    }
}

$output = @()
foreach ($key in $classesWithCommands.Keys) {
    if ($key) {
        $output += "[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof($key))]"
    }
}

# Update Commands.cs
$commandsFile = "c:\Users\Daryl Banks\source\repos\C3DKP26_FIXED\C3dProjects25\Commands.cs"
$cmdText = Get-Content $commandsFile -Raw
$joinedOutput = $output -join "`r`n"
$cmdText = $cmdText -replace '(?s)namespace RCS\.C3D2025\.Tools', "`r`n$joinedOutput`r`n`r`nnamespace RCS.C3D2025.Tools"
Set-Content -Path $commandsFile -Value $cmdText

Write-Host "Registered $($classesWithCommands.Keys.Count) command classes."
