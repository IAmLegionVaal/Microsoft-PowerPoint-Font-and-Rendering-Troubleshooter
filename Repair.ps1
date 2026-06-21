#requires -Version 5.1
<# Created by Dewald Pretorius. #>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [ValidateSet('Diagnose','ResetOfficeCache')][string]$Action='Diagnose',
    [string]$OutputPath=(Join-Path ([Environment]::GetFolderPath('Desktop')) 'PowerPoint_Rendering_Repair')
)
$ErrorActionPreference='Stop'
$CachePath="$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
New-Item -ItemType Directory -Path $OutputPath -Force|Out-Null
$Stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath=Join-Path $OutputPath "Repair_$Stamp.log"
function Write-RepairLog([string]$Message){$Line='{0:u} {1}' -f (Get-Date),$Message;Write-Host $Line;Add-Content -LiteralPath $LogPath -Value $Line}
$State=[ordered]@{
    Action=$Action
    PowerPointRunning=[bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
    CacheExists=(Test-Path -LiteralPath $CachePath)
    FontCount=@(Get-ChildItem "$env:WINDIR\Fonts" -ErrorAction SilentlyContinue).Count
}
$State|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $OutputPath "PreRepair_$Stamp.json") -Encoding UTF8
if($Action -eq 'Diagnose'){Write-RepairLog '[COMPLETE] Read-only snapshot saved.';exit 0}
try{
    if($PSCmdlet.ShouldProcess($CachePath,'Preserve and reset Office cache')){
        if(Get-Process POWERPNT -ErrorAction SilentlyContinue){throw 'Close PowerPoint before resetting the cache.'}
        if(Test-Path -LiteralPath $CachePath){
            $Backup="$CachePath.backup-$Stamp"
            Move-Item -LiteralPath $CachePath -Destination $Backup -Force
            New-Item -ItemType Directory -Path $CachePath -Force|Out-Null
            Write-RepairLog "[BACKUP] $Backup"
        }
    }
}catch{Write-RepairLog "[FAILED] $($_.Exception.Message)";exit 5}
Write-RepairLog '[COMPLETE] Repair completed.'
exit 0
