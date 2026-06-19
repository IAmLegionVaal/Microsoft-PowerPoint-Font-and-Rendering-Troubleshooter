#requires -Version 5.1
<# Created by Dewald Pretorius #>
param([string]$OutputPath)
if(-not $OutputPath){$OutputPath="$([Environment]::GetFolderPath('Desktop'))\PowerPoint_Font_Rendering_Reports"};New-Item $OutputPath -ItemType Directory -Force|Out-Null
$fonts=Get-ChildItem "$env:WINDIR\Fonts" -File -ErrorAction SilentlyContinue|Select-Object Name,Length,LastWriteTime;$display=Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue|Select-Object Name,DriverVersion,AdapterRAM,Status
@('POWERPOINT FONT AND RENDERING TROUBLESHOOTER','Created by Dewald Pretorius',"Generated: $(Get-Date)",'Display adapters:',($display|Format-Table -AutoSize|Out-String -Width 220),'Installed fonts:',($fonts|Format-Table -AutoSize|Out-String -Width 220),'Guidance: verify font availability and embedding, compare display scaling, test hardware acceleration, export to PDF, and check theme font substitution.')|Set-Content (Join-Path $OutputPath 'Report.txt') -Encoding UTF8