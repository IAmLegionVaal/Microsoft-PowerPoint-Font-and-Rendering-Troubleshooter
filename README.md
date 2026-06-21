# Microsoft PowerPoint Font and Rendering Troubleshooter

Created by **Dewald Pretorius**.

The repository includes the original font, substitution, display-scaling, rendering, and layout diagnostics plus a guarded `Repair.ps1` helper.

```powershell
.\Repair.ps1 -Action Diagnose
.\Repair.ps1 -Action ResetOfficeCache -WhatIf
.\Repair.ps1 -Action ResetOfficeCache -Confirm
```

PowerPoint must be closed before repair. Existing Office cache data is preserved as a timestamped backup, and the diagnostic snapshot records installed-font count and application state. The helper does not remove fonts or edit presentations. Source-reviewed for Windows PowerShell 5.1; not runtime-tested against every display or Office configuration.
