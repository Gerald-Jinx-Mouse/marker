# Make All Conda Environments Work in PowerShell 7

This guide fixes the issue where `conda env list`, `conda activate`, and other conda commands fail in PowerShell 7.5+ with errors like:

```
conda-script.py: error: argument COMMAND: invalid choice: ''
```

---

## What You Have

Three environments that need to work in PowerShell 7:

- **base** — `C:\ProgramData\anaconda3`
- **dev** — `C:\Users\jacob\.conda\envs\dev`
- **cuda** — `C:\Users\jacob\.ai-navigator\micromamba\envs\cuda`

All are activated via `conda activate`. Once conda works in PowerShell 7, all three will work.

---

## Approach: PowerShell Profile Workaround

Because base uses Python 3.13 and conda 25.1.1 does not support it, the most reliable fix is the **profile workaround**. It clears the problematic env vars before conda runs and works with any conda version.

---

### Step 1: Locate the PowerShell 7 Profile

The profile file is at:

```
C:\Users\jacob\Documents\PowerShell\Microsoft.PowerShell_profile.ps1
```

If it does not exist, create the folder and file:

```powershell
New-Item -ItemType Directory -Path "$env:USERPROFILE\Documents\PowerShell" -Force
New-Item -ItemType File -Path "$env:USERPROFILE\Documents\PowerShell\Microsoft.PowerShell_profile.ps1" -Force
notepad "$env:USERPROFILE\Documents\PowerShell\Microsoft.PowerShell_profile.ps1"
```

---

### Step 2: Add the Fix

Add one of the following at the **end** of the profile (after any existing conda init block).

**Option A — Simple (runs on every profile load):**

```powershell
$Env:_CE_M = $Env:_CE_CONDA = $null
```

**Option B — Conditional (only when invoking conda, PowerShell 7.5+):**

```powershell
if ($PSVersionTable.PSVersion -ge [version]"7.5.0") {
    $ExecutionContext.InvokeCommand.PreCommandLookupAction = {
        param ($CommandName, $CommandOrigin)
        if ($CommandName -eq "conda") {
            $Env:_CE_M = $null
            $Env:_CE_CONDA = $null
        }
    }
}
```

---

### Step 3: Ensure Conda Is Initialized for PowerShell 7

If conda is not yet initialized for PowerShell 7, add this block **before** the fix above:

```powershell
#region conda initialize
# !! Contents within this block are managed by 'conda init' !!
If (Test-Path "C:\ProgramData\anaconda3\Scripts\conda.exe") {
    (& "C:\ProgramData\anaconda3\Scripts\conda.exe" "shell.powershell" "hook") | Out-String | ?{$_} | Invoke-Expression
}
#endregion
```

---

### Step 4: Reload and Verify

1. Reload the profile:

   ```powershell
   . $PROFILE
   ```

2. Test:

   ```powershell
   conda env list
   conda activate dev
   conda activate C:\Users\jacob\.ai-navigator\micromamba\envs\cuda
   ```

3. Restart PowerShell 7 if needed.

---

## Alternative: Upgrade Conda to 25.11

If you prefer a conda upgrade instead of the profile workaround, run this in **Windows PowerShell 5.1** (where conda works):

```powershell
conda activate base
conda install -n base -c conda-forge conda=25.11
```

Conda 25.11 supports Python 3.13 and includes the PowerShell 7.5 fix. After upgrading, you can remove the profile workaround if you added it.

---

## Summary

| Action                    | Result                                                                 |
| ------------------------- | ---------------------------------------------------------------------- |
| Add workaround to profile | `conda env list`, `conda activate base`, `conda activate dev`, and activating the cuda env all work in PowerShell 7 |
| Upgrade to conda 25.11    | Same result, no profile change needed                                  |

The profile workaround is the safest option given a Python 3.13 base environment.
