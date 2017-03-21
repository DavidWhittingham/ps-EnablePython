Set-StrictMode -Version 3.0

$OLD_ENV_PATH = $null

function Disable-Python {
    <#
    .SYNOPSIS
    Disables any Python version enabled by the Enable-Python function.

    .DESCRIPTION
    The Disable-Python function restores the original PATH environment variable, removing any changes made by running the Enable-Python function.

    It will also check for the existance of a 'deactivate' function, and it will call it if it exists (assuming that it is from the virtualenv activate.ps1' script).
    #>

    [CmdletBinding()]

    param ()

    process {
        # test for a "deactivate" function, if it exists it probably means a virtualenv environment is in use
        # if so, call deactivate before restoring path
        if (
            ($deactivCmd = Get-Command "deactivate" -errorAction SilentlyContinue) -and
            ($deactivCmd.CommandType -eq "Function")
        ) {
            deactivate
        }

        if ($script:OLD_ENV_PATH) {
            # restore the original path
            $Env:PATH = $script:OLD_ENV_PATH
        }
    }
}

function Enable-Python {
    <#
    .SYNOPSIS
    Adds an installed version of Python to the current shell environment.

    .DESCRIPTION
    The Enable-Python function adds the install path and scripts path for the specified Python version to the PATH environment variable for this shell session.

    .EXAMPLE
    Enable-Python -Version 2.7 -Platform 32
    Enable Python v2.7 (x86-32).

    .EXAMPLE
    Enable-Python 3.4 64
    Enable Python v3.4 (x86-64) using short-hand syntax.

    .EXAMPLE
    Enable-Python 3.4
    Enable Python v3.4 using the highest available platform.
    #>

    [CmdletBinding()]

    param (
        [Parameter()]
        [ValidateScript({
            (Get-Python | ForEach-Object -process { $_.Version } | Sort-Object -Unique).Contains($_)
        })]
        # The version number of Python (in '<major>.<minor>' format) to enable.
        [string]$Version,

        [Parameter()]
        [ValidateSet(32,64)]
        # The Python platform version to enable (either 32 or 64).
        [int]$Platform
    )

    process {
        $installedPythons = Get-Python

        # From the Python versions installed, get the first version that matches the version number specified (and
        # optionally the platform).
        $foundVersion = $null
        foreach ($install in $installedPythons) {
            if ([string]::IsNullOrEmpty($Version) -or $install.Version -eq $Version) {
                if ($Platform) {
                    if ($Platform -eq $install.Platform) {
                        $foundVersion = $install
                        break
                    }
                }
                else
                {
                    $foundVersion = $install
                    break
                }
            }
        }

        if (!$foundVersion) {
            $pythonVersion = if ($Version) {
                    "Python $Version"
                } else {
                    "Python"
                }

            $errMessage = if ($Platform) {
                    "$pythonVersion (x86-$Platform) could not be found."
                } else {
                    "$pythonVersion could not be found."
                }

            throw $errMessage
        }

        # disable any existing Python version before enabling a new one
        Disable-Python

        # Save the existing path variable, then set the new path variable with the additional directories pre-pended.
        # Putting them at the start ensures the specified Python version will be the first one found (in case a Python
        # installation is already in the PATh variable).
        $script:OLD_ENV_PATH = $Env:PATH
        $Env:PATH = "$($foundVersion.InstallPath);$($foundVersion.ScriptsPath);$script:OLD_ENV_PATH"

        Write-Host "Python $($foundVersion.Version) (x86-$($foundVersion.Platform)) has been enabled."
    }
}

function Get-Python {
    [CmdletBinding()]
    [OutputType([PSObject])]
    param ()
    <#
    .SYNOPSIS
    Lists the Python versions installed on this computer.

    #>

    process {
        $versions = New-Object System.Collections.Generic.List[PSObject]

        [Microsoft.Win32.RegistryKey[]]$regKeys = Get-ChildItem -Path Registry::HKLM\Software\Python\PythonCore\ -ErrorAction "SilentlyContinue"
        [Microsoft.Win32.RegistryKey[]]$regKeys = $regKeys + (Get-ChildItem -Path Registry::HKLM\SOFTWARE\Wow6432Node\Python\PythonCore\ -ErrorAction "SilentlyContinue")

        foreach ($key in $regKeys) {
            if (Test-Path ("Registry::" + (Join-Path $key.Name "\InstallPath"))) {
                $versions.Add((createCPythonVersion $key))
            }
        }

        ($versions |
            Sort-Object -Property (
                @{Expression="Implementation"; Descending=$false},
                @{Expression="Version"; Descending=$true},
                @{Expression="Platform"; Descending=$true}
            )
        )
    }
}

function createCPythonVersion(
    [Microsoft.Win32.RegistryKey]$registryKey
) {
    # Create a new PythonVersion object from a CPython install registry key
    $newVersion = newPythonVersion
    $newVersion.Implementation = "CPython"
    $newVersion.InstallPath = (Get-ItemProperty -Path ("Registry::" + (Join-Path $key.Name "\InstallPath"))).'(Default)'
    $newVersion.Version = ($key.PSChildName)
    $newVersion.Platform = if ((is64Bit) -and !($registryKey -match "Wow6432Node")) {"64"} else {"32"}
    $newVersion
}

function is64Bit {
    # Check if this machine is 64-bit
    ((Get-WmiObject Win32_OperatingSystem).OSArchitecture -match '64-bit')
}

function newPythonVersion {
    # Create a new PSObject for storing Python version information
    $obj = New-Object -TypeName PSObject
    Add-Member -InputObject $obj -NotePropertyName InstallPath -NotePropertyValue $null
    Add-Member -InputObject $obj -NotePropertyName Implementation -NotePropertyValue $null
    Add-Member -InputObject $obj -NotePropertyName Platform -NotePropertyValue $null
    Add-Member -InputObject $obj -NotePropertyName Version -NotePropertyValue $null
    Add-Member -InputObject $obj -MemberType ScriptProperty -Name Name -Value {
        return ("{0} {1}, x86-{2}" -f $this.Implementation, $this.Version, $this.Platform)
    }
    Add-Member -InputObject $obj -MemberType ScriptProperty -Name ScriptsPath -Value {
        return (Join-Path $this.InstallPath "Scripts")
    }

    $defaultProperties = @("Name", "Version", "Platform")
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $obj | Add-Member MemberSet PSStandardMembers $PSStandardMembers

    $obj
}

Export-ModuleMember "*-*"