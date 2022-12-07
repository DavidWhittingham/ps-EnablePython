Set-StrictMode -Version 2.0

$ENV_VAR_BACKUP = @{}
$PROMPT_FUNCTION = $null

function Disable-Python {
    <#
.SYNOPSIS
Disables any Python version enabled by the Enable-Python function.

.DESCRIPTION
The Disable-Python function restores the original PATH environment variable, removing any changes made by running the
Enable-Python function.

It will also check for the existance of a 'deactivate' function, and it will call it if it exists (assuming that it is
from the virtualenv activate.ps1' script).
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

        disableConda
        restoreEnvVars
    }
}

function Enable-Python {
    <#
.SYNOPSIS
Adds an installed version of Python to the current shell environment.

.DESCRIPTION
The Enable-Python function adds the install path and scripts path for the specified Python version to the PATH
environment variable for this shell session.  If multiple Python distributions are found, they are sorted in the
following order:
    - Scope (current user installs sorted first)
    - Company (alphabetical)
    - Version
    - Platform (descending)
    - Tag
The core Python distribution ("PythonCore") is treated as a special case and always sorted first.  If multiple
versions are found, the top-sorted version is enabled.

.PARAMETER Company
The name of a company to filter Python distributions on.

.PARAMETER Tag
The tag value to filter Python distributions on.

.PARAMETER Version
The version number to filter Python distributions on.

.PARAMETER Platform
The underlying CPU architecture (either "32"-bit or "64"-bit) to filter Python distributions on.

.PARAMETER Scope
The operating system install scope (either "CurrentUser" or "AllUsers") to filter Python distributions on.

.PARAMETER PythonHome
Sets a custom path on the PYTHONHOME environment variable.

.PARAMETER NoPlatformUserBase
Instructs EnablePython not to set a custom PYTHONUSERBASE path in a platform-specific manner.  By default, Python uses
the same location for both 32-bit and 64-bit installations, which causes problems for binary components.

.EXAMPLE
Enable-Python
Gets all Python installs available, sorts them, and enables the top-most install (CurrentUser scope installs before
AllUser, PythonCore above others then alphabetical company name, highest version number, 64-bit before 32-bit).

.EXAMPLE
Enable-Python -Company ContinuumAnalytics
Gets all Python installs from Continuum Analytics, Inc., sorts them, and enables the top-most install (CurrentUser
scope installs before AllUser, highest version number, 64-bit before 32-bit).

.EXAMPLE
Enable-Python -Version 2.7 -Platform 32
Filters for 32-bit Python distributions with a version of 2.7.*, sorts them, and enables the top-most install
(CurrentUser scope installs before AllUser, PythonCore above others then alphabetical company name, highest version
number, 64-bit before 32-bit).

.EXAMPLE
Enable-Python 3.4 64
The Version and Platform parameters are positional, and can be specified without their respective names. Filters for
64-bit Python distributions with a version of 3.4.*, sorts them, and enables the top-most install (CurrentUser scope
installs before AllUser, PythonCore above others then alphabetical company name, highest version number, 64-bit before
32-bit).

.LINK
https://github.com/DavidWhittingham/ps-EnablePython

#>

    [CmdletBinding()]

    param (
        [Parameter()]
        [string]$Company,

        [Parameter()]
        [string]$Tag,

        [Parameter(Position = 1)]
        [string]$Version,

        [Parameter(Position = 2)]
        [ValidateSet(32, 64)]
        [Nullable[int]]$Platform,

        [Parameter()]
        [ValidateSet("CurrentUser", "AllUsers")]
        [string]$Scope,

        [Parameter()]
        [Alias("Home")]
        # Implemented validation, but leaving it disabled for now, not sure if there are any situations where
        # Test-Path might fail on a valid path
        # [ValidateScript({
        #     if ($_ -eq $null) {
        #         $true
        #     } elseif (Test-Path $_ -PathType Container) {
        #         $true
        #     } else {
        #         Throw [System.Management.Automation.ValidationMetadataException] "The path '${_}' is not a valid directory."
        #     }
        # })]
        [string]$PythonHome,

        [Parameter()]
        [switch]$NoPlatformUserBase
    )

    process {
        $getPythonArgs = @{
            "Company" = $Company;
            "Tag"     = $Tag;
            "Version" = $Version;
        }
        if ($null -ne $Platform) {
            $getPythonArgs.Platform = $Platform
        }
        if (![string]::IsNullOrWhiteSpace($Scope)) {
            $getPythonArgs.Scope = $Scope
        }

        $pythons = Get-Python @getPythonArgs

        if (!$pythons) {
            throw "No Python distribution could be found matching those search criteria."
        }

        if ($pythons.Length -gt 1) {
            Write-Information "Multiple Python distributions found matching that criteria, enabling the top choice..." -InformationAction Continue
        }

        # disable any existing Python version before enabling a new one
        Disable-Python

        $foundVersion = $pythons[0]

        # Save the existing path variable, then set the new path variable with the additional directories pre-pended.
        # Putting them at the start ensures the specified Python version will be the first one found (in case a Python
        # installation is already in the PATh variable).
        $script:ENV_VAR_BACKUP["PATH"] = $Env:PATH
        $Env:PATH = "$($foundVersion.InstallPath);$($foundVersion.ScriptsPath);$Env:PATH"

        # Save the existing PYTHONHOME variable, then ensure it is cleared so that the activated Python doesn't go off
        # looking at an incorrect home
        $script:ENV_VAR_BACKUP["PYTHONHOME"] = $Env:PYTHONHOME
        $Env:PYTHONHOME = $PythonHome

        # If configured to separate user base by platform, set a custom user base
        if ($NoPlatformUserBase -eq $false) {
            $script:ENV_VAR_BACKUP["PYTHONUSERBASE"] = $Env:PYTHONUSERBASE
            $Env:PYTHONUSERBASE = Join-Path -Path (Join-Path -Path $Env:APPDATA -ChildPath "EnablePython") `
                -ChildPath ("x86-{0}" -f $foundVersion.Platform)
        }

        # Get the user scripts path, add it to PATH as well
        $userScriptsPath = & $foundVersion.Executable -E -c 'import sysconfig; print(sysconfig.get_path(""scripts"", scheme=""nt_user""))'
        $Env:PATH = "$userScriptsPath;$Env:Path"

        # attempt to enable conda
        enableConda $foundVersion #-ErrorAction SilentlyContinue

        Write-Information """$($foundVersion.Name)"" has been enabled." -InformationAction Continue
    }
}

function Get-Python {
    <#
.SYNOPSIS
Gets the installed versions of Python on the current machine.

.DESCRIPTION
The Get-Python function finds Python installs on the current machine, collects their details and returns them as a
sorted list. Installs are sorted in the following order:
    - Scope (current user installs sorted first)
    - Company (alphabetical)
    - Version
    - Platform (descending)
    - Tag
The core Python distribution ("PythonCore") is treated as a special case and always sorted first.

.PARAMETER Company
The name of a company to filter Python distributions on.

.PARAMETER Tag
The tag value to filter Python distributions on.

.PARAMETER Version
The version number to filter Python distributions on.

.PARAMETER Platform
The underlying CPU architecture (either "32"-bit or "64"-bit) to filter Python distributions on.

.PARAMETER Scope
The operating system install scope (either "CurrentUser" or "AllUsers") to filter Python distributions on.

.EXAMPLE
Get-Python
Gets all Python installs available and sorts them.

.EXAMPLE
Get-Python -Company ContinuumAnalytics
Gets all Python installs from Continuum Analytics, Inc. and sorts them.

.EXAMPLE
Get-Python -Version 2.7 -Platform 32
Gets all 32-bit Python installs with a version of 2.7.* and sorts them.

.EXAMPLE
Get-Python 3.4 64
Gets all 364-bit Python installs with a version of 3.4.* and sorts them. The Version and Platform parameters are
positional, and can be specified without their respective names.

.LINK
https://github.com/DavidWhittingham/ps-EnablePython

#>

    [CmdletBinding()]
    [OutputType([System.Array])]

    param (
        [Parameter()]
        [string]$Company,

        [Parameter()]
        [string]$Tag,

        [Parameter(Position = 1)]
        [string]$Version,

        [Parameter(Position = 2)]
        [ValidateSet(32, 64)]
        [Nullable[int]]$Platform,

        [Parameter()]
        [ValidateSet("CurrentUser", "AllUsers")]
        [string]$Scope
    )

    <#
    .SYNOPSIS
    Lists the Python versions installed on this computer.

    #>

    process {
        $pythons = getPep514Pythons
        $arcGisProPython = getArcGisProPython

        if ($null -ne $arcGisProPython) {
            $pythons.Add($arcGisProPython) | Out-Null
        }

        # Sort the list of Python versions
        $sortProperties = (
            @{
                Expression = " Scope";
                Descending = $true
            },
            @{
                Expression = "Company";
                Descending = $false
            },
            @{
                Expression = "Version";
                Descending = $true
            },
            @{
                Expression = "Platform";
                Descending = $true
            },
            @{
                Expression = "Tag";
                Descending = $false
            }
        )

        # Apply filters
        if (![string]::IsNullOrWhiteSpace($Scope)) {
            $pythons = ($pythons | Where-Object { "$($_.Scope)" -like "$Scope" })
        }

        if (![string]::IsNullOrWhiteSpace($Company)) {
            $pythons = ($pythons | Where-Object { "$($_.Company)" -like "$Company" + "*" })
        }

        if (![string]::IsNullOrWhiteSpace($Tag)) {
            $pythons = ($pythons | Where-Object { "$($_.Tag)" -like "$Tag" + "*" })
        }

        if (![string]::IsNullOrWhiteSpace($Version)) {
            $pythons = ($pythons | Where-Object { "$($_.Version)" -like "$Version*" })
        }

        if ($null -ne $Platform) {
            $pythons = ($pythons | Where-Object { "$($_.Platform)" -like "$Platform" })
        }

        # Sort PythonCore first
        Write-Output -NoEnumerate @(
            @($pythons | Where-Object { $_.Company -eq "PythonCore" } | Sort-Object -Property $sortProperties) +
            @($pythons | Where-Object { $_.Company -ne "PythonCore" } | Sort-Object -Property $sortProperties)
        )
    }
}

function createPep514PythonVersion([Microsoft.Win32.RegistryKey]$tagKey, [string]$scope) {
    $parentKey = Get-Item $tagKey.PSParentPath
    $installPathKey = Get-Item (Join-Path $tagKey.PSPath "InstallPath")
    $pythonExecutable = getPythonCommand $installPathKey
    $version = getPythonVersion($pythonExecutable)
    $company = $parentKey.PSChildName
    $companyDisplayName = if ((Get-ItemProperty -Path $parentKey.PSPath) -match "DisplayName") {
        (Get-ItemProperty -Path $parentKey.PSPath)."DisplayName"
    }
    else {
        $null
    }
    $tagDisplayName = if ((Get-ItemProperty -Path $tagKey.PSPath) -match "DisplayName") {
        (Get-ItemProperty -Path $tagKey.PSPath)."DisplayName"
    }
    else {
        $null
    }

    # Create a new PSObject for storing Python version information
    $obj = New-Object -TypeName PSObject -Prop (@{
            "Company"            = $company;
            "CompanyDisplayName" = $companyDisplayName;
            "Tag"                = $tagKey.PSChildName;
            "TagDisplayName"     = $tagDisplayName;
            "InstallPath"        = (Get-ItemProperty -Path $installPathKey.PSPath)."(Default)";
            "Platform"           = if ((is64Bit) -and !($tagKey -match "Wow6432Node")) { "64" } else { "32" };
            "Executable"         = $pythonExecutable;
            "Version"            = $version;
            "Scope"              = $scope;
        })

    Add-Member -InputObject $obj -MemberType ScriptProperty -Name "Name" -Value {
        $company = if ($null -eq $this.CompanyDisplayName) { $this.Company } else { $this.CompanyDisplayName }
        $tag = if ($null -eq $this.TagDisplayName) { $this.Tag } else { $this.TagDisplayName }
        return ("{0} {1} (x86-{2})" -f $company, $tag, $this.Platform)
    }
    Add-Member -InputObject $obj -MemberType ScriptProperty -Name "ScriptsPath" -Value {
        return (Join-Path $this.InstallPath "Scripts")
    }

    $obj = setPythonVersionStandardMembers $obj

    $obj
}

function disableConda() {
    if (-not(Get-Module -Name "Conda")) {
        # do nothing if no Conda module loaded
        return
    }

    # Remove the Conda module that has been added
    Remove-Module -Name "Conda"

    # Restore the backup prompt, with a basic prompt fallback
    if ($null -ne $script:PROMPT_FUNCTION) {
        $promptFunction = $script:PROMPT_FUNCTION
    }
    else {
        $promptFunction = {
            "PS $($executionContext.SessionState.Path.CurrentLocation)$('>' * ($nestedPromptLevel + 1)) ";
        }
    }
    New-Item -Path function: -Name "global:prompt" -Value $promptFunction -Force > $null
    $script:PROMPT_FUNCTION = $null
}

function enableConda($pythonVersion) {
    $condaPath = Join-Path $pythonVersion.InstallPath "Scripts\conda.exe"

    if (-not(Test-Path $condaPath)) {
        return
    }

    # test is the Conda PS module exists, if so, try to activate Conda
    $condaPsModPath = Join-Path "$($pythonVersion.InstallPath)" "\shell\condabin\Conda.psm1"
    if (-not(Test-Path $condaPsModPath)) {
        return
    }

    # Conda messes with the PS prompt, backup it up so that we can restore it later
    $script:PROMPT_FUNCTION = Get-Item Function:\prompt

    # backup the current state of conda-related environment variables
    $script:ENV_VAR_BACKUP["CONDA_EXE"] = $Env:CONDA_EXE
    $script:ENV_VAR_BACKUP["_CE_M"] = $Env:_CE_M
    $script:ENV_VAR_BACKUP["_CE_CONDA"] = $Env:_CE_CONDA
    $script:ENV_VAR_BACKUP["_CONDA_ROOT"] = $Env:_CONDA_ROOT
    $script:ENV_VAR_BACKUP["_CONDA_EXE"] = $Env:_CONDA_EXE
    $script:ENV_VAR_BACKUP["CONDA_DEFAULT_ENV"] = $Env:CONDA_DEFAULT_ENV
    $script:ENV_VAR_BACKUP["CONDA_PREFIX"] = $Env:CONDA_PREFIX
    $script:ENV_VAR_BACKUP["CONDA_PROMPT_MODIFIER"] = $Env:CONDA_PROMPT_MODIFIER
    $script:ENV_VAR_BACKUP["CONDA_PYTHON_EXE"] = $Env:CONDA_PYTHON_EXE
    $script:ENV_VAR_BACKUP["CONDA_SHLVL"] = $Env:CONDA_SHLVL

    # set environment variables for conda activation
    $Env:CONDA_EXE = "$condaPath"
    $Env:_CONDA_ROOT = "$($pythonVersion.InstallPath)"
    $Env:_CONDA_EXE = "$condaPath"

    Import-Module $condaPsModPath -Global
    conda activate

    Add-CondaEnvironmentToPrompt
}

function getArcGisProPython() {
    $arcgisProRegKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\ESRI\ArcGISPro"

    if (-not(Test-Path $arcgisProRegKey)) {
        return $null
    }

    $condaEnv = Get-ItemProperty -Path $arcgisProRegKey -Name "PythonCondaEnv" -ErrorAction "SilentlyContinue"
    $pythonRoot = Get-ItemProperty -Path $arcgisProRegKey -Name "PythonCondaRoot" -ErrorAction "SilentlyContinue"

    if ((-not $condaEnv) -or (-not $pythonRoot)) {
        return $null
    }

    $pythonInstallDir = Join-Path -Path (Join-Path -Path $pythonRoot.PythonCondaRoot -ChildPath "envs") -ChildPath $condaEnv.PythonCondaEnv
    $pythonExecutable = Join-Path -path $pythonInstallDir -ChildPath "python.exe"

    if (-not(Test-Path $pythonExecutable)) {
        return $null
    }

    # format is "Platform|Tag|Version"
    $pythonInfoCommand = 'import platform; import sys; import sysconfig; print("{}|{}|{}".format(sysconfig.get_platform(), sys.winver, platform.python_version()))'
    $pythonInfo = (& $pythonExecutable -E -c $pythonInfoCommand).Split("|")

    $company = "Esri"
    $companyDisplayName = "Esri Inc."
    $platform = if ($pythonInfo[0] -eq "win-amd64") { "64" } elseif ($pythonInfo[0] -eq "win32") { "32" } else { $null }
    $tag = $pythonInfo[1]
    $tagDisplayName = "Python {0} ({1}-bit)" -f $tag, $platform
    $version = $pythonInfo[2]
    $name = "{0} {1}" -f $companyDisplayName, $tagDisplayName
    $scriptsPath = Join-Path $pythonInstallDir "Scripts"

    $arcgisProPython = New-Object -TypeName PSObject -Prop (@{
            "Company"            = $company;
            "CompanyDisplayName" = $companyDisplayName;
            "InstallPath"        = $pythonInstallDir;
            "Executable"         = $pythonExecutable;
            "Tag"                = $tag;
            "TagDisplayName"     = $tagDisplayName;
            "Platform"           = $platform;
            "Version"            = $version;
            "Scope"              = "AllUsers";
            "Name"               = $name;
            "ScriptsPath"        = $scriptsPath;
        })

    return setPythonVersionStandardMembers $arcgisProPython
}

function getPep514Pythons() {
    [OutputType([System.Collections.Generic.List[PSObject]])]

    $pythons = New-Object System.Collections.Generic.List[PSObject]

    $regKeyLocations = (
        @{
            path  = "HKCU\Software\Python\"
            scope = "CurrentUser"
        },
        @{
            path  = "HKLM\Software\Python\"
            scope = "AllUsers"
        }
    )

    if (is64Bit) {
        $regKeyLocations += @{
            path  = "HKLM\SOFTWARE\Wow6432Node\Python\"
            scope = "AllUsers"
        }
    }

    foreach ($location in $regKeyLocations) {
        $companyKeys = Get-ChildItem -Path "Registry::$($location.path)" -ErrorAction "SilentlyContinue"

        foreach ($companyKey in $companyKeys) {

            if ($companyKey.PSChildName -eq "PyLauncher") {
                # PyLauncher should be ignored
                continue
            }

            $tagKeys = Get-ChildItem -Path $companyKey.PSPath -ErrorAction "SilentlyContinue"

            foreach ($tagKey in $tagKeys) {
                if (Test-Path (Join-Path $tagKey.PSPath "\InstallPath")) {
                    # tests if a valid install actually exists, uninstalled version can leave the Company/Tag structure
                    $pythons.Add((createPep514PythonVersion $tagKey $location.scope)) | Out-Null
                }
            }
        }
    }

    Write-Output -NoEnumerate $pythons
}

function getPythonCommand([Microsoft.Win32.RegistryKey]$installPathKey) {
    # test for "ExecutablePath" folder
    if ((Get-ItemProperty -Path $installPathKey.PSPath) -match "ExecutablePath") {
        Get-Command (Get-ItemProperty -Path $installPathKey.PSPath)."ExecutablePath"
    }
    else {
        $pythonPath = Join-Path (Get-ItemProperty -Path $installPathKey.PSPath)."(Default)" "python.exe"
        if (Test-Path $pythonPath) {
            Get-Command $pythonPath
        }
        else {
            $null
        }
    }
}

function getPythonVersion([System.Management.Automation.ApplicationInfo]$executablePath) {
    if ($null -eq $executablePath) {
        $null
    }
    else {
        & $executablePath -E -c 'import platform; print(platform.python_version())'
    }
}

function is64Bit {
    # Check if this machine is 64-bit
    ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -match '64-bit')
}

function restoreEnvVars {
    foreach ($nvp in $script:ENV_VAR_BACKUP.GetEnumerator()) {
        $varName = $nvp.Name
        Set-Item "env:$varName" $nvp.Value
    }

    $script:ENV_VAR_BACKUP = @{}
}

function setPythonVersionStandardMembers($pythonVersion) {
    $defaultProperties = @("Name", "Company", "Tag", "Version", "Platform", "Scope")
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $pythonVersion | Add-Member MemberSet PSStandardMembers $PSStandardMembers

    $pythonVersion
}

Export-ModuleMember "*-*"