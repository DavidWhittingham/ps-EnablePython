Set-StrictMode -Version 2.0

$OLD_ENV_PATH = $null
$OLD_ENV_PYTHONHOME = $null
$OLD_ENV_PYTHONUSERBASE = $null
$RESTORE_ENV_VARS = $false
$RESTORE_USER_BASE = $false

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

        if ($script:RESTORE_ENV_VARS -eq $true) {
            # restore the original path
            $Env:PATH = $script:OLD_ENV_PATH

            # restore the original PYTHONHOME
            $Env:PYTHONHOME = $script:OLD_ENV_PYTHONHOME

            if ($script:RESTORE_USER_BASE -eq $true) {
                # restore the original PYTHONUSERBASE
                $Env:PYTHONUSERBASE = $script:OLD_ENV_PYTHONUSERBASE
                $script:RESTORE_USER_BASE = $false
            }

            $script:RESTORE_ENV_VARS = $false
        }
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

        [array]$pythons = Get-Python @getPythonArgs

        if (!$pythons) {
            throw "No Python distribution could be found matching those search criteria."
        }

        if ($pythons.Length -gt 1) {
            Write-Information "Multiple Python distributions found matching that criteria, enabling the top choice..." -InformationAction Continue
        }

        # disable any existing Python version before enabling a new one
        Disable-Python

        $foundVersion = $pythons[0]

        # Let EnablePython know it needs to restore the original environment variables on disabling Python
        $script:RESTORE_ENV_VARS = $true

        # Save the existing path variable, then set the new path variable with the additional directories pre-pended.
        # Putting them at the start ensures the specified Python version will be the first one found (in case a Python
        # installation is already in the PATh variable).
        $script:OLD_ENV_PATH = $Env:PATH
        $Env:PATH = "$($foundVersion.InstallPath);$($foundVersion.ScriptsPath);$script:OLD_ENV_PATH"

        # Save the existing PYTHONHOME variable, then ensure it is cleared so that the activated Python doesn't go off
        # looking at an incorrect home
        $script:OLD_ENV_PYTHONHOME = $Env:PYTHONHOME
        $Env:PYTHONHOME = $PythonHome

        # If configured to separate user base by platform, set a custom user base
        if ($NoPlatformUserBase -eq $false) {
            $script:OLD_ENV_PYTHONUSERBASE = $Env:PYTHONUSERBASE
            $Env:PYTHONUSERBASE = Join-Path -Path (Join-Path -Path $Env:APPDATA -ChildPath "EnablePython") `
                -ChildPath ("x86-{0}" -f $foundVersion.Platform)
            $script:RESTORE_USER_BASE = $true
        }

        # Get the user scripts path, add it to PATH as well
        $userScriptsPath = & $foundVersion.Executable -E -c 'import sysconfig; print(sysconfig.get_path(""scripts"", scheme=""nt_user""))'
        $Env:PATH = "$userScriptsPath;$Env:Path"

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
                        $pythons.Add((createPythonVersion $tagKey $location.scope))
                    }
                }
            }
        }

        # sort python core first
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

        @(
            @($pythons | Where-Object { $_.Company -eq "PythonCore" } | Sort-Object -Property $sortProperties) +
            @($pythons | Where-Object { $_.Company -ne "PythonCore" } | Sort-Object -Property $sortProperties)
        )
    }
}

function createPythonVersion([Microsoft.Win32.RegistryKey]$tagKey, [string]$scope) {
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

    $defaultProperties = @("Name", "Company", "Tag", "Version", "Platform", "Scope")
    $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet", [string[]]$defaultProperties)
    $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
    $obj | Add-Member MemberSet PSStandardMembers $PSStandardMembers

    $obj
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
    if ($executablePath -eq $null) {
        $null
    }
    else {
        & $executablePath -E -c 'from sys import version_info; print(""{0}.{1}.{2}"".format(version_info[0], version_info[1], version_info[2]))'
    }
}

function is64Bit {
    # Check if this machine is 64-bit
    ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -match '64-bit')
}

Export-ModuleMember "*-*"