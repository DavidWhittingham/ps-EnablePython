EnablePython
============
**EnablePython** is a PowerShell module with functions for listing and enabling (i.e. adding the installation and
scripts directories to your *PATH* environment variable) installed versions of Python.

Installation
------------
**EnablePython* is [availble on the PowerShell Gallery](https://www.powershellgallery.com/packages/EnablePython/), and
can be installed at the PowerShell prompt:

    PS> Install-Module -Name EnablePython

Using the Module
----------------

### Getting Python Installs
Once the module is installed and has been imported, you can list all of the installed Python versions by issuing the
`Get-Python` command. This will display all of the installed Python versions:
```
Name     : PythonCore 2.7 (x86-64)
Company  : PythonCore
Tag      : 2.7
Version  : 2.7.10
Platform : 64
Scope    : AllUsers

Name     : PythonCore 2.7 (x86-32)
Company  : PythonCore
Tag      : 2.7
Version  : 2.7.10
Platform : 32
Scope    : AllUsers

Name     : Continuum Analytics, Inc. Anaconda 4.5.11 (x86-64)
Company  : ContinuumAnalytics
Tag      : Anaconda37-64
Version  : 3.7.0
Platform : 64
Scope    : CurrentUser
```

You can use a number of parameters to filter the list:

- Scope (e.g. *CurrentUser*, *AllUsers*)
- Company (e.g. *PythonCore*, *ContinuumAnalytics*)
- Version (e.g. *3.5*, *2.7*)
- Platform (e.g. *64*, *32*)
- Tag (e.g. *Anaconda37-64*)

The **Company**, **Version**, and **Tag** filters all match from the start of the string with a wildcard at the end, so a value of *Continnuum* for **Company** would match *ContinuumAnalytics*, but a value of *conda* for **Tag** will not match *Anaconda37-64*.

The list of installs is sorted as per the order of properties listed above.  **Scope** sorts *CurrentUser* before *AllUsers*, **Company** is sorted alphabetically (except for *PythonCore*, which is always listed first if available), **Version** is sorted in descending order (highest available version first), **Platform** is sorted in descending order (64-bit before 32-bit), and Tag is sorted alphabetically.

#### Examples
Get all Python installs available and sort them:

    PS> Get-Python


    Name     : PythonCore 2.7 (x86-64)
    Company  : PythonCore
    Tag      : 2.7
    Version  : 2.7.10
    Platform : 64
    Scope    : AllUsers

    Name     : PythonCore 2.7 (x86-32)
    Company  : PythonCore
    Tag      : 2.7
    Version  : 2.7.10
    Platform : 32
    Scope    : AllUsers

    Name     : Continuum Analytics, Inc. Anaconda 4.5.11 (x86-64)
    Company  : ContinuumAnalytics
    Tag      : Anaconda37-64
    Version  : 3.7.0
    Platform : 64
    Scope    : CurrentUser

Get all Python installs from a company and sort them:

    PS> Get-Python -Company ContinuumAnalytics


    Name     : Continuum Analytics, Inc. Anaconda 4.5.11 (x86-64)
    Company  : ContinuumAnalytics
    Tag      : Anaconda37-64
    Version  : 3.7.0
    Platform : 64
    Scope    : CurrentUser

Gets all Python installs with a particular version/platform and sort them:

    PS> Get-Python -Version 2.7 -Platform 32


    Name     : PythonCore 2.7 (x86-32)
    Company  : PythonCore
    Tag      : 2.7
    Version  : 2.7.10
    Platform : 32
    Scope    : AllUsers

The **Version** and **Platform** parameters are positional, and can be specified without their respective names:
    PS> Get-Python 2.7 64


    Name     : PythonCore 2.7 (x86-64)
    Company  : PythonCore
    Tag      : 2.7
    Version  : 2.7.10
    Platform : 64
    Scope    : AllUsers

### Enabling Python Installs
To enable a Python installation on your current shell instance, use the `Enable-Python` function.  This function supports the same filtering and does the same sorting as the `Get-Python` function.  The top-most sorted Python install will be enabled, as long as one or more installs are found (given supplied parameters).

#### Examples
Get all Python installs available, sort them, and enable the top-most install:

    PS> Enable-Python
    Multiple Python distributions found matching that criteria, enabling the top choice...
    "PythonCore 2.7 (x86-64)" has been enabled.

Get all Python installs from a company, sort them, and enable the top-most install:

    PS> Enable-Python -Company ContinuumAnalytics
    "Continuum Analytics, Inc. Anaconda 4.5.11 (x86-64)" has been enabled.

Get all Python installs with a particular version/platform, sort them, and enable the top-most install:

    PS> Enable-Python -Version 2.7 -Platform 32
    "PythonCore 2.7 (x86-32)" has been enabled.

The **Version** and **Platform** parameters are positional, and can be specified without their respective names:
    PS> Enable-Python 2.7 64
    "PythonCore 2.7 (x86-64)" has been enabled.