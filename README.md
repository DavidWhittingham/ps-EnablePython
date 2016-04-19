EnablePython
============
**EnablePython** is a PowerShell module with functions for listing and enabling (i.e. adding the installation and 
scripts directories to your *PATH* environment variable) installed versions of Python.

Once the module is installed and has been imported, you can list all of the installed Python versions by issuing the
`Get-Python` command. This will display all of the installed Python versions:
```
Name                                    Version                                 Platform
----                                    -------                                 --------
CPython 3.4, x86-64                     3.4                                     64
CPython 2.7, x86-64                     2.7                                     64
CPython 2.7, x86-32                     2.7                                     32
```

To enable a Python installation on your current shell instance, use the `Enable-Python` function.

    Enable-Python -Version 2.7
    
You can optionally supply a platform parameter, either 32 or 64 bit, to select the platform version. If you do not 
supply a platform parameter and both 32-bit and 64-bit versions are installed for the version of Python requested, the
64-bit version is given preference.

    Enable-Python -Version 2.7 -Platform 64

Finally, both arguments are positional, so you can specify them without naming the parameters.

    Enable-Python 2.7 64