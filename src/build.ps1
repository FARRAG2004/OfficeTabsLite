# Builds Office Tabs Lite: add-in DLL -> Center app -> single-file installer (dist\).
# Needs only the .NET Framework 4 compiler that ships with Windows and the Office PIAs
# (Extensibility, office) in the GAC; their types are embedded, so the output needs no PIAs.
$ErrorActionPreference = 'Stop'
$root  = Split-Path $PSScriptRoot -Parent
$csc   = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$ext   = "C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\Extensibility.dll"
$off   = "C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll"
$icon  = "$root\assets\OfficeTabsLite.ico"
$build = "$root\build"
New-Item -ItemType Directory -Force $build, "$root\dist" | Out-Null

function Invoke-Csc([string[]] $cscArgs, [string] $what) {
    & $csc /nologo /platform:anycpu @cscArgs
    if ($LASTEXITCODE -ne 0) { throw "$what failed to compile" }
    "built $what"
}

Invoke-Csc @("/target:library", "/out:$build\OfficeTabsLite.dll",
             "/r:System.dll", "/r:System.Core.dll", "/r:Microsoft.CSharp.dll",
             "/r:System.Drawing.dll", "/r:System.Windows.Forms.dll",
             "/link:$ext", "/link:$off", "$root\src\OfficeTabsLite.cs") "add-in DLL"

Invoke-Csc @("/target:winexe", "/out:$build\OfficeTabsLite-Center.exe", "/win32icon:$icon",
             "/r:System.dll", "/r:System.Drawing.dll", "/r:System.Windows.Forms.dll",
             "/r:$build\OfficeTabsLite.dll", "$root\src\Center.cs") "Center app"

Invoke-Csc @("/target:winexe", "/out:$root\dist\OfficeTabsLite-Setup.exe", "/win32icon:$icon",
             "/r:System.dll", "/r:System.Windows.Forms.dll",
             "/resource:$build\OfficeTabsLite.dll,OfficeTabsLite.dll",
             "/resource:$build\OfficeTabsLite-Center.exe,OfficeTabsLite-Center.exe",
             "$root\src\Installer.cs") "installer"
