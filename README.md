<!-- markdownlint-disable MD033 -->
# Build-Module

Module Builder

## Syntax

```PowerShell
Build-Module
    -SourceFolder <String>
    -ModulePath <String>
    [-Depth <Int32> = 1]
    [<CommonParameters>]
```

## Description

Build a new module (`.psm1`) file from a folder containing PowerShell scripts (`.ps1` files) and other resources.

This module builder doesn't take care of the module manifest (`.psd1` file), but it simply builds the module file
from the scripts and resources in the specified folder while taking of the following:

* Merging the statements (e.g. `#Requires` and `using` statements)
* Preventing duplicates and collisions (e.g. duplicate function names)
* Ordering the statements based on their dependencies (e.g. classes inheritance)
* Formatting the output.
* Automatically exporting variables, functions, cmdlets, aliases and types.

It doesn't touch any module settings defined in the module manifest (`.psd1` file), such as the module Version,
NestedModules and ScripsToProcess. The only requirement is that the following settings are **not** defined
(or commented out):

* ~~FunctionsToExport = @()~~
* ~~VariablesToExport = @()~~
* ~~AliasesToExport = @()~~

These particular settings are automatically generated based on the cmdlets, variables and aliases defined in the
source scripts and eventually handled by the module (`.psm1`) file.

The general consensus behind this module builder is that the module author defines the items that should be
**loaded** (imported) by putting them in the module source folder and shouldn't be concerned with *invoking*
(dot-sourcing) any scripts knowing that this could lead to similar concerns as using the [`Invoke-Expression`]
cmdlet especially when working in a team. See also [https://github.com/PowerShell/PowerShell/issues/18740](#https-github-com-powershell-powershell-issues-18740).

This means that this module builder will only accept specific statements (blocks) and reject (with a warning) on
statements that require any invocation which potentially could lead to conflicts with other functions and types.

Anything that concerns a dynamic preparation of the module should be done by a specific module manifest setting
or scripted in the `ScriptsToProcess` setting of the module manifest.

## Accepted Statements

The accepted statements might be divided into different files and (sub)folders using any file name or folder name
with the exception of functions that need to be exported as cmdlets.

The accepted statement types are categorized and loaded in the following order:
* [Requirements](#requirements)
* [using statements](#using-statements)
* [enum types](#enum-types)
* [Classes](#classes)
* [Variables assignments](#variables-assignments)
* [(private) Functions](#private-functions)
* [(public) Cmdlets](#public-cmdlets)
* [Aliases](#aliases)
* [Format files](#format-files)

### Requirements

The module builder will merge the `#Requires` statements from the source scripts and will add them to the top of
the module file.

#### #Required -Version

If multiple `#required -version` statements are found, the highest version will be used.

#### #Required -PSEdition

If conflicting `#required -PSEdition` statements are found, a merge conflict exception is thrown.

#### #Required -Modules

If multiple `#required -Modules` statements are found, the module names will be merged and the highest version

#### #Required -RunAsAdministrator

If set in any script, the module builder will add the `#Requires -RunAsAdministrator` statement to the top of the
module file.

> [!TIP](#tip)
> Consider to make your function [self-elevating](https://stackoverflow.com/q/60209449/1701026).

### Using statements

In general `using` statements are merged and added to the module file except for the `using module` with will be
rejected and a warning will be shown.

#### using namespace <.NET-namespace>

The module builder will use the full namespace name and added or merged them accordingly.

#### using module <module-name>

The module builder will reject this statement and will suggest to use the module manifest instead.

#### using assembly <assembly-name>

The module builder will reformat the assembly path and merge the assembly names and add them to module file.

### Enum types

`Enum` and `Flags` types are reformatted using the explicit item value and added or merged them accordingly.

> [!NOTE](#note)
> All types are [automatically added to the TypeAccelerators list][2] to make sure they are publicly available.

### Classes

`Class` definitions are sorted based on any derived (custom) class dependency and added or accordingly.
If conflicting there are multiple classes with the same name a merge conflict exception is thrown unless the
content of the class is exactly the same.

> [!NOTE](#note)
> All types are [automatically added to the TypeAccelerators list][2] to make sure they are publicly available.

> [!WARNING](#warning)
> PowerShell classes do have some known [limitation][3] and know [issues][4] that might cause problems when using
> a module builder. For example, when dividing classes that are depended of each other over multiple files would
> lead to "*Unable to find type [<typename>](#typename)*" in the "PROBLEMS" tab. The only solution is to put these classes
> in the same file or neglect the specific problem.

### Variables assignments

The module builder will merge the variable assignments and add export them when the module is loaded.

> [!TIP](#tip)
> For variables that are dynamically assigned during module load time, consider to use the `ScriptsToProcess`
> setting in the module manifest instead or define the variable during the concerned function or class execution.

### (Private) Functions

Any function that is defined in the source scripts will be added to the module file as a private function.
Meaning the function will not be exported by the module builder and will not be available to the user
when the module is loaded. The function will only be available to other functions in the module file.
To export a function as a cmdlet, the function needs to be defined in a script file with the `.ps1` extension,
see: [(public) cmdlets](#public-cmdlets).

### (Public) cmdlets

Any (public) function that needs to be exported by the module builder is called a [cmdlets][5] in this design.
The module builder will recognize any PowerShell script file (`.ps1`) that contains a `param` block and will treat
it as a cmdlet. The name of the script file will be used as the cmdlet name. Any `Required` or `Using` statement
will be merged and added to the module file.

This module builder design enforces the use of advanced functions and prevents coincidentally interfering with
other cmdlets or other items in the module framework. See also [Add `ScriptsToInclude` to the Module Manifest][4].

### Aliases

This module builder only supports aliases for (public) cmdlets (exported functions). A cmdlet alias might be set
using the [Alias Attribute Declaration][6], this will export the alias when the module is loaded.

> [!NOTE](#note)
> The [`Set-Alias`](https://go.microsoft.com/fwlink/?LinkID=2096625) command statement is rejected as aliases should be avoided for private functions as they can
> make code difficult to read, understand and impact availability (see: [AvoidUsingCmdletAliases][7]).

### Format files

The module builder will accept PowerShell format files (`.ps1xml`) and will merge the view definitions.
For more details on formatting views, see: [about Types.ps1xml][8].

## Rejected Statements

In general, any statement that requires any invocation (dot-sourcing) is rejected by the module builder.
This includes any native cmdlet (and is not limited to) the following specific statements:

### Install-Module

The [`Install-Module`](https://go.microsoft.com/fwlink/?LinkID=398573) cmdlet is rejected along with other cmdlet commands, to specify scripts that run in the
module's session state, use the `NestedModules` manifest setting.

### Install-Module

The [New-Type](#new-type) cmdlet is rejected along with other cmdlet commands, to load any assembly or type definitions
written in a different language than PowerShell, use the `RequiredAssemblies` manifest setting or the
`ScriptsToProcess` manifest setting for any dynamic or conditional loading. Or consider to load the required
type just-in-time while executing the depended class or cmdlet.

## Examples

### <a id="example-1"><a id="example-re-build-a-new-module-file">Example 1: (Re)build a new module file</a></a>


Build a new module file from the scripts in the `.\Scripts` folder and save it to `.\MyModule.psm1`.

```PowerShell
Build-Module -SourceFolder .\Scripts -ModulePath .\MyModule.psm1
```

.PARAMETERS SourceFolder

The source folder containing the PowerShell scripts and resources to build the module from.

.PARAMETERS ModulePath

The path to the module file to create.

> [!WARNING](#warning)
> The module file will be overwritten if it already exists.

.PARAMETERS Depth

The depth of the source folder structure to search for scripts and resources. Default is `1`.

## Parameters

### <a id="-sourcefolder">`-SourceFolder` <a href="https://docs.microsoft.com/en-us/dotnet/api/System.String">&lt;String&gt;</a></a>

```powershell
Name:                       -SourceFolder
Aliases:                    # None
Type:                       [String]
Value (default):            # Undefined
Parameter sets:             # All
Mandatory:                  True
Position:                   # Named
Accept pipeline input:      False
Accept wildcard characters: False
```

### <a id="-modulepath">`-ModulePath` <a href="https://docs.microsoft.com/en-us/dotnet/api/System.String">&lt;String&gt;</a></a>

```powershell
Name:                       -ModulePath
Aliases:                    # None
Type:                       [String]
Value (default):            # Undefined
Parameter sets:             # All
Mandatory:                  True
Position:                   # Named
Accept pipeline input:      False
Accept wildcard characters: False
```

### <a id="-depth">`-Depth` <a href="https://docs.microsoft.com/en-us/dotnet/api/System.Int32">&lt;Int32&gt;</a></a>

```powershell
Name:                       -Depth
Aliases:                    # None
Type:                       [Int32]
Value (default):            1
Parameter sets:             # All
Mandatory:                  False
Position:                   # Named
Accept pipeline input:      False
Accept wildcard characters: False
```

## Related Links

* [cmdlet overview](https://learn.microsoft.com/powershell/scripting/developer/cmdlet/cmdlet-overview)
* [Exporting classes with type accelerators](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_classes#exporting-classes-with-type-accelerators)
* [Class limitations](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_classes#limitations)
* [Various PowerShell Class Issues](https://github.com/PowerShell/PowerShell/issues/6652)
* [Add ScriptsToInclude to the Module Manifest](https://github.com/PowerShell/PowerShell/issues/24253)
* [Alias Attribute Declaration](https://learn.microsoft.com/powershell/scripting/developer/cmdlet/alias-attribute-declaration)
* [Avoid using cmdlet aliases](https://learn.microsoft.com/powershell/utility-modules/psscriptanalyzer/rules/avoidusingcmdletaliases)
* [about Types.ps1xml](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_types.ps1xml)
<!-- -->


[1]: https://learn.microsoft.com/powershell/scripting/developer/cmdlet/cmdlet-overview "cmdlet overview"
[2]: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_classes#exporting-classes-with-type-accelerators "Exporting classes with type accelerators"
[3]: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_classes#limitations "Class limitations"
[4]: https://github.com/PowerShell/PowerShell/issues/6652 "Various PowerShell Class Issues"
[5]: https://github.com/PowerShell/PowerShell/issues/24253 "Add ScriptsToInclude to the Module Manifest"
[6]: https://learn.microsoft.com/powershell/scripting/developer/cmdlet/alias-attribute-declaration "Alias Attribute Declaration"
[7]: https://learn.microsoft.com/powershell/utility-modules/psscriptanalyzer/rules/avoidusingcmdletaliases "Avoid using cmdlet aliases"
[8]: https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_types.ps1xml "about Types.ps1xml"

[comment]: <> (Created with Get-MarkdownHelp: Install-Script -Name Get-MarkdownHelp)
