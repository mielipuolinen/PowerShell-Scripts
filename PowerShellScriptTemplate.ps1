#Requires -RunAsAdministrator 
#Requires -Version 6.0

<#
.SYNOPSIS
A brief description of script.
Prints two strings, by default "Hello World!"

.DESCRIPTION
A more detailed description of script.
This script combines two strings and returns it. Pipelining (input and output) and parameters are supported, see examples.

.PARAMETER String1
First string.

.PARAMETER String2
Second string.

.EXAMPLE
PS> PowerShellTemplate.ps1 -String1 "Hi" -String2 "GitHub!"
Hi GitHub!

.EXAMPLE
PS> "What's up" | PowerShellTemplate.ps1
What's up World!

.EXAMPLE
PS> PowerShellTemplate.ps1 | %{Return $PSItem}
Hello World!

.EXAMPLE
PS> PowerShellTemplate.ps1
Hello World!

.INPUTS
A list of type of objects that can be piped into this script.
String

.OUTPUTS
A list of type of objects that this script returns and therefore can be piped forward.
String

.LINK
A full link to a related topic. Could be also just a name of a related topic. Repeat this for all available topics.

.LINK
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help

.LINK
https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_requires

.LINK
https://docs.microsoft.com/en-us/powershell/scripting/overview

#>

[CmdletBinding()]
Param(
    [Parameter(ValueFromPipeline=$true)]
    [String]$String1 = "Hello",
    [String]$String2 = "World!"
)

Function MyFunction([String]$String1 = "", [String]$String2 = ""){
    Return "$($String1) $($String2)"
}

Return "$(MyFunction -String1 $String1 -String2 $String2)"