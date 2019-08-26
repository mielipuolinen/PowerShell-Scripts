<#
This can be run as local administrator on workstations or servers.
Up to you how you want to schedule or run it.
Anyhow it's a quick and easy way to clean user profiles from public workstations.
Won't touch logged in users or non-domain users as long as DomainSID is correctly included.
No prompts or errrors for users visible, up to you again to monitor.

Many ways to find DomainSID. DomainSID is a forest/domain wide identifier SID for domain. For example:
PS> Get-CimInstance -ClassName Win32_UserProfile | Where {$_.LocalPath -eq "C:\Users\domain.user"} | % {$_.SID.substring(0,($_.SID.length)-5);}
#>

$InactivityDays = 14
$DomainSID = "S-1-5-21-0123456789-012345678-012345678"

Try{
    $Profiles = Get-CimInstance -ClassName Win32_UserProfile | Where {$_.SID -Like "$DomainSID*"}
    ForEach($Profile in $Profiles){
        if($Profile.LastUseTime -gt ((Get-Date).AddDays(-$InactivityDays))){Continue}
        if($Profile.Loaded -eq "True"){Continue}
        $Profile | Remove-CimInstance
    }
}Catch{}