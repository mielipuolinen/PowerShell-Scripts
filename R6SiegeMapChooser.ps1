#Requires -RunAsAdministrator 
#Requires -Version 5.0

<#
.SYNOPSIS
Map Chooser with Graphical User-Interface for Rainbow 6: Siege
NOTE: This was patched by the end of 2018 by Ubisoft and doesn't work anymore. For archiving purposes only.

.DESCRIPTION
This script allows choosing maps for casual gamemode which isn't possible in the game itself.
The script works by renaming map files and therefore disabling map loading for unwanted maps.
Includes graphical user-interface.

.EXAMPLE
Confirm that $r6s_path is correct, pay attention to folder path format.
PS> R6SiegeMapChooser.ps1
Choose maps to play
(Re)start the game

.NOTES
This script disables maps by adding .deny extension to map file names.
In case you have problems with the script you can rename files back manually in $r6s_path folder.

It should be mentioned that I created this script for learning GUIs in PowerShell and for learning maps in the game which for some mind-boggling doesn't let you to choose maps for casual playing.

GUI created with .NET Framework. I can't remember which .NET and PS versions were required, however this should work on Win10 just fine.

Created in 2018-02-09
Author: Niko Mielikäinen
Git: https://github.com/mielipuolinen
#>

Set-StrictMode -Version Latest

$r6s_path = "C:\Program Files (x86)\Ubisoft\Ubisoft Game Launcher\games\Tom Clancy's Rainbow Six Siege\"

# DO NOT MAKE CHANGES BELOW UNLESS YOU WANT TO BREAK IT (OR FIX IT)

# import/compile ShowWindowAsync class from user32.dll
$showWindowAsync = Add-Type -MemberDefinition "[DllImport(`"user32.dll`")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);" -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
# hide the console window
$showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 0)

function fGetMaps{
    $pvpfiles = Get-ChildItem -Path $r6s_path -Filter "datapc64_pvp*"

    $counter = 1
    $filelist = @()

    while(1){
        $loop = 0

        foreach($pvpfile in $pvpfiles){
            if($pvpfile -like "*pvp$("{0:D2}" -f $counter)*"){ # pvp01, pvp02 ..
                $loop = 1
                $filelist += ,@($counter, $pvpfile)
            }
        }

        # because pvp18 is missing. New map, hype?!
        if($counter -eq 18 -and !$loop){
            $loop = 1
        }

        if($loop){
            $counter++
        }else{
            break
        }
    }

    $maplist = @()
    $old_counter = $counter
    $counter = 1
    while(1){
        foreach($file in $filelist){
            if($file[0] -eq $counter){
                $mapname = $file[1].toString()
                $mapname = $mapname -replace "datapc64_pvp\d\d_"
                $mapname = $mapname -replace "[_.].*"

                $maplist += ,@($counter, $mapname)

                $counter++
                break
            }
        }

        # because pvp18 is missing. New map, hype?!
        if($counter -eq 18){
            $counter++
        }

        # all maps (pvp01,pvp02..) gone through?
        if($counter -eq $old_counter){
            break
        }
    }

    foreach($map in $maplist){
        # fixing some map names
        switch ($map[1]){
            "clubhouse" {$map[1] = "Club House"}
            "university" {$map[1] = "Bartlett U."}
            "russiancafe" {$map[1] = "Kafe Dostoyesvky"}
            "temple" {$map[1] = "Skyscraper"}
            "ibiza" {$map[1] = "Coastline"}
            "themepark" {$map[1] = "Theme Park"}
            "italy" {$map[1] = "Villa"}
            default {}
        }
    
        # first character uppercase on every map name
        $map[1] = $map[1].Substring(0,1).ToUpper()+$map[1].Substring(1)
    }

    return $filelist,$maplist
}

function fMapControl{
    # reset/allow all mapfiles
    foreach($file in $filelist){
        if($file[1].Extension -eq ".deny"){
            $file[1] = $file[1] -replace ".deny"
            Rename-Item $($r6s_path+$file[1]+".deny") $($r6s_path+$file[1])
        }
    }
    
    # deny chosen mapfiles
    foreach($file in $filelist){
        foreach($denymap in $denymaps){
            if($file[1] -like "*pvp$("{0:D2}" -f $denymap)*" -and $file[1].Extension -ne ".deny"){
                Rename-Item $($r6s_path+$file[1]) $($r6s_path+$file[1]+".deny")
                $file[1] = "$($file[1]).deny"
            }
        }
    }
}

$h_fGetMaps = fGetMaps
$filelist = $h_fGetMaps[0]
$maplist = $h_fGetMaps[1]

[xml]$XML = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="R6S: Map Chooser" Height="410" Width="465">
        <Grid Background="#F0F0F0">
            <Label Content="Rainbow 6 Siege" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" FontSize="13"/>
            <Label Content="Map Chooser" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" FontSize="20"/>
            <Label Name="Lbl_Description" Content="Select maps to play." HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top"/>
            $(
                $marginLeft = 10
                $marginTop = 70

                foreach($map in $maplist){
                    "<CheckBox Name=`"Chk_$($map[0])`" Content=`"$($map[1])`" IsChecked=`"False`" HorizontalAlignment=`"Left`" Margin=`"$($marginLeft),$($marginTop),0,0`" VerticalAlignment=`"Top`"/>"
                    
                    $marginTop += 20
                    if($marginTop -gt 250){
                        $marginLeft += 140
                        $marginTop = 70
                    }
                }
            )
            <Button Name="Btn_Apply" Content="Apply Selection" HorizontalAlignment="Left" Margin="10,280,0,0" VerticalAlignment="Top" Width="100"/>
            <Button Name="Btn_About" Content="BeepBoop" HorizontalAlignment="Left" Margin="10,340,0,0" VerticalAlignment="Top" Width="100"/>
            <Button Name="Btn_SelectAll" Content="Select All" HorizontalAlignment="Left" Margin="150,280,0,0" VerticalAlignment="Top" Width="100"/>
            <Button Name="Btn_SelectCompMaps" Content="Select Comps" HorizontalAlignment="Left" Margin="10,310,0,0" VerticalAlignment="Top" Width="100"/>
            <Button Name="Btn_UnselectAll" Content="Unselect All" HorizontalAlignment="Left" Margin="150,310,0,0" VerticalAlignment="Top" Width="100"/>
            <Label Content="v1.1 / 09.02.2018" HorizontalAlignment="Left" Margin="150,340,0,0" VerticalAlignment="Top" FontSize="12"/>
        </Grid>
    </Window>
"@

#Add WPF and Windows Forms assemblies
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Windows.Forms

#Create the XAML reader using a new XML node reader
$GUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $XML))

#Create hooks to each named object in the XAML
$XML.SelectNodes("//*[@Name]") | ForEach-Object{Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name)}

$Btn_Apply.add_Click({
    $Lbl_Description.Content = “Chosen maps are now playable!”

    $denymaps = @(0)
    foreach($map in $maplist){
        if(!((Get-Variable -Name "Chk_$($map[0])").Value.IsChecked)){
            $denymaps += @($map[0])
        }
    }

    $h_fGetMaps = fGetMaps
    $filelist = $h_fGetMaps[0]
    $maplist = $h_fGetMaps[1]
    fMapControl
})

$Btn_SelectAll.add_Click({
    foreach($map in $maplist){
        (Get-Variable -Name "Chk_$($map[0])").Value.IsChecked = $true
    }
})

$Btn_SelectCompMaps.add_Click({
    foreach($map in $maplist){
        if(($map[0] -eq "2") -or ($map[0] -eq "4") -or ($map[0] -eq "7") -or ($map[0] -eq "8") -or ($map[0] -eq "10") -or ($map[0] -eq "12") -or ($map[0] -eq "13") -or ($map[0] -eq "15") -or ($map[0] -eq "16") -or ($map[0] -eq "17") -or ($map[0] -eq "19")){
            (Get-Variable -Name "Chk_$($map[0])").Value.IsChecked = $true
        }
    }
})

$Btn_UnselectAll.add_Click({
    foreach($map in $maplist){
        (Get-Variable -Name "Chk_$($map[0])").Value.IsChecked = $false
    }
})

$Btn_About.add_Click({
    $Lbl_Description.Content = “Greetings to all of you! Have some beeps! <3”
    $b=400;$m=500;$u=600;$d=100;$t=250;
    for($x=0; $x -lt 2; $x++){
        for($y=0; $y -lt 2; $y++){ [console]::beep($b,$t);[console]::beep($m,$t);[console]::beep($u,$t);[console]::beep($m,$t) }
        for($y=0; $y -lt 2; $y++){ [console]::beep($b-$d,$t);[console]::beep($m-$d,$t);[console]::beep($u-$d,$t);[console]::beep($m-$d,$t) }
    }
})
 
#Launch the GUI
$GUI.ShowDialog() | Out-Null
