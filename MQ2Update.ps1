$title = "MQ2Updater by Blasty v1.0"
$message = "Ensure to place this exefile in the root of your MQ2 installation folder.`nNote this will update any MQ2 installation or make a new clean install."
 
$update = New-Object System.Management.Automation.Host.ChoiceDescription "&Update MQ2", `
    "Update McAfee to latest virus definations (require internet connection)"
 
$options = [System.Management.Automation.Host.ChoiceDescription[]]($update)
 
$result = $host.ui.PromptForChoice($title, $message, $options, 0) 
 
Add-Type -AssemblyName System.IO.Compression.FileSystem
 
function UnzipIT
{
    param([string]$zipfile, [string]$outpath)
 
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}
 
switch ($result){
    0 {
        "`n`nUpdating... Please wait"
        # getting the avvdat file for latest version number
        "Downloading MQ2 Source files..." 
        try {                                             
            # Cleanup
            if (Test-Path "MacroQuest.zip") {
                del MacroQuest.zip
            }
            # Generating target url for update zip file with latest avv signature
            wget https://github.com/macroquest/macroquest/releases/download/rel-live/MacroQuest.zip -OutFile MacroQuest.zip
            Write-host ("Completed!") -foreground "green"
            # Unpack zipfile to target destanation
            try {
                Expand-Archive MacroQuest.zip -DestinationPath .\  -Force
            }
            catch {
                unzipIT MacroQuest.zip .\
            }
            $tid = Get-Date -format "yyyy.MM.dd HH:mm:ss"
            Write-host ("Update complete at " + $tid) -foreground "green"
            #Another cleanup
            if (Test-Path "MacroQuest.zip") {
                del MacroQuest.zip
            }
        }
        catch {
            Write-host ("Could not download MQ2 source") -foreground "red"
        }
    }
}
Pause