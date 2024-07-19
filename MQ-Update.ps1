# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
# Init PowerShell Gui
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#---------------------------------------------------------[Form]--------------------------------------------------------

[System.Windows.Forms.Application]::EnableVisualStyles()

$LocalUpdateForm                 = New-Object system.Windows.Forms.Form
$LocalUpdateForm.ClientSize      = '480,200'
$LocalUpdateForm.text            = "MQ-Updater"
$LocalUpdateForm.BackColor       = "#ffffff"
$LocalUpdateForm.TopMost         = $false
$LocalUpdateForm.StartPosition = "CenterScreen"

$Titel                           = New-Object system.Windows.Forms.Label
$Titel.text                      = "MQ-Updater by Blasty v1.1"
$Titel.AutoSize                  = $true
$Titel.width                     = 25
$Titel.height                    = 10
$Titel.location                  = New-Object System.Drawing.Point(20,20)
$Titel.Font                      = 'Microsoft Sans Serif,13'

$Description                     = New-Object system.Windows.Forms.Label
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 90
$Description.location            = New-Object System.Drawing.Point(20,50)
$Description.Font                = 'Microsoft Sans Serif,10'

$AddUpdateBtn                   = New-Object system.Windows.Forms.Button
$AddUpdateBtn.BackColor         = "#ff7b00"
$AddUpdateBtn.width             = 110
$AddUpdateBtn.height            = 30
$AddUpdateBtn.location          = New-Object System.Drawing.Point(360,150)
$AddUpdateBtn.Font              = 'Microsoft Sans Serif,10'
$AddUpdateBtn.ForeColor         = "#ffffff"
$AddUpdateBtn.Visible           = $true 

if (Test-Path "MacroQuest.exe"){ 
    try { 
        [IO.File]::OpenWrite("crashpad_handler.exe").close()
        $AddUpdateBtn.text = "Update MQ"
        $Description.text = "This will update any current MQ installation or make a new install.`n(Even MQ from RedGuides and also works with all class plugins)`n`nMake sure that MQ is NOT running.`n`nAll credits to the MQ devs!"
    }
    catch {
        $AddUpdateBtn.Visible = $false
        $Description.ForeColor = "#FF0000"
        $Description.text = "MQ is running - aborting!"
    }
} else { 
    $AddUpdateBtn.text = "New Install" 
    $Description.text = "This will make a new install as MQ is NOT detected!`nIf you do have MQ installed make sure to put this script in the root of your MQ folder."
}

$LocalUpdateForm.controls.AddRange(@($Titel,$Description,$AddUpdateBtn))

function UnzipIT{
	param([string]$zipfile, [string]$outpath)
	[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

function AddUpdate {
    $AddUpdateBtn.Visible = $false
	try {                                      
        # Cleanup
        if (Test-Path "MacroQuest.zip") {
            del MacroQuest.zip
        }
        $Description.text = "Downloading... Please be patient"        
        Import-Module BitsTransfer
        Start-BitsTransfer https://github.com/macroquest/macroquest/releases/download/rel-live/MacroQuest.zip MacroQuest.zip
        # Unpack zipfile to target destanation
        try {
            Expand-Archive MacroQuest.zip -DestinationPath .\  -Force
        }
        catch {
            unzipIT MacroQuest.zip .\
        }
        # declaring color
		$Description.ForeColor         = "#007500"
		$tid = Get-Date -format "yyyy.MM.dd HH:mm:ss"
        $Description.text = "Complete at " + $tid
        #Another cleanup
        if (Test-Path "MacroQuest.zip") {
            del MacroQuest.zip
        }
    }
    catch {
        $Description.ForeColor = "#f00000"
        $Description.text = "MQ update Failed!"
    }
	try {
		Start-BitsTransfer http://77.66.65.240/ZoneConnections.yaml resources/EasyFind/ZoneConnections.yaml
	}
	catch {
        $Description.ForeColor = "#f00000"
        $Description.text = "EasyFind Failed!"
    }
	try {
		Start-BitsTransfer http://77.66.65.240/guildhalllrg.navmesh resources/MQ2Nav/guildhalllrg.navmesh
	}
	catch {
        $Description.ForeColor = "#f00000"
        $Description.text = "Guildhall mesh failed Failed!"
    }
}

$AddUpdateBtn.Add_Click({ AddUpdate })
[void]$LocalUpdateForm.ShowDialog()