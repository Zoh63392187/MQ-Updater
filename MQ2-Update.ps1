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
$LocalUpdateForm.text            = "MQ2-Updater"
$LocalUpdateForm.BackColor       = "#ffffff"
$LocalUpdateForm.TopMost         = $false

$Titel                           = New-Object system.Windows.Forms.Label
$Titel.text                      = "MQ2-Updater by Blasty v1.1"
$Titel.AutoSize                  = $true
$Titel.width                     = 25
$Titel.height                    = 10
$Titel.location                  = New-Object System.Drawing.Point(20,20)
$Titel.Font                      = 'Microsoft Sans Serif,13'

$Description                     = New-Object system.Windows.Forms.Label
$Description.AutoSize            = $false
$Description.width               = 450
$Description.height              = 60
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
        $AddUpdateBtn.text = "Update MQ2"
        $Description.text = "This script will update any current MQ2 installation or make a new install.`n(Even MQ2 from RedGuides and also works with all class plugins)`n`nMake sure that MQ2 is NOT running."
    }
    catch {
        $AddUpdateBtn.Visible = $false
        $Description.ForeColor = "#FF0000"
        $Description.text = "MQ2 is running - aborting!"
    }
    
} else { 
    $AddUpdateBtn.text = "New Install" 
    $Description.text = "This script will update any current MQ2 installation or make a new install.`n(Even MQ2 from RedGuides and also works with all class plugins)`n`nMQ2 NOT detected!"
}

$LocalUpdateForm.controls.AddRange(@($Titel,$Description,$AddUpdateBtn))

function UnzipIT{
	param([string]$zipfile, [string]$outpath)
	[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

function AddUpdate {
    $AddUpdateBtn.Visible = $false
	try {                                      
        # declaring color
        $Description.ForeColor         = "#007500"
        # Cleanup
        if (Test-Path "MacroQuest.zip") {
            del MacroQuest.zip
        }
        $Description.text = "Downloading... Please be patient"
        # Generating target url for update zip file with latest avv signature
        wget https://github.com/macroquest/macroquest/releases/download/rel-live/MacroQuest.zip -OutFile MacroQuest.zip
        # Unpack zipfile to target destanation
        try {
            Expand-Archive MacroQuest.zip -DestinationPath .\  -Force
        }
        catch {
            unzipIT MacroQuest.zip .\
        }
        $tid = Get-Date -format "yyyy.MM.dd HH:mm:ss"
        $Description.text = "Update complete at " + $tid
        #Another cleanup
        if (Test-Path "MacroQuest.zip") {
            del MacroQuest.zip
        }
        
    }
    catch {
        $Description.ForeColor = "#f00000"
        $Description.text = "Update failed!"
    }
}

$AddUpdateBtn.Add_Click({ AddUpdate })
[void]$LocalUpdateForm.ShowDialog()