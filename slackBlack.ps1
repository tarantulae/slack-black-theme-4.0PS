#Must have 7Zip and ASAR addin installed 
#Check Prerequisites
if (!(Test-Path $env:LOCALAPPDATA\slack\app-4*))
{
    [System.Windows.MessageBox]::Show('Slack version 4 not installed','Error: Exiting...')
    exit
}
#Get directories
$AppDirectory = Get-ChildItem -Directory -Path $env:LOCALAPPDATA\slack -Filter "app-4*" | Sort-Object LastAccessTime -Descending | Select-Object -First 1 -ExpandProperty Name
$SlackDirectory = $env:LOCALAPPDATA + "\slack\" + $AppDirectory + "\resources"
$7zipInstall = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip |  Select-Object -ExpandProperty InstallLocation

#Check for 7z install
if(!$7zipInstall)
{
    [System.Windows.MessageBox]::Show('7zip not detected. Install 7zip and ASAR addin','Error: Exiting...')
    exit
}

#Check for ASAR 7z addin
if (!(Test-Path 'C:\Program Files\7-Zip\Formats\Asar*'))
{
    [System.Windows.MessageBox]::Show('ASAR addin must be installed. Download from: http://www.tc4shell.com/binary/Asar.zip','Error: Exiting...')
    exit
}

#Stop slack and Extract to temp directory
Get-Process slack -ErrorAction SilentlyContinue | Stop-Process -PassThru
& $7zipInstall\7z.exe x $SlackDirectory\app.asar "-o$SlackDirectory\app" -y

#Add css to ssb-interop
Add-Content blacktheme.txt -Value "
// First make sure the wrapper app is loaded
document.addEventListener('DOMContentLoaded', function() {

  // Fetch our CSS in parallel ahead of time
  const cssPath =
    'https://raw.githubusercontent.com/caiceA/slack-raw/master/slack-4';
  let cssPromise = fetch(cssPath).then((response) => response.text());
  let customCustomCSS = ``
   :root {
      /* Modify these to change your theme colors: */
      --primary: #61AFEF;
      --text: #ABB2BF;
      --background: #282C34;
      --background-elevated: #3B4048;
   }
   ``;

  // Insert a style tag into the wrapper view
  cssPromise.then((css) => {
    let s = document.createElement('style');
    s.type = 'text/css';
    s.innerHTML = css + customCustomCSS;
    document.head.appendChild(s);
  });
});
"

$blackcss = Get-Content blacktheme.txt
Add-Content $SlackDirectory\app\dist\ssb-interop.bundle.js -Value $blackcss
Remove-Item blacktheme.txt

#Rename old archive
Move-Item $SlackDirectory\app.asar $SlackDirectory\app.asar.original

#Archive new asar
& $7zipInstall\7z.exe a $SlackDirectory\app.asar $SlackDirectory\app