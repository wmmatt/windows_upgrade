# Set TLS
[Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([Net.SecurityProtocolType], 3072)

# Set / create dir
$dir = "$env:SystemDrive\Windows\LTSvc\packages"
New-Item $dir -ErrorAction SilentlyContinue | Out-Null
$file = "$($dir)\WindowsUpgrade.exe"

# Get the correct Upgrade Assistant
$os = (Get-ComputerInfo).OSName
switch -Wildcard ($os) {
    '*11*' { $url = 'https://go.microsoft.com/fwlink/?linkid=2171764 '; 'Detected Windows 11' }
    '*10*' { $url = 'https://go.microsoft.com/fwlink/?LinkID=799445 '; 'Detected Windows 10' }
    Default { Return 'Unsupported OS' }
}

# Delete existing upgrader if it exists
Remove-Item $file -Force -ErrorAction SilentlyContinue

# Download Upgrade Assistant
$webClient = New-Object System.Net.WebClient
$webClient.DownloadFile($url,$file)

# Install latest Windows 10/11
'Starting upgrade process...'
Start-Process -FilePath $file -ArgumentList '/quietinstall /skipeula /auto upgrade /copylogs $dir'

<#
%windir%\System32\WindowsPowerShell\v1.0\powershell.exe -Command "&{ [Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([Net.SecurityProtocolType], 3072); $dir = \"%windir%\LTSvc\packages\"; New-Item $dir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null; $webClient = New-Object System.Net.WebClient; $os = (Get-ComputerInfo).OSName; switch -Wildcard ($os) { '*11*' { $url = 'https://go.microsoft.com/fwlink/?linkid=2171764 '; 'Detected Windows 11' }; '*10*' { $url = 'https://go.microsoft.com/fwlink/?LinkID=799445 '; 'Detected Windows 10' }; Default { Return 'Unsupported OS' }}; $file = \"$dir\WindowsUpgrade.exe\"; Remove-Item $file -Force -ErrorAction SilentlyContinue; $webClient.DownloadFile($url,$file); 'Beginning Windows Upgrade...'; Start-Process -FilePath $file -ArgumentList '/quietinstall /skipeula /auto upgrade /copylogs $dir' }"
#>
