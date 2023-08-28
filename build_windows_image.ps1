Import-Module .\windows-imaging-tools\WinImageBuilder.psm1
Import-Module .\windows-imaging-tools\Config.psm1
Import-Module .\windows-imaging-tools\UnattendResources\ini.psm1

$type=$args[0]
$release=$args[1]

$ConfigFilePath = ".\config\$type\config-$release.ini"
Get-WimFileImagesInfo D:\windows-sources\$release\sources\install.wim
Set-IniFileValue -Path (Resolve-Path $ConfigFilePath) -Section "DEFAULT" `
                                      -Key "wim_file_path" `
                                      -Value "D:\windows-sources\$release\sources\install.wim"
New-WindowsOnlineImage -ConfigFilePath $ConfigFilePath
