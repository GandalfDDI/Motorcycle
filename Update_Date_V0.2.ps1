##########################################################################################################
#
# Update_Date_V0.X.ps1 - Input a directory, change the create, access and write to the current date/time
#			 For all files in and below that directory
#			 Note: Content Created and Date Last Saved In The Properties Tab is NOT changed
#			 As those are "special" properties in the metatdata of the file in the extended
#			 file properties. Example MP3 tag editing:
#			 http://rickgouin.com/use-powershell-to-edit-mp3-tags/
#
# Version 0.2 - Add or subtract minutes
# 
# Version 0.1 - Original
# 
##########################################################################################################

Add-Type -Assembly Microsoft.VisualBasic

$mypath = (Get-Item -Path ".\" -Verbose).FullName

Write-Host "Current Directory $mypath"

$nummin = Read-host "How Many Minutes (Return for 0 minutes)"
if ($nummin -eq "") {$nummin = 0}
if ([Microsoft.VisualBasic.Information]::IsNumeric($nummin) -eq $false) {Write-Host "Minutes is not a number";exit}
$numminn = $nummin

$files = Read-host "What is the name of the directory or files you want to update (Type in *.* to update all files in this folder)?"

Write-host "You are setting the time for the files to a date of " (Get-date).AddMinutes($numminn)

# $ErrorActionPreference= 'silentlycontinue'

$file_full_Path = $mypath + "\" + $files

if ((-not (Test-Path $file_full_Path -PathType Leaf)) -and (-not (Test-Path ($auth_file)))) {
			$exit = Read-host "The file $file_full_Path does not exist"
			exit
			}

Write-Host "Working on the file $file_full_Path"

Get-ChildItem  $file_full_Path -Recurse | % {$_.CreationTime = (Get-date).AddMinutes($numminn)}
Get-ChildItem  $file_full_Path -Recurse | % {$_.LastAccessTime = (Get-date).AddMinutes($numminn)}
Get-ChildItem  $file_full_Path -Recurse | % {$_.LastWriteTime = (Get-date).AddMinutes($numminn)}
