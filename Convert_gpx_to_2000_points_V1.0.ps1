#################################################################################################################
# Convert_gpx_to_2000_points_VX.Y.ps1 - Convert a excel file that has over 2000 points to less than that
# So that Google Maps will show the entire route
#
# Github - https://github.com/GandalfDDI/Motorcycle
# 
# Input - .csv File of all the routes combined taken from the Convert_gpx_to_Google_Plot_VX.Y.ps1 output
#
# Output:
# <name>_out.csv - A CSV file you input into Google Maps to create your own map with less than 2000 data points
#
# Go to https://www.google.com/maps/d/
# 1) Create a new map
# 2) Rename Map, add description, Rename Layer, Click Share and change from private to public to share publicly
# 3) Import --> Import CSV for the day, accept lat and lon, continue, select title and finish
# 4) Click paint bucket net to "All Items", choose More Icons and choose Custom Icon, choose Temp_Black_Circle.gif
# 5) Copy URL
#
# Version 1.0 - Initial
#################################################################################################################

$the_file = Read-Host "What is the name of the .CSV file?"
if (!(Test-Path $the_file)) {Write-Host "File $the_file not found. Press the enter key to exit"}
else {
	$file_in = Get-Content -Path $the_file
}

$outfile = $the_file.Split('.')[0] + "_out.csv"

$numtracks = $file_in.count

# We now have all the tracks. Count them and only export "X" tracks < 2000
# If the number of tracks deviced by the modulus is greater than 2000 then every $tottracks drop a track
$i = 1
$totcount = 0
$tottracks = $numtracks
if ($numtracks -gt 2000) {$tottracks = $numtracks / 2000}
# Put in the header
$file_in[0] | Add-Content $outfile
Write-Host "Number of tracks is $numtracks adding every $tottracks"
while ($i -lt $numtracks) {
# If the number of tracks is greater than 2000 then every $tottracks drop a track
	if ($i -gt $totcount) {
		$file_in[$i] | Add-Content $outfile
		$totcount = $totcount + $tottracks
		}
$i++
}

