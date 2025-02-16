#################################################################################################################
# Convert_gpx_to_Google_Plot_VX.Y.ps1 - Convert a GPS .gpx file (XML format) to a file that Google can plot for you
# Run the output file YYYY-MM-DD.html in Firefox or Chrome for correct output
#
# Github - https://github.com/GandalfDDI/Motorcycle
#
# Input - GPX File froma Garmin Device. Found under H:\Garmin\GPX or (older tracks) H:\Garmin\GPX\Archive
# Each GPX holds 3 days worth so the device will hold MAYBE 15 days worth before it rolls over ???
# Current.gpx
#
# Note: Some GPS's do not honor the time rollover. This script assumes that if the date is before 2008 (somewhat arbitrary
# choice in years) then the device is confused, therefore if (see https://en.wikipedia.org/wiki/GPS_Week_Number_Rollover ):
# Year is less than 2008 Add 7168 days
#
# To Do --> Sort the XML by date and dedup
#
# Output:
# YYYY-MM-DD.csv - A CSV file you input into Google Maps to create your own map
# YYYY-MM-DD.html - A map on Bing that you can bring up to show the same information, will not display as
#                   a web page on a web site? Maybe fix later, in the meantime use the Import Create Google Map instructions below
# YYYY-MM-DD_elev.csv - A Elevation Profile of the route in title, date, altitude, feet - The points will match the first CSV file
#
# Requirements for Formal-XML - Pscx package
# Command Line Install thru NuGet:
# Install-Module Pscx -Scope CurrentUser -AllowClobber
# Note:
# The following commands are already available on this system:'gcb,Expand-Archive,Format-Hex,Get-Hash,help,prompt,
# Get-Clipboard,Get-Help,Set-Clipboard'. This module 'Pscx' may override the existing commands
# 
# File to base64 variable for the icon symbol, create your own icon symbol and perform the following:
# $FilePath = Read-Host 'What is the filename of the new icon symbol you want to use?'
# $blackcircle = [Convert]::ToBase64String((Get-Content -Path $FilePath -Encoding Byte))
# $blackcircle > t.txt
# Copy the B64 in t.txt to the below string $blackcircle
# Output B64 to filename "Temp_Black_Circle.gif":
# $FilePath = "Temp_Black_Circle.gif"
# [Convert]::FromBase64String((Get-Clipboard)) | Set-Content -Path $FilePath -Encoding Byte
#
# Import / create a Google Map:
# Import the CSV file with the columns:
#	title,Date_Altitude,Cumulative_Distance,latitude,longitude
#	1,05/15/2020 Altitude: 47.77 Feet,0,44.631569,-124.050369
#	2,05/15/2020 Altitude: 49.34 Feet,0.005,44.631510,-124.050322
#	3,05/15/2020 Altitude: 52.49 Feet,0.015,44.631433,-124.050494
#	4,05/15/2020 Altitude: 73.00 Feet,0.043,44.631248,-124.051012
#
# Go to https://www.google.com/maps/d/
# 1) Create a new map
# 2) Rename Map, add description, Rename Layer, Click Share and change from private to public to share publicly
# 3) Import --> Import CSV for the day, accept lat and lon, continue, select title and finish
# 4) Click paint bucket net to "All Items", choose More Icons and choose Custom Icon, choose Temp_Black_Circle.gif
# 5) Copy URL
#
# Create Excel Elevation Profile:
# 1) Open YYYY-MM-DD_elev.csv
# 2) Selct Columns Cumulative_Distance_Miles and Altitude_feet
# 3) Insert --> Charts --> Scatter --> Scatter With Smooth Lines
#
# Version 2.00 - Update for Garmin's GpxExtensions/v3 TrackPointExtension/v2, this version only gathers the
#		 Lat/Lon, Elevation and Time into a list and plots that ignoring all other extension data
#
# Version 1.91 - Use the current folder where the script is running
#
# Version 1.9 - Correct counter on Delta Altitude
#
# Version 1.8 - Correct Delta Altitude, since the single data points jump around quite a bit average the last 'X'
#		since you get several data points per minute
#
# Version 1.7 - Fix Date / Time issue WRT crossing UTC Midnight, if the difference between data points is more
#		than "X" hours ($hours_between_plots) then assume a new plot
#
# Version 1.6 - Add a sheet for distance travelled vs Altitude so that an Excel plot showing a Elevation Profile can be done
#
# Version 1.5 - Convert altitude units from meters to feet ('cause this is 'Murica ...) and round to two digits
#		Seriously ... Change "$metric = 0" to "$metric = 1" and
#		I promise by the time you're done changing it, you'll feel right as rain.
#
# Version 1.4 - Limit number of data points to less than 2,000 as that is the max that can be displayed
#
# Version 1.3 - Show Date_Time_Altitude,Speed,Distance,Delta_Distance,Delta_Altitude_Gain
#
# Version 1.2 - Ask user if they want to show specific times, calculate speed if so.
# Example:
# PS> $a = New-Object System.Device.Location.GeoCoordinate 46.985861,-120.566050
# PS> $b = New-Object System.Device.Location.GeoCoordinate 46.998460,-120.581326
# PS> $a.GetDistanceTo($b)*3.2808399
# 5969.64998686707
# So 5969.65 feet
#
# Version 1.1 - Remove specific times so that speed cannot be calculated
#		Produce one map per day
#
# Version 1.0 - Initial
#################################################################################################################
# Metric variable. $metric = 1, yes use the metric system, $metric = 0 use the British Imperial System system
$metric = 0

$distsmall = 1.0
$distlarge = 1000.0
$distsmalldesc = "Meters"
$distlargedesc = "Kilometer"
$hours_between_plots = 4
$average_alt = 4
$average_alt_sum = 0

if ($metric -eq 0) {
	$distsmall = 3.2808399
	$distlarge = 5280.00
	$distsmalldesc = "Feet"
	$distlargedesc = "Mile"
	}


# Add library for Lat / Lon distance
Add-Type -AssemblyName System.Device

# Create Marker Icon file. To create your own see above
$FilePath = $PSScriptRoot + "\" + "Temp_Black_Circle.gif"
$blackcircle = "R0lGODlhBgAGAHAAACH5BAEAAPwALAAAAAAGAAYAhwAAAAAAMwAAZgAAmQAAzAAA/wArAAArMwArZgArmQArzAAr/wBVAABVMwBVZgBVmQBVzABV/wCAAACAMwCAZgCAmQCAzACA/wCqAACqMwCqZgCqmQCqzACq/wDVAADVMwDVZgDVmQDVzADV/wD/AAD/MwD/ZgD/mQD/zAD//zMAADMAMzMAZjMAmTMAzDMA/zMrADMrMzMrZjMrmTMrzDMr/zNVADNVMzNVZjNVmTNVzDNV/zOAADOAMzOAZjOAmTOAzDOA/zOqADOqMzOqZjOqmTOqzDOq/zPVADPVMzPVZjPVmTPVzDPV/zP/ADP/MzP/ZjP/mTP/zDP//2YAAGYAM2YAZmYAmWYAzGYA/2YrAGYrM2YrZmYrmWYrzGYr/2ZVAGZVM2ZVZmZVmWZVzGZV/2aAAGaAM2aAZmaAmWaAzGaA/2aqAGaqM2aqZmaqmWaqzGaq/2bVAGbVM2bVZmbVmWbVzGbV/2b/AGb/M2b/Zmb/mWb/zGb//5kAAJkAM5kAZpkAmZkAzJkA/5krAJkrM5krZpkrmZkrzJkr/5lVAJlVM5lVZplVmZlVzJlV/5mAAJmAM5mAZpmAmZmAzJmA/5mqAJmqM5mqZpmqmZmqzJmq/5nVAJnVM5nVZpnVmZnVzJnV/5n/AJn/M5n/Zpn/mZn/zJn//8wAAMwAM8wAZswAmcwAzMwA/8wrAMwrM8wrZswrmcwrzMwr/8xVAMxVM8xVZsxVmcxVzMxV/8yAAMyAM8yAZsyAmcyAzMyA/8yqAMyqM8yqZsyqmcyqzMyq/8zVAMzVM8zVZszVmczVzMzV/8z/AMz/M8z/Zsz/mcz/zMz///8AAP8AM/8AZv8Amf8AzP8A//8rAP8rM/8rZv8rmf8rzP8r//9VAP9VM/9VZv9Vmf9VzP9V//+AAP+AM/+AZv+Amf+AzP+A//+qAP+qM/+qZv+qmf+qzP+q///VAP/VM//VZv/Vmf/VzP/V////AP//M///Zv//mf//zP///wAAAAAAAAAAAAAAAAgSAPcBGAhAIMGD+wwOTHgQIcKAADs="
[Convert]::FromBase64String($blackcircle) | Set-Content -Path $FilePath -Encoding Byte

$csv_title = "title,Date_Altitude,Cumulative_Distance,latitude,longitude"
$csvelev_title = "title,Date,Cumulative_Distance_Miles,Altitude_feet,latitude,longitude"

# Create header and footer
$header = (
"<html>",
"<head>",
"<meta name='viewport' content='initial-scale=1.0, user-scalable=no' /><style type='text/css'>",
"  html { height: 100% }",
"  body { height: 100%; margin: 0; padding: 0 }",
"  #map_canvas { height: 100%}",
"</style>",
"<script type='text/javascript' src='http://dev.virtualearth.net/mapcontrol/mapcontrol.ashx?v=7.0'></script>",
"<script type='text/javascript'>",
"function mcpherGetCredentials() { return 'Ao1OcFRQ24TL8Ek0Mautt7PfMkvb7vZvsv48AkWDVnKHfvUuVGfv_bDgqG9VS-0J';}",
"function mcpherDataPopulate() { ",
"var mcpherData ={'cJobject':["
)

# Each point has a format as such, which goes between the $header and the $trailer variable
# {'title':'1','content':'\<b\>8/10/2014\</b\>\<br\>781 ft\<br\>','lat':'58.03824','lng':'-5.06937'},
# <snip>
# {'title':'1621','content':'\<b\>8/10/2014\</b\>\<br\>781 ft\<br\>','lat':'58.03730','lng':'-5.06900'},
#
# Then we add the trailer

$trailer = (
"{'title':'','content':'','lat':'','lng':''}]",
"};",
"return mcpherData; };",
"",
"//------",
"function initialize() {",
"    var mcpherData = mcpherDataPopulate();",
"    if (mcpherData.cJobject.length > 0) {",
"     var myOptions = {",
"   center: new Microsoft.Maps.Location(parseFloat(mcpherData.cJobject[0].lat), ",
"    parseFloat(mcpherData.cJobject[0].lng)), ",
"        mapTypeId: Microsoft.Maps.MapTypeId.auto,",
"        zoom : 15,",
"   credentials : mcpherGetCredentials()",
"     };",
"    // get parameters if any",
"     var qparams = mcpherGetqparams();",
"     if (qparams['zoom']) myOptions['zoom'] = parseInt(qparams['zoom']);",
"  // create the map",
"      var mapContainer = document.getElementById('map_canvas');",
"      map = new Microsoft.Maps.Map(mapContainer,",
"   {credentials: myOptions['credentials'],",
"    center: myOptions['center'],",
"    zoom: myOptions['zoom'],",
"    mapTypeId: myOptions['mapTypeId']}",
"   );",
"     map.entities.clear(); ",
"// add the excel data",
"     for ( var i = 0 ; i < mcpherData.cJobject.length;i++) ",
"                    mcpherAddMarker ( map, mcpherData.cJobject[i] );",
"     }",
"};",
"function mq() { return String.fromCharCode(34); }",
"function mcpherAddMarker(gMap, cj ) {",
"var pushpin= new Microsoft.Maps.Pushpin(map.getCenter(), null); ",
"  infoBoxStyle = mq() + ",
"   'background-color:White; border: medium solid DarkOrange; font-family: Sans-serif; width:auto; font-size: 70%;' + mq();",
"  hoverBoxStyle = mq() + ",
"   'background-color:Cornsilk;  border-width:0; width:auto; font-size: 70%; font-family: Sans-serif;' + mq();",
"  var gp = new Microsoft.Maps.Location(parseFloat(cj.lat), parseFloat(cj.lng));",
"  var marker = new Microsoft.Maps.Pushpin(gp,{icon: 'Temp_Black_Circle.gif', width:'20px', height:'20px'}); ",
"  gMap.entities.push(marker);",
"  var hoverBox = new Microsoft.Maps.Infobox(gp,  { ",
"  htmlContent: '<div style='+ hoverBoxStyle + '>' + cj.title + '</div>',visible:false} ); ",
"  gMap.entities.push(hoverBox);",
"  Microsoft.Maps.Events.addHandler(marker, 'mouseover', function() {",
"    hoverBox.setOptions({visible:true}); });",
"  Microsoft.Maps.Events.addHandler(marker, 'mouseout', function() {",
"    hoverBox.setOptions({visible:false}); ",
"                   infoBox.setOptions({visible:false}); });",
" if (cj.content){",
"   var infoBox = new Microsoft.Maps.Infobox(gp, { ",
"  htmlContent: '<div style='+ infoBoxStyle + '>' + cj.content + '</div>',visible:false} ); ",
"   gMap.entities.push(infoBox);",
"        Microsoft.Maps.Events.addHandler(marker, 'click', function() {",
"    infoBox.setOptions({visible:true}); });",
"        Microsoft.Maps.Events.addHandler(marker, 'mouseout', function() {",
"    hoverBox.setOptions({visible:false}); ",
"                   infoBox.setOptions({visible:false}); });",
"      }",
" else",
"    Microsoft.Maps.Events.addHandler(marker, 'mouseout', function() {",
"    hoverBox.setOptions({visible:false});  });",
"  return marker;",
"};",
"function mcpherGetqparams(){",
"      var qparams = new Array();",
"   var htmlquery = window.location.search.substring(1);",
"   var htmlparams = htmlquery.split('&');",
"   for ( var i=0; i < htmlparams.length;i++) {",
"     var k = htmlparams[i].indexOf('=');",
"     if (k > 0) qparams[ htmlparams[i].substring(0,k) ] = decodeURI(htmlparams [i].substring(k+1));",
" }    ",
" return qparams;",
"   };  ",
"</script>",
"</head>",
"<body onload='initialize()'>",
"  'GPX to Bing  mapping - http://ramblings.mcpher.com/Home/excelquirks/getmaps/bingmarker'",
"  <div id='map_canvas' style='width:100%; height:100%'></div>",
"</body>",
"</html>"
)

$dir = $PSScriptRoot

Set-Location $PSScriptRoot

Write-Host "Running script from" $dir
Write-Host "The GPX files are expected to be in this folder and ouput will land in this folder"

# Ask if we REALLY want the speed calculated
Write-Host "Reminder 1: In Garmin Basecamp select path and 'Export Selection' not export date"
Write-Host "Reminder 2: If you note speed in your graph and you publicly publish this map the police can use that as evidence that you broke the speed limit. Just sayin'"
$the_speed = Read-Host "Do you want the WHOLE ENCHALADA added to the tracks (Speed,Distance,Delta_Distance,Delta_Altitude_Gain, anything but 'HELL YES' is a no)?"

if ($the_speed -eq 'HELL YES') {$csv_title = "title,Date_Time_Altitude_" + $distsmalldesc + ",Speed (" + $distlargedesc + "s per hour),Distance (" + $distsmalldesc + "),Delta_Distance,Delta_Altitude_Gain,latitude,longitude"}

# Get the files and read in all the waypoints

$gpx_file = "start"

while ($gpx_file -ne "") {
	$the_file = Read-Host 'What is the GPX file with tracks? (Current.gpx --> Older version, CurrentTrackLog.gpx --> newer version if available)'
	if ($the_file -eq "") {Break}
	if (!(Test-Path $the_file)) {$file_not_found = Read-Host "File $the_file not found. Press the enter key to exit" ;exit}
	else {
		[xml]$xmlgpxfile = Get-Content -Path $the_file
		$xmlgpxtrks += $xmlgpxfile.gpx.trk
	}
}

# We now have all the tracks. Format them from the XML and put them into a variable, read them all and start adding to an HTML file
# One HTML File for trips that have a time break of less than 4 hours. Note that times are "local" time from the computer where the script
# was run, not UTC.
$tracks_xml = $xmlgpxtrks.trkseg | Format-Xml

# Since extension mess up the count, previously 4 lines between track, now "n" (not standard) between tracks
# Just make a list of the lines with the Lat/Lon, Elevation and Time
$tracks_all = $tracks_xml -split "\n"
$tracks_all_count = $tracks_all.Count
$tracks = New-Object System.Collections.Generic.List[System.String]
for ($cd = 0; $cd -lt $tracks_all_count; $cd++) {
	if ($tracks_all[$cd] -match "<trkpt lat=") {$tracks.Add($tracks_all[$cd])}
	if ($tracks_all[$cd] -match "<ele>") {$tracks.Add($tracks_all[$cd])}
	if ($tracks_all[$cd] -match "<time>") {$tracks.Add($tracks_all[$cd])}
	}

# Get initial day so that we can break up each map into single days, account for GPS Date less than 2008 and if so add date adjustment
[datetime]$dstring = (($tracks[2] -split "<time>")[1] -split "</time>")[0]
if ($dstring.Year -lt 2008) { $dstring = $dstring.AddDays(7168)}
$dstringold = $dstring
$i=1
$tracks_count = $tracks.Count

# Iterate through all of the tracks, when you get a track that is more than $hours_between_plots hours, put out that graph and start a new one
# When putting out a new graph make sure that number of data points is less than or equal to 2,000 data points
while ($i -lt $tracks_count) {
# Start with every data point. Adjust this number if it is greater than 2000
	$modtracks = 1
	$numtracks = 0
	$trackhtml = @() 
	$trackcsv = @() 
	$trackcsvelev = @()
	$cdayh = ""
# Count the total number of tracks per graph. For the date if the delta time between tracks is > 4 hours then assume a new day / track has started
	for ($cd = $i; $cd -lt $tracks_count; $cd++) {
		if ($tracks[$cd] -match "<time>") {
#			Write-Host "Convert " $tracks[$cd]
			[datetime]$dstring = (($tracks[$cd] -split "<time>")[1] -split "</time>")[0]
# If this is the first time through then initialize the filename and set the old date to the first date
			if ($cdayh -eq "") {$cdayh = $dstring.ToString("yyyy-MM-dd-HH"); $dstringold = $dstring}
			if ($dstring.Year -lt 2008) { $dstring = $dstring.AddDays(7168)}
			$numtracks++
# If we find that there are more than $hours_between_plots then bail out
#			Write-Host "Hours dstring" $dstring " and distring old" $dstringold "total hours " ($dstring - $dstringold).TotalHours
			if (($dstring - $dstringold).TotalHours -gt $hours_between_plots) {$cd = $tracks_count}
			else {$dstringold = $dstring}
			}
		}
# Since we went one past the track we want, subtract out a track
	$numtracks--
# If the number of tracks divided by the modulus is greater than 2000 then every $tottracks drop a track
	$totcount = 0
# Anything less than 1 and all tracks are included
	$tottracks = 0.9
# Again, the maps will not allow more than 2000 data points so only allow 2000 data points
	if ($numtracks -gt 2000) {$tottracks = $numtracks / 2000}

# Initialize variables for this set of tracks
	$old_lat = -1
	$old_lon = -1
	[datetime]$old_time = Get-Date
	$ddistance = 0
	$edistance = 0
	$daltitude = 0
	$old_alt = -1000
	$dstringold = ""
	$ttcount = 1
# If we have less than 5 tracks don't bother to make a map. Usually ocurrs when the device is turned on
# when connected via USB or just turned on for a short amount of time
	if ($numtracks -ge 5) {
	Write-Host "Number of tracks date $cdayh number of tracks is $numtracks dropping one after every $tottracks"
# Put out the track for that one day
		for ($cd = $i; $cd -lt ($i + $numtracks*3); $cd++) {
# Look for a track with lat / lon, elevation and time
			if ($tracks[$cd] -match 'trkpt lat="') {
				$tlat = $tracks[$cd]
				$tele = $tracks[$cd+1]
				$ttime = $tracks[$cd+2]
				[datetime]$dstring = (($ttime -split "<time>")[1] -split "</time>")[0]
				if ($dstring.Year -lt 2008) { $dstring = $dstring.AddDays(7168)}
# Check to ensure we have all the required data
				if (($tlat -match "lat=") -and ($tele -match "ele") -and ($ttime -match "time")) {
# Only put in every 'X'th track as calculated from totcount
				$trnum = ($cd - $i) / 3
#			Write-Host "trnum = $trnum totcount = $totcount cd $cd"
				if ($trnum -gt $totcount) {
# Create each HTML string Example:
# {'title':'1','content':'\<b\>8/10/2014\</b\>\<br\>781 ft\<br\>','lat':'58.03824','lng':'-5.06937'},
					[datetime]$dstring = (($ttime -split "<time>")[1] -split "</time>")[0]
# Check and see if the date is less than year 2008, if so then add 7168 days to get the correct date
					if ($dstring.Year -lt 2008) {$dstring = $dstring.AddDays(7168)}
					$dstringout = $dstring.ToString("MM/dd/yyyy")
# Get the elevation
					$estring = (($tele -split "<ele>")[1] -split "</ele>")[0]
# Convert to the appropriate unit of measure then convert back
					$enum1 = [decimal]$estring
					$enumber =  [math]::Round(($enum1 * $distsmall),2)
					$estring = [string]$enumber
					if ($old_alt -eq -1000) {$old_alt = [int]$estring}
# Get latitude and longitude
					$latstring = (($tlat -split '<trkpt lat="')[1] -split '"')[0]
					$lonstring = (($tlat -split 'lon="')[1] -split '"')[0]
# Claculate the delta distance from the last point
					$new_lat = [decimal]$latstring
					$new_lon = [decimal]$lonstring
					if ($old_lat -eq -1) {
						$old_lat = $new_lat
						$old_lon = $new_lon
						$old_time = $dstring
						}
					$oldpos = New-Object System.Device.Location.GeoCoordinate $old_lat,$old_lon
					$newpos = New-Object System.Device.Location.GeoCoordinate $new_lat,$new_lon
					$distance_sm = ($oldpos.GetDistanceTo($newpos)*$distsmall)
					$distance_lg = $distance_sm / $distlarge
					$edistance = $edistance + $distance_lg
					$dfstring = [string][math]::Round($edistance,3)

# Calculate speed here if wanted
					$speed = ""
					if ($the_speed -eq 'HELL YES') {
						$dstringout = $dstring.ToString("MM/dd/yyyy HH:mm:ss")
# PS> $a = New-Object System.Device.Location.GeoCoordinate 46.985861,-120.566050
# PS> $b = New-Object System.Device.Location.GeoCoordinate 46.998460,-120.581326
# PS> $a.GetDistanceTo($b)*3.2808399
						$dtime = ($dstring - $old_time).TotalHours
						$speed = $distance_lg / $dtime
						$ddistance = $ddistance + $distance_lg
						$destring = [decimal]$estring
#						Write-Host $ttcount $trnum $estring $daltitude $old_alt $average_alt_sum
						if (($ttcount % $average_alt) -eq 0) {
							$average_alt_sum = $average_alt_sum + $destring
							$destring = $average_alt_sum / $average_alt
							$average_alt_sum = 0
							if ($destring -gt $old_alt) {$daltitude = $daltitude + ($destring - $old_alt)}
							$old_alt = $destring
						}
						else {
							$average_alt_sum = $average_alt_sum + $destring
						}
# Round everything to a few digits of accuracy
						$dfspeed = [string][math]::Round($speed,2)
						$ddstring = [string][math]::Round($ddistance,2)
						$dastring = [string][math]::Round($daltitude,2)
						$tstring = [string]$ttcount
						$outstringh = "{'title':'" + $ttcount + "','content':'\<b\>" + $dstringout + "\</b\>\<br\>Altitude " + $estring + " " + $distsmalldesc + "\</b\>\<br\>" + "Speed " + $dfspeed + " Total Distance " + $ddstring + "\<br\>','lat':'" + $latstring + "','lng':'" + $lonstring + "'},"
						[string]$outstringc = $tstring + "," + $dstringout + " Altitude " + $distsmalldesc + " " + $estring + "," + $dfspeed + "," + $dfstring + "," + $ddstring + "," + $dastring + "," + $latstring + "," + $lonstring
						}
					else {
						[string]$outstringh = "{'title':'" + $ttcount + "','content':'\<b\>" + $dstringout + "\</b\>\<br\>Altitude: " + $estring + " " + $distsmalldesc + "\<br\>','lat':'" + $latstring + "','lng':'" + $lonstring + "'},"
						[string]$outstringc = [string]$ttcount + "," + $dstringout + " Altitude: " + $estring + " " + $distsmalldesc + "," + $dfstring + "," + $latstring + "," + $lonstring
						}
					[string]$outstringe = [string]$ttcount + "," + $dstringout + "," + $dfstring + "," + $estring + ","+ $latstring + "," + $lonstring
					$ttcount++
					$trackhtml += $outstringh
					$trackcsv += $outstringc
					$trackcsvelev += $outstringe
					$totcount = $totcount + $tottracks
#					Write-Host "trnum = $trnum totcount = $totcount Track $ttcount"
					$old_lat = $new_lat
					$old_lon = $new_lon
					$old_time = $dstring
				}
			}
			else
			{
				Write-Host "Found track that is incorrect $tlat $tele $ttime"
			}
		}
		}
# Write out sheets
		$fileout = "$cdayh.html"
		$filecsv = "$cdayh.csv"
		$filecsvelev = $cdayh + "_elev.csv"
		$thtml = $trackhtml.Count
		$tcsv = $trackcsv.Count
		Write-Host "For $fileout and $filecsv will output $thtml tracks and $tcsv tracks"
		$header | Out-File $fileout
		$trackhtml | Add-Content $fileout
		$trailer | Add-Content $fileout
		$csv_title | Add-Content $filecsv
		$trackcsv | Add-Content $filecsv
		$csvelev_title | Add-Content $filecsvelev
		$trackcsvelev | Add-Content $filecsvelev
	}
	else {	Write-Host "Skipping date $cdayh because number of tracks is $numtracks which is less than 5 data points"}
	$i = $i + ($numtracks*3) + 2
}

