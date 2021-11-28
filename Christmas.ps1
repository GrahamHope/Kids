function ConvertTo-Words
{
	$objSpeak = New-Object -ComObject SAPI.SpVoice
	$objSpeak.Speak($words)
}

$now = Get-Date
$today = $now.Date
$year = $today.Year
$Xmas = Get-Date 25/12/$year
$sleeps = ($xmas - $today).Days
$sleeps++
$words = "There are $sleeps sleeps until Christmas"
ConvertTo-Words $words | Out-Null