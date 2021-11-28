Param(
    [Parameter(Mandatory)]
    [String]
    $Name,
	
	[Parameter(Mandatory)]
    [ValidateRange(1,31)]
    [Int]
    $DayOfBirth,

    [Parameter(Mandatory)]
    [ValidateRange(1,12)]
    [Int]
    $MonthOfBirth
)

$now = Get-Date
$today = $now.Date
$year = $today.Year

function ConvertTo-Words
{
	$objSpeak = New-Object -ComObject SAPI.SpVoice
	$objSpeak.Speak($words)
}

if ($today.Day -eq $DayOfBirth -and $today.Month -eq $MonthOfBirth)
{
	$words = "Happy Birthday $Name!"
	$year = $null
}

elseif ($today.Day -gt $DayOfBirth -and $today.Month -eq $MonthOfBirth)
{
	$words = "You've just had a birthday $Name, behave and ask me next month!"
	$year = $null
}

elseif ($today.Month -gt $MonthOfBirth)
{
	$year++
}

if ($year)
{
	$birthday = Get-Date $DayOfBirth/$MonthOfBirth/$year
	$sleeps = ($birthday - $today).Days
	$sleeps++
}

if (-not $words)
{
	$words = "There are $sleeps sleeps until $name's birthday, Hooray!"
}

ConvertTo-Words $words