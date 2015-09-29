Param(
    [Parameter(Mandatory=$true)] [string]$emailAddress,
    [Parameter(Mandatory=$false)] $password,
    [string]$count = 200,
    [string]$saveTo = $HOME + "\Desktop\out.csv"
)


if($password.Length -eq 0){
    $password = Read-Host -AsSecureString -Prompt "Password"
    $secure = $true
}

$r = Invoke-WebRequest https://www.myfordmobile.com/content/mfm/app/site/login.html

$sessId = $r.headers.'Set-Cookie'.ToString() -match "JSESSIONID=([0-9a-z\-]*)"

$form = $r.Forms[0]

if($secure) { 
    $plainPass = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
}
else { $plainPass = $password }

$form.Fields["inputEmailAddress"] = $emailAddress
$form.Fields["inputPassword"] = $plainPass


$q = Invoke-WebRequest https://myfordmobile.com/services/webLoginPS -method Post -Body ('{"PARAMS":{"emailaddress":"' + $emailAddress + '","password":"' + $plainPass + '","persistent":"0","ck":"1435667401589","apiLevel":"1"}}') -Headers @{"Cookie" = "JSESSIONID=$($matches[1])"} -ContentType "application/json"

if( ! ($q.Content -match '"authToken":"([a-z0-9\-]*)') ) { Write-Error "Returned data doesn't contain authToken" }
$authToken = $matches[1]

$URL = "https://phev.myfordmobile.com/services/webTripAndChargeLogDesktopPS"
$Content = "application/json"

$time = [Math]::ceiling((New-TimeSpan -start (get-date "01/01/1970") -end (get-date)).TotalMilliseconds)

$body = '{"PARAMS":{"SESSIONID":"' + $authToken + '","PAGENUMBER":1,"PAGESIZE":' + $count + ',"ck":"' + $time + '","apiLevel":"1"}}'

$z = Invoke-WebRequest -Uri $URL -WebSession $mfm -Method POST -Body $body -ContentType $content

$obj = ($z.Content | ConvertFrom-Json).getTripandChargelogQuery.log 

if($obj.count -eq $count){ Write-Warning "You requested $count items, and the service returned $count items, so there might be more data!  Increase -count parameter." }

$obj | Export-Csv $saveTo -NoTypeInformation

Write-Host "File saved to $saveTo"
