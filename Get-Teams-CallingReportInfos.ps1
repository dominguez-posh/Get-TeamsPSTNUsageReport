$LocalCountry = "*Germany*"

$PriceListURI =  "https://bruellpublicblobs.blob.core.windows.net/pricelist/teamsDEPrices.xml"

$Pricing120min = 5.60
$Pricing1200min = 11.20
$PricingInternational = 22.50



$webClient = New-Object System.Net.WebClient 
$content = $webClient.DownloadString($PriceListURI)

$PriceList = [Management.Automation.PSSerializer]::Deserialize($content )



cls
Write-Host "Zuerst einen Report erstellen über Seite https://admin.teams.microsoft.com/analytics/reports. Hier auswählen 'PSTN and SMS (preview) usage' Dann eine Zeitrange angeben und rechts dann den Report als CSV herunterladen"

Write-Host

Write-Host "Im nächsten Schritt die Datei auswählen."
Pause


$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Documents (*.csv)|*.csv'
}
$null = $FileBrowser.ShowDialog()

Write-Host "Report Erstellung wird gestartet..."

$CallReport = Import-CSV $FileBrowser.FileName
$Users = ($CallReport | Sort-Object -Unique -Property UPN).UPN


$StartDate = [datetime]::ParseExact( ($CallReport| Select 'Start Time' | Sort-Object -Property 'Start Time')[0].'Start Time'.Split("T")[0],'yyyy-MM-dd',$null)
$EndDate = [datetime]::ParseExact( ($CallReport | Select 'Start Time' | Sort-Object -Property 'Start Time')[($CallReport  | Select 'Start Time' | Sort-Object -Property 'Start Time').Count -1].'Start Time'.Split("T")[0],'yyyy-MM-dd',$null)
$TotalDays = ($EndDate - $StartDate).TotalDays


Write-Host "Getting Calling Duration for all Countries..."
$AllOutboundCalls = $CallReport | ? "Call Direction" -like "Outbound" 
$CountryDuration = @()
foreach($Country in (($AllOutboundCalls | Sort-Object -Unique  "Destination Dialed" | ? "External Country" -ne "??" | select "Destination Dialed")."Destination Dialed")){

    
    $DurationSeconds = ($AllOutboundCalls | ? "Destination Dialed" -like $Country | ? "Duration Seconds" -ne "0")."Duration Seconds" 
    $DurationMinutes = ($DurationSeconds  | % {[Math]::Truncate($_ / 60)+1} | Measure-Object -Sum).Sum

    if ($Country -eq ""){$Country = "Unknown / Other"}

    
    $CountryDuration += [PSCustomObject]@{
        Country     = $Country
        CallingMinutes = $DurationMinutes
        PricePAYG = 0
        PriceSum = 0
    }
    


}

$OverAllTime = ($CountryDuration | Measure-Object  -Property CallingMinutes -sum).Sum

$CountryDuration = $CountryDuration | Sort-Object -Property CallingMinutes -Descending





Write-Host "Analyzing Price-List"
$Prices = @()

$PriceList  | % { 

$DestinationName =  $_."Ziel"

$Price = $_.'Minutentarif in EUR (Abrechnung in Zehnteln einer Minute)'

if($Price -like "*€*"){$Price = $Price.Replace("€","").Replace(',','')}



$Vorwahlen = $_."Vorwahl(en)".Split(",")

foreach($Vorwahl in $Vorwahlen){
        
    $Price = $Price.Replace(" Cent","")


     $Prices += [PSCustomObject]@{
        Vorwahl     = $Vorwahl
        DestinationName = $DestinationName
        Price = $Price


      }




}


}



   $prices = $Prices | ? Vorwahl -ne "" | ? Vorwahl -ne $Null | Sort-Object { $_.Vorwahl.length } 


Write-Host "Getting Prices"
$R=0
 foreach($Country in $CountryDuration){
    Write-Host "Processing " $Country.Country
    $Prefix = ($AllOutboundCalls | ? {$_."Destination Dialed" -EQ $Country.Country}) | select -First 1

    $DestNumber = $Prefix."Destination Number".Replace("*","").Replace("+","")


        $MatchingPrices = @()
        foreach($Price in $Prices){
            
            
            $SubstringLength = $Price.Vorwahl.Length

            if($SubstringLength -lt $DestNumber.Length){
                $Compare = $DestNumber.Substring(0, $SubstringLength)
            }
            if($Price.Vorwahl -eq $Compare){$MatchingPrices+=$Price}   
        }

    $Count = $MatchingPrices.Count
    if($Count -eq $Null){$Count = 1}




    $Price = [Decimal]$MatchingPrices[$Count-1].Price.Replace(",",".")


    $CountryDuration[$R].PricePAYG = $Price

    $CountryDuration[$R].PriceSum = [Math]::Round($CountryDuration[$R].CallingMinutes * $CountryDuration[$R].PricePAYG / 100,2)
    
    $R ++

} 

$UserCount = $users.Count




$NationalCallingMinutes =  ($CountryDuration | ? Country -like $LocalCountry | ? PricePAYG -NE 0 | Measure-Object -Property CallingMinutes -sum).sum
$InterNationalCallingMinutes = ($CountryDuration | ? Country -NotLike $LocalCountry  | ? PricePAYG -NE 0 | Measure-Object -Property CallingMinutes -sum).sum
$TollFreeCallingMinutes = ($CountryDuration  | ? PricePAYG -EQ 0 | Measure-Object -Property CallingMinutes -sum).sum


$UsersNational = ($AllOutboundCalls | ? "Domestic/International" -eq "Domestic" | Sort-Object -Property UPN -Unique | select upn).upn

$UsersInterNational = ($AllOutboundCalls | ? "Domestic/International" -eq "International" | Sort-Object -Property UPN -Unique | select upn).upn


$UserCountNational = $UsersNational.Count

$UserCountInterNational = $UsersINterNational.Count


$UsersNotCallingNational = (Compare-Object $Users $UsersNational).inputobject

$UsersNotCallingInternational = (Compare-Object $Users $UsersInterNational).inputobject

$UsersOnlyCallingNational = (Compare-Object $UsersNational  $UsersNotCallingInterNational -IncludeEqual | ? SideIndicator -eq "==").inputobject


$UsersNotCallingOutgoing = (Compare-Object $UsersNotCallingNational  $UsersNotCallingInternational -IncludeEqual | ? SideIndicator -eq "==").inputobject





#Report

cls
Write-Host "Betrachtung Zeitraum von" $StartDate "Bis" $EndDate -BackgroundColor Gray -ForegroundColor Black

Write-Host

Write-Host "Durchschnittliche ausgehende nationale Gesprächsminuten pro Tag: " ([math]::Round($NationalCallingMinutes / $TotalDays,2))
Write-Host "Durchschnittliche ausgehende nationale Gesprächsminuten pro Monat: " -NoNewline 
Write-Host ([math]::Round($NationalCallingMinutes / $TotalDays * 30.5,2)) -BackgroundColor White -ForegroundColor Black
Write-Host "Durchschnittliche ausgehende monatliche nationale Gesprächsminuten pro Benutzer: " -NoNewline 
Write-Host ([math]::Round($NationalCallingMinutes / $TotalDays * 30.5 / $UserCount,2)) -BackgroundColor White -ForegroundColor Black
Write-Host "Anzahl der Benutzer, die ausgehende nationale Gespräche geführt haben: " $UserCountNational

Write-Host


Write-Host "Durchschnittliche ausgehende internationale Gesprächsminuten pro Tag: " ([math]::Round($InterNationalCallingMinutes / $TotalDays,2))
Write-Host "Durchschnittliche ausgehende internationale Gesprächsminuten pro Monat: " -NoNewline 
Write-Host ([math]::Round($InterNationalCallingMinutes / $TotalDays * 30.5,2)) -BackgroundColor White -ForegroundColor Black
Write-Host "Durchschnittliche ausgehende monatliche internationale Gesprächsminuten pro Benutzer: " -NoNewline 
Write-Host ([math]::Round($InterNationalCallingMinutes / $TotalDays * 30.5 / $UserCount,2)) -BackgroundColor White -ForegroundColor Black
Write-Host "Anzahl der Benutzer, die ausgehende internationale Gespräche geführt haben: " $UserCountINterNational

Write-Host

Write-Host "Insgesamt haben" $Users.Count "Benutzer Ein und ausgehend telefoniert"

Write-Host

Write-Host "Davon haben folgende Benutzer nie ausgehend telefoniert und benötigen keinen Calling-Plan: "
Write-Host $UsersNotCallingOutgoing -Separator "," -BackgroundColor White -ForegroundColor Black

Write-Host

Write-Host "Und Folgende Benutzer haben nur national und nicht international telefoniert und benötigen keinen internationalen Calling-Plan:"
Write-Host $UsersOnlyCallingNational -Separator ", "  -BackgroundColor White -ForegroundColor Black

Write-Host
Write-Host


Write-Host "Folgende Kosten würden mit PAYG nur für die Inlandsgespräche pro Monat entstehen::" ([Math]::Round(($CountryDuration | ? Country -like $LocalCountry | Measure-Object -Property PriceSum -sum).Sum / $TotalDays * 30,2)) "Euro"


Write-Host "Folgende Kosten würden mit PAYG nur für die Auslandsgespräche pro Monat entstehen::" ([Math]::Round(($CountryDuration | ? Country -notlike $LocalCountry | Measure-Object -Property PriceSum -sum).Sum / $TotalDays * 30,2)) "Euro"


Write-Host
$AverageMinutePriceNational = ($CountryDuration | ? Country -like $LocalCountry | Measure-Object -sum -Property PriceSum).sum / ($CountryDuration | ? Country -like $LocalCountry | Measure-Object -sum -Property CallingMinutes).sum
$AverageMinutePriceNational = [Math]::Round($AverageMinutePriceNational * 100,2)
$AverageMinutePriceInterNational = ($CountryDuration | ? Country -notlike $LocalCountry | Measure-Object -sum -Property PriceSum).sum / ($CountryDuration | ? Country -notlike $LocalCountry | Measure-Object -sum -Property CallingMinutes).sum 
$AverageMinutePriceInterNational  = [Math]::Round($AverageMinutePriceInterNational  * 100,2)

$AverageMinutePriceOverAll = ($CountryDuration | Measure-Object -sum -Property PriceSum).sum / ($CountryDuration | Measure-Object -sum -Property CallingMinutes).sum
$AverageMinutePriceOverAll = [Math]::Round($AverageMinutePriceOverAll * 100,2)



Write-Host "Die durchschnittlichen Kosten für eine nationale Minute betrug" $AverageMinutePriceNational "Cent pro Minute"
Write-Host "Die durchschnittlichen Kosten für eine internationale Minute betrug" $AverageMinutePriceInterNational "Cent pro Minute"


Write-Host

$RestMinutes120 = [math]::Round((((([math]::Round($NationalCallingMinutes / $TotalDays * 30.5 / $UserCount,2))) ) -120),2) * $UserCount
if($RestMinutes120 -le 0){$RestMinutes120 = 0}

Write-Host "Bei Buchung eines 120 Minuten-Paketes für alle Benutzer entstehen basierend auf den analysierten Zeitraum durchschnittlich folgende Kosten:"
Write-Host "in einem Durchschnittlichen Monat, wird das Minutenkontingent um" $RestMinutes120 "Minuten überschritten."

Write-Host "Kosten für das Minutenpaket: " ($Pricing120min * $UserCount) "Euro"
Write-Host "Die zusätzlichen nationalen Minuten kosten pro Monat" ([Math]::Round(($RestMinutes120 * $AverageMinutePriceNational / 100),2)) "Euro"
Write-Host "Dazu kommen die Kosten für die internationalen Minuten: " ([Math]::Round(($CountryDuration | ? Country -notlike $LocalCountry  | measure-Object -sum -Property PriceSum).sum / 30,2)) "Euro"

$OverallCosts120min = (($Pricing120min * $UserCount) + ($RestMinutes120 * $AverageMinutePriceNational / 100) + ([Math]::Round(($CountryDuration | ? Country -notlike $LocalCountry  | measure-Object -sum -Property PriceSum).sum / 30,2))) 
$OverallCosts120min = [Math]::Round($OverallCosts120min,2)

Write-Host "Gesamtkosten bei einem gebuchten 120-Min Paket pro Monat: " $OverallCosts120min "Euro" -BackgroundColor White -ForegroundColor Black

Write-Host

$RestMinutes1200 = [math]::Round((((([math]::Round($NationalCallingMinutes / $TotalDays * 30.5 / $UserCount,2))) ) -1200),2) * $UserCount
if($RestMinutes1200 -le 0){$RestMinutes1200 = 0}

Write-Host "Bei Buchung eines 1200 Minuten-Paketes für alle Benutzer entstehen basierend auf den analysierten Zeitraum durchschnittlich folgende Kosten:"
Write-Host "in einem Durchschnittlichen Monat, wird das Minutenkontingent um" $RestMinutes1200 "Minuten überschritten."

Write-Host "Kosten für das Minutenpaket: " ($Pricing1200min * $UserCount) "Euro"
Write-Host "Die zusätzlichen nationalen Minuten kosten pro Monat" ([Math]::Round(($RestMinutes1200 * $AverageMinutePriceNational / 100),2)) "Euro"
Write-Host "Dazu kommen die Kosten für die internationalen Minuten: " ([Math]::Round(($CountryDuration | ? Country -notlike $LocalCountry  | measure-Object -sum -Property PriceSum).sum / 30,2)) "Euro"

$OverallCosts1200min = (($Pricing1200min * $UserCount) + ($RestMinutes120 * $AverageMinutePriceNational / 100) + ([Math]::Round(($CountryDuration | ? Country -notlike $LocalCountry  | measure-Object -sum -Property PriceSum).sum / 30,2))) 
$OverallCosts1200min = [Math]::Round($OverallCosts1200min,2)

Write-Host "Gesamtkosten bei einem gebuchten 1200-Min Paket pro Monat: " $OverallCosts1200min "Euro" -BackgroundColor White -ForegroundColor Black


Write-Host

$RestMinutesInternational = [math]::Round((((([math]::Round((($NationalCallingMinutes / $TotalDays * 30.5 / $UserCount) + ($InternationalCallingMinutes / $TotalDays * 30.5 / $UserCount)),2))) ) -1200),2) * $UserCount
if($RestMinutesInternational -le 0){$RestMinutesInternational = 0}

Write-Host "Bei Buchung eines Internationalen Minuten-Paketes für alle Benutzer entstehen basierend auf den analysierten Zeitraum durchschnittlich folgende Kosten:"
Write-Host "in einem Durchschnittlichen Monat, wird das Minutenkontingent um" $RestMinutesInternational "Minuten überschritten."

Write-Host "Kosten für das Minutenpaket: " ($PricingInternational * $UserCount) "Euro"
Write-Host "Die zusätzlichen nationalen Minuten kosten pro Monat" ([Math]::Round(($RestMinutesInternational * $AverageMinutePriceOverAll / 100),2)) "Euro"

$OverallCostsInternational = (($PricingInternational * $UserCount) + ([Math]::Round(($RestMinutesInternational * $AverageMinutePriceOverAll / 100),2)))
$OverallCostsInternational = [Math]::Round($OverallCostsInternational,2)

Write-Host "Gesamtkosten bei einem gebuchten Internationalen Paket pro Monat: " $OverallCostsInternational  "Euro" -BackgroundColor White -ForegroundColor Black



Write-Host
Write-Host

Write-Host "Gesamtkosten pro Monat, bei Minutengenauer Abrechnung: "  ([Math]::Round(($CountryDuration | Measure-Object -Property PriceSum -sum).Sum / $TotalDays * 30,2)) "Euro" -BackgroundColor White -ForegroundColor Black



Write-Host

Write-Host "Die Minuten sind wie folgt zusammengestellt:"

Write-Host
 $CountryDuration = $CountryDuration | Sort-Object -Property PriceSum -Descending

$CountryDuration | Format-Table | Out-String|% {Write-Host $_}






##Verteilung der Benutzer

$Output = @()

foreach($User in $Users){

    
    $Calls = $AllOutboundCalls | ? UPN -eq $User

    $CallSum = $Calls.Count ; if(-not $Callsum){$CallSum = 1}

    $InternationalCalls = $Calls | ? "Domestic/International" -like "International"
    $InternationalCallSum = $InternationalCalls.Count ; if(-not $InternationalCallSum ){$InternationalCallSum  = 1}
    if(-not $InternationalCalls){$InternationalCallSum  = 0}


    $nationalCalls = $Calls | ? "Domestic/International" -like "Domestic"
    $nationalCallSum = $nationalCalls.Count ; if(-not $nationalCallSum ){$nationalCallSum  = 1}
    if(-not $nationalCalls){$nationalCallSum  = 0}

     
    $CallingSecondsNational = ($nationalCalls | Measure-Object -Property "Duration Seconds" -Sum).sum

    $CallingSecondsInterNational = ($InternationalCalls |  Measure-Object -Property "Duration Seconds" -Sum).sum

    $CallingSecondsSharedCosts = ($Calls | ? "Destination Dialed" -like "*Shared*"  | Measure-Object -Property "Duration Seconds" -Sum).sum
    
    $Callingseconds0180 = ($Calls | ? "Destination Dialed" -like "*0180*"  | Measure-Object -Property "Duration Seconds" -Sum).sum

    $CallingSecondsTollFree = ($Calls | ? "Destination Dialed" -like "*Toll Free*"  | Measure-Object -Property "Duration Seconds" -Sum).sum
    

    $Output += [PSCustomObject]@{
            User     = $User
            CallSum = $CallSum
            InternationalCallSum = $InternationalCallSum 
            NationalCallSum = $nationalCallSum
            CallingSecondsNational = $CallingSecondsNational
            CallingSecondsInterNational = $CallingSecondsInterNational
            CallingSecondsSharedCosts = $CallingSecondsSharedCosts
            Callingseconds0180 = $Callingseconds0180
            CallingSecondsTollFree = $CallingSecondsTollFree
    }


}

Write-Host

Write-Host "Die Minuten sind wie folgt zusammengestellt"

$Output| Sort-Object -Property CallSum -Descending  | Format-Table | Out-String|% {Write-Host $_}





Pause
