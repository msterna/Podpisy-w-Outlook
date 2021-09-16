#################################################################################
#
#   Skrypt pobiera listę użytkowników z MSO 365 a następnie pobiera do pliku 
#   .\podpisy.csv ustawienia skrzynek pocztowych użytkowników wraz z ich
#   sygnaturami. 
#   Najpierw ustaw zmienną $sciezka
#
#################################################################################

$sciezka = 'C:\Users\Maciej\Documents\Programowanie\Skrypt\Podpis w Outlooku'
Set-Location $sciezka
if ((Test-Path -Path "$sciezka\podpisy.csv") -eq $true) {
    Remove-Item -Path "$sciezka\podpisy.csv"
}

if ((Test-Path -Path "$HOME\credentials.xmls") -eq $true) {
    $credentialsPath = Join-Path -Path $HOME -ChildPath credentials.xmls;
    $credentials = Import-CliXml $credentialsPath;
}
else {
    $credentials = Get-Credential;
    $credentialsPath = Join-Path -Path . -ChildPath credentials.xmls;
    $credentials | Export-CliXml $credentialsPath;
}
# Import-Module AzureADPreview
# Connect-AzureAD -Credential $credentials
Connect-MSOLService -Credential $credentials
Connect-ExchangeOnline -Credential $credentials
# Get-MsolUser -All | Export-Csv Users.csv -Encoding UTF8
$users = Get-MsolUser -All
$numer = 0

Write-Host "Pobieram podpisy dla:"

foreach ($user in $users) {
    $numer = $numer + 1
    Get-MailboxMessageConfiguration $user.UserPrincipalName | Export-Csv -Append -NoTypeInformation -Path .\podpisy.csv -Encoding UTF8

    [pscustomobject]@{
        Numer    = $numer;
        Imię     = $user.FirstName;
        Nazwisko = $user.LastName
    }
}
