#################################################################################
#
#   
#   Skrypt dodaje stopkę do skrzynek pocztowych użytkowników 
#   Następnie wstaw własne stopki w tekście w zmiennej $sygnatura_text 
#   i w html'u w zmiennej $sygnatura_html
#   Powodzenia
#
#################################################################################

$moja_lokalizacja = Get-Location
$sciezka = $moja_lokalizacja.Path
Set-Location $sciezka

if ((Test-Path -Path "$HOME\credentials.xmls") -eq $true) {
    $credentialsPath = Join-Path -Path $HOME -ChildPath credentials.xmls;
    $credentials = Import-CliXml $credentialsPath;
}
else {
    $credentials = Get-Credential;
    $credentialsPath = Join-Path -Path . -ChildPath credentials.xmls;
    $credentials | Export-CliXml $credentialsPath;
}
Connect-ExchangeOnline -Credential $credentials

$podpisy = Import-Csv .\podpisy.csv
$numer = 0

Write-Host "`n`nDodaję podpisy dla:"
foreach ($user in $podpisy) {
    $numer = $numer + 1
    $sygnatura_text = "e-mail: " + $user.Identity + "@sp5.szczecin.pl`n-------------------`nInformujemy, że Administratorem danych osobowych jest Szkoła Podstawowa nr 5 im. Henryka Sienkiewicza w Szczecinie. Szczegółowe informacje dotyczące przetwarzania danych, osobowych w Szkole Podstawowej nr 5 znajdują się na stronie internetowej pod adresem: www.sp5.szczecin.pl"
    $sygnatura_html = "<html><body><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div style=`"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)`"><span style=`"color:rgb(50,49,48); font-size:12pt`">e-mail: " + $user.Identity + "@sp5.szczecin.pl</span><br></div><div style=`"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)`"><span style=`"color:rgb(50,49,48); font-size:14.6667px; text-align:start; background-color:rgb(255,255,255); display:inline!important`">-------------------</span></div><div style=`"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)`"><span style=`"color:rgb(50,49,48); font-size:9pt; text-align:start; background-color:rgb(255,255,255); display:inline!important`">Informujemy, że Administratorem danych osobowych jest Szkoła Podstawowa nr 5 im. Henryka Sienkiewicza w Szczecinie. Szczegółowe informacje dotyczące przetwarzania danych, osobowych w Szkole Podstawowej nr 5 znajdują się na stronie internetowej pod adresem: <a href=`"http://www.sp5.szczecin.pl`" title=`"www.sp5.szczecin.pl`">www.sp5.szczecin.pl</a></span><br></div></body></html>"
    
    if (([bool]$user.SignatureText) -eq $false) {
        Set-MailboxMessageConfiguration $user.Identity -SignatureHtml $sygnatura_html -SignatureText $sygnatura_text 
    }

    Set-MailboxMessageConfiguration $user.Identity -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true
    
    [pscustomObject]@{
        Numer    = $numer ;
        Identity = $user.Identity ;
        Stan     = [bool]$user.SignatureText
    }
}
