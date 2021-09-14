#################################################################################
#
#   
#   Skrypt dodaje stopkę do skrzynek pocztowych użytkowników 
#   
#   
#
#################################################################################

$sciezka = 'C:\Users\Maciej\Documents\Programowanie\Skrypt\Podpis w Outlooku'
Set-Location $sciezka

# if ((Test-Path -Path "$HOME\credentials.xmls") -eq $true) {
#     $credentialsPath = Join-Path -Path $HOME -ChildPath credentials.xmls;
#     $credentials = Import-CliXml $credentialsPath;
# }
# else {
#     $credentials = Get-Credential;
#     $credentialsPath = Join-Path -Path . -ChildPath credentials.xmls;
#     $credentials | Export-CliXml $credentialsPath;
# }
# Connect-ExchangeOnline -Credential $credentials

$podpisy = Import-Csv .\podpisy_test_mini.csv
$sygnatura_text = "e-mail: " + $user.UserPrincipalName + "`n-------------------`nInformujemy, że Administratorem danych osobowych jest Szkoła Podstawowa nr 5 im. Henryka Sienkiewicza w Szczecinie. Szczegółowe informacje dotyczące przetwarzania danych, osobowych w Szkole Podstawowej nr 5 znajdują się na stronie internetowej pod adresem: www.sp5.szczecin.pl"
$sygnatura_html = "<html><body><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div style=`"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)`"><span style=`"color:rgb(50,49,48); font-size:12pt`">e-mail: " + $user.UserPrincipalName + "</span><br></div><div style=`"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)`"><span style=`"color:rgb(50,49,48); font-size:14.6667px; text-align:start; background-color:rgb(255,255,255); display:inline!important`">-------------------</span></div><div style=`"font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)`"><span style=`"color:rgb(50,49,48); font-size:9pt; text-align:start; background-color:rgb(255,255,255); display:inline!important`">Informujemy, że Administratorem danych osobowych jest Szkoła Podstawowa nr 5 im. Henryka Sienkiewicza w Szczecinie. Szczegółowe informacje dotyczące przetwarzania danych, osobowych w Szkole Podstawowej nr 5 znajdują się na stronie internetowej pod adresem:<a href=`"http://www.sp5.szczecin.pl`" title=`"www.sp5.szczecin.pl`">www.sp5.szczecin.pl</a></span><br></div></body></html>"
$numer = 0

Write-Host "Dodaję podpisy dla:"
foreach ($user in $podpisy) {
    $numer = $numer + 1
    Set-MailboxMessageConfiguration $user.Identity -SignatureHtml $sygnatura_html -SignatureText $sygnatura_text -AutoAddSignature $true -AutoAddSignatureOnMobile $true -AutoAddSignatureOnReply $true

    [pscustomObject]@{
        Numer    = $numer;
        Imię     = $user.FirstName;
        Nazwisko = $user.LastName
    }
}
