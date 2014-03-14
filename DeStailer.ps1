# Finds people who tried to send and email to a "Stale" contact
# Takes the recipient address and converts it into a X500 address which can be added to the end user's exchange object
# This will make future emails go through normally without deleting the cached contact

# How many days back to check
$days = 2

$server = Get-ExchangeServer * | where{$_.ServerRole -eq "HubTransport"}
$email = $server | %{Get-MessageTrackingLog -server $_.identity -EventId FAIL -Start (Get-Date).AddDays(-1 * $days)} | where{$_.recipients -match "IMCEAEX"}
$stale = $email | sort recipients -Unique | select Recipients | %{$_.recipients | %{$_}} | where{$_ -match "IMCEAEX"}

$stale | %{
    $ADDR = $_
    $REPL= @(@("_","/"), @("\+20"," "), @("\+28","("), @("\+29",")"), @("\+2C",","), @("\+5F", "_" ), @("\+40", "@" ), @("\+2E", "." )) 
    $REPL | FOREACH { $ADDR= $ADDR -REPLACE $_[0], $_[1] } 
    $ADDR= "X500:$ADDR" -REPLACE "IMCEAEX-","" -REPLACE "@.*$", "" 
    Write-Host $ADDR
}