## Users who will recieve the email.
$digestUsers = @(
                    #"CitrixTeam@beaumont.edu",
                    "Zachary.Brozowski@beaumont.org"
                )
$EmailFrom = "Sysman <sysman@beaumont.org>"
$SmtpServer = "mail.beaumont.org"

$servers = get-content "\\ms.beaumont.edu\share\ServerTeam\Zach\Script\Get-ServerIP&Subnet\servers.txt"
$output = @()

$scriptPath = "\\ms.beaumont.edu\share\ServerTeam\Zach\Script\Get-ServerIP&Subnet"
$username = "MSWBH\citrixadmin"
$passwordFile = $scriptPath + "\citrixadmin_cred.txt"
$password = get-content $passwordFile | ConvertTo-SecureString
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password


$serversProcessed = 0

foreach ($server in $servers) {

    Write-Progress -Activity "Gathering Network Info from Servers"`
        -CurrentOperation "Current: $($server)"`
        -Status "$($servers.Count - $serversProcessed) remaining"`
        -PercentComplete (($serversProcessed/$servers.Count)*100)

    $networkInfo = Get-WmiObject Win32_NetworkAdapterConfiguration -Credential $creds -ComputerName $server | ? {$_.IPEnabled}

    $object = New-Object PSObject -Property @{
        Server     = $Server
        IP         = $networkInfo.IPAddress |Out-String
        Subnet     = $networkInfo.IPSubnet | Out-String
    }
    $output+= $object
    $serversProcessed++
}


$fileName = "CitrixIP&Subnet.csv"
$output | Select-Object Server, IP , Subnet |Export-Csv -NoTypeInformation $fileName -Force
$output | Select-Object Server, IP , Subnet |Export-Csv -NoTypeInformation -Force -Path "\\ms.beaumont.edu\share\ServerTeam\Zach\Script\Get-ServerIP&Subnet\Output.csv"



$style =  "<style>body{font-family: Arial, sans-serrif;font-size: 8pt;}"
$style += "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style += "TH{border: 1px solid black; background: #dddddd; padding: 5px;}"
$style += "TD{border: 1px solid black; padding: 3px;}"
$style += "</style>"

$masterHead    = "<html><head>$style</head><body>"
$masterHead   += "<table><tr>The following Servers are not using the correct subnet.</tr></table><table><tr><th>Server</th><th>IP Address</th><th>Subnet</th></tr>"
$masterBody    = ""
$footer        = "</table></body></html>"


foreach($item in $output){
    if($item.Subnet -match "255.255.255.0"){
        $masterBody += "<tr><td>" + $item.server +"</td><td>" + $item.IP +"</td><td>" + $item.Subnet + "</td>`n"
    }
}

$masterMessage  = $masterHead
$masterMessage += $masterBody
$masterMessage += $footer

$messageparams = @{
    to          = $digestUsers
    Subject     = "Citrix Subnet Report"
    body        = $masterMessage
    from        = $EmailFrom
    smtpserver  = $SmtpServer
    bodyashtml  = $true
    attachments = @( "$fileName")
}
#Send-MailMessage @messageparams -UseSsl
