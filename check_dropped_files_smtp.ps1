Param (
	[string]$Path = "C:\inetpub\mailroot\Drop\",
	[string]$SMTPServer = "smtp.sendgrid.net",
	[string]$From = "postmaster@google.com",
	[string]$To = "alerts@google.com",
	[string]$Subject = "PROD - New file(s) detected in drop folder"
	)

## Creates an empty table
$table = New-Object System.Data.DataTable "Table"
$col1 = New-Object system.Data.DataColumn FileName,([string])
$col2 = New-Object system.Data.DataColumn Email,([string])
$col3 = New-Object system.Data.DataColumn Subject,([string])
$col4 = New-Object system.Data.DataColumn Message,([string])
$table.Columns.Add($col1)
$table.Columns.Add($col2)
$table.Columns.Add($col3)
$table.Columns.Add($col4)

$table2 = New-Object System.Data.DataTable "CountTable"
$col91 = New-Object system.Data.DataColumn Count,([string])
$col92 = New-Object system.Data.DataColumn Subject,([string])
$table2.Columns.Add($col91)
$table2.Columns.Add($col92)

## Initialize count and define filename based on today's date
$count = 0
$csvname = Get-Date -Format MM-dd-yyyy
$csvname +='-DroppedEmails'

## Define credentials for sendgrid. The password is saved in secure string format as endecauser, meaning this script will send email only if run as endecauser.
$pw = Get-Content C:\Temp\SMTP_scripts\Check_BAD_files_SMTP_sendgrid_creds.txt | ConvertTo-SecureString
$creds = New-Object System.Management.Automation.PSCredential "username",$pw
$SMTPMessage = @{
    To = $To
    From = $From
	Subject = "$Subject at $Path"
    Smtpserver = $SMTPServer
    Credential = $creds
}

$Files = Get-ChildItem $Path | Where { $_.LastWriteTime -ge [datetime]::Now.AddHours(-24) -and $_.Extension -eq ".EML" } 
If ($Files)
{	$SMTPBody = "`nFound new file(s) in Badmail folder, here is a summary:`n`n"
    foreach($File in $Files){
    $email = sls "Final-Recipient" $File | select -ExpandProperty line | %{$_.split(";")[1]}
    $Subject = sls "Subject:" $File | select -ExpandProperty line | select -Last 1 | %{$_.split(":")[1]}
    $diagnosticscode = sls "Diagnostic-Code:" $File | select -ExpandProperty line | %{$_.split(";")[1]}
    $row = $table.NewRow()
    $row.FileName = $File
    $row.Email = $email
    $row.Subject = $Subject.Trim()
    $row.Message = $diagnosticscode
    $table.Rows.Add($row)
    #$SMTPBody += "$File `t $email `t $diagnosticscode `n"
    $count +=1
    }
    #$File | ForEach { $SMTPBody += "$($_.FullName)`t Potatoes`n" }

<#  THIS SECTION WHEN ENABLED CAN GENERATE A HTML FORMATTED EMAIL
    $html = "<style>
    TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
    TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
    .odd  { background-color:#ffffff; }
    .even { background-color:#dddddd; }
    </style><table><tr><th>FileName</th><th>Email</th><th>Subject</th><th>Diagnostics Message</th></tr>"
    foreach ($row in $table.Rows)
    { 
        $html += "<tr><td>" + $row.FileName + "</td><td>" + $row.Email + "</td><td>" + $row.Subject + "</td><td>" + $row.Message + "</td></tr>"
    }
    $html += "</table>"
    Send-MailMessage @SMTPMessage -Body $html -BodyAsHtml
#>    
    $table.Rows | Export-Csv -NoTypeInformation -Path D:\$csvname.csv
    $t = $table.Rows | Group-Object Subject | select -Property Count, Name | Sort-Object Count -Descending
    $html3 = "<style>
    TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
    TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
    .odd  { background-color:#ffffff; }
    .even { background-color:#dddddd; }
    </style><table><tr><th>Count</th><th>Subject</th></tr>"
    foreach ($r in $t)
    { 
        Write-Output $r.Count
        $row2 = $table2.NewRow()
        $row2.Count = $r.Count
        $row2.Subject = $r.Name
        $table2.Rows.Add($row2)
        $html3 += "<tr><td>" + $row2.Count + "</td><td>" + $row2.Subject + "</td></tr>"
    }
    $html3 += "</table>"
    Send-MailMessage @SMTPMessage -Body "Found <b>$count</b> dropped emails on SMTP server, Here is the breakdown: <br><br> $html3" -BodyAsHtml -Attachments "D:\$csvname.csv"

}

else {
    $html = "///\\\`nNo new files in past 24 hours."
    Send-MailMessage @SMTPMessage -Body $html -BodyAsHtml

}
