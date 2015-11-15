<#
.Synopsis
   Sends invitations by email
.DESCRIPTION
   Sends invitations by email to a provided hashtable of recipients.
   You must define the recipients, server, port, credentials, from.
   The idea is not to run this over and over again for every recipient
   but to send a single email and include all attendees on the invite.
.EXAMPLE

   $Recipients = 
    @{
        "Lewis Roberts" = "lewis@lewisroberts.com";
        "Kym Jones" = "miss.kr.jones@gmail.com";
    }
    Send-WasteCollectionInvitations -To $Recipients `
                                    -Server 'smtp.gmail.com' `
                                    -Port 587 `
                                    -From 'John Doe <john.doe@gmail.com>' `
                                    -Credentials $SenderCredentials `
                                    -Subject 'Waste Collection' `
                                    -Body $Body
#>
Function Send-WasteCollectionInvitations
{
    Param
    (
        [Parameter(Mandatory=$true)]
        [hashtable]$To,

        [Parameter(Mandatory=$true)]
        [string]$Server,

        [Parameter(Mandatory=$true)]
        [int]$Port,

        [Parameter(Mandatory=$true)]
        [string]$From,

        [Parameter(Mandatory=$true)]
        [pscredential]$Credentials,

        [Parameter(Mandatory=$true)]
        [string]$Subject,

        [Parameter(Mandatory=$true)]
        [string]$Body

    )
    ########
    # SMTP #
    ########
    $Smtp = New-Object System.Net.Mail.SmtpClient
    $Smtp.Credentials = $Credentials

    $Smtp.Port = $Port
    $Smtp.Host = $Server
    $Smtp.EnableSsl = $true

    ########
    # MAIL #
    ########
    $Mail = New-Object System.Net.Mail.MailMessage

    $Mail.From = New-Object System.Net.Mail.MailAddress($From)
    $Mail.Subject = $Subject

    Foreach ($Recipient in $To.GetEnumerator()) {
        $Mail.To.Add($($Recipient.Value))
    }

    ########
    # BODY #
    ########

    $TodayUTC = (Get-Date).ToUniversalTime()
    $Description = ($Subject -replace "Waste Collection - ").ToLower()

    # Add the HTML message body to the email from the injected $Body.
    $Mail.AlternateViews.Add([Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/html'))
    
    # Build the invitation with StringBuilder
    $s = New-Object System.Text.StringBuilder
    [void]$s.AppendLine('BEGIN:VCALENDAR')
    [void]$s.AppendLine("PRODID:-//lewisroberts//EN")
    #[void]$s.AppendLine("CALSCALE:GREGORIAN")
    [void]$s.AppendLine("VERSION:2.0")
    [void]$s.AppendLine("METHOD:REQUEST")
    [void]$s.AppendLine("BEGIN:VEVENT")
    #[void]$s.AppendLine("TRANSP:TRANSPARENT")
    [void]$s.AppendLine([String]::Format("DTSTART:{0:yyyyMMddT190000Z}", $TodayUTC))
    [void]$s.AppendLine([String]::Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", $TodayUTC.AddMinutes(-1)))
    [void]$s.AppendLine([String]::Format("DTEND:{0:yyyyMMddT200000Z}", $TodayUTC))
    [void]$s.AppendLine("LOCATION:9 Eastgate Road, Holmes Chapel, Cheshire")
    [void]$s.AppendLine([String]::Format("UID:{0}", [Guid]::NewGuid()))
    [void]$s.AppendLine("ORGANIZER;CN=`""+$Mail.From.DisplayName+"`":MAILTO:"+$Mail.From.Address)
    Foreach ($Recipient in $To.GetEnumerator()) {
        [void]$s.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;CN="+$($Recipient.Name)+":mailto:"+$($Recipient.Value)+"")
    }
    #[void]$s.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=Waste Collections;X-NUM-GUESTS=0:mailto:bounce@lewisroberts.com")
    [void]$s.AppendLine("SUMMARY;ENCODING=QUOTED-PRINTABLE:$Subject")
    [void]$s.AppendLine("DESCRIPTION;ENCODING=QUOTED-PRINTABLE:This is a reminder for tomorrow's waste collection. The $Description must be placed at the kerbside by 6am.")
    [void]$s.AppendLine("SEQUENCE:0")
    [void]$s.AppendLine("STATUS:CONFIRMED")
    #[void]$s.AppendLine("PRIORITY:3")
    #[void]$s.AppendLine("CLASS:PUBLIC")
    [void]$s.AppendLine("BEGIN:VALARM")
    [void]$s.AppendLine("TRIGGER;RELATED=START:-PT00H15M00S")
    [void]$s.AppendLine("ACTION:DISPLAY")
    [void]$s.AppendLine("DESCRIPTION:Reminder")
    [void]$s.AppendLine("END:VALARM")
    [void]$s.AppendLine("END:VEVENT")
    [void]$s.Append("END:VCALENDAR")

    # Output to disk, if you wish. Outputs to required UTF-8 w/o BOM.
    #[IO.File]::WriteAllLines('calendarentry.ics', $s.ToString())

    # Create the invitation and add it to the message body.
    # Doing it this way gives a richer interface in Gmail
    $ContentType = New-Object System.Net.Mime.ContentType("text/calendar")
    $ContentType.Parameters.Add("method","REQUEST")
    $ContentType.Parameters.Add("name", "invite.ics")
    $Attachment = [Net.Mail.AlternateView]::CreateAlternateViewFromString($s.ToString(), $ContentType)
    $Mail.AlternateViews.Add($Attachment)

    # Send the email.
    $Smtp.Send($Mail)

    # Dispose of the resources.
    $Mail.Dispose()
    $Smtp.Dispose()
}

Export-ModuleMember -Function Send-WasteCollectionInvitations