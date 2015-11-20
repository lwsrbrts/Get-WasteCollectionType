<#
.Synopsis
   Sends invitations by email
.DESCRIPTION
   Sends invitations by email to a provided hashtable of recipients.
   You must define the recipients, server, port, credentials, from.
   The idea is not to run this over and over again for every recipient
   but to send a single email and include all attendees on the invite
   (usually the people responsible for putting the bin out!).
   If "automatically add events to calendar" is enabled on the users'
   Gmail calendar, they will get a reminder at 7pm in their time zone.

   The invitation is set to occur at 7pm the evening before the calculated
   date of collection. ie. If collection is on Monday 16th, the invitation
   will be set for Sunday 15th at 7pm.
.EXAMPLE

   $Recipients = 
    @{
        "Lewis Roberts" = "lewis@lewisroberts.com";
        Joe Bloggs" = "joe.bloggs@gmail.com";
    }
    Send-WasteCollectionInvitations -To $Recipients `
                                    -Server 'smtp.gmail.com' `
                                    -Port 587 `
                                    -From 'John Doe <john.doe@gmail.com>' `
                                    -Credentials $SenderCredentials `
                                    -Subject 'Waste Collection' `
                                    -EventDate $EventDate
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
        [datetime]$EventDate,

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

    $EventUTC = $EventDate.ToUniversalTime()
    
    # To have the description of the invite make grammatical sense,
    # I remove the default subject prefix from it.
    $Description = ($Subject -replace "Waste Collection - ").ToLower()

    # Add the HTML message body to the email from the injected $Body.
    $Mail.AlternateViews.Add([Net.Mail.AlternateView]::CreateAlternateViewFromString($Body, 'text/html'))
    
    # Build the invitation with StringBuilder
    # I've left some of the iCal bits in as a reminder of available options.
    $s = New-Object System.Text.StringBuilder
    [void]$s.AppendLine('BEGIN:VCALENDAR')
    [void]$s.AppendLine("PRODID:-//lewisroberts//EN")
    #[void]$s.AppendLine("CALSCALE:GREGORIAN")
    [void]$s.AppendLine("VERSION:2.0")
    [void]$s.AppendLine("METHOD:REQUEST")
    [void]$s.AppendLine("BEGIN:VEVENT")
    #[void]$s.AppendLine("TRANSP:TRANSPARENT")
    [void]$s.AppendLine([String]::Format("DTSTART:{0:yyyyMMddT190000Z}", $EventUTC.AddDays(-1)))
    [void]$s.AppendLine([String]::Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", (Get-Date).ToUniversalTime().AddMinutes(-1)))
    [void]$s.AppendLine([String]::Format("DTEND:{0:yyyyMMddT200000Z}", $EventUTC.AddDays(-1)))
    [void]$s.AppendLine("LOCATION:Holmes Chapel, Cheshire")
    [void]$s.AppendLine([String]::Format("UID:{0}", [Guid]::NewGuid()))
    [void]$s.AppendLine("ORGANIZER;CN=`""+$Mail.From.DisplayName+"`":MAILTO:"+$Mail.From.Address)
    Foreach ($Recipient in $To.GetEnumerator()) {
        [void]$s.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=ACCEPTED;CN="+$($Recipient.Name)+":mailto:"+$($Recipient.Value)+"")
    }
    #[void]$s.AppendLine("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=Waste Collections;X-NUM-GUESTS=0:mailto:waste_collections@lewisroberts.com")
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

    # Output the .ics file to disk, if you wish. Outputs to required UTF-8 w/o BOM.
    #[IO.File]::WriteAllLines('calendarentry.ics', $s.ToString())

    # Create the invitation and add it to the message body.
    # Doing it this way gives a richer interface in Gmail and probably Outlook.com
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