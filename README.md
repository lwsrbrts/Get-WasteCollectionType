# Get-WasteCollectionType
PowerShell scripts and a module to scrape waste (bin!) collection schedule data from Cheshire East council, send it as an email along with a "night before" reminder calendar invitation.

## Why?
Well, each week on the same day, I have a bin (waste/garbage) collection, like most households do. I found myself checking the online portal for Cheshire East Council for the type of bin that needed to go out on the kerb.
The first version of the script just scraped the site, cleaned up the formatting a little and sent an email with the data I'd scraped. Since I had a lot of time on my hands on an evening while working away from home, I decided to see if I could improve the process by sending a calendar event/invite as well. This gives the added benefit that there's a reminder that the bin (waste/garbage) needs to go out on the kerb for the following day and it tells me what bin by colour (we have three bins for different purposes in Cheshire East: Recyclables (Silver/Grey), Garden (Brown) & Other (Black)) that needs to go out.

## Really?
Yes, really. It's an over-engineered solution to a first world problem but as a I mentioned, I had a lot of time on my hands.

## Can I use this?
In its current guise, probably not unless you live in Holmes Chapel, Cheshire East, UK or have a decent grasp of how to change the scripts to suit your purposes. If you do have a decent grasp of PowerShell, I'm sure you could adapt the code to suit your purposes. The iCal event creation code is, I imagine, the most interesting part of this whole solution for most people looking at it. I don't mind admitting it took me nearly two weeks of mid-week evenings to work out how to properly format and send an iCal invite like this.

## So why did you put it on GitHub?
To learn about GitHub more than anything else. I'm not a developer in the truest definition, I've come from a Windows infrastructure/architecture background but I do have a fairly advanced grasp of PowerShell (at least, I think I do) and a desire to stop having to check a website every week to find out what bin I need to stick on the kerb the night before.