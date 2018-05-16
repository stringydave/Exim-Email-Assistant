# Exim Email Assistant

In our environment we use a Debian server, running Samba to make it emulate a Windows server, and Exim to handle mail.  Clients are Windows running Thunderbird as the mail client.  Users regularly ask us if they can enable out of office or redirect mail to someone else while they are away.  Exim _can_ do this, but the mechanism for doing so is to create .forward file in the user (Linux) home space.  This is fine if you're a Linux command line nerd, for normal windows users - not so much.

I couldn't find any plugins for Thunderbird which would create / manage this file, so I hacked this little .hta file to do the job.  It's pretty basic but hopefully will allow our users to manage their redirection and Out of Office needs.  The main reason was to figure out the steps required to make it work.

The user can do one (and only one) of:

* copy all inbound mail to another user
* specify or edit an Out of Office (Vacation) message
* turn off both of the above

Assumptions:

1. the Linux $HOME space (where exim will look for the file) and the windows %HOMESHARE% are both mapped to the same place, if not you'll need to adjust the code, or provide some hacks on the Linux side
1. the users have write access to the above space
1. the user username is different to the user email address (the part before the @), e.g. the email address for user jdoe is john.doe@company.com
1. running local .hta files is allowed

Notes:

Because the user username is different to the user email address, we need to use the alias syntax.

ToDo:

* cope with multiple aliases
* re-write it in a modern language

Maybe:

* show the user a summary of what we did when they turn off Out of Office
* make it look prettier