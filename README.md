# Exim Email Assistant

You _can_ make Exim handle email redirection, and out of Office (Vacation messages).  But the mechanism for doing so is to create .forward file in the user home space.

I couldn't find any plugins for Thunderbird which would create / manage this file, so I hacked this little .hta file to do the job.  It's pretty basic but hopefully will allow our users to manage their redirection and Out of Office needs.


The user can:

* Enter an email address, then copy all inbound mail to that address
  * optionally not receive a copy themselves (though this code is commented out at present)
* Specify or edit an Out of Office (Vacation) message
* Turn off both of the above

Assumptions:

1. The linux $HOME space (where exim will look for the file) and the windows %HOMESHARE% are both mapped to the same place, if not you'll need to adjust the code.
1. The users have write access to the above space
1. Running local .hta files is allowed