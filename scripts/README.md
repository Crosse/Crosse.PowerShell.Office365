# Check-Forwarding.ps1

The `Check-Forwarding.ps1` script is what JMU uses in a scheduled task to detect new and duplicate SMTP forwards in our Office 365 tenant domain.
It uses the functions provided by the [Crosse.PowerShell.Office365][] and [Crosse.PowerShell.Exchange][] PowerShell modules. so be sure that you have both of those modules somewhere in `$PSModulePath`.

In order to release the script, I had to make a few changes to the script params and some other minor tweaks.
The version of the script in this repository is **not** the same as the one we use in production, but it's pretty close.
If you encounter any issues, please [create an issue][new-issue] or open a pull request.

Once I have integrated the changes made here to the script we have in production, I will edit this file to remove the big scary warnings about it not being tested.

[Crosse.PowerShell.Exchange]: https://github.com/Crosse/Crosse.PowerShell.Exchange
[Crosse.PowerShell.Office365]: https://github.com/Crosse/Crosse.PowerShell.Office365
[new-issue]: https://github.com/Crosse/Crosse.PowerShell.Office365/issues/new
[new-pr]: https://github.com/Crosse/Crosse.PowerShell.Office365/issues/new