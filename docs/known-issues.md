# Known issues

This section lists limitations and other issues known to exist in Boffo at this time.

## Execution time limits

Depending on your account type, Boffo will run into Google quotas limiting maximum execution time. For G Suite users such as the Caltech Library, this limit is 30 minutes. For non-G Suite users, the limit is 6 minutes. At the time of this writing, Boffo can retrieve item records at a rate of about 70–100 records/second, so the 6 minute time limit will limit non-G suite users to looking up somewhere between 25,000–35,000 item records at a time.

There seems to be a limit of 10,000,000 cells in a spreadsheet, and this number applies to the total _across all sheets in a given spreadsheet_. If you keep all your results in a single Google spreadsheet and just keep adding new sheets/tabs to the same single spreadsheet, you can actually hit this limit if you have Boff return a lot of columns and a lot of results all the time. So just be mindful of this possibility. (The easiest solution is not to create new spreadsheets every once in a while.)
