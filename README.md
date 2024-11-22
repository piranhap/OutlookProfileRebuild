# Re-build an Outlook Profile
At some point for company xyz, we migrated all of our email from one provider to another. After pointing all the MX records and doing all the tasks that come with a mail migration. We, the IT Staff of XYZ, needed to re-build the profile for our new provider.

## Our use case
We needed a software/script/anything that would remove the profile and re-create it, then let autodiscover do its thing. So this is what we came up with.

## How it works

1. Have the user run the script, however you distribute it is up to you, this can be turned into an exe or msi with other tools. (out of scope for this post).
2. The script will (admin privileges highly recommended) close outlook, close all other applications that take ownership of office processes (do your due diligence here, since all environments are different, you can use [Process Explorer from Sysinternals](https://learn.microsoft.com/en-us/sysinternals/downloads/process-explorer)), remove the profile from the registry, remove local OST/NST files, re-create the profile 'outlook', and open outlook again. 
3. Outlook will now auto-discover and re-configure itself to use the new provider.

## Limitations

* Only tested on Windows 11/10
* Only tested on Outlook 2016
* If you think of any other limitations, please add an issue. 

## Nice to haves (to dos)

* Test with other versions of Outlook.
* Configure Outlook to sync all email from all time. 

## I hope you find this useful, if you have any questions please reach out through email.