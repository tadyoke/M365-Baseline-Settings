# M365-Baseline-Settings

Here are **a few things to note** about this script:

If your tenant has more than about 5000 users _Do Not Use This script_ - it will cause you timeout headaches and is very slow. I recommend using the export with one workload at a time. 

This requires that you have already installed locally the Azure AD, Exchange Online, PNP, SharePoint, etc. PowerShell modules - I might add this check later, but I didn't need it for now.

This is simplest using a cloud-only global admin account without MFA but using Conditional Access to restrict the account to connect from a single IP Address.

In my testing it worked fine with my global admin account with MFA but required more interaction than I wanted.

This script uses some of the Microsoft 365 DSC (microsoft/Microsoft365DSC on github) script components. The DSC is very helpful especially if you need to document tenant settings. And it doesn’t cost anything which is a bonus.

This script is really very basic and not the most elegant as far as code goes but it will do the following:
 1. Set some variables and ask for your tenant domain in the format domain.onmicrosoft.com
 2. Gets data for the following workloads: O365 tenant, Security and Compliance, Azure AD, Microsoft Teams, Exchange Online, Intune, SharePoint and OneDrive. Note that PowerPlatform and Planner have been removed for two separate reasons. I'm still working out how best to manage those.
 3. Checks the M365DSC PowerShell module is installed and that it is the current version
 4. Creates an M365DSC folder in the logged-in user's documents folder - change it if you need
 5. Creates a Monthly folder
 6. Creates a workload folder for the original export files the DSC creates
 7. Writes the workload xlsx files all to the monthly folder, converts them to csv, aggregates them into one xlsx file and removes the interim csv files

The idea was to get all this info into one file for historical reasons.
The Microsoft 365 DSC Compare utility can still be used against the ps1 files in each workload folder. I don't know yet if I will automate that.
