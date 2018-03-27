# supportops-bdma
Barcode DataManager Automation
#
This repository is to hold the MOVEit Central custom VBScript code used for processing the received barcode and DataManager
import files uploaded to the sFTP server by districts and the Contracts team.
#
The script collects all files in a given 'barcode' or 'datamanager' sub-folder in a district hierarchy, compresses them
into a single .zip file, and places that file in a 'ticketed' folder that is a sibling of the folder in which the files
are found. It also looks up the account for each folder, using the parent folder name as the "BAS SFTP Username" from
Salesforce account records. A lookup file is generated in Salesforce and placed in a location that the script can access.
From the lookup file, the account name, PID, and ISA# from the account record (plus the contact name and email address from
the DataManager Account Holder contact record) are placed in an email message, which is then sent to a Salesforce
email-to-case address to automatically create a Salesforce case that is assigned to the Itasca SupportOps team's
Data Integration Queue to be worked by the team.
#
The script relies on task parameters to customize its behavior so that it can be tested on the development environment
and deployed to the production environment without code changes. The script also embeds a UNC link to the archived
.zip file into the email body, so the team member working the case can find the file quickly.
