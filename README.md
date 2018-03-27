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
The script relies on task parameters to customize its behavior so that it can be tested on the development environment and deployed
to the production environment without code changes. The script checks for missing parameters and whether the lookup file and
compression program can be found, and whether the email server can be reached, and ends gracefully if a problem occurs. It ignores
files at the top level folder and reports in the log file the number of files found and the number of email sents. The scipt 
embeds a UNC link to the archived .zip file in the email body, so the team member working the case can find the file quickly.
#
Current task parameters used by the script:<p>
DI_Application: The type of application files the script is looking for. Current possible values: Barcode, DataManager<br>
DI_Base_Directory: The path to the parent directory in which all the files found by the Source element are found. Typically on the local S: drive.<br>
DI_Email_From: The default email address to which the From field of the emails will be set. SFDC will use it as the default To address when sending emails from the case. When a DM account holder is found, their email address will be used.  Currently using DoNotReply@hmhco.com.<br>
DI_Email_To: The email address to which the emails will be sent. Used for testing, and set to the SFDC email-to-case address on production.<br>
DI_Server_Folder: The topmost path of the file hierarchy under which all of the files in the sFTP server are found. Used to generate the UNC path to the archive file. Set to 'e7sftp' for regular customers, and to 'e7sftp\contracts' for contracts customers.<br>
DI_Server_Name: The network name of the sFTP server. Used to generate the UNC path to the archive file.<br>
DI_SMTP: The name of the SMTP server that will be used to create and sent the emails.<br>
DI_Lookup File: The path name of the lookup file created by exporting a Salesforce report containing the fields used by the script.<br>
DI_Zip File: The path to the 7Zip compression program on the local server that will be used to compress the import files.
