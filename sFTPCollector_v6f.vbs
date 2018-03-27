Option Explicit

' Declare constants

' Declare MOVEit log levels - used for the debugging function that writes status messages to the log
Const DL_Error_App  =  0 
Const DL_Error_Task = 10
Const DL_Warning    = 20
Const DL_OK_Task    = 30
Const DL_OK_File    = 40
Const DL_Debug_Some = 50
Const DL_Debug_More = 60
Const DL_Debug_All  = 70
' Maximum size (bytes) of a file attachment
' Const FileSizeMax   = 15360000
' File System Object I/O mode for Reading
Const ForReading = 1
' File System Object I/O mode for Writing
Const ForWriting = 2
' Const vbTextCompare = 1

' Declare variables
' intIndex - Counter variable for outer For loop. Equal to the number of files processed. Also used to give the compressed files a unique filename.
' intLoop - Counter variable for For loop that identifies the parent folder of the current file. Used after loop finished to get previous level directory name.
' intDepth - The deption of the directory in which the current file is found. If less than 2, then we are at the same level as the hostname - not good!
' intEmailCount - Couner variable for the number of email messages sent - informational only
' strError - The type of error. The Case statment handles each error type differently
' strErrMsg - The text of the error message to print to the log
' strCurrentHost - The current hostname. A change in this signals a new district.
' strLastHost - The prior hostname. CurentHost is compared to this to detect a new district.
' strTempDir - The MOVEit temporary directory. All found files are copied here and deleted when the task exists. The compressed file is also built here, then copied to the final directory.
' strFileList - The list of files found by the source of the task
' strLookupFilepath - The task parameter with the full pathnameof  the lookup file
' strZipExe - Tje task parameter with the full path of the compression utility
' strApplication - The task parameter that tells us which HMH application we're working on (Barcodes, DataManager, Edusoft, DataDirector)
' strCurrentPath - The path (starting from the host) to the parent directory of the current file. A change in this signals a new set of files.
' strLastPath - The path to the parent folder of the previous file. When this changes, zip up what we've found and send an email.
' strFilePaths - A pipe-delimited list of the files found so far. The compression tool will zip them up to be attached to the email.
' strFolderName - The parent directory immediately above the current file. We use this to determine the type of file, which is used to decide which 2-char product code to embed in the archive file name.
' arrFileList - An array that holds each filepath found by the task source in a separate bucket. The outer For loop uses it to process each file.
' arrPath - An array that holds each directory name in a separate buket. We use it to build the CurrentPath to the file, so we know when the parent folder changes.
' arrAccountList - An array that holds every line from the lookup file. The Lookup function searches it for the hostname, so that it can set values for the iMacro script.
' objFSO - The File System Object used to read the lookup file and copy the zip archive to the 'ticketed' folder of the parent directory.
' objCDO - The Collaboration Data Message Object used to create and send the email message
' objCDG - The Collaboration Data Configuration Object used to set the parameters of the CDO Message object
' objTS - The TextStream object used to read the lookup file and write the FileList.txt file used by the compression utility to build the zip archive

' Variables for main program

Dim intIndex, intLoop, intDepth, intEmailCount, strError, strErrMsg, strCurrentHost, strLastHost, strTempDir, strFileList, strLookupFilepath, strZipExe,  _
 strApplication, strCurrentPath, strLastPath, strFilePaths, strFolderName, arrFileList, arrPath, arrAccountList, objFSO, objCDO, objCDG, objTS

' Functions and Subroutines

Sub DLogMsg(DLevel, Message)

' Logging routine - prints log message only when MOVEit debug level matches value passed in
  If DLevel <= MIGetDebugLevel() then MILogMsg Message
  
End Sub


Sub PrintArray(arrPArray)

' Only used for debugging - list all values in an array to the log

Dim intElement

	DLogMsg DL_Debug_More,  "PrintArray:Starting"
	For intElement = 0 to UBound(arrPArray)
		MILogMsg "PrintArray:Element " & CStr(intElement) & " is: " & arrPArray(intElement)
	Next

End Sub


Function CaseCategory(strFolder)

' Associates sFTP folder name with Riverside Application using case insenstive match
	Select Case LCase(strFolder)
		Case "datamanager"
			CaseCategory = "Rostering"
		Case "datadirector"
			CaseCategory = "Data Integration"
		Case "edusoft"
			CaseCategory = "Data Integration"
		Case "barcode"
			CaseCategory = "Barcode"
		Case Else
			CaseCategory = "Data Integration"
	End Select

End Function


Function RiversideApplication(strFolder)

' Associates sFTP folder name with Riverside Application using case insenstive match
	Select Case LCase(strFolder)
		Case "datamanager"
			RiversideApplication = "E7_DataManager"
		Case "datadirector"
			RiversideApplication = "DataDirector"
		Case "edusoft"
			RiversideApplication = "Edusoft"
		Case "barcode"
			RiversideApplication = ""
		Case Else
			RiversideApplication = ""
	End Select

End Function


Function ProductCode (strFolder)

' Associates sFTP folder name with product code prefix used in filename using case insenstive match
	Select Case LCase(strFolder)
		Case "datamanager"
			ProductCode = "dm"
		Case "datadirector"
			ProductCode = "dd"
		Case "edusoft"
			ProductCode = "ed"
		Case "barcode"
			ProductCode = "bc"
		Case Else
			ProductCode = "ot"
	End Select

End Function


Sub SendEmail(strSEFrom, strSETo, strSESubject, strSEBody)

' Sends an email using the paramters passed in and the global CDO object
	DLogMsg DL_Debug_More, "SendEmail: Send email routine called."
	DLogMsg DL_Debug_Some, "SendEmail: Message From is: " & strSEFrom
	DLogMsg DL_Debug_Some, "SendEmail: Message To is: " & strSETo
	DLogMsg DL_Debug_Some, "SendEmail: Subject is: " & strSESubject
	DLogMsg DL_Debug_Some, "SendEmail: Body is: " & strSEBody
	With objCDO
		.From = strSEFrom
		.To = strSETo
		.Subject = strSESubject
		.TextBody = strSEBody
		.Send
	End With
	DLogMsg DL_Debug_More, "SendEmail: Message sent."
'	Clear CDO Message fields for the next iteration of the outer For loop
	With objCDO
		.From = ""
		.To = ""
		.Subject = ""
		.TextBody = ""
	End With
	intEmailCount = intEmailCount + 1

End Sub


Function LookupAccount(strPath, strHostname)

' Reads through the global array for the row that matches the strHostname argument and sets the other arguments to the values in the rest of the row.
' Sets the found value "#N/A" to "" so that the returned arguments will be empty if the record has that value in the field.

' This version was updated with a nested loop to search through the lookup file for the hostname. It begins with the parent folder in which the files
' were found and continues up the hierarchy until either the folder name (hostname) is found or the top-level directory is reached. If no hostname
' is found, the array is set to empty values except that the first bucket is set to the top-level directory name (the district 'hostname') and the
' second entry is set to "NOT FOUND" so the calling routine will know that the lookup failed and place a warning message in the email body.

' Declare variables
' i - Counter variable used for the outer For loop
' j - Counter variable for Inner For loop 
' intCount - Counter variable used to identify in which row of the lookup array the hostname is found
' arrAccount -  Array used to separate the fields of the lookup row into separate buckets
' arrFolders - Array used to hold the names of the folders the directory path
' strLine - The loop counter variable for the main For..Each loop
' strName - The hostname currently being searched for
' boolFound - Boolean variable to indicate whether the hostname was found
' intFieldMax - Number of fields in the lookup row. Used to redimension the arrAccount array and elminate extra buckets

Dim i, j, intCount, arrAccount, arrFolders, strLine, strName, boolFound, intFieldMax

'	DLogMsg DL_Debug_More, "LookupAccount: Function called"
'	DLogMsg DL_Debug_Some, "LookupAccount: Second line of file is: " & arrAccountList(1)
'	DLogMsg DL_Debug_Some, "LookupAccount: Number of lookup file rows is: " & CStr(Ubound(arrAccountList))
	DLogMsg DL_Debug_Some, "LookupAccount: Host to lookup is: " & strHostname
	DLogMsg DL_Debug_Some, "LookupAccount: Path to lookup is: " & strPath

'	How many fields are in the lookup file?
	intFieldMax = UBound(Split(arrAccountList(0), ","))
	DLogMsg DL_Debug_More, "LookupAccount: Number of lookup file fields is: " & CStr(intFieldMax)

'	Repeat loop until either hostname found or all the directory names have been searched for, starting with the folder
'	the file is in, and moving up through the folder path until either the hostname is found or the top-level folder is reached
	arrFolders = Split(strPath, "\")
'	Call PrintArray (arrFolders)
	DLogMsg DL_Debug_Some, "LookupAccount: Path depth is: " & CStr(UBound(arrFolders))
	For i = UBound(arrFolders) To 0 Step -1
		DLogMsg DL_Debug_Some, "LookupAccount: For: Directory to search is: " & arrFolders(i)
		intCount = 1
		For Each strLine in arrAccountList
			If strLine = "" Then
				DLogMsg DL_Debug_Some, "LookupAccount: For-For_Each-If1: Line " & CStr(intCount) & " is empty. "
				Exit For
			End If
' 			Row counter
			intCount = intCount + 1
'			Load each field into elements of an array
			arrAccount = Split(strLine, ",")
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: Loop Count: " & intCount
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: Current line of file is: " & strLine
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: First field of current line is: " & arrAccount(0)
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: Length of Hostname is: " & CStr(Len(arrFolders(i)))
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: Length of first field of current line is: " & CStr(Len(arrAccount(0)))
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: First character of first field is: " & CStr(Asc(Left(arrAccount(0),1)))
			DLogMsg DL_Debug_More, "LookupAccount: For-For_Each: Last character of first field is: " & CStr(Asc(Right(arrAccount(0),1)))
			If arrAccount(0) = arrFolders(i) Then
'				The hostname was found
				DLogMsg DL_Debug_Some, "LookupAccount: For-For_Each-If2: Account Found on line " & CStr(intCount - 1)
				DLogMsg DL_Debug_Some, "LookupAccount: For-For_Each-If2: Hostname is: " & arrAccount(0)
				DLogMsg DL_Debug_Some, "LookupAccount: For-For_Each-If2: File field2 is: " & arrAccount(2)
				For j = 0 to UBound(arrAccount)
					If arrAccount(j) = "#N/A" Then
						arrAccount(j) = ""
					End If
					boolFound = True
				Next
'				PrintArray arrAccount
'				Hostname was found, so exit internal For loop
				Exit For
			Else
				boolFound = False
				DLogMsg DL_Debug_More, "LookupAccount: For-For_Each-If2_Else: Account not found"
			End If
		Next
		If NOT boolFound Then
			DLogMsg DL_Debug_Some, "LookupAccount: For-For_Each-If3: Account not found"
		Else
'			Hostname was found, so exit outer For loop
			Exit For
		End If
	Next

'	If hostname was not found, set all fields to empty string,
'	 then set first field of the array to the top-level folder name so the Zip function can name the archive file,
'	 and the second field of the array to "NOT FOUND" to indicate that the hostname wasn't found in the lookup file 
	If NOT boolFound Then
		DLogMsg DL_Debug_Some, "LookupAccount: For-If1: ACCOUNT NOT FOUND"
		ReDim arrAccount(intFieldMax)
		For j = 0 to UBound(arrAccount)
			arrAccount(j) = ""
		Next
		arrAccount(0) = strHostname
		arrAccount(1) = "NOT FOUND"
	End If

'	Call PrintArray (arrAccount)
 	LookupAccount = arrAccount
	DLogMsg DL_Debug_Some, "LookupAccount: Loop count is: " & CStr(intCount)

End Function


Function ZipFiles(strZFilePaths, strZParentFolder, strZHost, intZPass)

' This code tries to create a Zip archive using the list of file paths that are passed in via the first argument. It uses the ParentFolder to determine
' where to put the compressed file (inside a 'ticketed' folder), and the Host argument to include a text string in the filename that identifies
' the application type. The Pass argument ensures that a unique filename is generated for multiple folders under a single hostname.

' Declare variables
' intIndex - Loop counter for the For loop that builds the filepath to the same level as the parent direcr=tory. The zip file will be copied to a 'ticketed' folder at this level.
' strZipFileName - The name of the zip archive that will be created
' strTempPath - The path from the hostnmae direcotry to the level just above the parent directory of the files. Built by the For loop.
' strCmd - The Windows command line that will invoke the compression utility and create the zip file.
' strResult - The value returned by the MIRunCommand function, Zero indicates success, anything else is failure.
' strDestinationFolder - the final location of the zip file after it is created
' arrParentPath - array used to hold each directory of the parent file path. Used by the For loop to build the TempPath.

Dim intIndex, strZipFileName, strTempPath, strCmd, strResult, strDestinationFolder, arrParentPath

	DLogMsg DL_Debug_Some, "ZipFiles: Pass is: " & CStr(intZPass)
	DLogMsg DL_Debug_Some, "ZipFiles: Parent Folder is: " & strZParentFolder
'	Call PrintArray(Split(strZFilePaths, "|"))

'	The destination folder will be at the same level as the Parent Folder, so we need to recreate the path up to the level above it
	arrParentPath = Split(strZParentFolder, "\")
	strTempPath = arrParentPath(0)
	For intIndex = 1 to Ubound(arrParentPath) - 1
		strTempPath = strTempPath & "\" & arrParentPath(intIndex)
	Next
	strZipFileName = MIMacro("[YYYY]-[MM]-[DD]-[HH]-[TT]-[SS]") & "-" & intZPass & "-" & ProductCode(strFoldername) & "-" & strZHost & ".zip"
	strDestinationFolder = MIGetTaskParam("DI_Base_Directory") & "\" & strTempPath  & "\ticketed"
'	Create full path for zip file
	ZipFiles = strDestinationFolder & "\" & strZipFileName
	DLogMsg DL_Debug_Some, "ZipFiles: Filelist is: " & strZFilePaths
	DLogMsg DL_Debug_Some, "ZipFiles: Zipped temp filepath is: " & strTempPath
	DLogMsg DL_Debug_Some, "ZipFiles: ZipFiles is: " & strZipFileName
	DLogMsg DL_Debug_Some, "ZipFiles: Zipped final filepath is: " & ZipFiles

'	Create a temp file to hold the list of files to compress. This will be passed as an argument to the compression utility.
	objFSO.CreateTextFile strTempDir & "\FileList.txt"
	Set objTS = objFSO.OpenTextFile(strTempDir & "\FileList.txt", ForWriting)
	objTS.Write Replace (strZFilePaths, "|", Chr(13))
	objTS.Close
	Set objTS = Nothing

'	Create the command line arugment to invoke the compression utility
	strCmd = strZipExe & " a " & strTempDir & "\" & strZipFileName & " -i@" & strTempDir & "\FileList.txt"
	DLogMsg DL_Debug_Some, "ZipFiles: If1: Command is: " & strCmd
	strResult = MIRunCommand(strCmd)
	If strResult <> 0 Then
		DLogMsg DL_Debug_Some, "ZipFiles: ERROR " & CStr(strResult) & " reported after compression utility called."
	End If
	DLogMsg DL_Debug_Some, "ZipFiles: Destination folder is: " & strDestinationFolder

'	If the 'ticketed' folder doesn't exist, create it
	If Not(objFSO.FolderExists(strDestinationFolder)) Then
		DLogMsg DL_Debug_Some, "ZipFiles: If2: Trying to create destination folder: " & strDestinationFolder
		objFSO.CreateFolder(strDestinationFolder)
	End If
	objFSO.CopyFile strTempDir & "\" & strZipFileName, ZipFiles
	ZipFiles = strDestinationFolder & "\" & strZipFileName

End Function


Sub SendFiles(strSFFileList, strSFPath, strSFHost, intSFPass)

' This is the code that prepares the email for sending. It uses the LookupAccount function to retrieve information about the SFDC account
' using the BAS SFTP account hostname as the lookup value.
'
' For DataManager files, the email's From address is set to the Contact email address to make it appear as if the email came from them.
' The SFDC report used to create this file selects the contact record in which the 'Role w/ RPC Products' field contains the text 'DataManager Account Holder'.
'
' Barcode customers do not have a unique contact record in SFDC, so we cannot mimic sending the email from a person in the district. The
' email is therefore sent using the default From address set into the task parameters.
'
' This version was updated to insert the account's MDR PID at the bottom so that SFDC will automatically make the case to belong
' to the parent account in which the BAS SFTP hostname is set, and to include the ISA number for barcode orders.
' 
' The expected structure of the lookup (currently created using an SFDC report) is as follows (the lookup array is zero-indexed):
' Column    Description
' 1         Account BAS SFTP Username
' 2         Account Name
' 3         Contact Full name (First Last)
' 4         Contact Email address
' 5         Account PID
' 6         Account ISA#


' Declare variables
' intAt - Position of the '@' symbol in the email address (from the start)
' intDot - Position of the '.' symbol in the email address (from the end) 
' strSFCompressedFile - The path to the zip file that is returned from the ZipFiles function
' strFileName - The loop variable for the For..Each loop that proceses each file in the strSFFileList argument. Adds each filename as separate row to the email body.
' strSFFrom - The email From address sent to the SendEmail function
' strSFTo - The email To address sent to the SendEmail function
' strSFSubject - The email Subject sent to the SendEmail function
' strSFBody - The email Body sent to the SendEmail function
' arrFieldList - Array used to hold each of the fields returned by the LookupAccount function
' arrFileList - Array used to hold each of the filenames passed in via the strSFFileList argument. The For..Each loop uses it to put them in the email body.
' arrFilePath - Array used to hold each directory in the current file path being processed in the For..Each loop. The last bucket is always the filename.

Dim intAt, intDot, strSFCompressedFile, strFileName, strSFFrom, strSFTo, strSFSubject, strSFBody, arrFieldList, arrFileList, arrFilePath

	DLogMsg DL_Debug_More, "SendFiles: Sub called."
	DLogMsg DL_Debug_Some, "SendFiles: Path is: " & strSFPath
	DLogMsg DL_Debug_Some, "SendFiles: Hostname is: " & strSFHost
	DLogMsg DL_Debug_Some, "SendFiles: File List is: " & strSFFileList

'	Split the pipe-delimited list of file paths into an array
	arrFileList = Split(strSFFileList, "|")
'	Look up the host folder name to get the account and contact information
	arrFieldList = LookupAccount(strSFPath, strSFHost)

'	Create the zip file, handling the error condition of the hostname not in the lookup file
	strSFCompressedFile = ZipFiles(strSFFileList, strSFPath, arrFieldList(0), intSFPass)

'	The following code block will try to determine the application type from the file's parent folder (instead of using the DI_Application parameter)
'	Extract the names of the folders in which the file was found into an array
'	arrFilePath = Split(strSFPath, "\")
'	Now we can handle the  email differently depending on the parent folder name, which is the last element in the array
'	Select Case arrFilePath(UBound(arrFilePath))

'	Determine the application type from the DI_Application parameter
	DLogMsg DL_Debug_Some, "SendFiles: Application is: " & strApplication 

' 	Prepare the email
	strSFTo = MIGetTaskParam("DI_Email_To")
	strSFBody = "The following file(s) were found in a '" & strFolderName & "' directory under hostname: " & arrFieldList(0) & Chr(13) & Chr(10) & Chr(13) & Chr(10)

'	Handle the case when the hostname is not found by giving the email a different subject and inserting a line into the body
	Select Case arrFieldList(1)
		Case ""
			strSFSubject = arrFieldList(0) & " has received new file(s) in a '" & strFolderName & "' directory"
		Case "NOT FOUND"
			strSFSubject = arrFieldList(0) & " has received new file(s) in a '" & strFolderName & "' directory"
			strSFBody = strSFBody & "-------------------------------------------------------------------------------------------------------------------------------------------------" & Chr(13) & Chr(10)
			strSFBody = strSFBody & " WARNING: The username was not found in lookup file - review Salesforce Contact and/or Account records." & Chr(13) & Chr(10)
			strSFBody = strSFBody & "-------------------------------------------------------------------------------------------------------------------------------------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)
		Case Else
			strSFSubject = arrFieldList(1) & " has received new file(s) in a '" & strFolderName & "' directory"
	End Select
	strSFBody = strSFBody & "Location of files: " & MIGetTaskParam("DI_Server_Name") & "\" & MIGetTaskParam("DI_Server_Folder")&  "\" & strSFPath & Chr(13) & Chr(10) & Chr(13) & Chr(10)
	For Each strFileName in arrFileList
'		Add only the filename, which will be the last element of the array we split each file path into
		arrFilePath = Split(strFileName,"\")
		strSFBody = strSFBody & arrFilePath(UBound(arrFilePath)) & Chr(13) & Chr(10)
	Next
	strSFBody = strSFBody & Chr(13) & Chr(10)

'	Add the Case Category, Riverside Application, Account Name, Contact Name, contact Email Address, and User Name (hostname) information from the lookup file that the iMacro uses to populate the SFDC fields
	strSFBody = strSFBody & "CC:" & CaseCategory(strFolderName) & "- RA:" & RiversideApplication(strApplication) & "- AN:" & arrFieldList(1) & "- CN:" & arrFieldList(2) & "- EA:" & arrFieldList(3) & "- UN:" & arrFieldList(0) & "- END" & Chr(13) & Chr(10)
'	We turn file path into a UNC, using the DI_Server_Name task parameter which is set to the network name of the sFTP server.
'	Then we find the location of the  directory set into the DI_Server_Folder task parameter in the path and grab everything after that in the path.
'	DLogMsg DL_Debug_Some, "SendFiles: Search result: " & CStr(InStr(strSFCompressedFile, "\e7sftp"))
	DLogMsg DL_Debug_Some, "SendFiles: File path: \\" & MIGetTaskParam("DI_Server_Name") & "\" & Right(strSFCompressedFile, Len(strSFCompressedFile) - InStr(strSFCompressedFile, "\" & MIGetTaskParam("DI_Server_Folder")))
	strSFBody = strSFBody & Chr(13) & Chr(10) & "File path is: \\" & MIGetTaskParam("DI_Server_Name") & "\" & Right(strSFCompressedFile, Len(strSFCompressedFile) - InStr(strSFCompressedFile, "\" & MIGetTaskParam("DI_Server_Folder")))
	strSFBody = strSFBody & Chr(13) & Chr(10) & Chr(13) & Chr(10)& "ISA#: " & arrFieldList(5)
	Select Case strApplication
'		We don't have a single point of contact for barcode orders, so use the default email address
		Case "Barcode"
			DLogMsg DL_Debug_Some, "SendFiles: Case: Barcode file(s)."
			strSFFrom = MIGetTaskParam("DI_Email_From")
		Case "DataManager"
			DLogMsg DL_Debug_Some, "SendFiles: Case: DataManager file(s)."
'			Perform simple format validation of retrieved email address to allow CDO 'from' field to be set
			intAt = InStr(arrFieldList(3),"@")
			intDot = InStrRev(arrFieldList(3),".")
			DLogMsg DL_Debug_Some, "SendFiles: Case: Lookup column 4 value: " & arrFieldList(3)
			DLogMsg DL_Debug_Some, "SendFiles: Case: @ found at: " & CStr(intAt)
			DLogMsg DL_Debug_Some, "SendFiles: Case: . found at: " & CStr(intDot)
			If intAt > 0 Then
				If intDot > intAt Then
'					Use the address from the lookup file
					DLogMsg DL_Debug_Some, "SendFiles: Case-If2: Using Lookup file From address"
					strSFFrom = arrFieldList(3)
				Else
'					Invalid entry in the lookup file, so use the default From address
					DLogMsg DL_Debug_Some, "SendFiles: Case-If2-Else: Using default From address"
					strSFFrom = MIGetTaskParam("DI_Email_From")			
				End If
			Else
'				Invalid entry in the lookup file, so use the default From address
				DLogMsg DL_Debug_Some, "SendFiles: Case-If1-Else: Using default From address"
				strSFFrom = MIGetTaskParam("DI_Email_From")
			End If
		Case Else
			DLogMsg DL_Debug_Some, "SendFiles: Case: Unknown file(s)."
			strSFFrom = MIGetTaskParam("DI_Email_From")
		
	End Select
	strSFBody = strSFBody & Chr(13) & Chr(10) & "[Ref-CustomerPID:" & arrFieldList(4) & "]"
	DLogMsg DL_Debug_Some, "SendFiles: Message from is: " & strSFFrom
	DLogMsg DL_Debug_Some, "SendFiles: Message to is: " & strSFTo
	DLogMsg DL_Debug_Some, "SendFiles: Subject line is: " & strSFSubject
	DLogMsg DL_Debug_Some, "SendFiles: Body is: " & strSFBody

'	Send the email
	Call SendEmail(strSFFrom, strSFTo, strSFSubject, strSFBody)
	
End Sub


' Main program
'
' The overal logic flow here is:
' 1.  Check that the task parameters are set
' 2.  Check that the system objects can be created
' 3.  Make sure that the compress program exists, and that the lookup file can be loaded
' 4.  Process the list of files found by the task Source
' 5.  Cleanup and exit


' Get MOVEit temp file directory
strTempDir = MICacheDir()
' Get list of file paths found by task Source - if this is empty, no files were found
strFileList = MICacheFiles()

' Get the task parameters
strApplication = MIGetTaskParam("DI_Application")
' Location of the data file in which account lookups will be performed. First field is sFTP hostname, second is Account Name, third is Contact first name,
'  fourth is contact last name, fifth is contact email address.
' Value of SFDC Riverside Application field (E7_DataManager, DataDirector, Edusoft)
strLookupFilepath = MIGetTaskParam("DI_Lookup_Filepath")
' Path to 7Zip compression utility executable
strZipExe = MIGetTaskParam("DI_Zip_Filepath")

' Check that task parameters were set - daisy chain all task parameter errors so we don't go through multiple tests to get them set
' Path at which the search for new files starts
If strApplication = "" Then strError = "PRM" : : strErrMsg = strErrMsg & "ERROR-DI_Application parameter has not been set."
If MIGetTaskParam("DI_Base_Directory") = "" Then strError = "PRM" : strErrMsg = "ERROR-DI_Base_Directory parameter has not been set."
If strLookupFilepath = "" Then strError = "PRM" : strErrMsg = strErrMsg & "ERROR-DI_Lookup_Filepath parameter has not been set."
If strZipExe = "" Then strError = "PRM" : strErrMsg = strErrMsg & "ERROR-DI_Zip_Filepath parameter has not been set."
' Default From address for emails
If MIGetTaskParam("DI_Email_From") = "" Then strError = "PRM" : strErrMsg = strErrMsg & "ERROR-DI_Email_From parameter has not been set."
' Destination for emails sent by this script. Will be set to SFDC email-to-case address
If MIGetTaskParam("DI_Email_To") = "" Then strError = "PRM" : strErrMsg = strErrMsg & "ERROR-DI_Email_To parameter has not been set."
' Name of SMTP mail server that will handle emails created by this script
If MIGetTaskParam("DI_Server_Folder") = "" Then strError = "PRM" : strErrMsg = "ERROR-DI_Server_Folder parameter has not been set."
If MIGetTaskParam("DI_Server_Name") = "" Then strError = "PRM" : strErrMsg = "ERROR-DI_Server_Name parameter has not been set."
If MIGetTaskParam("DI_SMTP") = "" Then strError = "PRM" : strErrMsg = strErrMsg & "ERROR-DI_SMTP parameter has not been set."
' If the file list is empty, then no new files were found. This allows the program to exit gracefully.
If strFileList = "" Then strError = "NNF" : strErrMsg = strErrMsg & "OK-NO FILES FOUND."

' Check that external files and objects are accessible
' Can a File System Object be created?
If strError = "" Then
	DLogMsg DL_Debug_Some, "main: Creating FSO object"
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If IsNull(objFSO) Then
		strError = "FSO" : strErrMsg = strErrMsg & "ERROR-File System Object could be not created"
	Else
'		Can account lookup file can be found?
		DLogMsg DL_Debug_Some, "main: Loading lookup file"
		If objFSO.FileExists(strLookupFilepath) Then
			strLookupFilepath = objFSO.GetFile(strLookupFilepath).ShortPath
'			Open an external CSV text file located in the folder specified by the [DI_Lookup_Filepath] parameter.
'			The file is expected to have 5 columns in which the first field is the key value.
'			Open the file
			Set objTS = objFSO.OpenTextFile(strLookupFilepath, ForReading)
			If IsNull(objTS) Then
				DLogMsg DL_Error_App, "main: Could not open file at: " & strLookupFilepath
				strError = "FSO" : strErrMsg = strErrMsg & "Could not open file at: " & strLookupFilepath
			End If

'			Read the contents of the file into a global array so we can perform lookups on it later - use Replace() to strip out double-quotes around fields
			arrAccountList = Split(Replace(objTS.ReadAll,Chr(34),""), Chr(10))
			objTS.Close
			Set objTS = Nothing
			DLogMsg DL_Debug_Some, "main: Number of lookup file rows is: " & CStr(Ubound(arrAccountList))
		Else
			strError = "LKP" : strErrMsg = strErrMsg & "ERROR-Accounts lookup file could not be found at: "& strLookupFilepath
		End If
	End If
End If

' Can a Collaboration Data Objects Message object be created?
If strError = "" Then
	DLogMsg DL_Debug_Some, "main: Creating CDO objects"
	Set objCDO = CreateObject("CDO.Message")
'	Can a Collaboration Data Objects Configuration object be created?
	Set objCDG = CreateObject("CDO.Configuration")
	If IsNull(objCDO) Then strError = "CDO" : strErrMsg = strErrMsg & "ERROR-CDO Message object could be not created"
	If IsNull(objCDG) Then
		strError = "CDO" : strErrMsg = strErrMsg & "ERROR-CDO Configuration object could be not created"
	Else
'		Set configuration field values
		With objCDG.Fields
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MIGetTaskParam("DI_SMTP")
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
		End With
'		Save the values and connect the configuration object to the message object
		objCDG.Fields.Update
		Set objCDO.Configuration = objCDG
		DLogMsg DL_Debug_Some, "main: CDO objects created."
	End If
	DLogMsg DL_Debug_More, "main: Message default server is: " & objCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")
	DLogMsg DL_Debug_More, "main: SMTP Port is: " & objCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")
End If

' Can the compression utility executable file can be found?
If strError = "" Then
	DLogMsg DL_Debug_Some, "main: Looking for Zip executable"
	If objFSO.FileExists(strZipExe) Then
	'	Obtain "safe" path to it
		strZipExe = objFSO.GetFile(strZipExe).ShortPath
	Else
		strError = "ZIP" : strErrMsg = strErrMsg & "ERROR-Compression executable could not be found at: " & strZipExe 
	End If
End If

 DLogMsg DL_Debug_Some, "main: DI_Application param is: " & strApplication
 DLogMsg DL_Debug_Some, "main: DI_Base_Directory param is: " & MIGetTaskParam("DI_Base_Directory")
 DLogMsg DL_Debug_Some, "main: DI_EmailFrom param is: " & MIGetTaskParam("DI_Email_From")
 DLogMsg DL_Debug_Some, "main: DI_EmailTo param is: " & MIGetTaskParam("DI_Email_To")
 DLogMsg DL_Debug_Some, "main: DI_SMTP param is: " & MIGetTaskParam("DI_SMTP")
 DLogMsg DL_Debug_Some, "main: DI_Lookup File param is: " & strLookupFilepath
 DLogMsg DL_Debug_Some, "main: DI_Zip File param is: " & strZipExe
 DLogMsg DL_Debug_Some, "main: DI_Server_Name param is: " & MIGetTaskParam("DI_Server_Name")
 DLogMsg DL_Debug_Some, "main: DI_Server_Folder param is: " & MIGetTaskParam("DI_Server_Folder")
 DLogMsg DL_Debug_Some, "main: Task filelist: " & strFileList
 DLogMsg DL_Debug_Some, "main: Error code is: " & strError

Select Case strError
	Case "CDO"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg

	Case "CDG"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg

	Case "FSO"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg

	Case "LKP"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg

	Case "NNF"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg
	
	Case "PRM"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg

	Case "ZIP"
		DLogMsg DL_Debug_Some, "main: " & strErrMsg

	Case ""
'		Create array of files found by task Source
		arrFileList = Split(strFileList, "|")
'		DLogMsg DL_Debug_Some, "main: Case: Path is1: " & strCurrentPath
'		DLogMsg DL_Debug_Some, "main: Case: Last path is1: " & strLastPath
'		DLogMsg DL_Debug_Some, "main: Case: Initial filelist: " & strFilePaths
'		DLogMsg DL_Debug_Some, "main: Case: Last host is1: " & strLastHost

'		This is the main processing routine. The Do...Loop allows graceful handling of error conditions while processing the file list.
'		An Exit Do ends processing of the current file, which starts the next iteration of the For loop - moving on to the next file..
'		It exits after the first pass, so the process ends normally when the For loop runs out of files
		For intIndex = 0 to UBound(arrFileList)
		Do
'			DLogMsg DL_Debug_Some, "main: Case-For1: Pass " & CStr(intIndex)
'			DLogMsg DL_Debug_Some, "main" Case-For1: Current Host is1: " & strCurrentHost
'			DLogMsg DL_Debug_Some, "main: Case-For1: Last Host is2: " & strLastHost

'			Split file path into an array
			arrPath = Split(arrFileList(intIndex), "\")
'			Max size of array equals depth of file in the path
			intDepth = UBound(arrPath)
'			DLogMsg DL_Debug_Some, "main: Case-For1: Directory depth: " & CStr(intDepth)
'			Catch a top-level file and skip it, but notify someone. The Exit Do allows the script to process the rest of the files without crashing.
			If intDepth < 1 Then
				Call SendEmail(MIGetTaskParam("DI_Email_From"), MIGetTaskParam("DI_Email_To"), "Error processing files", "Error on file: " & MIGetTaskParam("DI_Base_Directory") & "\" & arrFileList(intIndex))
				Exit Do
			End If
'			Intialize variables for the loop
'			Set current host to top-level folder name
			strCurrentHost = arrPath(0)
'			This next For loop builds the path up to the parent directory in which the file is found
'
'			This routine does NOT correctly handle the 'contracts' folder, which is a special case and should be treated as the 
'			top-level folder. Until this is correctly handled, the script will need to be deployed in separate MOVEit tasks
'			for files in the 'e7sftp' folder  and  files in the 'e7sftp/contracts' folder. The task for the 'e7sftp' folder
'			must have the source element set to exclude the 'contracts' folder.
'
			strCurrentPath = strCurrentHost
			For intLoop = 1 to intDepth - 1
				strCurrentPath = strCurrentPath & "\" & arrPath(intLoop)
				DLogMsg DL_Debug_More, "main:Case-For1-For1: CurrentPath is now1: " & strCurrentPath
			Next
			If strLastHost = "" Then
'				This is the first pass through the outer For loop, so set the last host to be the same as the current host.
'				It only needs to be done on the first pass, after that the loop sets it to the current host only when that changes.
				strLastHost = strCurrentHost
				strLastPath = strCurrentPath
				DLogMsg DL_Debug_Some, "main: Case-For1-If2: First pass."
'				(This code block can be used to determine the application type from the file's parent folder, instead of using the DI_Application parameter)
'				This retrieves the name of the folder in which the file is found - should tell us which file type it is
'				strFolderName = arrPath(intLoop - 1)

'				Use the value from the DI_Application task paramter for now
				strFoldername = strApplication				

				DLogMsg DL_Debug_More, "main:Case-For1-If2: Folder_Name is: " & strFolderName
			End If
'			DLogMsg DL_Debug_Some "main: Case-For1: CurrentPath is3: " & strCurrentPath
'			DLogMsg DL_Debug_Some, "main: Case-For1: Last path is2: " & strLastPath
'			DLogMsg DL_Debug_Some, "main: Case-For1: Current Host is2: " & strCurrentHost
'			DLogMsg DL_Debug_Some, "main: Case-For1: Last Host is3: " & strLastHost
'			DLogMsg DL_Debug_Some, "main: Case-For1: Directory depth2: " & CStr(intDepth)
			DLogMsg DL_Debug_Some, "main: Case-For1: Current filelist: " & strFilePaths

'			If the current path equals the last path, we're in the same directory, so add the file to the list
			If strCurrentPath = strLastPath Then
				strFilePaths = strFilePaths & strTempDir & "\" & strLastPath & "\" & arrPath(intDepth) & "|"
				strLastPath = strCurrentPath
				DLogMsg DL_Debug_Some, "main: Case-For1-If3: Filepath is now:  " & strFilePaths
			Else
'				The path has changed, so compress and send these files before moving on
'				Remove trailing | character from file path
				If Right(strFilePaths,1) = "|" Then strFilePaths = Left(strFilePaths, Len(strFilePaths) - 1)
				Call SendFiles(strFilePaths, strLastPath, strLastHost, intIndex)
'				Prepare for next set of files
				strCurrentHost = arrPath(0)
'				This builds the path up to the directory in which the file is found
				strCurrentPath = strCurrentHost
				For intLoop = 1 to intDepth - 1
					strCurrentPath = strCurrentPath & "\" & arrPath(intLoop)
				DLogMsg DL_Debug_More, "main: Case-For1-For-If3-Else-For1: CurrentPath is now2: " & strCurrentPath
				Next
'				(This code block can be used to determine the application type from the file's parent folder, instead of using the DI_Application parameter)
'				This retrieves the name of the folder in which the file is found - should tell us which file type it is
'				strFolderName = arrPath(intLoop - 1)

'				Use the value from the DI_Application task paramter for now
				strFoldername = strApplication				

				DLogMsg DL_Debug_More, "main: Case-For1-If3-Else: Folder_Name is: " & strFolderName
				strLastPath = strCurrentPath
				strFilePaths = strTempDir & "\" & strLastPath & "\" & arrPath(intDepth) & "|"
'				DLogMsg DL_Debug_Some, "main: Case-For1-If3-Else: CurrentPath is4: " & strCurrentPath
'				DLogMsg DL_Debug_Some, "main: Case-For1-If3-Else: Current Host is3: " & strCurrentHost
				strLastHost = strCurrentHost
				DLogMsg DL_Debug_More, "main: Case-For1-If3-Else: CurrentPath is5: " & strCurrentPath
'				DLogMsg DL_Debug_Some, "main: Case-For1-If3-Else: Last host is4: " & strLastHost
			End If
		Loop While False
		Next
'		Remove trailing | character from file path
		If Right(strFilePaths,1) = "|" Then strFilePaths = Left(strFilePaths, Len(strFilePaths) - 1)
		Call SendFiles(strFilePaths, strLastPath, strLastHost, intIndex)
'		DLogMsg DL_Debug_Some, "main: Case: Root Dir: " & strRoot
		DLogMsg DL_OK_File, "main: Case: " & CStr(UBound(arrFileList) + 1) & " FILES PROCESSED."

	Case Else
		DLogMsg DL_OK_File, "main: Case-Else: ERROR-Unhandled error occurred, strError =: " & strError
End Select

Set objFSO = Nothing
Set objCDO = Nothing
Set objCDG = Nothing

DLogMsg DL_OK_File, "main: Number of emails sent: " & CStr(intEmailCount)
DLogMsg DL_OK_File, "main: OK-FILE PROCESSING COMPLETED."
