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
' boolCDO - Boolean to track whether CDO mail objects were successfully created
' intIndex - Counter variable for outer For loop. Equal to the number of files processed. Also used to give the compressed files a unique filename.
' intLoop - Counter variable for For loop that identifies the parent folder of the current file. Used after loop finished to get previous level directory name.
' intDepth - The deption of the directory in which the current file is found. If less than 2, then we are at the same level as the hostname - not good!
' intEmailCount - Couner variable for the number of email messages sent - informational only
' intError - The type of error. The Case statment handles each error type differently
' strNL - A variable that holds the character(s) that start a new line for use in email and error messages
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
' strServerName - The name of the sFTP server for which the task is configured in the DI_Base_Directory parameter.
' strTopFolder - The top-level directory of the sFTP server, under which all of the account folders are created.
' strBaseDirectory - The path to the local direcotry. This code assumes that the directory will always be mapped to the S: drive letter.
' arrFileList - An array that holds each filepath found by the task source in a separate bucket. The outer For loop uses it to process each file.
' arrPath - An array that holds each directory name in a separate buket. We use it to build the CurrentPath to the file, so we know when the parent folder changes.
' arrAccountList - An array that holds every line from the lookup file. The Lookup function searches it for the hostname, so that it can set values for the iMacro script.
' objFSO - The File System Object used to read the lookup file and copy the zip archive to the 'ticketed' folder of the parent directory.
' objFO - The File Object that is a reference to the lllokup file.
' objCDO - The Collaboration Data Message Object used to create and send the email message
' objCDG - The Collaboration Data Configuration Object used to set the parameters of the CDO Message object
' objTS - The TextStream object used to read the lookup file and write the FileList.txt file used by the compression utility to build the zip archive

' Variables for main program

Dim boolCDO, intIndex, intLoop, intDepth, intEmailCount, intError, strNL, strErrMsg, strCurrentHost, strLastHost, strTempDir, strFileList, strLookupFilepath, _
 strZipExe, strApplication, strCurrentPath, strLastPath, strFilePaths, strFolderName, strServerName, strTopFolder, strBaseDirectory, arrFileList, arrPath, _
 arrAccountList, objFSO, objFO, objCDO, objCDG, objTS

' Functions and Subroutines

Sub DLogMsg(DLevel, Message)

' Logging routine - prints log message only when MOVEit debug level matches value passed in
  If DLevel <= MIGetDebugLevel() then MILogMsg Message
  
End Sub


Sub PrintArray(arrPArray)

' Only used for debugging - list all values in an array to the log

Dim intElement

	DLogMsg DL_Debug_More,  "PrintArray: Starting"
	For intElement = 0 to UBound(arrPArray)
		MILogMsg "PrintArray: Element " & CStr(intElement) & " is: " & arrPArray(intElement)
	Next

End Sub


Function ProductCode(strFolder)

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

'	Sends an email using the paramters passed in and the global CDO object
	DLogMsg DL_Debug_More, "SendEmail: Send email routine called."
	DLogMsg DL_Debug_Some, "SendEmail: Message From is: " & strSEFrom
	DLogMsg DL_Debug_Some, "SendEmail: Message To is: " & strSETo
	DLogMsg DL_Debug_Some, "SendEmail: Subject is: " & strSESubject
	DLogMsg DL_Debug_Some, "SendEmail: Body is: " & strSEBody

'	Make sure email object has been created	
	If IsObject(objCDO) And Not IsNull(objCDO) Then
		With objCDO
			.From = strSEFrom
			.To = strSETo
			.Subject = strSESubject
			.TextBody = strSEBody
			.Send
		End With
		DLogMsg DL_Debug_More, "SendEmail: Message sent."
'		Clear CDO Message fields for the next iteration of the outer For loop
		With objCDO
			.From = ""
			.To = ""
			.Subject = ""
			.TextBody = ""
		End With
	Else
		DLogMsg DL_Error_Task, "SendEmail: CDO object does not exist"
	End If

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

	DLogMsg DL_Debug_More, "LookupAccount: Function called"
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
	For i = UBound(arrFolders) To 0 Step - 1
		DLogMsg DL_Debug_Some, "LookupAccount: For: Directory to search is: " & arrFolders(i)
		intCount = 1
		For Each strLine in arrAccountList
			If strLine = "" Then
				DLogMsg DL_Warning, "LookupAccount: For-For-Each-If1: Line " & CStr(intCount) & " is empty. "
				Exit For
			End If
' 			Row counter
			intCount = intCount + 1
'			Load each field into elements of an array
			arrAccount = Split(strLine, ",")
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: Loop Count: " & intCount
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: Current line of file is: " & strLine
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: First field of current line is: " & arrAccount(0)
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: Length of Hostname is: " & CStr(Len(arrFolders(i)))
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: Length of first field of current line is: " & CStr(Len(arrAccount(0)))
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: First character of first field is: " & CStr(Asc(Left(arrAccount(0),1)))
			DLogMsg DL_Debug_More, "LookupAccount: For-For-Each: Last character of first field is: " & CStr(Asc(Right(arrAccount(0),1)))
			If arrAccount(0) = arrFolders(i) Then
'				The hostname was found
				DLogMsg DL_Debug_Some, "LookupAccount: For-For-Each-If2: Account Found on line " & CStr(intCount - 1)
				DLogMsg DL_Debug_Some, "LookupAccount: For-For-Each-If2: Hostname is: " & arrAccount(0)
				DLogMsg DL_Debug_Some, "LookupAccount: For-For-Each-If2: File field2 is: " & arrAccount(2)
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
				DLogMsg DL_Debug_More, "LookupAccount: For-For-Each-If2-Else: Account not found"
				boolFound = False
			End If
		Next
		If NOT boolFound Then
			DLogMsg DL_Debug_Some, "LookupAccount: For-For-Each-If3: Account not found"
		Else
'			Hostname was found, so exit outer For loop
			Exit For
		End If
	Next
	DLogMsg DL_Debug_Some, "LookupAccount: Loop count is: " & CStr(intCount)

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
' intResult - The value returned by the MIRunCommand function, Zero indicates success, anything else is failure.
' strDestinationFolder - the final location of the zip file after it is created
' arrParentPath - array used to hold each directory of the parent file path. Used by the For loop to build the TempPath.

Dim intIndex, strZipFileName, strTempPath, strCmd, intResult, strDestinationFolder, arrParentPath

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
	strDestinationFolder = strBaseDirectory & "\" & strTempPath  & "\ticketed"
'	Create full path for zip file
	ZipFiles = strDestinationFolder & "\" & strZipFileName
	DLogMsg DL_Debug_Some, "ZipFiles: Filelist is: " & strZFilePaths
	DLogMsg DL_Debug_Some, "ZipFiles: Zipped temp filepath is: " & strTempPath
	DLogMsg DL_Debug_Some, "ZipFiles: ZipFiles is: " & strZipFileName
	DLogMsg DL_Debug_Some, "ZipFiles: Zipped final filepath is: " & ZipFiles

'	Create a temp file to hold the list of files to compress. This will be passed as an argument to the compression utility.
	If Not IsNull(objFSO) Then
		objFSO.CreateTextFile strTempDir & "\FileList.txt"
		Set objTS = objFSO.OpenTextFile(strTempDir & "\FileList.txt", ForWriting)
		If Not IsNull(objTS) Then
			objTS.Write Replace (strZFilePaths, "|", Chr(13))
			objTS.Close
'			Release the memory for the object
			Set objTS = Nothing
		Else
			DLogMsg DL_Error_Task, "ZipFiles: TS object was not created"
		End If

'		Create the command line arugment to invoke the compression utility
		strCmd = strZipExe & " a " & strTempDir & "\" & strZipFileName & " -i@" & strTempDir & "\FileList.txt"
		DLogMsg DL_Debug_Some, "ZipFiles: If1: Command is: " & strCmd
'		Invoke the compression utility to create an archive with the files in it. This can return an error code, so capture it.
		intResult = MIRunCommand(strCmd)
		If intResult <> 0 Then
'			Return an error code to the task so it doesn't delete the files after a problem occurs
			DLogMsg DL_Error_Task, "ZipFiles: ERROR code " & CStr(intResult) & " reported after compression utility called."
			Call SendEmail(MIGetTaskParam("DI_Email_From"), MIGetTaskParam("DI_Error_To"), "Fatal error processing files in task " & MITaskname(), "ZipFiles: Error code " & CStr(intResult) & " reported after compression utility called.")
'			Set error condition so task aborts		
			MISetErrorDescription("ZipFiles: ERROR " & CStr(intResult) & " reported after compression utility called.")
			MISetErrorCode(intResult)
		End If
		DLogMsg DL_Debug_Some, "ZipFiles: Destination folder is: " & strDestinationFolder

'		If the 'ticketed' folder doesn't exist, create it
		If Not(objFSO.FolderExists(strDestinationFolder)) Then
			DLogMsg DL_Debug_Some, "ZipFiles: If2: Trying to create destination folder: " & strDestinationFolder
			objFSO.CreateFolder(strDestinationFolder)
		End If
		objFSO.CopyFile strTempDir & "\" & strZipFileName, ZipFiles
		ZipFiles = strDestinationFolder & "\" & strZipFileName
	Else
		DLogMsg DL_Error_Task, "ZipFiles: FSO object does not exist"
	End If

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

' Declare variables
' intAt - Position of the '@' symbol in the email address (from the start)
' intDot - Position of the '.' symbol in the email address (from the end) 
' strSFCompressedFile - The path to the zip file that is returned from the ZipFiles function
' strFileName - The loop variable for the For..Each loop that proceses each file in the strSFFileList argument. Adds each filename as separate row to the email body.
' strFiles - Temporary string to hold the list of files found for insertion into the email body
' strSFFrom - The email From address sent to the SendEmail function
' strSFTo - The email To address sent to the SendEmail function
' strSFSubject - The email Subject sent to the SendEmail function
' strSFBody - The email Body sent to the SendEmail function
' arrFieldList - Array used to hold each of the fields returned by the LookupAccount function
' arrFileList - Array used to hold each of the filenames passed in via the strSFFileList argument. The For..Each loop uses it to put them in the email body.
' arrFilePath - Array used to hold each directory in the current file path being processed in the For..Each loop. The last bucket is always the filename.

Dim intAt, intDot, strSFCompressedFile, strFileName, strFiles, strSFFrom, strSFTo, strSFSubject, strSFBody, arrFieldList, arrFileList, arrFilePath

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
'	First, build the list of files found that will be inserted into the email body
	For Each strFileName in arrFileList
'		Add only the filename, which will be the last element of the array we split each file path into
		arrFilePath = Split(strFileName,"\")
		strFiles = strFiles & arrFilePath(UBound(arrFilePath)) & strNL
	Next
	strFiles = strFiles & strNL
	DLogMsg DL_Debug_Some, "SendFiles: File list is: " & strFiles

'	Start setting the fields of the email
	strSFTo = MIGetTaskParam("DI_Email_To")
	strSFBody = "New file(s) were found in a '" & strFolderName & "' directory under hostname: " & arrFieldList(0) & strNL & strNL

'	Handle the case when the hostname is not found by giving the email a different subject and inserting a line into the body
	Select Case arrFieldList(1)
		Case ""
			strSFSubject = arrFieldList(0) & " has received new file(s) in a '" & strFolderName & "' directory"
		Case "NOT FOUND"
			strSFSubject = arrFieldList(0) & " has received new file(s) in a '" & strFolderName & "' directory"
			strSFBody = strSFBody & "-------------------------------------------------------------------------------------------------------------------------------------------------" & strNL
			strSFBody = strSFBody & " WARNING: The username was not found in the lookup file - review Salesforce Contact and/or Account records." & strNL
			strSFBody = strSFBody & "-------------------------------------------------------------------------------------------------------------------------------------------------" & strNL & strNL
		Case Else
			strSFSubject = arrFieldList(1) & " has received new file(s) in a '" & strFolderName & "' directory"
	End Select

'	Now insert the list of files we built earlier
	strSFBody = strSFBody & "The following files were found in: " & strTopFolder&  "\" & strSFPath & strNL & strNL
	strSFBody = strSFBody & strFiles

'	We turn file path into a UNC, using the server name and top-level directory name that we derived from the DI_Base_Directory task parameter.
	DLogMsg DL_Debug_Some, "SendFiles: File path is: \\" & strServerName & "\" & Right(strSFCompressedFile, Len(strSFCompressedFile) - InStr(strSFCompressedFile, "\" & strTopFolder))
	strSFBody = strSFBody & "File path is: \\" & strServerName & "\" & Right(strSFCompressedFile, Len(strSFCompressedFile) - InStr(strSFCompressedFile, "\" & strTopFolder)) & strNL & strNL

'	Add the sFTP Username and Account name information from the lookup file
	strSFBody = strSFBody & "sFTP Username: " & arrFieldList(0) & strNL
'	strSFBody = strSFBody & "Account Name: " & arrFieldList(1) & strNL

'	Next, we set the email From address. This differs between Barcode and DataManager orders, which have an Account Holder and Data Contact in SFDC
'	We will use this to also add the DataManager Data Contact person's name and email address
	Select Case LCase(strApplication)
'		We don't have a single point of contact for barcode orders, so use the default email address
		Case "barcode"
			DLogMsg DL_Debug_Some, "SendFiles: Case: Barcode file(s)."
			strSFFrom = MIGetTaskParam("DI_Email_From")
		Case "datamanager"
			DLogMsg DL_Debug_Some, "SendFiles: Case: DataManager file(s)."
			strSFBody = strSFBody & "Account Holder name: " & arrFieldList(2) & strNL
			strSFBody = strSFBody & "Account Holder email: " & arrFieldList(3) & strNL
'			Perform simple format validation of retrieved email address to allow CDO 'from' field to be set
			intAt = InStr(arrFieldList(3),"@")
			intDot = InStrRev(arrFieldList(3),".")
			DLogMsg DL_Debug_Some, "SendFiles: Case: Lookup column 4 value: " & arrFieldList(3)
			DLogMsg DL_Debug_More, "SendFiles: Case: @ found at: " & CStr(intAt)
			DLogMsg DL_Debug_More, "SendFiles: Case: . found at: " & CStr(intDot)
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
			DLogMsg DL_Debug_Some, "SendFiles: Case: Unknown Application: " & strApplication
			strSFFrom = MIGetTaskParam("DI_Email_From")
	End Select

'	Set the email To address and finish the email body with the ISA# and district PID
	DLogMsg DL_Debug_Some, "SendFiles: Message from is: " & strSFFrom
	DLogMsg DL_Debug_Some, "SendFiles: Message to is: " & strSFTo
	DLogMsg DL_Debug_Some, "SendFiles: Subject line is: " & strSFSubject
	DLogMsg DL_Debug_Some, "SendFiles: Body is: " & strSFBody
'	strSFBody = strSFBody & "ISA#: " & arrFieldList(5) & strNL
	strSFBody = strSFBody & "[Ref-CustomerPID:" & arrFieldList(4) & "]"

'	Send the email
	Call SendEmail(strSFFrom, strSFTo, strSFSubject, strSFBody)
	intEmailCount = intEmailCount + 1
	
End Sub


' Main program
'
' The overal logic flow here is:
' 1.  Check that the task parameters are set
' 2.  Check that the system objects can be created, and that the lookup file can be loaded
' 3.  Make sure that the compress program executable can be found
' 4.  Process the list of files found by the task Source
' 5.  Cleanup and exit

' This task requires 9 task parameters:

' DI_Application is the product folder this task is for. Set it to the name of the folder in the task source, such as datamanager or barcode
' DI_Base_Directory is the path to the top level of the files to be processed. 
'  For DataManager files, set the source to: <server>\<path to top level of district sFTP directories>\**\datamanager\*.*
'   odaftpfile01\opvftp_data\users\ia_form_e_cogat_form_7\e7sftp is an example
'  The "** will find all district-level folders and the "\datamanager" will find only files in the 'datamanager' sub-folders
'  For barcode, set the last part to "..\**\barcode"
' DI_Email_From is the apparent sender of the email message. Current default is "DoNotReply@hmhco.com"
' DI_Email_To is the email address to which the email will be sent. Must be a valid format or CDO object will throw an error.
' DI_Error_To is the email address to which the email about errors will be sent.. Must be a valid format or CDO object will throw an error.
' DI_Lookup_Filepath is the location of the CSV lookup file the script uses to insert Salesforce account info into the email.
' DI_SMTP is the name of the SMTP server through which emails will be sent.
' DI_Zip_Filepath is the path the local 7zip executable file that will be used to compress the files.
' DI_SMTP is the name of the SMTP server that will be used to send the email

' The expected structure of the lookup (currently created using SFDC reports) is as follows (the lookup array is zero-indexed):
' Column    Description
' 1         Account BAS SFTP Username
' 2         Account Name
' 3         Contact Full Name (First Last)
' 4         Contact Email address
' 5         Account PID
' 6         Account ISA#



' Store the Carriage Return + Line Feed characters in a variable to make the code look cleaner
strNL = Chr(13) & Chr(10)

' Log the values of the task parameters
DLogMsg DL_Debug_Some, "main: DI_Application parameter is: " & MIGetTaskParam("DI_Application")
DLogMsg DL_Debug_Some, "main: DI_Base_Directory parameter is: " & MIGetTaskParam("DI_Base_Directory")
DLogMsg DL_Debug_Some, "main: DI_Email_From parameter is: " & MIGetTaskParam("DI_Email_From")
DLogMsg DL_Debug_Some, "main: DI_Email_To parameter is: " & MIGetTaskParam("DI_Email_To")
DLogMsg DL_Debug_Some, "main: DI_Error_To parameter is: " & MIGetTaskParam("DI_Error_To")
DLogMsg DL_Debug_Some, "main: DI_Lookup_Filepath parameter is: " & MIGetTaskParam("DI_Lookup_Filepath")
DLogMsg DL_Debug_Some, "main: DI_SMTP parameter is: " & MIGetTaskParam("DI_SMTP")
DLogMsg DL_Debug_Some, "main: DI_Zip_Filepath parameter is: " & MIGetTaskParam("DI_Zip_Filepath")

' Get MOVEit temp file directory
strTempDir = MICacheDir()

' Get list of file paths found by task Source - if this is empty, no files were found
strFileList = MICacheFiles()
DLogMsg DL_Debug_Some, "main: Task filelist: " & strFileList

' If the file list is empty, then no new files were found. This allows the program to exit gracefully.
If strFileList = "" Then
	DLogMsg DL_OK_Task, "main: OK-NO FILES FOUND."
Else

' 	Check that task parameters were set - daisy chain all task parameter errors so we don't go through multiple tests to get them set
'	Value of SFDC Riverside Application field (E7_DataManager, DataDirector, Edusoft)
	If MIGetTaskParam("DI_Application") = "" Then intError = 1 : strErrMsg = "ERROR-DI_Application parameter has not been set."
'	Path at which the search for new files starts
	If MIGetTaskParam("DI_Base_Directory") = "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_Base_Directory parameter has not been set."
'	Default From address for emails
	If MIGetTaskParam("DI_Email_From") = "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_Email_From parameter has not been set."
'	Destination for case-creation emails sent by this script. Will be set to SFDC email-to-case address
	If MIGetTaskParam("DI_Email_To") = "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_Email_To parameter has not been set."
'	Destination for error emails sent by this script. Will be set to SFDC email-to-case address
	If MIGetTaskParam("DI_Error_To") = "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_Error_To parameter has not been set."
'	Location of the data file in which account lookups will be performed. First field is sFTP hostname, second is Account Name, third is Contact first name,
'	 fourth is contact last name, fifth is contact email address.
	If MIGetTaskParam("DI_Lookup_Filepath")= "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_Lookup_Filepath parameter has not been set."
'	Path of the parent directory in which the files are found
'	Name of SMTP mail server that will handle emails created by this script
	If MIGetTaskParam("DI_SMTP") = "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_SMTP parameter has not been set."
'	Path to 7zip compression utility executable file
	If MIGetTaskParam("DI_Zip_Filepath") = "" Then intError = 1 : strErrMsg = strErrMsg & strNL & "ERROR-DI_Zip_Filepath parameter has not been set."
'	Log error message if param(s) missing
	If intError > 0 Then DLogMsg DL_Error_Task, "main: Error code is: " & CStr(intError)

'	Initialize variables from the task parameters
	strApplication = MIGetTaskParam("DI_Application")
	strLookupFilepath = MIGetTaskParam("DI_Lookup_Filepath")
	strZipExe = MIGetTaskParam("DI_Zip_Filepath")

'	Derive the server name and top-level directory name from the DI_Base_Directory parameter
'	Split the parameter into an array
	strBaseDirectory = MIGetTaskParam("DI_Base_Directory")
	arrPath = Split(strBaseDirectory, "\")
'	Server name is the first entry, and the top-level directory is the last
'	We assume that the path to the files will be mapped to the local S: drive letter
	strServerName = arrPath(0)
	strTopFolder = arrPath(UBound(arrPath))
	strBaseDirectory = "S:\" & Right(strBaseDirectory, Len(strBaseDirectory) - InStr(strBaseDirectory, "\"))
	DLogMsg DL_Debug_Some, "main: Server name: " & strServerName
	DLogMsg DL_Debug_Some, "main: Top-level directory name: " & strTopFolder
	DLogMsg DL_Debug_Some, "main: Local base directory path: " & strBaseDirectory

'	Check that external files and objects are accessible

	If intError = 0 Then
'		Can a Collaboration Data Objects Message object be created?
		DLogMsg DL_Debug_More, "main: Creating CDO objects"
		Set objCDO = CreateObject("CDO.Message")
'		Can a Collaboration Data Objects Configuration object be created?
		Set objCDG = CreateObject("CDO.Configuration")
		If IsNull(objCDO) Then intError = 2 : strErrMsg = strErrMsg & strNL & "ERROR-CDO Message object could be not created"
		If IsNull(objCDG) Then
			intError = 2 : strErrMsg = strErrMsg & strNL & "ERROR-CDO Configuration object could be not created."
		Else
'			Set configuration field values
			With objCDG.Fields
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MIGetTaskParam("DI_SMTP")
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
			End With
'			Save the values and connect the configuration object to the message object
			objCDG.Fields.Update
			Set objCDO.Configuration = objCDG
			Set objCDG = Nothing
			DLogMsg DL_Debug_Some, "main: CDO objects created."
		End If
		DLogMsg DL_Debug_More, "main: Message default server is: " & objCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")
		DLogMsg DL_Debug_More, "main: SMTP Port is: " & objCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")
	End If

	If intError = 0 Then
'		Can a File System Object be created?
		DLogMsg DL_Debug_More, "main: Creating FSO object"
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If Not IsObject(objFSO) Or IsNull(objFSO) Then
			intError = 2 : strErrMsg = strErrMsg & strNL & "ERROR-File System Object could be not created"
		Else
'			Can the lookup file can be found?
			If objFSO.FileExists(strLookupFilepath) Then
				DLogMsg DL_Debug_More, "main: Lookup file found"
'				Can we read the lookup file?
				Set objFO = objFSO.GetFile(strLookupFilepath)
				If IsNull(objFO) Then
					DLogMsg DL_Error_Task, "main: Could not open lookup file at: " & strLookupFilepath
					intError = 2 : strErrMsg = strErrMsg & strNL & "ERROR-File System Object could be not created"
				End If
'				Open an external CSV text file located in the folder specified by the [DI_Lookup_Filepath] parameter.
'				The file is expected to have 6 columns in which the first field is the key value.
'				Open the file
				Set objTS = objFO.OpenAsTextStream(ForReading)
				If IsNull(objTS) Then
					DLogMsg DL_Error_Task, "main: Could not open lookup file at: " & strLookupFilepath
					intError = 3 : strErrMsg = strErrMsg & strNL & "ERROR-Could not open lookup file at: " & strLookupFilepath
				End If

'				Read the contents of the file into a global array so we can perform lookups on it later - use Replace() to strip out double-quotes around fields
				arrAccountList = Split(Replace(objTS.ReadAll,Chr(34),""), Chr(10))
				objTS.Close
'				Release the memory for the objects
				Set objTS = Nothing
				Set objFO = Nothing
				DLogMsg DL_Debug_More, "main: Number of lookup file rows is: " & CStr(Ubound(arrAccountList))
			Else
				intError = 3 : strErrMsg = strErrMsg & strNL & "ERROR-Lookup file could not be found at: "& strLookupFilepath
			End If
			DLogMsg DL_Debug_Some, "main: Lookup file loaded."
		End If
	End If

	If intError = 0 Then
'	 Can the compression utility executable file can be found?
		DLogMsg DL_Debug_More, "main: Looking for Zip executable"
		If objFSO.FileExists(strZipExe) Then
'			Obtain "safe" path to it
			strZipExe = objFSO.GetFile(strZipExe).ShortPath
		Else
			DLogMsg DL_Error_Task, "main: Compression executable could not be found."
			intError = 3 : strErrMsg = strErrMsg & strNL & "ERROR-Compression executable could not be found at: " & strZipExe
		End If
	End If

'	Main processing routine - only process files if the error condition is zero
	If intError = 0 Then
'		Create array of files found by task Source. The For..Loop will walk this array to process each file.
		arrFileList = Split(strFileList, "|")
'		DLogMsg DL_Debug_Some, "main: Path is1: " & strCurrentPath
'		DLogMsg DL_Debug_Some, "main: Last path is1: " & strLastPath
'		DLogMsg DL_Debug_Some, "main: Initial filelist: " & strFilePaths
'		DLogMsg DL_Debug_Some, "main: Last host is1: " & strLastHost

'		This is the file processing code. The Do...Loop allows graceful handling of error conditions while processing the file list.
'		An Exit Do ends processing of the current file, which starts the next iteration of the For loop - moving on to the next file.
'		It exits after the first pass, so the process ends normally when the For loop runs out of files.
		For intIndex = 0 to UBound(arrFileList)
		Do
			DLogMsg DL_Debug_More, "main: For-Do: Pass " & CStr(intIndex)
			DLogMsg DL_Debug_More, "main: For-Do: Current Host is1: " & strCurrentHost
			DLogMsg DL_Debug_More, "main: For-Do: Last Host is2: " & strLastHost

'			Split file path into an array - the elements of the array hold the names of each directory in the path
			arrPath = Split(arrFileList(intIndex), "\")
'			Max size of array equals depth of file in the path
			intDepth = UBound(arrPath)
			DLogMsg DL_Debug_More, "main: For-Do: Directory depth: " & CStr(intDepth)
			DLogMsg DL_Debug_More, "main: For-Do: Parent directory: " & arrPath(intDepth - 1)

'			Catch a top-level file and skip it, but notify someone. The Exit Do allows the script to process the rest of the files without crashing.
			If intDepth < 1 Then
				DLogMsg DL_Warning, "main: For-Do-If1: Top level file found: " & strBaseDirectory & "\" & arrFileList(intIndex)
				Call SendEmail(MIGetTaskParam("DI_Email_From"), MIGetTaskParam("DI_Error_To"), "Non-Fatal error processing files in task " & MITaskname(), "Top level file found: " & strBaseDirectory & "\" & arrFileList(intIndex))
				Exit Do
			End If

'			Catch a mismatch between the DI_Application task parameter and the parent folder name.
'			This won't make the script crash, but should be treated as a fatal error condition since it reveals a disconnect
'			 between the parameter value and task Source.
			If arrPath(intDepth - 1) <> strApplication Then
				intError = 4
				strErrMsg = "ERROR-DI_Application parameter '" & strApplication & "' does not match current directory " & arrPath(intDepth - 1)
				DLogMsg DL_Error_Task, "main: For-Do-If2: " & strErrMsg
				Call SendEmail(MIGetTaskParam("DI_Email_From"), MIGetTaskParam("DI_Error_To"), "Fatal error processing files in task " & MITaskname(), strErrMsg)
'				Set error condition so task aborts		
				MISetErrorDescription(strErrMsg)
				MISetErrorCode(intError)
				Exit Do			
			End If

'			Intialize variables for the loop
'			Set current host to top-level folder name
			strCurrentHost = arrPath(0)
			DLogMsg DL_Debug_Some, "main: For-Do: arrPath(0): " & strCurrentHost
'			This next For..Next loop builds the path up to the parent directory in which the file is found
			strCurrentPath = strCurrentHost
			For intLoop = 1 to intDepth - 1
				strCurrentPath = strCurrentPath & "\" & arrPath(intLoop)
				DLogMsg DL_Debug_More, "main: For-Do-For1: CurrentPath is now1: " & strCurrentPath
			Next
			If strLastHost = "" Then
'				This is the first pass through the outer For loop, so set the last host to be the same as the current host.
'				It only needs to be done on the first pass, after that the loop sets it to the current host only when that changes.
				DLogMsg DL_Debug_Some, "main: For-Do-If3: First pass."
				strLastHost = strCurrentHost
				strLastPath = strCurrentPath

'				(This code block can be used to determine the application type from the file's parent folder, instead of using the DI_Application parameter)
'				This retrieves the name of the folder in which the file is found - should tell us which application type it is
'				strFolderName = arrPath(intLoop - 1)
'				Use the value from the DI_Application task paramter for now
				strFoldername = strApplication

				DLogMsg DL_Debug_Some, "main: For-Do-If3: Folder_Name is: " & strFolderName
			End If
			DLogMsg DL_Debug_More, "main: For-Do: CurrentPath is3: " & strCurrentPath
			DLogMsg DL_Debug_More, "main: For-Do: Last path is2: " & strLastPath
			DLogMsg DL_Debug_More, "main: For-Do: Current Host is2: " & strCurrentHost
			DLogMsg DL_Debug_More, "main: For-Do: Last Host is3: " & strLastHost
			DLogMsg DL_Debug_More, "main: For-Do: Directory depth2: " & CStr(intDepth)
			DLogMsg DL_Debug_Some, "main: For-Do: Current filelist: " & strFilePaths

'			If the current path equals the last path, we're in the same directory, so append the file to the list
			If strCurrentPath = strLastPath Then
				strFilePaths = strFilePaths & strTempDir & "\" & strLastPath & "\" & arrPath(intDepth) & "|"
				strLastPath = strCurrentPath
				DLogMsg DL_Debug_Some, "main: For-Do-If4: Filepath is now:  " & strFilePaths
			Else
'				The path has changed, so compress and send the files in the list before moving on
'				Remove the trailing | character from the file path
				If Right(strFilePaths,1) = "|" Then strFilePaths = Left(strFilePaths, Len(strFilePaths) - 1)
				Call SendFiles(strFilePaths, strLastPath, strLastHost, intIndex)
'				Prepare for next set of files
				strCurrentHost = arrPath(0)
'				This builds the path up to the directory in which the file is found
				strCurrentPath = strCurrentHost
				For intLoop = 1 to intDepth - 1
					strCurrentPath = strCurrentPath & "\" & arrPath(intLoop)
				DLogMsg DL_Debug_More, "main: For-Do-If4-Else-For: CurrentPath is now2: " & strCurrentPath
				Next

'				(This code block can be used to determine the application type from the file's parent folder, instead of using the DI_Application parameter)
'				This retrieves the name of the folder in which the file is found - should tell us which file type it is
'				strFolderName = arrPath(intLoop - 1)
'				Use the value from the DI_Application task paramter for now
				strFoldername = strApplication

				DLogMsg DL_Debug_More, "main: For-Do-If4-Else: Folder_Name is: " & strFolderName
				strLastPath = strCurrentPath
				strFilePaths = strTempDir & "\" & strLastPath & "\" & arrPath(intDepth) & "|"
				DLogMsg DL_Debug_More, "main: For-Do-If4-Else: CurrentPath is4: " & strCurrentPath
				DLogMsg DL_Debug_More, "main: For-Do-If4-Else: Current Host is3: " & strCurrentHost
				strLastHost = strCurrentHost
				DLogMsg DL_Debug_More, "main: For-Do-If4-Else: CurrentPath is5: " & strCurrentPath
'				DLogMsg DL_Debug_Some, "main: For-Do-If4-Else: Last host is4: " & strLastHost
			End If
		Loop While False

'		If an error condition was set inside the loop, stop processing files and exit the For..Next
		If intError <> 0 Then Exit For
		Next

'		Only send an email about the last set of files if an error condition does not exist
		If intError = 0 Then 
'			Remove the trailing | character from the file path
			If Right(strFilePaths,1) = "|" Then strFilePaths = Left(strFilePaths, Len(strFilePaths) - 1)
			Call SendFiles(strFilePaths, strLastPath, strLastHost, intIndex)
'			Log the number of files processed
			DLogMsg DL_OK_Task, "main: " & CStr(UBound(arrFileList) + 1) & " FILES PROCESSED."
		End If
	Else
'		If the error code is non-zero, report it and abort the task
		DLogMsg DL_Error_Task, "main: " & strErrMsg
		Call SendEmail(MIGetTaskParam("DI_Email_From"), MIGetTaskParam("DI_Error_To"), "Fatal error processing files in task " & MITaskname(), strErrMsg)
'		Set error condition so task aborts		
		MISetErrorDescription(strErrMsg)
		MISetErrorCode(intError)
	End If

'	Cleanup by releasing memory for objects
	Set objFSO = Nothing
	Set objCDO = Nothing

End If

' Log the number of emails sent
If intEmailCount > 0 Then DLogMsg DL_OK_Task, "main: Number of emails sent: " & CStr(intEmailCount)

DLogMsg DL_OK_Task, "main: FILE PROCESSING COMPLETED."
