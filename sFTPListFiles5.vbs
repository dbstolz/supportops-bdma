Option Explicit

' Declare variables
Dim i, arrFileList

' Declare constants
Const DL_Error_App  =  0 
Const DL_Error_Task = 10
Const DL_Warning    = 20
Const DL_OK_Task    = 30
Const DL_OK_File    = 40
Const DL_Debug_Some = 50
Const DL_Debug_More = 60
Const DL_Debug_All  = 70
Const FileSizeMax   = 15360000


'
' Subroutines/Functions
'

Sub DLogMsg(DLevel, Message)
' Logging routine - prints log message only when MOVEit debug level matches value passed in

  if DLevel <= MIGetDebugLevel() then MILogMsg Message
End Sub

Sub PrintArray(arrPArray)
' Only used for debugging - list all values in an array to the log

Dim k
	DLogMsg DL_Debug_More, "PrintArray: Starting"
	For k = 0 to UBound(arrPArray)
		DLogMsg DL_Debug_Some, "PrintArray: Element " & CStr(k) & " is: " & arrPArray(k)
	Next
End Sub


Sub SendEmail(strFilesAttached)

Dim strSubject, strBody, strFileName, strFolder, arrSFileList, objEmail, objConf, obj

	Set objConf = CreateObject("CDO.Configuration")
	If IsNull(objConf) Then
		DLogMsg DL_Error_App, "SendEmail: CDO Configuration object not created"
		MISetErrorDescription("SendEmail: ERROR-CDO Configuration object could be not created")
		MISetErrorCode(2)
		Exit Sub
	End If
	With objConf.Fields
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MIGetTaskParam("DI_SMTP")
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
	End With
	objConf.Fields.Update
	DLogMsg DL_Debug_More, "SendEmail: Send email routine called"
	arrSFileList = Split(strFilesAttached, "|")
'	Call PrintArray(arrSFileList)
	strBody = "The following file(s) were found in the '" & MIGetTaskParam("DI_Application") & "' directory of '" & MIGetTaskParam("DI_Base_Directory") & "'." & Chr(13) & Chr(10) & Chr(13) & Chr(10)
	strSubject = "New file(s) were posted in the " & MIGetTaskParam("DI_Application") & " directory"
	For Each strFileName in arrSFileList
		strBody = strBody & "'" & MIGetTaskParam("DI_Base_Directory") & "\" & strFileName & "'" & Chr(13) & Chr(10)
	Next
	strBody = strBody & Chr(13) & Chr(10)
	DLogMsg DL_Debug_Some, "SendEmail: Subject line is: " & strSubject
	DLogMsg DL_Debug_Some, "SendEmail: Subject body is: " & strBody
	Set objEmail = CreateObject("CDO.Message")
'	For Each obj in objEmail.Fields
'		DLogMsg DL_Debug_More, "SendMail: Field name is: " & obj.Name
'	Next
	If IsNull(objEmail) Then
		DLogMsg DL_Error_App, "SendEmail: CDO Message object not created"
		MISetErrorDescription("SendEmail: ERROR-CDO Message object could be not created")
		MISetErrorCode(2)
		Exit Sub
	End If
	Set objEmail.Configuration = objConf
	Set objConf = Nothing
	DLogMsg DL_Debug_More, "SendEmail: CDO message default server is: " & objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")
	With objEmail
		.From = MIGetTaskParam("DI_Email_From")
		.To = MIGetTaskParam("DI_Email_To")
		.Subject = strSubject
		.TextBody = strBody
		.Send
	End With
	DLogMsg DL_Debug_More, "SendEmail: SMTP Port is: " & objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")
	Set objEmail = Nothing
End Sub

'
' Main program
'
' Get path of source files
' DLogMsg DL_Debug_Some,("main: Getting source path")
' DLogMsg DL_Debug_Some, "main: Full path: " & MIMacro("[FullPath]")
' DLogMsg DL_Debug_Some, "main: Folder name: " & MIMacro("[Foldername]")
' DLogMsg DL_Debug_Some, "main: Orig user: " & MIMacro("[OrigUser]")
' DLogMsg DL_Debug_Some, "main: Cache folder is: " & MICacheDir()
' DLogMsg DL_Debug_Some, "main: Cache filename is: " & MICacheFilename()
' DLogMsg DL_Debug_Some, "main: Cache files are: " & MICacheFiles()
' DLogMsg DL_Debug_Some, "main: Original filename is: " & MIGetOriginalFilename()
' DLogMsg DL_Debug_Some, "main: Directory listing is: " & MIDirGetListing()

' This task requires 5 local task parameters:

' DI_Application is the product folder this task is for. Set it to the name of the folder in the task source, such as datamanager or barcode
' DI_Base_Directory is the path to the top level of the files to be processed. 
'  For DataManager files, set the source to: <server>\<path to top level of district sFTP directories>\**\datamanager\*.*
'  The "** will find all district-level folders and the "\datamanager" will find only files in the 'datamanager' sub-folders
'  For barcode, set the last part to "..\**\barcode"
' DI_Email_From is the apparent sender of the email message. Current default is "DoNotReply@hmhco.com"
' DI_Email_To is the email address to which the email will be sent. Must be a valid format or CDO object will throw an error
' DI_SMTP is the name of the SMTP server that will be used to send the email

' Set the task to Use Original Filenames and set the process that calls this script to run Once after all downloads


DLogMsg DL_Debug_Some, "main: DI_Application parameter is: " & MIGetTaskParam("DI_Application")
DLogMsg DL_Debug_Some, "main: DI_Base_Directory parameter is: " & MIGetTaskParam("DI_Base_Directory")
DLogMsg DL_Debug_Some, "main: DI_Email_From parameter is: " & MIGetTaskParam("DI_Email_From")
DLogMsg DL_Debug_Some, "main: DI_Email_To parameter is: " & MIGetTaskParam("DI_Email_To")
DLogMsg DL_Debug_Some, "main: DI_SMTP parameter is: " & MIGetTaskParam("DI_SMTP")


' Obtain pipe delimited list of all files found and split them up into a new filelist array - each element will be a full file path
If MICacheFiles() <> "" Then
	DLogMsg DL_Debug_Some, "main: Cache filenames are: " & MICacheFiles()
	arrFileList = Split(MICacheFiles(), "|")
'	 Call PrintArray(arrFileList)
	DLogMsg DL_Debug_Some, "main: Count of file list: " & CStr(UBound(arrFileList) + 1)
'	PrintArray arrFileList
	Call SendEmail(MICacheFiles())
	DLogMsg DL_OK_Task, "main: COMPLETED - successfully processed " & CStr(UBound(arrFileList) + 1) & " files."
Else
	DLogMsg DL_OK_Task, "main: NO NEW FILES FOUND."
End If

