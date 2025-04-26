'***********************************************************************************************************************
'* Joplin to PDF Converter
'*
'* Created By  : Anthonel S. Eugenio
'* Date Created: 29-Mar-2025
'***********************************************************************************************************************
Option Explicit

'*----------------------------------------------------------------------------------------------------------------------
'* Variable Declarations
'*----------------------------------------------------------------------------------------------------------------------
Dim objFSO, objShellApp, objWScriptShell
Dim objNow, objTimeStampStart, objTimeStampEnd
Dim objTempInputFolder, objCleanedUpInputFolder, objLogsFolder, objOutputFolder, objInputFolder, objInputFile
Dim objCSVStream, objNewMarkdownFile, objResourceFile, objDestination
Dim objFiles, objSubFolder, objFile

Dim conMsgBoxTitle, conBatchFile

Dim strScriptDir, strBatchFile, str7zPath, strPath, strFileLocation, strFileName
Dim strTitle, strMsg, strInput, strInputFile, strOutputDirectory
Dim strDate, strTime, strFormattedDateTime
Dim strTempInputFolderName, strTempInputFolderPath, strTempCleanedUpFolderName, strTempCleanedUpFolderPath
Dim strTempLogsFolderName, strTempLogsFolderPath, strOutputFolderName, strOutputFolderPath
Dim strCommand
Dim strResourcesDirectory, strResourceFolder, strResourcesParentDirectory
Dim strCSVFileName, strCSVFullFileName
Dim strFileExtension, strTag, strResourceReplacement, strDetails, strFirstLine
Dim strNewMarkdownFile, strLine, strAttachmentName, strID, strParentID, strType, strNewFileName, strObjType
Dim strCSVLine, strText, strFolderName, strParentFolderName
Dim strMarkdownName, strMarkdownNameFullName, strDestinationPath, strOutputName, strOutputFullName
Dim strLogFile, strTempSearchFile, strNewFolderPath, strNewFolderName
Dim strConversionLogFileName, strConversionLogFullFileName, strTXTLine
Dim strSummaryFileName, strSummaryFullFileName

Dim arrPossiblePaths

Dim lngPos, lngStartPos, lngEndPos, lngFileLength, lngTotalFileCount, lngCounter, lngMaxJobs, lngJobCounter
Dim lngNumberOfLogs, lngDuration, lngMinutes, lngSeconds

Dim blnMsg, blnFileFound, blnDirectoryFound, blnProceed, blnCapture, blnExitLoop, blnFolderRenamed

'*----------------------------------------------------------------------------------------------------------------------
'* Initialization
'*----------------------------------------------------------------------------------------------------------------------

Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

Set objShellApp = CreateObject("Shell.Application")

Set objWScriptShell = CreateObject("WScript.Shell")
objWScriptShell.CurrentDirectory = strScriptDir

conMsgBoxTitle = "Joplin to PDF Bulk Converter"

Const ForReading = 1

'*----------------------------------------------------------------------------------------------------------------------
'* Check for Dependencies
'*----------------------------------------------------------------------------------------------------------------------

'Check the Windows Batch File.
conBatchFile = "MarkdownToPDF.bat" 'File Name of the Windows Batch File (*.bat)
strBatchFile = strScriptDir & "\" & conBatchFile
If Not objFSO.FileExists(strBatchFile) Then
	strTitle = conMsgBoxTitle & " | Missing Component"
	strMsg = "Batch file " & Trim(conBatchFile) & " not found." & vbCrLf & _
			"Please ensure the said file resides with this script."
	blnMsg = MsgBox(strMsg, vbOKOnly + vbInformation, strTitle)
	WScript.Quit
End If

'Check if the 7z program is installed in the system.
arrPossiblePaths = Array("C:\Program Files\7-Zip\7z.exe", "C:\Program Files (x86)\7-Zip\7z.exe")
str7zPath = ""
For Each strPath In arrPossiblePaths
    If objFSO.FileExists(strPath) Then
        str7zPath = strPath
        Exit For
    End If
Next

If str7zPath = "" Then
    strTitle = conMsgBoxTitle & " | Missing Component"
	strMsg = "The Program " & Chr(34) & "7-Zip File Manager" & Chr(34) & " is not installed or found in this PC. " & vbCrLf & _
			    "Please install the program first Then try again."
	blnMsg = MsgBox(strMsg, vbOKOnly + vbInformation, strTitle)
	WScript.Quit
End If

'***********************************************************************************************************************

'*----------------------------------------------------------------------------------------------------------------------
'* Get the JEX file.
'*----------------------------------------------------------------------------------------------------------------------
strTitle = conMsgBoxTitle
strMsg = "[Enter Joplin File]" & vbCrLf & vbCrLf & _
            "Please enter file name of the Joplin export file (e.g. jex-export.jex) with its full path."
strInput = ""
blnFileFound = False
Do While Not blnFileFound
	blnFileFound = True
	strInput = InputBox(strMsg, strTitle, strInput)
	If IsEmpty(strInput) Then
		strMsg = "Operation has been cancelled."
		blnMsg = MsgBox(strMsg, vbOKOnly + vbInformation, strTitle)
		WScript.Quit
	ElseIf Trim(strInput) = "" Then
        blnFileFound = False
        strMsg = "Enter Joplin File" & vbCrLf & vbCrLf & _
                   "Please enter file name of the Joplin export file (e.g. jex-export.jex) with its full path."
    Else
		If Not objFSO.FileExists(strInput) Then
			blnFileFound = False
            lngPos = InStr(strInput, "/")
            If lngPos = 0 Then
                strFileLocation = strScriptDir
            Else
                strFileLocation = Left(strInput, lngPos - 1)
            End If
            lngPos = lngPos + 1
            strFileName = Mid(strInput, lngPos, Len(strInput) - lngPos + 1)
            
			strMsg = "File " & Chr(34) & strFileName & Chr(34) & " not found in " & vbCrLf & _ 
                        Chr(34) & Trim(strFileLocation) & Chr(34) & "." & vbCrLf & vbCrLf & _
					    "Please enter file name of the Joplin export file (e.g. jex-export.jex) with its full path."
		Else
			strInputFile = objFSO.GetAbsolutePathName(strInput)
		End If
	End If
Loop

'*----------------------------------------------------------------------------------------------------------------------
'* Get the output folder/directory.
'*----------------------------------------------------------------------------------------------------------------------
strTitle = conMsgBoxTitle & " | Enter Directory"
strMsg = "[Enter Output Directory]" & vbCrLf & vbCrLf & _
            "Please enter the location where the converted files will be saved."
strInput = ""
blnDirectoryFound = False
Do While Not blnDirectoryFound
	blnDirectoryFound = True
	strInput = InputBox(strMsg, strTitle, strInput)
	If IsEmpty(strInput) Then
		strMsg = "Operation has been cancelled."
		blnMsg = MsgBox(strMsg, vbOKOnly + vbInformation, strTitle)
		WScript.Quit
	Else
		If Not objFSO.FolderExists(strInput) Then
			blnDirectoryFound = False
			strMsg = "Directory " & Chr(34) & Trim(strInput) & Chr(34) & " not found." & vbCrLf & vbCrLf & _
					    "Please enter the location where the converted files will be saved."
		Else
			strOutputDirectory = strInput
		End If
	End If
Loop

'*----------------------------------------------------------------------------------------------------------------------
'* Display Confirmation
'*----------------------------------------------------------------------------------------------------------------------
strTitle = conMsgBoxTitle
strMsg = "[Confirm Operation]" & vbCrLf & vbCrLf & _
            "Processing will start with the following details:" & vbCrLf & vbCrLf & _
			"Input File: " & Chr(34) & strInputFile & Chr(34) & vbCrLf & vbCrLf & _
			"Output Directory: " & Chr(34) & strOutputDirectory & Chr(34) &vbCrLf & vbCrLf & _
			"A message will be displayed once the conversion process is complete." & vbCrLf & vbCrLf & _
            "Select OK to proceed."
blnProceed = MsgBox(strMsg, vbOKCancel + vbInformation, strTitle)
If blnProceed = vbCancel Then
	strMsg = "Operation has been cancelled."
	blnMsg = MsgBox(strMsg, vbOKOnly + vbInformation, strTitle)
	WScript.Quit
End If

'***********************************************************************************************************************
'* Main Process
'***********************************************************************************************************************

'*----------------------------------------------------------------------------------------------------------------------
'* Create a Temporary Hidden folder for the following:
'*  - Extracted Markdown files from the .jex file.
'*  - Cleaned up version of Markdown files.
'*  - Final Output.
'* Folder name is based on Timestamp.
'*----------------------------------------------------------------------------------------------------------------------
objNow = Now()
objTimeStampStart = objNow
strDate = Year(objNow) & Right("0" & Month(objNow), 2) & Right("0" & Day(objNow), 2) 'Format Date
strTime = Right("0" & Hour(objNow), 2) & Right("0" & Minute(objNow), 2) & Right("0" & Second(objNow), 2) 'Format Time
strFormattedDateTime = strDate & strTime 'Combine Date and Time
'JTPI: Joplin To PDF Input
strTempInputFolderName = "JTPI" & strFormattedDateTime 'Joplin To PDF Input
strTempInputFolderPath = Trim(strOutputDirectory) & "\" & Trim(strTempInputFolderName)
Set objTempInputFolder = objFSO.CreateFolder(strTempInputFolderPath)
objTempInputFolder.Attributes = objTempInputFolder.Attributes + 2 '2 represents hidden attribute
'JTPC: Joplin to PDF Cleaned-up
strTempCleanedUpFolderName = "JTPC" & strFormattedDateTime 'Joplin To PDF Input
strTempCleanedUpFolderPath = Trim(strOutputDirectory) & "\" & Trim(strTempCleanedUpFolderName)
Set objCleanedUpInputFolder = objFSO.CreateFolder(strTempCleanedUpFolderPath)
objCleanedUpInputFolder.Attributes = objCleanedUpInputFolder.Attributes + 2 '2 represents hidden attribute
'JTPL: Joplin to PDF Conversion Log Files
strTempLogsFolderName = "JTPL" & strFormattedDateTime 'Joplin To PDF Input
strTempLogsFolderPath = Trim(strOutputDirectory) & "\" & Trim(strTempLogsFolderName)
Set objLogsFolder = objFSO.CreateFolder(strTempLogsFolderPath)
objLogsFolder.Attributes = objLogsFolder.Attributes + 2 '2 represents hidden attribute
'JTP: Joplin to PDF Final Output.
strOutputFolderName = "JTP_" & strFormattedDateTime 'Joplin To PDF Input
strOutputFolderPath = Trim(strOutputDirectory) & "\" & Trim(strOutputFolderName)
Set objOutputFolder = objFSO.CreateFolder(strOutputFolderPath)
objOutputFolder.Attributes = objOutputFolder.Attributes + 2 '2 represents hidden attribute

'*----------------------------------------------------------------------------------------------------------------------
'* Extract the .jex file to the temporary folder
'*----------------------------------------------------------------------------------------------------------------------
strCommand = """" & str7zPath & """ x """ & strInputFile & """ -o""" & strTempInputFolderPath & """ -y"
objWScriptShell.Run strCommand, 0, True

'*----------------------------------------------------------------------------------------------------------------------
'* Set up the resources folder.
'*----------------------------------------------------------------------------------------------------------------------
strResourcesDirectory = strTempInputFolderPath & "\" & "resources"
If Not objFSO.FolderExists(strResourcesDirectory) Then
    strResourcesDirectory = ""
End If

strResourceFolder = Replace(strResourcesDirectory, strTempInputFolderPath, "")
If Left(Trim(strResourceFolder), 1) = "\" Then
    strResourceFolder = Replace(strResourceFolder, "\", "", 1, 1)
End If

strResourcesParentDirectory = Replace(strResourcesDirectory, "\" & strResourceFolder, "")

'*----------------------------------------------------------------------------------------------------------------------
'* Gather the Markdown Files to: 
'*  - Get its details based on the Joplin specific tags, e.g. which is a File or a Folder. 
'*  - Store all the details in a CSV file.
'*  - Remove the Joplin specific tags and save the updated markdown file to a separate temporary folder.
'*----------------------------------------------------------------------------------------------------------------------
Dim objTextStream

'Specify the CSV file path (change this to your desired CSV file path)
strCSVFileName = "List_" & strFormattedDateTime & ".csv"
strCSVFullFileName = strOutputFolderPath & "\" & strCSVFileName

'Create or open the CSV file for writing
If objFSO.FileExists(strCSVFullFileName) Then
	objFSO.DeleteFile(strCSVFullFileName)
End If
Set objCSVStream = objFSO.CreateTextFile(strCSVFullFileName, True)

strFileExtension = "md"
strTag = "(:/"
strResourceReplacement = "(" & Trim(strResourceFolder) & "\"

Set objInputFolder = objFSO.GetFolder(strTempInputFolderPath)

'Loop through each file in the folder
For Each objInputFile In objInputFolder.Files

	If LCase(objFSO.GetExtensionName(objInputFile.Path)) = strFileExtension Then
		'Open the file for reading
		Set objTextStream = objInputFile.OpenAsTextStream(1)
		'Get the file name (without extension)
        strFileName = objFSO.GetBaseName(objInputFile.Name)
		
		'Initialize variables for capturing details
        blnCapture = False
        strDetails = ""
		strFirstLine = ""
		
		'Create the Output Markdown File
		strNewMarkdownFile = strTempCleanedUpFolderPath & "\" & strFileName & ".md"
		If objFSO.FileExists(strNewMarkdownFile) Then
			objFSO.DeleteFile(strNewMarkdownFile)
		End If
		Set objNewMarkdownFile = objFSO.CreateTextFile(strNewMarkdownFile, True)		
		
		'Read the file line by line
        Set objTextStream = objInputFile.OpenAsTextStream(1)
        Do Until objTextStream.AtEndOfStream
            strLine = objTextStream.ReadLine
            
			'Capture the first line to get the Title
            If strFirstLine = "" Then
				strFirstLine = strLine
				strFirstLine = ReplaceSpecialChars(strFirstLine)
            End If
			
			'Check if there's an attachment tag/characters.
			lngPos = InStr(strLine, strTag)
			If lngPos > 0 Then
				strLine = Replace(strLine, "(:/", strResourceReplacement)
                'Search the File in the resources folder.
                lngStartPos = InStr(strLine, strResourceReplacement) + Len(Trim(strResourceReplacement)) 
                lngEndPos = InStr(lngStartPos, strLine, ")") - 1
                lngFileLength = lngEndPos - lngStartPos + 1
                strAttachmentName = Mid(strLine, lngStartPos, lngFileLength)
                strLine = Replace(strLine, strAttachmentName, GetNameWithExtension(strResourcesDirectory, strAttachmentName)) & "{ width=60% }" & vbLf
                
				objNewMarkdownFile.Write strLine
			End If
			
			Do
				If lngPos > 0 Then Exit Do
				
				'Check if the line starts with the "id:" tag
				If Left(strLine, 3) = "id:" And InStr(strLine,strFileName) > 0 Then
                    strID = Trim(Mid(strLine,InStr(strLine,":")+1))
					blnCapture = True
				End If
				
				'Capture lines until the "type_:" tag
				If blnCapture Then
					If strDetails <> "" Then
						strDetails = strDetails & "|"
					End If
					strDetails = strDetails & strLine
				End If
				
                'Capture other details
                '---------------------
                'parent_id
                If Left(strLine, 10) = "parent_id:" Then
                    strParentID = Trim(Mid(strLine,InStr(strLine,":")+1))
                End If

				'Stop capturing after the "type_:" tag
				If Left(strLine, 6) = "type_:" Then
					strType = Trim(Mid(strLine,InStr(strLine,":")+1))
					blnCapture = False
				End If
				
				If blnCapture Or Left(strLine, 6) = "type_:" Then Exit Do
				
				'Write the Output
				strLine = strLine & vbLf
				objNewMarkdownFile.Write strLine
			Loop While False
            
        Loop
        objTextStream.Close
        objNewMarkdownFile.Close
        Set objNewMarkdownFile = Nothing
		
        'If it's a folder, delete the converted file.
        If strType <> "1" Then
            'Delete the File.
            objFSO.DeleteFile(strNewMarkdownFile)
        End If
            
		'Write the file name and details to the CSV file
		strNewFileName = strFirstLine
        If Len(Trim(strNewFileName)) > 30 Then
            strNewFileName = Left(strNewFileName, 30)
        End If
        strObjType = "File"
        If Trim(strType) = "2" Then
            strObjType = "Folder"
        ElseIf Trim(strType) = "1" Then
            strNewFileName = Chr(34) & Trim(strNewFileName) & ".pdf" & Chr(34)
        ElseIf Trim(strType) = "4" Then
            'Direct, see sample 0312704d3cd44f01be5bdc97f411e81a.md. This is 0312704d3cd44f01be5bdc97f411e81a.jpg
        End If

		strFileName = Chr(34) & strFileName & ".md" & Chr(34)
        strCSVLine = "File Name: " & strFileName & "|" & "Object Type: " & strObjType & "|" & "New Name: " & strNewFileName & "|" & strDetails
        objCSVStream.WriteLine strCSVLine
	End If
	
Next
objCSVStream.Close
Set objCSVStream = Nothing
Set objInputFolder = Nothing

'*----------------------------------------------------------------------------------------------------------------------
'* Read the CSV file to create the folders and its structure.
'* Get the total number of files to be converted.
'*----------------------------------------------------------------------------------------------------------------------
Set objInputFile = objFSO.OpenTextFile(strCSVFullFileName, ForReading)

lngTotalFileCount = 0
Do Until objInputFile.AtEndOfStream
    strLine = objInputFile.ReadLine
    If InStr(strLine, "|") > 0 Then
        'Check if the Object Type is a Folder.
        strText = Split(strLine, "|")(1)
        strObjType = Trim(Split(strText, ":")(1))

        If strObjType = "Folder" Then

            'Folder Name
            strText = Split(strLine, "|")(3)
            strFolderName = Trim(Split(strText, ":")(1))
            
            'Parent Folder
            strText = Split(strLine, "|")(10)
            strParentFolderName = Trim(Split(strText, ":")(1))

            'Create the folders
            CreateFolders strFolderName, strParentFolderName
        End If
    End If
    'type_: 1 is a File.
    If InStr(strLine, "|") > 0 And InStr(strLine, "type_: 1") > 0 Then
        'Check if the Object Type is a File.
        strText = Split(strLine, "|")(1)
        strObjType = Trim(Split(strText, ":")(1))

        'Check if the "type_" is 1.
        strText = Split(strLine, "|")(31)
        strType = Trim(Split(strText, ":")(1))

        If strObjType = "File" And strType = "1" Then
            lngTotalFileCount = lngTotalFileCount + 1
        End If
    End If
Loop

objInputFile.Close

'*----------------------------------------------------------------------------------------------------------------------
'* Read the CSV file to convert the files.
'*----------------------------------------------------------------------------------------------------------------------
Set objInputFile = objFSO.OpenTextFile(strCSVFullFileName, ForReading)
lngCounter = 0
lngMaxJobs = 10
lngJobCounter = 0
Set objLogsFolder = objFSO.GetFolder(strTempLogsFolderPath)

Do Until objInputFile.AtEndOfStream
    strLine = objInputFile.ReadLine
    'type_: 1 is a File.
    If InStr(strLine, "|") > 0 And InStr(strLine, "type_: 1") > 0 Then
        'Check if the Object Type is a File.
        strText = Split(strLine, "|")(1)
        strObjType = Trim(Split(strText, ":")(1))

        'Check if the "type_" is 1.
        strText = Split(strLine, "|")(31)
        strType = Trim(Split(strText, ":")(1))

        If strObjType = "File" And strType = "1" Then
            lngCounter = lngCounter + 1

            'Markdown File Name
            strText = Split(strLine, "|")(0)
            strMarkdownName = Trim(Replace(Split(strText, ":")(1), """", ""))
            strMarkdownNameFullName = Chr(34) & strTempCleanedUpFolderPath & "\" & strMarkdownName & Chr(34)

            'Folder Name
            strText = Split(strLine, "|")(4)
            strFolderName = Trim(Split(strText, ":")(1))
            
            'Find the full path.
            Set objDestination = SearchFolder(objFSO.GetFolder(strOutputFolderPath), strFolderName)
            strDestinationPath = objDestination.Path

            'Output File Name
            strText = Split(strLine, "|")(2)
            strOutputName = Trim(Replace(Split(strText, ":")(1), """", ""))
            strOutputFullName = Chr(34) & strDestinationPath & "\" & strOutputName & Chr(34)

            'Log File
            strLogFile = strTempLogsFolderPath & "\" & lngCounter & " - " & strMarkdownName & ".txt"

            blnExitLoop = False
            Do While Not blnExitLoop
                'Get the number of files inside the Logs Folder.
                Set objFiles = objLogsFolder.Files
                lngNumberOfLogs = objFiles.Count

                If lngJobCounter < lngMaxJobs Then
                    blnExitLoop = True
                End If

                If (lngNumberOfLogs > 0 And (lngNumberOfLogs Mod lngMaxJobs) = 0) Or _
                    (lngCounter > (Round(lngTotalFileCount / lngMaxJobs) * lngMaxJobs) And lngCounter <= lngTotalFileCount) Then
                    lngJobCounter = 0
                    blnExitLoop = True
                    'Wait for 10 seconds to avoid too much eating of resources.
                    WScript.Sleep 10000
                End If
            Loop

            lngJobCounter = lngJobCounter + 1
            strCommand = conBatchFile & " " & strMarkdownNameFullName & " " & strOutputFullName & " " & Chr(34) & strResourcesParentDirectory & Chr(34) & _
                            " " & Chr(34) & strLogFile & Chr(34)
            If (lngCounter > (Round(lngTotalFileCount / lngMaxJobs) * lngMaxJobs) And lngCounter <= lngTotalFileCount) Then
                objWScriptShell.Run strCommand, 0, True
            Else
                objWScriptShell.Run strCommand, 0, False
            End If
        End If
    End If
Loop

objInputFile.Close

'Get the number of files inside the Logs Folder.
Set objFiles = objLogsFolder.Files
lngNumberOfLogs = objFiles.Count

If lngCounter > 0 Then
    'Wait for a few more seconds to ensure everything is completed.
    WScript.Sleep 5000
    strCommand = ""
End If

'*----------------------------------------------------------------------------------------------------------------------
'* Rename the Folders.
'*----------------------------------------------------------------------------------------------------------------------
Dim objList
Set objList = CreateObject("System.Collections.ArrayList")
strOutputFolderPath = Trim(strOutputDirectory) & "\" & Trim(strOutputFolderName)
Set objOutputFolder = objFSO.GetFolder(strOutputFolderPath)

strTempSearchFile = strTempLogsFolderPath & "\" & "FolderSearch_" & strFormattedDateTime & ".txt"

'Collect all folders starting from the main folder
GetAllFolders objOutputFolder, objList
'Add the main folder last
objList.Add objOutputFolder

For Each objSubFolder In objList
    'Find the New Name of the Folder.
    strCommand = "cmd /c findstr /i /c:""\""" & objSubFolder.Name & ".md\""" & "|" & "Object Type: Folder"" " & Chr(34) & strCSVFullFileName & Chr(34) & _
                    " > " & Chr(34) & strTempSearchFile & Chr(34)
    objWScriptShell.Run strCommand, 0, True

    strLine = ""
    If objFSO.FileExists(strTempSearchFile) Then
        Set objFile = objFSO.OpenTextFile(strTempSearchFile, 1)
        Do Until objFile.AtEndOfStream
            strLine = strLine & objFile.ReadLine & vbCrLf
        Loop
        objFile.Close
        objFSO.DeleteFile strTempSearchFile 'Clean up the temporary file
    End If

    'Rename the Folder.
    If strLine <> "" Then
        strText = Split(strLine, "|")(2)
        strFolderName = Trim(Split(strText, ":")(1))
        strParentFolderName = objFSO.GetParentFolderName(objSubFolder.Path)
        strNewFolderPath = strParentFolderName & "\" & strFolderName
        lngCounter = 0
        blnFolderRenamed = False
        Do Until blnFolderRenamed
            If objFSO.FolderExists(strNewFolderPath) Then
                lngCounter = lngCounter + 1
                strNewFolderPath = strParentFolderName & "\" & strFolderName & " (" & lngCounter & ")"
            Else
                lngPos = InStrRev(strNewFolderPath, "\") + 1
                strNewFolderName = Mid(strNewFolderPath, lngPos, Len(strNewFolderPath)-lngPos+1)
                objSubFolder.Name = strNewFolderName
                blnFolderRenamed = True
            End If
        Loop
    End If
Next

'*----------------------------------------------------------------------------------------------------------------------
'* Consolidate the logs.
'*----------------------------------------------------------------------------------------------------------------------
Set objInputFolder = objFSO.GetFolder(strTempLogsFolderPath)

'Loop through each file in the folder
lngCounter = 0
For Each objInputFile In objInputFolder.Files
    
    If objInputFile.Size > 0 Then
        
        lngCounter = lngCounter + 1
        If lngCounter = 1 Then
            'Create the Conversion Log File
            strConversionLogFileName = "Conversion_Logs_" & strFormattedDateTime & ".txt"
            strConversionLogFullFileName = strOutputFolderPath & "\" & strConversionLogFileName
            'Create or open the TXT file for writing
            If objFSO.FileExists(strConversionLogFullFileName) Then
                objFSO.DeleteFile(strConversionLogFullFileName)
            End If
            'Write the Title
            strTXTLine = "Conversion Logs (Warnings, Errors)"
            strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
            objWScriptShell.Run strCommand, 0, True
            'Write new empty line
            strCommand = "cmd /c echo. >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
            objWScriptShell.Run strCommand, 0, True
            'Write the Note
            strTXTLine = "Note: Original Markdown files are found in " & strTempInputFolderPath
            strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
            objWScriptShell.Run strCommand, 0, True
            'Write new empty line
            strCommand = "cmd /c echo. >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
            objWScriptShell.Run strCommand, 0, True
        End If

        'Insert a divider
        strTXTLine = "--------------------------------------------------------------------------------------------------------------"
        strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
        objWScriptShell.Run strCommand, 0, True

        'Write the File Name of the Markdown File.
        strFileName = Trim(Split(Replace(objInputFile.Name,".txt",""), "-")(1))
        strTXTLine = "Input File: " & strFileName
        strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
        objWScriptShell.Run strCommand, 0, True

        'Write the File Location of the Markdown File.
        strTXTLine = "File Path: " & strTempInputFolderPath & "\" & strFileName
        strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
        objWScriptShell.Run strCommand, 0, True

        'Write the label "Log Contents:" before writing the contents of the log file.
        strTXTLine = "Log Contents:"
        strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
        objWScriptShell.Run strCommand, 0, True

        'Write new empty line
        strCommand = "cmd /c echo. >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
        objWScriptShell.Run strCommand, 0, True

        'Write the contents of the Log File.
        strCommand = "cmd /c type """ & objInputFile.Path & """ >> " & Chr(34) & strConversionLogFullFileName & Chr(34)
        objWScriptShell.Run strCommand, 0, True

    End If

    'Delete the Individual Log File.
    objFSO.DeleteFile(objInputFile.Path)

Next

'*----------------------------------------------------------------------------------------------------------------------
'* Clean up.
'*----------------------------------------------------------------------------------------------------------------------
'Delete the folder and its contents: JTPC*
strCommand = "cmd /c rmdir /s /q """ & strTempCleanedUpFolderPath & """"
objWScriptShell.Run strCommand, 0, True

'Delete the folder and its contents: JTPL*
strCommand = "cmd /c rmdir /s /q """ & strTempLogsFolderPath & """"
objWScriptShell.Run strCommand, 0, True

'*----------------------------------------------------------------------------------------------------------------------
'* Completion.
'*----------------------------------------------------------------------------------------------------------------------
'Unhide the Input and Result Folders, i.e. JTPI and JTP_ respectively.
objTempInputFolder.Attributes = objTempInputFolder.Attributes And Not 2
objOutputFolder.Attributes = objOutputFolder.Attributes And Not 2

'Calculate the total duration in seconds
objTimeStampEnd = Now()
lngDuration = DateDiff("s", objTimeStampStart, objTimeStampEnd)
'Convert total duration into minutes and seconds
lngMinutes = Int(lngDuration / 60)
lngSeconds = lngDuration Mod 60

strMsg = ""
If lngMinutes > 0 Then
    strMsg = Trim(strMsg) & " " & lngMinutes & " minute"
    If lngMinutes > 1 Then
        strMsg = Trim(strMsg) & "s"
    End If
    If lngSeconds > 0 Then
        strMsg = Trim(strMsg) & " and "
    End If
End If
If lngSeconds > 0 Then
    strMsg = Trim(strMsg) & " " & lngSeconds & " second"
    If lngSeconds > 1 Then
        strMsg = Trim(strMsg) & "s"
    End If
End If
strMsg = Trim(strMsg) & "."

'Generate the folder structure using windows batch command tree /f.
'Create the Summary File
strSummaryFileName = "Summary_" & strFormattedDateTime & ".txt"
strSummaryFullFileName = strOutputFolderPath & "\" & strSummaryFileName
'Create or open the TXT file for writing
If objFSO.FileExists(strSummaryFullFileName) Then
    objFSO.DeleteFile(strSummaryFullFileName)
End If
'Write the Title
strTXTLine = conMsgBoxTitle
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Insert a divider
strTXTLine = "----------------------------------------------------------------------------------------------------------------------------------"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Write the Completion Summary
strTXTLine = "Process completed in " & strMsg
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Insert a divider
strTXTLine = "----------------------------------------------------------------------------------------------------------------------------------"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Write note on conversion log
strTXTLine = "Please check the conversion log file " & Chr(34) & strConversionLogFileName & Chr(34) & " for errors or warnings."
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Write note on "List" csv file
strTXTLine = "File " & Chr(34) & strCSVFileName & Chr(34) & " contains the list of Markdown files with its attribute and can be used for reference and checking."
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Insert a divider
strTXTLine = "----------------------------------------------------------------------------------------------------------------------------------"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Write the Directories
'Directory of
strTXTLine = "Directory of"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Markdown Files
strTXTLine = "    Markdown Files: " & strTempInputFolderPath
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Output Files
strTXTLine = "    Output Files: " & strOutputFolderPath
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Insert a divider
strTXTLine = "----------------------------------------------------------------------------------------------------------------------------------"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Write the label "List of files converted"
strTXTLine = "List of Files converted (folder structure)"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Write new empty line
strCommand = "cmd /c echo. >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Insert the structure
strCommand = "cmd /c tree /f /a """ & strOutputFolderPath & """ >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True
'Insert a divider
strTXTLine = "----------------------------------------------------------------------------------------------------------------------------------"
strCommand = "cmd /c echo " & strTXTLine & " >> " & Chr(34) & strSummaryFullFileName & Chr(34)
objWScriptShell.Run strCommand, 0, True

'Display Completion Message.
strMsg = "Conversion Completed." & vbCrLf & vbCrLf & _
            "Process completed in " & Trim(strMsg)

strMsg = Trim(strMsg) & vbCrLf & vbCrLf & _
            "Converted files are found in " & Chr(34) & strOutputFolderPath & Chr(34) & vbCrLf & vbCrLf & _
            "The following files have been generated in the abovementioned folder for reference and checking if needed:" & vbCrLf & _
            "    " & strSummaryFileName & vbCrLf & _
            "    " & strConversionLogFileName & vbCrLf & _
            "    " & strCSVFileName & vbCrLf & vbCrLf & _
            "Actual Markdown files extracted from " & Chr(34) & strInputFile & Chr(34) & " and its resources are found in " & vbCrLf & _
            Chr(34) & strTempInputFolderPath & Chr(34)

MsgBox strMsg, vbOKOnly + vbInformation, strTitle

'Open the Converted Files Folder
strCommand = "explorer.exe /n, " & strOutputFolderPath
objWScriptShell.Run strCommand

WScript.Quit


'***********************************************************************************************************************
'* Functions and Subroutines 
'***********************************************************************************************************************

'Function to replace specified characters with space
Function ReplaceSpecialChars(str)

    Dim charsToReplace, char

    charsToReplace = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "•", "	", "")
    For Each char In charsToReplace
        str = Replace(str, char, " ")
    Next

	str = Replace(str, vbCrLf, " ")
	str = Replace(str, vbCr, " ")
    ReplaceSpecialChars = str

End Function

'Function to search for file and return the name with extension.
Function GetNameWithExtension(strDirectory, strSearchString)

    Dim objResourceFolder

    Set objResourceFolder = objFSO.GetFolder(strDirectory)

    GetNameWithExtension = strSearchString
    
    For Each objResourceFile In objResourceFolder.Files
        If InStr(1, objResourceFile.Name, strSearchString, vbTextCompare) > 0 Then
            GetNameWithExtension = objResourceFile.Name
            Exit For
        End If
    Next

End Function

'Function to Create Folders
Sub CreateFolders(strChild, strParent)

    Dim strParentFolderPath, strChildFolderPath
    Dim objParentFolder

    If Trim(strParent) = "" Then
        strParentFolderPath = strOutputFolderPath
    Else
        strParentFolderPath = strOutputFolderPath & "\" &  strParent
        Set objParentFolder = SearchFolder(objFSO.GetFolder(strOutputFolderPath), strParent)
        If objParentFolder Is Nothing Then 
            'Parent folder not found, create it in the root directory
            objFSO.CreateFolder strParentFolderPath
        Else
            'Parent folder found, get its path
            strParentFolderPath = objParentFolder.Path
        End If
    End If

    strChildFolderPath = strParentFolderPath & "\" & strChild
    If Not objFSO.FolderExists(strChildFolderPath) Then
        objFSO.CreateFolder strChildFolderPath
    End If

    strChildFolderPath = strOutputFolderPath & "\" & strChild
    If objFSO.FolderExists(strChildFolderPath) And _
        strParentFolderPath <> strOutputFolderPath Then
        'Use windows batch command robocopy to move the folder.
        strChildFolderPath = Chr(34) & Trim(strChildFolderPath) & Chr(34)
        strParentFolderPath = Chr(34) & Trim(strParentFolderPath) & "\" & strChild & Chr(34)
        strCommand = "cmd /c robocopy " & strChildFolderPath & " " & strParentFolderPath & " /E /MOVE"
        objWScriptShell.Run strCommand, 0, True
    End If

End Sub

'Recursive function to search for the parent folder
Function SearchFolder(rootFolder, searchFolderName)

    Dim objSubFolder, objFoundFolder

    Set objFoundFolder = Nothing

    For Each objSubFolder In rootFolder.SubFolders
        If objSubFolder.Name = searchFolderName Then
            Set SearchFolder = objSubFolder
            Exit Function
        Else
            Set objFoundFolder = SearchFolder(objSubFolder, searchFolderName)
            If Not objFoundFolder Is Nothing Then
                Set SearchFolder = objFoundFolder
                Exit Function
            End If
        End If
    Next

    Set SearchFolder = Nothing

End Function

'Function to recursively collect all folders
Function GetAllFolders(parentFolder, objList)

    Dim subFolders, subFolder

    Set subFolders = parentFolder.SubFolders

    For Each subFolder In subFolders
        GetAllFolders subFolder, objList 'Recursion
        objList.Add subFolder 'Add to list after processing subfolders
    Next

End Function