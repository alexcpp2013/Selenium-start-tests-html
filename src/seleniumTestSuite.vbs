Function fileRead(ByVal fileAllName, ByRef content)

	Const ForReading = 1
	Const TristateTrue = -1
	Set FSO = CreateObject("Scripting.FileSystemObject") 
	If(FSO.FileExists(fileAllName)) Then
		Set TS = FSO.GetFile(fileAllName)
		If(TS.Size > 0) Then
			'TristateUseDefault
			Set TS = FSO.OpenTextFile(fileAllName, ForReading, False, TristateUseDefault)
			content = TS.ReadAll
			MsgBox content
			TS.Close
			Set FSO = Nothing   
		Else
			MsgBox "FR: Nothing to read. 2" 
			fileRead = 2
			Set FSO = Nothing 
			Exit Function 
		End If
		 
	Else
		MsgBox "FR: No input file. 1"
		fileRead = 1
		Exit Function
	End If
	
	fileRead = 0

End Function

Function textFind(ByRef content, ByVal findText)

	If(content <> "") Then
		If(findText <> "") Then
			textFind = InStr(content, findText) 
		Else 
			MsgBox "TF:  Nothing to Find. 2"
			textFind = 2
			Exit Function
		End If
	Else  
		MsgBox "TF:  Nothing in Find. 1"
		textFind = 1
		Exit Function
	End If

End Function

Function fileManage()

	Dim strIn
	Dim strSelenium
	Dim strTempFile
	Dim strTestSuiteFile
	Dim strCommandWindows
	Dim strCommandSelenium
	Dim strBrowser
	Dim strPort
	Dim strHost
	strHost = ""
	strPort = "-port 4444"
	strBrowser = "*firefox"
	strSelenium = "C:\selenium-server-standalone-2.33.0.jar"
	strCommandWindows = "cmd /K java -jar"
	strCommandSelenium = "-htmlSuite"
	strtestSuiteFile = "C:\HTML\h1.html"
	strTempFile = "C:\HTML\testResults.txt"
	strIn = strCommandWindows & " " &  strSelenium &  " " & strPort & " " &  strCommandSelenium & " " & strBrowser & " " & strHost & " " & strTestSuiteFile & " " &  strTempFile
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")
	intRes = WshShell.Run(strIn, 1, TRUE)
	
	fileName = strTempFile
	findText = "Error"
	Dim content
	Dim result
	result = fileRead(fileAllName, content)
	If (result = 0) Then
		result = textFind(content, findText)	
	End If	
	
	If(result = 0) Then
		MsgBox "FM: Good! 0"
	Else 
		MsgBox "There are Error! 1"
		Dim WshShellFile
		set WshShellFile = WScript.CreateObject("WScript.Shell")
		WshShellFile.Run("cmd /K start" & " " & fileName)
	End If
	content = ""

End Function


'----Manage----------------



Call fileManage()
