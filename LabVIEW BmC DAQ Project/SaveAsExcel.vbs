Option Explicit											' for debugging

' initialize variables
Dim objExcel, objFileToRead, strFileFullPath, strFileName, strNewName, objShell, objBook, xlClosed

' Open txt file that specifies what tdms file to open
' create file ssytem object
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("fileSaveInfo.txt",1)
strFileFullPath = objFileToRead.ReadLine				' read full file path from txt
strFileName = objFileToRead.ReadLine					' read file name from txt
strNewName = objFileToRead.ReadLine						' read new name to save file as
objFileToRead.Close										' close txt file
Set objFileToRead = Nothing								' release objFileToRead


On Error Resume Next									' continue if error is encountered
WScript.Timeout = 20									' kill script after 10 seconds in case stuck in loop
xlClosed = 0											' xlClosed: whether or not other excel instances have been closed

' loop through all open instances of excel and close 
Do While xlClosed = 0
	Set objExcel = GetObject(,"Excel.Application") 		' object invoked is disconnected. Get ref to open excel instance
	
	' output any errors encountered
	If Err Then
		' uncomment line below to see errors
		'MsgBox "Notice: Error " & Err.Number & " " & Err.Description
		xlClosed = 1		
		Err.clear
	End If
	
	' exit loop if all excel applications have been closed
	If objExcel is Nothing Then
		xlClosed = 1
	Else
		' Go through open workbooks and close each, does not save changes
		For Each objBook In objExcel.Workbooks
			objBook.Close False
		Next 
		' close excel application instance
		objExcel.Quit
	End If
	' set objExcel to Nothing again and loop over 
	Set objExcel = Nothing
Loop


' Launch tdms file
Set objShell = WScript.CreateObject("WScript.Shell")		' open command prompt
objShell.run DblQuote(strFileFullPath)						' run selected tdms file
Wscript.Sleep(10000)											' pause script - MUST KEEP! Excel needs a second to load or GetObject returns empty
' Wscript.Sleep(5000)											' pause script - MUST KEEP! Excel needs a second to load or GetObject returns empty


' Save tdms file
Set objExcel = GetObject(,"Excel.Application") 									' Get ref to open excel instance
objExcel.DisplayAlerts = False													' Disable excel prompts (file overwrite prompt)

' find workbook with same file name as specified in txt file
If  objExcel.Workbooks(1).Sheets(1).Name = strFileName & " (root)" Then
	objExcel.Workbooks(1).Activate												' set this workbook as the active workbook
	'Err.Clear
	objExcel.Workbooks(1).SaveAs strNewName										' save file as xlsx
	'If Err.Number <> 0 Then
	'	Wscript.Echo Err.Description
	'End If
	objExcel.Workbooks(1).Close False											' close workbook

	objExcel.Workbooks(1).Saved = True
	objExcel.Workbooks(1).Close False											' close workbook

End If

' release objects and quit excel application
Set objBook = Nothing
Set objShell = Nothing  
objExcel.Quit
Set ObjExcel = Nothing

' function that places an input string into double quotes
Function DblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function