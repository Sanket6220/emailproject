Option Explicit
Dim xlApp,xlBook

Set xlApp = CreateObject("Excel.Application")

Set xlBook = xlApp.Workbooks.Open( GetCurrentFolder() & "\FinalSheet.xlsm")
xlApp.Run "Main"
xlBook.Close
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing

WScript.Echo "Processed Successfully."
WScript.Quit

Function getCurrentFolder()
	Dim FSO
	Set fso = CreateObject("Scripting.FileSystemObject")
	GetCurrentFolder = FSO.GetAbsolutePathName(".")
End Function