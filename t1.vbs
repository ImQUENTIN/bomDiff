' we can pass the excel file name and worksheet as the parameter and using Wscript.Arguments to use it in the script
 
if(Wscript.Arguments.Count<3) then
 msgbox "The count of Arguments is "&Wscript.Arguments.Count
 'Quit VBS script
 'Wscript.Quit
End if

Set oExcel=CreateObject("excel.application")

'@@@ path
Set oWorkBook=oExcel.Workbooks.Open("d:/bomDiff/temp.xlsx")

'!!! use the worksheet "Prop_Methods"
Set oSheet=oWorkBook.Sheets("Sheet1")

'Get the used range
'Set Sheet = oExcel.Worksheets("Prop_Methods").UsedRange
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("c:\testfile.txt", True)
a.Close
' Read=1 Write=2 Append =8
Set a = fs.OpenTextFile("c:\testfile.txt", 8,false)
‘we can write more value to the text file by using loop
value=oSheet.cells(4,2)
a.WriteLine(value)
value=osheet.cells(5,2)
a.WriteLine(value)
value=osheet.cells(6,2)
a.WriteLine(value)
a.Close
Set oSheet=Nothing 
oExcel.Quit
