''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   BOM 差异表                          ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
fileName = "bomDiff"

'  1. 拖拽需要对比的两个文件放置到脚本上                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oldBom, newBom
set arg = wScript.arguments

' 1.1 操作说明
if arg.count > 2 then
	' 拖拽两个以上的文件
	msgbox "只允许比较两个BOM表"& vbCrLf & "请重新运行！",0,"qyx's tips"
	wScript.quit
elseif arg.count = 0 then
	' 没有拖拽文件
	msgbox "使用方法：直接拖拽两个BOM表至此脚本文件上"_
	       & vbCrLf & "请重新运行！",0,"qyx's tips"
	wScript.quit
end if
' xFilename = oExcel.GetOpenFilename("Excel Files (*.xls), *.xls")


' 1.2 已成功拖拽两个文件
ans = msgbox("请确认" & vbCrLf & vbCrLf &_
	  "新文件为："& arg(0) & vbCrLf & "旧文件为："& arg(1), 4, "qyx's tips")
	  
if ans = vbYes then
	newBom = arg(0)
	oldBom = arg(1)
else
	newBom = arg(1)
	oldBom = arg(0)
end if

'  2. 打开excel表格                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oExcel, bookNew, bookOld, bookDiff

currentPath = createObject("scripting.filesystemObject").getfile(wscript.scriptfullName).parentfolder.path

set oExcel = createObject("excel.application")
oExcel.visible = true
' oExcel.caption = "qyx's vbs"
set bookDiff = oExcel.workBooks.add
oExcel.save currentPath & "\" & fileName & ".xls"

msgbox "pause"
' set bookDiff = 
' set bookNew = oExcel.workBooks.open(newBom)
' set bookOld = oExcel.workBooks.open(oldBom)

'不显示提示信息,这样保存的时候就不会提示是否要覆盖原文件 
oExcel.DisplayAlerts=FALSE 

' bookDiff.save
' bookOld.save
' bookNew.save
' bookDiff.saved = true

' bookDiff.close
' bookOld.close
' bookNew.close
oExcel.quit



