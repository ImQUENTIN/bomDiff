''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   BOM 差异表生成脚本                  ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
fileName = "qyx"

'  1. 拖拽需要对比的两个文件放置到脚本上                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oldBomPath, newBomPath, testBom
set arg = wScript.arguments
set oExcel = createObject("excel.application")
currentPath = createObject("scripting.filesystemObject").getfile(wscript.scriptfullName).parentfolder.path
diffBomPath = currentPath & "\" & fileName & ".xls"

' 1.1 操作说明
if arg.count = 0 then
	' 没有拖拽文件： 
	
	' a) 默认同时选中两个BOM表
	testBom = oExcel.getOpenFileName("BOM, *.xls, BOM, *.xlsx", 1, "选中需要对比的BOM表【Tips: 如果新旧表在同一目录下，可以同时选中两个文件；否则，建议先选择旧文件】", ,true)
	if not isArray(testBom) then
		msgbox "没有选择文件，脚本结束运行。", vbOKOnly, "→_→  中招了吧 -_-|||"
		wScript.quit
	else
		' 数组上下限， 下限为1， 上限为选择的文件数
		'msgbox "上限：" & uBound(testBom) & vbCrLf & "下限：" & lBound(testBom)
		if(uBound(testBom) > 2) then
			' 用户选择的文件多于两个，非设计。
			msgbox "最多只能选中2个文件，脚本结束运行。", vbOKOnly, "→_→  中招了吧 -_-|||"
			wScript.quit
		elseif uBound(testBom) = 1 then
			' 用户仅选择1个文件时，需要再选择另外一个文件。
			msgbox "由于您只选中了1个文件，暂将其视为旧BOM，现在请再选择新BOM。", vbOKOnly, "qyx's tips"
			oldBomPath = testBom
			newBomPath = oExcel.getOpenFileName("Another BOM, *.xls, Another BOM, *.xlsx", 1, "请选择另一张BOM表")
	    else
			' 用户同时选择了两个文件。
			oldBomPath = testBom(1)
			newBomPath = testBom(2)
		end if
	end if
	
	' b) 一次选中一个BOM表
	' msgbox "请先选择【旧】的BOM表",0,"qyx's tips"  
	' oldBomPath = oExcel.GetOpenFilename("旧BOM表(*.xls), *.xls, 旧BOM表(*.xlsx), *.xlsx")	 
	
	' msgbox "请再选择【新】的BOM表",0,"qyx's tips"  
	' newBomPath = oExcel.GetOpenFilename("新BOM表(*.xls), *.xls, 新BOM表(*.xlsx), *.xlsx")	  

	' msgbox "旧表：" & oldBomPath & vbcrlf & "新表：" & newBomPath, 0, "qyx's tips"
	
elseif arg.count > 2 or arg.count = 1 then
	' 拖拽两个以上的文件
	msgbox "最多且只能拖拽2个文件，脚本结束运行。", vbOKOnly, "→_→  中招了吧 -_-|||"
	wScript.quit
else
	newBomPath = arg(0)
	oldBomPath = arg(1)
end if 

' 1.2 已成功选择了两个文件
ans = msgbox("请确认" & vbCrLf & vbCrLf &_
	  "新文件为："& newBomPath & vbCrLf & "旧文件为："& oldBomPath, vbYesNo, "请确认新旧文件顺序是否正确！")
	  
if ans = vbNo then
    tempBom = newBomPath
	newBomPath  = oldBomPath
	oldBomPath  = tempBom
	msgbox "请确认" & vbCrLf & vbCrLf &_
	  "新文件为："& newBomPath & vbCrLf & "旧文件为："& oldBomPath, vbYes, "请确认新旧文件顺序是否正确！"
end if	

'  2. 打开excel表格                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
with oExcel

.visible = true				' 调试使用
.caption = "qyx's vbs"
.workBooks.add				' 新建一个工作簿
.DisplayAlerts = false		' 不显示警告

' 2.1 保留一个表做为差异表
do until .worksheets.count = 4
.worksheets.add
loop
.worksheets(1).name = "diff"
.worksheets(2).name = "new"
.worksheets(3).name = "del"
.worksheets(4).name = "chg"

' 2.2 保存该工作表
if not .activeworkbook.saved then
.activeWorkBook.saveAs diffBomPath, 56
end if 
.activeWorkBook.close

' 3. 比较计算
' 3.1 
set bookDiff = .workBooks.open(diffBomPath)
set bookNew  = .workBooks.open(newBomPath)
set bookOld  = .workBooks.open(oldBomPath)

'不显示提示信息,这样保存的时候就不会提示是否要覆盖原文件 


bookDiff.worksheets("diff").cells(1,1).value = "hello"

' 保存及退出
bookDiff.save
bookOld.close
bookNew.close
bookDiff.close
.quit

end with


