''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   BOM �����                          ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
fileName = "bomDiff"

'  1. ��ק��Ҫ�Աȵ������ļ����õ��ű���                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oldBom, newBom
set arg = wScript.arguments

' 1.1 ����˵��
if arg.count > 2 then
	' ��ק�������ϵ��ļ�
	msgbox "ֻ����Ƚ�����BOM��"& vbCrLf & "���������У�",0,"qyx's tips"
	wScript.quit
elseif arg.count = 0 then
	' û����ק�ļ�
	msgbox "ʹ�÷�����ֱ����ק����BOM�����˽ű��ļ���"_
	       & vbCrLf & "���������У�",0,"qyx's tips"
	wScript.quit
end if
' xFilename = oExcel.GetOpenFilename("Excel Files (*.xls), *.xls")


' 1.2 �ѳɹ���ק�����ļ�
ans = msgbox("��ȷ��" & vbCrLf & vbCrLf &_
	  "���ļ�Ϊ��"& arg(0) & vbCrLf & "���ļ�Ϊ��"& arg(1), 4, "qyx's tips")
	  
if ans = vbYes then
	newBom = arg(0)
	oldBom = arg(1)
else
	newBom = arg(1)
	oldBom = arg(0)
end if

'  2. ��excel���                          
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

'����ʾ��ʾ��Ϣ,���������ʱ��Ͳ�����ʾ�Ƿ�Ҫ����ԭ�ļ� 
oExcel.DisplayAlerts=FALSE 

' bookDiff.save
' bookOld.save
' bookNew.save
' bookDiff.saved = true

' bookDiff.close
' bookOld.close
' bookNew.close
oExcel.quit



