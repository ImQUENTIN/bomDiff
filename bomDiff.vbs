''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   BOM ��������ɽű�                  ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
fileName = "qyx"

'  1. ��ק��Ҫ�Աȵ������ļ����õ��ű���                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
dim oldBomPath, newBomPath, testBom
set arg = wScript.arguments
set oExcel = createObject("excel.application")
currentPath = createObject("scripting.filesystemObject").getfile(wscript.scriptfullName).parentfolder.path
diffBomPath = currentPath & "\" & fileName & ".xls"

' 1.1 ����˵��
if arg.count = 0 then
	' û����ק�ļ��� 
	
	' a) Ĭ��ͬʱѡ������BOM��
	testBom = oExcel.getOpenFileName("BOM, *.xls, BOM, *.xlsx", 1, "ѡ����Ҫ�Աȵ�BOM��Tips: ����¾ɱ���ͬһĿ¼�£�����ͬʱѡ�������ļ������򣬽�����ѡ����ļ���", ,true)
	if not isArray(testBom) then
		msgbox "û��ѡ���ļ����ű��������С�", vbOKOnly, "��_��  �����˰� -_-|||"
		wScript.quit
	else
		' ���������ޣ� ����Ϊ1�� ����Ϊѡ����ļ���
		'msgbox "���ޣ�" & uBound(testBom) & vbCrLf & "���ޣ�" & lBound(testBom)
		if(uBound(testBom) > 2) then
			' �û�ѡ����ļ���������������ơ�
			msgbox "���ֻ��ѡ��2���ļ����ű��������С�", vbOKOnly, "��_��  �����˰� -_-|||"
			wScript.quit
		elseif uBound(testBom) = 1 then
			' �û���ѡ��1���ļ�ʱ����Ҫ��ѡ������һ���ļ���
			msgbox "������ֻѡ����1���ļ����ݽ�����Ϊ��BOM����������ѡ����BOM��", vbOKOnly, "qyx's tips"
			oldBomPath = testBom
			newBomPath = oExcel.getOpenFileName("Another BOM, *.xls, Another BOM, *.xlsx", 1, "��ѡ����һ��BOM��")
	    else
			' �û�ͬʱѡ���������ļ���
			oldBomPath = testBom(1)
			newBomPath = testBom(2)
		end if
	end if
	
	' b) һ��ѡ��һ��BOM��
	' msgbox "����ѡ�񡾾ɡ���BOM��",0,"qyx's tips"  
	' oldBomPath = oExcel.GetOpenFilename("��BOM��(*.xls), *.xls, ��BOM��(*.xlsx), *.xlsx")	 
	
	' msgbox "����ѡ���¡���BOM��",0,"qyx's tips"  
	' newBomPath = oExcel.GetOpenFilename("��BOM��(*.xls), *.xls, ��BOM��(*.xlsx), *.xlsx")	  

	' msgbox "�ɱ�" & oldBomPath & vbcrlf & "�±�" & newBomPath, 0, "qyx's tips"
	
elseif arg.count > 2 or arg.count = 1 then
	' ��ק�������ϵ��ļ�
	msgbox "�����ֻ����ק2���ļ����ű��������С�", vbOKOnly, "��_��  �����˰� -_-|||"
	wScript.quit
else
	newBomPath = arg(0)
	oldBomPath = arg(1)
end if 

' 1.2 �ѳɹ�ѡ���������ļ�
ans = msgbox("��ȷ��" & vbCrLf & vbCrLf &_
	  "���ļ�Ϊ��"& newBomPath & vbCrLf & "���ļ�Ϊ��"& oldBomPath, vbYesNo, "��ȷ���¾��ļ�˳���Ƿ���ȷ��")
	  
if ans = vbNo then
    tempBom = newBomPath
	newBomPath  = oldBomPath
	oldBomPath  = tempBom
	msgbox "��ȷ��" & vbCrLf & vbCrLf &_
	  "���ļ�Ϊ��"& newBomPath & vbCrLf & "���ļ�Ϊ��"& oldBomPath, vbYes, "��ȷ���¾��ļ�˳���Ƿ���ȷ��"
end if	

'  2. ��excel���                          
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
with oExcel

.visible = true				' ����ʹ��
.caption = "qyx's vbs"
.workBooks.add				' �½�һ��������
.DisplayAlerts = false		' ����ʾ����

' 2.1 ����һ������Ϊ�����
do until .worksheets.count = 4
.worksheets.add
loop
.worksheets(1).name = "diff"
.worksheets(2).name = "new"
.worksheets(3).name = "del"
.worksheets(4).name = "chg"

' 2.2 ����ù�����
if not .activeworkbook.saved then
.activeWorkBook.saveAs diffBomPath, 56
end if 
.activeWorkBook.close

' 3. �Ƚϼ���
' 3.1 
set bookDiff = .workBooks.open(diffBomPath)
set bookNew  = .workBooks.open(newBomPath)
set bookOld  = .workBooks.open(oldBomPath)

'����ʾ��ʾ��Ϣ,���������ʱ��Ͳ�����ʾ�Ƿ�Ҫ����ԭ�ļ� 


bookDiff.worksheets("diff").cells(1,1).value = "hello"

' ���漰�˳�
bookDiff.save
bookOld.close
bookNew.close
bookDiff.close
.quit

end with


