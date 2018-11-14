''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   arguments                           ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'  1. property                           
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1.1 Arguments
set arg = wScript.arguments
for i = 0 to arg.count - 1
  msgbox arg(i), 0, "Arguments"
next

' 1.2 Fullname ���ԣ���ʾWScript.exe����ȫ�޶�·��
msgbox wScript.fullname,0,"Full Name"

' 1.3 Interactive ���ԣ�����ʽ��������ģʽ
msgbox wScript.interactive, 0, "Interactive"

' 1.4 Name ���ԣ����� WScript ����������ִ���ļ���������
msgbox wScript.name, 0 ,"Name"

' 1.5 Path ���ԣ����ذ���������ִ���ļ���CScript.exe �� WScript.exe����·������
msgbox wScript.path, 0, "Path"

' 1.6 ScriptFullName ���ԣ����ر��ű��ļ�����ȫ�޶�·��
msgbox wScript.scriptFullName, 0, "ScriptFullName"

' 1.7 ScriptName ���ԣ����ر��ű��ļ�������
msgbox wScript.scriptName, 0, "ScriptName"

' 1.8 Version ���ԣ�����Windows�ű������İ汾
msgbox wScript.version, 0, "Version" 

''  2. Method                           
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2.1 CreateObject Method
'     �﷨��CreateObject(appname.objectType, [servename])
'     ���ͣ�appname, ��Ҫ�� Variant(�ַ���)���ṩ�ö����Ӧ�ó�������
'           objecttype, ��Ҫ��Variant����������������ͻ����ࡣ
'           servename����ѡ��Variant��Ҫ�����ϴ��������������������ơ�
'
'     ˵����Ҫ����ActiveX����ֻ�轫CreateObject���صĶ��󸳸�һ�����������
'     ���ӣ�Set oExcel = CreateObject("Excel.Application")

' ����һ�������������ʹ�ö�̬�������������ö���
Dim oExcel
Set oExcel = CreateObject("Excel.Application")

' 1) ʹExcel�ɼ�
oExcel.Visible = true

' 2) ����Excel������
oExcel.caption = "qyx's vbs"

' 3) ���һ���µĹ�����
oExcel.workBooks.add

' 4) ���Ѵ��ڵĹ�����
' oExcel.workbooks.open("d:\temp.xlsx")

' 5) ���õ�2������ҳΪ�������
oExcel.worksheets(2).activate
' ����
' oExcel.worksheets("Sheet2").activate

' 6) ����Ԫ��ֵ
oExcel.cells(1,1).value = "This is column A, row 1"

' 7) ����ָ���еĸ߶ȣ���λ����, 0.035cm��
oExcel.activeSheet.rows(2).rowHeight = 1/0.035 ' 1cm

' 8) ����ָ���еĿ�ȣ���λ���ַ�������
oExcel.activeSheet.columns(1).columnWidth = 5

' 9) �ڵ�8��֮ǰ�����ҳ��
oExcel.worksheets(1).rows(8).pagebreak = 1

' 10) �ڵ�8��֮ǰɾ����ҳ��
oExcel.worksheets(1).columns(8).pagebreak = 0

' 11) ָ���߿��߿��
'     ˵����1-�� 2-�� 3-�� 4-�� 5-\ 6-/
oExcel.activeSheet.range("B3:D4").borders(5).weight = 3

' 12) �����1�е�4�е�Ԫ��ʽ
oExcel.activeSheet.cells(1,4).clearcontents
' oExcel.activeSheet.cells(1,4).value = ""

' 13) ���õ�һ����������
oExcel.activeSheet.rows(1).font.name = "����"
oExcel.activesheet.rows(1).font.color = vbRed
oExcel.activeSheet.rows(1).font.bold = true
oExcel.activesheet.rows(1).font.underLine = true

' 14) ҳ������
' a) ҳü
oExcel.activeSheet.pageSetup.centerHeader = "������ʾ"

' b) ҳ��
oExcel.activeSheet.pageSetup.centerFooter = "��&Pҳ"

' c) ҳü�����˱߾�2cm
oExcel.activeSheet.pageSetup.headerMargin = 2/0.035 

' d) ҳ�ŵ��׶˱߾�3cm
oExcel.activeSheet.pageSetup.footerMargin = 3/0.035

' e) ���߾�2cm
oExcel.activeSheet.pageSetup.topMargin = 2/0.035

' f) �ױ߾�2cm
oExcel.activeSheet.pageSetup.bottomMargin = 2/0.035

' g) ��߾�2cm
oExcel.activeSheet.pageSetup.leftMargin = 2/0.035

' h) �ұ߾�2cm
oExcel.activeSheet.pageSetup.rightMargin = 2/0.035

' i) ҳüˮƽ����
oExcel.activeSheet.pageSetup.centerVertically = 2/0.035

' k) ��ӡ��Ԫ������
oExcel.activeSheet.pageSetup.printGridLines = true

' 15) ������ճ������
' a) ��������������
' oExcel.activeSheet.copy	' δ����

' b) ����ָ������
oExcel.activeSheet.range("A1:E2").copy

' c) ��A1λ�ÿ�ʼճ��
oExcel.activeSheet.range("A1").pasteSpecial

' d) ���ļ�β����ʼճ��
' oExcel.activeSheet.range.pasteSpecial 'δ����

' 16) ����һ�л�һ��
oExcel.activeSheet.rows(2).insert
oExcel.activeSheet.columns(1).insert

' 17) ɾ��һ�л�һ��
oExcel.activeSheet.rows(2).delete
oExcel.activeSheet.columns(1).delete

' 18) ��ӡԤ��������
oExcel.activeSheet.printPreview

' 19) ��ӡ���������
oExcel.activeSheet.printOut

' 20) ��������
oExcel.activeWorkBook.saveAs "d:\te.xls", 56

' 21) �ر��˳�
' �رչ�����
oExcel.activeWorkBook.close

' ʹ��Ӧ�ó�������quit�����ر�Excel
oExcel.Quit

' �ͷŸö������
Set oExcel = Nothing

