'��ʼ
Option Explicit

Dim mdl ' the current model
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "There is no Active Model"
End If

Dim HaveExcel
Dim RQ
RQ = vbYes 'MsgBox("Is  Excel Installed on your machine ?", vbYesNo + vbInformation, "Confirmation")
If RQ = vbYes Then
   HaveExcel = True
   ' Open & Create  Excel Document
   Dim x1
   Dim excel
   Dim oSheet
   Dim j
   Dim count
   Dim ExcelPath
   ExcelPath = CreateObject("WScript.Shell").Exec("mshta vbscript:""<input type=file id=f><script>f.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(f.value)[close()];</script>""").StdOut.ReadAll
   if instr (ExcelPath,"[") <> 0 then
      ExcelPath = left(ExcelPath, instr(ExcelPath,"[")-1)
   end if
   Set x1 = CreateObject("Excel.Application")
   Set excel = x1.Workbooks.Open(ExcelPath) 'ָ�� excel�ĵ�·��
   for j=1 to excel.Worksheets.Count
      Set oSheet = excel.Sheets(j)
      oSheet.Activate
      'doCheckSheet x1, mdl
      newTable x1, mdl
      'x1.Workbooks(1).Worksheets("d_site").Activate 'ָ��Ҫ�򿪵�sheet����
      'if j =3 then
      '   exit for
      'end if
   Next
   MsgBox "�������� ��ṹ���� " + CStr(count), vbOK + vbInformation, " ��"
Else
   HaveExcel = False
End If


'a x1, mdl
sub doCheckSheet(x1, mdl)
   'MsgBox oSheet.name
   if oSheet.name = "�汾���¼�¼" Or oSheet.name = "����" Then
      Exit Sub
   End If
   
   newTable x1, mdl
end sub

function getTable(tname)
   dim t_table
   dim column
   for each t_table in mdl.tables
      if(tname = t_table.name) then
         for each column in t_table.columns
            column.delete()
         next
         set getTable = t_table
         exit function
      end if
   next
   set getTable = mdl.Tables.CreateNew
end function



function getCol(cname,table)
   dim column
   for each column in table.columns
      if(cname = column.name) then
         set getCol = column
         exit function
      end if
   next
   set getCol = table.Columns.CreateNew '����һ��/�ֶ�
end function

sub newTable(x1, mdl)
   dim rwIndex 
   dim tableName
   dim colname
   dim comment
   dim col
   dim table
   dim suffix
   
   'on error Resume Next
   if oSheet.name = "�汾���¼�¼" Or oSheet.name = "����" Then
      'MsgBox oSheet.name
      Exit Sub
   End If
   tableName = oSheet.name
   'while getTable(tableName)=1
   '   tableName = tableName+"_"
   'wend
   
   set table = getTable(tableName)
   'set table = mdl.Tables.CreateNew '����һ�� ��ʵ��
   table.Name = tableName 'ָ�� ����������� Excel�ĵ����У�Ҳ���� .Cells(rwIndex, 3).Value ����ָ��
   table.Code = tableName 'ָ�� ����
   count = count + 1
   For rwIndex = 3 To 1000 'ָ��Ҫ������ Excel�б� ���ڵ�1���� ��ͷ�� �ӵ�2�п�ʼ
      'th x1.Workbooks(1).Worksheets("d_site")
      'MsgBox oSheet.Cells(rwIndex, 1).Value
      If oSheet.Cells(rwIndex, 1).Value = "" Then
         Exit For
      End If

      'set col = getCol(oSheet.Cells(rwIndex, 2).Value,table)
      set col = table.Columns.CreateNew '����һ��/�ֶ�
      col.Name = oSheet.Cells(rwIndex, 2).Value
      'MsgBox col.Name, vbOK + vbInformation, "��"
      col.Code = oSheet.Cells(rwIndex, 2).Value 'ָ������
      col.DataType = oSheet.Cells(rwIndex, 4).Value 'ָ������������
      col.Comment = oSheet.Cells(rwIndex, 3).Value 'ָ����˵��
      If oSheet.Cells(rwIndex, 5).Value = "PK" Then
         col.Primary = true 'ָ������
      End If
      'End With
   Next
   Exit Sub
End sub
