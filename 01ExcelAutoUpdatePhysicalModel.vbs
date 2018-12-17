'开始
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
   Set excel = x1.Workbooks.Open(ExcelPath) '指定 excel文档路径
   for j=1 to excel.Worksheets.Count
      Set oSheet = excel.Sheets(j)
      oSheet.Activate
      'doCheckSheet x1, mdl
      newTable x1, mdl
      'x1.Workbooks(1).Worksheets("d_site").Activate '指定要打开的sheet名称
      'if j =3 then
      '   exit for
      'end if
   Next
   MsgBox "生成数据 表结构共计 " + CStr(count), vbOK + vbInformation, " 表"
Else
   HaveExcel = False
End If


'a x1, mdl
sub doCheckSheet(x1, mdl)
   'MsgBox oSheet.name
   if oSheet.name = "版本更新记录" Or oSheet.name = "汇总" Then
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
   set getCol = table.Columns.CreateNew '创建一列/字段
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
   if oSheet.name = "版本更新记录" Or oSheet.name = "汇总" Then
      'MsgBox oSheet.name
      Exit Sub
   End If
   tableName = oSheet.name
   'while getTable(tableName)=1
   '   tableName = tableName+"_"
   'wend
   
   set table = getTable(tableName)
   'set table = mdl.Tables.CreateNew '创建一个 表实体
   table.Name = tableName '指定 表名，如果在 Excel文档里有，也可以 .Cells(rwIndex, 3).Value 这样指定
   table.Code = tableName '指定 表名
   count = count + 1
   For rwIndex = 3 To 1000 '指定要遍历的 Excel行标 由于第1行是 表头， 从第2行开始
      'th x1.Workbooks(1).Worksheets("d_site")
      'MsgBox oSheet.Cells(rwIndex, 1).Value
      If oSheet.Cells(rwIndex, 1).Value = "" Then
         Exit For
      End If

      'set col = getCol(oSheet.Cells(rwIndex, 2).Value,table)
      set col = table.Columns.CreateNew '创建一列/字段
      col.Name = oSheet.Cells(rwIndex, 2).Value
      'MsgBox col.Name, vbOK + vbInformation, "列"
      col.Code = oSheet.Cells(rwIndex, 2).Value '指定列名
      col.DataType = oSheet.Cells(rwIndex, 4).Value '指定列数据类型
      col.Comment = oSheet.Cells(rwIndex, 3).Value '指定列说明
      If oSheet.Cells(rwIndex, 5).Value = "PK" Then
         col.Primary = true '指定主键
      End If
      'End With
   Next
   Exit Sub
End sub
