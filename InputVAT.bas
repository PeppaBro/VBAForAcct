Public Sub InputVAT()
    Dim CnnExcel As Object
    Dim SqlStr As String
    Dim Wb As Workbook
    Dim arr
    Set CnnExcel = CreateObject("ADODB.Connection")
    CnnExcel.Open "provider=Microsoft.ACE.OLEDB.12.0;extended properties='excel 8.0;hdr=yes;imex=1';data source=" & ActiveWorkbook.FullName
    SqlStr = "Select 发票代码,发票号码,开票日期,销方名称,金额,税额 From [第1页$]"
    Set Wb = Workbooks.Add(xlWBATWorksheet)
    'NormalFont
    Cells.Font.Name = "宋体"
    Cells.Font.Size = 9
    Cells(2, 1).CopyFromRecordset CnnExcel.Execute(SqlStr)
    CnnExcel.Close: Set CnnExcel = Nothing
    Columns("E:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    'NumberFormatLocal
    Columns("A:B").NumberFormatLocal = "@"
    Columns("D:D").NumberFormatLocal = "@"
    Columns("F:F").NumberFormatLocal = "@"
    Columns("M:Q").NumberFormatLocal = "@"
    Columns("C:C").NumberFormatLocal = "yyyy-mm-dd"
    Columns("E:E").NumberFormatLocal = "G/通用格式"
    Columns("G:G").NumberFormatLocal = "_   * #,##0.0000_ ;_   * -#,##0.0000_ ;_   * ""-""????_ ;_ @_ "
    Columns("H:L").NumberFormatLocal = "_   * #,##0.00_ ;_   * -#,##0.00_ ;_   * ""-""??_ ;_ @_ "
    'FormulaR1C1
    Cells(2, 12).FormulaR1C1 = "=ROUND(RC[-2]+RC[-1],2)"
    Range("L2:L" & [j1].CurrentRegion.Rows.Count).FillDown
    [H2].FormulaR1C1 = "=IFERROR(RC[2]/RC[-1],"""")"
    
    [i2].FormulaR1C1 = "=IFERROR(RC[3]/RC[-2],"""")"
    Range("H2:I" & [j1].CurrentRegion.Rows.Count).FillDown
    'Title
    [a1].Resize(1, 17) = [{"发票代码","发票号码","开票日期","销方名称","存货编码","品名","数量","不含税单价","含税单价","金额","税额","价税合计","本期","类别","FSC声明","备注","辅助品名"}]
    'TypeSetting
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, FieldInfo:=Array(1, 2)
    Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, FieldInfo:=Array(1, 2)
    Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, FieldInfo:=Array(1, 1)
    Columns("D:D").TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, FieldInfo:=Array(1, 1)
    Columns("J:J").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, FieldInfo:=Array(1, 1)
    Columns("K:K").TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, FieldInfo:=Array(1, 1)
    Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, FieldInfo:=Array(1, 1)
    ActiveWindow.FreezePanes = False
    Range("E2").Select
    ActiveWindow.FreezePanes = True
    Columns("A:Q").AutoFilter
    Columns("A:Q").EntireColumn.AutoFit
End Sub