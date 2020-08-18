Sub createEx()
Application.ScreenUpdating = False
Const sfPath = "D:\Ex\sf-express\"
Const stoPath = "D:\Ex\sto\"
Const Sender = "佩奇"
Const SenderCop = "中国进出口贸易有限公司"
Const SenderAddr = "唐人街666号"
Const SenderTel = "18888888888"
Dim N, CounterEx, sRow, eRow
sRow = 0
eRow = 0
N = Sheets("流水").[A65536].End(xlUp).Row
Debug.Print Sheets("流水").Name
For CounterEx = 1 To N
    If Sheets("流水").Cells(CounterEx, 1).Font.Color = RGB(0, 255, 0) Then sRow = CounterEx
    If Sheets("流水").Cells(CounterEx, 1).Font.Color = RGB(255, 0, 0) Then eRow = CounterEx
Next CounterEx
    If eRow = 0 Then eRow = sRow
Dim ArrEx As Variant
'getExpressData
ArrEx = Sheets("流水").Range("A" & sRow & ":I" & eRow)
Dim RowEx, ColEx
RowEx = UBound(ArrEx, 1)
ColEx = UBound(ArrEx, 2)
Dim sfArr() As Variant
Dim stoArr() As Variant
Dim CounterSf As Integer
Dim CounterSto As Integer
CounterSf = 0
CounterSto = 0
Dim i
For i = 1 To RowEx
    Select Case (ArrEx(i, 9))
    Case "申通"
        CounterSto = CounterSto + 1
        ReDim Preserve stoArr(1 To 4, 1 To CounterSto)
        stoArr(1, CounterSto) = ArrEx(i, 1)
        stoArr(2, CounterSto) = ArrEx(i, 3)
        stoArr(3, CounterSto) = ArrEx(i, 6) & ArrEx(i, 2)
        stoArr(4, CounterSto) = ArrEx(i, 7)
    Case "顺丰月结"
'        ReDim Preserve sfArr(1 To 57, 1 To counterSf)
'        sfArr(1, counterSf) = counterSf
'        sfArr(2, counterSf) = ArrEx(i, 1)
'        sfArr(7, counterSf) = senderCop
'        sfArr(8, counterSf) = sender
'        sfArr(9, counterSf) = senderTel
'        sfArr(10, counterSf) = senderAddr
'        sfArr(11, counterSf) = ArrEx(i, 2)
'        sfArr(12, counterSf) = ArrEx(i, 3)
'        sfArr(13, counterSf) = ArrEx(i, 7)
'        sfArr(14, counterSf) = ArrEx(i, 6)
'        sfArr(15, counterSf) = "文件"
'        sfArr(21, counterSf) = "1"
'        sfArr(22, counterSf) = "顺丰标快（陆运）"
'        sfArr(23, counterSf) = "寄付月结"
'        sfArr(24, counterSf) = "5981455314"
        CounterSf = CounterSf + 1
        ReDim Preserve sfArr(1 To 44, 1 To CounterSf)
        sfArr(1, CounterSf) = ArrEx(i, 1)
        sfArr(2, CounterSf) = SenderCop
        sfArr(3, CounterSf) = Sender
        sfArr(4, CounterSf) = SenderTel
        sfArr(5, CounterSf) = SenderAddr
        sfArr(6, CounterSf) = ArrEx(i, 2)
        sfArr(7, CounterSf) = ArrEx(i, 3)
'        sfArr(8, CounterSf) = ArrEx(i, 7)
        sfArr(9, CounterSf) = ArrEx(i, 7)
        sfArr(10, CounterSf) = ArrEx(i, 6)
        sfArr(11, CounterSf) = "文件"
        sfArr(12, CounterSf) = "1"
        sfArr(15, CounterSf) = "寄付月结"
        sfArr(16, CounterSf) = "顺丰标快（陆运）"
        sfArr(17, CounterSf) = "1"
    
    Case "顺丰到付"
'        CounterSf = CounterSf + 1
'        ReDim Preserve sfArr(1 To 57, 1 To CounterSf)
'        sfArr(1, CounterSf) = CounterSf
'        sfArr(2, CounterSf) = ArrEx(i, 1)
'        sfArr(7, CounterSf) = SenderCop
'        sfArr(8, CounterSf) = Sender
'        sfArr(9, CounterSf) = SenderTel
'        sfArr(10, CounterSf) = SenderAddr
'        sfArr(11, CounterSf) = ArrEx(i, 2)
'        sfArr(12, CounterSf) = ArrEx(i, 3)
'        sfArr(13, CounterSf) = ArrEx(i, 7)
'        sfArr(14, CounterSf) = ArrEx(i, 6)
'        sfArr(15, CounterSf) = "文件"
'        sfArr(17, CounterSf) = "1"
'        sfArr(21, CounterSf) = "1"
'        sfArr(22, CounterSf) = "顺丰标快（陆运）"
'        sfArr(23, CounterSf) = "到付现结"
        CounterSf = CounterSf + 1
        ReDim Preserve sfArr(1 To 44, 1 To CounterSf)
        sfArr(1, CounterSf) = ArrEx(i, 1)
        sfArr(2, CounterSf) = SenderCop
        sfArr(3, CounterSf) = Sender
        sfArr(4, CounterSf) = SenderTel
        sfArr(5, CounterSf) = SenderAddr
        sfArr(6, CounterSf) = ArrEx(i, 2)
        sfArr(7, CounterSf) = ArrEx(i, 3)
'        sfArr(8, CounterSf) = ArrEx(i, 7)
        sfArr(9, CounterSf) = ArrEx(i, 7)
        sfArr(10, CounterSf) = ArrEx(i, 6)
        sfArr(11, CounterSf) = "文件"
        sfArr(12, CounterSf) = "1"
        sfArr(15, CounterSf) = "到付现结"
        sfArr(16, CounterSf) = "顺丰标快（陆运）"
        sfArr(17, CounterSf) = "1"

    Case Else
    End Select
Next i
Debug.Print ArrEx(1, 1)
Application.DisplayAlerts = False
'sf-init
Dim sfEx As Workbook
Set sfEx = Workbooks.Add(xlWBATWorksheet)
Cells.NumberFormatLocal = "@"
Cells.Font.Name = "宋体"
Cells.Font.Size = 9
'[g1] = "寄件方信息"
'[g1].Resize(1, 4).Merge
'[k1] = "收件方信息"
'[k1].Resize(1, 4).Merge
'[o1] = "商品信息"
'[o1].Resize(1, 6).Merge
'[u1] = "发货信息"
'[u1].Resize(1, 37).Merge
'[a2].Resize(1, 57) = Array("序号", "订单号", "运单号", "子单号", "签回单号", "寄方备注", "寄方公司", "寄方姓名", "寄方联系方式", "寄方地址", "收方公司", "收方姓名", "收方联系方式", "收方地址", "托寄物内容", "商品编码", "托寄物数量", "商品单价/元", "商品货号", "商品属性", "包裹件数", "业务类型", "付款方式", "月结卡号", "包裹重量/KG", "代收金额", "代收卡号", "保价金额", "是否签回单", "派送日期", "派送时段", "是否自取", "是否保单配送", "是否票据专送", "是否易碎宝", "易碎宝服务费/元", "是否口令签收", "标准化包装（元）", "个性化包装（元）", "其它费用（元）", "超长超重服务费", "是否双人派送", "长(cm)", "宽(cm)", "高(cm)", "扩展字段1", "扩展字段2", "扩展字段3", "扩展字段4", "扩展字段5", "温区", "签单返还范本", "保鲜服务", "WOW基础", "WOW尊享", "是否到付优惠", "优惠月结卡号")
'[a1].Resize(2, 57).HorizontalAlignment = xlCenterAcrossSelection
'[a1].Resize(2, 57).VerticalAlignment = xlCenter
'[a3].Resize(counterSf, 57) = Application.Transpose(sfArr)
[A1].Resize(1, 44) = Array("用户订单号", "寄件公司", "寄件人", "寄件电话", "寄件详细地址", "收件公司", "收件人", "收件电话", "收件手机", "收件详细地址", "托寄物内容", "托寄物数量", "包裹重量", "寄方备注", "运费付款方式", "业务类型", "件数", "代收金额", "保价金额", "个性化包装", "签回单", "自取件", "电子验收", "是否超长超重", "超长超重服务费", "保鲜服务", "保单配送", "拍照验证", "票据专送", "口令签收", "等通知派送", "温度追溯（离线）", "是否定时派送", "派送日期", "派送时段", "长（cm）", "宽（cm）", "高（cm）", "体积（cm3）", "扩展字段1", "扩展字段2", "扩展字段3", "扩展字段4", "扩展字段5")
[A1].Resize(1, 44).HorizontalAlignment = xlCenterAcrossSelection
[A1].Resize(1, 44).VerticalAlignment = xlCenter
[A2].Resize(CounterSf, 44) = Application.Transpose(sfArr)
Cells.EntireColumn.AutoFit
ActiveWindow.FreezePanes = False
[c3].Select
ActiveWindow.FreezePanes = True
sfEx.SaveAs Filename:=sfPath & "顺丰" & Format(Date, "YYMMDD") & "@" & sRow & "~" & eRow & ".xls", FileFormat:=xlWorkbookNormal, CreateBackup:=False
sfEx.Close Savechanges:=True
Set sfEx = Nothing
'sto-init
Dim stoEx As Workbook
Set stoEx = Workbooks.Add(xlWBATWorksheet)
Cells.NumberFormatLocal = "@"
Cells.Font.Name = "宋体"
Cells.Font.Size = 9

[A1].Resize(1, 4) = Array("备注", "姓名", "详细地址", "电话")
On Error Resume Next
[A2].Resize(CounterSto, 4) = Application.Transpose(stoArr)
Cells.EntireColumn.AutoFit
ActiveWindow.FreezePanes = False
[b2].Select
ActiveWindow.FreezePanes = True
stoEx.SaveAs Filename:=stoPath & "申通" & Format(Date, "YYMMDD") & "@" & sRow & "~" & eRow & ".xls", FileFormat:=xlWorkbookNormal, CreateBackup:=False
stoEx.Close Savechanges:=True
Set stoEx = Nothing
Application.ScreenUpdating = False
End Sub
