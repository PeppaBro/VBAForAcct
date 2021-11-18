Sub 购销合同()
Dim myMerge As MailMerge, I As Integer, myname As String, Mypath As String
Application.ScreenUpdating = False
Mypath = ActiveDocument.Path & "\"
If Dir(Mypath & "合同及收货证明" & Format(Date, "YYMMDD"), vbDirectory) <> "" Then
   Else
       MkDir Mypath & "合同及收货证明" & Format(Date, "YYMMDD")
 End If
Set myMerge = ActiveDocument.MailMerge
With myMerge.DataSource
    If .Parent.State = wdMainAndDataSource Then
        .ActiveRecord = wdFirstRecord
        For I = 1 To .RecordCount
            .FirstRecord = I
            .LastRecord = I
            .Parent.Destination = wdSendToNewDocument
            myname = .DataFields(4).Value & "-" & .DataFields(2).Value & "-购销合同"
            .ActiveRecord = wdNextRecord
            .Parent.Execute  '每次合并一个数据记录
           With ActiveDocument
                .Content.Characters.Last.Previous.Delete  '删除分节符
                .SaveAs Filename:=Mypath & "合同及收货证明" & Format(Date, "YYMMDD") & "\" & myname & ".doc", FileFormat:=wdFormatDocument97
                .Close  '关闭生成的文档（已保存）
            End With
        
        Next
    End If
End With
Application.ScreenUpdating = True
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 收货证明()
'主文档的类型为信函
'合并全部数据记录
'假设主文档已链接好数据源，可以进行正常的邮件合并
Dim myMerge As MailMerge, I As Integer, myname As String, Mypath As String
Application.ScreenUpdating = False
Mypath = ActiveDocument.Path & "\"
If Dir(Mypath & "合同及收货证明" & Format(Date, "YYMMDD"), vbDirectory) <> "" Then
   Else
       MkDir Mypath & "合同及收货证明" & Format(Date, "YYMMDD")
 End If
Set myMerge = ActiveDocument.MailMerge
With myMerge.DataSource
    If .Parent.State = wdMainAndDataSource Then
        .ActiveRecord = wdFirstRecord
        For I = 1 To .RecordCount
            .FirstRecord = I
            .LastRecord = I
            .Parent.Destination = wdSendToNewDocument
            myname = .DataFields(4).Value & "-" & .DataFields(2).Value & "-收货证明"
            .ActiveRecord = wdNextRecord
            .Parent.Execute  '每次合并一个数据记录
           With ActiveDocument
                .Content.Characters.Last.Previous.Delete  '删除分节符
                .SaveAs Filename:=Mypath & "合同及收货证明" & Format(Date, "YYMMDD") & "\" & myname & ".doc", FileFormat:=wdFormatDocument97
                .Close  '关闭生成的文档（已保存）
            End With
        
        Next
    End If
End With
Application.ScreenUpdating = True
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 保越分割()
Dim SrcDoc As Document, NewDoc As Document
Dim NewDocName As String, Subname As String
Dim N As Integer, I As Integer
Dim SR As Range, NR As Range
Dim Mypath As String, Cop As String, Con As String
Dim RStr As String
Dim RREs As Long
    Set SrcDoc = ActiveDocument
    Mypath = "D:\Contracts\保越\"
    Subname = Trim(Split(SrcDoc.Name, ".")(0))
    If Dir(Mypath & Format(Date, "YYMMDD"), vbDirectory) <> "" Then
    Else
    MkDir Mypath & Format(Date, "YYMMDD")
    End If
    If Dir(Mypath & Format(Date, "YYMMDD") & "\" & Subname, vbDirectory) <> "" Then
    Else
    MkDir Mypath & Format(Date, "YYMMDD") & "\" & Subname
    End If
    Set SR = SrcDoc.Content
    N = ActiveDocument.Content.Information(wdNumberOfPagesInDocument)
    SR.Collapse wdCollapseStart
    SR.Select
    For I = 1 To N Step 1
        Set NewDoc = Documents.Add
        SrcDoc.Activate
        SrcDoc.Bookmarks("\page").Range.Copy
        SrcDoc.Windows(1).Activate
        Application.Browser.Target = wdBrowsePage
        Application.Browser.Next
        NewDoc.Activate
        NewDoc.Windows(1).Selection.Paste
        ActiveDocument.Content.Characters.Last.Previous.Delete
        Set NR = NewDoc.Content
        NR.SetRange Start:=NewDoc.Paragraphs(21).Range.Start, End:=NewDoc.Paragraphs(21).Range.End
        Cop = NR.Text
        NR.SetRange Start:=NewDoc.Paragraphs(3).Range.Start, End:=NewDoc.Paragraphs(3).Range.End
        Con = NR.Text
        NewDocName = Mypath & Format(Date, "YYMMDD") & "\" & Subname & "\" & Cop & "-" & Con & "-收货证明.doc"
        NewDocName = Replace(NewDocName, Chr(10), "", , , vbBinaryCompare)
        NewDocName = Replace(NewDocName, Chr(13), "", , , vbBinaryCompare)
        NewDocName = Replace(NewDocName, Chr(32), "", , , vbBinaryCompare)
        NewDoc.SaveAs NewDocName
        NewDoc.Close False
        Set NR = Nothing
    Next I
    Set SR = Nothing
    Set NewDoc = Nothing
    SrcDoc.Close False
    Set SrcDoc = Nothing
    RStr = "C:\Program Files\WinRAR\WinRAR.exe  a -ep1 " & Mypath & Format(Date, "YYMMDD") & "\" & Subname & ".rar" & " " & Mypath & Format(Date, "YYMMDD") & "\" & Subname & "\*.*"
    RREs = Shell(RStr, vbHide)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 保越分割PDF()
    Dim NewDocName$, SubName$, CoName$, ContrNum$, tPage%, MyPath$, i%, RarStr$, RarRlt&, EplRlt&
    Dim SrcDoc As Document, NewDoc As Document
    Dim SrcRng As Range, NewRng As Range 
    Application.ScreenUpdating = False
    Set SrcDoc = ActiveDocument
    tPage = SrcDoc.BuiltInDocumentProperties(wdPropertyPages)
    ActiveDocument.Sections(1).Range.Select  
    Do Until (Selection.Information(wdActiveEndPageNumber) = tPage)
        If Selection.Find.Execute("商品明细") Then
            Application.Run "RestartNumbering"
        End If
        Selection.Range.Next(Unit:=wdSection, Count:=1).Select
        DoEvents
    Loop
    If Selection.Find.Execute("商品明细") Then Application.Run "RestartNumbering"
    Beep
    MyPath = "D:\Contracts\保越\"
    SubName = Trim(Split(SrcDoc.Name, ".")(0))
    SubName = Replace(SubName, Chr(10), "", , , vbBinaryCompare)
    SubName = Replace(SubName, Chr(13), "", , , vbBinaryCompare)
    SubName = Replace(SubName, Chr(32), "", , , vbBinaryCompare)  
    If Dir(MyPath & Format(Date, "YYMMDD"), vbDirectory) <> "" Then
    Else
        MkDir MyPath & Format(Date, "YYMMDD")
    End If
    If Dir(MyPath & Format(Date, "YYMMDD") & "\" & SubName, vbDirectory) <> "" Then
    Else
        MkDir MyPath & Format(Date, "YYMMDD") & "\" & SubName
    End If
    Set SrcRng = SrcDoc.Content
    SrcRng.Collapse wdCollapseStart
    SrcRng.Select
    For i = 1 To tPage Step 2
        Set NewDoc = Documents.Add
        SrcDoc.Activate
        SrcDoc.Bookmarks("\Section").Range.Copy
        SrcDoc.Windows(1).Activate
        Application.Browser.Target = wdBrowseSection
        Application.Browser.Next
        NewDoc.Activate
        NewDoc.Windows(1).Selection.Paste
        Set NewRng = NewDoc.Content
        ContrNum = Trim(Replace(NewRng.Paragraphs(2).Range.Text, "编号：", ""))
        CoName = Trim(Replace(NewRng.Paragraphs(8).Range.Text, "买方：", ""))
        NewDocName = MyPath & Format(Date, "YYMMDD") & "\" & SubName & "\" & CoName & "-" & ContrNum & ".pdf"
        NewDocName = Replace(NewDocName, Chr(10), "", , , vbBinaryCompare)
        NewDocName = Replace(NewDocName, Chr(13), "", , , vbBinaryCompare)
        NewDocName = Replace(NewDocName, Chr(32), "", , , vbBinaryCompare)
        With NewDoc
            .Content.Characters.Last.Previous.Delete
            .SaveAs FileName:=NewDocName, FileFormat:=wdFormatPDF
            .Close wdDoNotSaveChanges
        End With
        Set NewRng = Nothing
    Next i
    Beep
    Set NewDoc = Nothing
    Set SrcRng = Nothing
    SrcDoc.Close wdDoNotSaveChanges
    Set SrcDoc = Nothing
    Application.ScreenUpdating = True  
    RarStr = "C:\Program Files\WinRAR\WinRAR.exe  a -ep1 " & MyPath & Format(Date, "YYMMDD") & "\" & SubName & ".rar" & " " & MyPath & Format(Date, "YYMMDD") & "\" & SubName & "\*.*"
    RarRlt = Shell(RarStr, vbHide)
    EplRlt = Shell("explorer.exe /n, /e," & MyPath & Format(Date, "YYMMDD") & "\" & SubName, vbNormalFocus)
    If Word.Application.Windows.Count = 0 Then Application.Quit
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub 提货委托书()
Dim myMerge As MailMerge, I As Integer, myname As String, Mypath As String
Application.ScreenUpdating = False
Mypath = ActiveDocument.Path & "\"
If Dir(Mypath & "合同及收货证明" & Format(Date, "YYMMDD"), vbDirectory) <> "" Then
   Else
       MkDir Mypath & "合同及收货证明" & Format(Date, "YYMMDD")
 End If
Set myMerge = ActiveDocument.MailMerge
With myMerge.DataSource
    If .Parent.State = wdMainAndDataSource Then
        .ActiveRecord = wdFirstRecord
        For I = 1 To .RecordCount
            .FirstRecord = I
            .LastRecord = I
            .Parent.Destination = wdSendToNewDocument
            myname = .DataFields(4).Value & "-" & .DataFields(2).Value 
            .ActiveRecord = wdNextRecord
            .Parent.Execute  '每次合并一个数据记录
           With ActiveDocument
                .Content.Characters.Last.Previous.Delete  '删除分节符
                .SaveAs Filename:=Mypath & "合同及收货证明" & Format(Date, "YYMMDD") & "\" & myname & ".pdf", FileFormat:=wdFormatPDF
                .Close wdDoNotSaveChanges
            End With
        Next
    End If
End With
Application.ScreenUpdating = True
End Sub
