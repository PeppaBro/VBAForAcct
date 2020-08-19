Attribute VB_Name = "Proc"
Public XorKey(8) As Byte
Public rst()
Public SqlStr$, sCode$
Public CnnServer As Object
'Subs & Functions without any error handling

Private Sub InitXorKey()
    XorKey(0) = CByte("&H" & "B2")
    XorKey(1) = CByte("&H" & "09")
    XorKey(2) = CByte("&H" & "AA")
    XorKey(3) = CByte("&H" & "55")
    XorKey(4) = CByte("&H" & "93")
    XorKey(5) = CByte("&H" & "6D")
    XorKey(6) = CByte("&H" & "84")
    XorKey(7) = CByte("&H" & "47")
End Sub

Private Function SimpleEncrypt(Str As String) As String
    Dim i%, j%
    SimpleEncrypt = ""
    j = 0
    For i = 1 To Len(Str)
        SimpleEncrypt = SimpleEncrypt & WorksheetFunction.Dec2Hex(CByte(Asc(Mid(Str, i, 1))) Xor XorKey(j), 2)
        j = (j + 1) Mod 8
    Next
End Function

Private Function SimpleDecrypt(Str As String) As String
    Dim i%, j%
    SimpleDecrypt = ""
    j = 0
    For i = 1 To Int(Len(Str) / 2)
        SimpleDecrypt = SimpleDecrypt & Chr(WorksheetFunction.Hex2Dec(Mid(Str, i * 2 - 1, 2)) Xor XorKey(j))
        j = (j + 1) Mod 8
    Next
End Function

Public Sub RegUserConfigures()
    InitXorKey
    SaveSetting appname:="BYMY", section:="SQL Server", Key:="UserName", Setting:=Application.InputBox(prompt:="请输入用户名", Title:="SQL Server用户名", Default:="SA")
    SaveSetting appname:="BYMY", section:="SQL Server", Key:="UserPassword", Setting:=SimpleEncrypt(Application.InputBox(prompt:="请输入密码", Title:="SQL Server密码"))
End Sub

Private Function GetUserName() As String
    GetUserName = GetSetting("BYMY", "SQL Server", "UserName")
End Function

Private Function GetUserPassword() As String
    GetUserPassword = SimpleDecrypt(GetSetting("BYMY", "SQL Server", "UserPassword"))
End Function

Private Sub ReVal()
    rst = CnnServer.Execute(SqlStr).GetRows
End Sub

Private Function Indent(t As Integer) As String
    Indent = WorksheetFunction.Rept(Chr(vbKeySpace), 4 * t)
End Function

'Private Function SQLRetIC() As String
'    SQLRetIC = SQLRetIC & "SELECT citemcode "
'    SQLRetIC = SQLRetIC & "   ,citemname "
'    SQLRetIC = SQLRetIC & "FROM GL_AllItemName "
'    SQLRetIC = SQLRetIC & "WHERE citem_class = '00'"
'End Function

Private Function SQLRet(TableSource As String, ColumnCode As String, ColumnName As String, Optional SearchCondition As String = "") As String
    SQLRet = SQLRet & "SELECT " & ColumnCode & vbCrLf
    SQLRet = SQLRet & "," & ColumnName & vbCrLf
    SQLRet = SQLRet & "FROM " & TableSource & vbCrLf
    If SearchCondition <> "" Then SQLRet = SQLRet & "WHERE " & SearchCondition
End Function

'Private Sub GetRst()
'    Dim ConStr$, usrN$, usrPW$
'    InitXorKey
'    usrN = GetUserName()
'    usrPW = GetUserPassword()
'    ConStr = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & usrN & ";pAssword = " & usrPW & ";Initial Catalog=UFDATA_006_2017;Data Source=192.168.10.250;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
'    Set CnnServer = CreateObject("ADODB.Connection")
'    CnnServer.Open ConStr
'    ReVal
'    CnnServer.Close: Set CnnServer = Nothing
'End Sub

'Private Function CodeRetIC() As String
'    Dim U%, i%
'    U = UBound(rst, 2)
'    CodeRetIC = "Public Function RetIC(Str As String) As String" & vbCrLf
'    CodeRetIC = CodeRetIC & Indent(1) & "Dim ArrI(0 To " & U & ") As String, ArrC(0 To " & U & ") As String" & vbCrLf
'    For i = 0 To U
'        CodeRetIC = CodeRetIC & Indent(1) & "ArrI(" & i & ") = " & Chr(34) & rst(0, i) & Chr(34) & ":ArrC(" & i & ") = " & Chr(34) & rst(1, i) & Chr(34) & vbCrLf
'    Next
'    CodeRetIC = CodeRetIC & Indent(1) & "If IsNumeric(Str) Then" & vbCrLf
'    CodeRetIC = CodeRetIC & Indent(2) & "If IsInArray(Str, ArrI) Then RetIC = ArrC(ArrIdx(Str, ArrI)): Exit Function" & vbCrLf
'    CodeRetIC = CodeRetIC & Indent(1) & "Else" & vbCrLf
'    CodeRetIC = CodeRetIC & Indent(2) & "If IsInArray(Str, ArrC) Then RetIC = ArrI(ArrIdx(Str, ArrC)): Exit Function" & vbCrLf
'    CodeRetIC = CodeRetIC & Indent(1) & "End If" & vbCrLf
'    CodeRetIC = CodeRetIC & Indent(1) & "RetIC = " & Chr(34) & Chr(34) & vbCrLf
'    CodeRetIC = CodeRetIC & "End Function"
'End Function

Public Function CodeStr(FuncName As String, ArrCode As String, ArrName As String) As String
    Dim U%, i%
    U = UBound(rst, 2)
    CodeStr = "Public Function " & FuncName & "(Str As String) As String" & vbCrLf
    CodeStr = CodeStr & Indent(1) & "Dim " & ArrCode & "(0 To " & U & ") As String, " & ArrName & "(0 To " & U & ") As String" & vbCrLf
    For i = 0 To U
        CodeStr = CodeStr & Indent(1) & ArrCode & "(" & i & ") = " & Chr(34) & rst(0, i) & Chr(34) & ":" & ArrName & "(" & i & ") = " & Chr(34) & rst(1, i) & Chr(34) & vbCrLf
    Next
    CodeStr = CodeStr & Indent(1) & "If Str = " & Chr(34) & Chr(34) & "Then " & FuncName & " = " & Chr(34) & Chr(34) & " : Exit Function" & vbCrLf
    CodeStr = CodeStr & Indent(1) & "If IsNumeric(Str) Then" & vbCrLf
    CodeStr = CodeStr & Indent(2) & "If IsInArray(Str, " & ArrCode & ") Then " & FuncName & " = " & ArrName & "(ArrIdx(Str, " & ArrCode & ")): Exit Function" & vbCrLf
    CodeStr = CodeStr & Indent(1) & "Else" & vbCrLf
    CodeStr = CodeStr & Indent(2) & "If IsInArray(Str, " & ArrName & ") Then " & FuncName & " = " & ArrCode & "(ArrIdx(Str, " & ArrName & ")): Exit Function" & vbCrLf
    CodeStr = CodeStr & Indent(1) & "End If" & vbCrLf
    CodeStr = CodeStr & Indent(1) & "" & FuncName & " = " & Chr(34) & Chr(34) & vbCrLf
    CodeStr = CodeStr & "End Function"
End Function

Public Function IsInArray(mStr As String, Arr As Variant) As Boolean
  IsInArray = (UBound(Filter(Arr, mStr)) > -1)
End Function

Public Function ArrIdx(mStr As String, Arr As Variant) As Integer
    Dim i%
    For i = 0 To UBound(Arr)
        If Arr(i) = mStr Then ArrIdx = i: Exit For
    Next
End Function

'Public Sub UpdateRetIC()
'    Dim s%, C%, CodeStr$
'    GetRst
'    CodeStr = CodeRetIC
'    With Application.VBE.VBProjects("UserDefinedFunctions").VBComponents("Func").CodeModule
'        s = .ProcStartLine("RetIC", vbext_pk_Proc)
'        C = .ProcCountLines("RetIC", vbext_pk_Proc)
'        .DeleteLines s, C
'        .AddFromString CodeStr
'    End With
'End Sub

Private Sub OpenCnn()
    Dim ConStr$, usrN$, usrPW$
    InitXorKey
    usrN = GetUserName()
    usrPW = GetUserPassword()
    ConStr = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & usrN & ";pAssword = " & usrPW & ";Initial Catalog=UFDATA_006_2017;Data Source=192.168.10.250;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False"
    Set CnnServer = CreateObject("ADODB.Connection")
    CnnServer.Open ConStr
End Sub

Private Sub CloseCnn()
    CnnServer.Close: Set CnnServer = Nothing
End Sub

Private Sub DelFunc(FuncName As String)
    Dim s%, c%
    With Application.VBE.VBProjects("UserDefinedFunctions").VBComponents("Func").CodeModule
        s = .ProcStartLine(FuncName, vbext_pk_Proc)
        c = .ProcCountLines(FuncName, vbext_pk_Proc)
        .DeleteLines s, c
    End With
End Sub

Private Sub AddFunc(Str As String)
    Application.VBE.VBProjects("UserDefinedFunctions").VBComponents("Func").CodeModule.AddFromString Str
End Sub

Public Sub UpdFunc()
'+----------+-----------------+--------------+--------------+--------------------+
'! FuncName !   TableSource   !  ColumnCode  !  ColumnName  ! SearchCondition    !
'+----------+-----------------+--------------+--------------+--------------------+
'! RetItem  !  GL_AllItemName !  citemcode   !  citemname   ! citem_class = '00' !
'+----------+-----------------+--------------+--------------+--------------------+
'! RetVen   !     vendor      !   cVenCode   !   cVenName   !                    !
'+----------+-----------------+--------------+--------------+--------------------+
'! RetCus   !    Customer     !   cCusCode   !   cCusName   !                    !
'+----------+-----------------+--------------+--------------+--------------------+
'! RetCode  !      code       !    ccode     !  ccode_name  !  iyear = 2020      !
'+----------+-----------------+--------------+--------------+--------------------+
On Error Resume Next
    DelFunc "RetItem"
    DelFunc "RetVen"
    DelFunc "RetCus"
    DelFunc "RetCode"
On Error GoTo 0
    OpenCnn
    
    SqlStr = SQLRet("GL_AllItemName", "citemcode", "citemname", "citem_class = '00'"): ReVal
    sCode = CodeStr("RetItem", "citemcode", "citemname"): AddFunc sCode
    
    SqlStr = SQLRet("vendor", "cVenCode", "cVenName"): ReVal
    sCode = CodeStr("RetVen", "cVenCode", "cVenName"): AddFunc sCode
    
    SqlStr = SQLRet("Customer", "cCusCode", "cCusName"): ReVal
    sCode = CodeStr("RetCus", "cCusCode", "cCusName"): AddFunc sCode

    SqlStr = SQLRet("code", "ccode", "ccode_name", "iyear = 2020"): ReVal
    sCode = CodeStr("RetCode", "ccode", "ccode_name"): AddFunc sCode

    CloseCnn
End Sub



