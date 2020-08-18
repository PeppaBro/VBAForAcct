VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private dict_pKeyValues As Collection
Private dict_pKeys() As Variant
Private dict_pItems() As Variant
Private dict_pObjectKeys As Collection
Private dict_pCompareMode As CompareMethod

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

Public Property Get CompareMode() As CompareMethod
    CompareMode = dict_pCompareMode
End Property

Public Property Let CompareMode(Value As CompareMethod)
    If Me.Count > 0 Then
        err.Raise 5 
    End If
    dict_pCompareMode = Value
End Property

Public Property Get Count() As Long
    Count = dict_pKeyValues.Count
End Property

Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Key)
    If Not IsEmpty(dict_KeyValue) Then
        If VBA.IsObject(dict_KeyValue(2)) Then
            Set Item = dict_KeyValue(2)
        Else
            Item = dict_KeyValue(2)
        End If
    Else
    End If
End Property

Public Property Let Item(Key As Variant, Value As Variant)
    If Me.Exists(Key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, Value
    Else
        dict_AddKeyValue Key, Value
    End If
End Property

Public Property Set Item(Key As Variant, Value As Variant)
    If Me.Exists(Key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, Value
    Else
        dict_AddKeyValue Key, Value
    End If
End Property

Public Property Let Key(Previous As Variant, Updated As Variant)
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Previous)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_ReplaceKeyValue dict_KeyValue, Updated, dict_KeyValue(2)
    End If
End Property

Public Sub Add(Key As Variant, Item As Variant)
    If Not Me.Exists(Key) Then
        dict_AddKeyValue Key, Item
    Else
        err.Raise 457
    End If
End Sub

Public Function Exists(Key As Variant) As Boolean
    Exists = Not IsEmpty(dict_GetKeyValue(Key))
End Function

Public Function Items() As Variant
    If Me.Count > 0 Then
        Items = dict_pItems
    Else
        Items = VBA.Split("")
    End If
End Function

Public Function Keys() As Variant
    If Me.Count > 0 Then
        Keys = dict_pKeys
    Else
        Keys = VBA.Split("")
    End If
End Function

Public Sub Remove(Key As Variant)
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Key)
    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_RemoveKeyValue dict_KeyValue
    Else
        err.Raise 32811
    End If
End Sub

Public Sub RemoveAll()
    Set dict_pKeyValues = New Collection
    Erase dict_pKeys
    Erase dict_pItems
End Sub

Private Function dict_GetKeyValue(dict_Key As Variant) As Variant
    On Error Resume Next
    dict_GetKeyValue = dict_pKeyValues(dict_GetFormattedKey(dict_Key))
    err.Clear
End Function

Private Sub dict_AddKeyValue(dict_Key As Variant, dict_Value As Variant, Optional dict_Index As Long = -1)
    If Me.Count = 0 Then
        ReDim dict_pKeys(0 To 0)
        ReDim dict_pItems(0 To 0)
    Else
        ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) + 1)
        ReDim Preserve dict_pItems(0 To UBound(dict_pItems) + 1)
    End If
    Dim dict_FormattedKey As String
    dict_FormattedKey = dict_GetFormattedKey(dict_Key)
    If dict_Index >= 0 And dict_Index < dict_pKeyValues.Count Then
        Dim dict_i As Long
        For dict_i = UBound(dict_pKeys) To dict_Index + 1 Step -1
            dict_pKeys(dict_i) = dict_pKeys(dict_i - 1)
            If VBA.IsObject(dict_pItems(dict_i - 1)) Then
                Set dict_pItems(dict_i) = dict_pItems(dict_i - 1)
            Else
                dict_pItems(dict_i) = dict_pItems(dict_i - 1)
            End If
        Next dict_i
        dict_pKeys(dict_Index) = dict_Key
        If VBA.IsObject(dict_Value) Then
            Set dict_pItems(dict_Index) = dict_Value
        Else
            dict_pItems(dict_Index) = dict_Value
        End If
        dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey, Before:=dict_Index + 1
    Else
        If VBA.IsObject(dict_Key) Then
            Set dict_pKeys(UBound(dict_pKeys)) = dict_Key
        Else
            dict_pKeys(UBound(dict_pKeys)) = dict_Key
        End If
        If VBA.IsObject(dict_Value) Then
            Set dict_pItems(UBound(dict_pItems)) = dict_Value
        Else
            dict_pItems(UBound(dict_pItems)) = dict_Value
        End If
        dict_pKeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey
    End If
End Sub

Private Sub dict_ReplaceKeyValue(dict_KeyValue As Variant, dict_Key As Variant, dict_Value As Variant)
    Dim dict_Index As Long
    Dim dict_i As Integer
    dict_Index = dict_GetKeyIndex(dict_KeyValue(1))
    dict_RemoveKeyValue dict_KeyValue, dict_Index
    dict_AddKeyValue dict_Key, dict_Value, dict_Index
End Sub

Private Sub dict_RemoveKeyValue(dict_KeyValue As Variant, Optional ByVal dict_Index As Long = -1)
    Dim dict_i As Long
    If dict_Index = -1 Then
        dict_Index = dict_GetKeyIndex(dict_KeyValue(1))
    End If
    If dict_Index >= 0 And dict_Index <= UBound(dict_pKeys) Then
        For dict_i = dict_Index To UBound(dict_pKeys) - 1
            dict_pKeys(dict_i) = dict_pKeys(dict_i + 1)
            If VBA.IsObject(dict_pItems(dict_i + 1)) Then
                Set dict_pItems(dict_i) = dict_pItems(dict_i + 1)
            Else
                dict_pItems(dict_i) = dict_pItems(dict_i + 1)
            End If
        Next dict_i
        If UBound(dict_pKeys) = 0 Then
            Erase dict_pKeys
            Erase dict_pItems
        Else
            ReDim Preserve dict_pKeys(0 To UBound(dict_pKeys) - 1)
            ReDim Preserve dict_pItems(0 To UBound(dict_pItems) - 1)
        End If
    End If
    dict_pKeyValues.Remove dict_KeyValue(0)
    dict_RemoveObjectKey dict_KeyValue(1)
End Sub

Private Function dict_GetFormattedKey(dict_Key As Variant) As String
    If VBA.IsObject(dict_Key) Then
        dict_GetFormattedKey = dict_GetObjectKey(dict_Key)
    ElseIf VarType(dict_Key) = VBA.vbBoolean Then
        dict_GetFormattedKey = IIf(dict_Key, "-1__-1", "0__0")
    ElseIf VarType(dict_Key) = VBA.vbString Then
        dict_GetFormattedKey = dict_Key

        If Me.CompareMode = CompareMethod.BinaryCompare Then
            Dim dict_Lowercase As String
            dict_Lowercase = ""
            Dim dict_i As Integer
            Dim dict_Char As String
            Dim dict_Ascii As Integer
            For dict_i = 1 To VBA.Len(dict_GetFormattedKey)
                dict_Char = VBA.Mid$(dict_GetFormattedKey, dict_i, 1)
                dict_Ascii = VBA.Asc(dict_Char)
                If dict_Ascii >= 97 And dict_Ascii <= 122 Then
                    dict_Lowercase = dict_Lowercase & dict_Char
                Else
                    dict_Lowercase = dict_Lowercase & "_"
                End If
            Next dict_i
            If dict_Lowercase <> "" Then
                dict_GetFormattedKey = dict_GetFormattedKey & "__" & dict_Lowercase
            End If
        End If
    Else
        dict_GetFormattedKey = VBA.CStr(dict_Key) & "__" & CStr(dict_Key)
    End If
End Function

Private Function dict_GetObjectKey(dict_ObjKey As Variant) As String
    Dim dict_i As Integer
    For dict_i = 1 To dict_pObjectKeys.Count
        If dict_pObjectKeys.Item(dict_i) Is dict_ObjKey Then
            dict_GetObjectKey = "__object__" & dict_i
            Exit Function
        End If
    Next dict_i
    dict_pObjectKeys.Add dict_ObjKey
    dict_GetObjectKey = "__object__" & dict_pObjectKeys.Count
End Function

Private Sub dict_RemoveObjectKey(dict_ObjKey As Variant)
    Dim dict_i As Integer
    For dict_i = 1 To dict_pObjectKeys.Count
        If dict_pObjectKeys.Item(dict_i) Is dict_ObjKey Then
            dict_pObjectKeys.Remove dict_i
            Exit Sub
        End If
    Next dict_i
End Sub

Private Function dict_GetKeyIndex(dict_Key As Variant) As Long
    Dim dict_i As Long
    For dict_i = 0 To UBound(dict_pKeys)
        If VBA.IsObject(dict_pKeys(dict_i)) And VBA.IsObject(dict_Key) Then
            If dict_pKeys(dict_i) Is dict_Key Then
                dict_GetKeyIndex = dict_i
                Exit For
            End If
        ElseIf VBA.IsObject(dict_pKeys(dict_i)) Or VBA.IsObject(dict_Key) Then
        ElseIf dict_pKeys(dict_i) = dict_Key Then
            dict_GetKeyIndex = dict_i
            Exit For
        End If
    Next dict_i
End Function

Private Sub Class_Initialize()
    Set dict_pKeyValues = New Collection
    Erase dict_pKeys
    Erase dict_pItems
    Set dict_pObjectKeys = New Collection
End Sub

Private Sub Class_Terminate()
    Set dict_pKeyValues = Nothing
    Set dict_pObjectKeys = Nothing
End Sub
