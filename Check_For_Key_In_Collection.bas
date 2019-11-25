Attribute VB_Name = "Check_For_Key_In_Collection"
Public Function HasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    If IsEmpty(coll(strKey)) = True Then
        HasKey = False
        'MsgBox (" - My Var")
    Else
        HasKey = True
        'MsgBox (coll(strKey) & " - My Key")
    End If
    'err.Number = 0
    err.Clear
End Function
