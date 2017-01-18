Attribute VB_Name = "GrantHolder"
Public Function getUsername() As String
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    getUsername = wshShell.ExpandEnvironmentStrings("%USERNAME%")
End Function

Public Function nowUserIsAdminMinPart() As Boolean
    nowUserIsAdminMinPart = isAdminFunction("admin.minpart")
End Function

Public Function nowUserIsAdminControlNovelty() As Boolean
    nowUserIsAdminControlNovelty = isAdminFunction("admin.control.novelty")
End Function

Public Function nowUserIsAdminShowBox() As Boolean
    nowUserIsAdminShowBox = isAdminFunction("admin.showBox")
End Function

Private Function isAdmin(arrAdmin As String, username As String) As Boolean
    Dim arr As Variant
    arr = Split(arrAdmin, ",")
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = username Then
            isAdmin = True
            Exit Function
        End If
    Next i
    isAdmin = False
End Function

Private Function isAdminFunction(strPropertyAdmin As String) As Boolean
    Dim arrAdmin As String
    arrAdmin = Properties.getProperties(strPropertyAdmin)
    Dim un As String
    un = getUsername()
    Dim b As Boolean
    b = isAdmin(arrAdmin, un)
    isAdminFunction = b
End Function
