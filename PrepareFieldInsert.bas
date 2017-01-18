Attribute VB_Name = "PrepareFieldInsert"
Public Function prepareString(str As String) As String
    If str = "" Then
        prepareString = "null"
    Else
        prepareString = "'" & Replace(str, "'", "''") & "'"
    End If
End Function

Public Function prepareDate(str As String) As String
    If str = "" Then
        prepareDate = "null"
    Else
        prepareDate = "cast('" & str & "' as date)"
    End If
End Function

Public Function prepareNumber(str As String) As String
    If str = "" Then
        prepareNumber = "null"
    Else
        prepareNumber = Replace(str, ",", ".")
    End If
End Function

Public Function prepareDateTime(str As String) As String
    If str = "" Then
        prepareDateTime = "null"
    Else
        prepareDateTime = "CONVERT(Datetime, '" & Format(str, "yyyy-mm-dd hh:mm:ss") & "', 120)"
    End If
End Function


