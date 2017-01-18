Attribute VB_Name = "ErrorCatch"
Public Sub catch(err As ErrObject)
    If err.number <> 0 Then
        MsgBox err.Description, vbOKOnly + vbExclamation, "Ошибка"
    End If
    Application.ScreenUpdating = True
End Sub
