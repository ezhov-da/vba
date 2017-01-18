Attribute VB_Name = "Util"
Private Const FILE_HOLDER_DATE As String = "DATE_HOLDER_MIN.txt"
Private Const FILE_HOLDER_LOG As String = "E:\LOG_MIN_PART.txt"
Private Const WRITE_LOG As Boolean = False

Public Sub addColumnCalculate()
    Dim ado As ADODB.connection
    Set ado = connection.getADO
    
    Dim text As String
    
    r = 3
    
    Do While (Cells(r, 1) <> "")
        one = Left(Cells(r, 41).Value, 10)
        two = Left(Cells(r, 42).Value, 10)
        three = Cells(r, 43).Formula
        text = "insert into OTZ.dbo.T_E_wassort_insertBatchAddComment values('" & Cells(r, 18) & "', '" & one & "', '" & two & "', '" & three & "')"
        Call Util.writeLog(text)
        ado.execute (text)
        r = r + 1
    Loop
End Sub

'этот метод устанавливает последнюю дату выгрузки шаблона
Public Sub setDateTimeUnload()
    Dim lastDate As String
    lastDate = Format(Now, "yyyy-MM-dd HH:mm:ss")
    Open ThisWorkbook.Path & "\" & FILE_HOLDER_DATE For Output As #1
    Print #1, lastDate
    Close #1
End Sub

Public Function getLastDateTimeUnload() As String
    Dim file As String
    file = ThisWorkbook.Path & "\" & FILE_HOLDER_DATE
    If Dir(file) = "" Then
        getLastDateTimeUnload = ""
    Else
        Open file For Input As #1
        Dim s As String
        Input #1, s
        Close #1
        getLastDateTimeUnload = s
    End If
End Function

Public Sub writeLog(text As String)
    If (WRITE_LOG) Then
        Debug.Print text
    End If
End Sub

Public Sub protectSheetsTwoThree()
        Call protectSheet(2)
        Call protectSheet(3)
End Sub


Public Sub protectSheet(numberSheet As Integer)
        Worksheets(numberSheet).Protect Password:="ezhov_da", DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True, AllowInsertingRows _
        :=True, AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
        AllowUsingPivotTables:=True
End Sub

Public Sub paintRow(row As Long, colorIndexxxx As Integer)
    Dim rStr As String
    rStr = CStr(row) & ":" & CStr(row)
    Rows(rStr).Interior.colorIndex = colorIndexxxx
End Sub

Public Sub setTextToCell(row As Long, columnText As Integer, text As String)
    Cells(row, columnText).Value = text
End Sub


Public Function getSystemSeparator() As String
    Dim separator As String
    separator = Mid(CStr(1 / 2), 2, 1)
    getSystemSeparator = separator
End Function



