Attribute VB_Name = "MonthGenerator"
Option Explicit

Public Function generate() As String
    Dim collectionMonth As New Collection
    Set collectionMonth = monthGenerator(12, "mmmm yyyy")
    
    Dim strArray As String
    Dim i As Integer
    For i = 1 To collectionMonth.Count
        strArray = strArray & collectionMonth.Item(i) & ";"
    Next i
    
    generate = Left(strArray, Len(strArray) - 1)
End Function


Public Function monthGenerator(countMonthFuture As Integer, frmtDateOut As String) As Collection
    Dim collectionMonth As New Collection

    Dim textMonth As String
    
    Dim monthNum As Integer
    monthNum = Month(Now)
    Dim yearNum As Integer
    yearNum = Year(Now)
    
    Dim counter As Integer
    
    Dim dateNow As Date
    
    dateNow = CDate("01." & CStr(monthNum) & "." & CStr(yearNum))
    
    Dim dateCurrent As Date
    dateCurrent = dateNow
    For counter = 0 To countMonthFuture
        textMonth = Format(dateCurrent, frmtDateOut)
        dateCurrent = dateCurrent + 31
        collectionMonth.Add CStr(textMonth)
    Next counter
    Set monthGenerator = collectionMonth
End Function
