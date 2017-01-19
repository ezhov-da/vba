Attribute VB_Name = "Properties"
Public Function getProperties(key As String) As String
    Dim query  As String
    query = "select settings from OTZ.dbo.T_E_wassort_settings where nameSettings = '" & key & "'"
    Dim setting As String
    
    Dim ado As ADODB.connection
    Set ado = connection.getADO
    
    Dim myRecordset As ADODB.Recordset
    Set myRecordset = ado.execute(query)
    

    Do Until myRecordset.EOF

        setting = myRecordset("settings")

    myRecordset.MoveNext
    Loop

    myRecordset.Close
    Set myRecordset = Nothing
    
    getProperties = setting
End Function
