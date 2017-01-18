VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionMsSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ado As ADODB.connection

Property Get getADO() As ADODB.connection
    Set getADO = ado
End Property
                   
Private Sub Class_Initialize()
    Set ado = New ADODB.connection
    ConnectString = "DRIVER=SQL Server;SERVER=otz-prod1;Trusted_Connection=Yes;"
    ado.ConnectionTimeout = 0
    ado.CommandTimeout = 0
    ado.Open ConnectString
End Sub
