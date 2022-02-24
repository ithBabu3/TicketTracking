VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SetUpConnection()
    Dim dbServer, dbName As String
    Dim configObject As New Confiuration
    Dim configString As String
    

    configString = readConfig(App.Path & "\" & "Config.ini")
    dbServer = Trim(configObject.getConfigFrom(configString, "Db_Server"))
    dbName = Trim(configObject.getConfigFrom(configString, "Db_Name"))
    
    Set myConnection = New ADODB.Connection
    myConnection.ConnectionString = "DRIVER=SQL Server;SERVER=" & dbServer & ";DATABASE=" & dbName & ";"
    Call myConnection.Open
    
    SetUpConnection = myConnection
End Function

Public Function readConfig(ByRef configPath As String) As String
    Dim fsObject As New FileSystemObject
    On Error Resume Next
    
    readConfig = fsObject.OpenTextFile(configPath, ForReading, False).ReadAll
End Function
