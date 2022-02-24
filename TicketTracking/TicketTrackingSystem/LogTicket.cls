VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

 
Private ticketId As Integer
Private loggedBy As String
Private raisedBy As Date
Private severity As String
Private ticketDesc As String
Private resolvedBy As String
Private resolution As String
Private resolvedDate As Date
Private status As String





Public Property Let setTicketId(ByVal newValue As Integer)
    ticketId = newValue
End Property
 
Public Property Get getTicketId() As Integer
    getTicketId = ticketId
End Property
 


Public Property Let setloggedBy(ByVal newValue As String)
    loggedBy = newValue
End Property
 
Public Property Get getloggedBy() As String
    getloggedBy = loggedBy
End Property



Public Property Let setraisedBy(ByVal newValue As Date)
    raisedBy = newValue
End Property
 
Public Property Get getraisedBy() As Date
    getraisedBy = raisedBy
End Property
 
 
 Public Property Let setseverity(ByVal newValue As String)
    severity = newValue
End Property

Public Property Get getseverity() As String
    getseverity = severity
End Property
 

 Public Property Let setticketDesc(ByVal newValue As String)
    ticketDesc = newValue
End Property
 
Public Property Get getticketDesc() As String
    getticketDesc = ticketDesc
End Property


Public Property Let setresolvedBy(ByVal newValue As String)
    resolvedBy = newValue
End Property
 
Public Property Get getresolvedBy() As String
    getresolvedBy = resolvedBy
End Property


Public Property Let setresolution(ByVal newValue As String)
    resolution = newValue
End Property
 
Public Property Get getresolution() As String
    getresolution = resolution
End Property


Public Property Let setresolvedDate(ByVal newValue As Date)
    resolvedDate = newValue
End Property
 
Public Property Get getresolvedDate() As Date
    getresolvedDate = resolvedDate
End Property



Public Property Let setstatus(ByVal newValue As String)
    status = newValue
End Property
 
Public Property Get getstatus() As String
    getstatus = status
End Property


Public Function save() As String
    Dim sql As String
    Dim isExist As Boolean
    Dim message As String
    Dim empAuthObj As EmployeeAuthentication
    

    On Error GoTo error_handler
    
      
    setstatus = "Open"
    

    setraisedBy = Format(Now, "mm/dd/yy hh:mm")
 
    
    sql = "SET DATEFORMAT DMY" & vbCrLf & _
          "INSERT INTO ticket([loggedBy],[raisedDate],[severity],[ticketDesc],[resolvedBy],[resolution],[resolvedDate],[status])"
    sql = sql & "VALUES(" & "'" & mEmpId & "'," & "'" & raisedBy & "'," & "'" & severity & "'," & "'" & ticketDesc & "'," & "''," & "''," & "''," & "'" & status & "'" & ")"
    

    Call myConnection.Execute(sql)
    
    
    message = "Ok"
    save = message
    
    Exit Function
    
error_handler:
    message = "Error! " & Err.Number & ": " & Err.Description
    
End Function


Public Function viewTickets(Optional ByVal statusVal As String, Optional ByVal severityVal As String) As Boolean
    Dim myRecSet As New ADODB.Recordset
   

    Dim sql As String
    
    If (status = "" And severity = "") Then
        sql = "select * from ticket"
    End If
    
     Set myRecSet = myConnection.Execute(sql)
   
   

     While Not myRecSet.EOF
                    ticetid = myRecSet(ticketId)
                    loggedBy = myRecSet("loggedBy")
                    raisedBy = myRecSet("raisedDate")
                    severity = myRecSet("severity")
                    ticketDesc = myRecSet("ticketDesc")
                    resolvedBy = myRecSet("resolvedBy")
                    resolution = myRecSet("resolution")
                    resolvedDate = myRecSet("resolvedDate")
                    status = myRecSet("status")

                    loadEmployeeAuth = True
        Wend
    
    While Not myRecSet.EOF
        ReDim Preserve arr(arrSize)
        arr(arrSize) = myRecSet("dept").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
    
    getDep = arr
  
End Function



'Public Function getTickets() As String()
'
'
'ElseIf (status <> "") Then
'        sql = "SELECT  * FROM ticket AS t WITH(NOLOCK)" & _
'            " WHERE t.status = '" & statusVal & "' "
'    ElseIf severity <> "" Then
'            sql = "SELECT  * FROM ticket AS t WITH(NOLOCK)" & _
'            " WHERE t.severity = '" & severityVal & "' "
'    Else
'        sql = "SELECT  * FROM ticket AS t WITH(NOLOCK)" & _
'          " WHERE t.severity = '" & severityVal & "' AND t.status =  '" & statusVal & "'"
'
'    Dim myRecSet As ADODB.Recordset
'    Set myRecSet = myConnection.Execute("SELECT * FROM ticket as t")
'
'   While Not myRecSet.EOF
'        ReDim Preserve arrUserNames(arrSize)
'        arrUserNames(arrSize) = myRecSet("employeeName").Value
'        arrSize = arrSize + 1
'        myRecSet.MoveNext
'    Wend
'
'
'    loadEmployeeAuth = False
'End Function


Public Function closeTickets(ByVal ticketId As Integer, ByVal resolution As String) As String

    Dim sql As String
    Dim isExist As Boolean
    Dim message As String
    Dim empAuthObj As EmployeeAuthentication
    

    On Error GoTo error_handler
    
      
    setstatus = "Closed"
    

    
    setresolvedDate = Format(Now, "mm/dd/yy hh:mm")
 
    
    
    sql = "SET DATEFORMAT DMY" & vbCrLf & _
          "UPDATE  ticket SET "
    sql = sql & "status = " & "'" & status & "'," & "resolvedBy =" & "'" & mEmpId & "'," & "resolvedDate = " & "'" & resolvedDate & "'," & "resolution =" & "'" & resolution & "'" & _
           "WHERE ticketId =" & "'" & ticketId & "'" & ""
    

    Call myConnection.Execute(sql)
    
    
    message = "Ok"
    closeTickets = message

    
    Exit Function
    
error_handler:
    message = "Error! " & Err.Number & ": " & Err.Description
  
End Function



Public Function getTicketCollection() As Dictionary
      Set getTicketCollection = ticketCollection
End Function


Public Function getSearchCollection(Optional ByVal status As String, Optional ByVal seveirty As String) As Dictionary
      Call getSearchTicketCollection(status, seveirty)
      Set getSearchCollection = searchCollection
End Function




