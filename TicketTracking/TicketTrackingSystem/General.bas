Attribute VB_Name = "General"
Public myConnection As ADODB.Connection
Public ticketCollection As New Dictionary
Public searchCollection As New Dictionary
Private arr() As String
Private arrUserNames() As String




Public Type ticket
    loggedBy As String
    rasiedDate As Date
    severity As String
    ticketDes As String
    resolvedBy As String
    resolutio As String
    resolvedDate As Date
    status As String
End Type


Public ticketArr(8) As ticket


Public Function getDep() As String()
    Dim myRecSet As ADODB.Recordset
    Set myRecSet = myConnection.Execute("SELECT e.dept FROM employee as e GROUP BY e.dept")
    
    While Not myRecSet.EOF
        ReDim Preserve arr(arrSize)
        arr(arrSize) = myRecSet("dept").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
   
    Call getTickets

    getDep = arr
    
End Function

Public Function getAllUsersName() As String()
    Dim myRecSet As ADODB.Recordset
    Set myRecSet = myConnection.Execute("SELECT e.employeeName FROM employee as e")
    
    While Not myRecSet.EOF
        ReDim Preserve arrUserNames(arrSize)
        arrUserNames(arrSize) = myRecSet("employeeName").Value
        arrSize = arrSize + 1
        myRecSet.MoveNext
    Wend
    getAllUsersName = arrUserNames
    
End Function

Public Function getTickets() As Dictionary

     Dim myRecSet As ADODB.Recordset

     Set myRecSet = myConnection.Execute("SELECT * FROM ticket as t")

     While Not myRecSet.EOF
    
        Dim logDataCollection As New Collection

        Call logDataCollection.Add(myRecSet("loggedBy").Value)
        Call logDataCollection.Add(myRecSet("raisedDate").Value)
        Call logDataCollection.Add(myRecSet("severity").Value)
        Call logDataCollection.Add(myRecSet("ticketDesc").Value)
        Call logDataCollection.Add(myRecSet("resolvedBy").Value)
        Call logDataCollection.Add(myRecSet("resolution").Value)
        Call logDataCollection.Add(myRecSet("resolvedDate").Value)
        Call logDataCollection.Add(myRecSet("status").Value)
        
        Call ticketCollection.Add(myRecSet("ticketId").Value, logDataCollection)
        
        Set logDataCollection = Nothing
        
        
        myRecSet.MoveNext
    
    Wend

    Set getTickets = ticketCollection

End Function


Public Function getSearchTickets(Optional ByVal status As String, Optional ByVal sevirity As String) As Dictionary

     Dim myRecSet As ADODB.Recordset
    
    If status = "" Then
    Set myRecSet = myConnection.Execute("SELECT * FROM ticket as t where severity = '" & sevirity & "'")

    ElseIf sevirity = "" Then
         Set myRecSet = myConnection.Execute("SELECT * FROM ticket as t where status = '" & status & "'")
    ElseIf status = "" And sevirity = "" Then
         Set getSearchTickets = ticketCollection
         Exit Function
    Else
    Set myRecSet = myConnection.Execute("SELECT * FROM ticket as t where severity = '" & sevirity & "' and status = '" & status & "'  ")
    End If
    
     While Not myRecSet.EOF
    
        Dim logDataCollection As New Collection

        Call logDataCollection.Add(myRecSet("loggedBy").Value)
        Call logDataCollection.Add(myRecSet("raisedDate").Value)
        Call logDataCollection.Add(myRecSet("severity").Value)
        Call logDataCollection.Add(myRecSet("ticketDesc").Value)
        Call logDataCollection.Add(myRecSet("resolvedBy").Value)
        Call logDataCollection.Add(myRecSet("resolution").Value)
        Call logDataCollection.Add(myRecSet("resolvedDate").Value)
        Call logDataCollection.Add(myRecSet("status").Value)
        
        Call searchCollection.Add(myRecSet("ticketId").Value, logDataCollection)
        
        Set logDataCollection = Nothing
        
        
        myRecSet.MoveNext
    
    Wend

    Set getSearchTickets = searchCollection

End Function




Public Function getGenTicketCollection() As Dictionary

    Set getTicketCollection = ticketCollection

End Function

Public Function getSearchTicketCollection(Optional ByVal status As String, Optional ByVal seveirty As String) As Dictionary
    Call getSearchTickets(status, seveirty)
    Set getSearchTicketCollection = searchCollection

End Function
