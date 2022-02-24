VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Private eid As String
Private employeeName As String
Private hireDate As Date
Private dept As String

Public Property Let setEmployeeId(ByVal newValue As String)
    eid = newValue
End Property
 
Public Property Get getEmployeeId() As String
    getCustomerId = eid
End Property
 


Public Property Let setemployeeName(ByVal newValue As String)
    employeeName = newValue
End Property
 
Public Property Get getemployeeName() As String
    getemployeeName = employeeName
End Property


Public Property Let sethireDate(ByVal newValue As Date)
    hireDate = newValue
End Property

Public Property Get gethireDate() As Date
    gethireDate = hireDate
End Property



Public Property Let setdept(ByVal newValue As String)
    dept = newValue
End Property
 
Public Property Get getdept() As String
    getdept = dept
End Property
 
 
 
Public Function loadEmployee(ByVal empId As String, ByVal password As String) As String
    Dim myRecSet As ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT  *" & _
          "FROM employeeAuthentication AS e WITH(NOLOCK) WHERE e.eid = " & empId & "e.eid =" & password
    
    Set myRecSet = myConnection.Execute(sql)
    
    loadEmployee = "Employee not found!"
    
    
    
    
    
    If Not myRecSet Is Nothing Then
        If myRecSet.State = adStateOpen Then
            If myRecSet.RecordCount < 0 Then
            
                
'                eid = empId
'                employeeName = myRecSet("")
'                hireDate
'                dept
'
                
                loadEmployee = "Ok"
                
                Exit Function
            End If
        End If
    End If
    
    Set myRecSet = Nothing
    
End Function

 
 
 
