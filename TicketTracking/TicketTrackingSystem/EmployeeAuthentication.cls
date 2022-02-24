VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmployeeAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private eid As String
Private userId As String
Private password As String


Private empObj As New Employee
Private empAuthObj As New EmployeeAuthentication


Public Property Let setEmployeeId(ByVal newValue As String)
    eid = newValue
End Property
 
Public Property Get getEmployeeId() As String
    getCustomerId = eid
End Property
 

Public Property Let setuserId(ByVal newValue As String)
    userId = newValue
End Property
 
Public Property Get getuserId() As String
    getuserId = userId
End Property



Public Property Let setpassword(ByVal newValue As String)
    password = newValue
End Property
 
Public Property Get getpassword() As String
    getpassword = password
End Property
 
 
 
Public Function loadEmployeeAuth(ByVal empId As String, ByVal password As String, ByVal dept As String) As Boolean
    Dim myRecSet As New ADODB.Recordset
   

    Dim sql As String
  
    
    
   sql = "SELECT  * FROM employeeAuthentication AS ea WITH(NOLOCK)" & _
         "INNER  JOIN employee as e on e.eid=ea.eid" & _
          " WHERE ea.eid = '" & empId & "' AND ea.password =  '" & password & "' AND e.dept =  '" & dept & "'"

    Set myRecSet = myConnection.Execute(sql)
    
    
   
    loadEmployeeAuth = False

    If Not myRecSet Is Nothing Then
        If myRecSet.State = adStateOpen Then
            If Not myRecSet.EOF Then
                    
                    eid = empId
                    userId = myRecSet("userId")
                    password = myRecSet("password")
                    
                    getEmployeeObj.setemployeeName = myRecSet("employeeName")
                    getEmployeeObj.sethireDate = myRecSet("hireDate")
                    mDept = myRecSet("dept")
                    mEmpId = empId
                    
                    
                    
                    loadEmployeeAuth = True
                    
                    Exit Function
                    
                End If
            End If
    End If
  

    
End Function

 
 Public Function getAllDep() As String()
          getAllDep = getDep()
 End Function
 Public Function getAllUsers() As String()
         getAllUsers = getAllUsersName
        
 End Function


Public Function getEmployeeObj() As Employee
       Set getEmployeeObj = empObj
End Function

Public Function dept() As String
     dept = mDept
End Function

