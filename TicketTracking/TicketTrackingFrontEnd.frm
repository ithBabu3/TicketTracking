VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "avinash@123"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   480
      Left            =   4440
      TabIndex        =   1
      Top             =   3240
      Width           =   990
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   480
      Left            =   1800
      TabIndex        =   0
      Top             =   3240
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin VB.ComboBox ComDept 
         Height          =   315
         ItemData        =   "TicketTrackingFrontEnd.frx":0000
         Left            =   2640
         List            =   "TicketTrackingFrontEnd.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtEmpId 
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Text            =   "M100103"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label lblDepartment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label lblEmployeeID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim empAuth As New EmployeeAuthentication

    
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSubmit_Click()
  
    Dim strError As String
    
    If empAuth.loadEmployeeAuth(txtEmpId.Text, txtPassword.Text, ComDept.Text) Then
                 MDIMainForm.Show
    Else
        MsgBox "Data not found"
      End If

    
End Sub

Private Sub Form_Load()

    Dim dbConnection As New dbConnection
    dbConnection.SetUpConnection
    
     Dim arr() As String
     arr = empAuth.getAllDep()
    
    For i = 0 To UBound(arr)
         ComDept.AddItem (arr(i))
    Next
     


    
End Sub




