VERSION 5.00
Begin VB.Form frmCreateNewTicket 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
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
   ScaleHeight     =   4845
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraE 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblSeverity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Severity"
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
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Width           =   1005
      End
      Begin VB.Label lblTicketDescrption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Descrption"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   2145
      End
   End
   Begin VB.TextBox txtDesc 
      Height          =   1095
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.ComboBox CombSeverity 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3120
      Width           =   3855
   End
   Begin VB.ComboBox ComUserName 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Top             =   4320
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3960
      TabIndex        =   0
      Top             =   4320
      Width           =   990
   End
End
Attribute VB_Name = "frmCreateNewTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim empAuth As New EmployeeAuthentication
        
Private Sub cmdCancel_Click()
        Unload Me
End Sub

Private Sub cmdSave_Click()

        Dim logTicketObj As New logTicket

        logTicketObj.setticketDesc = txtDesc.Text
        logTicketObj.setseverity = CombSeverity.Text
        logTicketObj.setloggedBy = empAuth.getuserId

        Call logTicketObj.Save
        
        MsgBox "Ticket has been submited"
        
End Sub

Private Sub Form_Load()
    
     Dim arr() As String
     arr = empAuth.getAllUsers()

    For i = 0 To UBound(arr)
         ComUserName.AddItem (arr(i))
    Next

    CombSeverity.AddItem ("Major")
    CombSeverity.AddItem ("Critical")
    CombSeverity.AddItem ("Normal")
End Sub



