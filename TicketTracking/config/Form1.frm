VERSION 5.00
Begin VB.Form frmCustomer 
   Caption         =   "ADD Customers"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10410
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
   ScaleHeight     =   5025
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.TextBox District 
         Height          =   405
         Left            =   7320
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtState 
         Height          =   405
         Left            =   7320
         TabIndex        =   12
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtCity 
         Height          =   405
         Left            =   7320
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtCusId 
         Height          =   405
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         TabIndex        =   7
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtMobile 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txtCustomer 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label lblDistrict 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   15
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   13
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label lblCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   11
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label lblEmailID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   3
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label lblMobileNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   2
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerName"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   1
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label lblCustomerId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomerId"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()

    Dim customerObj As New customer

    customerObj.setCustomerId = txtCusId.Text
    customerObj.setCustomerName = txtCustomer.Text
    customerObj.setMobileNum = txtMobile.Text
    customerObj.setEmailId = txtMobile.Text

    customerObj.save
    
End Sub

