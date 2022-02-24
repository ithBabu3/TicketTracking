VERSION 5.00
Begin VB.Form frmCloseTickets 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
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
   ScaleHeight     =   3555
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   3120
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.TextBox txtResol 
         Height          =   1335
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox ComTicket 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblResolution 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblTicketId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TicketId"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmCloseTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim logTicket As New logTicket


Private Sub cmdSave_Click()
    If logTicket.closeTickets(ComTicket.Text, txtResol.Text) = "Ok" Then
        MsgBox "Ticket Closed"
    Else
        MsgBox "Error Occur in data base"
      End If
      
        
End Sub

Private Sub Form_Load()
    Dim eObj As Object
    Set eObj = logTicket.getTicketCollection
    Dim vArray As Variant
    vArray = eObj.Keys
    For i = 0 To eObj.Count - 1
        ComTicket.AddItem vArray(i)
    Next
    
End Sub
