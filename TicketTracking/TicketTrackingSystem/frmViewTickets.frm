VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmViewTickets 
   Caption         =   "View Tickets"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10710
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
   ScaleHeight     =   7185
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10335
      Begin VB.ListBox List1 
         BackColor       =   &H8000000F&
         Height          =   5715
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   7935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   8040
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         Begin VB.CommandButton cmdReset 
            Caption         =   "Reset"
            Height          =   360
            Left            =   1200
            TabIndex        =   9
            Top             =   2160
            Width           =   870
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   360
            Left            =   120
            TabIndex        =   7
            Top             =   2160
            Width           =   990
         End
         Begin VB.ComboBox ComSevier 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1560
            Width           =   1935
         End
         Begin VB.ComboBox ComStatus 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblSeverity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Severity"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label lblStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   555
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   30
         Index           =   0
         Left            =   3600
         TabIndex        =   1
         Top             =   3360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmViewTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim logTicket As New logTicket

Private Sub cmdReset_Click()
    List1.Clear
    Call loadData
    
End Sub

Private Sub cmdSearch_Click()


    List1.Clear


 Dim key As Variant
    Dim i As Integer

    Dim cObj, eObj As Object

   Set eObj = logTicket.getSearchCollection(ComStatus.Text, ComSevier.Text)
   
   
   
   cmdReset.Visible = True
   
    

    For Each key In eObj.Keys
       Set cObj = eObj.Item(key)

     Dim loggedBY, Desc, severity, rasiedDate, resolDate, resolBy, status, resolution As String
    
       For i = 1 To cObj.Count
       
           If i = 1 Then
               If IsNull(cObj.Item(i)) = True Then
                loggedBY = "NULL"
             Else
             loggedBY = cObj.Item(i)
             End If
             
            
                ElseIf i = 2 Then
                 If IsNull(cObj.Item(i)) = True Then
                    rasiedDate = "NULL"
                Else
                rasiedDate = cObj.Item(i)
                End If
                
                
            ElseIf i = 3 Then
                If IsNull(cObj.Item(i)) = True Then
                    severity = "NULL"
                Else
                severity = cObj.Item(i)
                End If
                
            ElseIf i = 4 Then
                If IsNull(cObj.Item(i)) = True Then
                    Desc = "NULL"
                Else
                Desc = cObj.Item(i)
                End If
                

          ElseIf i = 5 Then
               If IsNull(cObj.Item(i)) = True Then
                    resolBy = "NULL"
                Else
                resolBy = cObj.Item(i)
                End If
                
            
            ElseIf i = 6 Then
                If IsNull(cObj.Item(i)) = True Then
                    resolution = "NULL"
                Else
                resolution = cObj.Item(i)
                End If
                

            ElseIf i = 7 Then
                If IsNull(cObj.Item(i)) = True Then
                    resolDate = "NULL"
                Else
                resolDate = cObj.Item(i)
                End If
                
            ElseIf i = 8 Then
                If IsNull(cObj.Item(i)) = True Then
                    status = "NULL"
                Else
                status = cObj.Item(i)
                End If

       
        End If
        
          
            
    
       Next
        List1.AddItem "Severity :   " & severity & vbTab & vbTab & vbTab & vbTab & vbTab & "RasiedDate : " & rasiedDate
        List1.AddItem " "
        List1.AddItem vbTab & "LoggedBY  :" & vbTab & loggedBY
        List1.AddItem vbTab & "Description :" & vbTab & Desc
        'List1.AddItem vbTab & "RasiedDate :" & vbTab & rasiedDate
        List1.AddItem vbTab & "Resolved By :" & vbTab & resolBy
        List1.AddItem vbTab & "Resolution :" & vbTab & resolution
        List1.AddItem vbTab & "Resolved By :" & vbTab & resolDate
        List1.AddItem vbTab & "Status :" & vbTab & vbTab & status
        List1.AddItem ""
        List1.AddItem vbTab & "------------------------------------------------------------------------------------------------------"
    Next
   
  
   
End Sub

Private Sub Form_Load()
       
  
    ComStatus.AddItem ("Open")
     ComStatus.AddItem ("Closed")
     
     ComSevier.AddItem ("Critical")
     ComSevier.AddItem ("Normal")
     ComSevier.AddItem ("Major")
     
   
    cmdReset.Visible = False
    
     Call loadData
     
   
        
End Sub

Public Sub loadData()
 Dim key As Variant
    Dim i As Integer

    Dim cObj, eObj As Object

   Set eObj = logTicket.getTicketCollection

    For Each key In logTicket.getTicketCollection.Keys
       Set cObj = logTicket.getTicketCollection.Item(key)
        Dim loggedBY, Desc, severity, rasiedDate, resolDate, resolBy, status, resolution As String
    
       For i = 1 To cObj.Count
       
           If i = 1 Then
               If IsNull(cObj.Item(i)) = True Then
                loggedBY = "NULL"
             Else
             loggedBY = cObj.Item(i)
             End If
             
            
                ElseIf i = 2 Then
                 If IsNull(cObj.Item(i)) = True Then
                    rasiedDate = "NULL"
                Else
                rasiedDate = cObj.Item(i)
                End If
                
                
            ElseIf i = 3 Then
                If IsNull(cObj.Item(i)) = True Then
                    severity = "NULL"
                Else
                severity = cObj.Item(i)
                End If
                
            ElseIf i = 4 Then
                If IsNull(cObj.Item(i)) = True Then
                    Desc = "NULL"
                Else
                Desc = cObj.Item(i)
                End If
                

          ElseIf i = 5 Then
               If IsNull(cObj.Item(i)) = True Then
                    resolBy = "NULL"
                Else
                resolBy = cObj.Item(i)
                End If
                
            
            ElseIf i = 6 Then
                If IsNull(cObj.Item(i)) = True Then
                    resolution = "NULL"
                Else
                resolution = cObj.Item(i)
                End If
                

            ElseIf i = 7 Then
                If IsNull(cObj.Item(i)) = True Then
                    resolDate = "NULL"
                Else
                resolDate = cObj.Item(i)
                End If
                
            ElseIf i = 8 Then
                If IsNull(cObj.Item(i)) = True Then
                    status = "NULL"
                Else
                status = cObj.Item(i)
                End If

       
        End If
        
          
            
    
       Next
        List1.AddItem "Severity :   " & severity & vbTab & vbTab & vbTab & vbTab & vbTab & "RasiedDate : " & rasiedDate
        List1.AddItem " "
        List1.AddItem vbTab & "LoggedBY  :" & vbTab & loggedBY
        List1.AddItem vbTab & "Description :" & vbTab & Desc
        'List1.AddItem vbTab & "RasiedDate :" & vbTab & rasiedDate
        List1.AddItem vbTab & "Resolved By :" & vbTab & resolBy
        List1.AddItem vbTab & "Resolution :" & vbTab & resolution
        List1.AddItem vbTab & "Resolved By :" & vbTab & resolDate
        List1.AddItem vbTab & "Status :" & vbTab & vbTab & status
        List1.AddItem ""
        List1.AddItem vbTab & "------------------------------------------------------------------------------------------------------"


    Next
   
End Sub

