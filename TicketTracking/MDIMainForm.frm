VERSION 5.00
Begin VB.MDIForm MDIMainForm 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5520
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10740
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu cTickets 
      Caption         =   "Create Ticket"
   End
   Begin VB.Menu closeTickets 
      Caption         =   "Close a Ticket"
   End
   Begin VB.Menu DTickets 
      Caption         =   "Display Closed Tickets"
   End
   Begin VB.Menu vTickets 
      Caption         =   "View Tickets"
   End
   Begin VB.Menu mLogOut 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "MDIMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub closeTickets_Click()
        frmCloseTickets.Show

End Sub

Private Sub cTickets_Click()
        frmCreateNewTicket.Show
End Sub

Private Sub LTickets_Click()

End Sub

Private Sub DTickets_Click()

Dim crApp As New CRAXDRT.Application
Dim crRpt As New CRAXDRT.Report

Dim filepath As String

filepath = "D:\eurofins\Training\ShopOn\ShopOn\print.rpt"
    

Set crRpt = crApp.OpenReport(filepath)
frmDisplayTickets.CRViewer91.ReportSource = crRpt
frmDisplayTickets.CRViewer91.ViewReport
frmDisplayTickets.Show

    
End Sub

Private Sub MDIForm_Load()
       Dim empAutObj As New EmployeeAuthentication
       
       vTickets.Visible = False
       closeTickets.Visible = False
        
       If (empAutObj.dept = "DEVOPS") Then
        vTickets.Visible = True
        cTickets.Visible = False
        closeTickets.Visible = True
       End If
          
End Sub

Private Sub mLogOut_Click()
        frmLogin.Show
        Unload Me
End Sub

Private Sub vTickets_Click()
            frmViewTickets.Show
End Sub
