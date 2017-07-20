VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   525
   ClientWidth     =   10725
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileAdmin 
         Caption         =   "Admin"
         Begin VB.Menu mnuFileAdminLogin 
            Caption         =   "Login"
         End
      End
      Begin VB.Menu mnuFileCustomer 
         Caption         =   "Customer"
         Begin VB.Menu mnuFileCustomerRegister 
            Caption         =   "Register"
         End
         Begin VB.Menu mnuFileCustomerLogin 
            Caption         =   "Login"
         End
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuFileAdminLogin_Click()
    Call LoginAdmin.Show
    mnuFileCustomer.Enabled = False
End Sub

Private Sub mnuFileExit_Click()
    Dim ans As String
    ans = MsgBox("Are Sure You Want To Exit This Program", vbYesNo, "Comfirmation")
    If ans = vbYes Then
        End
    Else
    End If
End Sub

Private Sub mnuFileCustomerLogin_Click()
    Call LoginCustomer.Show
    mnuFileCustomerRegister.Enabled = False
    mnuFileAdmin.Enabled = False
End Sub

Private Sub mnuFileCustomerRegister_Click()
    Call CustomerRegister.Show
    mnuFileCustomerLogin.Enabled = False
    mnuFileAdmin.Enabled = False
End Sub

Private Sub mnuHelpAbout_Click()
    Call frmAbout.Show
End Sub
