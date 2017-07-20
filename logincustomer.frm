VERSION 5.00
Begin VB.Form LoginCustomer 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   14415
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   15375
      TabIndex        =   0
      Top             =   9600
      Width           =   15375
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5775
      Left            =   -720
      TabIndex        =   15
      Top             =   3960
      Width           =   15375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   0
      TabIndex        =   14
      Top             =   -1560
      Width           =   15375
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   840
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   360
         Top             =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2760
      TabIndex        =   6
      Top             =   4080
      Width           =   8775
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3480
         TabIndex        =   10
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3480
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2160
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME :"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "CUSTOMER LOGIN FORM"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   7575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   14175
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact : 06-6794563"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   2400
         Width           =   13935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No 1, 46, 47, 48, 50, Jalan Tampin, Taman Bukit Emas, 70450 Seremban, Negeri Sembilan."
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   975
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   10095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "FURNITURE ORDERING SYSTEM"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   27.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   13695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "KEMUDI YAKIN FURNITURE STORE"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   36
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   13935
      End
   End
End
Attribute VB_Name = "LoginCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub Command2_Click()
    LoginSucceeded = False
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = True
    MDIForm1!mnuFileCustomerRegister.Enabled = True
    MDIForm1!mnuFileCustomerLogin.Enabled = True
    MDIForm1!mnuFileAdmin.Enabled = True
End Sub
Private Sub Command1_Click()
    Dim struser As String
    Dim strpass As String

    OpenKYAKINDatabase
    OpenUserTable 'open student table in the registration database file
    
    struser = Text1
    strpass = Text2
    If Not Text1 = "" Then
        rs.Index = "user_username"
        rs.Seek "=", struser
        If rs.NoMatch Then
            MsgBox "INVALID USERNAME OR PASSWORD.. TRY AGAIN..", vbExclamation, "LOGIN"
            MsgBox "If u not a member, please register first", vbInformation, "LOGIN"
            Text1.Text = ""
            Text2.Text = ""
            Text1.SetFocus
        ElseIf (rs!user_username = struser And rs!Password = strpass) Then
            passn = rs!Password
            usern = rs!user_username
                LoginSucceeded = True
                Call HomeCustomer.Show
                Unload Me
                MDIForm1!mnuFileAdmin.Enabled = False
                MDIForm1!mnuFileCustomer.Enabled = False
        Else
            MsgBox "INVALID USERNAME OR PASSWORD.. TRY AGAIN..", vbExclamation, "LOGIN"
                MsgBox "If u not a member, please register first", vbInformation, "LOGIN"
                Text1.Text = ""
                Text2.Text = ""
                Text1.SetFocus
        End If
    Else
        MsgBox "Please make sure that you have inserted the username and password.", vbOKOnly, "Error"
    End If
End Sub

Private Sub Timer1_Timer()
Frame4.Top = Frame4.Top + 100
If Frame4.Top >= 4080 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
Frame3.Top = Frame3.Top - 100
If Frame3.Top <= -1560 Then
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub
