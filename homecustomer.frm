VERSION 5.00
Begin VB.Form HomeCustomer 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   14430
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   14430
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   30
      Left            =   480
      Top             =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H000040C0&
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   14175
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "CUSTOMER PROFILE"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   4575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "BUY PRODUCT"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "M  A  I  N    M  E  N  U"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   27.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   13455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Height          =   855
      Left            =   3000
      TabIndex        =   6
      Top             =   6480
      Width           =   8175
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "QUIT"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Height          =   855
      Left            =   3000
      TabIndex        =   5
      Top             =   5520
      Width           =   8175
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " CONTINUE"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   21.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14175
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
         TabIndex        =   4
         Top             =   600
         Width           =   13935
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
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   13695
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1320
         Width           =   11055
      End
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
         TabIndex        =   1
         Top             =   2280
         Width           =   13935
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "M  A  I  N    M  E  N  U"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   13455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   48
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1095
      Left            =   4080
      TabIndex        =   14
      Top             =   3600
      Width           =   6495
   End
End
Attribute VB_Name = "HomeCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
CustomerProfile.Show
Unload Me
MDIForm1!mnuFileCustomer.Enabled = False
MDIForm1!mnuFileAdmin.Enabled = False
End Sub

Private Sub Command4_Click()
Frame4.Visible = False
End Sub

Private Sub Command5_Click()
Customer.Show
Unload Me
MDIForm1!mnuFileAdmin.Enabled = False
End Sub

Private Sub Form_Load()
Frame4.Visible = False
End Sub

Private Sub Label3_Click()
Frame4.Visible = True
Command4.SetFocus
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontSize = 24
Label3.ForeColor = &H80&
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontSize = 22
Label3.ForeColor = &HC0&
End Sub

Private Sub Label6_Click()
    Call LoginCustomer.Show
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = True
    MDIForm1!mnuFileCustomerLogin.Enabled = True
    MDIForm1!mnuFileCustomerRegister.Enabled = False
    MDIForm1!mnuFileAdmin.Enabled = False
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontSize = 24
Label6.ForeColor = &H80&
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.FontSize = 22
Label6.ForeColor = &HC0&
End Sub

Private Sub Timer1_Timer()
Label8.Left = Label8.Left + 100
If Label8.Left >= 8800 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Label8.Left = Label8.Left - 100
If Label8.Left <= 0 Then
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub


